"""
yomiage.py — 選択テキスト読み上げアプリ
テキストを選択して Ctrl+Alt+R を押すと Windows TTS (日本語) で読み上げる。
読み上げ中に Esc を押すと即座に中断できる。
タスクバーのトレイアイコンとして常駐する。

依存パッケージ:
  pip install pystray pyperclip pillow

ホットキー:
  Ctrl+Alt+R : 選択テキストを読み上げ（読み上げ中なら停止）
  Esc        : 読み上げ中のみ有効。読み上げを即座に停止。
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import logging
import os
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path

# ログ設定
_LOG_PATH = Path(__file__).parent / "yomiage.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(_LOG_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stderr),
    ],
)
log = logging.getLogger("yomiage")

import pyperclip
import pystray
from PIL import Image, ImageDraw, ImageFont

# DPI awareness
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass


# =====================================================================
# Windows 定数
# =====================================================================
_INPUT_KEYBOARD   = 1
_KEYEVENTF_KEYUP  = 0x0002
_VK_CONTROL       = 0x11
_VK_C             = 0x43
_VK_R             = 0x52
_VK_ESCAPE        = 0x1B

_MOD_ALT          = 0x0001
_MOD_CONTROL      = 0x0002
_MOD_CTRL_ALT     = _MOD_ALT | _MOD_CONTROL

_WM_HOTKEY        = 0x0312
_WM_USER          = 0x0400

_HOTKEY_READ      = 1

_PS_FLAGS = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW

# クリップボード待ち時間
_CLIP_WAIT_1 = 0.15   # 最初の待機 (秒)
_CLIP_WAIT_2 = 0.25   # リトライ時の追加待機 (秒)


# =====================================================================
# ctypes SendInput — Ctrl+C 送信
# =====================================================================
_PUL = ctypes.POINTER(ctypes.c_ulong)


class _KeyBdInput(ctypes.Structure):
    _fields_ = [
        ("wVk",         wt.WORD),
        ("wScan",       wt.WORD),
        ("dwFlags",     wt.DWORD),
        ("time",        wt.DWORD),
        ("dwExtraInfo", _PUL),
    ]


class _InputUnion(ctypes.Union):
    _fields_ = [("ki", _KeyBdInput), ("_pad", ctypes.c_byte * 28)]


class _INPUT(ctypes.Structure):
    _fields_ = [("type", wt.DWORD), ("ii", _InputUnion)]


_VK_MENU    = 0x12   # Alt
_VK_SHIFT   = 0x10


def _send_ctrl_c() -> None:
    # まず修飾キー (Ctrl, Alt, Shift) をすべてリリースする。
    # ホットキー Ctrl+Alt+R の直後はこれらが押下状態のままなので、
    # そのまま Ctrl+C を送ると Ctrl+Alt+C になりコピーが効かない。
    release = (_INPUT * 3)(
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_MENU,    0, _KEYEVENTF_KEYUP, 0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_SHIFT,   0, _KEYEVENTF_KEYUP, 0, None))),
    )
    ctypes.windll.user32.SendInput(3, release, ctypes.sizeof(_INPUT))
    time.sleep(0.05)

    # Ctrl+C を送信
    inputs = (_INPUT * 4)(
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_CONTROL, 0, 0,               0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_C,       0, 0,               0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_C,       0, _KEYEVENTF_KEYUP, 0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0, None))),
    )
    ctypes.windll.user32.SendInput(4, inputs, ctypes.sizeof(_INPUT))


# =====================================================================
# TTSEngine — Popen + terminate() で即座に中断可能
# =====================================================================
class TTSEngine:
    def __init__(self):
        self._queue: queue.Queue[str | None] = queue.Queue()
        self._proc: subprocess.Popen | None = None
        self._proc_lock = threading.Lock()
        self._speaking = threading.Event()  # 読み上げ中かどうか
        self._done = threading.Event()      # 読み上げ完了待ち用
        self._done.set()                    # 初期状態は「完了」
        self._thread = threading.Thread(target=self._worker, daemon=True)
        self._thread.start()

    @property
    def is_speaking(self) -> bool:
        return self._speaking.is_set()

    def speak(self, text: str) -> None:
        if not text.strip():
            return
        # キュー内の古いアイテムを破棄
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        self._done.clear()  # 「未完了」にセット
        self._queue.put(text)

    def wait_done(self, timeout: float | None = None) -> bool:
        """読み上げ完了（または中断）まで待機。True=完了、False=タイムアウト"""
        return self._done.wait(timeout=timeout)

    def stop_current(self) -> None:
        """即座に読み上げを中断する"""
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        with self._proc_lock:
            if self._proc and self._proc.poll() is None:
                try:
                    self._proc.terminate()
                    log.info("TTS プロセスを終了しました")
                except Exception:
                    pass
        self._speaking.clear()

    def stop(self) -> None:
        self.stop_current()
        self._queue.put(None)

    def _worker(self) -> None:
        while True:
            text = self._queue.get()
            if text is None:
                break
            self._speak(text)

    def _speak(self, text: str) -> None:
        safe = text.replace("'", "''")
        ps_script = (
            "Add-Type -AssemblyName System.Speech; "
            "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer; "
            "$s.SelectVoice('Microsoft Haruka Desktop'); "
            f"$s.Speak('{safe}')"
        )
        log.info(f"TTS 開始 ({len(text)} 文字)")
        self._speaking.set()
        try:
            proc = subprocess.Popen(
                ["powershell", "-NoProfile", "-Command", ps_script],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=_PS_FLAGS,
            )
            with self._proc_lock:
                self._proc = proc
            proc.wait(timeout=180)
            rc = proc.returncode
            if rc == 0:
                log.info("TTS 完了")
            else:
                stderr = proc.stderr.read()[:300] if proc.stderr else ""
                log.error(f"TTS 異常終了 rc={rc} {stderr}")
        except subprocess.TimeoutExpired:
            proc.terminate()
            log.error("TTS タイムアウト (180秒)")
        except Exception as e:
            log.error(f"TTS エラー: {e}")
        finally:
            with self._proc_lock:
                self._proc = None
            self._speaking.clear()
            self._done.set()  # 完了を通知


# =====================================================================
# ClipboardManager — クリア → Ctrl+C → リトライ付き読み取り
# =====================================================================
class ClipboardManager:

    def get_selected_text(self) -> str:
        """選択テキストを取得する。失敗時は空文字列を返す。"""
        original = self._read()

        # クリップボードをクリアしてから Ctrl+C
        self._clear()
        _send_ctrl_c()
        time.sleep(_CLIP_WAIT_1)

        text = self._read()
        if not text.strip():
            # リトライ: アプリによっては遅い
            time.sleep(_CLIP_WAIT_2)
            text = self._read()

        # 元のクリップボードを復元（新テキストが取れた場合のみ）
        if text.strip():
            # 取得成功 → 元のクリップボードを復元
            if original and original != text:
                try:
                    pyperclip.copy(original)
                except Exception:
                    pass
        else:
            # 取得失敗 → 元のクリップボードを復元
            if original:
                try:
                    pyperclip.copy(original)
                except Exception:
                    pass

        return text.strip()

    def _read(self) -> str:
        try:
            return pyperclip.paste() or ""
        except Exception:
            return ""

    def _clear(self) -> None:
        try:
            pyperclip.copy("")
        except Exception:
            pass


# =====================================================================
# HotkeyHandler — Ctrl+Alt+R + Esc (読み上げ中のみ動的に登録/解除)
# =====================================================================
_HOTKEY_ESC = 2
_MSG_REGISTER_ESC   = _WM_USER + 1
_MSG_UNREGISTER_ESC = _WM_USER + 2


class HotkeyHandler:
    def __init__(self, tts: TTSEngine, clipboard: ClipboardManager):
        self._tts = tts
        self._clipboard = clipboard
        self._lock = threading.Lock()
        self._thread_id: int | None = None
        self._esc_registered = False

    def register(self) -> None:
        t = threading.Thread(target=self._hotkey_loop, daemon=True)
        t.start()

    def _register_esc(self) -> None:
        """ホットキースレッドに Esc 登録を依頼"""
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(
                self._thread_id, _MSG_REGISTER_ESC, 0, 0)

    def _unregister_esc(self) -> None:
        """ホットキースレッドに Esc 解除を依頼"""
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(
                self._thread_id, _MSG_UNREGISTER_ESC, 0, 0)

    def _hotkey_loop(self) -> None:
        user32 = ctypes.windll.user32
        self._thread_id = ctypes.windll.kernel32.GetCurrentThreadId()

        ok = user32.RegisterHotKey(None, _HOTKEY_READ, _MOD_CTRL_ALT, _VK_R)
        if not ok:
            log.error("RegisterHotKey(Ctrl+Alt+R) 失敗")
            return
        log.info("Ctrl+Alt+R を登録しました")

        msg = wt.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == _WM_HOTKEY:
                if msg.wParam == _HOTKEY_READ:
                    if self._tts.is_speaking:
                        log.info("Ctrl+Alt+R → 読み上げ停止")
                        self._tts.stop_current()
                        self._do_unregister_esc(user32)
                    else:
                        log.info("Ctrl+Alt+R → 読み上げ開始")
                        t = threading.Thread(
                            target=self._speak_selected_text, daemon=True)
                        t.start()
                elif msg.wParam == _HOTKEY_ESC:
                    log.info("Esc → 読み上げ停止")
                    self._tts.stop_current()
                    self._do_unregister_esc(user32)

            elif msg.message == _MSG_REGISTER_ESC:
                self._do_register_esc(user32)
            elif msg.message == _MSG_UNREGISTER_ESC:
                self._do_unregister_esc(user32)

        user32.UnregisterHotKey(None, _HOTKEY_READ)
        if self._esc_registered:
            user32.UnregisterHotKey(None, _HOTKEY_ESC)

    def _do_register_esc(self, user32) -> None:
        if not self._esc_registered:
            if user32.RegisterHotKey(None, _HOTKEY_ESC, 0, _VK_ESCAPE):
                self._esc_registered = True
                log.info("Esc ホットキー登録（読み上げ中）")
            else:
                log.error("Esc ホットキー登録失敗")

    def _do_unregister_esc(self, user32) -> None:
        if self._esc_registered:
            user32.UnregisterHotKey(None, _HOTKEY_ESC)
            self._esc_registered = False
            log.info("Esc ホットキー解除")

    def _speak_selected_text(self) -> None:
        if not self._lock.acquire(blocking=False):
            return
        try:
            # ユーザーが Ctrl+Alt+R のキーを離すのを待つ
            time.sleep(0.3)
            text = self._clipboard.get_selected_text()
            if text:
                log.info(f"読み上げ: {text[:60]}{'...' if len(text) > 60 else ''}")
                self._register_esc()
                self._tts.speak(text)
                self._tts.wait_done()  # 読み上げ完了 or 中断まで待機
                self._unregister_esc()
            else:
                log.info("テキストが取得できませんでした")
        finally:
            self._lock.release()


# =====================================================================
# TrayApp
# =====================================================================
class TrayApp:
    def __init__(self, tts: TTSEngine, hotkey: HotkeyHandler):
        self._tts = tts
        self._hotkey = hotkey
        img = self._create_icon_image()
        menu = pystray.Menu(
            pystray.MenuItem("停止", self._on_stop),
            pystray.MenuItem("終了", self._on_quit),
        )
        self._icon = pystray.Icon(
            "yomiage", img, "読み上げ  Ctrl+Alt+R  /  Esc で停止", menu,
        )

    def _create_icon_image(self) -> Image.Image:
        size = 64
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.ellipse([2, 2, size - 2, size - 2], fill=(30, 120, 220, 255))
        text = "読"
        font = None
        for name in ("msgothic.ttc", "meiryo.ttc", "yumin.ttf"):
            try:
                font = ImageFont.truetype(name, 34)
                break
            except Exception:
                pass
        if font is None:
            font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text(
            ((size - tw) / 2 - bbox[0], (size - th) / 2 - bbox[1] - 2),
            text, font=font, fill="white",
        )
        return img

    def _on_stop(self, icon, item) -> None:
        self._tts.stop_current()
        log.info("トレイメニュー → 停止")

    def _on_quit(self, icon, item) -> None:
        self._tts.stop()
        icon.stop()

    def run(self) -> None:
        self._hotkey.register()
        log.info("起動完了。Ctrl+Alt+R で読み上げ、Esc で停止。")
        self._icon.run()


# =====================================================================
# Entry point
# =====================================================================
if __name__ == "__main__":
    try:
        log.info(f"=== yomiage 起動 (Python {sys.version}) ===")
        log.info(f"スクリプト: {__file__}")
        tts = TTSEngine()
        clipboard = ClipboardManager()
        hotkey = HotkeyHandler(tts, clipboard)
        app = TrayApp(tts, hotkey)
        app.run()
    except Exception:
        log.exception("致命的エラーで終了しました")
        sys.exit(1)
