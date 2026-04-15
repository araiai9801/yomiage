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
# TTSEngine — edge-tts (Microsoft Nanami) + 即座に中断可能
# =====================================================================
# 音声名: ja-JP-NanamiNeural (女性), ja-JP-KeitaNeural (男性)
_TTS_VOICE = "ja-JP-NanamiNeural"
_TTS_RATE  = "+20%"   # 読み上げ速度 (例: "+10%", "+20%", "+50%", "-10%")



# MCI (winmm.dll) を Python から直接呼び出す — PowerShell 不要
_winmm = ctypes.windll.winmm
_mciSendStringW = _winmm.mciSendStringW
_mciSendStringW.argtypes = [
    ctypes.c_wchar_p,       # lpszCommand
    ctypes.c_wchar_p,       # lpszReturnString
    ctypes.c_uint,          # cchReturn
    ctypes.c_void_p,        # hwndCallback
]
_mciSendStringW.restype = ctypes.c_int


def _mci(cmd: str) -> int:
    """MCI コマンドを送信して戻り値を返す"""
    ret = _mciSendStringW(cmd, None, 0, None)
    if ret != 0:
        log.debug(f"MCI '{cmd}' → rc={ret}")
    return ret


def _mci_status(cmd: str) -> str:
    """MCI status コマンドを送信して結果文字列を返す"""
    buf = ctypes.create_unicode_buffer(256)
    _mciSendStringW(cmd, buf, 256, None)
    return buf.value


class TTSEngine:
    def __init__(self):
        self._queue: queue.Queue[str | None] = queue.Queue()
        self._gen_proc: subprocess.Popen | None = None
        self._proc_lock = threading.Lock()
        self._speaking = threading.Event()
        self._stop_flag = threading.Event()
        self._done = threading.Event()
        self._done.set()
        self._thread = threading.Thread(target=self._worker, daemon=True)
        self._thread.start()

    @property
    def is_speaking(self) -> bool:
        return self._speaking.is_set()

    def speak(self, text: str) -> None:
        if not text.strip():
            return
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        self._done.clear()
        self._queue.put(text)

    def wait_done(self, timeout: float | None = None) -> bool:
        return self._done.wait(timeout=timeout)

    def stop_current(self) -> None:
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        self._stop_flag.set()
        # edge-tts 生成中ならプロセスを kill
        with self._proc_lock:
            if self._gen_proc and self._gen_proc.poll() is None:
                try:
                    self._gen_proc.terminate()
                    log.info("edge-tts 生成プロセスを終了しました")
                except Exception:
                    pass
        # MCI 再生中なら即停止
        _mci("stop tts")
        _mci("close tts")
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
        log.info(f"TTS 開始 ({len(text)} 文字, {_TTS_VOICE})")
        self._speaking.set()
        self._stop_flag.clear()
        tmp_mp3 = Path(__file__).parent / "_tts_output.mp3"
        try:
            # ---- 1. edge-tts で MP3 生成 ----
            python_exe = sys.executable or (
                r"C:\Users\arai\AppData\Local\Programs\Python\Python313\pythonw.exe"
            )
            python_exe = python_exe.replace("pythonw.exe", "python.exe")
            gen_proc = subprocess.Popen(
                [
                    python_exe, "-m", "edge_tts",
                    "--voice", _TTS_VOICE,
                    "--rate", _TTS_RATE,
                    "--text", text,
                    "--write-media", str(tmp_mp3),
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                text=True,
                creationflags=_PS_FLAGS,
            )
            with self._proc_lock:
                self._gen_proc = gen_proc
            gen_proc.wait(timeout=60)
            with self._proc_lock:
                self._gen_proc = None

            if self._stop_flag.is_set():
                log.info("TTS 中断（生成中にキャンセル）")
                return
            if gen_proc.returncode != 0:
                stderr = gen_proc.stderr.read()[:300] if gen_proc.stderr else ""
                log.error(f"edge-tts 生成失敗 rc={gen_proc.returncode}: {stderr}")
                return
            if not tmp_mp3.exists():
                log.error("edge-tts: MP3 ファイルが生成されませんでした")
                return

            log.info(f"MP3 生成完了 ({tmp_mp3.stat().st_size} bytes)")

            # ---- 2. Python から直接 MCI で MP3 再生 ----
            mp3_path_str = str(tmp_mp3)
            rc = _mci(f'open "{mp3_path_str}" type mpegvideo alias tts')
            if rc != 0:
                log.error(f"MCI open 失敗 rc={rc}")
                return

            rc = _mci("play tts")
            if rc != 0:
                log.error(f"MCI play 失敗 rc={rc}")
                _mci("close tts")
                return

            # 再生完了をポーリングで待つ（stop_flag で即中断可能）
            while not self._stop_flag.is_set():
                status = _mci_status("status tts mode")
                if status != "playing":
                    break
                time.sleep(0.1)

            _mci("stop tts")
            _mci("close tts")

            if self._stop_flag.is_set():
                log.info("TTS 中断（Esc/停止）")
            else:
                log.info("TTS 完了")

        except subprocess.TimeoutExpired:
            with self._proc_lock:
                if self._gen_proc and self._gen_proc.poll() is None:
                    self._gen_proc.terminate()
                self._gen_proc = None
            log.error("TTS タイムアウト（edge-tts 生成）")
        except Exception as e:
            log.error(f"TTS エラー: {e}")
        finally:
            with self._proc_lock:
                self._gen_proc = None
            self._speaking.clear()
            self._done.set()
            try:
                tmp_mp3.unlink(missing_ok=True)
            except Exception:
                pass


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
