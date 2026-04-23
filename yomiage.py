"""
yomiage.py — 選択テキスト読み上げアプリ
テキストを選択して Ctrl+Alt+R を押すと Windows TTS (日本語) で読み上げる。
読み上げ中に Esc を押すと即座に中断できる。
タスクバーのトレイアイコンとして常駐する。

依存パッケージ:
  pip install pystray pyperclip pillow

ホットキー:
  Ctrl+Alt+R : 選択テキストを読み上げ（読み上げ中なら停止）
  Ctrl+Alt+E : テキストカーソル位置から文末まで読み上げ
  Esc        : 読み上げ中のみ有効。読み上げを即座に停止。
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import io
import logging
import math
import os
import queue
import struct
import subprocess
import sys
import threading
import time
import tkinter as tk
import wave
import winsound
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
_KEYEVENTF_KEYUP       = 0x0002
_KEYEVENTF_EXTENDEDKEY = 0x0001
_VK_CONTROL       = 0x11
_VK_C             = 0x43
_VK_R             = 0x52
_VK_E             = 0x45
_VK_END           = 0x23
_VK_ESCAPE        = 0x1B

_MOD_ALT          = 0x0001
_MOD_CONTROL      = 0x0002
_MOD_CTRL_ALT     = _MOD_ALT | _MOD_CONTROL

_WM_HOTKEY        = 0x0312
_WM_USER          = 0x0400

_HOTKEY_READ      = 1
_HOTKEY_READ_END  = 3   # Ctrl+Alt+E
_HOTKEY_PAUSE     = 4   # Tab（読み上げ中のみ）

_PS_FLAGS = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW


# =====================================================================
# ビープ音 WAV データ（メモリ上で生成）
# =====================================================================
def _make_beep_wav(freq: int = 800, duration_ms: int = 150,
                   volume: float = 0.15, sample_rate: int = 44100) -> bytes:
    """フェードイン/アウト付きサイン波 WAV を bytes で返す（柔らかい音）"""
    n_samples = int(sample_rate * duration_ms / 1000)
    fade = int(n_samples * 0.25)   # 前後25%をフェード
    buf = io.BytesIO()
    with wave.open(buf, "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)  # 16-bit
        wf.setframerate(sample_rate)
        for i in range(n_samples):
            # フェードエンベロープ
            if i < fade:
                env = i / fade
            elif i > n_samples - fade:
                env = (n_samples - i) / fade
            else:
                env = 1.0
            val = int(32767 * volume * env * math.sin(2 * math.pi * freq * i / sample_rate))
            wf.writeframes(struct.pack("<h", val))
    return buf.getvalue()


_BEEP_WAV = _make_beep_wav(800, 60, 0.15)   # 800Hz, 60ms, 音量15%, フェード付き


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
_VK_RIGHT   = 0x27   # →（拡張キー）
_VK_TAB     = 0x09
_VK_F_KEY   = 0x46   # F
_VK_A_KEY   = 0x41   # A
_VK_V_KEY   = 0x56   # V
_VK_RETURN  = 0x0D


def _release_modifiers() -> None:
    """修飾キー (Ctrl/Alt/Shift) をすべてリリースする"""
    release = (_INPUT * 3)(
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_MENU,    0, _KEYEVENTF_KEYUP, 0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_SHIFT,   0, _KEYEVENTF_KEYUP, 0, None))),
    )
    ctypes.windll.user32.SendInput(3, release, ctypes.sizeof(_INPUT))
    time.sleep(0.05)


def _send_ctrl_c() -> None:
    """修飾キーをリリースしてから Ctrl+C を送信"""
    _release_modifiers()
    inputs = (_INPUT * 4)(
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_CONTROL, 0, 0,                0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_C,       0, 0,                0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_C,       0, _KEYEVENTF_KEYUP, 0, None))),
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(_VK_CONTROL, 0, _KEYEVENTF_KEYUP, 0, None))),
    )
    ctypes.windll.user32.SendInput(4, inputs, ctypes.sizeof(_INPUT))


def _send_one_key(vk: int, flags: int = 0) -> None:
    inp = (_INPUT * 1)(
        _INPUT(_INPUT_KEYBOARD, _InputUnion(ki=_KeyBdInput(vk, 0, flags, 0, None))),
    )
    ctypes.windll.user32.SendInput(1, inp, ctypes.sizeof(_INPUT))


def _send_ctrl_shift_end_then_copy() -> None:
    """Ctrl+Shift+End で選択 → Ctrl+C でコピー（End は拡張キーとして送信）"""
    _release_modifiers()

    # Step1: Ctrl+Shift+End（End は KEYEVENTF_EXTENDEDKEY が必要）
    _send_one_key(_VK_CONTROL, 0)
    time.sleep(0.02)
    _send_one_key(_VK_SHIFT, 0)
    time.sleep(0.02)
    _send_one_key(_VK_END, _KEYEVENTF_EXTENDEDKEY)                       # End down (拡張)
    time.sleep(0.02)
    _send_one_key(_VK_END, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP)    # End up (拡張)
    time.sleep(0.02)
    _send_one_key(_VK_SHIFT, _KEYEVENTF_KEYUP)
    time.sleep(0.02)
    _send_one_key(_VK_CONTROL, _KEYEVENTF_KEYUP)

    # 選択確定を待つ
    time.sleep(0.25)

    # Step2: Ctrl+C でコピー
    log.info("Ctrl+C 送信開始")
    _send_one_key(_VK_CONTROL, 0)
    time.sleep(0.02)
    _send_one_key(_VK_C, 0)
    time.sleep(0.02)
    _send_one_key(_VK_C, _KEYEVENTF_KEYUP)
    time.sleep(0.02)
    _send_one_key(_VK_CONTROL, _KEYEVENTF_KEYUP)



# =====================================================================
# チャンクスクロール — ブラウザ限定で Ctrl+F によるスクロール表示
# =====================================================================
def _is_chromium_browser(hwnd: int) -> bool:
    """Edge / Chrome など Chromium 系ブラウザのウィンドウかどうかを判定"""
    buf = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
    return "Chrome_WidgetWin" in buf.value


def _scroll_to_chunk(chunk_text: str, hwnd: int) -> None:
    """Ctrl+F でチャンク先頭へスクロール（Edge/Chrome/Word/メモ帳 共通）。
    Ctrl+A は使わない（Word で文書全体選択になる恐れがあるため）。
    Ctrl+F を開くと検索ボックスがフォーカスされ前回語が選択状態になるので
    Ctrl+V で上書き貼り付けするだけで安全に動作する。"""
    if not hwnd or not chunk_text.strip():
        return
    try:
        # フォーカスを対象ウィンドウに戻す
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.15)

        # 検索キーワード: 最初の 20 文字
        keyword = chunk_text.strip()[:20]
        if not keyword:
            return
        pyperclip.copy(keyword)
        time.sleep(0.05)

        # Ctrl+F — 検索バー/ダイアログを開く
        _send_one_key(_VK_CONTROL, 0);  time.sleep(0.02)
        _send_one_key(_VK_F_KEY, 0);    time.sleep(0.02)
        _send_one_key(_VK_F_KEY, _KEYEVENTF_KEYUP); time.sleep(0.02)
        _send_one_key(_VK_CONTROL, _KEYEVENTF_KEYUP)
        time.sleep(0.5)    # 検索バーが開いてフォーカスが移るのを待つ

        # Ctrl+V — 貼り付け（検索ボックスに前回語が選択されていれば上書き）
        _send_one_key(_VK_CONTROL, 0);  time.sleep(0.02)
        _send_one_key(_VK_V_KEY, 0);    time.sleep(0.02)
        _send_one_key(_VK_V_KEY, _KEYEVENTF_KEYUP); time.sleep(0.02)
        _send_one_key(_VK_CONTROL, _KEYEVENTF_KEYUP)
        time.sleep(0.15)

        # Enter — 検索実行（ページ/文書がスクロール）
        _send_one_key(_VK_RETURN, 0);   time.sleep(0.05)
        _send_one_key(_VK_RETURN, _KEYEVENTF_KEYUP)
        time.sleep(0.25)

        # Escape — 検索バー/ダイアログを閉じる（スクロール位置は維持）
        _send_one_key(_VK_ESCAPE, 0);   time.sleep(0.05)
        _send_one_key(_VK_ESCAPE, _KEYEVENTF_KEYUP)
        time.sleep(0.1)

    except Exception as e:
        log.warning(f"チャンクスクロール失敗: {e}")


# =====================================================================
# OverlayWindow — 読み上げ中のテキストを画面下部に表示
# =====================================================================
class OverlayWindow:
    """読み上げ中のチャンクを画面下部に表示する常に最前面のオーバーレイ"""

    def __init__(self):
        self._queue: queue.Queue = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self):
        root = tk.Tk()
        root.withdraw()
        root.overrideredirect(True)       # タイトルバーなし
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.88)
        root.configure(bg="#1a1a1a")

        # 画面下部中央に配置
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        w, h = min(1000, screen_w - 100), 120
        x = (screen_w - w) // 2
        y = screen_h - h - 80
        root.geometry(f"{w}x{h}+{x}+{y}")

        # タイトル行
        title_lbl = tk.Label(
            root, text="", font=("Meiryo", 10),
            fg="#88ccff", bg="#1a1a1a", anchor="w",
        )
        title_lbl.pack(fill="x", padx=15, pady=(8, 0))

        # 本文
        text_lbl = tk.Label(
            root, text="", font=("Meiryo", 16, "bold"),
            fg="#ffffff", bg="#1a1a1a", wraplength=w - 30,
            justify="left", anchor="w",
        )
        text_lbl.pack(expand=True, fill="both", padx=15, pady=(4, 10))

        def poll():
            try:
                while True:
                    msg = self._queue.get_nowait()
                    if msg is None:
                        root.destroy()
                        return
                    action, a, b = msg
                    if action == "show":
                        title_lbl.config(text=a)
                        text_lbl.config(text=b)
                        root.deiconify()
                        root.lift()
                    elif action == "update":
                        title_lbl.config(text=a)
                        text_lbl.config(text=b)
                    elif action == "hide":
                        root.withdraw()
            except queue.Empty:
                pass
            root.after(50, poll)

        root.after(50, poll)
        root.mainloop()

    def show(self, title: str, text: str) -> None:
        self._queue.put(("show", title, text))

    def update(self, title: str, text: str) -> None:
        self._queue.put(("update", title, text))

    def hide(self) -> None:
        self._queue.put(("hide", "", ""))


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


import re as _re

# 文の区切りパターン（日本語の句点・感嘆符・疑問符・改行など）
_SENTENCE_SPLIT = _re.compile(r'(?<=[。！？!?\n])\s*')

# 1チャンクの最大文字数（これ以上は強制分割）
_CHUNK_MAX_CHARS = 120


def _split_into_chunks(text: str) -> list[str]:
    """テキストを文単位のチャンクに分割する"""
    # まず文で分割
    sentences = _SENTENCE_SPLIT.split(text.strip())
    sentences = [s.strip() for s in sentences if s.strip()]

    if not sentences:
        return [text.strip()] if text.strip() else []

    # 短い文はまとめる、長い文は分割
    chunks: list[str] = []
    current = ""
    for s in sentences:
        if len(s) > _CHUNK_MAX_CHARS:
            # 長い文: 現在のバッファを先にフラッシュ
            if current:
                chunks.append(current)
                current = ""
            # 句読点なしの長文は読点で分割を試みる
            parts = _re.split(r'(?<=[、,，])\s*', s)
            sub = ""
            for p in parts:
                if len(sub) + len(p) > _CHUNK_MAX_CHARS and sub:
                    chunks.append(sub)
                    sub = p
                else:
                    sub += p
            if sub:
                chunks.append(sub)
        elif len(current) + len(s) > _CHUNK_MAX_CHARS:
            # バッファがいっぱい → フラッシュ
            if current:
                chunks.append(current)
            current = s
        else:
            current += s
    if current:
        chunks.append(current)

    return chunks if chunks else [text.strip()]


class TTSEngine:
    def __init__(self, overlay: OverlayWindow | None = None):
        self._overlay = overlay
        self._queue: queue.Queue[str | None] = queue.Queue()
        self._gen_proc: subprocess.Popen | None = None
        self._proc_lock = threading.Lock()
        self._speaking = threading.Event()
        self._stop_flag = threading.Event()
        self._pause_flag = threading.Event()   # セットされたら一時停止
        self._done = threading.Event()
        self._done.set()
        self._source_hwnd: int = 0             # Ctrl+Alt+E 発火時のウィンドウ
        self._current_chunk: str = ""          # 一時停止時の表示用
        self._python_exe: str | None = None
        self._thread = threading.Thread(target=self._worker, daemon=True)
        self._thread.start()

    def _get_python_exe(self) -> str:
        if self._python_exe is None:
            exe = sys.executable or (
                r"C:\Users\arai\AppData\Local\Programs\Python\Python313\pythonw.exe"
            )
            self._python_exe = exe.replace("pythonw.exe", "python.exe")
        return self._python_exe

    @property
    def is_speaking(self) -> bool:
        return self._speaking.is_set()

    def speak(self, text: str, source_hwnd: int = 0) -> None:
        if not text.strip():
            return
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        self._source_hwnd = source_hwnd
        self._pause_flag.clear()
        self._done.clear()
        self._queue.put(text)

    def pause(self) -> None:
        """Tab キーによる一時停止"""
        if self._speaking.is_set() and not self._pause_flag.is_set():
            _mci("pause tts")
            self._pause_flag.set()
            if self._overlay is not None:
                self._overlay.update("⏸ 一時停止中 — Tab で再開", self._current_chunk)
            log.info("TTS 一時停止")

    def resume(self) -> None:
        """Tab キーによる再開"""
        if self._speaking.is_set() and self._pause_flag.is_set():
            _mci("resume tts")
            self._pause_flag.clear()
            if self._overlay is not None:
                self._overlay.update("▶ 読み上げ再開", self._current_chunk)
            log.info("TTS 再開")

    def wait_done(self, timeout: float | None = None) -> bool:
        return self._done.wait(timeout=timeout)

    def stop_current(self) -> None:
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        self._stop_flag.set()
        self._pause_flag.clear()   # 一時停止も解除
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

    def _generate_mp3(self, text: str, mp3_path: Path) -> bool:
        """edge-tts で MP3 を生成。成功なら True"""
        proc = subprocess.Popen(
            [
                self._get_python_exe(), "-m", "edge_tts",
                "--voice", _TTS_VOICE,
                "--rate", _TTS_RATE,
                "--text", text,
                "--write-media", str(mp3_path),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=_PS_FLAGS,
        )
        with self._proc_lock:
            self._gen_proc = proc
        try:
            proc.wait(timeout=60)
        except subprocess.TimeoutExpired:
            proc.terminate()
            return False
        finally:
            with self._proc_lock:
                self._gen_proc = None
        if proc.returncode != 0:
            stderr = proc.stderr.read()[:300] if proc.stderr else ""
            log.error(f"edge-tts 生成失敗 rc={proc.returncode}: {stderr}")
            return False
        return mp3_path.exists()

    def _play_mp3(self, mp3_path: Path) -> None:
        """MCI で MP3 を再生し、完了 or stop_flag まで待つ"""
        mp3_path_str = str(mp3_path)
        rc = _mci(f'open "{mp3_path_str}" type mpegvideo alias tts')
        if rc != 0:
            log.error(f"MCI open 失敗 rc={rc}")
            return
        rc = _mci("play tts")
        if rc != 0:
            log.error(f"MCI play 失敗 rc={rc}")
            _mci("close tts")
            return
        # 再生完了をポーリングで待つ（一時停止中はスリープのみ）
        while not self._stop_flag.is_set():
            if self._pause_flag.is_set():
                time.sleep(0.05)
                continue
            status = _mci_status("status tts mode")
            if status not in ("playing", "paused"):
                break
            time.sleep(0.1)
        _mci("stop tts")
        _mci("close tts")

    def _speak(self, text: str) -> None:
        chunks = _split_into_chunks(text)
        log.info(f"TTS 開始 ({len(text)} 文字, {len(chunks)} チャンク, {_TTS_VOICE})")
        self._speaking.set()
        self._stop_flag.clear()

        tmp_dir = Path(__file__).parent
        generated_files: list[Path] = []

        try:
            for i, chunk in enumerate(chunks):
                if self._stop_flag.is_set():
                    log.info("TTS 中断")
                    break

                mp3_file = tmp_dir / f"_tts_chunk_{i}.mp3"
                generated_files.append(mp3_file)

                # ---- 先読み生成スレッド（次のチャンクを並行生成） ----
                next_mp3: Path | None = None
                prefetch_thread: threading.Thread | None = None
                prefetch_ok: list[bool] = [False]

                if i + 1 < len(chunks):
                    next_mp3 = tmp_dir / f"_tts_chunk_{i + 1}.mp3"
                    generated_files.append(next_mp3)

                    def _prefetch(txt=chunks[i + 1], path=next_mp3):
                        prefetch_ok[0] = self._generate_mp3(txt, path)

                # ---- 現在のチャンクを生成 ----
                if i == 0:
                    # 最初のチャンクだけビープ音付き
                    beep_stop = threading.Event()

                    def _beep_loop():
                        while not beep_stop.is_set():
                            try:
                                winsound.PlaySound(
                                    _BEEP_WAV,
                                    winsound.SND_MEMORY | winsound.SND_NOSTOP,
                                )
                            except Exception:
                                pass
                            beep_stop.wait(0.8)

                    beep_thread = threading.Thread(target=_beep_loop, daemon=True)
                    beep_thread.start()
                    ok = self._generate_mp3(chunk, mp3_file)
                    beep_stop.set()
                    beep_thread.join(timeout=1)
                else:
                    ok = self._generate_mp3(chunk, mp3_file)

                if not ok or self._stop_flag.is_set():
                    break

                log.info(f"チャンク {i+1}/{len(chunks)} 再生 ({len(chunk)} 文字)")

                # Ctrl+Alt+E の場合のみ: 読み上げ位置にスクロール＆ハイライト
                if self._source_hwnd:
                    _scroll_to_chunk(chunk, self._source_hwnd)

                # オーバーレイを更新（現在のチャンクを表示）
                self._current_chunk = chunk
                if self._overlay is not None:
                    title = f"読み上げ中 {i+1}/{len(chunks)}"
                    if i == 0:
                        self._overlay.show(title, chunk)
                    else:
                        self._overlay.update(title, chunk)

                # 再生開始と同時に次のチャンクを先読み生成
                if next_mp3 is not None:
                    prefetch_thread = threading.Thread(target=_prefetch, daemon=True)
                    prefetch_thread.start()

                self._play_mp3(mp3_file)

                # 先読み完了を待つ
                if prefetch_thread is not None:
                    prefetch_thread.join(timeout=60)

                if self._stop_flag.is_set():
                    log.info("TTS 中断（Esc/停止）")
                    break
            else:
                log.info("TTS 完了")

        except Exception as e:
            log.error(f"TTS エラー: {e}")
        finally:
            with self._proc_lock:
                self._gen_proc = None
            self._speaking.clear()
            self._done.set()
            # オーバーレイを隠す
            if self._overlay is not None:
                self._overlay.hide()
            # 一時ファイル削除
            for f in generated_files:
                try:
                    f.unlink(missing_ok=True)
                except Exception:
                    pass


# =====================================================================
# ClipboardManager — クリア → Ctrl+C → リトライ付き読み取り
# =====================================================================
class ClipboardManager:

    def _copy_and_read(self, send_fn) -> str:
        """send_fn でキー送信 → クリップボード読み取り → 元のクリップボードを復元"""
        original = self._read()
        self._clear()
        send_fn()
        time.sleep(_CLIP_WAIT_1)
        text = self._read()
        if not text.strip():
            time.sleep(_CLIP_WAIT_2)
            text = self._read()
        # 元のクリップボードを復元
        if original and original != text:
            try:
                pyperclip.copy(original)
            except Exception:
                pass
        return text.strip()

    def get_selected_text(self) -> str:
        """選択テキストを取得する。失敗時は空文字列を返す。"""
        return self._copy_and_read(_send_ctrl_c)

    def get_text_from_cursor(self) -> str:
        """テキストカーソル位置から文末までを取得する。失敗時は空文字列を返す。"""
        original = self._read()
        self._clear()
        _send_ctrl_shift_end_then_copy()
        # 選択+コピー後はクリップボード更新が遅いアプリがあるため長めに待つ
        time.sleep(0.25)
        text = self._read()
        if not text.strip():
            time.sleep(0.3)
            text = self._read()
        # 元のクリップボードを復元
        if original and original != text:
            try:
                pyperclip.copy(original)
            except Exception:
                pass
        log.info(f"カーソルから取得: {len(text)} 文字")
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
_MSG_REGISTER_ESC    = _WM_USER + 1
_MSG_UNREGISTER_ESC  = _WM_USER + 2
_MSG_REGISTER_TAB    = _WM_USER + 3
_MSG_UNREGISTER_TAB  = _WM_USER + 4


class HotkeyHandler:
    def __init__(self, tts: TTSEngine, clipboard: ClipboardManager):
        self._tts = tts
        self._clipboard = clipboard
        self._lock = threading.Lock()
        self._thread_id: int | None = None
        self._esc_registered = False
        self._tab_registered = False

    def register(self) -> None:
        t = threading.Thread(target=self._hotkey_loop, daemon=True)
        t.start()

    def _register_esc(self) -> None:
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(
                self._thread_id, _MSG_REGISTER_ESC, 0, 0)

    def _unregister_esc(self) -> None:
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(
                self._thread_id, _MSG_UNREGISTER_ESC, 0, 0)

    def _register_tab(self) -> None:
        """ホットキースレッドに Tab 登録を依頼"""
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(
                self._thread_id, _MSG_REGISTER_TAB, 0, 0)

    def _unregister_tab(self) -> None:
        """ホットキースレッドに Tab 解除を依頼"""
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(
                self._thread_id, _MSG_UNREGISTER_TAB, 0, 0)

    def _hotkey_loop(self) -> None:
        user32 = ctypes.windll.user32
        self._thread_id = ctypes.windll.kernel32.GetCurrentThreadId()

        ok = user32.RegisterHotKey(None, _HOTKEY_READ, _MOD_CTRL_ALT, _VK_R)
        if not ok:
            log.error("RegisterHotKey(Ctrl+Alt+R) 失敗")
            return
        log.info("Ctrl+Alt+R を登録しました")

        ok2 = user32.RegisterHotKey(None, _HOTKEY_READ_END, _MOD_CTRL_ALT, _VK_E)
        if ok2:
            log.info("Ctrl+Alt+E を登録しました（カーソル位置から末尾まで読み上げ）")
        else:
            log.warning("RegisterHotKey(Ctrl+Alt+E) 失敗（他アプリが使用中の可能性）")

        msg = wt.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == _WM_HOTKEY:
                if msg.wParam == _HOTKEY_READ:
                    if self._tts.is_speaking:
                        log.info("Ctrl+Alt+R → 読み上げ停止")
                        self._tts.stop_current()
                        self._do_unregister_esc(user32)
                    else:
                        log.info("Ctrl+Alt+R → 読み上げ開始（選択テキスト）")
                        t = threading.Thread(
                            target=self._speak_selected_text, daemon=True)
                        t.start()
                elif msg.wParam == _HOTKEY_READ_END:
                    if self._tts.is_speaking:
                        log.info("Ctrl+Alt+E → 読み上げ停止")
                        self._tts.stop_current()
                        self._do_unregister_esc(user32)
                    else:
                        hwnd = user32.GetForegroundWindow()
                        log.info(f"Ctrl+Alt+E → 読み上げ開始（カーソル位置から末尾, hwnd={hwnd}）")
                        t = threading.Thread(
                            target=self._speak_from_cursor, args=(hwnd,), daemon=True)
                        t.start()
                elif msg.wParam == _HOTKEY_ESC:
                    log.info("Esc → 読み上げ停止")
                    self._tts.stop_current()
                    self._do_unregister_esc(user32)
                    self._do_unregister_tab(user32)
                elif msg.wParam == _HOTKEY_PAUSE:
                    if self._tts._pause_flag.is_set():
                        log.info("Tab → 再開")
                        self._tts.resume()
                    else:
                        log.info("Tab → 一時停止")
                        self._tts.pause()

            elif msg.message == _MSG_REGISTER_ESC:
                self._do_register_esc(user32)
            elif msg.message == _MSG_UNREGISTER_ESC:
                self._do_unregister_esc(user32)
            elif msg.message == _MSG_REGISTER_TAB:
                self._do_register_tab(user32)
            elif msg.message == _MSG_UNREGISTER_TAB:
                self._do_unregister_tab(user32)

        user32.UnregisterHotKey(None, _HOTKEY_READ)
        user32.UnregisterHotKey(None, _HOTKEY_READ_END)
        if self._esc_registered:
            user32.UnregisterHotKey(None, _HOTKEY_ESC)
        if self._tab_registered:
            user32.UnregisterHotKey(None, _HOTKEY_PAUSE)

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

    def _do_register_tab(self, user32) -> None:
        if not self._tab_registered:
            if user32.RegisterHotKey(None, _HOTKEY_PAUSE, 0, _VK_TAB):
                self._tab_registered = True
                log.info("Tab ホットキー登録（一時停止/再開）")
            else:
                log.warning("Tab ホットキー登録失敗（他アプリが使用中の可能性）")

    def _do_unregister_tab(self, user32) -> None:
        if self._tab_registered:
            user32.UnregisterHotKey(None, _HOTKEY_PAUSE)
            self._tab_registered = False
            log.info("Tab ホットキー解除")

    def _speak_selected_text(self) -> None:
        if not self._lock.acquire(blocking=False):
            return
        try:
            time.sleep(0.3)  # キーを離すのを待つ
            text = self._clipboard.get_selected_text()
            if text:
                log.info(f"読み上げ: {text[:60]}{'...' if len(text) > 60 else ''}")
                self._register_esc()
                self._register_tab()
                self._tts.speak(text)
                self._tts.wait_done()
                self._unregister_esc()
                self._unregister_tab()
            else:
                log.info("テキストが取得できませんでした")
        finally:
            self._lock.release()

    def _speak_from_cursor(self, hwnd: int = 0) -> None:
        if not self._lock.acquire(blocking=False):
            return
        try:
            time.sleep(0.4)  # キーを離すのを待つ
            # ホットキー発火時のウィンドウに確実にフォーカスを戻す
            if hwnd:
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                time.sleep(0.1)
            text = self._clipboard.get_text_from_cursor()
            if text:
                # クリップボード取得完了後に選択を解除（全反転を消す）
                _release_modifiers()
                time.sleep(0.05)
                _send_one_key(_VK_RIGHT, _KEYEVENTF_EXTENDEDKEY)
                time.sleep(0.05)
                _send_one_key(_VK_RIGHT, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP)
                time.sleep(0.05)
                log.info(f"カーソルから末尾まで読み上げ: {text[:60]}{'...' if len(text) > 60 else ''}")
                self._register_esc()
                self._register_tab()
                self._tts.speak(text, source_hwnd=hwnd)   # hwnd を渡してスクロール有効化
                self._tts.wait_done()
                self._unregister_esc()
                self._unregister_tab()
            else:
                log.info("カーソル位置からテキストが取得できませんでした")
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
            "yomiage", img,
            "読み上げ  Ctrl+Alt+R(選択) / Ctrl+Alt+E(末尾まで) / Tab(一時停止) / Esc(停止)",
            menu,
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
        overlay = OverlayWindow()
        tts = TTSEngine(overlay)
        clipboard = ClipboardManager()
        hotkey = HotkeyHandler(tts, clipboard)
        app = TrayApp(tts, hotkey)
        app.run()
    except Exception:
        log.exception("致命的エラーで終了しました")
        sys.exit(1)
