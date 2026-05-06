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
import json
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
import urllib.request
import urllib.error
import wave
import winsound
from pathlib import Path

# tomllib は Python 3.11+ 標準。3.10 以下なら tomli を試行
try:
    import tomllib
except ModuleNotFoundError:
    try:
        import tomli as tomllib  # type: ignore
    except ModuleNotFoundError:
        tomllib = None  # type: ignore

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
_VK_S             = 0x53
_VK_T             = 0x54
_VK_END           = 0x23
_VK_ESCAPE        = 0x1B

_MOD_ALT          = 0x0001
_MOD_CONTROL      = 0x0002
_MOD_SHIFT        = 0x0004
_MOD_CTRL_ALT     = _MOD_ALT | _MOD_CONTROL

_WM_HOTKEY        = 0x0312
_WM_USER          = 0x0400

_HOTKEY_READ      = 1
_HOTKEY_READ_END  = 3   # Ctrl+Alt+E
_HOTKEY_PAUSE     = 4   # Tab（読み上げ中のみ）
_HOTKEY_TOGGLE_POS = 5  # Shift+Tab（読み上げ中のみ・オーバーレイ位置上下切替）
_HOTKEY_SUMMARY   = 6   # Ctrl+Alt+S（生成AIで要約してから読み上げ）

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


def _is_chromium_browser(hwnd: int) -> bool:
    """Edge / Chrome など Chromium 系ブラウザのウィンドウかどうかを判定"""
    buf = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
    return "Chrome_WidgetWin" in buf.value


def _scroll_to_chunk(chunk_text: str, hwnd: int) -> None:
    """Chromium 系ブラウザ限定: Ctrl+F でチャンク先頭へスクロール。
    【重要】Escape を送信しない — SendInput の Escape が RegisterHotKey の
    「Esc=停止」を誤発動させるため。Find バーは開いたままにする。
    Word・メモ帳等ではスキップ（Ctrl+F の挙動が異なり干渉するため）。"""
    if not hwnd or not chunk_text.strip():
        return
    if not _is_chromium_browser(hwnd):
        return   # ブラウザ以外はスキップ
    try:
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        time.sleep(0.15)

        keyword = chunk_text.strip()[:20]
        if not keyword:
            return
        pyperclip.copy(keyword)
        time.sleep(0.05)

        # Ctrl+F — 検索バーを開く（既に開いていれば前回語が選択状態）
        _send_one_key(_VK_CONTROL, 0);  time.sleep(0.02)
        _send_one_key(_VK_F_KEY, 0);    time.sleep(0.02)
        _send_one_key(_VK_F_KEY, _KEYEVENTF_KEYUP); time.sleep(0.02)
        _send_one_key(_VK_CONTROL, _KEYEVENTF_KEYUP)
        time.sleep(0.4)

        # Ctrl+V — 貼り付け
        # Chrome/Edge はテキストを貼り付けた時点でライブ検索・ページスクロールする。
        # Enter を押すと「次のマッチへ移動」になりズレるため送信しない。
        _send_one_key(_VK_CONTROL, 0);  time.sleep(0.02)
        _send_one_key(_VK_V_KEY, 0);    time.sleep(0.02)
        _send_one_key(_VK_V_KEY, _KEYEVENTF_KEYUP); time.sleep(0.02)
        _send_one_key(_VK_CONTROL, _KEYEVENTF_KEYUP)
        time.sleep(0.25)   # スクロールアニメーション完了を待つ
        # ※ Enter・Escape は送信しない

    except Exception as e:
        log.warning(f"チャンクスクロール失敗: {e}")


# =====================================================================
# Config — config.toml を読み込み（無ければデフォルト値で作成）
# =====================================================================
_CONFIG_PATH = Path(__file__).parent / "config.toml"

# config.toml が存在しないときに書き出すデフォルト中身
_DEFAULT_CONFIG_TOML = """\
# =============================================================
# yomiage 設定ファイル
# =============================================================
# このファイルを編集して保存後、yomiage を再起動すると反映されます。
# # で始まる行はコメント（無視されます）。
# =============================================================

[summary]
# -------------------------------------------------------------
# 要約読み上げ機能（Ctrl+Alt+S で呼び出す）
# -------------------------------------------------------------

# 要約モードを有効にするか (true / false)
enabled = true

# 使用する AI プロバイダ
#   "openai"    - OpenAI ChatGPT (推奨・最安・最速)
#   "anthropic" - Anthropic Claude
#   "gemini"    - Google Gemini (無料枠あり)
#
# それぞれ環境変数からAPIキーを読み込みます:
#   openai    → OPENAI_API_KEY
#   anthropic → ANTHROPIC_API_KEY
#   gemini    → GEMINI_API_KEY
provider = "openai"

# モデル名
# OpenAI:    "gpt-4o-mini" (推奨/安価) / "gpt-4o" (高品質)
# Anthropic: "claude-haiku-4-5" / "claude-sonnet-4-5"
# Gemini:    "gemini-2.0-flash" / "gemini-2.5-pro"
model = "gpt-4o-mini"

# 要約の長さ
#   "short"  - 短く（元の30%程度）
#   "medium" - 中程度（50%程度）★推奨
#   "long"   - 詳しめ（70%程度）
#   "bullet" - 要点を箇条書きで列挙
length = "medium"

# 追加の指示（オプション・空文字でもOK）
# 例: "話し言葉で簡潔に" / "ビジネス調で" / "中学生にも分かるように"
extra_instruction = "話し言葉で簡潔に"

# API エラー / APIキー未設定時の動作
#   "fallback" - 元のテキストをそのまま読み上げる ★推奨
#   "skip"     - 何もしない（停止）
on_error = "fallback"

# API のタイムアウト秒数
timeout_s = 30

# 要約処理チャンクの最大文字数
# 長文を 1 回で要約させると AI が極端に短くまとめがち。
# このサイズで区切って AI に渡し、各部分の要約を連結する。
#   小さくする (例: 800)  → 詳細が残る・APIコールが増える
#   大きくする (例: 3000) → コール少ないが圧縮されすぎる
#   推奨: 1500
chunk_chars = 1500
"""


class Config:
    """config.toml を読み込んで設定値を保持する"""

    def __init__(self):
        # デフォルト値
        self.summary_enabled: bool = True
        self.summary_provider: str = "openai"
        self.summary_model: str = "gpt-4o-mini"
        self.summary_length: str = "medium"
        self.summary_extra_instruction: str = "話し言葉で簡潔に"
        self.summary_on_error: str = "fallback"
        self.summary_timeout_s: int = 30
        # 要約処理の単位（文字数）。長文時にこの単位で区切って AI に渡す。
        self.summary_chunk_chars: int = 1500

        # ファイルが無ければ作成
        if not _CONFIG_PATH.exists():
            try:
                _CONFIG_PATH.write_text(_DEFAULT_CONFIG_TOML, encoding="utf-8")
                log.info(f"デフォルト設定ファイルを作成: {_CONFIG_PATH}")
            except Exception as e:
                log.warning(f"config.toml 作成失敗: {e}")

        # 読み込み
        if tomllib is None:
            log.warning("tomllib が利用できません。デフォルト設定を使用します。")
            return
        try:
            with _CONFIG_PATH.open("rb") as f:
                data = tomllib.load(f)
            s = data.get("summary", {})
            self.summary_enabled = bool(s.get("enabled", self.summary_enabled))
            self.summary_provider = str(s.get("provider", self.summary_provider)).lower()
            self.summary_model = str(s.get("model", self.summary_model))
            self.summary_length = str(s.get("length", self.summary_length)).lower()
            self.summary_extra_instruction = str(s.get("extra_instruction", self.summary_extra_instruction))
            self.summary_on_error = str(s.get("on_error", self.summary_on_error)).lower()
            self.summary_timeout_s = int(s.get("timeout_s", self.summary_timeout_s))
            self.summary_chunk_chars = int(s.get("chunk_chars", self.summary_chunk_chars))
            log.info(
                f"設定読み込み完了: provider={self.summary_provider} "
                f"model={self.summary_model} length={self.summary_length} "
                f"chunk_chars={self.summary_chunk_chars}"
            )
        except Exception as e:
            log.warning(f"config.toml 読み込み失敗、デフォルト使用: {e}")


# グローバル設定インスタンス
_CONFIG = Config()


# =====================================================================
# SummaryEngine — 生成 AI でテキストを要約
# =====================================================================
class SummaryEngine:
    """テキストを生成 AI で要約する。プロバイダごとに HTTP で API 呼び出し。"""

    # 長さ別の指示
    _LENGTH_INSTRUCTIONS = {
        "short":  "元の文章を30%程度の長さに圧縮した要約",
        "medium": "元の文章を50%程度の長さに圧縮した要約",
        "long":   "元の文章を70%程度の長さに圧縮した要約",
        "bullet": "重要なポイントを箇条書き（・）で列挙した要約",
    }

    def __init__(self, config: Config):
        self._config = config

    def is_available(self) -> tuple[bool, str]:
        """要約機能が使える状態か (有効/APIキーあり) を確認"""
        if not self._config.summary_enabled:
            return False, "要約モードが無効です（config.toml の enabled=false）"
        env_name = self._env_var_name()
        if not env_name:
            return False, f"未対応のプロバイダ: {self._config.summary_provider}"
        if not os.environ.get(env_name):
            return False, f"環境変数 {env_name} が設定されていません"
        return True, ""

    def _env_var_name(self) -> str:
        m = {
            "openai": "OPENAI_API_KEY",
            "anthropic": "ANTHROPIC_API_KEY",
            "gemini": "GEMINI_API_KEY",
        }
        return m.get(self._config.summary_provider, "")

    def _build_prompt(self, text: str, is_part: bool = False) -> str:
        length_instr = self._LENGTH_INSTRUCTIONS.get(
            self._config.summary_length,
            self._LENGTH_INSTRUCTIONS["medium"],
        )
        extra = self._config.summary_extra_instruction.strip()
        extra_part = f"\n追加指示: {extra}" if extra else ""
        part_note = (
            "\n注意: これは長文の一部分です。この部分の内容のみを要約し、"
            "「以下は」「この文章は」などの前置きは付けないでください。"
            if is_part else ""
        )
        return (
            f"以下の日本語テキストを要約してください。\n"
            f"形式: {length_instr}{extra_part}{part_note}\n"
            f"出力は要約本文のみ（前置き・解説・記号枠は不要）。\n\n"
            f"--- 元のテキスト ---\n{text}\n--- ここまで ---"
        )

    def _split_for_summary(self, text: str, max_chars: int) -> list[str]:
        """要約 API に渡すためのチャンク分割。文末(。！？!?改行)で区切る。"""
        if len(text) <= max_chars:
            return [text]
        import re as _re
        # 文末候補で区切る（区切り文字は前のチャンクに含める）
        parts = _re.split(r'(?<=[。！？!?\n])', text)
        chunks: list[str] = []
        cur = ""
        for p in parts:
            if not p:
                continue
            if len(cur) + len(p) > max_chars and cur:
                chunks.append(cur)
                cur = p
            else:
                cur += p
        if cur.strip():
            chunks.append(cur)
        return chunks

    def summarize(self, text: str,
                  progress_cb=None) -> list[tuple[str, str]] | None:
        """要約結果を [(元セクション, 要約), ...] のリストで返す。失敗時は None。
        progress_cb(i, n) を渡すとチャンクごとに進捗を通知する。
        短文の場合はリスト長 1 のリストになる。
        """
        ok, reason = self.is_available()
        if not ok:
            log.warning(f"要約スキップ: {reason}")
            return None

        # 入力を要約処理単位に分割
        chunks = self._split_for_summary(text, self._config.summary_chunk_chars)
        log.info(f"要約処理チャンク数: {len(chunks)} (単位 ~{self._config.summary_chunk_chars}文字)")

        is_part = len(chunks) > 1
        sections: list[tuple[str, str]] = []
        for i, ck in enumerate(chunks):
            log.info(f"要約 {i+1}/{len(chunks)} ({len(ck)} 文字)")
            if progress_cb:
                progress_cb(i + 1, len(chunks))
            s = self._summarize_one(ck, is_part=is_part)
            if s is None:
                log.warning(f"要約 {i+1}/{len(chunks)} 失敗 → 元テキストをそのまま採用")
                s = ck
            sections.append((ck.strip(), s.strip()))
        return sections

    def _summarize_one(self, text: str, is_part: bool) -> str | None:
        """単一チャンクを要約する。失敗時は None。"""
        prompt = self._build_prompt(text, is_part=is_part)
        provider = self._config.summary_provider
        try:
            if provider == "openai":
                return self._call_openai(prompt)
            elif provider == "anthropic":
                return self._call_anthropic(prompt)
            elif provider == "gemini":
                return self._call_gemini(prompt)
            else:
                log.warning(f"未対応のプロバイダ: {provider}")
                return None
        except Exception as e:
            log.error(f"要約API呼び出し失敗: {e}")
            return None

    def _http_post_json(self, url: str, headers: dict, body: dict) -> dict:
        data = json.dumps(body).encode("utf-8")
        req = urllib.request.Request(url, data=data, headers=headers, method="POST")
        with urllib.request.urlopen(req, timeout=self._config.summary_timeout_s) as resp:
            raw = resp.read().decode("utf-8")
        return json.loads(raw)

    def _call_openai(self, prompt: str) -> str | None:
        api_key = os.environ.get("OPENAI_API_KEY", "")
        url = "https://api.openai.com/v1/chat/completions"
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }
        body = {
            "model": self._config.summary_model,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.3,
        }
        result = self._http_post_json(url, headers, body)
        try:
            return result["choices"][0]["message"]["content"].strip()
        except (KeyError, IndexError):
            log.error(f"OpenAI 応答パース失敗: {str(result)[:300]}")
            return None

    def _call_anthropic(self, prompt: str) -> str | None:
        api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        url = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json",
        }
        body = {
            "model": self._config.summary_model,
            "max_tokens": 2000,
            "messages": [{"role": "user", "content": prompt}],
        }
        result = self._http_post_json(url, headers, body)
        try:
            parts = result.get("content", [])
            for p in parts:
                if p.get("type") == "text":
                    return p.get("text", "").strip()
            return None
        except Exception:
            log.error(f"Anthropic 応答パース失敗: {str(result)[:300]}")
            return None

    def _call_gemini(self, prompt: str) -> str | None:
        api_key = os.environ.get("GEMINI_API_KEY", "")
        model = self._config.summary_model
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        headers = {"Content-Type": "application/json"}
        body = {"contents": [{"parts": [{"text": prompt}]}]}
        result = self._http_post_json(url, headers, body)
        try:
            return result["candidates"][0]["content"]["parts"][0]["text"].strip()
        except (KeyError, IndexError):
            log.error(f"Gemini 応答パース失敗: {str(result)[:300]}")
            return None


# =====================================================================
# OverlayWindow — 読み上げ中のテキストを画面下部に表示
# =====================================================================
class OverlayWindow:
    """読み上げ中のチャンクを画面下部 or 上部に表示する常に最前面のオーバーレイ"""

    def __init__(self):
        self._queue: queue.Queue = queue.Queue()
        self._position = "bottom"  # "bottom" or "top"
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    # 1段モード（通常）と 2段モード（要約: 上=元セクション、下=読み上げ中）の高さ
    _H_SINGLE = 120
    _H_DUAL   = 280
    # 2段モードでの上段固定高さ（残りが下段になる）
    _H_CONTEXT = 110

    def _run(self):
        root = tk.Tk()
        root.withdraw()
        root.overrideredirect(True)       # タイトルバーなし
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.88)
        root.configure(bg="#1a1a1a")

        # 画面サイズ
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        w = min(1000, screen_w - 100)
        x = (screen_w - w) // 2

        # 現在表示中の高さ（1段 or 2段）
        state = {"h": self._H_SINGLE, "dual": False}

        def _apply_position(pos: str):
            h = state["h"]
            if pos == "top":
                y = 40
            else:
                y = screen_h - h - 80
            root.geometry(f"{w}x{h}+{x}+{y}")

        # ---------- 上段: コンテキスト（要約モードで使用） ----------
        # 高さを固定して、長文が下段を押し出さないようにする
        context_frame = tk.Frame(root, bg="#0d0d0d", height=self._H_CONTEXT)
        context_frame.pack_propagate(False)  # 中身に応じて伸縮しないようにする
        context_title_lbl = tk.Label(
            context_frame, text="", font=("Meiryo", 11, "bold"),
            fg="#ffaa00", bg="#0d0d0d", anchor="w",
        )
        context_title_lbl.pack(fill="x", padx=15, pady=(6, 0))
        context_text_lbl = tk.Label(
            context_frame, text="", font=("Meiryo", 12),
            fg="#cccccc", bg="#0d0d0d", wraplength=w - 30,
            justify="left", anchor="nw",
        )
        context_text_lbl.pack(expand=True, fill="both", padx=15, pady=(2, 6))
        # 区切り線
        divider = tk.Frame(root, bg="#666666", height=2)

        # ---------- 下段: 通常表示（読み上げ中の文） ----------
        main_frame = tk.Frame(root, bg="#1a1a1a")
        title_lbl = tk.Label(
            main_frame, text="", font=("Meiryo", 10),
            fg="#88ccff", bg="#1a1a1a", anchor="w",
        )
        title_lbl.pack(fill="x", padx=15, pady=(8, 0))
        text_lbl = tk.Label(
            main_frame, text="", font=("Meiryo", 16, "bold"),
            fg="#ffffff", bg="#1a1a1a", wraplength=w - 30,
            justify="left", anchor="w",
        )
        text_lbl.pack(expand=True, fill="both", padx=15, pady=(4, 10))
        main_frame.pack(expand=True, fill="both")

        def _enter_dual_mode():
            if state["dual"]:
                return
            # main_frame を一旦外して、上段→区切り→下段の順に詰め直す
            main_frame.pack_forget()
            context_frame.pack(fill="x")
            divider.pack(fill="x")
            main_frame.pack(expand=True, fill="both")
            state["dual"] = True
            state["h"] = self._H_DUAL
            _apply_position(self._position)

        def _exit_dual_mode():
            if not state["dual"]:
                return
            context_frame.pack_forget()
            divider.pack_forget()
            state["dual"] = False
            state["h"] = self._H_SINGLE
            _apply_position(self._position)

        _apply_position("bottom")

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
                        _exit_dual_mode()
                        root.withdraw()
                    elif action == "set_context":
                        context_title_lbl.config(text=a)
                        context_text_lbl.config(text=b)
                        _enter_dual_mode()
                    elif action == "clear_context":
                        _exit_dual_mode()
                    elif action == "toggle_position":
                        new_pos = "top" if self._position == "bottom" else "bottom"
                        self._position = new_pos
                        _apply_position(new_pos)
                        log.info(f"オーバーレイ位置を{new_pos}に変更")
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

    def toggle_position(self) -> None:
        """オーバーレイの位置を上下で切り替える"""
        self._queue.put(("toggle_position", "", ""))

    def set_context(self, title: str, text: str) -> None:
        """上段に元セクション（要約元）を表示する"""
        self._queue.put(("set_context", title, text))

    def clear_context(self) -> None:
        """上段（コンテキスト）を非表示にして 1 段表示に戻す"""
        self._queue.put(("clear_context", "", ""))


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
        log.warning(f"MCI '{cmd}' → rc={ret}")  # DEBUG→WARNING: ログに常時出力
    return ret


def _mci_status(cmd: str) -> str:
    """MCI status コマンドを送信して結果文字列を返す"""
    buf = ctypes.create_unicode_buffer(256)
    _mciSendStringW(cmd, buf, 256, None)
    return buf.value


import re as _re

# 文の区切りパターン（日本語の句点・感嘆符・疑問符・改行など）
_SENTENCE_SPLIT = _re.compile(r'(?<=[。！？!?\n])\s*')

# TTS に送る前に除去する不要記号
# ボックス罫線・ブロック要素・装飾記号など edge-tts が読み上げられない文字
_NOISE_CHARS = _re.compile(
    r'['
    r'─-╿'   # ボックス罫線（─ │ ┌ など）
    r'▀-▟'   # ブロック要素（█ ▀ など）
    r'■-◿'   # 幾何学図形（■ ▲ ● など）
    r'☀-⛿'   # 各種記号（☀ ★ など）— 必要なら外してください
    r'～'          # ～（全角チルダ）
    r']+'
)


def _clean_for_tts(text: str) -> str:
    """TTS に不要な記号を除去し、読み上げやすい形に整える"""
    # 不要記号を空白に置換
    cleaned = _NOISE_CHARS.sub(' ', text)
    # 連続する空白・全角スペースを 1 つにまとめる
    cleaned = _re.sub(r'[ 　]{2,}', ' ', cleaned).strip()
    return cleaned

# 短文を結合するときの上限文字数
# 句点（。！？!?）ごとに分割し、この文字数に達するまで結合する。
# 読点（、）での分割は _CHUNK_FORCE_SPLIT 超えの超長文にのみ使用。
_CHUNK_MAX_CHARS   = 100   # 短文まとめ上限
_CHUNK_FORCE_SPLIT = 150   # この長さを超える文節のみ読点で強制分割


def _split_into_chunks(text: str) -> list[str]:
    """テキストを文単位のチャンクに分割する。
    分割は句点（。！？!?\\n）のみ。読点（、）は超長文の最終手段。
    """
    sentences = _SENTENCE_SPLIT.split(text.strip())
    sentences = [s.strip() for s in sentences if s.strip()]

    if not sentences:
        return [text.strip()] if text.strip() else []

    chunks: list[str] = []
    current = ""
    for s in sentences:
        if len(current) + len(s) > _CHUNK_MAX_CHARS:
            if current:
                chunks.append(current)
            # 一文自体が非常に長い場合のみ読点で分割
            if len(s) > _CHUNK_FORCE_SPLIT:
                parts = _re.split(r'(?<=[、,，])\s*', s)
                sub = ""
                for p in parts:
                    if len(sub) + len(p) > _CHUNK_FORCE_SPLIT and sub:
                        chunks.append(sub)
                        sub = p
                    else:
                        sub += p
                if sub:
                    chunks.append(sub)
                current = ""
            else:
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
        self._done = threading.Event()
        self._done.set()
        self._source_hwnd: int = 0             # Ctrl+Alt+E 発火時のウィンドウ
        self._current_chunk: str = ""          # 一時停止時の表示用
        self._python_exe: str | None = None
        # ---- Tab チャンク境界一時停止 ----
        self._tab_lock = threading.Lock()
        self._tab_pause_requested: bool = False  # Tab押下→次境界で停止
        self._tab_paused: bool = False           # 現在チャンク境界で停止中
        self._tab_resume_event = threading.Event()
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
        self._done.clear()
        self._queue.put(text)

    def toggle_tab_pause(self) -> None:
        """Tab キー: チャンク境界での一時停止 / 再開トグル"""
        if not self._speaking.is_set():
            return
        with self._tab_lock:
            if self._tab_paused:
                # 現在停止中 → 再開
                self._tab_paused = False
                self._tab_resume_event.set()
                log.info("Tab: 再開")
                if self._overlay is not None:
                    self._overlay.update("▶ 読み上げ再開", self._current_chunk)
            elif self._tab_pause_requested:
                # 停止予約中 → キャンセル
                self._tab_pause_requested = False
                log.info("Tab: 一時停止キャンセル")
            else:
                # 再生中 → 次のチャンク境界で停止予約
                self._tab_pause_requested = True
                log.info("Tab: 次チャンク境界で一時停止予定")

    def wait_done(self, timeout: float | None = None) -> bool:
        return self._done.wait(timeout=timeout)

    def stop_current(self) -> None:
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        self._stop_flag.set()
        # チャンク境界一時停止中なら解除してループを抜けさせる
        with self._tab_lock:
            self._tab_pause_requested = False
            self._tab_paused = False
            self._tab_resume_event.set()
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
        # 古いダブルバッファエイリアスも念のため閉じる（過去バージョン互換）
        for _a in ("tts_0", "tts_1"):
            _mci(f"stop {_a}")
            _mci(f"close {_a}")
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
        # TTS に不適切な記号を除去してから送る
        text = _clean_for_tts(text)
        if not text:
            log.warning("クリーニング後にテキストが空になりました。スキップします。")
            return False
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
        # 再生完了をポーリングで待つ
        while not self._stop_flag.is_set():
            status = _mci_status("status tts mode")
            if status != "playing":
                break
            time.sleep(0.05)
        _mci("stop tts")
        _mci("close tts")

    def _speak(self, text: str) -> None:
        # クリーニング後に空になるチャンク（罫線のみなど）を事前除外
        chunks = [c for c in _split_into_chunks(text) if _clean_for_tts(c).strip()]
        if not chunks:
            log.info("読み上げ可能なテキストがありません")
            return
        log.info(f"TTS 開始 ({len(text)} 文字, {len(chunks)} チャンク, {_TTS_VOICE})")
        self._speaking.set()
        self._stop_flag.clear()

        with self._tab_lock:
            self._tab_pause_requested = False
            self._tab_paused = False
            self._tab_resume_event.clear()

        # MP3 一時ファイルは OneDrive 外に置く
        tmp_dir = Path(os.environ.get("LOCALAPPDATA", os.environ.get("TEMP", "."))) / "yomiage"
        try:
            tmp_dir.mkdir(parents=True, exist_ok=True)
        except Exception as _e:
            log.warning(f"一時フォルダ作成失敗 → スクリプト隣を使用: {_e}")
            tmp_dir = Path(__file__).parent

        # 各チャンクの MP3 ファイルパス（重複防止に乱数を含める）
        import uuid
        _session_id = uuid.uuid4().hex[:8]
        mp3_files = [tmp_dir / f"_tts_{_session_id}_{i}.mp3" for i in range(len(chunks))]

        # ---- バックグラウンド: 全チャンクの MP3 を順次生成 ----
        # MCI には触れない（メインスレッドのみが MCI を扱う）
        gen_done = [False] * len(chunks)
        gen_lock = threading.Lock()

        def _gen_worker():
            for idx, ch in enumerate(chunks):
                if self._stop_flag.is_set():
                    return
                ok = self._generate_mp3(ch, mp3_files[idx])
                with gen_lock:
                    gen_done[idx] = ok
                if not ok:
                    log.warning(f"チャンク {idx+1} 生成失敗")

        gen_thread = threading.Thread(target=_gen_worker, daemon=True)
        gen_thread.start()

        # ---- 開始ビープ ----
        try:
            winsound.PlaySound(_BEEP_WAV, winsound.SND_MEMORY)
        except Exception:
            pass

        def _wait_chunk_ready(idx: int, timeout_s: float = 60.0) -> bool:
            """指定チャンクの MP3 が生成完了するまで待つ"""
            deadline = time.monotonic() + timeout_s
            while not self._stop_flag.is_set() and time.monotonic() < deadline:
                with gen_lock:
                    if gen_done[idx]:
                        return mp3_files[idx].exists() and mp3_files[idx].stat().st_size > 100
                # ジェネレータースレッドが終了して未完なら失敗
                if not gen_thread.is_alive():
                    return mp3_files[idx].exists() and mp3_files[idx].stat().st_size > 100
                time.sleep(0.05)
            return False

        try:
            for i, chunk in enumerate(chunks):
                if self._stop_flag.is_set():
                    log.info("TTS 中断")
                    break

                # MP3 生成完了を待つ
                if not _wait_chunk_ready(i):
                    if self._stop_flag.is_set():
                        break
                    log.warning(f"チャンク {i+1} の MP3 が未生成 → スキップ")
                    continue

                mp3_file = mp3_files[i]

                # 念のため前回の alias を閉じる
                _mci("stop tts")
                _mci("close tts")

                # MCI open
                rc = _mci(f'open "{mp3_file}" type mpegvideo alias tts')
                if rc != 0:
                    log.error(f"MCI open 失敗 rc={rc} → スキップ")
                    continue

                # MCI play
                rc = _mci("play tts")
                if rc != 0:
                    log.error(f"MCI play 失敗 rc={rc} → スキップ")
                    _mci("close tts")
                    continue

                log.info(f"チャンク {i+1}/{len(chunks)} 再生 ({len(chunk)} 文字)")

                if self._source_hwnd:
                    _scroll_to_chunk(chunk, self._source_hwnd)

                self._current_chunk = chunk
                if self._overlay is not None:
                    title = f"読み上げ中 {i+1}/{len(chunks)}"
                    if i == 0:
                        self._overlay.show(title, chunk)
                    else:
                        self._overlay.update(title, chunk)

                # ---- 再生完了待ち ----
                # フェーズ1: "playing" 状態になるまで待つ（最大 2 秒）
                _p1_deadline = time.monotonic() + 2.0
                while not self._stop_flag.is_set() and time.monotonic() < _p1_deadline:
                    if _mci_status("status tts mode") == "playing":
                        break
                    time.sleep(0.02)

                # フェーズ2: 文字数ベースの推定時間 90% を最低保証して完了待ち
                _estimated_s = max(len(chunk) * 0.08, 1.0)
                _p2_start = time.monotonic()
                while not self._stop_flag.is_set():
                    mode = _mci_status("status tts mode")
                    elapsed = time.monotonic() - _p2_start
                    if mode != "playing" and elapsed >= _estimated_s * 0.9:
                        break
                    time.sleep(0.05)

                _mci("stop tts")
                _mci("close tts")

                if self._stop_flag.is_set():
                    log.info("TTS 中断（Esc/停止）")
                    break

                # ---- Tab によるチャンク境界一時停止 ----
                with self._tab_lock:
                    should_pause = self._tab_pause_requested
                    if should_pause:
                        self._tab_pause_requested = False
                        self._tab_paused = True

                if should_pause:
                    self._tab_resume_event.clear()
                    if self._overlay is not None:
                        self._overlay.update("⏸ 一時停止中 — Tab で再開", self._current_chunk)
                    log.info("Tab 一時停止（チャンク境界）")
                    self._tab_resume_event.wait()
                    with self._tab_lock:
                        self._tab_paused = False
                    if self._stop_flag.is_set():
                        log.info("TTS 中断（一時停止中に Esc）")
                        break
                    log.info("Tab 再開")
            else:
                log.info("TTS 完了")

        except Exception as e:
            log.error(f"TTS エラー: {e}", exc_info=True)
        finally:
            self._stop_flag.set()  # ジェネレータースレッドを止める
            _mci("stop tts")
            _mci("close tts")
            gen_thread.join(timeout=3)
            with self._tab_lock:
                self._tab_pause_requested = False
                self._tab_paused = False
                self._tab_resume_event.set()
            with self._proc_lock:
                self._gen_proc = None
            self._speaking.clear()
            self._done.set()
            if self._overlay is not None:
                self._overlay.hide()
            for f in mp3_files:
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
_MSG_REGISTER_ESC          = _WM_USER + 1
_MSG_UNREGISTER_ESC        = _WM_USER + 2
_MSG_REGISTER_TAB          = _WM_USER + 3
_MSG_UNREGISTER_TAB        = _WM_USER + 4
_MSG_REGISTER_SHIFT_TAB    = _WM_USER + 5
_MSG_UNREGISTER_SHIFT_TAB  = _WM_USER + 6


class HotkeyHandler:
    def __init__(self, tts: TTSEngine, clipboard: ClipboardManager,
                 summary_engine: SummaryEngine | None = None):
        self._tts = tts
        self._clipboard = clipboard
        self._summary = summary_engine
        self._lock = threading.Lock()
        self._thread_id: int | None = None
        self._esc_registered = False
        self._tab_registered = False
        self._shift_tab_registered = False

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

    def _register_shift_tab(self) -> None:
        """ホットキースレッドに Shift+Tab 登録を依頼"""
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(
                self._thread_id, _MSG_REGISTER_SHIFT_TAB, 0, 0)

    def _unregister_shift_tab(self) -> None:
        """ホットキースレッドに Shift+Tab 解除を依頼"""
        if self._thread_id:
            ctypes.windll.user32.PostThreadMessageW(
                self._thread_id, _MSG_UNREGISTER_SHIFT_TAB, 0, 0)

    def _hotkey_loop(self) -> None:
        user32 = ctypes.windll.user32
        self._thread_id = ctypes.windll.kernel32.GetCurrentThreadId()

        ok = user32.RegisterHotKey(None, _HOTKEY_READ, _MOD_CTRL_ALT, _VK_R)
        if not ok:
            log.warning("RegisterHotKey(Ctrl+Alt+R) 失敗 — 既存インスタンスを終了して再試行します")
            # 同じ yomiage.py を実行中の古いプロセスを終了する
            _current_pid = os.getpid()
            try:
                subprocess.run(
                    ["wmic", "process", "where",
                     f"(name='pythonw.exe' or name='python.exe')"
                     f" and ProcessId!='{_current_pid}'"
                     f" and CommandLine like '%yomiage%'",
                     "delete"],
                    capture_output=True, timeout=5,
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
            except Exception as _e:
                log.warning(f"既存インスタンス終了試行: {_e}")
            time.sleep(1.2)
            ok = user32.RegisterHotKey(None, _HOTKEY_READ, _MOD_CTRL_ALT, _VK_R)
            if not ok:
                log.error("RegisterHotKey(Ctrl+Alt+R) 再試行も失敗 — 別アプリが使用中の可能性")
                return
            log.info("Ctrl+Alt+R を登録しました（古いインスタンス終了後に再登録）")
        log.info("Ctrl+Alt+R を登録しました")

        ok2 = user32.RegisterHotKey(None, _HOTKEY_READ_END, _MOD_CTRL_ALT, _VK_E)
        if ok2:
            log.info("Ctrl+Alt+E を登録しました（カーソル位置から末尾まで読み上げ）")
        else:
            log.warning("RegisterHotKey(Ctrl+Alt+E) 失敗（他アプリが使用中の可能性）")

        ok3 = user32.RegisterHotKey(None, _HOTKEY_SUMMARY, _MOD_CTRL_ALT, _VK_S)
        if ok3:
            log.info("Ctrl+Alt+S を登録しました（要約してから読み上げ）")
        else:
            log.warning("RegisterHotKey(Ctrl+Alt+S) 失敗（他アプリが使用中の可能性）")

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
                    self._do_unregister_shift_tab(user32)
                elif msg.wParam == _HOTKEY_PAUSE:
                    log.info("Tab → チャンク境界一時停止/再開")
                    self._tts.toggle_tab_pause()
                elif msg.wParam == _HOTKEY_TOGGLE_POS:
                    log.info("Shift+Tab → オーバーレイ上下切替")
                    if self._tts._overlay is not None:
                        self._tts._overlay.toggle_position()
                elif msg.wParam == _HOTKEY_SUMMARY:
                    if self._tts.is_speaking:
                        log.info("Ctrl+Alt+S → 読み上げ停止")
                        self._tts.stop_current()
                        self._do_unregister_esc(user32)
                    else:
                        log.info("Ctrl+Alt+S → 要約読み上げ開始（選択テキスト）")
                        t = threading.Thread(
                            target=self._speak_selected_text_summary, daemon=True)
                        t.start()

            elif msg.message == _MSG_REGISTER_ESC:
                self._do_register_esc(user32)
            elif msg.message == _MSG_UNREGISTER_ESC:
                self._do_unregister_esc(user32)
            elif msg.message == _MSG_REGISTER_TAB:
                self._do_register_tab(user32)
            elif msg.message == _MSG_UNREGISTER_TAB:
                self._do_unregister_tab(user32)
            elif msg.message == _MSG_REGISTER_SHIFT_TAB:
                self._do_register_shift_tab(user32)
            elif msg.message == _MSG_UNREGISTER_SHIFT_TAB:
                self._do_unregister_shift_tab(user32)

        user32.UnregisterHotKey(None, _HOTKEY_READ)
        user32.UnregisterHotKey(None, _HOTKEY_READ_END)
        user32.UnregisterHotKey(None, _HOTKEY_SUMMARY)
        if self._esc_registered:
            user32.UnregisterHotKey(None, _HOTKEY_ESC)
        if self._tab_registered:
            user32.UnregisterHotKey(None, _HOTKEY_PAUSE)
        if self._shift_tab_registered:
            user32.UnregisterHotKey(None, _HOTKEY_TOGGLE_POS)

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

    def _do_register_shift_tab(self, user32) -> None:
        if not self._shift_tab_registered:
            if user32.RegisterHotKey(None, _HOTKEY_TOGGLE_POS, _MOD_SHIFT, _VK_TAB):
                self._shift_tab_registered = True
                log.info("Shift+Tab ホットキー登録（オーバーレイ上下切替）")
            else:
                log.warning("Shift+Tab ホットキー登録失敗（他アプリが使用中の可能性）")

    def _do_unregister_shift_tab(self, user32) -> None:
        if self._shift_tab_registered:
            user32.UnregisterHotKey(None, _HOTKEY_TOGGLE_POS)
            self._shift_tab_registered = False
            log.info("Shift+Tab ホットキー解除")

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
                self._register_shift_tab()
                self._tts.speak(text)
                self._tts.wait_done()
                self._unregister_esc()
                self._unregister_tab()
                self._unregister_shift_tab()
            else:
                log.info("テキストが取得できませんでした")
        finally:
            self._lock.release()

    def _speak_selected_text_summary(self) -> None:
        """選択テキストを生成 AI で要約してから読み上げる"""
        if not self._lock.acquire(blocking=False):
            return
        try:
            time.sleep(0.3)
            text = self._clipboard.get_selected_text()
            if not text:
                log.info("要約対象テキストが取得できませんでした")
                return
            log.info(f"要約対象: {text[:60]}{'...' if len(text) > 60 else ''} ({len(text)} 文字)")

            if self._summary is None:
                log.warning("要約エンジン未初期化")
                return

            ok, reason = self._summary.is_available()
            if not ok:
                log.warning(f"要約スキップ: {reason}")
                # フォールバック判定
                if self._summary._config.summary_on_error == "fallback":
                    log.info("→ 元のテキストをそのまま読み上げます")
                    self._speak_text_with_overlay(text)
                else:
                    # エラー内容を読み上げ
                    self._speak_text_with_overlay(f"要約できませんでした。{reason}")
                return

            # オーバーレイに「要約中...」表示
            overlay = getattr(self._tts, "_overlay", None)
            if overlay is not None:
                overlay.show("🤖 要約中...", "AI に問い合わせています。少々お待ちください。")

            # チャンクごとの進捗をオーバーレイに表示
            def _progress(i: int, n: int):
                if overlay is not None and n > 1:
                    overlay.update(
                        f"🤖 要約中... {i}/{n}",
                        "セクションごとに AI で要約しています。少々お待ちください。",
                    )

            # 要約 API 呼び出し
            t0 = time.monotonic()
            sections = self._summary.summarize(text, progress_cb=_progress)
            elapsed = time.monotonic() - t0
            log.info(f"要約 API 応答時間: {elapsed:.2f}秒")

            if not sections:
                log.warning("要約失敗")
                if overlay is not None:
                    overlay.hide()
                if self._summary._config.summary_on_error == "fallback":
                    log.info("→ 元のテキストをそのまま読み上げます")
                    self._speak_text_with_overlay(text)
                return

            n = len(sections)
            log.info(f"要約セクション数: {n}")

            # 前回の Esc 停止で立った _stop_flag をクリア
            # （クリアしないと最初のセクションで即 break してしまう）
            self._tts._stop_flag.clear()

            # ホットキー（Esc/Tab/Shift+Tab）はループ全体で 1 回だけ登録
            self._register_esc()
            self._register_tab()
            self._register_shift_tab()
            try:
                for i, (orig_section, summary_text) in enumerate(sections):
                    if self._tts._stop_flag.is_set():
                        log.info("要約読み上げを途中で中断")
                        break
                    log.info(
                        f"セクション {i+1}/{n} 読み上げ "
                        f"(元={len(orig_section)}文字, 要約={len(summary_text)}文字)"
                    )
                    # 上段に元セクションを表示
                    # 上段は高さ固定なので、文字数も短めに（先頭160字程度）抑える
                    if overlay is not None:
                        ctx = orig_section if len(orig_section) <= 160 \
                            else orig_section[:160] + " …"
                        ctx_label = (
                            f"📄 元のセクション {i+1}/{n}"
                            if n > 1 else "📄 元の文章"
                        )
                        overlay.set_context(ctx_label, ctx)
                        # 下段の古い「要約中...」表示をすぐに置き換える。
                        # MP3 生成に数秒かかるので、その間「音声準備中」を表示。
                        # （実際の再生開始時に _speak 内で再度 update される）
                        section_title = f"📝 要約 {i+1}/{n} 準備中..." if n > 1 else "📝 要約 準備中..."
                        overlay.update(section_title, summary_text[:120] + (" …" if len(summary_text) > 120 else ""))
                    # 下段に要約を読み上げ（既存の TTS パイプライン）
                    self._tts.speak(summary_text)
                    self._tts.wait_done()
            finally:
                self._unregister_esc()
                self._unregister_tab()
                self._unregister_shift_tab()
                if overlay is not None:
                    overlay.clear_context()
        finally:
            self._lock.release()

    def _speak_text_with_overlay(self, text: str) -> None:
        """共通: テキストを Esc/Tab/Shift+Tab ホットキー登録付きで読み上げる"""
        self._register_esc()
        self._register_tab()
        self._register_shift_tab()
        self._tts.speak(text)
        self._tts.wait_done()
        self._unregister_esc()
        self._unregister_tab()
        self._unregister_shift_tab()

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
                # クリップボード取得完了後、Left arrow で選択を解除しカーソルを先頭（元の位置）に戻す
                _release_modifiers()
                time.sleep(0.05)
                _send_one_key(0x25, _KEYEVENTF_EXTENDEDKEY)           # VK_LEFT（拡張キー）
                time.sleep(0.05)
                _send_one_key(0x25, _KEYEVENTF_EXTENDEDKEY | _KEYEVENTF_KEYUP)
                time.sleep(0.05)
                log.info(f"カーソルから末尾まで読み上げ: {text[:60]}{'...' if len(text) > 60 else ''}")
                self._register_esc()
                self._register_tab()
                self._register_shift_tab()
                self._tts.speak(text, source_hwnd=hwnd)   # hwnd を渡してスクロール有効化
                self._tts.wait_done()
                self._unregister_esc()
                self._unregister_tab()
                self._unregister_shift_tab()
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
        summary_engine = SummaryEngine(_CONFIG)
        hotkey = HotkeyHandler(tts, clipboard, summary_engine)
        app = TrayApp(tts, hotkey)
        app.run()
    except Exception:
        log.exception("致命的エラーで終了しました")
        sys.exit(1)
