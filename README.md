# yomiage — 選択テキスト読み上げアプリ

Windows 上でテキストを選択してホットキーを押すと、日本語音声で読み上げるシステムトレイ常駐アプリ。

## 機能

- **Ctrl+Alt+R** : 選択テキストを日本語音声（Microsoft Nanami Neural）で読み上げ
- **Esc** : 読み上げを即座に停止（読み上げ中のみ有効）
- **Ctrl+Alt+R 再押し** : 読み上げ中なら停止
- システムトレイアイコンから停止・終了

## 必要環境

- Windows 10/11
- Python 3.10+
- インターネット接続（edge-tts による音声合成に必要）

## インストール

```bash
pip install pystray pyperclip pillow edge-tts
```

## 起動方法

バッチファイルをダブルクリック:

```
yomiage.bat
```

または直接実行:

```bash
pythonw yomiage.py
```

## 仕組み

1. `RegisterHotKey` API でグローバルホットキー（Ctrl+Alt+R）を登録
2. ホットキー検出 → `SendInput` で Ctrl+C を送信して選択テキストをクリップボードにコピー
3. `edge-tts` で Microsoft Nanami Neural 音声の MP3 を生成
4. `winmm.dll` MCI API で MP3 を直接再生（PowerShell 不要）
5. 読み上げ中は Esc キーを動的に登録し、即座に中断可能
