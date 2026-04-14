# yomiage — 選択テキスト読み上げアプリ

Windows 上でテキストを選択してホットキーを押すと、日本語音声で読み上げるシステムトレイ常駐アプリ。

## 機能

- **Ctrl+Alt+R** : 選択テキストを日本語音声（Microsoft Haruka）で読み上げ
- **Esc** : 読み上げを即座に停止（読み上げ中のみ有効）
- **Ctrl+Alt+R 再押し** : 読み上げ中なら停止
- システムトレイアイコンから停止・終了

## 必要環境

- Windows 10/11
- Python 3.10+
- 日本語音声パック（Microsoft Haruka Desktop）

## インストール

```bash
pip install pystray pyperclip pillow
```

## 起動方法

```bash
pythonw yomiage.py
```

またはバッチファイルをダブルクリック:

```
yomiage.bat
```

## 仕組み

1. `RegisterHotKey` API でグローバルホットキー（Ctrl+Alt+R）を登録
2. ホットキー検出 → `SendInput` で Ctrl+C を送信して選択テキストをクリップボードにコピー
3. PowerShell の `System.Speech.Synthesis.SpeechSynthesizer` で日本語音声読み上げ
4. 読み上げ中は Esc キーを動的に登録し、即座に中断可能
