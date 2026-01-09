# 📝 議事録自動生成アプリ

会議音声から議事録を自動生成するローカルアプリです。

## 🎯 機能

- **音声文字起こし**: Whisperを使用して高精度な日本語文字起こし
- **テンプレート対応**: JSONテンプレートで様々な議事録フォーマットに対応
- **Excel出力**: 指定のExcelフォーマットに自動入力して出力

## 📋 必要なもの

- Python 3.9以上
- ffmpeg
- Whisper

## 🚀 セットアップ

### 1. ffmpegをインストール

```bash
brew install ffmpeg
```

### 2. Python仮想環境を作成

```bash
cd kaigi-app
python3 -m venv venv
source venv/bin/activate
```

### 3. 依存関係をインストール

```bash
pip install -r requirements.txt
```

## 💻 使い方

### アプリを起動

```bash
source venv/bin/activate
python meeting_app.py
```

### 操作手順

1. **テンプレート選択**: 使用する議事録テンプレートを選択
2. **音声選択**: 会議の録音ファイル（m4a/mp3/wav）を選択
3. **文字起こし**: 「文字起こしを実行」ボタンをクリック
4. **ChatGPTへ貼り付け**: 生成されたプロンプトをChatGPTに貼り付け
5. **JSON貼り付け**: ChatGPTが返したJSONをアプリに貼り付け
6. **Excel生成**: 「Excelを作成」ボタンをクリック

## 📁 フォルダ構成

```
kaigi-app/
├── meeting_app.py          # メインアプリ
├── requirements.txt        # 依存関係
├── templates/              # テンプレートJSON
│   └── shozokucho.json
├── excel_templates/        # Excelテンプレート
│   └── 所属長会議まとめ.xlsx
└── output/                 # 出力ファイル
```

## 🔧 テンプレートの追加方法

### 1. JSONテンプレートを作成

`templates/` に新しいJSONファイルを追加:

```json
{
  "name": "会議名",
  "excel_template": "excel_templates/テンプレート.xlsx",
  "sheet": "シート名",
  "sections": [
    { "key": "datetime", "label": "日時", "cell": "A2" },
    { "key": "item1", "label": "項目1", "cell": "B3" }
  ],
  "chatgpt_prompt": {
    "style_rules": [
      "箇条書きで記載",
      "重要事項は強調"
    ]
  }
}
```

### 2. Excelテンプレートを配置

`excel_templates/` に対応するExcelファイルを配置

## 🍎 Macアプリ化（オプション）

```bash
source venv/bin/activate
pyinstaller --onefile --windowed meeting_app.py
```

`dist/meeting_app` が生成されます。

## 📝 Whisperモデルについて

| モデル | 精度 | 速度 | 推奨用途 |
|--------|------|------|----------|
| tiny | 低 | 最速 | テスト用 |
| base | 低〜中 | 速い | 短い音声 |
| small | 中 | 普通 | 一般的な会議 |
| medium | 高 | やや遅い | 推奨（デフォルト） |
| large | 最高 | 遅い | 高精度が必要な場合 |

## ⚠️ 注意事項

- ChatGPTの出力は必ずJSONのみにしてください
- 2時間を超える長時間音声は分割することを推奨
- 初回のWhisper実行時はモデルのダウンロードが発生します
