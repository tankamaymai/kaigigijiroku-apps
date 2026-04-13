# 議事録メーカー セットアップ指示書

> **このファイルは AI エージェント（Cursor, ChatGPT 等）向けの指示書です。**
> 上から順に実行してください。すべて完了するとアプリがブラウザで開きます。

---

## 前提

- このファイルはプロジェクトルートに配置されています
- プロジェクトには `scripts/setup_mac.sh`（Mac用）と `scripts/setup_windows.ps1`（Windows用）が含まれています
- セットアップ完了後、ユーザーは http://localhost:8000 で議事録作成アプリを使えます

---

## 手順

### ステップ 1: OS を判定する

以下のコマンドでOSを判定してください。

```bash
uname -s
```

- 出力が `Darwin` → **Mac** です。ステップ 2A に進んでください。
- コマンドがエラーになる、または `MINGW`/`MSYS` を含む → **Windows** です。ステップ 2B に進んでください。

---

### ステップ 2A: Mac の場合

セットアップスクリプトを実行します。

```bash
cd /path/to/kaigigijiroku-apps
chmod +x scripts/setup_mac.sh
./scripts/setup_mac.sh
```

> `/path/to/kaigigijiroku-apps` は実際のプロジェクトディレクトリのパスに置き換えてください。
> Cursor で実行している場合はワークスペースルートで実行すれば OK です。

スクリプトは以下を自動で行います:
1. Homebrew の確認/インストール
2. Python 3, ffmpeg のインストール
3. Python 仮想環境の作成と依存パッケージのインストール
4. Ollama のインストールと起動
5. gemma4 モデルのダウンロード（サイズはモデルにより異なります）
6. アプリ起動とブラウザ表示

**gemma4 のダウンロードに数分〜数十分かかります。** 途中で止まっているように見えても、ダウンロードが進行中なので待ってください。

---

### ステップ 2B: Windows の場合

PowerShell を **管理者として** 実行し、以下を実行します。

```powershell
cd C:\path\to\kaigigijiroku-apps
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\scripts\setup_windows.ps1
```

> `C:\path\to\kaigigijiroku-apps` は実際のプロジェクトディレクトリのパスに置き換えてください。

スクリプトは以下を自動で行います:
1. Python 3 の確認/インストール（winget 使用）
2. ffmpeg のインストール（winget 使用）
3. Python 仮想環境の作成と依存パッケージのインストール
4. Ollama のインストールと起動
5. gemma4 モデルのダウンロード（サイズはモデルにより異なります）
6. アプリ起動とブラウザ表示

**gemma4 のダウンロードに数分〜数十分かかります。** 途中で止まっているように見えても、ダウンロードが進行中なので待ってください。

---

### ステップ 3: 検証

スクリプト完了後、以下のコマンドで動作を確認してください。

```bash
curl -s http://localhost:8000/api/health
```

期待する応答（JSON）:
```json
{
  "status": "ok",
  "faster_whisper": true
}
```

- `status` が `ok` → アプリは正常に動作しています
- `faster_whisper` が `true` → 高速文字起こしエンジンが有効です

ブラウザで http://localhost:8000 が開いていることを確認してください。

---

## トラブルシューティング

### Homebrew のインストールで止まる（Mac）
Xcode Command Line Tools のインストールが求められる場合があります。表示に従って許可してください。

### `pip install` でエラーが出る
```bash
pip install --upgrade pip setuptools wheel
pip install -r requirements.txt
```

### Ollama に接続できない
```bash
# Mac
ollama serve &

# Windows (PowerShell)
Start-Process ollama -ArgumentList "serve" -WindowStyle Hidden
```
起動後、3〜5秒待ってから再試行してください。

### gemma4 のダウンロードが途中で止まった
```bash
ollama pull gemma4
```
再実行すれば途中から再開されます。

### ポート 8000 が使用中
```bash
# Mac
lsof -ti :8000 | xargs kill -9

# Windows (PowerShell)
Get-Process -Id (Get-NetTCPConnection -LocalPort 8000).OwningProcess | Stop-Process -Force
```

### Windows で winget が使えない
Windows 10 バージョン 1709 以降に標準搭載されています。Microsoft Store から「アプリ インストーラー」を更新してください。

---

## 次回の起動方法

### Mac
```bash
cd /path/to/kaigigijiroku-apps
source venv/bin/activate
python server.py
```

### Windows
```powershell
cd C:\path\to\kaigigijiroku-apps
.\venv\Scripts\Activate.ps1
python server.py
```

ブラウザで http://localhost:8000 を開いてください。
