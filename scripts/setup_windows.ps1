# ========================================
# 議事録メーカー - Windows セットアップスクリプト
# 全自動でアプリが使える状態にします
# PowerShell で実行してください
# ========================================

$ErrorActionPreference = "Stop"
$APP_DIR = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$LOG_FILE = Join-Path $APP_DIR "setup.log"

function Log($msg)  { Write-Host "[INFO] $msg" -ForegroundColor Cyan;  Add-Content $LOG_FILE "[INFO] $msg" }
function Ok($msg)   { Write-Host "[OK]   $msg" -ForegroundColor Green; Add-Content $LOG_FILE "[OK]   $msg" }
function Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow; Add-Content $LOG_FILE "[WARN] $msg" }
function Fail($msg) { Write-Host "[FAIL] $msg" -ForegroundColor Red;   Add-Content $LOG_FILE "[FAIL] $msg"; exit 1 }

"" | Out-File $LOG_FILE

Write-Host "=================================================="
Write-Host "  議事録メーカー セットアップ (Windows)"
Write-Host "=================================================="
Write-Host ""

# ------------------------------------------
# 1. Python 3
# ------------------------------------------
Log "Python 3 を確認中..."
$pythonCmd = $null
if (Get-Command python -ErrorAction SilentlyContinue) {
    $ver = python --version 2>&1
    if ($ver -match "Python 3") { $pythonCmd = "python" }
}
if (-not $pythonCmd -and (Get-Command python3 -ErrorAction SilentlyContinue)) {
    $pythonCmd = "python3"
}

if ($pythonCmd) {
    $pyVer = & $pythonCmd --version 2>&1
    Ok "Python インストール済み ($pyVer)"
} else {
    Log "Python 3 をインストール中..."
    try {
        winget install Python.Python.3.11 --accept-source-agreements --accept-package-agreements
        $pythonCmd = "python"
        Ok "Python 3 インストール完了（ターミナルを再起動してください）"
    } catch {
        Fail "Python のインストールに失敗しました。https://www.python.org/downloads/ から手動でインストールしてください。"
    }
}

# ------------------------------------------
# 2. ffmpeg
# ------------------------------------------
Log "ffmpeg を確認中..."
if (Get-Command ffmpeg -ErrorAction SilentlyContinue) {
    Ok "ffmpeg インストール済み"
} else {
    Log "ffmpeg をインストール中..."
    try {
        winget install Gyan.FFmpeg --accept-source-agreements --accept-package-agreements
        Ok "ffmpeg インストール完了"
    } catch {
        try {
            choco install ffmpeg -y
            Ok "ffmpeg インストール完了 (choco)"
        } catch {
            Warn "ffmpeg の自動インストールに失敗しました。https://ffmpeg.org/download.html から手動でインストールしてください。"
        }
    }
}

# ------------------------------------------
# 3. Python 仮想環境 & 依存関係
# ------------------------------------------
Log "Python 仮想環境を作成中..."
Set-Location $APP_DIR

if (-not (Test-Path "venv")) {
    & $pythonCmd -m venv venv
    Ok "仮想環境を作成しました"
} else {
    Ok "仮想環境は既に存在します"
}

$activateScript = Join-Path $APP_DIR "venv\Scripts\Activate.ps1"
. $activateScript

Log "Python パッケージをインストール中（数分かかります）..."
pip install --upgrade pip -q
pip install -r requirements.txt -q
Ok "Python パッケージ インストール完了"

# ------------------------------------------
# 4. Ollama
# ------------------------------------------
Log "Ollama を確認中..."
if (Get-Command ollama -ErrorAction SilentlyContinue) {
    Ok "Ollama インストール済み"
} else {
    Log "Ollama をインストール中..."
    try {
        winget install Ollama.Ollama --accept-source-agreements --accept-package-agreements
        Ok "Ollama インストール完了"
    } catch {
        Warn "Ollama の自動インストールに失敗しました。https://ollama.com/download から手動でインストールしてください。"
    }
}

# Ollama サーバー起動
Log "Ollama サーバーを起動中..."
try {
    $response = Invoke-RestMethod -Uri "http://localhost:11434/api/tags" -TimeoutSec 3 -ErrorAction SilentlyContinue
    Ok "Ollama サーバーは既に起動しています"
} catch {
    Start-Process ollama -ArgumentList "serve" -WindowStyle Hidden
    Start-Sleep -Seconds 5
    try {
        $response = Invoke-RestMethod -Uri "http://localhost:11434/api/tags" -TimeoutSec 5 -ErrorAction SilentlyContinue
        Ok "Ollama サーバー起動完了"
    } catch {
        Warn "Ollama サーバーの起動に時間がかかっています。後で ollama serve を実行してください。"
    }
}

# ------------------------------------------
# 5. LLM モデルダウンロード
# ------------------------------------------
Log "Ollama モデルを確認中..."
$models = ollama list 2>&1
if ($models -match "gemma4") {
    Ok "gemma4 モデル ダウンロード済み"
} else {
    Log "gemma4 モデルをダウンロード中（サイズはモデルにより異なります。数分〜数十分かかることがあります）..."
    ollama pull gemma4
    Ok "gemma4 モデル ダウンロード完了"
}

# ------------------------------------------
# 6. アプリ起動
# ------------------------------------------
Log "議事録メーカーを起動中..."
Set-Location $APP_DIR
. $activateScript

$serverJob = Start-Process -FilePath $pythonCmd -ArgumentList "server.py" -WorkingDirectory $APP_DIR -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 8

try {
    $health = Invoke-RestMethod -Uri "http://localhost:8000/api/health" -TimeoutSec 10 -ErrorAction SilentlyContinue
    Ok "アプリ起動完了"
} catch {
    Start-Sleep -Seconds 15
    try {
        $health = Invoke-RestMethod -Uri "http://localhost:8000/api/health" -TimeoutSec 10 -ErrorAction SilentlyContinue
        Ok "アプリ起動完了"
    } catch {
        Fail "アプリの起動に失敗しました。ログを確認してください: $LOG_FILE"
    }
}

# ------------------------------------------
# 7. ブラウザで開く
# ------------------------------------------
Start-Process "http://localhost:8000"

Write-Host ""
Write-Host "=================================================="
Write-Host "  セットアップ完了!"
Write-Host "=================================================="
Write-Host ""
Write-Host "  ブラウザで http://localhost:8000 が開きます"
Write-Host "  Ollama（ローカルLLM）がデフォルトで選択されています"
Write-Host "  API Key 不要でそのまま使えます"
Write-Host ""
Write-Host "  次回の起動方法:"
Write-Host "    cd $APP_DIR"
Write-Host "    .\venv\Scripts\Activate.ps1"
Write-Host "    python server.py"
Write-Host ""
Write-Host "  ログファイル: $LOG_FILE"
Write-Host "=================================================="
