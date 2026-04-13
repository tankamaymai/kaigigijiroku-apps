#!/bin/bash
set -e

# ========================================
# 議事録メーカー - Mac セットアップスクリプト
# 全自動でアプリが使える状態にします
# ========================================

APP_DIR="$(cd "$(dirname "$0")/.." && pwd)"
LOG_FILE="$APP_DIR/setup.log"

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BLUE='\033[0;34m'
NC='\033[0m'

log()  { echo -e "${BLUE}[INFO]${NC} $1" | tee -a "$LOG_FILE"; }
ok()   { echo -e "${GREEN}[OK]${NC}   $1" | tee -a "$LOG_FILE"; }
warn() { echo -e "${YELLOW}[WARN]${NC} $1" | tee -a "$LOG_FILE"; }
fail() { echo -e "${RED}[FAIL]${NC} $1" | tee -a "$LOG_FILE"; exit 1; }

echo "" > "$LOG_FILE"

echo "=================================================="
echo "  議事録メーカー セットアップ (Mac)"
echo "=================================================="
echo ""

# ------------------------------------------
# 1. Homebrew
# ------------------------------------------
log "Homebrew を確認中..."
if command -v brew &>/dev/null; then
    ok "Homebrew インストール済み"
else
    log "Homebrew をインストール中..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    # Apple Silicon の場合 PATH を通す
    if [ -f /opt/homebrew/bin/brew ]; then
        eval "$(/opt/homebrew/bin/brew shellenv)"
        echo 'eval "$(/opt/homebrew/bin/brew shellenv)"' >> ~/.zprofile
    fi
    ok "Homebrew インストール完了"
fi

# ------------------------------------------
# 2. Python 3
# ------------------------------------------
log "Python 3 を確認中..."
if command -v python3 &>/dev/null; then
    PY_VERSION=$(python3 --version 2>&1)
    ok "Python インストール済み ($PY_VERSION)"
else
    log "Python 3 をインストール中..."
    brew install python@3.11
    ok "Python 3 インストール完了"
fi

# ------------------------------------------
# 3. ffmpeg
# ------------------------------------------
log "ffmpeg を確認中..."
if command -v ffmpeg &>/dev/null; then
    ok "ffmpeg インストール済み"
else
    log "ffmpeg をインストール中..."
    brew install ffmpeg
    ok "ffmpeg インストール完了"
fi

# ------------------------------------------
# 4. Python 仮想環境 & 依存関係
# ------------------------------------------
log "Python 仮想環境を作成中..."
cd "$APP_DIR"

if [ ! -d "venv" ]; then
    python3 -m venv venv
    ok "仮想環境を作成しました"
else
    ok "仮想環境は既に存在します"
fi

source venv/bin/activate

log "Python パッケージをインストール中（数分かかります）..."
pip install --upgrade pip -q
pip install -r requirements.txt -q
ok "Python パッケージ インストール完了"

# ------------------------------------------
# 5. Ollama
# ------------------------------------------
log "Ollama を確認中..."
if command -v ollama &>/dev/null; then
    ok "Ollama インストール済み"
else
    log "Ollama をインストール中..."
    brew install ollama
    ok "Ollama インストール完了"
fi

# Ollama サーバー起動
log "Ollama サーバーを起動中..."
if curl -s http://localhost:11434/api/tags &>/dev/null; then
    ok "Ollama サーバーは既に起動しています"
else
    ollama serve &>/dev/null &
    sleep 3
    if curl -s http://localhost:11434/api/tags &>/dev/null; then
        ok "Ollama サーバー起動完了"
    else
        warn "Ollama サーバーの起動に時間がかかっています。後で ollama serve を実行してください。"
    fi
fi

# ------------------------------------------
# 6. LLM モデルダウンロード
# ------------------------------------------
log "Ollama モデルを確認中..."
if ollama list 2>/dev/null | grep -q "gemma4"; then
    ok "gemma4 モデル ダウンロード済み"
else
    log "gemma4 モデルをダウンロード中（サイズはモデルにより異なります。数分〜数十分かかることがあります）..."
    ollama pull gemma4
    ok "gemma4 モデル ダウンロード完了"
fi

# ------------------------------------------
# 7. アプリ起動
# ------------------------------------------
log "議事録メーカーを起動中..."

# 既にポート8000が使われていれば停止
if lsof -ti :8000 &>/dev/null; then
    warn "ポート8000が使用中のため、既存プロセスを停止します"
    lsof -ti :8000 | xargs kill -9 2>/dev/null || true
    sleep 2
fi

cd "$APP_DIR"
source venv/bin/activate
python server.py &
SERVER_PID=$!
sleep 5

if curl -s http://localhost:8000/api/health &>/dev/null; then
    ok "アプリ起動完了"
else
    sleep 10
    if curl -s http://localhost:8000/api/health &>/dev/null; then
        ok "アプリ起動完了"
    else
        fail "アプリの起動に失敗しました。ログを確認してください: $LOG_FILE"
    fi
fi

# ------------------------------------------
# 8. ブラウザで開く
# ------------------------------------------
open http://localhost:8000

echo ""
echo "=================================================="
echo "  セットアップ完了!"
echo "=================================================="
echo ""
echo "  ブラウザで http://localhost:8000 が開きます"
echo "  Ollama（ローカルLLM）がデフォルトで選択されています"
echo "  API Key 不要でそのまま使えます"
echo ""
echo "  次回の起動方法:"
echo "    cd $APP_DIR"
echo "    source venv/bin/activate"
echo "    python server.py"
echo ""
echo "  ログファイル: $LOG_FILE"
echo "=================================================="
