#!/usr/bin/env python3
"""
議事録メーカー - Webサーバー
FastAPI + Whisper + ChatGPT API
"""
import os
import json
import glob
import subprocess
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, HTMLResponse
from openpyxl import load_workbook

try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# ========================================
# 設定
# ========================================

APP_DIR = Path(__file__).parent
TEMPLATE_DIR = APP_DIR / "templates"
EXCEL_TEMPLATE_DIR = APP_DIR / "excel_templates"
OUTPUT_DIR = APP_DIR / "output"
WEB_DIR = APP_DIR / "web"

# outputフォルダがなければ作成
OUTPUT_DIR.mkdir(exist_ok=True)

# ========================================
# FastAPI アプリ
# ========================================

app = FastAPI(
    title="議事録メーカー API",
    description="音声から議事録を自動生成するAPI",
    version="2.0.0"
)

# 静的ファイル配信
app.mount("/static", StaticFiles(directory=WEB_DIR / "static"), name="static")


# ========================================
# ユーティリティ関数
# ========================================

def list_templates() -> list:
    """テンプレート一覧を取得"""
    files = sorted(glob.glob(str(TEMPLATE_DIR / "*.json")))
    templates = []
    for f in files:
        with open(f, "r", encoding="utf-8") as fp:
            data = json.load(fp)
            templates.append({
                "filename": os.path.basename(f),
                "name": data.get("name", ""),
                "sections": data.get("sections", [])
            })
    return templates


def load_template(filename: str) -> dict:
    """テンプレートを読み込む"""
    path = TEMPLATE_DIR / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail=f"テンプレートが見つかりません: {filename}")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def run_whisper(audio_path: str, model_name: str = "medium") -> str:
    """Whisperで文字起こし"""
    print(f"[WHISPER] 入力ファイル: {audio_path}")
    print(f"[WHISPER] モデル: {model_name}")
    print(f"[WHISPER] 出力先: {OUTPUT_DIR}")
    
    result = subprocess.run(
        ["whisper", audio_path, "--language", "Japanese", "--model", model_name, "--output_dir", str(OUTPUT_DIR)],
        capture_output=True,
        text=True
    )
    
    print(f"[WHISPER] 終了コード: {result.returncode}")
    if result.stdout:
        print(f"[WHISPER] stdout: {result.stdout[:500]}")
    if result.stderr:
        print(f"[WHISPER] stderr: {result.stderr[:500]}")
    
    if result.returncode != 0:
        raise HTTPException(status_code=500, detail=f"Whisper実行エラー: {result.stderr}")
    
    # 生成されたテキストファイルを探す
    base = os.path.splitext(os.path.basename(audio_path))[0]
    txt_path = OUTPUT_DIR / f"{base}.txt"
    
    print(f"[WHISPER] 期待するファイル: {txt_path}")
    
    # ファイルが見つからない場合、outputフォルダ内の最新の.txtを探す
    if not txt_path.exists():
        print(f"[WHISPER] ファイルが見つかりません。outputフォルダを検索...")
        txt_files = list(OUTPUT_DIR.glob("*.txt"))
        print(f"[WHISPER] 見つかった.txtファイル: {txt_files}")
        
        if txt_files:
            # 最新のファイルを使用
            txt_path = max(txt_files, key=lambda p: p.stat().st_mtime)
            print(f"[WHISPER] 最新のファイルを使用: {txt_path}")
        else:
            raise HTTPException(status_code=500, detail="文字起こしファイルが生成されませんでした")
    
    with open(txt_path, "r", encoding="utf-8") as f:
        return f.read().strip()


def build_prompt(tpl: dict, transcript: str) -> str:
    """ChatGPT用プロンプトを構築"""
    sections = "\n".join([f'- "{s["key"]}": {s["label"]}の報告内容' for s in tpl["sections"]])
    rules = "\n".join([f"- {r}" for r in tpl.get("chatgpt_prompt", {}).get("style_rules", [])])
    
    # 期待するJSONの例を作成
    example_json = {s["key"]: f"（{s['label']}の内容）" for s in tpl["sections"]}
    import json
    example_str = json.dumps(example_json, ensure_ascii=False, indent=2)
    
    return f"""あなたは会議議事録作成のプロです。以下の文字起こしから、各項目の内容を抽出してください。

【テンプレート名】
{tpl["name"]}

【抽出する項目とキー】
{sections}

【ルール】
{rules}
- 該当する内容がない場合は空文字 "" を設定

【出力形式（厳守）】
以下の形式のJSONのみを返してください。コードブロック（```）は絶対に使わないでください。

{example_str}

【会議の文字起こし】
{transcript}
"""


def call_chatgpt(api_key: str, prompt: str, model: str = "gpt-4o-mini") -> dict:
    """ChatGPT APIを呼び出し"""
    if not OPENAI_AVAILABLE:
        raise HTTPException(status_code=500, detail="OpenAIライブラリがインストールされていません")
    
    client = OpenAI(api_key=api_key)
    
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": "あなたは議事録作成のプロフェッショナルです。会議の内容を正確かつ簡潔にまとめてください。"},
            {"role": "user", "content": prompt}
        ],
        temperature=0.3,
        max_tokens=4000
    )
    
    content = response.choices[0].message.content
    print(f"[DEBUG] ChatGPT raw response: {content[:500]}...")
    
    # JSONをパース
    if "```json" in content:
        content = content.split("```json")[1]
    if "```" in content:
        content = content.split("```")[0]
    content = content.strip()
    
    try:
        parsed = json.loads(content)
        print(f"[DEBUG] Parsed JSON keys: {list(parsed.keys())}")
        
        # ネストされた "sections" キーがある場合は展開
        if "sections" in parsed and isinstance(parsed["sections"], dict):
            parsed = parsed["sections"]
            print(f"[DEBUG] Extracted from 'sections': {list(parsed.keys())}")
        
        return parsed
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=500, detail=f"ChatGPTの応答をJSONとして解析できません: {e}\n応答: {content[:200]}")


def write_excel(tpl: dict, data: dict, output_name: str) -> Path:
    """Excelファイルを生成"""
    from openpyxl.styles import Alignment
    
    xlsx_path = APP_DIR / tpl["excel_template"]
    
    if not xlsx_path.exists():
        raise HTTPException(status_code=404, detail=f"Excelテンプレートが見つかりません: {xlsx_path}")
    
    wb = load_workbook(xlsx_path)
    ws = wb[tpl["sheet"]]
    
    print(f"[EXCEL] === 書き込み開始 ===")
    print(f"[EXCEL] テンプレート: {xlsx_path}")
    print(f"[EXCEL] データキー: {list(data.keys())}")
    print(f"[EXCEL] セクション: {[(s['key'], s['cell']) for s in tpl['sections']]}")
    
    written_count = 0
    for s in tpl["sections"]:
        key = s["key"]
        cell_ref = s["cell"]
        prefix = s.get("prefix", "")  # ラベルのプレフィックス
        value = data.get(key, "")
        
        # 値が文字列でない場合は文字列に変換
        if value is not None and not isinstance(value, str):
            value = str(value)
        
        if not value:
            print(f"[EXCEL] スキップ: '{key}' (値なし)")
            continue
        
        # プレフィックス（ラベル）+ データを結合して書き込む
        full_value = prefix + value if prefix else value
        ws[cell_ref] = full_value
        
        # テキスト折り返しと上揃えを設定
        ws[cell_ref].alignment = Alignment(wrap_text=True, vertical='top')
        
        print(f"[EXCEL] ✅ {cell_ref} ← '{key}': {full_value[:50]}...")
        written_count += 1
    
    print(f"[EXCEL] === 書き込み完了: {written_count}件 ===")
    
    output_path = OUTPUT_DIR / output_name
    wb.save(output_path)
    
    return output_path


# ========================================
# APIエンドポイント
# ========================================

@app.get("/", response_class=HTMLResponse)
async def index():
    """メインページ"""
    index_path = WEB_DIR / "index.html"
    with open(index_path, "r", encoding="utf-8") as f:
        return f.read()


@app.get("/template-manager", response_class=HTMLResponse)
async def template_manager():
    """テンプレート管理ページ"""
    page_path = WEB_DIR / "template-manager.html"
    with open(page_path, "r", encoding="utf-8") as f:
        return f.read()


@app.get("/api/templates")
async def get_templates():
    """テンプレート一覧を取得"""
    return {"templates": list_templates()}


@app.post("/api/process")
async def process_audio(
    file: UploadFile = File(...),
    whisper_model: str = Form("medium"),
    gpt_model: str = Form("gpt-4o-mini"),
    template: str = Form("shozokucho.json"),
    api_key: str = Form(...)
):
    """
    音声ファイルを処理して議事録を生成
    """
    tmp_path = None
    
    try:
        # 一時ファイルに保存
        suffix = os.path.splitext(file.filename)[1]
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name
        
        print(f"[INFO] 音声ファイル保存: {tmp_path} ({len(content)} bytes)")
        
        # 1. 文字起こし
        print(f"[INFO] Whisper開始 (モデル: {whisper_model})")
        transcript = run_whisper(tmp_path, whisper_model)
        print(f"[INFO] 文字起こし完了: {len(transcript)}文字")
        
        # 2. テンプレート読み込み
        tpl = load_template(template)
        print(f"[INFO] テンプレート読み込み: {tpl['name']}")
        
        # 3. プロンプト生成
        prompt = build_prompt(tpl, transcript)
        
        # 4. ChatGPT API呼び出し
        print(f"[INFO] ChatGPT API呼び出し (モデル: {gpt_model})")
        data = call_chatgpt(api_key, prompt, gpt_model)
        print(f"[INFO] ChatGPT応答受信")
        
        # 5. Excel生成
        base = os.path.splitext(file.filename)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"{tpl['name']}_{base}_{timestamp}.xlsx"
        
        output_path = write_excel(tpl, data, output_name)
        print(f"[INFO] Excel生成完了: {output_path}")
        
        return {
            "success": True,
            "filename": output_name,
            "path": str(OUTPUT_DIR),
            "transcript_length": len(transcript),
            "transcript": transcript,  # 文字起こし結果を追加
            "summary": data,  # AI要約結果を追加
            "sections": list(data.keys())
        }
        
    except HTTPException:
        raise
    except Exception as e:
        print(f"[ERROR] {type(e).__name__}: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        # 一時ファイル削除
        if tmp_path and os.path.exists(tmp_path):
            os.unlink(tmp_path)


@app.get("/api/download/{filename}")
async def download_file(filename: str):
    """生成したExcelファイルをダウンロード"""
    file_path = OUTPUT_DIR / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="ファイルが見つかりません")
    
    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.post("/api/open-folder")
async def open_folder():
    """出力フォルダを開く"""
    subprocess.run(["open", str(OUTPUT_DIR)])
    return {"success": True}


@app.get("/api/health")
async def health_check():
    """ヘルスチェック"""
    return {
        "status": "ok",
        "openai_available": OPENAI_AVAILABLE,
        "templates_count": len(list_templates())
    }


# ========================================
# テンプレート管理API
# ========================================

from template_manager import analyze_excel_template, create_template_config, save_template_config


@app.post("/api/templates/upload")
async def upload_template_excel(file: UploadFile = File(...)):
    """
    Excelテンプレートをアップロードして解析
    """
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Excelファイル(.xlsx, .xls)のみ対応しています")
    
    # Excelファイルを保存
    excel_path = EXCEL_TEMPLATE_DIR / file.filename
    
    content = await file.read()
    with open(excel_path, "wb") as f:
        f.write(content)
    
    print(f"[INFO] Excelテンプレート保存: {excel_path}")
    
    # 解析
    try:
        analysis = analyze_excel_template(excel_path)
        return {
            "success": True,
            "filename": file.filename,
            "analysis": analysis
        }
    except Exception as e:
        # エラー時はファイルを削除
        if excel_path.exists():
            excel_path.unlink()
        raise HTTPException(status_code=500, detail=f"Excel解析エラー: {e}")


@app.post("/api/templates/create")
async def create_template(
    excel_filename: str = Form(...),
    template_name: str = Form(...),
    sections: str = Form(...),  # JSON文字列
    style_rules: str = Form(None)  # JSON文字列（オプション）
):
    """
    テンプレート設定を作成
    """
    excel_path = EXCEL_TEMPLATE_DIR / excel_filename
    
    if not excel_path.exists():
        raise HTTPException(status_code=404, detail=f"Excelファイルが見つかりません: {excel_filename}")
    
    try:
        sections_list = json.loads(sections)
        rules_list = json.loads(style_rules) if style_rules else None
        
        config = create_template_config(
            excel_path=excel_path,
            template_name=template_name,
            sections=sections_list,
            style_rules=rules_list
        )
        
        # 設定ファイルを保存
        config_path = save_template_config(config, TEMPLATE_DIR)
        
        print(f"[INFO] テンプレート設定保存: {config_path}")
        
        return {
            "success": True,
            "config": config,
            "config_file": config_path.name
        }
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"JSONパースエラー: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/templates/{template_name}/preview")
async def preview_template(template_name: str):
    """
    テンプレートのプレビュー情報を取得
    """
    tpl = load_template(template_name)
    excel_path = APP_DIR / tpl["excel_template"]
    
    if not excel_path.exists():
        raise HTTPException(status_code=404, detail="Excelテンプレートが見つかりません")
    
    wb = load_workbook(excel_path)
    ws = wb[tpl["sheet"]]
    
    # セル情報を取得
    cells_info = []
    for section in tpl["sections"]:
        cell = ws[section["cell"]]
        cells_info.append({
            "key": section["key"],
            "label": section["label"],
            "cell": section["cell"],
            "current_value": str(cell.value) if cell.value else "",
            "row": cell.row,
            "col": cell.column
        })
    
    return {
        "name": tpl["name"],
        "sheet": tpl["sheet"],
        "sections": cells_info,
        "style_rules": tpl.get("chatgpt_prompt", {}).get("style_rules", [])
    }


@app.delete("/api/templates/{template_name}")
async def delete_template(template_name: str):
    """
    テンプレートを削除
    """
    template_path = TEMPLATE_DIR / template_name
    
    if not template_path.exists():
        raise HTTPException(status_code=404, detail="テンプレートが見つかりません")
    
    # テンプレート設定を読み込んでExcelファイルも削除
    try:
        with open(template_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        excel_path = APP_DIR / config.get("excel_template", "")
        
        # 設定ファイル削除
        template_path.unlink()
        
        # Excelファイル削除（他のテンプレートで使用されていない場合）
        if excel_path.exists():
            # 他のテンプレートで使用されているかチェック
            other_templates = [t for t in list_templates() if t != template_name]
            is_used = False
            for other in other_templates:
                other_config = load_template(other)
                if other_config.get("excel_template") == config.get("excel_template"):
                    is_used = True
                    break
            
            if not is_used:
                excel_path.unlink()
        
        return {"success": True, "deleted": template_name}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ========================================
# 起動
# ========================================

if __name__ == "__main__":
    import uvicorn
    print("=" * 50)
    print("📝 議事録メーカー Web版")
    print("=" * 50)
    print(f"🌐 http://localhost:8000")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8000)
