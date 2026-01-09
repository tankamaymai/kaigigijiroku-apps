/**
 * 議事録メーカー - Minutes Maker
 * Frontend JavaScript
 */

// ========================================
// 状態管理
// ========================================

const state = {
    apiKey: '',
    whisperModel: 'medium',
    gptModel: 'gpt-4o-mini',
    selectedTemplate: 'shozokucho.json',
    selectedFile: null,
    isProcessing: false
};

// ========================================
// DOM要素取得（遅延初期化）
// ========================================

let elements = {};

function initElements() {
    elements = {
        // 設定
        apiKeyInput: document.getElementById('api-key'),
        whisperModelSelect: document.getElementById('whisper-model'),
        gptModelSelect: document.getElementById('gpt-model'),
        
        // アップロード
        uploadZone: document.getElementById('upload-zone'),
        audioFileInput: document.getElementById('audio-file'),
        fileInfo: document.getElementById('file-info'),
        fileName: document.getElementById('file-name'),
        fileSize: document.getElementById('file-size'),
        uploadContent: document.querySelector('.upload-content'),
        
        // 実行
        runBtn: document.getElementById('run-btn'),
        
        // プログレス
        progressCard: document.getElementById('progress-card'),
        progressTitle: document.getElementById('progress-title'),
        progressPercent: document.getElementById('progress-percent'),
        progressFill: document.getElementById('progress-fill'),
        
        // 結果
        resultCard: document.getElementById('result-card'),
        resultFilename: document.getElementById('result-filename'),
        resultPath: document.getElementById('result-path'),
        downloadBtn: document.getElementById('download-btn'),
        openFolderBtn: document.getElementById('open-folder-btn'),
        
        // ログ
        logContent: document.getElementById('log-content')
    };
}

// ========================================
// ユーティリティ
// ========================================

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function formatTime(date = new Date()) {
    return date.toTimeString().slice(0, 8);
}

function log(message, type = 'info') {
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.innerHTML = `
        <span class="log-time">${formatTime()}</span>
        <span class="log-message">${message}</span>
    `;
    elements.logContent.appendChild(entry);
    elements.logContent.scrollTop = elements.logContent.scrollHeight;
}

function clearLog() {
    elements.logContent.innerHTML = '';
    log('ログをクリアしました', 'info');
}

// ========================================
// 設定管理
// ========================================

function loadSettings() {
    const saved = localStorage.getItem('minutesMakerSettings');
    if (saved) {
        const settings = JSON.parse(saved);
        state.apiKey = settings.apiKey || '';
        state.whisperModel = settings.whisperModel || 'medium';
        state.gptModel = settings.gptModel || 'gpt-4o-mini';
        
        if (elements.apiKeyInput) elements.apiKeyInput.value = state.apiKey;
        if (elements.whisperModelSelect) elements.whisperModelSelect.value = state.whisperModel;
        if (elements.gptModelSelect) elements.gptModelSelect.value = state.gptModel;
    }
}

function saveSettings() {
    state.apiKey = elements.apiKeyInput.value;
    state.whisperModel = elements.whisperModelSelect.value;
    state.gptModel = elements.gptModelSelect.value;
    
    localStorage.setItem('minutesMakerSettings', JSON.stringify({
        apiKey: state.apiKey,
        whisperModel: state.whisperModel,
        gptModel: state.gptModel
    }));
    
    log('設定を保存しました', 'success');
    alert('設定を保存しました！');
}

// ========================================
// ファイルアップロード
// ========================================

function setupFileUpload() {
    const zone = elements.uploadZone;
    const input = elements.audioFileInput;
    
    if (!zone || !input) {
        console.error('Upload elements not found');
        return;
    }
    
    // ブラウズリンクのクリック
    const browseLink = zone.querySelector('.browse-link');
    if (browseLink) {
        browseLink.addEventListener('click', (e) => {
            e.stopPropagation();
            input.click();
        });
    }
    
    // アップロードゾーン全体のクリック
    zone.addEventListener('click', (e) => {
        // ファイル情報が表示されている場合は、削除ボタン以外のクリックでファイル選択
        if (state.selectedFile) {
            // 削除ボタンのクリックは除外
            if (e.target.classList.contains('remove-file')) return;
        }
        input.click();
    });
    
    // ファイル選択時
    input.addEventListener('change', (e) => {
        console.log('File input changed', e.target.files);
        if (e.target.files && e.target.files.length > 0) {
            handleFileSelect(e.target.files[0]);
        }
    });
    
    // ドラッグ&ドロップ
    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.add('dragover');
    });
    
    zone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.remove('dragover');
    });
    
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.remove('dragover');
        
        console.log('File dropped', e.dataTransfer.files);
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleFileSelect(e.dataTransfer.files[0]);
        }
    });
    
    // ファイル削除ボタン
    const removeBtn = zone.querySelector('.remove-file');
    if (removeBtn) {
        removeBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            removeFile();
        });
    }
}

function handleFileSelect(file) {
    console.log('Handling file:', file.name, file.type);
    
    // 対応フォーマットチェック
    const validExtensions = ['.m4a', '.mp3', '.wav', '.mp4', '.webm', '.ogg', '.flac'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    
    if (!validExtensions.includes(ext)) {
        log(`非対応のファイル形式です: ${ext}`, 'error');
        alert('対応していないファイル形式です。\n対応形式: m4a, mp3, wav, mp4, webm, ogg, flac');
        return;
    }
    
    state.selectedFile = file;
    
    // UI更新
    elements.fileName.textContent = file.name;
    elements.fileSize.textContent = formatFileSize(file.size);
    
    // アップロードコンテンツを非表示、ファイル情報を表示
    if (elements.uploadContent) elements.uploadContent.classList.add('hidden');
    if (elements.fileInfo) elements.fileInfo.classList.remove('hidden');
    
    // 実行ボタン有効化
    updateRunButton();
    
    log(`ファイル選択: ${file.name} (${formatFileSize(file.size)})`, 'info');
}

function removeFile() {
    state.selectedFile = null;
    elements.audioFileInput.value = '';
    
    // アップロードコンテンツを表示、ファイル情報を非表示
    if (elements.uploadContent) elements.uploadContent.classList.remove('hidden');
    if (elements.fileInfo) elements.fileInfo.classList.add('hidden');
    
    updateRunButton();
    log('ファイルを削除しました', 'info');
}

function updateRunButton() {
    const canRun = state.selectedFile && !state.isProcessing;
    elements.runBtn.disabled = !canRun;
}

// ========================================
// テンプレート選択
// ========================================

async function loadTemplates() {
    try {
        const response = await fetch('/api/templates');
        if (!response.ok) throw new Error('テンプレート取得失敗');
        
        const data = await response.json();
        const templateList = document.querySelector('.template-list');
        
        if (!templateList || !data.templates || data.templates.length === 0) {
            return;
        }
        
        // テンプレートアイコンのマッピング
        const icons = ['🏥', '📋', '📝', '📊', '📄', '🗂️', '📑', '🗒️'];
        
        // テンプレートリストをクリア
        templateList.innerHTML = '';
        
        // テンプレートを追加
        data.templates.forEach((tpl, index) => {
            const item = document.createElement('div');
            item.className = 'template-item' + (index === 0 ? ' active' : '');
            item.dataset.template = tpl.filename;
            item.innerHTML = `
                <div class="template-icon">${icons[index % icons.length]}</div>
                <div class="template-info">
                    <h3>${tpl.name}</h3>
                    <p>${tpl.sections.length}項目</p>
                </div>
                <span class="check-mark">✓</span>
            `;
            templateList.appendChild(item);
        });
        
        // 最初のテンプレートを選択
        if (data.templates.length > 0) {
            state.selectedTemplate = data.templates[0].filename;
        }
        
        // クリックイベントを設定
        setupTemplateSelection();
        
        log(`テンプレート ${data.templates.length}件を読み込みました`, 'info');
        
    } catch (error) {
        console.error('テンプレート読み込みエラー:', error);
    }
}

function setupTemplateSelection() {
    const items = document.querySelectorAll('.template-item');
    
    items.forEach(item => {
        item.addEventListener('click', () => {
            // 全てのアクティブを解除
            items.forEach(i => i.classList.remove('active'));
            // クリックしたアイテムをアクティブに
            item.classList.add('active');
            
            state.selectedTemplate = item.dataset.template;
            log(`テンプレート選択: ${item.querySelector('h3').textContent}`, 'info');
        });
    });
}

// ========================================
// プログレス表示
// ========================================

function showProgress() {
    elements.progressCard.classList.remove('hidden');
    elements.resultCard.classList.add('hidden');
}

function hideProgress() {
    elements.progressCard.classList.add('hidden');
}

function updateProgress(percent, title, currentStep) {
    elements.progressPercent.textContent = `${percent}%`;
    elements.progressTitle.textContent = title;
    elements.progressFill.style.width = `${percent}%`;
    
    // ステップ表示更新
    for (let i = 1; i <= 3; i++) {
        const step = document.getElementById(`step-${i}`);
        if (step) {
            step.classList.remove('active', 'completed');
            
            if (i < currentStep) {
                step.classList.add('completed');
            } else if (i === currentStep) {
                step.classList.add('active');
            }
        }
    }
}

// ========================================
// 結果表示
// ========================================

function showResult(filename, path) {
    hideProgress();
    elements.resultCard.classList.remove('hidden');
    elements.resultFilename.textContent = filename;
    elements.resultPath.textContent = path;
}

// ========================================
// API通信
// ========================================

async function runPipeline() {
    if (!state.selectedFile) {
        alert('ファイルを選択してください');
        return;
    }
    
    if (!state.apiKey) {
        alert('OpenAI API Keyを設定してください');
        return;
    }
    
    state.isProcessing = true;
    updateRunButton();
    showProgress();
    
    log('='.repeat(40), 'info');
    log('🚀 議事録作成を開始します', 'info');
    
    try {
        // Step 1: 文字起こし
        updateProgress(10, '🎙️ 音声を文字起こし中...', 1);
        log(`🎙️ Whisperで文字起こし開始 (モデル: ${state.whisperModel})`, 'info');
        
        const formData = new FormData();
        formData.append('file', state.selectedFile);
        formData.append('whisper_model', state.whisperModel);
        formData.append('gpt_model', state.gptModel);
        formData.append('template', state.selectedTemplate);
        formData.append('api_key', state.apiKey);
        
        updateProgress(30, '🎙️ 文字起こし中...', 1);
        
        const response = await fetch('/api/process', {
            method: 'POST',
            body: formData
        });
        
        // レスポンスのテキストを取得
        const responseText = await response.text();
        console.log('Response status:', response.status);
        console.log('Response text:', responseText);
        
        if (!response.ok) {
            let errorMessage = 'サーバーエラー';
            try {
                const error = JSON.parse(responseText);
                errorMessage = error.detail || errorMessage;
            } catch (e) {
                errorMessage = responseText || errorMessage;
            }
            throw new Error(errorMessage);
        }
        
        // JSONをパース
        let result;
        try {
            result = JSON.parse(responseText);
        } catch (e) {
            throw new Error('サーバーからの応答を解析できませんでした');
        }
        
        // Step 2: AI要約完了
        updateProgress(70, '🤖 AI要約完了', 2);
        log('✅ ChatGPT応答を受信', 'success');
        
        // Step 3: Excel生成完了
        updateProgress(100, '📊 Excel生成完了', 3);
        log(`✅ Excel出力完了: ${result.filename}`, 'success');
        
        log('='.repeat(40), 'info');
        log('🎉 議事録作成が完了しました！', 'success');
        
        // 文字起こし結果を表示
        if (result.transcript) {
            const transcriptText = document.getElementById('transcript-text');
            if (transcriptText) {
                transcriptText.textContent = result.transcript;
            }
            log(`📝 文字起こし: ${result.transcript_length}文字`, 'info');
        }
        
        // AI要約結果を表示
        if (result.summary) {
            const summaryContent = document.getElementById('summary-content');
            if (summaryContent) {
                let summaryHtml = '';
                for (const [key, value] of Object.entries(result.summary)) {
                    summaryHtml += `
                        <div class="summary-item">
                            <div class="summary-item-label">${key}</div>
                            <div class="summary-item-value">${value || '(データなし)'}</div>
                        </div>
                    `;
                }
                summaryContent.innerHTML = summaryHtml;
            }
        }
        
        showResult(result.filename, result.path);
        
        // ダウンロードボタンにURLを設定
        elements.downloadBtn.onclick = () => {
            window.location.href = `/api/download/${result.filename}`;
        };
        
    } catch (error) {
        log(`❌ エラー: ${error.message}`, 'error');
        hideProgress();
        alert(`エラー: ${error.message}`);
    } finally {
        state.isProcessing = false;
        updateRunButton();
    }
}

// ========================================
// デモモード（バックエンドなしでのテスト用）
// ========================================

async function runDemoMode() {
    if (!state.selectedFile) {
        alert('ファイルを選択してください');
        return;
    }
    
    state.isProcessing = true;
    updateRunButton();
    showProgress();
    
    log('='.repeat(40), 'info');
    log('🚀 議事録作成を開始します', 'info');
    
    // Step 1
    updateProgress(10, '🎙️ 音声を文字起こし中...', 1);
    log(`🎙️ Whisperで文字起こし開始 (モデル: ${state.whisperModel})`, 'info');
    await sleep(2000);
    
    updateProgress(30, '🎙️ 文字起こし中...', 1);
    log('📝 文字起こし処理中...', 'info');
    await sleep(1500);
    
    log('✅ 文字起こし完了 (2,450文字)', 'success');
    
    // Step 2
    updateProgress(50, '🤖 AI要約中...', 2);
    log('🤖 ChatGPT APIで議事録作成中...', 'info');
    await sleep(2000);
    
    log('✅ ChatGPT応答を受信', 'success');
    updateProgress(70, '🤖 AI要約完了', 2);
    
    // Step 3
    updateProgress(85, '📊 Excel生成中...', 3);
    log('📊 Excelファイル生成中...', 'info');
    await sleep(1000);
    
    const filename = `所属長会議_${state.selectedFile.name.split('.')[0]}_${new Date().toISOString().slice(0,10)}.xlsx`;
    
    updateProgress(100, '📊 Excel生成完了', 3);
    log(`✅ Excel出力完了: ${filename}`, 'success');
    
    log('='.repeat(40), 'info');
    log('🎉 議事録作成が完了しました！', 'success');
    
    showResult(filename, 'output/');
    
    state.isProcessing = false;
    updateRunButton();
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ========================================
// イベントリスナー設定
// ========================================

function setupEventListeners() {
    // 設定保存
    const saveBtn = document.querySelector('.save-settings');
    if (saveBtn) {
        saveBtn.addEventListener('click', saveSettings);
    }
    
    // API Key入力時に状態更新
    if (elements.apiKeyInput) {
        elements.apiKeyInput.addEventListener('input', () => {
            state.apiKey = elements.apiKeyInput.value;
            updateRunButton();
        });
    }
    
    // パスワード表示切替
    const toggleBtn = document.querySelector('.toggle-visibility');
    if (toggleBtn) {
        toggleBtn.addEventListener('click', () => {
            const input = elements.apiKeyInput;
            input.type = input.type === 'password' ? 'text' : 'password';
        });
    }
    
    // 実行ボタン
    if (elements.runBtn) {
        elements.runBtn.addEventListener('click', () => {
            // API Keyがあれば本番モード、なければデモモード
            if (state.apiKey) {
                runPipeline();
            } else {
                runDemoMode();
            }
        });
    }
    
    // ログクリア
    const clearLogBtn = document.querySelector('.clear-log');
    if (clearLogBtn) {
        clearLogBtn.addEventListener('click', clearLog);
    }
    
    // フォルダを開く
    if (elements.openFolderBtn) {
        elements.openFolderBtn.addEventListener('click', () => {
            fetch('/api/open-folder', { method: 'POST' });
            log('📁 出力フォルダを開きました', 'info');
        });
    }
}

// ========================================
// 初期化
// ========================================

async function init() {
    console.log('Initializing Minutes Maker...');
    
    initElements();
    loadSettings();
    setupFileUpload();
    await loadTemplates();  // テンプレートを動的に読み込む
    setupEventListeners();
    updateRunButton();
    
    log('システム準備完了', 'success');
    console.log('Minutes Maker initialized successfully');
}

// DOMContentLoaded
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
} else {
    init();
}
