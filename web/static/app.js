/**
 * 議事録メーカー - Minutes Maker
 * Frontend JavaScript
 */

// ========================================
// 状態管理
// ========================================

const state = {
    aiProvider: 'ollama',
    apiKey: '',
    geminiApiKey: '',
    ollamaEndpoint: 'http://localhost:11434',
    ollamaModel: '',
    whisperModel: 'medium',
    gptModel: 'gpt-4o-mini',
    geminiModel: 'gemini-2.0-flash',
    selectedTemplate: 'shozokucho.json',
    outputFormat: 'excel',
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
        aiProviderSelect: document.getElementById('ai-provider'),
        apiKeyInput: document.getElementById('api-key'),
        geminiApiKeyInput: document.getElementById('gemini-api-key'),
        ollamaEndpointInput: document.getElementById('ollama-endpoint'),
        ollamaModelSelect: document.getElementById('ollama-model'),
        ollamaRefreshBtn: document.getElementById('ollama-refresh'),
        whisperModelSelect: document.getElementById('whisper-model'),
        gptModelSelect: document.getElementById('gpt-model'),
        geminiModelSelect: document.getElementById('gemini-model'),
        openaiSettings: document.getElementById('openai-settings'),
        geminiSettings: document.getElementById('gemini-settings'),
        ollamaSettings: document.getElementById('ollama-settings'),
        outputFormatSelect: document.getElementById('output-format'),
        dictionaryText: document.getElementById('dictionary-text'),
        saveDictionaryBtn: document.querySelector('.save-dictionary'),
        
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
        state.aiProvider = settings.aiProvider || 'ollama';
        state.apiKey = settings.apiKey || '';
        state.geminiApiKey = settings.geminiApiKey || '';
        state.ollamaEndpoint = settings.ollamaEndpoint || 'http://localhost:11434';
        state.ollamaModel = settings.ollamaModel || '';
        state.whisperModel = settings.whisperModel || 'medium';
        state.gptModel = settings.gptModel || 'gpt-4o-mini';
        state.geminiModel = settings.geminiModel || 'gemini-2.0-flash';
        state.outputFormat = settings.outputFormat || 'excel';
        
        if (elements.aiProviderSelect) elements.aiProviderSelect.value = state.aiProvider;
        if (elements.apiKeyInput) elements.apiKeyInput.value = state.apiKey;
        if (elements.geminiApiKeyInput) elements.geminiApiKeyInput.value = state.geminiApiKey;
        if (elements.ollamaEndpointInput) elements.ollamaEndpointInput.value = state.ollamaEndpoint;
        if (elements.whisperModelSelect) elements.whisperModelSelect.value = state.whisperModel;
        if (elements.gptModelSelect) elements.gptModelSelect.value = state.gptModel;
        if (elements.geminiModelSelect) elements.geminiModelSelect.value = state.geminiModel;
        if (elements.outputFormatSelect) elements.outputFormatSelect.value = state.outputFormat;
    }
    switchProvider(state.aiProvider);
}

function saveSettings() {
    state.aiProvider = elements.aiProviderSelect.value;
    state.apiKey = elements.apiKeyInput.value;
    state.geminiApiKey = elements.geminiApiKeyInput.value;
    state.ollamaEndpoint = elements.ollamaEndpointInput.value;
    state.ollamaModel = elements.ollamaModelSelect.value;
    state.whisperModel = elements.whisperModelSelect.value;
    state.gptModel = elements.gptModelSelect.value;
    state.geminiModel = elements.geminiModelSelect.value;
    state.outputFormat = elements.outputFormatSelect ? elements.outputFormatSelect.value : 'excel';
    
    localStorage.setItem('minutesMakerSettings', JSON.stringify({
        aiProvider: state.aiProvider,
        apiKey: state.apiKey,
        geminiApiKey: state.geminiApiKey,
        ollamaEndpoint: state.ollamaEndpoint,
        ollamaModel: state.ollamaModel,
        whisperModel: state.whisperModel,
        gptModel: state.gptModel,
        geminiModel: state.geminiModel,
        outputFormat: state.outputFormat
    }));
    
    log('設定を保存しました', 'success');
    alert('設定を保存しました！');
}

async function loadDictionary() {
    try {
        const res = await fetch('/api/dictionary');
        const data = await res.json();
        const entries = data.entries || {};
        const lines = Object.entries(entries).map(([k, v]) => `${k},${v}`);
        if (elements.dictionaryText) {
            elements.dictionaryText.value = lines.join('\n');
        }
    } catch (e) {
        console.warn('辞書の読み込みに失敗:', e);
    }
}

function parseDictionaryText(text) {
    const entries = {};
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
    for (const line of lines) {
        const sep = line.includes(',') ? ',' : (line.includes('=') ? '=' : null);
        if (sep) {
            const idx = line.indexOf(sep);
            const key = line.slice(0, idx).trim();
            const val = line.slice(idx + 1).trim();
            if (key && val) entries[key] = val;
        }
    }
    return entries;
}

async function saveDictionary() {
    if (!elements.dictionaryText) return;
    const text = elements.dictionaryText.value;
    const entries = parseDictionaryText(text);
    try {
        const res = await fetch('/api/dictionary', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ entries })
        });
        if (!res.ok) throw new Error('保存に失敗しました');
        log(`辞書を保存しました（${Object.keys(entries).length}件）`, 'success');
        alert('辞書を保存しました！');
    } catch (e) {
        log(`辞書の保存に失敗: ${e.message}`, 'error');
        alert(`エラー: ${e.message}`);
    }
}

function switchProvider(provider) {
    state.aiProvider = provider;
    
    const showMap = {
        ollama:  { ollama: true,  openai: false, gemini: false },
        openai:  { ollama: false, openai: true,  gemini: false },
        gemini:  { ollama: false, openai: false, gemini: true  },
    };
    const show = showMap[provider] || showMap.ollama;
    
    if (elements.ollamaSettings) elements.ollamaSettings.style.display = show.ollama ? 'block' : 'none';
    if (elements.openaiSettings) elements.openaiSettings.style.display = show.openai ? 'block' : 'none';
    if (elements.geminiSettings) elements.geminiSettings.style.display = show.gemini ? 'block' : 'none';
    
    if (elements.ollamaModelSelect) elements.ollamaModelSelect.style.display = show.ollama ? 'block' : 'none';
    if (elements.gptModelSelect) elements.gptModelSelect.style.display = show.openai ? 'block' : 'none';
    if (elements.geminiModelSelect) elements.geminiModelSelect.style.display = show.gemini ? 'block' : 'none';
    
    const modelLabel = document.getElementById('model-label');
    if (modelLabel) {
        const labels = { ollama: 'Ollamaモデル', openai: 'GPTモデル', gemini: 'Geminiモデル' };
        modelLabel.textContent = labels[provider] || 'AIモデル';
    }
    
    if (provider === 'ollama') {
        refreshOllamaModels();
    }
}

function getActiveApiKey() {
    if (state.aiProvider === 'ollama') return '__ollama__';
    if (state.aiProvider === 'gemini') return state.geminiApiKey;
    return state.apiKey;
}

function getActiveModel() {
    if (state.aiProvider === 'ollama') return state.ollamaModel;
    if (state.aiProvider === 'gemini') return state.geminiModel;
    return state.gptModel;
}

// ========================================
// Ollama モデル管理
// ========================================

async function refreshOllamaModels() {
    const endpoint = elements.ollamaEndpointInput ? elements.ollamaEndpointInput.value : state.ollamaEndpoint;
    const statusIcon = document.getElementById('ollama-status-icon');
    const statusText = document.getElementById('ollama-status-text');
    
    if (statusIcon) statusIcon.textContent = '⏳';
    if (statusText) statusText.textContent = '接続確認中...';
    
    try {
        const resp = await fetch(`/api/ollama/models?endpoint=${encodeURIComponent(endpoint)}`);
        
        if (!resp.ok) {
            if (statusIcon) statusIcon.textContent = '❌';
            if (statusText) statusText.textContent = 'Ollamaに接続できません';
            log('Ollamaに接続できません。ollama serve で起動してください。', 'error');
            return;
        }
        
        const data = await resp.json();
        const select = elements.ollamaModelSelect;
        if (!select) return;
        
        select.innerHTML = '';
        
        if (data.models.length === 0) {
            select.innerHTML = '<option value="">モデルがありません</option>';
            if (statusIcon) statusIcon.textContent = '⚠️';
            if (statusText) statusText.textContent = '接続OK（モデルなし）';
            log('Ollamaにモデルがインストールされていません。ollama pull gemma3 等でインストールしてください。', 'error');
            return;
        }
        
        data.models.forEach((m, i) => {
            const opt = document.createElement('option');
            opt.value = m.name;
            opt.textContent = `${m.name} (${m.size})`;
            if (state.ollamaModel === m.name || (i === 0 && !state.ollamaModel)) {
                opt.selected = true;
                state.ollamaModel = m.name;
            }
            select.appendChild(opt);
        });
        
        if (statusIcon) statusIcon.textContent = '✅';
        if (statusText) statusText.textContent = `接続OK（${data.models.length}モデル）`;
        log(`Ollama接続OK: ${data.models.length}モデル利用可能`, 'success');
        
    } catch (e) {
        if (statusIcon) statusIcon.textContent = '❌';
        if (statusText) statusText.textContent = '接続エラー';
        console.error('Ollama接続エラー:', e);
    }
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
        
        if (!templateList) return;
        
        const icons = ['🏥', '📋', '📝', '📊', '📄', '🗂️', '📑', '🗒️'];
        
        templateList.innerHTML = '';
        
        // 「テンプレートなし」を先頭に追加
        const freeformItem = document.createElement('div');
        freeformItem.className = 'template-item active';
        freeformItem.dataset.template = '__none__';
        freeformItem.innerHTML = `
            <div class="template-icon">✨</div>
            <div class="template-info">
                <h3>テンプレートなし</h3>
                <p>AIが自動で議事録を構成</p>
            </div>
            <span class="check-mark">✓</span>
        `;
        templateList.appendChild(freeformItem);
        state.selectedTemplate = '__none__';
        
        // 既存テンプレートを追加
        if (data.templates && data.templates.length > 0) {
            data.templates.forEach((tpl, index) => {
                const item = document.createElement('div');
                item.className = 'template-item';
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
        }
        
        setupTemplateSelection();
        
        const totalCount = (data.templates ? data.templates.length : 0) + 1;
        log(`テンプレート ${totalCount}件を読み込みました（テンプレートなし含む）`, 'info');
        
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

function formatAudioTime(sec) {
    if (sec == null || !Number.isFinite(sec) || sec < 0) return '—';
    const m = Math.floor(sec / 60);
    const s = Math.floor(sec % 60);
    return `${m}:${s.toString().padStart(2, '0')}`;
}

function resetTranscribeDetailUI() {
    const detail = document.getElementById('transcribe-progress-detail');
    const label = document.getElementById('transcribe-progress-label');
    const fill = document.getElementById('transcribe-audio-fill');
    const cur = document.getElementById('transcribe-time-current');
    const tot = document.getElementById('transcribe-time-total');
    const preview = document.getElementById('transcribe-preview');
    if (detail) detail.classList.add('hidden');
    if (label) label.textContent = '音声を解析しています…';
    if (fill) {
        fill.style.width = '0%';
        fill.classList.remove('indeterminate');
    }
    if (cur) cur.textContent = '0:00';
    if (tot) tot.textContent = '—';
    if (preview) preview.textContent = '';
}

function showTranscribeDetailUI() {
    const detail = document.getElementById('transcribe-progress-detail');
    if (detail) detail.classList.remove('hidden');
}

function applyTranscribeStreamEvent(ev) {
    const label = document.getElementById('transcribe-progress-label');
    const fill = document.getElementById('transcribe-audio-fill');
    const cur = document.getElementById('transcribe-time-current');
    const tot = document.getElementById('transcribe-time-total');
    const preview = document.getElementById('transcribe-preview');

    if (ev.type === 'transcribe_start') {
        showTranscribeDetailUI();
        if (label) label.textContent = ev.message || '文字起こしを開始しました';
        if (tot) tot.textContent = ev.duration > 0 ? formatAudioTime(ev.duration) : '取得中…';
    }
    if (ev.type === 'transcribe_segment') {
        showTranscribeDetailUI();
        if (ev.cli_mode) {
            if (label) label.textContent = 'Whisper CLI で処理中（完了までお待ちください）';
            if (fill) fill.classList.add('indeterminate');
            return;
        }
        if (fill) {
            fill.classList.remove('indeterminate');
            const pct = typeof ev.percent === 'number' ? Math.min(100, ev.percent) : 0;
            fill.style.width = `${pct}%`;
        }
        if (cur) cur.textContent = formatAudioTime(ev.seconds_end);
        if (tot && ev.duration > 0) tot.textContent = formatAudioTime(ev.duration);
        if (label) {
            label.textContent = ev.duration > 0
                ? `約 ${formatAudioTime(ev.seconds_end)} / ${formatAudioTime(ev.duration)} まで文字起こし済み`
                : `セグメント ${ev.segment_index || 0} を処理中`;
        }
        if (preview && ev.preview) {
            preview.textContent = ev.preview;
        }
    }
    if (ev.type === 'transcribe_done') {
        if (fill) {
            fill.classList.remove('indeterminate');
            fill.style.width = '100%';
        }
        if (label) label.textContent = `文字起こし完了（${ev.chars || 0}文字）`;
        if (cur && ev.duration > 0) cur.textContent = formatAudioTime(ev.duration);
        if (tot && ev.duration > 0) tot.textContent = formatAudioTime(ev.duration);
    }
    if (ev.type === 'ai_start') {
        if (label) label.textContent = ev.message || 'AIで要約しています';
    }
    if (ev.type === 'file_write_start') {
        if (label) label.textContent = ev.message || 'ファイルを書き出しています';
    }
}

function showProgress() {
    elements.progressCard.classList.remove('hidden');
    elements.resultCard.classList.add('hidden');
    resetTranscribeDetailUI();
}

function hideProgress() {
    elements.progressCard.classList.add('hidden');
    resetTranscribeDetailUI();
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

/**
 * NDJSON ストリームで処理進捗を受け取り、最終 result を返す
 */
async function consumeProcessStream(formData, onEvent) {
    const response = await fetch('/api/process-stream', {
        method: 'POST',
        body: formData
    });

    if (!response.ok) {
        const t = await response.text();
        let errMsg = 'サーバーエラー';
        try {
            const j = JSON.parse(t);
            if (Array.isArray(j.detail)) {
                errMsg = j.detail.map((x) => x.msg || x).join(', ');
            } else {
                errMsg = j.detail || errMsg;
            }
        } catch (e) {
            if (t) errMsg = t.slice(0, 300);
        }
        throw new Error(errMsg);
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let result = null;

    while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed) continue;
            let ev;
            try {
                ev = JSON.parse(trimmed);
            } catch (e) {
                continue;
            }
            if (ev.type === 'ping') continue;
            if (ev.type === 'done') {
                result = ev.result;
                continue;
            }
            if (ev.type === 'error') {
                throw new Error(ev.detail || '処理エラー');
            }
            if (onEvent) onEvent(ev);
        }
    }

    if (!result) {
        throw new Error('サーバーから完了応答がありませんでした');
    }
    return result;
}

async function runPipeline() {
    if (!state.selectedFile) {
        alert('ファイルを選択してください');
        return;
    }
    
    if (state.aiProvider === 'ollama') {
        if (!state.ollamaModel) {
            alert('Ollamaモデルが選択されていません。Ollamaに接続してモデルを取得してください。');
            return;
        }
    } else if (!getActiveApiKey() || getActiveApiKey() === '__ollama__') {
        const providerName = state.aiProvider === 'gemini' ? 'Gemini' : 'OpenAI';
        alert(`${providerName} API Keyを設定してください`);
        return;
    }
    
    state.isProcessing = true;
    updateRunButton();
    showProgress();
    
    log('='.repeat(40), 'info');
    log('🚀 議事録作成を開始します', 'info');
    
    try {
        // Step 1: 文字起こし
        updateProgress(5, '🎙️ 音声を文字起こし中...', 1);
        log(`🎙️ Whisperで文字起こし開始 (モデル: ${state.whisperModel})`, 'info');

        const formData = new FormData();
        formData.append('file', state.selectedFile);
        formData.append('whisper_model', state.whisperModel);
        formData.append('gpt_model', getActiveModel());
        formData.append('template', state.selectedTemplate);
        formData.append('output_format', elements.outputFormatSelect?.value || state.outputFormat);
        formData.append('ai_provider', state.aiProvider);
        if (state.aiProvider === 'ollama') {
            formData.append('ollama_endpoint', state.ollamaEndpoint);
        } else {
            formData.append('api_key', getActiveApiKey());
        }

        const result = await consumeProcessStream(formData, (ev) => {
            applyTranscribeStreamEvent(ev);
            if (ev.type === 'transcribe_segment' && !ev.cli_mode && typeof ev.percent === 'number') {
                const overall = 5 + (ev.percent / 100) * 42;
                updateProgress(Math.round(overall), '🎙️ 文字起こし中...', 1);
            }
            if (ev.type === 'transcribe_start') {
                updateProgress(8, '🎙️ 文字起こしを準備中...', 1);
            }
            if (ev.type === 'ai_start') {
                updateProgress(55, '🤖 AI要約中...', 2);
            }
            if (ev.type === 'file_write_start') {
                updateProgress(88, '📄 ファイル生成中...', 3);
            }
        });
        
        const providerLabels = { ollama: 'Ollama（ローカル）', gemini: 'Gemini', openai: 'ChatGPT' };
        const providerLabel = providerLabels[state.aiProvider] || 'AI';
        
        // Step 2: AI要約完了
        updateProgress(70, '🤖 AI要約完了', 2);
        log(`✅ ${providerLabel}応答を受信`, 'success');
        
        // Step 3: ファイル生成完了
        const outputFormat = elements.outputFormatSelect?.value || state.outputFormat;
        const formatLabels = { excel: 'Excel', text: 'テキスト', docx: 'Word' };
        const formatLabel = formatLabels[outputFormat] || 'ファイル';
        updateProgress(100, `📄 ${formatLabel}生成完了`, 3);
        log(`✅ ${formatLabel}出力完了: ${result.filename}`, 'success');
        
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
    
    // 辞書保存
    if (elements.saveDictionaryBtn) {
        elements.saveDictionaryBtn.addEventListener('click', saveDictionary);
    }
    
    // プロバイダー切替
    if (elements.aiProviderSelect) {
        elements.aiProviderSelect.addEventListener('change', () => {
            switchProvider(elements.aiProviderSelect.value);
            updateRunButton();
        });
    }
    
    // API Key入力時に状態更新
    if (elements.apiKeyInput) {
        elements.apiKeyInput.addEventListener('input', () => {
            state.apiKey = elements.apiKeyInput.value;
            updateRunButton();
        });
    }
    if (elements.geminiApiKeyInput) {
        elements.geminiApiKeyInput.addEventListener('input', () => {
            state.geminiApiKey = elements.geminiApiKeyInput.value;
            updateRunButton();
        });
    }
    
    // Ollama関連
    if (elements.ollamaRefreshBtn) {
        elements.ollamaRefreshBtn.addEventListener('click', () => {
            state.ollamaEndpoint = elements.ollamaEndpointInput.value;
            refreshOllamaModels();
        });
    }
    if (elements.ollamaModelSelect) {
        elements.ollamaModelSelect.addEventListener('change', () => {
            state.ollamaModel = elements.ollamaModelSelect.value;
        });
    }
    
    // パスワード表示切替
    document.querySelectorAll('.toggle-visibility').forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.dataset.target;
            const input = document.getElementById(targetId);
            if (input) input.type = input.type === 'password' ? 'text' : 'password';
        });
    });
    
    // 実行ボタン
    if (elements.runBtn) {
        elements.runBtn.addEventListener('click', () => {
            if (state.aiProvider === 'ollama' || getActiveApiKey()) {
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
    await loadDictionary();
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
