// DOCX Formatter Web App
// Handles drag & drop, file upload, API communication, and history

(function() {
    'use strict';

    // DOM Elements
    const templateDropZone = document.getElementById('templateDropZone');
    const templateInput = document.getElementById('templateInput');
    const templatePreview = document.getElementById('templatePreview');
    const templateFileName = document.getElementById('templateFileName');
    const templateFileSize = document.getElementById('templateFileSize');
    const removeTemplate = document.getElementById('removeTemplate');

    const contentDropZone = document.getElementById('contentDropZone');
    const contentInput = document.getElementById('contentInput');
    const contentPreview = document.getElementById('contentPreview');
    const contentFileName = document.getElementById('contentFileName');
    const contentFileSize = document.getElementById('contentFileSize');
    const removeContent = document.getElementById('removeContent');

    const outputName = document.getElementById('outputName');
    const apiUrl = document.getElementById('apiUrl');
    const submitBtn = document.getElementById('submitBtn');
    const debugBtn = document.getElementById('debugBtn');
    const progressSection = document.getElementById('progressSection');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const errorSection = document.getElementById('errorSection');
    const errorMessage = document.getElementById('errorMessage');
    const successSection = document.getElementById('successSection');
    const successDetails = document.getElementById('successDetails');
    const downloadLink = document.getElementById('downloadLink');
    const resetBtn = document.getElementById('resetBtn');
    const historyList = document.getElementById('historyList');

    // Debug elements
    const debugSection = document.getElementById('debugSection');
    const debugSummary = document.getElementById('debugSummary');
    const debugMatches = document.getElementById('debugMatches');
    const debugUnmatched = document.getElementById('debugUnmatched');
    const closeDebug = document.getElementById('closeDebug');

    // State
    let templateFile = null;
    let contentFile = null;
    let isProcessing = false;

    // ============================
    // Helpers
    // ============================

    function formatBytes(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    function formatDate(date) {
        return new Intl.DateTimeFormat('en-US', {
            month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit'
        }).format(date);
    }

    function show(el) { el.classList.remove('hidden'); }
    function hide(el) { el.classList.add('hidden'); }

    function setProgress(percent, text) {
        progressFill.style.width = percent + '%';
        if (text) progressText.textContent = text;
    }

    function updateSubmitButton() {
        const enabled = templateFile && contentFile && !isProcessing;
        submitBtn.disabled = !enabled;
        debugBtn.disabled = !enabled;
    }

    // ============================
    // History Management
    // ============================

    const HISTORY_KEY = 'docx_formatter_history';
    const MAX_HISTORY = 10;

    function loadHistory() {
        try {
            return JSON.parse(localStorage.getItem(HISTORY_KEY)) || [];
        } catch {
            return [];
        }
    }

    function saveHistory(history) {
        localStorage.setItem(HISTORY_KEY, JSON.stringify(history.slice(0, MAX_HISTORY)));
    }

    function addHistoryItem(item) {
        const history = [item, ...loadHistory()].slice(0, MAX_HISTORY);
        saveHistory(history);
        renderHistory();
    }

    function renderHistory() {
        const history = loadHistory();
        if (history.length === 0) {
            historyList.innerHTML = '<p class="empty-history">No recent formats yet</p>';
            return;
        }
        historyList.innerHTML = history.map(item => `
            <div class="history-item">
                <div class="history-icon">📄</div>
                <div class="history-info">
                    <p class="history-name">${escapeHtml(item.name || 'formatted.docx')}</p>
                    <p class="history-meta">${formatDate(new Date(item.date))}${item.size ? ' · ' + formatBytes(item.size) : ''}</p>
                </div>
                <span class="history-status ${item.status}">${item.status === 'success' ? 'Done' : 'Failed'}</span>
            </div>
        `).join('');
    }

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    // ============================
    // File Handling
    // ============================

    function setupDropZone(dropZone, input, onFile) {
        dropZone.addEventListener('click', () => input.click());
        input.addEventListener('change', (e) => {
            if (e.target.files.length) onFile(e.target.files[0]);
        });

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
            }, false);
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
        });

        dropZone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files.length) onFile(files[0]);
        }, false);
    }

    function setTemplateFile(file) {
        if (!file.name.toLowerCase().endsWith('.docx')) {
            showError('Template must be a .docx file');
            return;
        }
        templateFile = file;
        templateFileName.textContent = file.name;
        templateFileSize.textContent = formatBytes(file.size);
        templateDropZone.classList.add('has-file');
        show(templatePreview);
        hide(errorSection);
        updateSubmitButton();
    }

    function setContentFile(file) {
        if (!file.name.toLowerCase().endsWith('.docx')) {
            showError('Content must be a .docx file');
            return;
        }
        contentFile = file;
        contentFileName.textContent = file.name;
        contentFileSize.textContent = formatBytes(file.size);
        contentDropZone.classList.add('has-file');
        show(contentPreview);
        hide(errorSection);
        updateSubmitButton();
    }

    function clearTemplate() {
        templateFile = null;
        templateInput.value = '';
        templateDropZone.classList.remove('has-file');
        hide(templatePreview);
        updateSubmitButton();
    }

    function clearContent() {
        contentFile = null;
        contentInput.value = '';
        contentDropZone.classList.remove('has-file');
        hide(contentPreview);
        updateSubmitButton();
    }

    // ============================
    // API Communication
    // ============================

    function showError(message) {
        errorMessage.textContent = message;
        show(errorSection);
        hide(progressSection);
        hide(successSection);
    }

    async function handleSubmit() {
        if (!templateFile || !contentFile || isProcessing) return;

        isProcessing = true;
        updateSubmitButton();
        hide(errorSection);
        hide(successSection);
        show(progressSection);
        setProgress(10, 'Uploading template...');

        const formData = new FormData();
        formData.append('template', templateFile);
        formData.append('content', contentFile);

        const filename = outputName.value.trim();
        if (filename) formData.append('output_filename', filename);

        setProgress(30, 'Uploading content...');

        try {
            const baseUrl = apiUrl.value.replace(/\/$/, '');
            setProgress(50, 'Formatting document...');

            const startTime = Date.now();
            const response = await fetch(`${baseUrl}/api/v1/format/template-upload`, {
                method: 'POST',
                body: formData,
            });

            const elapsed = Date.now() - startTime;
            setProgress(90, 'Processing response...');

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(errorData.detail?.message || errorData.detail || `Server error: ${response.status}`);
            }

            const blob = await response.blob();
            const downloadUrl = URL.createObjectURL(blob);
            const outputFilename = filename || 'formatted.docx';

            setProgress(100, 'Complete!');

            // Show success
            downloadLink.href = downloadUrl;
            downloadLink.download = outputFilename;
            successDetails.textContent = `${outputFilename} · ${formatBytes(blob.size)} · ${Math.round(elapsed / 100) / 10}s`;
            show(successSection);
            hide(progressSection);

            // Save to history
            addHistoryItem({
                name: outputFilename,
                size: blob.size,
                date: new Date().toISOString(),
                status: 'success',
            });

        } catch (err) {
            showError(err.message || 'Network error. Check API URL and try again.');
            addHistoryItem({
                name: filename || 'formatted.docx',
                date: new Date().toISOString(),
                status: 'error',
            });
        } finally {
            isProcessing = false;
            updateSubmitButton();
        }
    }

    function resetForm() {
        clearTemplate();
        clearContent();
        outputName.value = 'formatted.docx';
        hide(progressSection);
        hide(errorSection);
        hide(successSection);
        hide(debugSection);
    }

    // ============================
    // Debug
    // ============================

    function getConfidenceClass(confidence) {
        if (confidence >= 0.8) return 'high';
        if (confidence >= 0.5) return 'medium';
        return 'low';
    }

    function renderDebugSummary(data) {
        const llmStatus = data.llm_available
            ? (data.llm_used ? '✅ Used' : '⏭️ Skipped')
            : '❌ Not configured';

        debugSummary.innerHTML = `
            <div class="debug-stat">
                <div class="debug-stat-value">${data.template_styles_count}</div>
                <div class="debug-stat-label">Template Styles</div>
            </div>
            <div class="debug-stat">
                <div class="debug-stat-value">${data.content_paragraphs_count}</div>
                <div class="debug-stat-label">Content Paragraphs</div>
            </div>
            <div class="debug-stat">
                <div class="debug-stat-value">${data.matches.length}</div>
                <div class="debug-stat-label">Matches</div>
            </div>
            <div class="debug-stat">
                <div class="debug-stat-value">${data.unmatched_styles.length}</div>
                <div class="debug-stat-label">Unmatched</div>
            </div>
            <div class="debug-stat">
                <div class="debug-stat-value" style="font-size:14px">${llmStatus}</div>
                <div class="debug-stat-label">LLM</div>
            </div>
        `;
    }

    function renderDebugMatches(matches) {
        if (!matches.length) {
            debugMatches.innerHTML = '<p style="padding:20px;text-align:center;color:var(--gray-400)">No matches found</p>';
            return;
        }

        debugMatches.innerHTML = matches.map(m => `
            <div class="debug-match-item">
                <span class="debug-match-pass ${m.pass_name}">${m.pass_name}</span>
                <span class="debug-match-style" title="${escapeHtml(m.source_style)}">${escapeHtml(m.source_style)}</span>
                <span class="debug-match-arrow">→</span>
                <span class="debug-match-style" title="${escapeHtml(m.target_style)}">${escapeHtml(m.target_style)}</span>
                <span class="debug-match-confidence ${getConfidenceClass(m.confidence)}">${Math.round(m.confidence * 100)}%</span>
                <span class="debug-match-preview">${escapeHtml(m.preview || '')}</span>
            </div>
        `).join('');
    }

    function renderDebugUnmatched(unmatched) {
        if (!unmatched.length) {
            hide(debugUnmatched);
            return;
        }
        show(debugUnmatched);
        debugUnmatched.innerHTML = `
            <h4>⚠️ Unmatched Styles (${unmatched.length})</h4>
            <div class="debug-unmatched-list">
                ${unmatched.map(s => `<span class="debug-unmatched-tag">${escapeHtml(s)}</span>`).join('')}
            </div>
        `;
    }

    async function handleDebug() {
        if (!templateFile || !contentFile || isProcessing) return;

        isProcessing = true;
        updateSubmitButton();
        hide(errorSection);
        hide(successSection);
        hide(debugSection);
        show(progressSection);
        setProgress(30, 'Analyzing styles...');

        const formData = new FormData();
        formData.append('template', templateFile);
        formData.append('content', contentFile);

        try {
            const baseUrl = apiUrl.value.replace(/\/$/, '');
            setProgress(60, 'Running debug pipeline...');

            const response = await fetch(`${baseUrl}/api/v1/format/template-upload/debug`, {
                method: 'POST',
                body: formData,
            });

            setProgress(90, 'Rendering results...');

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                throw new Error(errorData.detail?.message || errorData.detail || `Server error: ${response.status}`);
            }

            const data = await response.json();

            renderDebugSummary(data);
            renderDebugMatches(data.matches);
            renderDebugUnmatched(data.unmatched_styles);

            show(debugSection);
            hide(progressSection);
            setProgress(0, '');

        } catch (err) {
            showError(err.message || 'Network error. Check API URL and try again.');
        } finally {
            isProcessing = false;
            updateSubmitButton();
        }
    }

    // ============================
    // Event Listeners
    // ============================

    setupDropZone(templateDropZone, templateInput, setTemplateFile);
    setupDropZone(contentDropZone, contentInput, setContentFile);

    removeTemplate.addEventListener('click', (e) => {
        e.stopPropagation();
        clearTemplate();
    });

    removeContent.addEventListener('click', (e) => {
        e.stopPropagation();
        clearContent();
    });

    submitBtn.addEventListener('click', handleSubmit);
    debugBtn.addEventListener('click', handleDebug);
    resetBtn.addEventListener('click', resetForm);
    closeDebug.addEventListener('click', () => hide(debugSection));

    // Load history on init
    renderHistory();

})();
