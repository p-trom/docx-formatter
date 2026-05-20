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
        submitBtn.disabled = !templateFile || !contentFile || isProcessing;
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
    resetBtn.addEventListener('click', resetForm);

    // Load history on init
    renderHistory();

})();
