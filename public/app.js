// DOM Elements
const tabs = document.querySelectorAll('.tab');
const panels = document.querySelectorAll('.converter-panel');

// DOCX to MD elements
const docxDropzone = document.getElementById('docx-dropzone');
const docxInput = document.getElementById('docx-input');
const docxFileInfo = document.getElementById('docx-file-info');
const convertDocxBtn = document.getElementById('convert-docx-btn');
const docxResult = document.getElementById('docx-result');
const mdOutput = document.getElementById('md-output');
const copyMdBtn = document.getElementById('copy-md-btn');
const downloadMdBtn = document.getElementById('download-md-btn');

// MD to DOCX elements
const mdInput = document.getElementById('md-input');
const mdInputArea = document.getElementById('md-input-area');
const convertMdBtn = document.getElementById('convert-md-btn');

// Loading & Toast
const loadingOverlay = document.getElementById('loading');
const toast = document.getElementById('toast');

// State
let selectedDocxFile = null;

// Tab switching
tabs.forEach(tab => {
    tab.addEventListener('click', () => {
        tabs.forEach(t => t.classList.remove('active'));
        panels.forEach(p => p.classList.remove('active'));

        tab.classList.add('active');
        document.getElementById(tab.dataset.tab).classList.add('active');
    });
});

// Drag and drop for DOCX
docxDropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    docxDropzone.classList.add('dragover');
});

docxDropzone.addEventListener('dragleave', () => {
    docxDropzone.classList.remove('dragover');
});

docxDropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    docxDropzone.classList.remove('dragover');

    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.doc') || file.name.endsWith('.docx'))) {
        handleDocxFile(file);
    } else {
        showToast('请上传 .doc 或 .docx 文件', 'error');
    }
});

docxDropzone.addEventListener('click', () => {
    docxInput.click();
});

docxInput.addEventListener('change', (e) => {
    if (e.target.files[0]) {
        handleDocxFile(e.target.files[0]);
    }
});

function handleDocxFile(file) {
    selectedDocxFile = file;
    docxFileInfo.textContent = `已选择: ${file.name} (${formatFileSize(file.size)})`;
    docxFileInfo.classList.add('show');
    convertDocxBtn.disabled = false;
}

// Convert DOCX to Markdown
convertDocxBtn.addEventListener('click', async () => {
    if (!selectedDocxFile) return;

    showLoading(true);

    try {
        const formData = new FormData();
        formData.append('file', selectedDocxFile);

        const response = await fetch('/api/docx-to-md', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            mdOutput.value = result.markdown;
            docxResult.classList.add('show');
            showToast('转换成功！', 'success');
        } else {
            showToast(result.error || '转换失败', 'error');
        }
    } catch (error) {
        showToast('转换失败: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
});

// Copy Markdown
copyMdBtn.addEventListener('click', async () => {
    try {
        await navigator.clipboard.writeText(mdOutput.value);
        showToast('已复制到剪贴板！', 'success');
    } catch (error) {
        // Fallback for older browsers
        mdOutput.select();
        document.execCommand('copy');
        showToast('已复制到剪贴板！', 'success');
    }
});

// Download Markdown
downloadMdBtn.addEventListener('click', () => {
    const blob = new Blob([mdOutput.value], { type: 'text/markdown' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = selectedDocxFile ? selectedDocxFile.name.replace(/\.docx?$/, '.md') : 'converted.md';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('下载已开始！', 'success');
});

// MD file input
mdInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            mdInputArea.value = e.target.result;
            showToast('文件已加载！', 'success');
        };
        reader.readAsText(file);
    }
});

// Convert Markdown to DOCX
convertMdBtn.addEventListener('click', async () => {
    const markdown = mdInputArea.value.trim();

    if (!markdown) {
        showToast('请输入 Markdown 内容', 'error');
        return;
    }

    showLoading(true);

    try {
        const response = await fetch('/api/md-to-docx', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ markdown })
        });

        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || '转换失败');
        }

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'converted.docx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        showToast('转换成功，文件已下载！', 'success');
    } catch (error) {
        showToast('转换失败: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
});

// Utility functions
function showLoading(show) {
    loadingOverlay.classList.toggle('show', show);
}

function showToast(message, type = '') {
    toast.textContent = message;
    toast.className = 'toast';
    if (type) toast.classList.add(type);
    toast.classList.add('show');

    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}
