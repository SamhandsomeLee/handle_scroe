/**
 * 成绩导入转换工具 - 前端 JavaScript
 */

let currentFile = null;
let currentFilePath = null;
let currentOutputFilename = null;

// 页面加载完成
document.addEventListener('DOMContentLoaded', function() {
    setupEventListeners();
});

/**
 * 设置事件监听
 */
function setupEventListeners() {
    const fileInput = document.getElementById('fileInput');
    const uploadBox = document.querySelector('.upload-box');
    
    // 文件选择
    fileInput.addEventListener('change', handleFileSelect);
    
    // 拖放
    uploadBox.addEventListener('dragover', handleDragOver);
    uploadBox.addEventListener('dragleave', handleDragLeave);
    uploadBox.addEventListener('drop', handleDrop);
    
    // 点击上传框
    uploadBox.addEventListener('click', function() {
        fileInput.click();
    });
}

/**
 * 处理文件选择
 */
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        uploadFile(file);
    }
}

/**
 * 处理拖放
 */
function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    document.querySelector('.upload-box').classList.add('dragover');
}

function handleDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();
    document.querySelector('.upload-box').classList.remove('dragover');
}

function handleDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    document.querySelector('.upload-box').classList.remove('dragover');
    
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        uploadFile(files[0]);
    }
}

/**
 * 上传文件
 */
function uploadFile(file) {
    // 验证文件类型
    const validTypes = ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    if (!validTypes.includes(file.type) && !file.name.endsWith('.xls') && !file.name.endsWith('.xlsx')) {
        showError('不支持的文件格式，请上传 .xls 或 .xlsx 文件');
        return;
    }
    
    // 验证文件大小
    const maxSize = 50 * 1024 * 1024; // 50MB
    if (file.size > maxSize) {
        showError('文件过大，请上传小于 50MB 的文件');
        return;
    }
    
    currentFile = file;
    
    // 显示进度
    showProgress('上传中...');
    
    // 创建 FormData
    const formData = new FormData();
    formData.append('file', file);
    
    // 上传
    fetch('/api/upload', {
        method: 'POST',
        body: formData
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            currentFilePath = data.filepath;
            displayFileInfo(data);
            hideProgress();
        } else {
            showError(data.error || '上传失败');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        showError('上传失败: ' + error.message);
    });
}

/**
 * 显示文件信息
 */
function displayFileInfo(data) {
    // 显示文件名
    document.getElementById('fileName').textContent = data.filename;
    document.getElementById('sheetCount').textContent = data.total_sheets;
    
    // 显示工作表信息
    const tableBody = document.getElementById('sheetTable');
    tableBody.innerHTML = '';
    
    data.sheets.forEach(sheet => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${sheet.name}</td>
            <td>${sheet.rows}</td>
            <td>${sheet.cols}</td>
            <td>${sheet.student_count}</td>
        `;
        tableBody.appendChild(row);
    });
    
    // 显示信息和操作按钮
    document.getElementById('infoSection').style.display = 'block';
    document.getElementById('actionSection').style.display = 'block';
}

/**
 * 转换文件
 */
function convertFile() {
    if (!currentFilePath) {
        showError('请先选择文件');
        return;
    }
    
    // 禁用按钮
    document.querySelectorAll('.btn').forEach(btn => btn.disabled = true);
    
    // 显示进度
    showProgress('转换中...');
    updateProgress(30);
    
    // 发送转换请求
    fetch('/api/convert', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            filepath: currentFilePath
        })
    })
    .then(response => response.json())
    .then(data => {
        updateProgress(100);
        
        if (data.success) {
            currentOutputFilename = data.output_filename;
            setTimeout(() => {
                hideProgress();
                showSuccess(data.message);
            }, 500);
        } else {
            showError(data.error || '转换失败');
        }
    })
    .catch(error => {
        console.error('Error:', error);
        showError('转换失败: ' + error.message);
    })
    .finally(() => {
        // 启用按钮
        document.querySelectorAll('.btn').forEach(btn => btn.disabled = false);
    });
}

/**
 * 下载文件
 */
function downloadFile() {
    if (!currentOutputFilename) {
        showError('文件不存在');
        return;
    }
    
    const url = `/api/download/${encodeURIComponent(currentOutputFilename)}`;
    
    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error('下载失败: ' + response.statusText);
            }
            return response.blob();
        })
        .then(blob => {
            // 创建 blob URL
            const blobUrl = window.URL.createObjectURL(blob);
            
            // 创建下载链接
            const a = document.createElement('a');
            a.href = blobUrl;
            a.download = currentOutputFilename;
            document.body.appendChild(a);
            a.click();
            
            // 清理
            window.URL.revokeObjectURL(blobUrl);
            document.body.removeChild(a);
        })
        .catch(error => {
            console.error('Error:', error);
            showError('下载失败: ' + error.message);
        });
}

/**
 * 重置表单
 */
function resetForm() {
    // 隐藏所有部分
    document.getElementById('infoSection').style.display = 'none';
    document.getElementById('actionSection').style.display = 'none';
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('resultSection').style.display = 'none';
    document.getElementById('errorSection').style.display = 'none';
    
    // 重置文件输入
    document.getElementById('fileInput').value = '';
    
    // 重置变量
    currentFile = null;
    currentFilePath = null;
    currentOutputFilename = null;
    
    // 清理临时文件
    fetch('/api/cleanup', {
        method: 'POST'
    })
    .catch(error => console.error('Cleanup error:', error));
}

/**
 * 显示进度
 */
function showProgress(message) {
    document.getElementById('progressText').textContent = message;
    document.getElementById('progressFill').style.width = '0%';
    document.getElementById('progressSection').style.display = 'block';
    document.getElementById('resultSection').style.display = 'none';
    document.getElementById('errorSection').style.display = 'none';
}

/**
 * 更新进度
 */
function updateProgress(percentage) {
    document.getElementById('progressFill').style.width = percentage + '%';
}

/**
 * 隐藏进度
 */
function hideProgress() {
    document.getElementById('progressSection').style.display = 'none';
}

/**
 * 显示成功
 */
function showSuccess(message) {
    document.getElementById('resultTitle').textContent = '转换成功！';
    document.getElementById('resultMessage').textContent = message;
    document.getElementById('resultSection').style.display = 'block';
    document.getElementById('errorSection').style.display = 'none';
}

/**
 * 显示错误
 */
function showError(message) {
    document.getElementById('errorMessage').textContent = message;
    document.getElementById('errorSection').style.display = 'block';
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('resultSection').style.display = 'none';
    
    // 启用按钮
    document.querySelectorAll('.btn').forEach(btn => btn.disabled = false);
}
