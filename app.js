// Store multiple files data
let filesData = new Map(); // Map<fileName, {data, selectedRows}>
let currentActiveFile = null;

// Elements
const fileInput = document.getElementById('fileInput');
const dataTable = document.getElementById('dataTable');
const tableBody = document.getElementById('tableBody');
const loading = document.getElementById('loading');
const exportBtn = document.getElementById('exportBtn');
const fileList = document.getElementById('fileList');
const fileItems = document.getElementById('fileItems');
const fileCount = document.getElementById('fileCount');

// Toast container
const toastContainer = document.getElementById('toastContainer');

// Event Listeners
fileInput.addEventListener('change', handleFileSelect);
exportBtn.addEventListener('click', exportAllFiles);

// Toast notification function
function showToast(message, type = 'info', title = '') {
    // Set default titles
    const defaultTitles = {
        'info': 'Thông báo',
        'success': 'Thành công',
        'error': 'Lỗi',
        'warning': 'Cảnh báo'
    };
    
    const toastTitle = title || defaultTitles[type] || 'Thông báo';
    
    // Set icon based on type
    const icons = {
        'info': 'ℹ️',
        'success': '✅',
        'error': '❌',
        'warning': '⚠️'
    };
    const icon = icons[type] || icons['info'];
    
    // Create toast element
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
        <div class="toast-icon">${icon}</div>
        <div class="toast-content">
            <div class="toast-title">${toastTitle}</div>
            <div class="toast-message">${message}</div>
        </div>
        <button class="toast-close">✕</button>
    `;
    
    // Add to container
    toastContainer.appendChild(toast);
    
    // Close button
    const closeBtn = toast.querySelector('.toast-close');
    closeBtn.addEventListener('click', () => removeToast(toast));
    
    // Click to close
    toast.addEventListener('click', (e) => {
        if (!e.target.classList.contains('toast-close')) {
            removeToast(toast);
        }
    });
    
    // Auto remove after 5 seconds
    setTimeout(() => removeToast(toast), 5000);
}

function removeToast(toast) {
    if (toast && toast.parentElement) {
        toast.style.animation = 'fadeOut 0.3s ease';
        setTimeout(() => {
            if (toast.parentElement) {
                toast.remove();
            }
        }, 300);
    }
}

// Handle file selection (multiple files)
function handleFileSelect(e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;
    
    showLoading();
    processMultipleFiles(files);
}

// Process multiple files
async function processMultipleFiles(files) {
    let successCount = 0;
    let errorFiles = [];
    
    for (const file of files) {
        if (!file.name.match(/\.(xlsx|xls)$/)) {
            errorFiles.push(`${file.name} (không phải file Excel)`);
            continue;
        }
        
        try {
            const data = await readFileData(file);
            filesData.set(file.name, {
                data: data,
                selectedRows: new Set()
            });
            successCount++;
        } catch (error) {
            errorFiles.push(`${file.name} (${error.message})`);
        }
    }
    
    hideLoading();
    updateFileList();
    
    // Display first file
    if (filesData.size > 0 && !currentActiveFile) {
        const firstFileName = Array.from(filesData.keys())[0];
        switchToFile(firstFileName);
    }
    
    // Show result
    if (successCount > 0) {
        showToast(`Đã tải lên thành công ${successCount} file!`, 'success');
    }
    
    if (errorFiles.length > 0) {
        showToast(`Không thể tải lên:\n${errorFiles.join('\n')}`, 'error');
    }
    
    // Reset file input
    fileInput.value = '';
}

// Read file data
function readFileData(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = () => reject(new Error('Lỗi khi đọc file'));
        reader.readAsArrayBuffer(file);
    });
}

// Update file list display
function updateFileList() {
    if (filesData.size === 0) {
        fileList.style.display = 'none';
        return;
    }
    
    fileList.style.display = 'block';
    fileCount.textContent = filesData.size;
    
    let html = '';
    filesData.forEach((fileData, fileName) => {
        const isActive = fileName === currentActiveFile;
        html += `
            <div class="file-item ${isActive ? 'active' : ''}" onclick="switchToFile('${escapeHtml(fileName)}')">
                <span class="file-item-name" title="${escapeHtml(fileName)}">📄 ${escapeHtml(fileName)}</span>
                <button class="file-item-remove" onclick="event.stopPropagation(); removeFile('${escapeHtml(fileName)}')" title="Xóa file">✕</button>
            </div>
        `;
    });
    
    fileItems.innerHTML = html;
}

// Switch to a different file
function switchToFile(fileName) {
    const fileData = filesData.get(fileName);
    if (!fileData) return;
    
    // Save current selections
    if (currentActiveFile) {
        saveCurrentSelections();
    }
    
    currentActiveFile = fileName;
    displayData(fileData.data);
    updateFileList();
    
    // Restore selections
    restoreSelections(fileData.selectedRows);
}

// Remove a file
function removeFile(fileName) {
    filesData.delete(fileName);
    
    if (currentActiveFile === fileName) {
        currentActiveFile = null;
        if (filesData.size > 0) {
            const firstFileName = Array.from(filesData.keys())[0];
            switchToFile(firstFileName);
        } else {
            tableBody.innerHTML = '<tr class="empty-row"><td colspan="8" class="empty-message">📂 Bạn cần tải file lên để xem dữ liệu</td></tr>';
        }
    }
    
    updateFileList();
}

// Save current selections
function saveCurrentSelections() {
    if (!currentActiveFile) return;
    
    const fileData = filesData.get(currentActiveFile);
    if (!fileData) return;
    
    const checkedBoxes = document.querySelectorAll('.row-checkbox:checked');
    fileData.selectedRows = new Set(
        Array.from(checkedBoxes).map(cb => parseInt(cb.dataset.row))
    );
}

// Restore selections
function restoreSelections(selectedRows) {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    checkboxes.forEach(checkbox => {
        const rowIndex = parseInt(checkbox.dataset.row);
        checkbox.checked = selectedRows.has(rowIndex);
    });
    
    // Update select all checkbox
    const selectAll = document.getElementById('selectAll');
    if (selectAll) {
        const allChecked = Array.from(checkboxes).every(cb => cb.checked);
        selectAll.checked = allChecked && checkboxes.length > 0;
    }
}

// Display data in table
function displayData(data) {
    if (!data || data.length === 0) {
        showToast('File không có dữ liệu', 'warning');
        return;
    }

    // Columns to display
    const columnsToShow = [
        'Mã sản phẩm',
        'Sản phẩm',
        'Lượt truy cập sản phẩm',
        'Sản phẩm (Thêm vào giỏ hàng)',
        'Sản phẩm (Đơn đã xác nhận)',
        'Doanh số (Đơn đã xác nhận) (VND)',
        'Tỷ lệ chuyển đổi (Đơn đã xác nhận)'
    ];

    const headers = data[0];
    
    // Find column indices
    const columnIndices = columnsToShow.map(col => {
        const index = headers.findIndex(h => h && h.toString().trim() === col);
        return { name: col, index: index };
    }).filter(col => col.index !== -1);

    if (columnIndices.length === 0) {
        showToast('Không tìm thấy các cột cần hiển thị trong file', 'warning');
        return;
    }

    let html = '';
    
    // Data rows - only selected columns
    for (let i = 1; i < data.length; i++) {
        html += '<tr>';
        
        // Checkbox column
        html += `<td class="checkbox-col"><input type="checkbox" class="row-checkbox" data-row="${i}"></td>`;
        
        const row = data[i];
        
        columnIndices.forEach(col => {
            const cell = row[col.index] !== undefined ? row[col.index] : '';
            html += `<td>${escapeHtml(String(cell))}</td>`;
        });
        html += '</tr>';
    }
    
    tableBody.innerHTML = html;
    
    // Add checkbox to header
    const headerCheckbox = document.querySelector('thead th.checkbox-col');
    if (headerCheckbox && !document.getElementById('selectAll')) {
        headerCheckbox.innerHTML = '<input type="checkbox" id="selectAll">';
    }
    
    // Add event listeners for checkboxes
    setupCheckboxListeners();
}

// Setup checkbox event listeners
function setupCheckboxListeners() {
    const selectAll = document.getElementById('selectAll');
    const rowCheckboxes = document.querySelectorAll('.row-checkbox');
    
    if (!selectAll) return;
    
    // Remove old listeners by cloning
    const newSelectAll = selectAll.cloneNode(true);
    selectAll.parentNode.replaceChild(newSelectAll, selectAll);
    
    // Select all checkbox
    newSelectAll.addEventListener('change', function() {
        rowCheckboxes.forEach(checkbox => {
            checkbox.checked = this.checked;
        });
    });
    
    // Individual checkboxes
    rowCheckboxes.forEach(checkbox => {
        checkbox.addEventListener('change', function() {
            const allChecked = Array.from(rowCheckboxes).every(cb => cb.checked);
            newSelectAll.checked = allChecked;
        });
    });
}

// Export all files with selected products
async function exportAllFiles() {
    if (filesData.size === 0) {
        showToast('Vui lòng tải file lên trước!', 'warning');
        return;
    }
    
    // Save current selections before exporting
    saveCurrentSelections();
    
    let totalExported = 0;
    let filesWithSelections = 0;
    let errorFiles = [];
    
    showLoading();
    
    // Debug: Log what we're about to export
    console.log('Exporting files:', Array.from(filesData.entries()).map(([name, data]) => ({
        name,
        selectedCount: data.selectedRows.size,
        totalRows: data.data.length - 1
    })));
    
    for (const [fileName, fileData] of filesData) {
        if (fileData.selectedRows.size === 0) {
            console.log(`Skipping ${fileName} - no selections`);
            continue;
        }
        
        filesWithSelections++;
        
        try {
            console.log(`Exporting ${fileName} with ${fileData.selectedRows.size} rows`);
            await exportSingleFile(fileName, fileData.data, fileData.selectedRows);
            totalExported += fileData.selectedRows.size;
            
            // Small delay between downloads
            await new Promise(resolve => setTimeout(resolve, 300));
        } catch (error) {
            console.error(`Error exporting ${fileName}:`, error);
            errorFiles.push(`${fileName}: ${error.message}`);
        }
    }
    
    hideLoading();
    
    if (filesWithSelections === 0) {
        showToast('Vui lòng chọn ít nhất một sản phẩm từ bất kỳ file nào để xuất!', 'warning');
    } else {
        if (errorFiles.length > 0) {
            showToast(`Đã xuất ${filesWithSelections} file với tổng ${totalExported} sản phẩm.\n\nLỗi:\n${errorFiles.join('\n')}`, 'warning');
        } else {
            showToast(`Đã xuất ${filesWithSelections} file với tổng ${totalExported} sản phẩm!`, 'success');
        }
    }
}

// Export single file
async function exportSingleFile(fileName, data, selectedRows) {
    // Columns to export
    const columnsToShow = [
        'Mã sản phẩm',
        'Sản phẩm',
        'Lượt truy cập sản phẩm',
        'Sản phẩm (Thêm vào giỏ hàng)',
        'Sản phẩm (Đơn đã xác nhận)',
        'Doanh số (Đơn đã xác nhận) (VND)',
        'Tỷ lệ chuyển đổi (Đơn đã xác nhận)'
    ];
    
    const headers = data[0];
    
    // Find column indices
    const columnIndices = columnsToShow.map(col => {
        const index = headers.findIndex(h => h && h.toString().trim() === col);
        return index;
    }).filter(index => index !== -1);
    
    // Create workbook with ExcelJS
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sản phẩm đã chọn');
    
    // Add header row
    const headerRow = columnIndices.map(index => headers[index]);
    worksheet.addRow(headerRow);
    
    // Style header row - Gray background with wrap text
    const headerRowObj = worksheet.getRow(1);
    headerRowObj.height = 40;
    headerRowObj.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFA6A6A6' }
        };
        cell.font = {
            name: 'Arial',
            bold: true,
            color: { argb: 'FF000000' },
            size: 10
        };
        cell.alignment = {
            vertical: 'middle',
            horizontal: 'center',
            wrapText: true
        };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FF000000' } },
            left: { style: 'thin', color: { argb: 'FF000000' } },
            bottom: { style: 'thin', color: { argb: 'FF000000' } },
            right: { style: 'thin', color: { argb: 'FF000000' } }
        };
    });
    
    // Add data rows
    Array.from(selectedRows).forEach(rowIndex => {
        const row = data[rowIndex];
        const dataRow = columnIndices.map(colIndex => {
            return row[colIndex] !== undefined ? row[colIndex] : '';
        });
        const addedRow = worksheet.addRow(dataRow);
        
        // Apply Arial font size 10 to data rows
        addedRow.eachCell((cell) => {
            cell.font = {
                name: 'Arial',
                size: 10
            };
        });
    });
    
    // Set column widths
    worksheet.columns.forEach((column) => {
        column.width = 25;
    });
    
    // Generate Excel file and download
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
}

// UI helpers
function showLoading() {
    loading.style.display = 'block';
}

function hideLoading() {
    loading.style.display = 'none';
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
