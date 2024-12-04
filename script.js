let defaultFiles = ['ActivityRewardConfig.xlsx', 'RewardConfig.xlsx', 'RewardConfigMall.xlsx', 'RewardConfigUniqueDraw.xlsx'];
let searchResults = [];
let currentFolderPath = '';
let dirHandle = null;

// 初始化页面
window.onload = function() {
    loadConfig();
    // initializeFileList();
    // 设置默认文件名列表
    document.getElementById('fileNames').value = defaultFiles.join(',');
};

// 加载配置
function loadConfig() {
    const savedFiles = localStorage.getItem('excelFiles');
    if (savedFiles) {
        defaultFiles = JSON.parse(savedFiles);
        document.getElementById('fileNames').value = defaultFiles.join(',');
    } else {
        document.getElementById('fileNames').value = defaultFiles.join(',');
    }
    
    const savedPath = localStorage.getItem('folderPath');
    if (savedPath) {
        currentFolderPath = savedPath;
        document.getElementById('folderPath').value = savedPath;
    }
    
    const savedColumns = localStorage.getItem('columnNames');
    if (savedColumns) {
        const columns = JSON.parse(savedColumns);
        document.getElementById('columnNames').value = `${columns.id},${columns.randomType},${columns.fixed}`;
        updateSearchColumnOptions(`${columns.id},${columns.randomType},${columns.fixed}`);
    } else {
        document.getElementById('columnNames').value = 'Id,randomtype,fixed';
        updateSearchColumnOptions('Id,randomtype,fixed');
    }
    
    const savedSearchColumn = localStorage.getItem('searchColumn');
    if (savedSearchColumn) {
        document.getElementById('searchColumn').value = savedSearchColumn;
    }
}

// 添加新函数：更新搜索列下拉菜单
function updateSearchColumnOptions(columnNamesStr) {
    const columns = columnNamesStr.split(',').map(c => c.trim());
    const searchColumnSelect = document.getElementById('searchColumn');
    searchColumnSelect.innerHTML = ''; // 清空现有选项
    
    columns.forEach(column => {
        const option = document.createElement('option');
        option.value = column.toLowerCase();
        option.textContent = column;
        searchColumnSelect.appendChild(option);
    });
}

// 修改选择文件夹函数
async function selectFolder() {
    try {
        dirHandle = await window.showDirectoryPicker();
        currentFolderPath = dirHandle.name;
        document.getElementById('folderPath').value = dirHandle.name;
        localStorage.setItem('folderPath', dirHandle.name);
    } catch (err) {
        console.error('选择文件夹失败:', err);
    }
}

// 验证文件是否存在
async function validateFiles() {
    const fileNames = document.getElementById('fileNames').value.split(',').map(f => f.trim()).filter(f => f);
    const validationDiv = document.getElementById('fileValidation');
    validationDiv.innerHTML = '';
    
    if (!dirHandle) return;

    const nonExistentFiles = [];
    for (let fileName of fileNames) {
        try {
            await dirHandle.getFileHandle(fileName);
        } catch {
            nonExistentFiles.push(fileName);
        }
    }
    
    if (nonExistentFiles.length > 0) {
        validationDiv.innerHTML = `警告：以下文件不存在：${nonExistentFiles.join(', ')}`;
    }
    
    localStorage.setItem('excelFiles', JSON.stringify(fileNames));
    defaultFiles = fileNames;
}

// 开始搜索
async function startSearch() {
    if (!dirHandle) {
        alert('请先选择文件夹路径');
        return;
    }
    
    const target = document.getElementById('searchTarget').value;
    if (!target) {
        alert('请输入搜索内容');
        return;
    }

    const fileNames = document.getElementById('fileNames').value.split(',').map(f => f.trim()).filter(f => f);
    if (fileNames.length === 0) {
        alert('请输入至少一个Excel文件名');
        return;
    }

    // 解析列名
    const columnNames = document.getElementById('columnNames').value.split(',').map(c => c.trim());
    if (columnNames.length !== 3) {
        alert('请输入正确的列名（需要3个列名，用逗号分隔）');
        return;
    }

    const [idCol, randomTypeCol, fixedCol] = columnNames;
    localStorage.setItem('columnNames', JSON.stringify({
        id: idCol,
        randomType: randomTypeCol,
        fixed: fixedCol
    }));

    const searchColumn = document.getElementById('searchColumn').value;
    localStorage.setItem('searchColumn', searchColumn);

    // 修复：使用正确的方式获取选中的列名
    const searchColumnSelect = document.getElementById('searchColumn');
    const searchColumnName = columnNames[searchColumnSelect.selectedIndex];

    // 直接使用选择的列名进行搜索，不需要switch判断
    // const searchColumnName = columnNames[searchColumnSelect.selectedIndex];

    // 根据选择的列进行搜索
    // let searchColumnName;
    // switch(searchColumn) {
    //     case 'id':
    //         searchColumnName = idCol;
    //         break;
    //     case 'randomType':
    //         searchColumnName = randomTypeCol;
    //         break;
    //     case 'fixed':
    //         searchColumnName = fixedCol;
    //         break;
    // }

    searchResults = [];
    // showProgress();
    
    let processedFiles = 0;
    const totalFiles = fileNames.length;

    for (let fileName of fileNames) {
        try {
            const fileHandle = await dirHandle.getFileHandle(fileName);
            const file = await fileHandle.getFile();
            await processExcel(file, target, idCol, randomTypeCol, fixedCol, searchColumnName);
        } catch (err) {
            console.error(`处理文件 ${fileName} 时出错:`, err);
        }
        processedFiles++;
        // updateProgress((processedFiles / totalFiles) * 100);
    }

    displayResults();
    document.querySelector('.btn-success').style.display = 'block';
}

// 处理Excel文件
async function processExcel(file, target, idCol, randomTypeCol, fixedCol, searchColumnName) {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);

    for (let sheetName of workbook.SheetNames) {
        if (!sheetName.startsWith('$')) {
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet, {range: 4});

            jsonData.forEach(row => {
                const searchValue = String(row[searchColumnName] || '');
                if (searchValue.includes(target)) {
                    searchResults.push({
                        'Excel文件': file.name,
                        '工作表': sheetName,
                        'ID': row[idCol],
                        '随机类型': row[randomTypeCol],
                        '固定值': row[fixedCol]
                    });
                }
            });
        }
    }
}

// 显示进度条
// function showProgress() {
//     const progressBar = document.querySelector('.progress');
//     progressBar.style.display = 'block';
//     updateProgress(0); // 初始化进度为0
// }

// // 更新进度
// function updateProgress(percentage) {
//     const progressBar = document.querySelector('.progress-bar');
//     progressBar.style.width = `${percentage}%`;
//     progressBar.setAttribute('aria-valuenow', percentage);
// }

// 显示结果
function displayResults() {
    const resultsDiv = document.getElementById('searchResults');
    if (searchResults.length === 0) {
        resultsDiv.innerHTML = '<p>未找到匹配结果</p>';
        return;
    }

    let html = '<table class="table"><thead><tr>';
    const headers = Object.keys(searchResults[0]);
    headers.forEach(header => {
        html += `<th>${header}</th>`;
    });
    html += '</tr></thead><tbody>';

    searchResults.forEach(result => {
        html += '<tr>';
        headers.forEach(header => {
            html += `<td>${result[header]}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';
    resultsDiv.innerHTML = html;
}

// 下载结果
function downloadResults() {
    if (searchResults.length === 0) return;

    const target = document.getElementById('searchTarget').value;
    const now = new Date();
    const timestamp = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}_${now.getHours().toString().padStart(2,'0')}${now.getMinutes().toString().padStart(2,'0')}${now.getSeconds().toString().padStart(2,'0')}`;
    const filename = `搜索结果_${timestamp}_${target}.xlsx`;

    const ws = XLSX.utils.json_to_sheet(searchResults);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "搜索结果");
    XLSX.writeFile(wb, filename);
}