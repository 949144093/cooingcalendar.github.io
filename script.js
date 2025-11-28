// 数据存储
let nameList = [];
let drawnList = [];
let currentEditIndex = -1;

// DOM元素
const excelFileInput = document.getElementById('excelFile');
const fileInfo = document.getElementById('fileInfo');
const downloadTemplateBtn = document.getElementById('downloadTemplateBtn');
const showAllCheckbox = document.getElementById('showAll');
const tableBody = document.getElementById('tableBody');
const totalCountSpan = document.getElementById('totalCount');
const eligibleCountSpan = document.getElementById('eligibleCount');
const drawnCountSpan = document.getElementById('drawnCount');
const drawCountInput = document.getElementById('drawCount');
const drawBtn = document.getElementById('drawBtn');
const resetBtn = document.getElementById('resetBtn');
const wheelGraphic = document.getElementById('wheelGraphic');
const wheelCanvas = document.getElementById('wheelCanvas');
const wheelCtx = wheelCanvas.getContext('2d');
const wheelText = document.getElementById('wheelText');
const wheelPointer = document.getElementById('wheelPointer');
const resultsList = document.getElementById('resultsList');
const clearResultsBtn = document.getElementById('clearResultsBtn');
const editModal = document.getElementById('editModal');
const closeModal = document.getElementById('closeModal');
const editNameInput = document.getElementById('editName');
const editAttendanceInput = document.getElementById('editAttendance');
const saveEditBtn = document.getElementById('saveEditBtn');
const cancelEditBtn = document.getElementById('cancelEditBtn');

// 下载模板文件
downloadTemplateBtn.addEventListener('click', downloadTemplate);

function downloadTemplate() {
    // 创建示例数据
    const templateData = [
        { '姓名': '张三', '出勤次数': 8 },
        { '姓名': '李四', '出勤次数': 7 },
        { '姓名': '王五', '出勤次数': 6 },
        { '姓名': '赵六', '出勤次数': 8 },
        { '姓名': '钱七', '出勤次数': 7 },
        { '姓名': '孙八', '出勤次数': 5 },
        { '姓名': '周九', '出勤次数': 8 },
        { '姓名': '吴十', '出勤次数': 7 }
    ];
    
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(templateData);
    
    // 设置列宽
    ws['!cols'] = [
        { wch: 15 }, // 姓名列
        { wch: 12 }  // 出勤次数列
    ];
    
    // 添加工作表到工作簿
    XLSX.utils.book_append_sheet(wb, ws, '抽奖名单');
    
    // 生成Excel文件并下载
    const fileName = '抽奖名单模板.xlsx';
    XLSX.writeFile(wb, fileName);
    
    // 提示用户
    fileInfo.textContent = `模板文件 "${fileName}" 已下载，请填写数据后重新导入`;
    fileInfo.style.color = '#4caf50';
    setTimeout(() => {
        fileInfo.textContent = '';
    }, 5000);
}

// Excel文件读取
excelFileInput.addEventListener('change', handleFileUpload);

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    fileInfo.textContent = `已选择文件: ${file.name}`;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            // 解析数据
            nameList = jsonData.map((row, index) => {
                // 尝试不同的列名
                const name = row['姓名'] || row['名字'] || row['name'] || row['Name'] || Object.values(row)[0];
                const attendance = row['出勤次数'] || row['出勤'] || row['attendance'] || row['Attendance'] || Object.values(row)[1] || 0;
                
                return {
                    id: index,
                    name: String(name || '').trim(),
                    attendance: parseInt(attendance) || 0,
                    drawn: false
                };
            }).filter(item => item.name); // 过滤掉空名字
            
            // 重置抽奖状态
            drawnList = [];
            nameList.forEach(item => item.drawn = false);
            
            updateTable();
            updateStats();
            // 重新启用自动建议
            drawCountAuto = true;
            updateDrawCount();
            drawBtn.disabled = false;
            
            // 绘制转盘
            drawWheel();
            
            alert(`成功导入 ${nameList.length} 条数据！`);
        } catch (error) {
            alert('文件读取失败，请确保文件格式正确！\n' + error.message);
            console.error(error);
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// 更新表格
function updateTable() {
    const showAll = showAllCheckbox.checked;
    const filteredList = showAll ? nameList : nameList.filter(item => item.attendance >= 7);
    
    if (filteredList.length === 0) {
        tableBody.innerHTML = '<tr class="empty-row"><td colspan="5" class="empty-message">暂无数据</td></tr>';
        return;
    }
    
    tableBody.innerHTML = filteredList.map((item, index) => {
        const isEligible = item.attendance >= 7;
        const isDrawn = item.drawn;
        const rowClass = isEligible ? 'eligible' : 'ineligible';
        const statusClass = isEligible ? 'status-eligible' : 'status-ineligible';
        const statusText = isEligible ? '可抽奖' : '不可抽奖';
        const drawnText = isDrawn ? ' (已抽中)' : '';
        
        return `
            <tr class="${rowClass}" data-id="${item.id}">
                <td>${index + 1}</td>
                <td>${item.name}</td>
                <td>${item.attendance}/8</td>
                <td>
                    <span class="status-badge ${statusClass}">${statusText}${drawnText}</span>
                </td>
                <td>
                    <div class="action-buttons">
                        <button class="btn btn-edit" onclick="editItem(${item.id})">编辑</button>
                        <button class="btn btn-delete" onclick="deleteItem(${item.id})">删除</button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

// 更新统计信息
function updateStats() {
    const total = nameList.length;
    const eligible = nameList.filter(item => item.attendance >= 7 && !item.drawn).length;
    const drawn = drawnList.length;
    
    totalCountSpan.textContent = total;
    eligibleCountSpan.textContent = eligible;
    drawnCountSpan.textContent = drawn;
}

// 是否自动根据人数计算抽奖数量（第一次导入/重置时自动，开始抽奖后不再改动用户设置）
let drawCountAuto = true;

// 更新抽奖数量（5:1比例）
function updateDrawCount() {
    const eligible = nameList.filter(item => item.attendance >= 7 && !item.drawn).length;
    const suggestedCount = Math.floor(eligible / 5);
    
    // 只在允许自动模式时才改动实际输入值
    if (drawCountAuto) {
        drawCountInput.value = suggestedCount || 1;
    }

    drawCountInput.placeholder = `建议: ${suggestedCount || 1}人`;
}

// 筛选切换
showAllCheckbox.addEventListener('change', function() {
    updateTable();
    // 注意：转盘只显示可抽奖人员，不受筛选影响
});

// 编辑功能
function editItem(id) {
    const item = nameList.find(i => i.id === id);
    if (!item) return;
    
    currentEditIndex = id;
    editNameInput.value = item.name;
    editAttendanceInput.value = item.attendance;
    editModal.style.display = 'block';
}

// 删除功能
function deleteItem(id) {
    if (!confirm('确定要删除这条记录吗？')) return;
    
    nameList = nameList.filter(item => item.id !== id);
    drawnList = drawnList.filter(item => item.id !== id);
    
    updateTable();
    updateStats();
    updateDrawCount();
    drawWheel();
}

// 保存编辑
saveEditBtn.addEventListener('click', function() {
    if (currentEditIndex === -1) return;
    
    const name = editNameInput.value.trim();
    const attendance = parseInt(editAttendanceInput.value) || 0;
    
    if (!name) {
        alert('姓名不能为空！');
        return;
    }
    
    if (attendance < 0 || attendance > 8) {
        alert('出勤次数必须在0-8之间！');
        return;
    }
    
    const item = nameList.find(i => i.id === currentEditIndex);
    if (item) {
        item.name = name;
        item.attendance = attendance;
        
        // 如果出勤次数不足，取消已抽中状态
        if (attendance < 7 && item.drawn) {
            item.drawn = false;
            drawnList = drawnList.filter(i => i.id !== item.id);
        }
        
        updateTable();
        updateStats();
        updateDrawCount();
        drawWheel();
    }
    
    editModal.style.display = 'none';
    currentEditIndex = -1;
});

// 取消编辑
cancelEditBtn.addEventListener('click', function() {
    editModal.style.display = 'none';
    currentEditIndex = -1;
});

closeModal.addEventListener('click', function() {
    editModal.style.display = 'none';
    currentEditIndex = -1;
});

window.addEventListener('click', function(event) {
    if (event.target === editModal) {
        editModal.style.display = 'none';
        currentEditIndex = -1;
    }
});

// 当前转盘旋转角度（度）
let wheelRotation = 0;

// 绘制转盘（使用 Canvas，实现扇形和名字沿扇形半径排布）
function drawWheel() {
    const eligibleList = nameList.filter(item => item.attendance >= 7 && !item.drawn);

    // 重置旋转角度展示
    wheelRotation = 0;
    if (wheelGraphic) {
        wheelGraphic.style.transform = 'rotate(0deg)';
    }

    const size = wheelCanvas.width;
    const cx = size / 2;
    const cy = size / 2;
    const radius = size / 2 - 12; // 预留边框

    // 清空画布
    wheelCtx.clearRect(0, 0, size, size);

    if (eligibleList.length === 0) {
        wheelText.textContent = '等待抽奖';
        return;
    }

    const anglePerSegment = (Math.PI * 2) / eligibleList.length;

    eligibleList.forEach((item, index) => {
        const startAngle = index * anglePerSegment;
        const endAngle = startAngle + anglePerSegment;

        // 扇形背景
        wheelCtx.beginPath();
        wheelCtx.moveTo(cx, cy);
        wheelCtx.arc(cx, cy, radius, startAngle, endAngle, false);
        wheelCtx.closePath();
        const hue = (index * 137.5) % 360;
        wheelCtx.fillStyle = `hsl(${hue}, 70%, 60%)`;
        wheelCtx.fill();

        // 扇形边框
        wheelCtx.strokeStyle = '#ffffff';
        wheelCtx.lineWidth = 2;
        wheelCtx.stroke();

        // 名字（沿扇形半径显示）
        const midAngle = startAngle + anglePerSegment / 2;
        const textRadius = radius * 0.7;

        wheelCtx.save();
        wheelCtx.translate(cx, cy);
        wheelCtx.rotate(midAngle);
        wheelCtx.textAlign = 'center';
        wheelCtx.textBaseline = 'middle';
        wheelCtx.fillStyle = '#ffffff';
        wheelCtx.font = 'bold 16px "Microsoft YaHei", "Segoe UI", sans-serif';
        wheelCtx.fillText(item.name, textRadius, 0);
        wheelCtx.restore();
    });

    wheelText.textContent = `共${eligibleList.length}人`;
}

// 抽奖功能
drawBtn.addEventListener('click', startLottery);

let currentDrawIndex = 0;
let drawQueue = [];

function startLottery() {
    const eligibleList = nameList.filter(item => item.attendance >= 7 && !item.drawn);
    
    if (eligibleList.length === 0) {
        alert('没有可抽奖的人员了！');
        return;
    }
    
    const drawCount = parseInt(drawCountInput.value) || 1;
    const actualDrawCount = Math.min(drawCount, eligibleList.length);
    
    if (actualDrawCount <= 0) {
        alert('抽取人数必须大于0！');
        return;
    }
    
    // 一旦开始抽奖，就不再自动修改输入框的值
    drawCountAuto = false;
    
    // 禁用按钮
    drawBtn.disabled = true;
    
    // 随机抽取不重复的人员
    const shuffled = [...eligibleList].sort(() => Math.random() - 0.5);
    const winners = shuffled.slice(0, actualDrawCount);
    
    // 设置抽奖队列
    drawQueue = winners;
    currentDrawIndex = 0;
    
    // 开始第一次抽奖
    performSingleDraw();
}

function performSingleDraw() {
    if (currentDrawIndex >= drawQueue.length) {
        // 所有抽奖完成
        drawBtn.disabled = nameList.filter(item => item.attendance >= 7 && !item.drawn).length > 0;
        return;
    }
    
    const winner = drawQueue[currentDrawIndex];
    const eligibleList = nameList.filter(item => item.attendance >= 7 && !item.drawn);
    
    // 找到中奖者在转盘上的索引
    const winnerIndex = eligibleList.findIndex(item => item.id === winner.id);
    
    if (winnerIndex === -1) {
        currentDrawIndex++;
        performSingleDraw();
        return;
    }

    // 计算转盘需要旋转的角度（指针固定在中心，朝上）
    const anglePerSegment = 360 / eligibleList.length;
    const segmentCenter = winnerIndex * anglePerSegment + anglePerSegment / 2;

    // 当前转盘角度（0-360 归一化）
    const normalizedCurrent = ((wheelRotation % 360) + 360) % 360;

    // 多转几圈增加动画效果
    const extraRotations = 4; // 多转 4 圈
    // 指针朝上，对应 270°；目标是让中奖扇形的中线最终落在该方向
    const pointerDirection = 270;
    const deltaRotation = extraRotations * 360 + (pointerDirection - segmentCenter) - normalizedCurrent;
    const totalRotation = wheelRotation + deltaRotation;
    wheelRotation = totalRotation;

    // 设置转盘旋转（使用 CSS transition 实现平滑动画）
    if (wheelGraphic) {
        wheelGraphic.style.transform = `rotate(${totalRotation}deg)`;
    }
    wheelText.textContent = '抽奖中...';
    
    // 等待动画完成（与 CSS 中 3.5s 动画时长匹配）
    setTimeout(() => {
        // 动画完成
        wheelText.textContent = winner.name;
        
        // 标记为已抽中
        winner.drawn = true;
        drawnList.push(winner);
        
        // 显示中奖结果
        displaySingleResult(winner);
        
        // 更新界面
        updateTable();
        updateStats();
        updateDrawCount();
        
        // 重新绘制转盘（移除已抽中的人）
        setTimeout(() => {
            drawWheel();
            
            // 继续下一次抽奖
            currentDrawIndex++;
            if (currentDrawIndex < drawQueue.length) {
                setTimeout(() => {
                    performSingleDraw();
                }, 1000);
            } else {
                drawBtn.disabled = nameList.filter(item => item.attendance >= 7 && !item.drawn).length > 0;
            }
        }, 1500);
    }, 3500);
}

function displaySingleResult(winner) {
    const resultItem = document.createElement('div');
    resultItem.className = 'result-item';
    resultItem.textContent = `🎉 ${winner.name} (出勤: ${winner.attendance}/8)`;
    resultsList.appendChild(resultItem);
    
    // 移除空提示
    const emptyResults = resultsList.querySelector('.empty-results');
    if (emptyResults) {
        emptyResults.remove();
    }
    
    // 显示清空按钮
    clearResultsBtn.style.display = 'block';
    
    // 滚动到底部
    resultsList.scrollTop = resultsList.scrollHeight;
}

// 重置抽奖
resetBtn.addEventListener('click', function() {
    if (!confirm('确定要重置所有抽奖记录吗？这将清空所有中奖记录。')) return;
    
    nameList.forEach(item => item.drawn = false);
    drawnList = [];
    resultsList.innerHTML = '<div class="empty-results">暂无中奖记录</div>';
    clearResultsBtn.style.display = 'none';
    
    // 重置转盘和指针
    wheelRotation = 0;
    if (wheelGraphic) {
        wheelGraphic.style.transform = 'rotate(0deg)';
    }
    
    updateTable();
    updateStats();
    // 重置时重新允许自动建议
    drawCountAuto = true;
    updateDrawCount();
    drawWheel();
    drawBtn.disabled = nameList.filter(item => item.attendance >= 7).length === 0;
});

// 清空中奖记录
clearResultsBtn.addEventListener('click', function() {
    if (!confirm('确定要清空中奖记录吗？')) return;
    
    resultsList.innerHTML = '<div class="empty-results">暂无中奖记录</div>';
    clearResultsBtn.style.display = 'none';
});

// 抽奖数量输入变化时更新
drawCountInput.addEventListener('input', function() {
    // 用户手动修改后，不再自动覆盖这个值
    drawCountAuto = false;

    const eligible = nameList.filter(item => item.attendance >= 7 && !item.drawn).length;
    const value = parseInt(this.value) || 0;
    
    if (value > eligible) {
        this.value = eligible;
        alert(`最多只能抽取 ${eligible} 人！`);
    }
});

