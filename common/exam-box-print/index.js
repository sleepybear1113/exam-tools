// 初始化时生成所有20个输入框
function initializeSubjectInputs() {
    const container = document.getElementById('subjectInputs');
    container.innerHTML = '';

    for (let i = 1; i <= 20; i++) {
        const subjectDiv = document.createElement('div');
        subjectDiv.className = 'subject-row';
        subjectDiv.innerHTML = `
                    <div class="subject-inputs">
                        <label for="subjectType${i}">科目${i}名称</label>
                        <input type="text" class="input-field" id="subjectType${i}">

                        <label for="totalBags${i}">袋数总计</label>
                        <input type="number" class="input-field" id="totalBags${i}" min="1">

                        <label for="bagsPerBox${i}">每箱袋数</label>
                        <input type="number" class="input-field" id="bagsPerBox${i}" min="1">

                        <label for="maxExtraBags${i}">尾箱超限袋数</label>
                        <input type="number" class="input-field" id="maxExtraBags${i}" value="0">
                    </div>
                `;
        container.appendChild(subjectDiv);
    }

    updateVisibleInputs();
}

// 更新显示的输入框数量
function updateVisibleInputs() {
    const count = parseInt(document.getElementById('subjectCount').value);
    const rows = document.getElementsByClassName('subject-row');

    // 更新每行的显示状态
    for (let i = 0; i < rows.length; i++) {
        if (i < count) {
            rows[i].classList.add('visible');
        } else {
            rows[i].classList.remove('visible');
        }
    }
}

// 页面加载时初始化
window.onload = function () {
    // 初始化选择器
    const select = document.getElementById('subjectCount');
    for (let i = 1; i <= 20; i++) {
        const option = document.createElement('option');
        let s = String(i);
        option.value = s;
        option.textContent = s;
        select.appendChild(option);
    }

    // 初始化所有输入框
    initializeSubjectInputs();

    let subjectDataPlaceholder="每行一个科目，数据用空格分隔。\n格式：科目名称 袋数总计 每箱袋数 尾箱超限袋数\n示例：\nCET4     191    60    10\n学考语文 122    30    3\n兼容Excel。可以在Excel中将科目的4列数据填写后，将数据复制到本输入框（不要复制标题）"
    document.getElementById('subjectData').placeholder = subjectDataPlaceholder + "";

    // 添加事件监听器
    addEventListeners();
}

// 添加事件监听器函数
function addEventListeners() {
    // 监听地区输入框变化
    const regionInput = document.getElementById('region');
    regionInput.addEventListener('input', function() {
        // 使用防抖，避免频繁触发
        clearTimeout(this.debounceTimer);
        this.debounceTimer = setTimeout(() => {
            if (hasValidSubjects()) {
                generateLabels();
            }
        }, 500);
    });

    // 监听单页重复显示单选按钮变化
    const duplicateRadios = document.querySelectorAll('input[name="duplicate"]');
    duplicateRadios.forEach(radio => {
        radio.addEventListener('change', function() {
            if (hasValidSubjects()) {
                generateLabels();
            }
        });
    });

    // 监听科目数量变化
    const subjectCountSelect = document.getElementById('subjectCount');
    subjectCountSelect.addEventListener('change', function() {
        updateVisibleInputs();
        // 延迟一点时间让输入框更新完成
        setTimeout(() => {
            if (hasValidSubjects()) {
                generateLabels();
            }
        }, 100);
    });

    // 监听科目输入框变化
    for (let i = 1; i <= 20; i++) {
        const subjectTypeInput = document.getElementById(`subjectType${i}`);
        const totalBagsInput = document.getElementById(`totalBags${i}`);
        const bagsPerBoxInput = document.getElementById(`bagsPerBox${i}`);
        const maxExtraBagsInput = document.getElementById(`maxExtraBags${i}`);

        // 为每个输入框添加监听器
        [subjectTypeInput, totalBagsInput, bagsPerBoxInput, maxExtraBagsInput].forEach(input => {
            input.addEventListener('input', function() {
                // 使用防抖，避免频繁触发
                clearTimeout(this.debounceTimer);
                this.debounceTimer = setTimeout(() => {
                    if (hasValidSubjects()) {
                        generateLabels();
                    }
                }, 500);
            });
        });
    }
}

// 检查是否有有效的科目数据
function hasValidSubjects() {
    const visibleSubjectCount = parseInt(document.getElementById('subjectCount').value) || 0;
    
    for (let i = 1; i <= visibleSubjectCount; i++) {
        const subjectType = document.getElementById(`subjectType${i}`).value;
        const totalBags = parseInt(document.getElementById(`totalBags${i}`).value) || 0;
        const bagsPerBox = parseInt(document.getElementById(`bagsPerBox${i}`).value) || 0;

        if (subjectType && totalBags > 0 && bagsPerBox > 0) {
            return true;
        }
    }
    return false;
}

// generateLabels 函数修改
function generateLabels() {
    const region = document.getElementById('region').value || '地区名称';
    const printArea = document.getElementById('printArea');
    const isDuplicate = document.querySelector('input[name="duplicate"]:checked').value === "1";
    printArea.innerHTML = '';

    // 重置页面计数器
    printArea.style.counterReset = 'page-counter';

    let boxes = [];

    // 获取当前显示的科目数量
    const visibleSubjectCount = parseInt(document.getElementById('subjectCount').value) || 0;

    // 只处理可见的科目（前N个）
    for (let i = 1; i <= visibleSubjectCount; i++) {
        const subjectType = document.getElementById(`subjectType${i}`).value;
        const totalBags = parseInt(document.getElementById(`totalBags${i}`).value) || 0;
        const bagsPerBox = parseInt(document.getElementById(`bagsPerBox${i}`).value) || 0;
        const maxExtraBags = parseInt(document.getElementById(`maxExtraBags${i}`).value) || 0;

        if (subjectType && totalBags > 0 && bagsPerBox > 0) {
            let boxCounts = calculateBoxCounts(totalBags, bagsPerBox, maxExtraBags);
            let cumulativeBags = 0;

            boxCounts.forEach((count, index) => {
                cumulativeBags += count;
                boxes.push(createBox({
                    region: region,
                    boxNumber: `${index + 1}/${boxCounts.length}`,
                    count: count,
                    type: subjectType,
                    totalBags: totalBags,
                    cumulativeBags: cumulativeBags
                }));
            });
        }
    }

    let htmlElements = [];
    // 分页显示标签
    if (isDuplicate) {
        // 重复显示模式：每页显示相同的内容
        for (let i = 0; i < boxes.length; i++) {
            const page = document.createElement('div');
            page.className = 'page';

            // 上半部分
            const container1 = document.createElement('div');
            container1.className = 'box-container';
            container1.appendChild(boxes[i].cloneNode(true));
            page.appendChild(container1);

            // 下半部分（复制上半部分）
            const container2 = document.createElement('div');
            container2.className = 'box-container';
            container2.appendChild(boxes[i].cloneNode(true));
            page.appendChild(container2);

            htmlElements.push(page);
        }
    } else {
        // 原有的显示模式：每页显示两个不同的标签
        for (let i = 0; i < boxes.length; i += 2) {
            const page = document.createElement('div');
            page.className = 'page';

            const container1 = document.createElement('div');
            container1.className = 'box-container';
            container1.appendChild(boxes[i]);
            page.appendChild(container1);

            if (i + 1 < boxes.length) {
                const container2 = document.createElement('div');
                container2.className = 'box-container';
                container2.appendChild(boxes[i + 1]);
                page.appendChild(container2);
            }

            htmlElements.push(page);
        }
    }

    setTimeout(() => htmlElements.forEach(e => printArea.appendChild(e)), 50);
}

function createBox({region, boxNumber, count, type, totalBags, cumulativeBags}) {
    const box = document.createElement('div');
    box.className = 'box';

    box.innerHTML = `
                <div class="region">${region}</div>
                <div class="box-number">
                    <div class="box-number-row">
                        <span class="box-label">箱号</span>
                        <span class="box-number-content">${boxNumber}</span>
                    </div>
                    <div class="divider"></div>
                    <div class="box-number-row">
                        <span class="box-label">本箱袋数</span>
                        <span class="box-number-content">${count}</span>
                    </div>
                    <div class="divider"></div>
                    <div class="box-number-row">
                        <span class="box-label">袋数</span>
                        <span class="box-number-content">${cumulativeBags}/${totalBags}</span>
                    </div>
                </div>
                <div class="subject-type">${type}</div>
            `;
    return box;
}

// 新增计算箱数的函数
function calculateBoxCounts(totalBags, bagsPerBox, maxExtraBags) {
    let boxCounts = [];
    let remainingBags = totalBags;

    // 如果总数小于等于每箱袋数，直接返回一箱
    if (totalBags <= bagsPerBox) {
        return [totalBags];
    }

    // 先计算完整的箱子
    while (remainingBags > bagsPerBox) {
        boxCounts.push(bagsPerBox);
        remainingBags -= bagsPerBox;
    }

    // 处理最后的不完整箱
    if (remainingBags > 0) {
        // 检查是否需要合并最后两箱
        if (boxCounts.length > 0 &&
            (remainingBags + bagsPerBox) <= (bagsPerBox + maxExtraBags)) {
            // 合并最后两箱
            boxCounts[boxCounts.length - 1] += remainingBags;
        } else {
            // 添加新的一箱
            boxCounts.push(remainingBags);
        }
    }

    return boxCounts;
}

// 切换手动输入区域的显示/隐藏
function toggleManualInput() {
    const area = document.getElementById('manualInputArea');
    const btn = document.getElementById('toggleBtn');
    const isVisible = area.classList.contains('visible');

    if (isVisible) {
        area.classList.remove('visible');
        btn.textContent = '手动输入';
    } else {
        area.classList.add('visible');
        btn.textContent = '隐藏输入';
    }
}

// 修改样例填充函数
function fillSampleData() {
    const textarea = document.getElementById('subjectData');
    const currentContent = textarea.value.trim();

    // 检查是否已有内容
    if (currentContent) {
        if (!confirm('文本框中已有内容，是否清空并填充样例数据？')) {
            return;
        }
    }

    textarea.value = "CET4\t191\t60\t10\nCET6\t122\t30\t5\n学考语文\t88\t40\t8\n日语\t156\t50\t10\n13000英语专升本\t95\t35\t5";
}

// 修改解析函数，增加数据验证
function parseSubjectData() {
    const data = document.getElementById('subjectData').value;
    const lines = data.trim().split('\n');
    const validLines = lines.filter(line => line.trim()); // 过滤空行

    // 处理并验证每行数据
    const validData = validLines.filter(line => {
        const parts = line.trim().split(/\s+/);
        // 检查是否至少有科目名称、袋数总计和每箱袋数
        return parts.length >= 3 &&
            !isNaN(parts[1]) && // 验证袋数总计是数字
            !isNaN(parts[2]);   // 验证每箱袋数是数字
    });

    if (validData.length === 0) {
        alert('没有找到有效的数据行！');
        return;
    }

    // 设置科目数量
    const select = document.getElementById('subjectCount');
    select.value = validData.length;
    updateVisibleInputs();

    // 处理每行数据
    validData.forEach((line, index) => {
        const parts = line.trim().split(/\s+/);
        const subjectIndex = index + 1;

        // 设置科目名称
        document.getElementById(`subjectType${subjectIndex}`).value = parts[0];
        // 设置袋数总计
        document.getElementById(`totalBags${subjectIndex}`).value = parts[1];
        // 设置每箱袋数
        document.getElementById(`bagsPerBox${subjectIndex}`).value = parts[2];
        // 设置尾箱超限袋数（如果有且是有效数字）
        document.getElementById(`maxExtraBags${subjectIndex}`).value =
            (parts[3] && !isNaN(parts[3])) ? parts[3] : '0';
    });

    // 解析完成后自动生成标签
    setTimeout(() => {
        if (hasValidSubjects()) {
            generateLabels();
        }
    }, 100);
}
