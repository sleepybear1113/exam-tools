// 测试样例数据
const testSample = `张三 - 北京大学
李四 - 清华大学
王五 - 复旦大学
来自浙江嘉兴的测试人员aaaa
赵六 - 上海交通大学
钱七 - 浙江大学
孙八 - 南京大学
周九 - 武汉大学
吴十 - 华中科技大学
这是另外一条水印数据     和空格后面的内容
abcdefghijklmnopqrstuvwx`;

// 纸张尺寸定义 (mm)
const PAPER_SIZES = {
    A3: {portrait: [297, 420], landscape: [420, 297]},
    A4: {portrait: [210, 297], landscape: [297, 210]},
};

// 浏览器默认打印边距 (mm) - 使用更小的边距来适应浏览器默认设置
const PRINT_MARGIN = 5;

function mmToPx(mm) {
    // 96 DPI 转换: 1mm ≈ 3.78px
    return mm * 3.78;
}

function padNumber(num, length) {
    return num.toString().padStart(length, '0');
}

function generate() {
    const loading = document.getElementById('loading');
    const preview = document.getElementById('preview');

    loading.style.display = 'block';
    preview.style.display = 'none';

    setTimeout(() => {
        generateWatermarks();
        loading.style.display = 'none';
        preview.style.display = 'block';
    }, 200);
}

function loadTestSample() {
    const input = document.getElementById('input');
    const currentContent = input.value.trim();


    // 如果当前有内容，询问是否清空
    if (currentContent) {
        if (confirm('当前已有内容，是否清空并加载测试样例？')) {
            input.value = testSample;
            generate();
        }
    } else {
        // 如果没有内容，直接加载
        input.value = testSample;
        generate();
    }
}

function setPageSize() {
    const paperSize = document.getElementById('paperSize').value;
    const orientation = document.getElementById('orientation').value;

    // 动态设置CSS变量
    const [widthMm, heightMm] = PAPER_SIZES[paperSize][orientation];
    document.documentElement.style.setProperty('--page-width', widthMm + 'mm');
    document.documentElement.style.setProperty('--page-height', heightMm + 'mm');
}

function generateWatermarks() {
    // 确保页面尺寸设置正确
    setPageSize();

    const inputLines = document.getElementById('input').value.trim().split('\n').filter(l => l.trim() !== '');
    if (inputLines.length === 0) {
        alert('请输入至少一行水印文本内容！');
        return;
    }

    const paperSize = document.getElementById('paperSize').value;
    const orientation = document.getElementById('orientation').value;
    const gapX = parseFloat(document.getElementById('gapX').value);
    const gapY = parseFloat(document.getElementById('gapY').value);
    const angle = parseFloat(document.getElementById('angleDeg').value);
    const fontSize = parseInt(document.getElementById('fontSize').value, 10);
    const fontFamily = document.getElementById('fontFamily').value;
    const opacity = parseFloat(document.getElementById('opacity').value);
    const showLineNumber = document.getElementById('showLineNumber').checked;

    const preview = document.getElementById('preview');
    preview.innerHTML = `
            <div class="preview-header">
                <h3>📋 预览效果</h3>
                <p>以下是生成的水印预览，每页对应一行输入文本</p>
            </div>
        `;

    // 获取纸张尺寸
    const [pageWidthMm, pageHeightMm] = PAPER_SIZES[paperSize][orientation];

    // 计算可用内容区域 (减去打印边距)
    const contentWidthMm = pageWidthMm - 2 * PRINT_MARGIN;
    const contentHeightMm = pageHeightMm - 2 * PRINT_MARGIN;

    // 转换为像素
    const pageWidthPx = mmToPx(pageWidthMm);
    const pageHeightPx = mmToPx(pageHeightMm);
    const contentWidthPx = mmToPx(contentWidthMm);
    const contentHeightPx = mmToPx(contentHeightMm);

    // 创建测量元素
    const measureDiv = document.createElement('div');
    measureDiv.style.position = 'absolute';
    measureDiv.style.visibility = 'hidden';
    measureDiv.style.fontSize = fontSize + 'px';
    measureDiv.style.fontFamily = fontFamily;
    measureDiv.style.whiteSpace = 'nowrap';
    document.body.appendChild(measureDiv);

    const maxLineNum = inputLines.length;
    const maxDigits = maxLineNum.toString().length;

    inputLines.forEach((line, index) => {
        const text = showLineNumber
            ? `${padNumber(index + 1, maxDigits)} ${line.trim()}`
            : line.trim();

        // 保留空格和Tab
        const htmlText = text.replace(/ /g, '&nbsp;').replace(/\t/g, '&nbsp;&nbsp;&nbsp;&nbsp;');

        // 测量文字尺寸
        measureDiv.innerHTML = htmlText;
        const textWidthPx = measureDiv.offsetWidth;
        const textHeightPx = fontSize;

        // 计算旋转后的文字尺寸
        const angleRad = angle * Math.PI / 180;
        const rotatedWidth = Math.abs(textWidthPx * Math.cos(angleRad)) + Math.abs(textHeightPx * Math.sin(angleRad));
        const rotatedHeight = Math.abs(textWidthPx * Math.sin(angleRad)) + Math.abs(textHeightPx * Math.cos(angleRad));

        // 创建页面容器 - 使用完整的纸张尺寸
        const page = document.createElement('div');
        page.className = 'page';
        // 使用完整的纸张尺寸，让CSS处理边距
        page.style.width = pageWidthPx + 'px';
        page.style.height = pageHeightPx + 'px';
        page.style.position = 'relative';
        page.style.overflow = 'hidden';
        page.style.boxSizing = 'border-box';

        const contentDiv = document.createElement('div');
        contentDiv.className = 'page-content';
        contentDiv.style.width = '100%';
        contentDiv.style.height = '100%';
        contentDiv.style.position = 'relative';
        contentDiv.style.overflow = 'hidden';

        // 计算水印间距 (像素)
        const gapXPx = gapX * 3.78;
        const gapYPx = gapY * 3.78;

        // 计算水印网格 - 基于用户设置的间距
        const cols = Math.max(1, Math.floor(contentWidthPx / gapXPx) + 1);
        const rows = Math.max(1, Math.floor(contentHeightPx / gapYPx) + 1);

        // 生成水印
        for (let row = 0; row < rows; row++) {
            for (let col = 0; col < cols; col++) {
                const watermark = document.createElement('div');
                watermark.className = 'watermark';
                watermark.innerHTML = htmlText;

                // 计算位置 (基于用户设置的间距)
                const x = col * gapXPx;
                const y = row * gapYPx;

                // 检查水印是否完全在页面内（考虑旋转后的尺寸）
                const watermarkRight = x + rotatedWidth;
                const watermarkBottom = y + rotatedHeight;

                // 如果水印完全在页面内，则添加
                if (watermarkRight <= contentWidthPx && watermarkBottom <= contentHeightPx) {
                    watermark.style.left = (PRINT_MARGIN * 3.78 + x) + 'px';
                    watermark.style.top = (PRINT_MARGIN * 3.78 + y) + 'px';
                    watermark.style.transform = `rotate(${angle}deg)`;
                    watermark.style.fontSize = fontSize + 'px';
                    watermark.style.opacity = opacity;
                    watermark.style.fontFamily = fontFamily;
                    watermark.style.position = 'absolute';

                    contentDiv.appendChild(watermark);
                }
                // 如果水印部分超出，则截断显示
                else if (x < contentWidthPx && y < contentHeightPx) {
                    watermark.style.left = (PRINT_MARGIN * 3.78 + x) + 'px';
                    watermark.style.top = (PRINT_MARGIN * 3.78 + y) + 'px';
                    watermark.style.transform = `rotate(${angle}deg)`;
                    watermark.style.fontSize = fontSize + 'px';
                    watermark.style.opacity = opacity;
                    watermark.style.fontFamily = fontFamily;
                    watermark.style.position = 'absolute';

                    // 添加裁剪样式，确保超出部分不显示
                    watermark.style.overflow = 'hidden';
                    watermark.style.maxWidth = Math.max(0, contentWidthPx - x) + 'px';
                    watermark.style.maxHeight = Math.max(0, contentHeightPx - y) + 'px';

                    contentDiv.appendChild(watermark);
                }
            }
        }

        page.appendChild(contentDiv);
        preview.appendChild(page);
    });

    document.body.removeChild(measureDiv);
}

// 页面加载完成后自动聚焦到输入框
document.addEventListener('DOMContentLoaded', function () {
    document.getElementById('input').focus();

    // 初始化页面尺寸
    setPageSize();

    // 添加实时预览功能（仅对非文本输入控件）
    const inputs = ['paperSize', 'orientation', 'gapX', 'gapY', 'angleDeg', 'fontSize', 'opacity', 'showLineNumber', 'fontFamily'];
    inputs.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.addEventListener('change', function () {
                // 如果是纸张设置改变，更新页面尺寸
                if (id === 'paperSize' || id === 'orientation') {
                    setPageSize();
                }
                if (document.getElementById('input').value.trim()) {
                    generate();
                }
            });
        }
    });
});

// 设置 placeholder 文本
document.getElementById('input').placeholder = testSample;
