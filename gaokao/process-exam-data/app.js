function getActiveComboSpecs() {
    const specs = [];
    COMBO_SPECS.combos3.forEach(spec => {
        if (comboUsage[spec.key] > 0) specs.push(spec);
    });
    COMBO_SPECS.combos2.forEach(spec => {
        if (comboUsage[spec.key] > 0) specs.push(spec);
    });
    COMBO_SPECS.combos1.forEach(spec => {
        if (comboUsage[spec.key] > 0) specs.push(spec);
    });
    if (comboUsage[COMBO_SPECS.comboZero.key] > 0) specs.push(COMBO_SPECS.comboZero);
    if (comboUsage[COMBO_SPECS.comboPartial.key] > 0) specs.push(COMBO_SPECS.comboPartial);
    return specs;
}

function buildComboColumns() {
    return getActiveComboSpecs().map(spec => ({key: spec.key, label: spec.label}));
}

(function () {
    'use strict';

    const SUBJECTS = [
        {key: 'KM_YW', label: '语文'},
        {key: 'KM_SX', label: '数学'},
        {key: 'KM_YY', label: '外语'},
        {key: 'KM_SZ', label: '政治'},
        {key: 'KM_LS', label: '历史'},
        {key: 'KM_DL', label: '地理'},
        {key: 'KM_WL', label: '物理'},
        {key: 'KM_HX', label: '化学'},
        {key: 'KM_SW', label: '生物'},
        {key: 'KM_JS', label: '技术'},
    ];
    const SUBJECT_KEY_SET = new Set(SUBJECTS.map(s => s.key));
    const NON_LANG_SUBJECTS = SUBJECTS.filter(subj => subj.key !== 'KM_YY');
    const NON_LANG_YS_SUBJECTS = SUBJECTS.filter(subj => subj.key !== 'KM_YY' && subj.key !== 'KM_YW' && subj.key !== 'KM_SX');

    const LANGUAGE_CODES = [
        ['1', '英语'],
        ['2', '俄语'],
        ['3', '日语'],
        ['4', '德语'],
        ['5', '法语'],
        ['6', '西班牙语'],
        ['0', '未选'],
    ];
    const WYYZ_CODE_TO_LABEL = LANGUAGE_CODES.reduce((acc, [code, name]) => {
        acc[code] = name;
        return acc;
    }, {});
    const LANGUAGE_LABELS = LANGUAGE_CODES.map(([, name]) => name);

    const YSTYSK_MAP = {
        '30': '书法类',
        '31': '美术与设计类',
        '32': '音乐类',
        '33': '舞蹈类',
        '34': '表（导）演类',
        '36': '播音主持类',
        '39': '戏曲类',
        '41': '体育类',
    };
    const YSTYSK_LABELS = Object.values(YSTYSK_MAP);

    const STORAGE_KEYS = {
        bmd: 'processExamData.bmdMapping',
        county: 'processExamData.countyMapping',
    };

    const HEADER_ALIAS = new Map([
        ['BMD', 'BMD'],
        ['报名点', 'BMD'],
        ['报名点代码', 'BMD'],
        ['报名点编号', 'BMD'],
        ['BMDM', 'BMD'],
        ['YDDH', 'YDDH'],
        ['移动电话', 'YDDH'],
        ['联系电话', 'YDDH'],
        ['SHZT', 'SHZT'],
        ['审核状态', 'SHZT'],
        ['YSTYSK', 'YSTYSK'],
        ['YSTS', 'YSTYSK'],
        ['术科', 'YSTYSK'],
        ['术科项目', 'YSTYSK'],
        ['缴费状态', '缴费状态'],
        ['WYYZ', 'WYYZ'],
        ['外语语种', 'WYYZ'],
        ['外语语种代码', 'WYYZ'],
    ]);

    const els = {
        zone: document.getElementById('dropzone'),
        fileInput: document.getElementById('fileInput'),
        btnPick: document.getElementById('btnPick'),
        btnClear: document.getElementById('btnClear'),
        btnToggleMode: document.getElementById('btnToggleMode'),
		btnResetBmd: document.getElementById('btnResetBmd'),
		btnResetCounty: document.getElementById('btnResetCounty'),
        tableHead: document.querySelector('#resultTable thead'),
        tableBody: document.querySelector('#resultTable tbody'),
        tableWrapper: document.getElementById('tableWrapper'),
        copyWrapper: document.getElementById('copyWrapper'),
        copyOutput: document.getElementById('copyOutput'),
        statusLine: document.getElementById('statusLine'),
        loading: document.getElementById('loadingOverlay'),
        loadingText: document.getElementById('loadingText'),
        bmdMappingInput: document.getElementById('bmdMappingInput'),
        countyMappingInput: document.getElementById('countyMappingInput'),
    };

    const defaultMappings = {
        bmd: `73666\t嘉兴市区单独考试报名点
73800\t嘉善县招生考试办公室
73966\t平湖市单招考试报名点
74001\t海宁单独考试报名点（运动训练等）
74101\t海盐县单独考试报名点
74277\t桐乡市单独考试报名点（不参加统考）
93601\t嘉兴市第一中学
93602\t嘉兴市秀州中学
93603\t嘉兴市第三中学
93604\t嘉兴市第四高级中学
93605\t嘉兴市第五高级中学
93606\t北京师范大学附属嘉兴南湖高级中学
93607\t嘉兴市二十一世纪外国语学校
93608\t北京师范大学南湖附属学校
93609\t嘉兴外国语学校
93610\t嘉兴高级中学
93611\t嘉兴一中实验学校
93612\t嘉兴市秀洲现代实验学校
93613\t嘉兴市秀水高级中学
93614\t嘉兴开明中学
93616\t嘉兴市禾光学校
93618\t嘉兴一实学校
93620\t清华附中嘉兴实验高级中学
93622\t嘉州学校
93625\t嘉兴市教育考试院(社会考生报名点)
93699\t秀州中学(新疆班)
93801\t嘉善县高级中学
93802\t嘉善中学
93803\t嘉善第二高级中学
93804\t嘉善县树人学校
93805\t嘉善县综合高级中学
93806\t嘉善县新世纪学校
93807\t嘉善电大附属综合高级中学
93901\t浙江省平湖中学
93903\t平湖市当湖高级中学
93904\t平湖市新华爱心高级中学
93905\t平湖市乍浦高级中学
93907\t平湖杭州湾实验学校
93908\t平湖市毅进卡迪夫公学高级中学有限公司
93909\t平湖市稚川实验中学
94001\t海宁市高级中学
94002\t海宁中学
94003\t海宁市第一中学
94004\t海宁市紫微高级中学
94005\t海宁一中新疆部
94006\t海宁市宏达高级中学
94007\t海宁市南苑高级中学
94008\t海宁市开明学校
94009\t海宁市职业高级中学
94010\t浙江青年艺术学校
94011\t海宁市教育考试中心（普高报名）
94012\t海宁市静安高级中学
94101\t海盐元济高级中学
94102\t海盐高级中学
94103\t海盐第二高级中学
94105\t海盐县教育局招生办
94106\t海盐职成教综合班
94201\t桐乡高级中学
94202\t桐乡第一中学
94203\t桐乡第二中学
94204\t桐乡茅盾中学
94208\t桐乡技师学院
94209\t社会考生报名点（茅盾中学）
94214\t桐乡教师进修学校
94215\t桐乡凤鸣高级中学
83692\t嘉兴技师学院(高职)
83693\t嘉兴市建筑工业学校
83696\t嘉兴市教育考试院（高职）
83698\t嘉兴市秀水中等专业学校
83891\t嘉善中专
83892\t嘉善电大学院
83893\t嘉善信息技术工程学校
83991\t平湖市职业中等专业学校
83992\t嘉兴市交通学校
84001\t海宁市职业高级中学
84003\t海宁市服装职业技术学校
84006\t海宁技师学院
84007\t海宁市卫生学校
84008\t海宁市教育考试中心（职高报名）
84101\t海盐职业教育中心
84102\t海盐县教育局招生办（高职）
84103\t海盐县商贸学校
84201\t桐乡技师学院
84202\t桐乡技师学院南校区（高职）
84204\t桐乡市卫生学校
84209\t桐乡考试中心（高职）
83692\t嘉兴技师学院(高职)
83693\t嘉兴市建筑工业学校
83696\t嘉兴市教育考试院（高职）
83698\t嘉兴市秀水中等专业学校
83891\t嘉善中专
83892\t嘉善电大学院
83893\t嘉善信息技术工程学校
83991\t平湖市职业中等专业学校
83992\t嘉兴市交通学校
84001\t海宁市职业高级中学
84003\t海宁市服装职业技术学校
84006\t海宁技师学院
84007\t海宁市卫生学校
84008\t海宁市教育考试中心（职高报名）
84101\t海盐职业教育中心
84102\t海盐县教育局招生办（高职）
84103\t海盐县商贸学校
84201\t桐乡技师学院
84202\t桐乡技师学院南校区（高职）
84204\t桐乡市卫生学校
84209\t桐乡考试中心（高职）
63601\t嘉兴市第一中学
63602\t嘉兴市秀州中学
63603\t嘉兴市第三中学
63604\t嘉兴市第四高级中学
63605\t嘉兴市第五高级中学
63606\t北京师范大学附属嘉兴南湖高级中学
63607\t嘉兴市二十一世纪外国语学校
63608\t北京师范大学南湖附属学校
63609\t嘉兴外国语学校
63610\t嘉兴高级中学
63611\t嘉兴一中实验学校
63612\t嘉兴市秀洲现代实验学校
63613\t嘉兴市秀水高级中学
63614\t嘉兴市开明中学
63615\t嘉兴一实学校2
63616\t嘉兴市禾光学校
63618\t嘉兴一实学校
63620\t清华嘉兴附中
63625\t嘉兴市教育考试院
63661\t一中2
63662\t秀中2
63664\t四高2
63666\t北师大2
63667\t21世纪2
63669\t嘉外2
63670\t嘉高2
63673\t秀水高中2
63699\t嘉兴市秀州中学新疆部
63801\t嘉善高级中学
63802\t嘉善中学
63803\t嘉善第二高级中学
63804\t嘉善社会考生报名点
63806\t嘉善新世纪学校
63811\t嘉善高级中学（往届）
63812\t嘉善中学（往届）
63813\t嘉善第二高级中学（往届）
63816\t嘉善新世纪学校（往届）
63891\t嘉善中专
63893\t嘉善信息技术工程学校
63901\t浙江省平湖中学
63902\t浙江省平湖中学社会考生报名点
63903\t平湖市当湖高级中学
63904\t平湖市新华爱心高级中学
63905\t平湖市乍浦高级中学
63906\t平湖市乍浦高级中学社会考生报名点
63907\t平湖杭州湾实验学校
63908\t平湖市毅进卡迪夫公学高级中学有限公司
63909\t平湖市稚川实验中学
63910\t平湖市职业中等专业学校
63911\t平湖市卡迪夫高中社会生报名点
64001\t海宁市高级中学
64002\t海宁中学
64003\t海宁市第一中学
64004\t海宁市紫微高级中学
64005\t海宁一中新疆部
64006\t宏达高级中学
64007\t海宁市南苑高级中学
64008\t海宁市静安高级中学
64011\t海宁市教育考试中心
64101\t元济高级中学
64102\t海盐高级中学
64103\t海盐第二高级中学
64104\t海盐县招生办（社会考生报名点）
64201\t桐乡市高级中学
64202\t浙江省桐乡第一中学
64203\t浙江省桐乡第二中学
64204\t桐乡市茅盾中学
64209\t桐乡市教育考试中心（社会报名）
64215\t桐乡市凤鸣高级中学
`,
        county: `36 市本级
38 嘉善县
39 平湖市
40 海宁市
41 海盐县
42 桐乡市`,
    };

    let bmdMap = new Map();
    let countyMap = new Map();

    let resultColumns = [];
    let resultRows = [];
    let currentMode = 'table';
    let lastDataset = null; // { type: 'xuexuan' | 'gaokao', rows, headerSet }

    const COMBO_SPECS = buildComboSpecs();
    const COMBO_KEY_MAP = buildComboKeyMap(COMBO_SPECS);

    const nextFrame = () => new Promise(resolve => setTimeout(resolve, 0));

    init();

    function init() {
        restoreMappings();
        wireEvents();
        enablePageDnd();
        updateStatus('等待上传文件…');
        renderEmptyTable();
    }

    function wireEvents() {
        els.btnPick.addEventListener('click', () => els.fileInput.click());
        els.fileInput.addEventListener('change', (e) => {
            const file = e.target.files && e.target.files[0];
            if (file) handleFile(file);
            els.fileInput.value = '';
        });
        els.btnClear.addEventListener('click', clearResults);
        els.btnToggleMode.addEventListener('click', toggleMode);
		els.btnResetBmd.addEventListener('click', () => resetMapping('bmd'));
		els.btnResetCounty.addEventListener('click', () => resetMapping('county'));

        ['input', 'change'].forEach(evt => {
            els.bmdMappingInput.addEventListener(evt, () => handleMappingChange('bmd'));
            els.countyMappingInput.addEventListener(evt, () => handleMappingChange('county'));
        });
    }

    function enablePageDnd() {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => {
            document.addEventListener(evt, e => {
                e.preventDefault();
                e.stopPropagation();
            });
        });
        document.addEventListener('dragover', () => els.zone.classList.add('dragging'));
        document.addEventListener('dragleave', () => els.zone.classList.remove('dragging'));
        document.addEventListener('drop', (e) => {
            els.zone.classList.remove('dragging');
            const files = e.dataTransfer?.files;
            if (files && files.length) {
                handleFile(files[0]);
            }
        });
    }

    function restoreMappings() {
        const storedBmd = localStorage.getItem(STORAGE_KEYS.bmd);
        const storedCounty = localStorage.getItem(STORAGE_KEYS.county);
        els.bmdMappingInput.value = storedBmd ?? defaultMappings.bmd;
        els.countyMappingInput.value = storedCounty ?? defaultMappings.county;
        parseMappings();
    }

    function handleMappingChange(type) {
        if (type === 'bmd') {
            localStorage.setItem(STORAGE_KEYS.bmd, els.bmdMappingInput.value);
        } else {
            localStorage.setItem(STORAGE_KEYS.county, els.countyMappingInput.value);
        }
        parseMappings();
        if (lastDataset) {
            processCurrentDataset();
        }
    }

    function parseMappings() {
        bmdMap = textToMap(els.bmdMappingInput.value);
        countyMap = textToMap(els.countyMappingInput.value);
    }

	function resetMapping(type) {
		if (type === 'bmd') {
			els.bmdMappingInput.value = defaultMappings.bmd;
			localStorage.setItem(STORAGE_KEYS.bmd, defaultMappings.bmd);
		} else {
			els.countyMappingInput.value = defaultMappings.county;
			localStorage.setItem(STORAGE_KEYS.county, defaultMappings.county);
		}
		parseMappings();
		if (lastDataset) {
			processCurrentDataset();
		}
	}

    function textToMap(text) {
        const map = new Map();
        text.split(/\n+/).forEach(line => {
            const trimmed = line.trim();
            if (!trimmed) return;
            const [code, ...nameParts] = trimmed.split(/\s+/);
            if (!code || !nameParts.length) return;
            map.set(code, nameParts.join(' '));
        });
        return map;
    }

    function toggleMode() {
        if (!resultRows.length) return;
        currentMode = currentMode === 'table' ? 'copy' : 'table';
        if (currentMode === 'copy') {
            els.tableWrapper.classList.add('d-none');
            els.copyWrapper.classList.remove('d-none');
            els.btnToggleMode.textContent = '切换至表格模式';
        } else {
            els.tableWrapper.classList.remove('d-none');
            els.copyWrapper.classList.add('d-none');
            els.btnToggleMode.textContent = '切换至复制模式';
        }
    }

    function clearResults() {
        resultColumns = [];
        resultRows = [];
        lastDataset = null;
        renderEmptyTable();
        updateCopyOutput();
        updateStatus('已清空结果，等待上传文件…');
        els.btnToggleMode.disabled = true;
        if (currentMode === 'copy') toggleMode();
    }

    function renderEmptyTable() {
        els.tableHead.innerHTML = '';
        els.tableBody.innerHTML = `
			<tr>
				<td class="text-center text-muted" colspan="5">暂无数据</td>
			</tr>
		`;
    }

    function updateStatus(text) {
        els.statusLine.textContent = text;
    }

    function showLoading(message) {
        if (message) els.loadingText.textContent = message;
        els.loading.classList.remove('d-none');
        document.body.classList.add('page-is-loading');
    }

    function hideLoading() {
        els.loading.classList.add('d-none');
        document.body.classList.remove('page-is-loading');
    }

    async function readWorkbook(file) {
        const arrayBuffer = await file.arrayBuffer();
        await nextFrame();
        const workbook = XLSX.read(arrayBuffer, {type: 'array'});
        await nextFrame();
        const firstSheetName = workbook.SheetNames[0];
        if (!firstSheetName) return {header: [], rows: []};
        const sheet = workbook.Sheets[firstSheetName];
        const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
        if (!data.length) return {header: [], rows: []};
        const [header, ...rows] = data;
        return {header, rows};
    }

    async function handleFile(file) {
        if (!file || !/\.xlsx?$/i.test(file.name)) {
            alert('请拖入或选择 .xls / .xlsx 文件');
            return;
        }
        showLoading('正在读取文件…');
        const startTime = performance.now();
        try {
            await nextFrame();
            const {header, rows} = await readWorkbook(file);
            if (!header.length) {
                throw new Error('未读取到表头，请确认文件内容。');
            }
            const normalizedHeader = header.map(normalizeHeader);
            const headerSet = new Set(normalizedHeader);
            const datasetType = detectDatasetType(headerSet);
            if (!datasetType) {
                alert('文件格式不符合统计要求，未找到 KM_YY 或 YDDH 字段。');
                return;
            }
            const objects = await rowsToObjects(normalizedHeader, rows);
            lastDataset = {
                type: datasetType,
                rows: objects,
                headerSet,
            };
            await processCurrentDataset(startTime);
        } catch (err) {
            console.error(err);
            alert(err.message || '读取文件失败，请确认文件内容。');
        } finally {
            hideLoading();
        }
    }

    function normalizeHeader(label) {
        const raw = String(label ?? '').trim();
        if (!raw) return '';
        const base = raw.split(/[\s（(【\[]/)[0];
        if (HEADER_ALIAS.has(base)) return HEADER_ALIAS.get(base);
        const upper = base.toUpperCase();
        if (HEADER_ALIAS.has(upper)) return HEADER_ALIAS.get(upper);
        if (SUBJECT_KEY_SET.has(upper)) return upper;
        const firstPart = upper.split('_')[0];
        if (HEADER_ALIAS.has(firstPart)) return HEADER_ALIAS.get(firstPart);
        return upper;
    }

    function detectDatasetType(headerSet) {
        if (headerSet.has('KM_YY')) return 'xuexuan';
        if (headerSet.has('YDDH')) return 'gaokao';
        return null;
    }

    async function rowsToObjects(header, rows) {
        const result = [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const obj = {};
            header.forEach((key, idx) => {
                if (!key) return;
                if (!(key in obj)) obj[key] = row[idx];
            });
            result.push(obj);
            if (i % 400 === 0) await nextFrame();
        }
        return result;
    }

    async function processCurrentDataset(startTime) {
        if (!lastDataset) return;
        const startedAt = typeof startTime === 'number' ? startTime : performance.now();
        showLoading('正在统计数据…');
        await nextFrame();
        const {type} = lastDataset;
        if (type === 'xuexuan') {
            const result = await aggregateXuexuan(lastDataset);
            applyResult(result, startedAt);
        } else if (type === 'gaokao') {
            const result = await aggregateGaokao(lastDataset);
            applyResult(result, startedAt);
        }
        hideLoading();
    }

    function applyResult(result, startTime) {
        resultColumns = result.columns;
        resultRows = result.rows;
        renderTable();
        updateCopyOutput();
        els.btnToggleMode.disabled = resultRows.length === 0;
        if (currentMode === 'copy') toggleMode();
        const duration = ((performance.now() - startTime) / 1000).toFixed(3);
        updateStatus(buildStatusText(result.meta, duration));
    }

    function buildStatusText(meta, durationSec) {
        if (!meta) return `处理完成，用时 ${durationSec} 秒。`;
        const parts = [];
        if (meta.type === 'xuexuan') {
            parts.push(`学选考：保留 ${meta.kept} 条`);
            if (meta.filtered) parts.push(`过滤 ${meta.filtered} 条（缴费/审核）`);
            if (meta.missingBmd) parts.push(`缺少报名点映射 ${meta.missingBmd} 条`);
            if (meta.missingCounty) parts.push(`缺少县区映射 ${meta.missingCounty} 条`);
        } else {
            parts.push(`高考：保留 ${meta.kept} 条`);
            if (meta.filtered) parts.push(`过滤 ${meta.filtered} 条（审核未通过）`);
            if (meta.missingBmd) parts.push(`缺少报名点映射 ${meta.missingBmd} 条`);
            if (meta.missingCounty) parts.push(`缺少县区映射 ${meta.missingCounty} 条`);
        }
        parts.push(`耗时 ${durationSec} 秒`);
        return parts.join('，');
    }

    function renderTable() {
        if (!resultColumns.length || !resultRows.length) {
            renderEmptyTable();
            return;
        }
        const thead = document.createElement('tr');
        resultColumns.forEach(col => {
            const th = document.createElement('th');
            th.textContent = col.label;
            thead.appendChild(th);
        });
        els.tableHead.innerHTML = '';
        els.tableHead.appendChild(thead);

        els.tableBody.innerHTML = '';
        resultRows.forEach(row => {
            const tr = document.createElement('tr');
            if (row.rowType === 'countySummary') tr.classList.add('county-summary-row');
            if (row.rowType === 'grandTotal') tr.classList.add('grand-total-row');
            resultColumns.forEach(col => {
                const td = document.createElement('td');
                td.textContent = row[col.key] ?? '';
                tr.appendChild(td);
            });
            els.tableBody.appendChild(tr);
        });
    }

    function updateCopyOutput() {
        if (!resultColumns.length || !resultRows.length) {
            els.copyOutput.value = '';
            return;
        }
        const headerLine = resultColumns.map(col => col.label).join('\t');
        const lines = resultRows.map(row => resultColumns.map(col => row[col.key] ?? '').join('\t'));
        els.copyOutput.value = [headerLine, ...lines].join('\n');
    }

    function formatCount(count) {
        return count > 0 ? String(count) : '';
    }

    function extractCountyCode(bmdCode) {
        if (!bmdCode || bmdCode.length < 3) return '';
        return bmdCode.slice(1, 3);
    }

    function normalizeSubjectStatus(value) {
        const text = String(value ?? '').trim();
        if (!text) return '';
        if (text === '0') return '';
        if (text === '1') return '学考';
        if (text === '2') return '选考';
        if (text === '高考') return '学考'; // 视为学考场次的外语
        return text;
    }

    function normalizeShzt(value) {
        return String(value ?? '').trim();
    }

    function normalizeWyyz(value) {
        const text = String(value ?? '').trim();
        if (!text) return '未选';
        if (WYYZ_CODE_TO_LABEL[text]) return WYYZ_CODE_TO_LABEL[text];
        return text;
    }

    async function aggregateXuexuan(dataset) {
        const {rows, headerSet} = dataset;
        const hasFeeColumn = headerSet.has('缴费状态');
        const statsByCounty = new Map();
        let kept = 0;
        let filtered = 0;
        let missingBmd = 0;
        let missingCounty = 0;

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (hasFeeColumn) {
                const fee = String(row['缴费状态'] ?? '').trim();
                if (fee !== '已缴费') {
                    filtered++;
                    continue;
                }
            } else {
                const shzt = normalizeShzt(row['SHZT']);
                if (shzt !== '审核通过' && shzt !== '3') {
                    filtered++;
                    continue;
                }
            }

            const bmdCode = String(row['BMD'] ?? '').trim();
            if (!bmdCode) {
                filtered++;
                continue;
            }
            const bmdName = bmdMap.get(bmdCode) ?? '';
            if (!bmdName) missingBmd++;
            const countyCode = extractCountyCode(bmdCode);
            if (!countyCode) {
                missingCounty++;
                continue;
            }
            const countyName = countyMap.get(countyCode) ?? '';
            if (!countyName) missingCounty++;

            const countyStats = ensureCounty(statsByCounty, countyCode, countyName);
            const bmdStats = ensureBmd(countyStats, bmdCode, bmdName);
            bmdStats.total += 1;

            const selectedKeys = [];
            SUBJECTS.forEach(subj => {
                const status = normalizeSubjectStatus(row[subj.key]);
                if (status === '学考') {
                    bmdStats.xuekao[subj.key] += 1;
                } else if (status === '选考') {
                    bmdStats.xuankao[subj.key] += 1;
                    if (subj.key !== 'KM_YY') {
                        selectedKeys.push(subj.key);
                    }
                }
            });

            const lang = normalizeWyyz(row['WYYZ']);
            if (!bmdStats.wyyz[lang]) bmdStats.wyyz[lang] = 0;
            bmdStats.wyyz[lang] += 1;
            if (lang && lang !== '未选') {
                bmdStats.foreignTotal += 1;
            }
            recordComboSelection(bmdStats.comboCounts, selectedKeys);

            kept++;
            if (i % 400 === 0) await nextFrame();
        }

        const activeComboSpecs = deriveActiveComboSpecs(statsByCounty);
        const columns = buildXuexuanColumns(activeComboSpecs);
        const rowsData = buildHierarchicalRows(statsByCounty, activeComboSpecs);
        return {
            columns,
            rows: rowsData,
            meta: {type: 'xuexuan', kept, filtered, missingBmd, missingCounty},
        };
    }

    function ensureCounty(map, code, name) {
        if (!map.has(code)) {
            map.set(code, {
                code,
                name,
                bmds: new Map(),
            });
        }
        return map.get(code);
    }

    function ensureBmd(countyStats, code, name) {
        if (!countyStats.bmds.has(code)) {
            countyStats.bmds.set(code, {
                rowType: 'bmd',
                countyCode: countyStats.code,
                countyName: countyStats.name,
                bmdCode: code,
                bmdName: name,
                total: 0,
                foreignTotal: 0,
                xuekao: SUBJECTS.reduce((acc, subj) => {
                    acc[subj.key] = 0;
                    return acc;
                }, {}),
                xuankao: SUBJECTS.reduce((acc, subj) => {
                    acc[subj.key] = 0;
                    return acc;
                }, {}),
                wyyz: LANGUAGE_LABELS.reduce((acc, label) => {
                    acc[label] = 0;
                    return acc;
                }, {}),
                comboCounts: initComboCounts(),
            });
        }
        return countyStats.bmds.get(code);
    }

    function buildXuexuanColumns(activeComboSpecs) {
        const base = [
            {key: 'countyCode', label: '县区代码'},
            {key: 'countyName', label: '县区名称'},
            {key: 'bmdCode', label: '报名点代码'},
            {key: 'bmdName', label: '报名点名称'},
            {key: 'total', label: '总数'},
        ];
        const xuekao = NON_LANG_SUBJECTS.map(subj => ({
            key: `xk_${subj.key}`,
            label: `学考_${subj.label}`,
        }));
        const xuankao = NON_LANG_YS_SUBJECTS.map(subj => ({
            key: `sk_${subj.key}`,
            label: `选考_${subj.label}`,
        }));
        const langs = [];
        LANGUAGE_LABELS.forEach(label => {
            if (label === '未选') {
                langs.push({key: 'lang_TOTAL', label: '外语_总数'});
            }
            langs.push({
                key: `lang_${label}`,
                label: `外语_${label}`,
            });
        });
        const comboCols = buildComboColumns(activeComboSpecs);
        return [
            ...base,
            ...xuekao,
            ...xuankao,
            ...langs,
            ...comboCols,
        ];
    }

    function buildHierarchicalRows(statsByCounty, activeComboSpecs) {
        const rows = [];
        const countyCodes = Array.from(statsByCounty.keys()).sort();
        const grand = initAggregateRow('grandTotal');

        countyCodes.forEach(code => {
            const county = statsByCounty.get(code);
            const countySummary = initAggregateRow('countySummary', code, county.name);
            const bmdCodes = Array.from(county.bmds.keys()).sort();

            bmdCodes.forEach(bmdCode => {
                const summary = county.bmds.get(bmdCode);
                rows.push(summaryToRow(summary, activeComboSpecs));
                accumulate(countySummary, summary);
                accumulate(grand, summary);
            });

            if (countySummary.total > 0) {
                rows.push(summaryToRow(countySummary, activeComboSpecs));
            }
        });

        if (grand.total > 0) {
            rows.push(summaryToRow(grand, activeComboSpecs));
        }

        return rows;
    }

    function initAggregateRow(rowType, countyCode = '', countyName = '') {
        return {
            rowType,
            countyCode,
            countyName,
            bmdCode: rowType === 'grandTotal' ? '——' : '合计',
            bmdName: rowType === 'grandTotal' ? '总计' : '县区汇总',
            total: 0,
            foreignTotal: 0,
            xuekao: SUBJECTS.reduce((acc, subj) => {
                acc[subj.key] = 0;
                return acc;
            }, {}),
            xuankao: SUBJECTS.reduce((acc, subj) => {
                acc[subj.key] = 0;
                return acc;
            }, {}),
            wyyz: LANGUAGE_LABELS.reduce((acc, label) => {
                acc[label] = 0;
                return acc;
            }, {}),
            comboCounts: initComboCounts(),
        };
    }

    function accumulate(target, source) {
        target.total += source.total;
        SUBJECTS.forEach(subj => {
            target.xuekao[subj.key] += source.xuekao[subj.key];
            target.xuankao[subj.key] += source.xuankao[subj.key];
        });
        target.foreignTotal += source.foreignTotal;
        LANGUAGE_LABELS.forEach(label => {
            target.wyyz[label] += source.wyyz[label];
        });
        Object.keys(source.comboCounts).forEach(key => {
            target.comboCounts[key] += source.comboCounts[key];
        });
    }

    function summaryToRow(summary, activeComboSpecs = []) {
        const row = {
            rowType: summary.rowType ?? 'bmd',
            countyCode: summary.countyCode ?? '',
            countyName: summary.countyName ?? '',
            bmdCode: summary.bmdCode ?? '',
            bmdName: summary.bmdName ?? '',
            total: summary.total ? String(summary.total) : '',
        };
        NON_LANG_SUBJECTS.forEach(subj => {
            row[`xk_${subj.key}`] = formatCount(summary.xuekao[subj.key]);
            row[`sk_${subj.key}`] = formatCount(summary.xuankao[subj.key]);
        });
        row['lang_TOTAL'] = formatCount(summary.foreignTotal);
        LANGUAGE_LABELS.forEach(label => {
            row[`lang_${label}`] = formatCount(summary.wyyz[label]);
        });
        activeComboSpecs.forEach(spec => {
            row[spec.key] = formatCount(summary.comboCounts[spec.key]);
        });
        return row;
    }

    async function aggregateGaokao(dataset) {
        const {rows} = dataset;
        const statsByCounty = new Map();
        let kept = 0;
        let filtered = 0;
        let missingBmd = 0;
        let missingCounty = 0;

        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const shzt = normalizeShzt(row['SHZT']);
            if (shzt !== '审核通过' && shzt !== '3') {
                filtered++;
                continue;
            }
            const bmdCode = String(row['BMD'] ?? '').trim();
            if (!bmdCode) {
                filtered++;
                continue;
            }
            const bmdName = bmdMap.get(bmdCode) ?? '';
            if (!bmdName) missingBmd++;
            const countyCode = extractCountyCode(bmdCode);
            if (!countyCode) {
                missingCounty++;
                continue;
            }
            const countyName = countyMap.get(countyCode) ?? '';
            if (!countyName) missingCounty++;
            const countyStats = ensureCounty(statsByCounty, countyCode, countyName);
            const bmdStats = ensureGaokaoBmd(countyStats, bmdCode, bmdName);
            bmdStats.total += 1;

            const ystyskRaw = String(row['YSTYSK'] ?? '').trim();
            if (ystyskRaw) {
                const codes = Array.from(new Set(ystyskRaw.split('|').map(code => code.trim()).filter(Boolean)));
                codes.forEach(code => {
                    const label = YSTYSK_MAP[code] ?? `${code}类`;
                    if (!bmdStats.arts[label]) bmdStats.arts[label] = 0;
                    bmdStats.arts[label] += 1;
                });
            }

            kept++;
            if (i % 500 === 0) await nextFrame();
        }

        const columns = buildGaokaoColumns(statsByCounty);
        const rowsData = buildGaokaoRows(statsByCounty, columns);
        return {
            columns,
            rows: rowsData,
            meta: {type: 'gaokao', kept, filtered, missingBmd, missingCounty},
        };
    }

    function ensureGaokaoBmd(countyStats, code, name) {
        if (!countyStats.bmds.has(code)) {
            countyStats.bmds.set(code, {
                rowType: 'bmd',
                countyCode: countyStats.code,
                countyName: countyStats.name,
                bmdCode: code,
                bmdName: name,
                total: 0,
                arts: {},
            });
        }
        return countyStats.bmds.get(code);
    }

    function buildGaokaoColumns(statsByCounty) {
        const base = [
            {key: 'countyCode', label: '县区代码'},
            {key: 'countyName', label: '县区名称'},
            {key: 'bmdCode', label: '报名点代码'},
            {key: 'bmdName', label: '报名点名称'},
            {key: 'total', label: '总数'},
        ];
        const dynamicLabels = new Set(YSTYSK_LABELS);
        statsByCounty.forEach(county => {
            county.bmds.forEach(bmd => {
                Object.keys(bmd.arts).forEach(label => dynamicLabels.add(label));
            });
        });
        const extras = Array.from(dynamicLabels).map(label => ({
            key: `art_${label}`,
            label,
        }));
        return [...base, ...extras];
    }

    function buildGaokaoRows(statsByCounty, columns) {
        const rows = [];
        const countyCodes = Array.from(statsByCounty.keys()).sort();
        const grand = {
            rowType: 'grandTotal',
            countyCode: '',
            countyName: '',
            bmdCode: '——',
            bmdName: '总计',
            total: 0,
            arts: {},
        };

        countyCodes.forEach(code => {
            const county = statsByCounty.get(code);
            const countySummary = {
                rowType: 'countySummary',
                countyCode: county.code,
                countyName: county.name,
                bmdCode: '合计',
                bmdName: '县区汇总',
                total: 0,
                arts: {},
            };
            const bmdCodes = Array.from(county.bmds.keys()).sort();
            bmdCodes.forEach(bmdCode => {
                const summary = county.bmds.get(bmdCode);
                rows.push(gaokaoSummaryToRow(summary, columns));
                accumulateGaokao(countySummary, summary);
                accumulateGaokao(grand, summary);
            });
            if (countySummary.total > 0) {
                rows.push(gaokaoSummaryToRow(countySummary, columns));
            }
        });
        if (grand.total > 0) {
            rows.push(gaokaoSummaryToRow(grand, columns));
        }
        return rows;
    }

    function accumulateGaokao(target, source) {
        target.total += source.total;
        Object.keys(source.arts).forEach(label => {
            if (!target.arts[label]) target.arts[label] = 0;
            target.arts[label] += source.arts[label];
        });
    }

    function gaokaoSummaryToRow(summary, columns) {
        const row = {
            rowType: summary.rowType ?? 'bmd',
            countyCode: summary.countyCode ?? '',
            countyName: summary.countyName ?? '',
            bmdCode: summary.bmdCode ?? '',
            bmdName: summary.bmdName ?? '',
            total: summary.total ? String(summary.total) : '',
        };
        columns.forEach(col => {
            if (col.key.startsWith('art_')) {
                const label = col.label;
                row[col.key] = formatCount(summary.arts[label] ?? 0);
            }
        });
        return row;
    }

    function buildComboSpecs() {
        const combos3 = generateCombos(NON_LANG_SUBJECTS, 3);
        const combos2 = generateCombos(NON_LANG_SUBJECTS, 2);
        const combos1 = generateCombos(NON_LANG_SUBJECTS, 1);
        const comboZero = {key: 'combo_0', label: '选考_0科目'};
        const comboPartial = {key: 'combo_LT3', label: '选考_不足3科'};
        return {combos3, combos2, combos1, comboZero, comboPartial};
    }

    function generateCombos(subjects, size) {
        const results = [];
        const combo = [];
        const dfs = (start, depth) => {
            if (depth === size) {
                const keys = combo.map(item => item.key);
                results.push({
                    key: `combo_${keys.join('_')}`,
                    label: combo.map(item => item.label).join('_'),
                    size,
                    keys,
                });
                return;
            }
            for (let i = start; i <= subjects.length - (size - depth); i++) {
                combo[depth] = subjects[i];
                dfs(i + 1, depth + 1);
            }
        };
        if (size > 0) dfs(0, 0);
        return results;
    }

    function buildComboKeyMap(specs) {
        const map = new Map();
        [...specs.combos3, ...specs.combos2, ...specs.combos1].forEach(spec => {
            map.set(spec.keys.join('|'), spec);
        });
        return map;
    }

    function initComboCounts() {
        const obj = {};
        [...COMBO_SPECS.combos3, ...COMBO_SPECS.combos2, ...COMBO_SPECS.combos1].forEach(spec => {
            obj[spec.key] = 0;
        });
        obj[COMBO_SPECS.comboZero.key] = 0;
        obj[COMBO_SPECS.comboPartial.key] = 0;
        return obj;
    }

    function recordComboSelection(comboCounts, selectedKeys) {
        const count = selectedKeys.length;
        if (count === 0) {
            comboCounts[COMBO_SPECS.comboZero.key] += 1;
            return;
        }
        if (count > 3) {
            comboCounts[COMBO_SPECS.comboPartial.key] += 1;
            return;
        }
        const key = selectedKeys.join('|');
        const spec = COMBO_KEY_MAP.get(key);
        if (spec) {
            comboCounts[spec.key] += 1;
        }
        if (count < 3) {
            comboCounts[COMBO_SPECS.comboPartial.key] += 1;
        }
    }

    function deriveActiveComboSpecs(statsByCounty) {
        const usage = initComboCounts();
        statsByCounty.forEach(county => {
            county.bmds.forEach(bmd => {
                Object.keys(bmd.comboCounts).forEach(key => {
                    if (bmd.comboCounts[key] > 0) {
                        usage[key] = 1;
                    }
                });
            });
        });
        const specs = [];
        COMBO_SPECS.combos3.forEach(spec => {
            if (usage[spec.key]) specs.push(spec);
        });
        COMBO_SPECS.combos2.forEach(spec => {
            if (usage[spec.key]) specs.push(spec);
        });
        COMBO_SPECS.combos1.forEach(spec => {
            if (usage[spec.key]) specs.push(spec);
        });
        if (usage[COMBO_SPECS.comboZero.key]) specs.push(COMBO_SPECS.comboZero);
        if (usage[COMBO_SPECS.comboPartial.key]) specs.push(COMBO_SPECS.comboPartial);
        return specs;
    }

    function buildComboColumns(activeComboSpecs = []) {
        return activeComboSpecs.map(spec => ({key: spec.key, label: spec.label}));
    }
})();

