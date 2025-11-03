(function () {
	const REQUIRED_HEADERS = ["BMXH", "XM", "KM_YW"];
	const SUBJECT_KEYS = [
		["KM_YW", "语文"],
		["KM_SX", "数学"],
		["KM_YY", "英语"],
		["KM_SZ", "政治"],
		["KM_LS", "历史"],
		["KM_DL", "地理"],
		["KM_WL", "物理"],
		["KM_HX", "化学"],
		["KM_SW", "生物"],
		["KM_JS", "技术"],
	];

	const WYYZ_MAP = {
		0: "未报名",
		1: "英语",
		2: "俄语",
		3: "日语",
		4: "德语",
		5: "法语",
		6: "西班牙语",
	};

	const els = {
		year: document.getElementById("yearSelect"),
		month: document.getElementById("monthSelect"),
		day: document.getElementById("daySelect"),
		dz: document.getElementById("dropzone"),
		fileInput: document.getElementById("fileInput"),
		btnPick: document.getElementById("btnPick"),
		btnClear: document.getElementById("btnClear"),
		btnExportAll: document.getElementById("btnExportAll"),
		btnExportZip: document.getElementById("btnExportZip"),
		tbody: document.querySelector("#resultTable tbody"),
		loading: document.getElementById("loadingOverlay"),
	};

	let processedRows = [];
	let sourceRows = [];

	function initDateSelectors() {
		const now = new Date();
		const month = now.getMonth() + 1;
		const baseYear = month >= 9 ? now.getFullYear() + 1 : now.getFullYear();

		els.year.innerHTML = '';
		for (let y = baseYear - 2; y <= baseYear + 2; y++) {
			const opt = document.createElement('option');
			opt.value = String(y);
			opt.textContent = String(y) + "年";
			if (y === baseYear) opt.selected = true;
			els.year.appendChild(opt);
		}

		els.month.innerHTML = '';
		for (let m = 1; m <= 12; m++) {
			const opt = document.createElement('option');
			opt.value = String(m);
			opt.textContent = String(m) + "月";
			if (m === 6) opt.selected = true; // 默认6月
			els.month.appendChild(opt);
		}

		els.day.innerHTML = '';
		for (let d = 1; d <= 31; d++) {
			const opt = document.createElement('option');
			opt.value = String(d);
			opt.textContent = String(d) + "日";
			if (d === 30) opt.selected = true; // 默认30日
			els.day.appendChild(opt);
		}
	}

	function getTargetDate() {
		const y = parseInt(els.year.value, 10);
		const m = parseInt(els.month.value, 10);
		const d = parseInt(els.day.value, 10);
		return new Date(y, m - 1, d);
	}

	function fmt2(n) {
		return (Math.round(n * 100) / 100).toFixed(2);
	}

	function parseWyyz(val) {
		if (val === null || val === undefined || val === "") return WYYZ_MAP[0];
		const n = Number(val);
		if (!Number.isFinite(n)) return String(val);
		return WYYZ_MAP[n] ?? String(val);
	}

	function getBirthDateFromId(sfzh) {
		if (!sfzh || typeof sfzh !== 'string') return null;
		const s = sfzh.trim();
		if (s.length === 18) {
			const y = parseInt(s.substring(6, 10), 10);
			const m = parseInt(s.substring(10, 12), 10);
			const d = parseInt(s.substring(12, 14), 10);
			if (y > 1900 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
				return new Date(y, m - 1, d);
			}
		}
		return null;
	}

	function diffYears(birth, target) {
		if (!birth || !target) return null;
		const msPerDay = 24 * 60 * 60 * 1000;
		const diffDays = Math.floor((target - birth) / msPerDay);
		return diffDays / 365;
	}

	function readWorkbook(arrayBuffer) {
		const wb = XLSX.read(arrayBuffer, { type: 'array' });
		const first = wb.SheetNames[0];
		const ws = wb.Sheets[first];
		const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
		if (!data || data.length === 0) return { header: [], rows: [] };
		const header = data[0].map(h => String(h).trim());
		const rows = data.slice(1).map(r => r.map(c => c));
		return { header, rows };
	}

	function headerCheck(header) {
		const hs = new Set(header);
		return REQUIRED_HEADERS.every(h => hs.has(h));
	}

	function rowsToObjects(header, rows) {
		return rows.map(r => {
			const obj = {};
			header.forEach((h, i) => { obj[h] = r[i] ?? ''; });
			return obj;
		});
	}

	function validateRow(row, targetDate) {
		const tips = [];
		const grade = String(row['NJ'] ?? '').trim();
		const shzt = String(row['SHZT'] ?? '').trim();
		if (shzt === '审核未通过') {
			return { skip: true };
		}

		if (grade === '一') {
			tips.push('请注意高一学生报名考试');
		} else if (grade === '二') {
			// 二年级：政治、物理应为“学考”，其他应为空白
			const mustXuekao = { 'KM_SZ': '政治', 'KM_WL': '物理' };
			Object.keys(mustXuekao).forEach(k => {
				if (String(row[k] ?? '').trim() !== '学考') {
					tips.push(`${mustXuekao[k]}应为学考`);
				}
			});
			SUBJECT_KEYS.forEach(([k, name]) => {
				if (!(k in mustXuekao)) {
					if (String(row[k] ?? '').trim() !== '') {
						tips.push(`${name}应为空白`);
					}
				}
			});
		} else if (grade === '三') {
			let electiveCount = 0;
			SUBJECT_KEYS.forEach(([k]) => {
				const v = String(row[k] ?? '').trim();
				if (v === '选考') electiveCount += 1;
			});
			if (electiveCount !== 3) {
				tips.push('选考科目少于3门');
			}
			if (String(row['KM_YY'] ?? '').trim() !== '高考') {
				tips.push('未报名外语科目');
			}
		}

		// 年龄判断（学籍辅号为空时）
		const xjfh = String(row['XJFH'] ?? '').trim();
		if (!xjfh) {
			const birth = getBirthDateFromId(String(row['SFZH'] ?? '').trim());
			const years = diffYears(birth, targetDate);
			if (years !== null && years < 18) {
				tips.push(`报名年龄为${fmt2(years)}岁`);
			}
		}

		return { skip: false, tips };
	}

	function buildDisplayRow(row, tips) {
		const electiveNames = [];
		const xuekaoNames = [];
		let countXuan = 0;
		let countXue = 0;
		SUBJECT_KEYS.forEach(([k, name]) => {
			const v = String(row[k] ?? '').trim();
			if (v === '选考') {
				if (k !== 'KM_YY') electiveNames.push(name); // 选考科目中去除英语
				countXuan += 1;
			}
			if (v === '学考') {
				xuekaoNames.push(name);
				countXue += 1;
			}
		});
		const foreignLang = parseWyyz(row['WYYZ']);
		return {
			BMXH: String(row['BMXH'] ?? '').trim(),
			SFZH: String(row['SFZH'] ?? '').trim(),
			XM: String(row['XM'] ?? '').trim(),
			XJFH: String(row['XJFH'] ?? '').trim(),
			BMD: String(row['BMD'] ?? '').trim(),
			NJ: String(row['NJ'] ?? '').trim(),
			BJ: String(row['BJ'] ?? '').trim(),
			ELECTIVE_SUBJECTS: electiveNames.join('、'),
			COUNT_XK: countXuan,
			XUEKAO_SUBJECTS: xuekaoNames.join('、'),
			COUNT_XKX: countXue,
			WYYZ_LABEL: foreignLang,
			TIPS: tips.join('；')
		};
	}

	function renderTable(rows) {
		els.tbody.innerHTML = '';
		rows.forEach((r, idx) => {
			const tr = document.createElement('tr');
			const cells = [
				idx + 1,
				r.BMXH, r.SFZH, r.XM, r.XJFH, r.BMD, r.NJ, r.BJ,
				r.ELECTIVE_SUBJECTS, r.COUNT_XK, r.XUEKAO_SUBJECTS, r.COUNT_XKX, r.WYYZ_LABEL, r.TIPS
			];
			cells.forEach(c => {
				const td = document.createElement('td');
				td.textContent = c == null ? '' : String(c);
				tr.appendChild(td);
			});
			els.tbody.appendChild(tr);
		});
		const hasData = rows.length > 0;
		els.btnExportAll.disabled = !hasData;
		els.btnExportZip.disabled = !hasData;
	}

	function enablePageDnd() {
		['dragenter','dragover','dragleave','drop'].forEach(evt => {
			document.addEventListener(evt, e => {
				e.preventDefault();
				e.stopPropagation();
			});
		});
		document.addEventListener('dragover', () => { els.dz.classList.add('dragging'); });
		document.addEventListener('dragleave', () => { els.dz.classList.remove('dragging'); });
		document.addEventListener('drop', (e) => {
			els.dz.classList.remove('dragging');
			const files = e.dataTransfer && e.dataTransfer.files ? e.dataTransfer.files : [];
			if (files.length) handleFile(files[0]);
		});
	}

function showLoading() { if (els.loading) els.loading.classList.remove('d-none'); }
function hideLoading() { if (els.loading) els.loading.classList.add('d-none'); }
function nextTick() { return new Promise(r => setTimeout(r, 0)); }

async function handleFile(file) {
		if (!file || !file.name.endsWith('.xlsx')) {
			alert('请拖入 .xlsx 文件');
			return;
		}
		showLoading();
		try {
			await nextTick();
			const buf = await file.arrayBuffer();
			await nextTick();
			const { header, rows } = readWorkbook(buf);
			if (!headerCheck(header)) {
				alert('文件格式不正确，缺少必要字段（BMXH、XM、KM_YW）');
				return;
			}
			await nextTick();
			const objects = rowsToObjects(header, rows);
			// 保留可参与计算的原始数据（过滤审核未通过）
			sourceRows = objects.filter(r => String(r['SHZT'] ?? '').trim() !== '审核未通过');
			await rebuildProcessedAsync();
		} finally {
			hideLoading();
		}
	}

async function rebuildProcessedAsync() {
	const targetDate = getTargetDate();
	processedRows = [];
	for (let i = 0; i < sourceRows.length; i++) {
		const row = sourceRows[i];
		const { skip, tips } = validateRow(row, targetDate);
		if (skip) continue;
		if (!tips || tips.length === 0) continue; // 仅保留有提示信息的
		processedRows.push(buildDisplayRow(row, tips));
		if (i % 500 === 0) await nextTick();
	}
	renderTable(processedRows);
}

	function exportAll() {
		if (!processedRows.length) return;
		const wb = XLSX.utils.book_new();
		const sheetData = processedRows.map(r => ({
			"报名序号": r.BMXH,
			"身份证号": r.SFZH,
			"姓名": r.XM,
			"学籍辅号": r.XJFH,
			"报名点": r.BMD,
			"年级": r.NJ,
			"班级": r.BJ,
			"报考的选考科目": r.ELECTIVE_SUBJECTS,
			"报考选考数量": r.COUNT_XK,
			"报考的学考科目": r.XUEKAO_SUBJECTS,
			"报考学考数量": r.COUNT_XKX,
			"报考的外语科目": r.WYYZ_LABEL,
			"提示": r.TIPS,
		}));
		const ws = XLSX.utils.json_to_sheet(sheetData);
		XLSX.utils.book_append_sheet(wb, ws, '校验结果');
		XLSX.writeFile(wb, '选考报名数据校验-全部.xlsx');
	}

	async function exportZipByBmd() {
		if (!processedRows.length) return;
		showLoading();
		const byBmd = new Map();
		for (const r of processedRows) {
			const key = r.BMD || '未知报名点';
			if (!byBmd.has(key)) byBmd.set(key, []);
			byBmd.get(key).push(r);
		}
		const zip = new JSZip();
		for (const [bmd, arr] of byBmd.entries()) {
			const wb = XLSX.utils.book_new();
			const sheetData = arr.map(r => ({
				"报名序号": r.BMXH,
				"身份证号": r.SFZH,
				"姓名": r.XM,
				"学籍辅号": r.XJFH,
				"报名点": r.BMD,
				"年级": r.NJ,
				"班级": r.BJ,
				"报考的选考科目": r.ELECTIVE_SUBJECTS,
				"报考选考数量": r.COUNT_XK,
				"报考的学考科目": r.XUEKAO_SUBJECTS,
				"报考学考数量": r.COUNT_XKX,
				"报考的外语科目": r.WYYZ_LABEL,
				"提示": r.TIPS,
			}));
			const ws = XLSX.utils.json_to_sheet(sheetData);
			XLSX.utils.book_append_sheet(wb, ws, '校验结果');
			const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
			const filename = `选考报名数据校验-${bmd}.xlsx`;
			zip.file(filename, wbout);
		}
		const blob = await zip.generateAsync({ type: 'blob' });
		downloadBlob(blob, '按报名点导出.zip');
		hideLoading();
	}

	function downloadBlob(blob, filename) {
		const url = URL.createObjectURL(blob);
		const a = document.createElement('a');
		a.href = url;
		a.download = filename;
		document.body.appendChild(a);
		a.click();
		document.body.removeChild(a);
		URL.revokeObjectURL(url);
	}

	function wireEvents() {
		enablePageDnd();
		els.btnPick.addEventListener('click', () => els.fileInput.click());
		els.fileInput.addEventListener('change', (e) => {
			const f = e.target.files && e.target.files[0];
			if (f) handleFile(f);
		});
		[els.year, els.month, els.day].forEach(sel => {
			sel.addEventListener('change', () => {
				if (sourceRows.length) {
					showLoading();
					rebuildProcessedAsync().finally(hideLoading);
				}
			});
		});
		els.btnClear.addEventListener('click', () => { processedRows = []; renderTable(processedRows); });
		els.btnExportAll.addEventListener('click', exportAll);
		els.btnExportZip.addEventListener('click', exportZipByBmd);
	}

	// init
	initDateSelectors();
	wireEvents();
})();


