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
		bmdList: document.getElementById("bmdList"),
		btnFilterAll: document.getElementById("btnFilterAll"),
		btnFilterInverse: document.getElementById("btnFilterInverse"),
		btnFilterConfirm: document.getElementById("btnFilterConfirm"),
		ageCheck: document.getElementById("ageCheckEnabled"),
		btnXjfhFilterConfirm: document.getElementById("btnXjfhFilterConfirm"),
		xjfhHasCount: document.getElementById("xjfh_has_count"),
		xjfhNoneCount: document.getElementById("xjfh_none_count"),
		xjfhHasCheck: document.getElementById("xjfh_has"),
		xjfhNoneCheck: document.getElementById("xjfh_none"),
	};

	let processedRows = [];
	let sourceRows = [];
	let selectedBmds = new Set();
	let allBmds = [];
	let bmdCounts = {};
	let pendingSelectedBmds = new Set();
	let selectedXjfhStatus = new Set(['has', 'none']);
	let pendingSelectedXjfhStatus = new Set(['has', 'none']);
	let xjfhCounts = { has: 0, none: 0 };

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
		const bmxh = String(row['BMXH'] ?? '').trim();
		const grade = String(row['NJ'] ?? '').trim();
		const shzt = String(row['SHZT'] ?? '').trim();
		if (shzt === '审核未通过') {
			return { skip: true };
		}

		if (bmxh.startsWith('A')) {
			// A开头逻辑
			if (grade === '一') {
				// 高一年级正常需要同时选择学考的（历史、地理、化学、生物）
				const mustXuekao = { 'KM_LS': '历史', 'KM_DL': '地理', 'KM_HX': '化学', 'KM_SW': '生物' };
				Object.keys(mustXuekao).forEach(k => {
					if (String(row[k] ?? '').trim() !== '学考') {
						tips.push(`未报名${mustXuekao[k]}学考`);
					}
				});
			} else if (grade === '二') {
				// A开头报名序号，均不作选考的校验，只学考的校验
			}
			// wyyz不做校验 (A开头不校验)
			// 高三和99不做科目校验 (默认不加tips即可)
		} else {
			// 原逻辑（X开头或其他）
			if (grade === '一') {
				tips.push('请注意高一学生报名考试');
			} else if (grade === '二') {
				// 二年级：政治、物理未报名学考，其他不应报名
				const mustXuekao = { 'KM_SZ': '政治', 'KM_WL': '物理' };
				Object.keys(mustXuekao).forEach(k => {
					if (String(row[k] ?? '').trim() !== '学考') {
						tips.push(`未报名${mustXuekao[k]}学考`);
					}
				});
				SUBJECT_KEYS.forEach(([k, name]) => {
					if (!(k in mustXuekao)) {
						if (String(row[k] ?? '').trim() !== '') {
							tips.push(`${name}不应报名考试`);
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
		}

		// 年龄判断（学籍辅号为空时）
		const isAgeCheckEnabled = els.ageCheck ? els.ageCheck.checked : true;
		if (isAgeCheckEnabled) {
			const xjfh = String(row['XJFH'] ?? '').trim();
			if (!xjfh) {
				const birth = getBirthDateFromId(String(row['SFZH'] ?? '').trim());
				const years = diffYears(birth, targetDate);
				if (years !== null && years < 18) {
					tips.push(`报名年龄为${fmt2(years)}岁`);
				}
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
			cells.forEach((c, i) => {
				const td = document.createElement('td');
				td.textContent = c == null ? '' : String(c);
				// 如果是提示列（最后一列），添加红色加粗样式
				if (i === cells.length - 1) {
					td.classList.add('text-danger', 'fw-bold');
				}
				tr.appendChild(td);
			});
			els.tbody.appendChild(tr);
		});
		const hasData = sourceRows.length > 0;
		els.btnExportAll.disabled = !hasData;
		els.btnExportZip.disabled = !hasData;

		// 只有在存在多个报名点时才显示导出ZIP按钮
		if (allBmds.length > 1) {
			els.btnExportZip.classList.remove('d-none');
		} else {
			els.btnExportZip.classList.add('d-none');
		}
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
		// 稍微延迟一下，确保 Loading 显示出来后再进行耗时操作
		await new Promise(r => setTimeout(r, 50));
		try {
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
			
			// 初始化报名点筛选
			const bmdSet = new Set();
			sourceRows.forEach(r => {
				const bmd = String(r['BMD'] ?? '').trim() || '未知报名点';
				bmdSet.add(bmd);
			});
			allBmds = Array.from(bmdSet).sort();
			selectedBmds = new Set(allBmds);
			pendingSelectedBmds = new Set(selectedBmds);
			selectedXjfhStatus = new Set(['has', 'none']);
			pendingSelectedXjfhStatus = new Set(['has', 'none']);
			updateBmdFilterUI();
			updateXjfhFilterUI();

			await rebuildProcessedAsync();
		} finally {
			hideLoading();
		}
	}

	async function rebuildProcessedAsync() {
		const targetDate = getTargetDate();
		processedRows = [];
		const activeBmdSet = new Set();
		const counts = {};
		const xCounts = { has: 0, none: 0 };
		for (let i = 0; i < sourceRows.length; i++) {
			const row = sourceRows[i];
			const { skip, tips } = validateRow(row, targetDate);
			if (skip) continue;
			if (!tips || tips.length === 0) continue; // 仅保留有提示信息的
			const disp = buildDisplayRow(row, tips);
			processedRows.push(disp);
			const bmdKey = disp.BMD || '未知报名点';
			activeBmdSet.add(bmdKey);
			counts[bmdKey] = (counts[bmdKey] || 0) + 1;

			if (disp.XJFH) {
				xCounts.has += 1;
			} else {
				xCounts.none += 1;
			}

			if (i % 500 === 0) await nextTick();
		}
		
		// 更新报名点列表，只显示有数据的
		allBmds = Array.from(activeBmdSet).sort();
		bmdCounts = counts;
		xjfhCounts = xCounts;
		// 如果是第一次加载或者之前的选择不再有效，重置选择
		if (selectedBmds.size === 0) {
			selectedBmds = new Set(allBmds);
		} else {
			// 保留仍然存在的选择
			const next = new Set();
			selectedBmds.forEach(bmd => {
				if (activeBmdSet.has(bmd)) next.add(bmd);
			});
			selectedBmds = next;
		}
		pendingSelectedBmds = new Set(selectedBmds);
		updateBmdFilterUI();
		updateXjfhFilterUI();
		
		applyFilterAndRender();
	}

function applyFilterAndRender() {
	const filtered = processedRows.filter(r => {
		const bmdMatch = selectedBmds.has(r.BMD || '未知报名点');
		const xjfhType = r.XJFH ? 'has' : 'none';
		const xjfhMatch = selectedXjfhStatus.has(xjfhType);
		return bmdMatch && xjfhMatch;
	});
	renderTable(filtered);
}

function updateXjfhFilterUI() {
	if (els.xjfhHasCount) els.xjfhHasCount.textContent = xjfhCounts.has;
	if (els.xjfhNoneCount) els.xjfhNoneCount.textContent = xjfhCounts.none;
	if (els.xjfhHasCheck) els.xjfhHasCheck.checked = pendingSelectedXjfhStatus.has('has');
	if (els.xjfhNoneCheck) els.xjfhNoneCheck.checked = pendingSelectedXjfhStatus.has('none');
}

function updateBmdFilterUI() {
	if (!els.bmdList) return;
	els.bmdList.innerHTML = '';
	allBmds.forEach(bmd => {
		const div = document.createElement('div');
		div.className = 'form-check bmd-filter-item';
		const id = `bmd_${Math.random().toString(36).substr(2, 9)}`;
		const checked = pendingSelectedBmds.has(bmd);
		const count = bmdCounts[bmd] || 0;
		div.innerHTML = `
			<input class="form-check-input bmd-checkbox" type="checkbox" value="${bmd}" id="${id}" ${checked ? 'checked' : ''}>
			<label class="form-check-label" for="${id}">${bmd}(${count})</label>
		`;
		els.bmdList.appendChild(div);
	});
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
		[els.year, els.month, els.day, els.ageCheck].forEach(sel => {
			if (!sel) return;
			sel.addEventListener('change', () => {
				if (sourceRows.length) {
					showLoading();
					rebuildProcessedAsync().finally(hideLoading);
				}
			});
		});
		els.btnClear.addEventListener('click', () => { 
			processedRows = []; 
			sourceRows = [];
			allBmds = [];
			bmdCounts = {};
			xjfhCounts = { has: 0, none: 0 };
			selectedBmds = new Set();
			selectedXjfhStatus = new Set(['has', 'none']);
			updateBmdFilterUI();
			updateXjfhFilterUI();
			renderTable(processedRows); 
		});
		els.btnExportAll.addEventListener('click', exportAll);
		els.btnExportZip.addEventListener('click', exportZipByBmd);

		// 筛选相关的事件
		if (els.bmdList) {
			els.bmdList.addEventListener('change', (e) => {
				if (e.target.classList.contains('bmd-checkbox')) {
					const val = e.target.value;
					if (e.target.checked) {
						pendingSelectedBmds.add(val);
					} else {
						pendingSelectedBmds.delete(val);
					}
				}
			});
		}

		if (els.btnFilterAll) {
			els.btnFilterAll.addEventListener('click', (e) => {
				e.stopPropagation();
				pendingSelectedBmds = new Set(allBmds);
				updateBmdFilterUI();
			});
		}

		if (els.btnFilterInverse) {
			els.btnFilterInverse.addEventListener('click', (e) => {
				e.stopPropagation();
				const next = new Set();
				allBmds.forEach(bmd => {
					if (!pendingSelectedBmds.has(bmd)) next.add(bmd);
				});
				pendingSelectedBmds = next;
				updateBmdFilterUI();
			});
		}

		if (els.btnFilterConfirm) {
			els.btnFilterConfirm.addEventListener('click', () => {
				selectedBmds = new Set(pendingSelectedBmds);
				applyFilterAndRender();
				// 自动关闭下拉菜单 (Bootstrap 5 方式)
				const dropdownElement = document.getElementById("bmdFilterMenu").previousElementSibling;
				if (dropdownElement) {
					const dropdown = bootstrap.Dropdown.getInstance(dropdownElement);
					if (dropdown) dropdown.hide();
				}
			});
		}

		if (els.btnXjfhFilterConfirm) {
			els.btnXjfhFilterConfirm.addEventListener('click', () => {
				selectedXjfhStatus = new Set(pendingSelectedXjfhStatus);
				applyFilterAndRender();
				const dropdownElement = document.getElementById("xjfhFilterMenu").previousElementSibling;
				if (dropdownElement) {
					const dropdown = bootstrap.Dropdown.getInstance(dropdownElement);
					if (dropdown) dropdown.hide();
				}
			});
		}

		// 处理学籍辅号勾选变动
		const xjfhFilterMenu = document.getElementById('xjfhFilterMenu');
		if (xjfhFilterMenu) {
			xjfhFilterMenu.addEventListener('change', (e) => {
				if (e.target.classList.contains('xjfh-checkbox')) {
					const val = e.target.value;
					if (e.target.checked) {
						pendingSelectedXjfhStatus.add(val);
					} else {
						pendingSelectedXjfhStatus.delete(val);
					}
				}
			});
			xjfhFilterMenu.addEventListener('click', (e) => {
				e.stopPropagation();
			});
		}

		// 阻止下拉菜单点击自动关闭
		const filterMenu = document.getElementById('bmdFilterMenu');
		if (filterMenu) {
			filterMenu.addEventListener('click', (e) => {
				e.stopPropagation();
			});
		}
	}

	// init
	initDateSelectors();
	wireEvents();
})();


