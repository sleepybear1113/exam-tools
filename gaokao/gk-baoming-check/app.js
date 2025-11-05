(function () {
    const els = {
        dz: document.getElementById('dropzone'),
        fileInput: document.getElementById('fileInput'),
        btnPick: document.getElementById('btnPick'),
        btnClear: document.getElementById('btnClear'),
        btnExportAll: document.getElementById('btnExportAll'),
        chkZip: document.getElementById('chkZip'),
        groupLenSelect: document.getElementById('groupLenSelect'),
        btnExportZip: document.getElementById('btnExportZip'),
        stats: document.getElementById('statsContainer'),
        btnExportStats: document.getElementById('btnExportStats'),
        warnTbody: document.querySelector('#warnTable tbody'),
        warnCount: document.getElementById('warningCount'),
        loading: document.getElementById('loadingOverlay'),
    };

    let records = []; // 原始解析后记录（每页一条）
    let warnings = []; // {ksHao, bj, xm, tips[]}
    let stats = null; // 统计数据

    // 省级行政区划关键字（简写，匹配开头两个字）
    const PROVINCE_PREFIXES = [
        "北京", "天津", "上海", "重庆", "河北", "山西", "内蒙", "辽宁", "吉林", "黑龙", "江苏", "浙江",
        "安徽", "福建", "江西", "山东", "河南", "湖北", "湖南", "广东", "广西", "海南", "四川", "贵州",
        "云南", "西藏", "陕西", "甘肃", "青海", "宁夏", "新疆", "香港", "澳门", "台湾"
    ];

    const FIELD_LABELS = [
        "考生号", "姓名", "性别", "民族", "毕业学校", "外语语种", "报考科类", "职业类别", "政治面貌",
        "毕业类别", "思想品德考核结果", "体育达标结果", "考生类别", "报名类别", "邮寄地址", "证件号码",
        "报名点", "参加高校招生英语面试"
    ];

    function showLoading() {
        if (els.loading) els.loading.classList.remove('d-none');
    }

    function hideLoading() {
        if (els.loading) els.loading.classList.add('d-none');
    }

    function nextTick() {
        return new Promise(r => setTimeout(r, 0));
    }

    function enablePageDnd() {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => {
            document.addEventListener(evt, e => {
                e.preventDefault();
                e.stopPropagation();
            });
        });
        document.addEventListener('dragover', () => {
            els.dz.classList.add('dragging');
        });
        document.addEventListener('dragleave', () => {
            els.dz.classList.remove('dragging');
        });
        document.addEventListener('drop', e => {
            els.dz.classList.remove('dragging');
            const files = e.dataTransfer && e.dataTransfer.files ? e.dataTransfer.files : [];
            if (files.length) handlePdf(files[0]);
        });
    }

    function setControlsEnabled(enabled) {
        els.btnClear.disabled = !enabled;
        els.btnExportAll.disabled = !enabled;
        if (els.btnExportStats) els.btnExportStats.disabled = !enabled;
        els.chkZip.disabled = !enabled;
        els.groupLenSelect.disabled = !enabled || !els.chkZip.checked;
        els.btnExportZip.disabled = !enabled || !els.chkZip.checked;
    }

    function normalizeSpaces(s) {
        return s.replace(/\s+/g, ' ').trim();
    }

    function parsePageTextToRecord(text) {
        const rec = {};
        const normalized = normalizeSpaces(text);
        const tokens = normalized.split(' ');
        for (let i = 0; i < tokens.length; i++) {
            const t = tokens[i];
            if (i === 1) {
                rec["报名类别"] = t;
                continue;
            }
            if (FIELD_LABELS.includes(t)) {
                const next = tokens[i + 1] || '';
                rec[t] = next;
                // 特殊规则：班级紧跟在证件号码的值之后
                if (t === '证件号码') {
                    const maybeClass = tokens[i + 2] || '';
                    if (maybeClass) rec['班级'] = maybeClass;
                    i += 2; // 跳过证件号码的值与班级
                } else {
                    i += 1;
                }
            }
        }
        return rec;
    }

    async function extractPdfText(file) {
        const pdfjsLib = window.pdfjsLib;
        if (!pdfjsLib) throw new Error('pdf.js 未加载');
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
        const pages = [];
        for (let i = 0; i < pdf.numPages; i++) {
            const page = await pdf.getPage(i + 1);
            const content = await page.getTextContent();
            const text = content.items.map(it => it.str).join(' ');
            pages.push(text);
        }
        return pages;
    }

    function validateHeader(pages) {
        if (!pages.length) return '未能读取PDF内容';
        const first = normalizeSpaces(pages[0]);
        const ok = /^20\d{2}年浙江省普通高校/.test(first);
        if (!ok) return '上传的PDF格式不正确';
        return '';
    }

    function buildStats(recs) {
        const counters = (key) => {
            const map = new Map();
            recs.forEach(r => {
                const v = String(r[key] || '').trim();
                if (!v) return;
                map.set(v, (map.get(v) || 0) + 1);
            });
            return map;
        };
        const by = {
            毕业学校: counters('毕业学校'),
            民族: counters('民族'),
            班级: counters('班级'),
            外语语种: counters('外语语种'),
            报考科类: counters('报考科类'),
            职业类别: counters('职业类别'),
            政治面貌: counters('政治面貌'),
            毕业类别: counters('毕业类别'),
            思想品德考核结果: counters('思想品德考核结果'),
            体育达标结果: counters('体育达标结果'),
            考生类别: counters('考生类别'),
        };
        // 每个班级下报考科类数量
        const classToKelei = new Map();
        recs.forEach(r => {
            const bj = String(r['班级'] || '').trim() || '未知班级';
            const kl = String(r['报考科类'] || '').trim() || '未知科类';
            if (!classToKelei.has(bj)) classToKelei.set(bj, new Map());
            const m = classToKelei.get(bj);
            m.set(kl, (m.get(kl) || 0) + 1);
        });

        // 报名点-每班-报考科类数量
        const bmdKeyToClassToKelei = new Map();
        recs.forEach(r => {
            const ksHao = String(r['考生号'] || '').trim();
            const bmdName = String(r['报名点'] || '').trim();
            const {displayKey} = deriveBmdKeyFromKsh(ksHao, bmdName);
            const bj = String(r['班级'] || '').trim() || '未知班级';
            const kl = String(r['报考科类'] || '').trim() || '未知科类';
            if (!bmdKeyToClassToKelei.has(displayKey)) bmdKeyToClassToKelei.set(displayKey, new Map());
            const classMap = bmdKeyToClassToKelei.get(displayKey);
            if (!classMap.has(bj)) classMap.set(bj, new Map());
            const m = classMap.get(bj);
            m.set(kl, (m.get(kl) || 0) + 1);
        });

        return {by, classToKelei, bmdKeyToClassToKelei};
    }

    function deriveBmdKeyFromKsh(ksHao, bmdName) {
        const s = String(ksHao || '');
        let base = '';
        let n910 = '';
        if (s.length >= 10) {
            base = s.substring(4, 8); // 第5-8位
            n910 = s.substring(8, 10); // 第9-10位
        }
        const n910Num = parseInt(n910, 10);
        let prefix = '7';
        if (Number.isFinite(n910Num)) {
            if (n910Num === 13 || n910Num === 14 || n910Num === 15) prefix = '9';
            else if (n910Num < 13) prefix = '8';
            else prefix = '7';
        }
        const fullCode = (prefix + base) || '';
        const displayKey = fullCode ? (bmdName ? `${fullCode}-${bmdName}` : fullCode) : (bmdName || '未知报名点');
        return {fullCode, displayKey};
    }

    function sexFromId(sfzh) {
        const s = String(sfzh || '').trim();
        if (s.length !== 18) return '';
        const n17 = parseInt(s.substring(16, 17), 10);
        if (!Number.isFinite(n17)) return '';
        return (n17 % 2 === 1) ? '男' : '女';
    }

    function validateRecord(r) {
        const tips = [];
        // 邮寄地址省市
        const addr = String(r['邮寄地址'] || '').trim();
        if (addr) {
            const prefix2 = addr.substring(0, 2);
            if (!PROVINCE_PREFIXES.some(p => p.startsWith(prefix2))) {
                tips.push('邮寄地址格式需有省市');
            }
        }
        // 思想品德/体育
        if (String(r['思想品德考核结果'] || '').trim() === '不合格') tips.push('思想品德考核结果为不合格');
        if (String(r['体育达标结果'] || '').trim() === '不达标') tips.push('体育达标结果不达标');
        if (String(r['民族'] || '').trim() === '其他') tips.push('民族为其他');
        // 英语面试：当且仅当有外语语种时检查
        const hasWyyz = !!String(r['外语语种'] || '').trim();
        const hasInterviewField = Object.prototype.hasOwnProperty.call(r, '参加高校招生英语面试') && !!String(r['参加高校招生英语面试'] || '').trim();
        if (hasWyyz && !hasInterviewField) tips.push('未报名英语面试');
        // 证件号码与性别
        const idSex = sexFromId(r['证件号码']);
        const sex = String(r['性别'] || '').trim();
        if (idSex && sex && idSex !== sex) tips.push('性别与证件号码不符');
        // 报名类别与考生类别
        const bmlb = String(r['报名类别'] || '').trim();
        const kslb = String(r['考生类别'] || '').trim();
        const conflict1 = (bmlb.includes('往届生和应届非新课改')) && (kslb.includes('城市应届') || kslb.includes('农村应届'));
        const conflict2 = (bmlb.includes('应届新课改')) && (kslb.includes('城市往届') || kslb.includes('农村往届'));
        if (conflict1 || conflict2) tips.push('报名类别与考生类别不符');
        return tips;
    }

    function renderStats(s) {
        const container = els.stats;
        container.innerHTML = '';
        const addSection = (title, map) => {
            const div = document.createElement('div');
            div.className = 'mb-2';
            const h = document.createElement('div');
            h.className = 'fw-bold mb-1';
            h.textContent = title;
            div.appendChild(h);
            const ul = document.createElement('ul');
            ul.className = 'mb-0';
            Array.from(map.entries()).sort((a, b) => b[1] - a[1]).forEach(([k, v]) => {
                const li = document.createElement('li');
                li.textContent = `${k}：${v}`;
                ul.appendChild(li);
            });
            div.appendChild(ul);
            container.appendChild(div);
        };
        addSection('毕业学校', s.by['毕业学校']);
        addSection('民族', s.by['民族']);
        // 班级：按数字顺序（无数字的排末尾）
        (() => {
            const title = '班级';
            const map = s.by['班级'];
            const div = document.createElement('div');
            div.className = 'mb-2';
            const h = document.createElement('div');
            h.className = 'fw-bold mb-1';
            h.textContent = title;
            div.appendChild(h);
            const ul = document.createElement('ul');
            ul.className = 'mb-0';
            const toNum = (k) => {
                const m = String(k).match(/\d+/);
                return m ? parseInt(m[0], 10) : Infinity;
            };
            Array.from(map.entries()).sort((a, b) => {
                const na = toNum(a[0]);
                const nb = toNum(b[0]);
                if (na === nb) return String(a[0]).localeCompare(String(b[0]));
                return na - nb;
            }).forEach(([k, v]) => {
                const li = document.createElement('li');
                li.textContent = `${k}：${v}`;
                ul.appendChild(li);
            });
            div.appendChild(ul);
            container.appendChild(div);
        })();
        addSection('外语语种', s.by['外语语种']);
        addSection('报考科类', s.by['报考科类']);
        addSection('职业类别', s.by['职业类别']);
        addSection('政治面貌', s.by['政治面貌']);
        addSection('毕业类别', s.by['毕业类别']);
        addSection('思想品德考核结果', s.by['思想品德考核结果']);
        addSection('体育达标结果', s.by['体育达标结果']);
        addSection('考生类别', s.by['考生类别']);
        // 班级下报考科类
        const classDiv = document.createElement('div');
        classDiv.className = 'mt-3';
        const ch = document.createElement('div');
        ch.className = 'fw-bold mb-1';
        ch.textContent = '每个班级下报考科类数量';
        classDiv.appendChild(ch);
        const ul = document.createElement('ul');
        ul.className = 'mb-0';
        const toNum = (k) => {
            const m = String(k).match(/\d+/);
            return m ? parseInt(m[0], 10) : Infinity;
        };
        Array.from(s.classToKelei.entries()).sort((a, b) => {
            const na = toNum(a[0]);
            const nb = toNum(b[0]);
            if (na === nb) return String(a[0]).localeCompare(String(b[0]));
            return na - nb;
        }).forEach(([bj, map]) => {
            const li = document.createElement('li');
            const parts = Array.from(map.entries()).map(([k, v]) => `${k}:${v}`).join('，');
            li.textContent = `${bj}：${parts}`;
            ul.appendChild(li);
        });
        classDiv.appendChild(ul);
        container.appendChild(classDiv);

        // 报名点-每班-报考科类数量（报名点代码小的在前，展示为 代码-报名点）
        const bmdDiv = document.createElement('div');
        bmdDiv.className = 'mt-3';
        const bh = document.createElement('div');
        bh.className = 'fw-bold mb-1';
        bh.textContent = '按报名点分组的每班报考科类数量（代码-报名点）';
        bmdDiv.appendChild(bh);
        const bul = document.createElement('ul');
        bul.className = 'mb-0';
        const parseCodeNum = (key) => {
            const m = String(key).match(/^([0-9]+)-?/);
            return m ? parseInt(m[1], 10) : Infinity;
        };
        Array.from(s.bmdKeyToClassToKelei.entries()).sort((a, b) => parseCodeNum(a[0]) - parseCodeNum(b[0])).forEach(([bmdKey, classMap]) => {
            const li = document.createElement('li');
            const inner = document.createElement('div');
            inner.className = 'mb-1';
            inner.textContent = bmdKey;
            li.appendChild(inner);
            const innerUl = document.createElement('ul');
            Array.from(classMap.entries()).sort((a, b) => {
                const na = toNum(a[0]);
                const nb = toNum(b[0]);
                if (na === nb) return String(a[0]).localeCompare(String(b[0]));
                return na - nb;
            }).forEach(([bj, kmMap]) => {
                const cli = document.createElement('li');
                const parts = Array.from(kmMap.entries()).map(([k, v]) => `${k}:${v}`).join('，');
                cli.textContent = `${bj}：${parts}`;
                innerUl.appendChild(cli);
            });
            li.appendChild(innerUl);
            bul.appendChild(li);
        });
        bmdDiv.appendChild(bul);
        container.appendChild(bmdDiv);
    }

    function renderWarnings(arr) {
        els.warnTbody.innerHTML = '';
        arr.forEach((w, idx) => {
            const tr = document.createElement('tr');
            const cells = [idx + 1, w.ksHao || '', w.bmdCode || '', w.bmdName || '', w.bj || '', w.xm || '', w.tips.join('；')];
            cells.forEach(c => {
                const td = document.createElement('td');
                td.textContent = c;
                tr.appendChild(td);
            });
            els.warnTbody.appendChild(tr);
        });
        els.warnCount.textContent = `${arr.length} 条`;
    }

    async function handlePdf(file) {
        if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
            alert('请拖入 PDF 文件');
            return;
        }
        showLoading();
        try {
            await nextTick();
            const pages = await extractPdfText(file);
            const headerErr = validateHeader(pages);
            if (headerErr) {
                alert(headerErr);
                return;
            }
            const parsed = pages.map(p => parsePageTextToRecord(p));
            records = parsed;
            stats = buildStats(records);
            warnings = [];
            records.forEach(r => {
                const tips = validateRecord(r);
                if (tips.length) {
                    const ksHao = String(r['考生号'] || '').trim();
                    const bmdName = String(r['报名点'] || '').trim();
                    const {fullCode} = deriveBmdKeyFromKsh(ksHao, bmdName);
                    warnings.push({
                        ksHao,
                        bmdCode: fullCode || '',
                        bmdName,
                        bj: String(r['班级'] || '').trim(),
                        xm: String(r['姓名'] || '').trim(),
                        tips
                    });
                }
            });
            renderStats(stats);
            renderWarnings(warnings);
            setControlsEnabled(true);
        } finally {
            hideLoading();
        }
    }

    function exportAllExcel() {
        if (!warnings.length) return;
        const wb = XLSX.utils.book_new();
        const sheetData = warnings.map(w => ({
            "考生号": w.ksHao,
            "报名点代码": w.bmdCode,
            "报名点": w.bmdName,
            "班级": w.bj,
            "姓名": w.xm,
            "提示": w.tips.join('；'),
        }));
        const ws = XLSX.utils.json_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(wb, ws, '提示信息');
        // 附加统计页（可读性）
        const statRows = [];
        const pushMap = (title, m) => {
            statRows.push({分类: title, 选项: '', 数量: ''});
            Array.from(m.entries()).forEach(([k, v]) => statRows.push({分类: title, 选项: k, 数量: v}));
        };
        pushMap('毕业学校', stats.by['毕业学校']);
        pushMap('民族', stats.by['民族']);
        pushMap('班级', stats.by['班级']);
        pushMap('外语语种', stats.by['外语语种']);
        pushMap('报考科类', stats.by['报考科类']);
        pushMap('职业类别', stats.by['职业类别']);
        pushMap('政治面貌', stats.by['政治面貌']);
        pushMap('毕业类别', stats.by['毕业类别']);
        pushMap('思想品德考核结果', stats.by['思想品德考核结果']);
        pushMap('体育达标结果', stats.by['体育达标结果']);
        pushMap('考生类别', stats.by['考生类别']);
        const ws2 = XLSX.utils.json_to_sheet(statRows);
        XLSX.utils.book_append_sheet(wb, ws2, '统计信息');
        XLSX.writeFile(wb, '高考报名-提示与统计.xlsx');
    }

    function exportStatsExcel() {
        if (!stats) return;
        const wb = XLSX.utils.book_new();
        // sheet 1: 基础统计
        const baseRows = [];
        const pushMap = (title, m) => {
            baseRows.push({分类: title, 选项: '', 数量: ''});
            Array.from(m.entries()).forEach(([k, v]) => baseRows.push({分类: title, 选项: k, 数量: v}));
        };
        pushMap('毕业学校', stats.by['毕业学校']);
        pushMap('民族', stats.by['民族']);
        pushMap('班级', stats.by['班级']);
        pushMap('外语语种', stats.by['外语语种']);
        pushMap('报考科类', stats.by['报考科类']);
        pushMap('职业类别', stats.by['职业类别']);
        pushMap('政治面貌', stats.by['政治面貌']);
        pushMap('毕业类别', stats.by['毕业类别']);
        pushMap('思想品德考核结果', stats.by['思想品德考核结果']);
        pushMap('体育达标结果', stats.by['体育达标结果']);
        pushMap('考生类别', stats.by['考生类别']);
        const ws1 = XLSX.utils.json_to_sheet(baseRows);
        XLSX.utils.book_append_sheet(wb, ws1, '基础统计');

        // sheet 2: 班级-报考科类数量
        const classRows = [];
        Array.from(stats.classToKelei.entries()).forEach(([bj, map]) => {
            Array.from(map.entries()).forEach(([kl, cnt]) => {
                classRows.push({班级: bj, 报考科类: kl, 数量: cnt});
            });
        });
        const ws2 = XLSX.utils.json_to_sheet(classRows);
        XLSX.utils.book_append_sheet(wb, ws2, '班级-报考科类');

        // sheet 3: 报名点(代码-名称)-班级-报考科类数量 + 报名点合计
        const bmdRows = [];
        Array.from(stats.bmdKeyToClassToKelei.entries()).forEach(([bmdKey, classMap]) => {
            let total = 0;
            Array.from(classMap.values()).forEach(kmMap => {
                Array.from(kmMap.values()).forEach(v => total += v);
            });
            bmdRows.push({报名点: bmdKey, 班级: '', 报考科类: '合计', 数量: total});
            Array.from(classMap.entries()).forEach(([bj, kmMap]) => {
                Array.from(kmMap.entries()).forEach(([kl, cnt]) => {
                    bmdRows.push({报名点: bmdKey, 班级: bj, 报考科类: kl, 数量: cnt});
                });
            });
        });
        const ws3 = XLSX.utils.json_to_sheet(bmdRows);
        XLSX.utils.book_append_sheet(wb, ws3, '报名点-班级-科类');

        XLSX.writeFile(wb, '高考报名-统计.xlsx');
    }

    async function exportZip() {
        if (!warnings.length) return;
        const mode = String(els.groupLenSelect.value);
        const zip = new JSZip();
        const byGroup = new Map();
        warnings.forEach(w => {
            const code = String(w.bmdCode || '');
            let key = '未知分组';
            if (code) key = (mode === 'code3') ? code.substring(0, 3) : code;
            if (!byGroup.has(key)) byGroup.set(key, []);
            byGroup.get(key).push(w);
        });
        for (const [g, arr] of byGroup.entries()) {
            const wb = XLSX.utils.book_new();
            const sheetData = arr.map(w => ({
                "考生号": w.ksHao,
                "报名点代码": w.bmdCode,
                "报名点": w.bmdName,
                "班级": w.bj,
                "姓名": w.xm,
                "提示": w.tips.join('；')
            }));
            const ws = XLSX.utils.json_to_sheet(sheetData);
            XLSX.utils.book_append_sheet(wb, ws, '提示信息');
            const wbout = XLSX.write(wb, {bookType: 'xlsx', type: 'array'});
            zip.file(`提示-${g}.xlsx`, wbout);
        }
        const blob = await zip.generateAsync({type: 'blob'});
        downloadBlob(blob, mode === 'code3' ? '按报名点代码前3位分组.zip' : '按报名点代码分组.zip');
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
        els.fileInput.addEventListener('change', e => {
            const f = e.target.files && e.target.files[0];
            if (f) handlePdf(f);
        });
        els.btnClear.addEventListener('click', () => {
            records = [];
            warnings = [];
            stats = null;
            els.stats.innerHTML = '';
            renderWarnings([]);
            setControlsEnabled(false);
        });
        els.btnExportAll.addEventListener('click', exportAllExcel);
        if (els.btnExportStats) els.btnExportStats.addEventListener('click', exportStatsExcel);
        els.chkZip.addEventListener('change', () => {
            els.groupLenSelect.disabled = !els.chkZip.checked || els.btnExportAll.disabled;
            els.btnExportZip.disabled = !els.chkZip.checked || els.btnExportAll.disabled;
        });
        els.btnExportZip.addEventListener('click', exportZip);
    }

    // init
    wireEvents();
})();


