/**
 * 美术与设计类统考 · 考生信息桌贴
 *
 * 流程：导入含 ZWH（座位号）的名单 → 按考场分组、按座位号排序 → 按「列数 × 行数」分页 → A4 纵向打印。
 *
 * 每张桌贴展示 6 项信息（全部必需，不做取舍）：
 *   座位号（最大）、姓名、性别（由身份证第 17 位识别）、考场、准考证号：xxxx、考生号：xxxx。
 *
 * 版式说明：
 *   1. 默认 3 列 × 5 行（每页 15 张），可改为 2–3 列、5–15 行；
 *   2. 字号不写死，按 A4 可用打印区域（194mm × 281mm）与行列数计算：
 *      单元格宽 = 194 / 列数，单元格高 = 281 / 行数，
 *      每类文字取「高度占比」与「宽度占比」的较小值，行列变化都会同步适配；
 *   3. 排列形式自动选择：格子偏方/偏窄时上下排列（座位号在上、信息在下），
 *      格子偏扁宽时改为左右排列（座位号在左、信息在右），两种排版分别算出可用字号，
 *      取「最小字号更大」的一种，避免扁宽格子里硬塞 5 行导致文字过小。
 *
 * 所有解析均在浏览器本地完成，不发起网络请求，也不写入本地存储。
 */
(function (global) {
    'use strict';

    /* ------------------------------------------------------------------ */
    /* 常量                                                                */
    /* ------------------------------------------------------------------ */

    var SHEET = {width: 210, height: 297, padX: 8, padY: 8};   // A4 纵向（mm）
    var GRID_W = SHEET.width - SHEET.padX * 2;                 // 194mm
    var GRID_H = SHEET.height - SHEET.padY * 2;                // 281mm
    var MM_PER_PT = 25.4 / 72;                                 // 1pt = 0.3528mm

    var CELL_PAD = 1.2;        // 单元格内边距（mm）
    var SAFE = 0.97;           // 安全系数，避免取整后顶到单元格边框
    var META_EM = 11.9;        // 「考生号：14 位数字」的估算宽度（em）
    var NAME_EM = 4.5;         // 「姓名 + 性别」的估算宽度（em）
    var SEAT_EM = 2.2;         // 两位座位号的估算宽度（em）

    /* 上下排列：座位号 3.2b、姓名 1.35b、三行明细 0.78b，行距 0.28b */
    var V = {seat: 3.2, name: 1.35, meta: 0.78, gap: 0.28};
    var V_HEIGHT = V.seat * 1.02 + V.name * 1.18 + 3 * V.meta * 1.25 + 4 * V.gap; // 8.90b
    var V_WIDTH = META_EM * V.meta;                                              // 9.28b

    /* 左右排列：右侧姓名 1.5b、三行明细 1b，行距 0.28b */
    var H = {name: 1.5, meta: 1.0, gap: 0.28};
    var H_HEIGHT = H.name * 1.18 + 3 * H.meta * 1.25 + 4 * H.gap;                 // 6.64b
    var H_WIDTH = META_EM * H.meta;                                              // 11.9b
    var H_NAME_WIDTH = NAME_EM * H.name;                                         // 6.75b

    var DEFAULT_COLS = 3;
    var DEFAULT_ROWS = 5;
    var COLS_OPTIONS = [2, 3];
    var ROWS_OPTIONS = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15];

    var HEADER_SCAN = 20;
    var NO_ROOM = '未编号';

    /* 需要的列：key 为内部字段名，aliases 为可识别的表头写法（兼容中文表头与导出字段名）。 */
    var FIELDS = [
        {key: 'zwh', label: 'ZWH（座位号）', aliases: ['ZWH', '座位号', '座位'], required: true},
        {key: 'kc', label: 'KC（考场号）', aliases: ['KC', '考场号', '考场'], required: true},
        {key: 'xm', label: 'XM（姓名）', aliases: ['XM', '姓名', '考生姓名'], required: true},
        {key: 'zkzh', label: 'ZKZH（准考证号）', aliases: ['ZKZH', '准考证号', '准考证'], required: true},
        {key: 'ksh', label: 'KSH（考生号）', aliases: ['KSH', '考生号'], required: true},
        {key: 'sfzh', label: 'SFZH（身份证号）', aliases: ['SFZH', '身份证号', '身份证'], required: false},
        {key: 'xbdm', label: 'XBDM（性别代码）', aliases: ['XBDM', '性别代码', '性别'], required: false},
        {key: 'kdmc', label: 'KDMC（考点）', aliases: ['KDMC', '考点名称', '考点'], required: false}
    ];

    /* ------------------------------------------------------------------ */
    /* 文本与表头处理                                                      */
    /* ------------------------------------------------------------------ */

    function normalizeHeader(value) {
        return String(value === null || value === undefined ? '' : value)
            .replace(/[\s\u3000\uFEFF]/g, '')
            .replace(/_+$/, '')
            .toUpperCase();
    }

    function cellText(value) {
        if (value === null || value === undefined) return '';
        if (typeof value === 'number') return String(value);
        return String(value).replace(/^[\s\u3000\uFEFF']+|[\s\u3000\uFEFF]+$/g, '');
    }

    /** 表头行匹配：返回 {map: {字段: 列下标}, missing: [缺失列说明]} */
    function matchHeader(row) {
        var normalized = row.map(normalizeHeader);
        var map = {};
        var missing = [];

        FIELDS.forEach(function (field) {
            var aliases = field.aliases.map(normalizeHeader);
            var hit = -1;
            for (var i = 0; i < normalized.length; i++) {
                if (normalized[i] && aliases.indexOf(normalized[i]) >= 0) {
                    hit = i;
                    break;
                }
            }
            if (hit >= 0) {
                map[field.key] = hit;
            } else if (field.required) {
                missing.push(field.label);
            }
        });

        return {map: map, missing: missing};
    }

    /** 座位号统一为两位（1 / 01 → 01；100 保持 100） */
    function normalizeZwh(value) {
        var text = cellText(value).replace(/[\s\u3000\uFEFF]/g, '');
        if (!text) return '';
        if (/^\d+$/.test(text)) {
            return text.length < 2 ? '0' + text : text;
        }
        return text;
    }

    /** 考场号统一为三位：1 / 01 / 001 → 001 */
    function normalizeRoom(value) {
        var text = cellText(value).replace(/[\s\u3000\uFEFF]/g, '');
        if (!text) return '';
        if (/^\d+$/.test(text)) {
            return text.length >= 3 ? text : ('000' + text).slice(-3);
        }
        return text;
    }

    /**
     * 性别：优先用身份证第 17 位（奇男偶女），
     * 身份证缺失或格式不对时退回 XBDM（1=男、2=女）。
     */
    function resolveGender(sfzh, xbdm) {
        var id = String(sfzh || '').replace(/[\s\u3000\uFEFF]/g, '');
        if (/^\d{17}[\dXx]$/.test(id)) {
            return Number(id.charAt(16)) % 2 === 1 ? '男' : '女';
        }
        var code = String(xbdm || '').trim();
        if (code === '1') return '男';
        if (code === '2') return '女';
        return '';
    }

    function compareRoom(a, b) {
        var na = parseInt(a, 10);
        var nb = parseInt(b, 10);
        if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
        return String(a).localeCompare(String(b), 'zh-CN', {numeric: true});
    }

    /** 同一考场内按座位号升序；座位号相同再按准考证号 */
    function compareSeat(a, b) {
        var za = Number(a.zwh);
        var zb = Number(b.zwh);
        if (isFinite(za) && isFinite(zb) && za !== zb) return za - zb;
        var byZwh = String(a.zwh).localeCompare(String(b.zwh), 'zh-CN', {numeric: true});
        if (byZwh !== 0) return byZwh;
        return String(a.zkzh).localeCompare(String(b.zkzh), 'zh-CN', {numeric: true});
    }

    /* ------------------------------------------------------------------ */
    /* 字号与排列计算                                                      */
    /* ------------------------------------------------------------------ */

    function roundHalf(value) {
        return Math.round(value * 2) / 2;
    }

    function mmToPt(mm) {
        return mm / MM_PER_PT;
    }

    /**
     * 按每页行列数计算字号与排列方式。
     * @returns {{perPage:number, mode:'vertical'|'horizontal', seat:number, name:number, meta:number,
     *            gap:number, cellW:number, cellH:number}}
     *          字号单位为 pt，gap / cellW / cellH 为 mm。
     */
    function computeLayout(cols, rows) {
        var perCols = Math.max(1, Number(cols) || DEFAULT_COLS);
        var perRows = Math.max(1, Number(rows) || DEFAULT_ROWS);
        var cellW = GRID_W / perCols;
        var cellH = GRID_H / perRows;
        var aw = Math.max(1, cellW - CELL_PAD * 2);
        var ah = Math.max(1, cellH - CELL_PAD * 2);

        // 上下排列：座位号在上，姓名 + 三行明细在下
        var bv = Math.min((ah * SAFE) / V_HEIGHT, (aw * SAFE) / V_WIDTH);
        var seatV = V.seat * bv;
        var vMin = Math.min(V.name * bv, V.meta * bv);

        // 左右排列：座位号占左侧 36%，右侧四行
        var leftW = aw * 0.36;
        var rightW = Math.max(1, aw - leftW - 2);
        var bh = Math.min(
            (ah * SAFE) / H_HEIGHT,
            (rightW * SAFE) / H_WIDTH,
            (rightW * SAFE) / H_NAME_WIDTH
        );
        var hMin = Math.min(H.name * bh, H.meta * bh);

        // 两种排布各算一遍，取「最小字号」更大的那一种：
        // 格子偏方/偏窄时上下排列更划算，格子偏扁宽时左右排列能塞下更大的明细文字。
        var mode = vMin >= hMin ? 'vertical' : 'horizontal';
        var sizes = mode === 'vertical'
            ? {seat: seatV, name: V.name * bv, meta: V.meta * bv, gap: V.gap * bv}
            : {
                // 座位号始终不超过上下排列时的尺寸，保证「张数越少字号越大」不被排列切换破坏
                seat: Math.min(seatV, ah * 0.85, leftW / 1.2),
                name: H.name * bh,
                meta: H.meta * bh,
                gap: H.gap * bh
            };

        return {
            perPage: perCols * perRows,
            mode: mode,
            seat: roundHalf(mmToPt(sizes.seat)),
            name: roundHalf(mmToPt(sizes.name)),
            meta: roundHalf(mmToPt(sizes.meta)),
            gap: Math.round(sizes.gap * 10) / 10,
            cellW: Math.round(cellW * 10) / 10,
            cellH: Math.round(cellH * 10) / 10
        };
    }

    /* ------------------------------------------------------------------ */
    /* 解析 Excel                                                          */
    /* ------------------------------------------------------------------ */

    function parseWorkbookBuffer(buffer) {
        var xlsx = global.XLSX;
        if (!xlsx) throw new Error('未加载 SheetJS 解析库，请检查 lib/sheetjs 路径是否正确。');

        var workbook = xlsx.read(new Uint8Array(buffer), {type: 'array'});
        var sheetNames = workbook.SheetNames || [];
        var fallbackHeaders = [];

        for (var s = 0; s < sheetNames.length; s++) {
            var sheet = workbook.Sheets[sheetNames[s]];
            if (!sheet) continue;

            var aoa = xlsx.utils.sheet_to_json(sheet, {
                header: 1,
                raw: false,
                defval: '',
                blankrows: false
            });

            var head = null;
            for (var i = 0; i < aoa.length && i < HEADER_SCAN; i++) {
                var row = aoa[i] || [];
                if (!fallbackHeaders.length) {
                    fallbackHeaders = row.map(cellText).filter(Boolean).slice(0, 12);
                }
                var result = matchHeader(row);
                if (result.missing.length === 0) {
                    head = {index: i, map: result.map};
                    break;
                }
            }
            if (!head) continue;

            var records = [];
            var noSeat = 0;
            var noName = 0;
            var noGender = 0;

            for (var r = head.index + 1; r < aoa.length; r++) {
                var line = aoa[r] || [];
                var zwh = normalizeZwh(line[head.map.zwh]);
                var xm = cellText(line[head.map.xm]);
                var kc = normalizeRoom(line[head.map.kc]) || NO_ROOM;
                var zkzh = cellText(line[head.map.zkzh]);
                var ksh = cellText(line[head.map.ksh]);
                var sfzh = head.map.sfzh === undefined ? '' : cellText(line[head.map.sfzh]);
                var xbdm = head.map.xbdm === undefined ? '' : cellText(line[head.map.xbdm]);
                var kdmc = head.map.kdmc === undefined ? '' : cellText(line[head.map.kdmc]);

                if (!zwh && !xm && !zkzh && !ksh && kc === NO_ROOM) continue;
                // 没有座位号就无法贴到桌面，直接跳过并统计
                if (!zwh) {
                    noSeat++;
                    continue;
                }
                if (!xm) noName++;

                var gender = resolveGender(sfzh, xbdm);
                if (!gender) noGender++;

                records.push({
                    zwh: zwh,
                    xm: xm,
                    gender: gender,
                    kc: kc,
                    zkzh: zkzh,
                    ksh: ksh,
                    kdmc: kdmc
                });
            }

            if (!records.length) {
                throw new Error('工作表「' + sheetNames[s] + '」中没有读取到有效数据行（需要 ZWH、KC、XM、ZKZH、KSH 五列且有数据）。');
            }

            return {
                sheetName: sheetNames[s],
                records: records,
                noSeat: noSeat,
                noName: noName,
                noGender: noGender
            };
        }

        throw new Error('未找到必需的列：ZWH（座位号）、KC（考场号）、XM（姓名）、ZKZH（准考证号）、KSH（考生号）。'
            + (fallbackHeaders.length ? '当前文件首个表头行为：' + fallbackHeaders.join('、') : ''));
    }

    /* ------------------------------------------------------------------ */
    /* 分页                                                                */
    /* ------------------------------------------------------------------ */

    /**
     * 按考场分组 → 组内按座位号排序 → 按 perPage 切页（缺位补空，保证网格整齐）。
     * @returns Array<{key, kc, part, parts, count, slots}>
     */
    function buildPages(records, cols, rows) {
        var perPage = Math.max(1, (Number(cols) || DEFAULT_COLS) * (Number(rows) || DEFAULT_ROWS));

        var grouped = {};
        var order = [];
        (records || []).forEach(function (item) {
            if (!grouped[item.kc]) {
                grouped[item.kc] = [];
                order.push(item.kc);
            }
            grouped[item.kc].push(item);
        });

        order.sort(compareRoom);

        var pages = [];
        order.forEach(function (kc) {
            var list = grouped[kc].slice().sort(compareSeat);
            var parts = Math.ceil(list.length / perPage) || 1;

            for (var start = 0; start < list.length; start += perPage) {
                var chunk = list.slice(start, start + perPage);
                var slots = chunk.map(function (item) {
                    return {
                        zwh: item.zwh,
                        xm: item.xm,
                        gender: item.gender,
                        kc: kc,
                        zkzh: item.zkzh,
                        ksh: item.ksh
                    };
                });
                while (slots.length < perPage) slots.push(null);

                pages.push({
                    key: kc + '-' + (pages.length + 1),
                    kc: kc,
                    part: 1,
                    parts: parts,
                    count: chunk.length,
                    slots: slots
                });
            }
        });

        // 同一考场内的页码需要单独编号
        var partCursor = {};
        pages.forEach(function (page) {
            partCursor[page.kc] = (partCursor[page.kc] || 0) + 1;
            page.part = partCursor[page.kc];
        });

        return pages;
    }

    /** 数据质量提示 */
    function collectWarnings(parsed) {
        var warnings = [];
        var records = parsed.records || [];
        var seen = {};
        var duplicate = 0;
        var noRoom = 0;
        var roomMax = {};

        records.forEach(function (item) {
            var seatKey = item.kc + '|' + item.zwh;
            if (seen[seatKey]) duplicate++;
            else seen[seatKey] = true;
            if (item.kc === NO_ROOM) noRoom++;
            roomMax[item.kc] = (roomMax[item.kc] || 0) + 1;
        });

        if (duplicate) warnings.push('有 ' + duplicate + ' 个「考场 + 座位号」重复，请核对原始名单。');
        if (noRoom) warnings.push('有 ' + noRoom + ' 条记录考场号为空，已归入「' + NO_ROOM + '」。');
        if (parsed.noSeat) warnings.push('有 ' + parsed.noSeat + ' 条记录缺少座位号（ZWH），无法生成桌贴，已跳过。');
        if (parsed.noName) warnings.push('有 ' + parsed.noName + ' 条记录姓名为空。');
        if (parsed.noGender) warnings.push('有 ' + parsed.noGender + ' 条记录无法识别性别（身份证号缺失或格式异常，且没有性别代码）。');

        var over = Object.keys(roomMax).filter(function (kc) {
            return roomMax[kc] > 30;
        });
        if (over.length) {
            warnings.push('考场 ' + over.join('、') + ' 的人数超过 30 人，请确认座位号是否重复。');
        }

        return warnings;
    }

    /* 供调试与本地校验使用 */
    global.DeskSticker = {
        DEFAULT_COLS: DEFAULT_COLS,
        DEFAULT_ROWS: DEFAULT_ROWS,
        COLS_OPTIONS: COLS_OPTIONS,
        ROWS_OPTIONS: ROWS_OPTIONS,
        GRID_W: GRID_W,
        GRID_H: GRID_H,
        computeLayout: computeLayout,
        parseWorkbookBuffer: parseWorkbookBuffer,
        buildPages: buildPages,
        collectWarnings: collectWarnings,
        resolveGender: resolveGender
    };

    if (typeof global.Vue === 'undefined') return; // Node/测试环境下只暴露纯函数

    /* ------------------------------------------------------------------ */
    /* 页面逻辑                                                            */
    /* ------------------------------------------------------------------ */

    var app = global.Vue.createApp({
        data: function () {
            return {
                colsOptions: COLS_OPTIONS,
                rowsOptions: ROWS_OPTIONS,

                cols: DEFAULT_COLS,
                rows: DEFAULT_ROWS,

                fileState: 'idle',   // idle | reading | ready | error
                fileName: '',
                errorMsg: '',
                dragOver: false,

                parsed: null,
                records: [],
                warnings: [],

                toastText: ''
            };
        },

        computed: {
            uploadClass: function () {
                return {
                    'is-dragging': this.dragOver,
                    'is-ready': this.fileState === 'ready',
                    'is-error': this.fileState === 'error'
                };
            },

            layout: function () {
                return computeLayout(this.cols, this.rows);
            },

            gridStyle: function () {
                var layout = this.layout;
                return {
                    '--cols': this.cols,
                    '--rows': this.rows,
                    '--seat-size': layout.seat + 'pt',
                    '--name-size': layout.name + 'pt',
                    '--meta-size': layout.meta + 'pt',
                    '--gap': layout.gap + 'mm'
                };
            },

            /** 字号过小时的提示（六项信息都必须展示，因此只提示不隐藏） */
            sizeHint: function () {
                var layout = this.layout;
                if (layout.meta >= 6) return '';
                return '当前每页 ' + layout.perPage + ' 张，明细字号约 ' + layout.meta
                    + ' pt，打印后偏小；建议减少行数或列数。';
            },

            pages: function () {
                if (!this.records.length) return [];
                return buildPages(this.records, this.cols, this.rows);
            },

            pageCount: function () {
                return this.pages.length;
            },

            roomCount: function () {
                var seen = {};
                this.records.forEach(function (item) {
                    seen[item.kc] = true;
                });
                return Object.keys(seen).length;
            }
        },

        beforeUnmount: function () {
            if (this._toastTimer) clearTimeout(this._toastTimer);
        },

        methods: {
            pickFile: function () {
                if (this.$refs.picker) this.$refs.picker.click();
            },

            onPick: function (event) {
                var file = event.target.files && event.target.files[0];
                if (file) this.handleFile(file);
                event.target.value = '';
            },

            onDragOver: function () {
                this.dragOver = true;
            },

            onDragLeave: function () {
                this.dragOver = false;
            },

            onDrop: function (event) {
                this.dragOver = false;
                var file = event.dataTransfer.files && event.dataTransfer.files[0];
                if (file) this.handleFile(file);
            },

            handleFile: function (file) {
                if (!/\.(xlsx|xls)$/i.test(file.name || '')) {
                    this.fail('仅支持 .xlsx / .xls 文件。');
                    return;
                }

                var self = this;
                this.fileState = 'reading';
                this.fileName = file.name;
                this.errorMsg = '';

                var reader = new FileReader();
                reader.onload = function () {
                    try {
                        self.load(reader.result, file.name);
                    } catch (err) {
                        self.fail(err && err.message ? err.message : '文件解析失败，请检查文件内容。');
                    }
                };
                reader.onerror = function () {
                    self.fail('文件读取失败，请重新选择。');
                };
                reader.readAsArrayBuffer(file);
            },

            load: function (buffer, name) {
                var parsed = parseWorkbookBuffer(buffer);
                this.parsed = parsed;
                this.records = parsed.records;
                this.warnings = collectWarnings(parsed);
                this.fileState = 'ready';
                this.fileName = name;
                this.errorMsg = '';
                this.showToast('已读取 ' + parsed.records.length + ' 名考生');
            },

            fail: function (message) {
                this.fileState = 'error';
                this.errorMsg = message;
                this.parsed = null;
                this.records = [];
                this.warnings = [];
            },

            doPrint: function () {
                if (!this.pages.length) return;
                window.print();
            },

            reset: function () {
                this.fileState = 'idle';
                this.fileName = '';
                this.errorMsg = '';
                this.dragOver = false;
                this.parsed = null;
                this.records = [];
                this.warnings = [];
                if (this.$refs.picker) this.$refs.picker.value = '';
            },

            showToast: function (text) {
                this.toastText = text;
                if (this._toastTimer) clearTimeout(this._toastTimer);
                this._toastTimer = setTimeout(function () {
                    this.toastText = '';
                }.bind(this), 2600);
            }
        }
    });

    global.__deskStickerApp = app.mount('#app');

})(typeof window !== 'undefined' ? window : globalThis);
