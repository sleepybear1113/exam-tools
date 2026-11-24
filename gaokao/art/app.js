/**
 * 美术与设计类统考 · 考生诚信考试承诺签名表
 *
 * 流程：拖入「美术报名库」Excel → 按考场分组 → 每 30 人一页 → 生成可打印的 A4 表格。
 * 版面参照 sample/嘉兴市2027年美术与设计类统一考试考生诚信考试承诺签名表.xlsx，
 * 考场号直接拼在标题末尾（如 …承诺签名表-001），表格内只保留 座位号 / 准考证号 / 姓名 / 签名。
 *
 * 说明：所有解析均在浏览器本地完成，不发起任何网络请求，也不写入本地存储。
 */
(function (global) {
    'use strict';

    /* ------------------------------------------------------------------ */
    /* 常量                                                                */
    /* ------------------------------------------------------------------ */

    var DEFAULT_PREFIX = '嘉兴市';
    var TITLE_SUFFIX = '年美术与设计类专业统一考试考生诚信考试承诺签名表';

    /* 承诺书正文：中间一句需要加粗、下划线并放大一号 */
    var PLEDGE_HEAD = '我已经学习了《浙江省普通高校招生美术与设计类专业统一考试考场规则》及《国家教育考试违规处理办法》（摘录一），'
        + '充分认识到违反考场规则及考试违纪行为所应承担的严重后果。';
    var PLEDGE_EM = '特别是，在考试开始信号发出前答题，或在考试结束信号发出后继续答题，均属考试违纪行为，一经查实，将被取消该科目考试成绩。';
    var PLEDGE_TAIL = '我郑重承诺，将严格遵守上述规定，自觉维护考试的公平与公正。';
    var PLEDGE = PLEDGE_HEAD + PLEDGE_EM + PLEDGE_TAIL;

    var ROWS_PER_PAGE = 30;   // 一个考场不超过 30 人，即一页最多 30 行
    var HEADER_SCAN = 20;     // 表头最多向下探测的行数
    var NO_ROOM = '未编号';

    /* 需要的列：key 为内部字段名，aliases 为可识别的表头写法。
       KSH 目前不展示（后续与高考报名库 BMXH_ 关联时才需要），因此不作为必需列。 */
    var FIELDS = [
        {key: 'zkzh', label: 'ZKZH（准考证号）', aliases: ['ZKZH', '准考证号', '准考证'], required: true},
        {key: 'xm', label: 'XM（姓名）', aliases: ['XM', '姓名', '考生姓名'], required: true},
        {key: 'kc', label: 'KC（考场号）', aliases: ['KC', '考场号', '考场'], required: true},
        {key: 'ksh', label: 'KSH（考生号）', aliases: ['KSH', '考生号'], required: false}
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

    /** 考场号统一为三位：1 / 01 / 001 → 001 */
    function normalizeRoom(value) {
        var text = cellText(value);
        if (!text) return '';
        if (/^\d+$/.test(text)) {
            return text.length >= 3 ? text : ('000' + text).slice(-3);
        }
        return text;
    }

    function compareRoom(a, b) {
        var na = parseInt(a, 10);
        var nb = parseInt(b, 10);
        if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
        return String(a).localeCompare(String(b), 'zh-CN', {numeric: true});
    }

    function compareZkzh(a, b) {
        var na = Number(a.zkzh);
        var nb = Number(b.zkzh);
        if (isFinite(na) && isFinite(nb) && na !== nb) return na - nb;
        return String(a.zkzh).localeCompare(String(b.zkzh), 'zh-CN', {numeric: true});
    }

    /** 考试年份：当年 10 月及以后都算作第二年的考试（如 2026-11 → 2027） */
    function computeExamYear(now) {
        var date = now instanceof Date ? now : new Date();
        return date.getMonth() >= 9 ? date.getFullYear() + 1 : date.getFullYear();
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
                raw: true,
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
            for (var r = head.index + 1; r < aoa.length; r++) {
                var line = aoa[r] || [];
                var zkzh = cellText(line[head.map.zkzh]);
                var xm = cellText(line[head.map.xm]);
                var ksh = cellText(line[head.map.ksh]);
                var kc = normalizeRoom(line[head.map.kc]);

                // 整行为空则跳过（导出文件常见的尾部空行）
                if (!zkzh && !xm && !ksh && !kc) continue;
                // 没有准考证号无法确定座位，直接忽略
                if (!zkzh) continue;

                records.push({zkzh: zkzh, xm: xm, kc: kc || NO_ROOM, ksh: ksh, seat: 0});
            }

            if (!records.length) {
                throw new Error('工作表「' + sheetNames[s] + '」中没有读取到有效数据行。');
            }

            return {sheetName: sheetNames[s], records: records};
        }

        throw new Error('未找到必需的列：ZKZH、XM、KC、KSH。'
            + (fallbackHeaders.length ? '当前文件首个表头行为：' + fallbackHeaders.join('、') : ''));
    }

    /* ------------------------------------------------------------------ */
    /* 座位号与分页                                                        */
    /* ------------------------------------------------------------------ */

    /**
     * 座位号优先按准考证号推算（同一考场内准考证号连续递增，每 30 人循环），
     * 推算结果不是 1~30 的完整排列时退回按排序序号编号。
     */
    function assignSeats(list) {
        var numbers = list.map(function (item) {
            return Number(item.zkzh);
        });
        var usable = numbers.length > 0 && numbers.every(function (n) {
            return isFinite(n) && n > 0;
        });
        var seats = [];

        if (usable) {
            var first = numbers[0];
            seats = numbers.map(function (n) {
                return (((n - first) % ROWS_PER_PAGE) + ROWS_PER_PAGE) % ROWS_PER_PAGE + 1;
            });
            var used = {};
            usable = seats.every(function (seat) {
                if (seat < 1 || seat > ROWS_PER_PAGE || used[seat]) return false;
                used[seat] = true;
                return true;
            });
        }

        if (!usable) {
            seats = list.map(function (item, index) {
                return index + 1;
            });
        }

        list.forEach(function (item, index) {
            item.seat = seats[index];
        });

        return list;
    }

    /** 记录 → 页（每页一个考场分段，不足 30 行补空行） */
    function buildPages(records) {
        var grouped = {};
        var order = [];

        records.forEach(function (item) {
            if (!grouped[item.kc]) {
                grouped[item.kc] = [];
                order.push(item.kc);
            }
            grouped[item.kc].push(item);
        });

        order.sort(compareRoom);

        var pages = [];
        var warnings = [];

        order.forEach(function (kc) {
            var list = grouped[kc].slice().sort(compareZkzh);
            assignSeats(list);

            var chunks = [];
            for (var i = 0; i < list.length; i += ROWS_PER_PAGE) {
                chunks.push(list.slice(i, i + ROWS_PER_PAGE));
            }

            if (chunks.length > 1) {
                warnings.push('考场 ' + kc + ' 共 ' + list.length
                    + ' 人，超过每页 30 人，已自动拆分为 ' + chunks.length + ' 页。');
            }

            chunks.forEach(function (chunk, ci) {
                var rows = chunk.slice();
                while (rows.length < ROWS_PER_PAGE) rows.push(null);

                pages.push({
                    key: kc + '-' + (ci + 1),
                    kc: kc,
                    count: chunk.length,
                    part: ci + 1,
                    parts: chunks.length,
                    rows: rows
                });
            });
        });

        return {pages: pages, warnings: warnings};
    }

    /** 数据质量提示 */
    function collectWarnings(records) {
        var seen = {};
        var duplicate = 0;
        var noName = 0;
        var noRoom = 0;

        records.forEach(function (item) {
            if (seen[item.zkzh]) duplicate++;
            else seen[item.zkzh] = true;
            if (!item.xm) noName++;
            if (item.kc === NO_ROOM) noRoom++;
        });

        var warnings = [];
        if (duplicate) warnings.push('存在 ' + duplicate + ' 个重复的准考证号，请核对原始数据。');
        if (noName) warnings.push('有 ' + noName + ' 条记录姓名为空。');
        if (noRoom) warnings.push('有 ' + noRoom + ' 条记录考场号为空，已归入「' + NO_ROOM + '」考场。');
        return warnings;
    }

    /** 组装标题：前缀 + 年份 + 固定后缀 + 「-考场号」 */
    function buildTitle(prefix, year, room) {
        var head = String(prefix || '').replace(/[\s\u3000]+$/, '');
        var title = head + year + TITLE_SUFFIX;
        if (room) title += '-' + room;
        return title;
    }

    /* 供调试与本地校验使用 */
    global.ArtPledge = {
        DEFAULT_PREFIX: DEFAULT_PREFIX,
        TITLE_SUFFIX: TITLE_SUFFIX,
        PLEDGE: PLEDGE,
        ROWS_PER_PAGE: ROWS_PER_PAGE,
        parseWorkbookBuffer: parseWorkbookBuffer,
        buildPages: buildPages,
        collectWarnings: collectWarnings,
        assignSeats: assignSeats,
        computeExamYear: computeExamYear,
        buildTitle: buildTitle
    };

    if (typeof global.Vue === 'undefined') return; // Node/测试环境下只暴露纯函数

    /* ------------------------------------------------------------------ */
    /* 页面逻辑                                                            */
    /* ------------------------------------------------------------------ */

    var XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    /** base64（内嵌模板）→ ArrayBuffer */
    function base64ToArrayBuffer(b64) {
        var bin = global.atob(b64);
        var len = bin.length;
        var bytes = new Uint8Array(len);
        for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
        return bytes.buffer;
    }

    var app = global.Vue.createApp({
        data: function () {
            return {
                pledgeHead: PLEDGE_HEAD,
                pledgeEm: PLEDGE_EM,
                pledgeTail: PLEDGE_TAIL,
                titlePrefix: DEFAULT_PREFIX,
                examYear: computeExamYear(),

                pageDragging: false,
                cardDragging: '',

                // 美术报名库（主表）
                fileState: 'idle',   // idle | reading | ready | error
                fileName: '',
                errorMsg: '',
                records: [],
                pages: [],
                warnings: [],

                // 高考报名库
                gkState: 'idle',
                gkFileName: '',
                gkErrorMsg: '',
                gkRecords: [],

                // 报名点映射（与数据统计工具共用 localStorage）
                bmdMap: {},
                bmdMapCount: 0,

                // 手机收缴登记表
                phoneGroups: [],
                phoneWarnings: [],
                phoneMatched: 0,
                phoneUnmatched: 0,
                phoneUnmatchedList: [],
                downloadingCode: '',
                zipping: false,
                zipProgress: '',

                toastText: '',
                toastType: 'info',

                // 当前展示的模块：overview | pledge | phone
                view: 'overview'
            };
        },

        computed: {
            hasAny: function () {
                return this.records.length > 0 || this.gkRecords.length > 0;
            },

            uploadClass: function () {
                return {
                    'is-dragging': this.pageDragging || this.cardDragging === 'art',
                    'is-ready': this.fileState === 'ready',
                    'is-error': this.fileState === 'error'
                };
            },

            gkUploadClass: function () {
                return {
                    'is-dragging': this.pageDragging || this.cardDragging === 'gk',
                    'is-ready': this.gkState === 'ready',
                    'is-error': this.gkState === 'error'
                };
            },

            /** 不含考场号的标题主体，用于页面提示 */
            baseTitle: function () {
                return buildTitle(this.titlePrefix, this.examYear, '');
            },

            stats: function () {
                var rooms = {};
                this.records.forEach(function (item) {
                    rooms[item.kc] = true;
                });
                return {
                    total: this.records.length,
                    rooms: Object.keys(rooms).length,
                    pages: this.pages.length
                };
            }
        },

        mounted: function () {
            this.onWindowDragOver = this.handleWindowDragOver.bind(this);
            this.onWindowDrop = this.handleWindowDrop.bind(this);

            window.addEventListener('dragover', this.onWindowDragOver);
            window.addEventListener('drop', this.onWindowDrop);

            this.refreshBmdMap(true);
        },

        beforeUnmount: function () {
            window.removeEventListener('dragover', this.onWindowDragOver);
            window.removeEventListener('drop', this.onWindowDrop);
            if (this._hideTimer) clearTimeout(this._hideTimer);
            if (this._toastTimer) clearTimeout(this._toastTimer);
        },

        methods: {
            /* ---------------- 模块切换 ---------------- */

            openView: function (view) {
                // 承诺书必须有美术库；收缴表允许进入后再提示缺少哪个文件
                if (view === 'pledge' && !this.records.length) return;
                this.view = view;
                this.$nextTick(function () {
                    window.scrollTo({top: 0, behavior: 'smooth'});
                });
            },

            /* ---------------- 标题 ---------------- */

            pageTitle: function (page) {
                return buildTitle(this.titlePrefix, this.examYear, page.kc);
            },

            /* ---------------- 文件选择与拖拽 ---------------- */

            pickFile: function (kind) {
                var ref = kind === 'gk' ? 'gkPicker' : 'artPicker';
                if (this.$refs[ref]) this.$refs[ref].click();
            },

            onPickFile: function (kind, event) {
                var file = event.target.files && event.target.files[0];
                if (file) this.handleFile(file, kind);
                event.target.value = '';
            },

            onCardDragOver: function (kind) {
                this.pageDragging = false;
                this.cardDragging = kind;
            },

            onCardDragLeave: function (kind) {
                if (this.cardDragging === kind) this.cardDragging = '';
            },

            onCardDrop: function (kind, event) {
                this.cardDragging = '';
                this.pageDragging = false;
                var file = event.dataTransfer.files && event.dataTransfer.files[0];
                if (file) this.handleFile(file, kind);
            },

            hasFiles: function (event) {
                var transfer = event.dataTransfer;
                if (!transfer) return false;
                if (transfer.files && transfer.files.length > 0) return true;

                var types = transfer.types || [];
                for (var i = 0; i < types.length; i++) {
                    if (types[i] === 'Files') return true;
                }
                return false;
            },

            handleWindowDragOver: function (event) {
                if (!this.hasFiles(event)) return;
                // 必须阻止默认行为，否则浏览器会直接打开被拖入的文件
                event.preventDefault();
                this.pageDragging = true;

                // dragleave 在子元素之间移动时也会触发，改用定时器判断是否已离开页面
                if (this._hideTimer) clearTimeout(this._hideTimer);
                this._hideTimer = setTimeout(function () {
                    this.pageDragging = false;
                }.bind(this), 400);
            },

            handleWindowDrop: function (event) {
                if (this._hideTimer) clearTimeout(this._hideTimer);
                this.pageDragging = false;
                if (!this.hasFiles(event)) return;
                event.preventDefault();

                var file = event.dataTransfer.files && event.dataTransfer.files[0];
                if (file) this.handleFile(file, '');
            },

            /* ---------------- 解析 ---------------- */

            /** 未指定类型时，按表头自动识别是美术库还是高考库 */
            guessKind: function (buffer) {
                try {
                    parseWorkbookBuffer(buffer);
                    return 'art';
                } catch (err) {
                    // 不是美术库，继续判断
                }
                try {
                    if (global.PhoneRegister) {
                        global.PhoneRegister.parseGaokaoBuffer(buffer);
                        return 'gk';
                    }
                } catch (err) {
                    // 也不是高考库
                }
                return '';
            },

            handleFile: function (file, kind) {
                if (!/\.(xlsx|xls)$/i.test(file.name || '')) {
                    this.fail(kind, '仅支持 .xlsx / .xls 文件。');
                    return;
                }

                var self = this;
                var type = kind || '';
                if (type === 'gk') {
                    this.gkState = 'reading';
                    this.gkFileName = file.name;
                    this.gkErrorMsg = '';
                } else if (type === 'art') {
                    this.fileState = 'reading';
                    this.fileName = file.name;
                    this.errorMsg = '';
                }

                var reader = new FileReader();
                reader.onload = function () {
                    try {
                        var detected = type || self.guessKind(reader.result);
                        if (!detected) {
                            self.showToast('无法识别该文件：既不是美术报名库（ZKZH/XM/KC），也不是高考报名库（BMXH_/BJ_/BMD_）。', 'error');
                            return;
                        }
                        if (detected === 'art') {
                            self.loadArt(reader.result, file.name);
                        } else {
                            self.loadGaokao(reader.result, file.name);
                        }
                    } catch (err) {
                        self.fail(type, err && err.message ? err.message : '文件解析失败，请检查文件内容。');
                    }
                };
                reader.onerror = function () {
                    self.fail(type, '文件读取失败，请重新选择。');
                };
                reader.readAsArrayBuffer(file);
            },

            loadArt: function (buffer, name) {
                var parsed = parseWorkbookBuffer(buffer);
                var built = buildPages(parsed.records);

                this.records = parsed.records;
                this.pages = built.pages;
                this.warnings = collectWarnings(parsed.records).concat(built.warnings);
                this.fileState = 'ready';
                this.fileName = name;
                this.errorMsg = '';

                this.refreshPhoneGroups();
            },

            loadGaokao: function (buffer, name) {
                var parsed = global.PhoneRegister.parseGaokaoBuffer(buffer);

                this.gkRecords = parsed.records;
                this.gkState = 'ready';
                this.gkFileName = name;
                this.gkErrorMsg = '';

                this.refreshPhoneGroups();
            },

            fail: function (kind, message) {
                if (kind === 'gk') {
                    this.gkState = 'error';
                    this.gkErrorMsg = message;
                    this.gkRecords = [];
                } else {
                    this.fileState = 'error';
                    this.errorMsg = message;
                    this.records = [];
                    this.pages = [];
                    this.warnings = [];
                    // 主表失效时退回功能选择，避免停留在空的承诺书页
                    if (this.view === 'pledge') this.view = 'overview';
                }
                this.refreshPhoneGroups();
            },

            /* ---------------- 手机收缴登记表 ---------------- */

            refreshBmdMap: function (silent) {
                var PR = global.PhoneRegister;
                if (!PR) return;

                var result = PR.readBmdMap();
                this.bmdMap = result.map;
                this.bmdMapCount = result.count;
                this.refreshPhoneGroups();

                if (!silent) {
                    this.showToast(result.available
                        ? '已读取 ' + result.count + ' 条报名点映射'
                        : '无法读取本机映射，浏览器可能禁用了本地存储', result.available ? 'success' : 'error');
                }
            },

            refreshPhoneGroups: function () {
                var PR = global.PhoneRegister;
                if (!PR || !this.records.length || !this.gkRecords.length) {
                    this.phoneGroups = [];
                    this.phoneWarnings = [];
                    this.phoneMatched = 0;
                    this.phoneUnmatched = 0;
                    this.phoneUnmatchedList = [];
                    return;
                }

                var result = PR.groupByBmd(this.records, this.gkRecords, this.bmdMap);
                var warnings = result.warnings.slice();

                var missingNames = result.groups.filter(function (group) {
                    return !group.name;
                }).map(function (group) {
                    return group.code;
                });
                if (missingNames.length) {
                    warnings.push('报名点 ' + missingNames.join('、')
                        + ' 尚未在映射表中配置名称：表内只显示报名点代码，文件名中的名称为「未配置名称」。');
                }

                this.phoneGroups = result.groups;
                this.phoneWarnings = warnings;
                this.phoneMatched = result.matched;
                this.phoneUnmatched = result.unmatched.length;
                this.phoneUnmatchedList = result.unmatched.slice(0, 200);
            },

            getTemplateBuffer: function () {
                if (!this._templatePromise) {
                    var b64 = global.ART_PHONE_TEMPLATE_B64;
                    if (!b64) {
                        return Promise.reject(new Error('未找到内嵌的手机收缴登记表模板（template-20xx.js）。'));
                    }
                    this._templatePromise = Promise.resolve(base64ToArrayBuffer(b64));
                }
                return this._templatePromise;
            },

            downloadGroup: function (group) {
                if (this.downloadingCode) return;

                var self = this;
                this.downloadingCode = group.code;

                this.getTemplateBuffer().then(function (template) {
                    return global.PhoneRegister.buildWorkbook(template, {year: self.examYear, group: group});
                }).then(function (buffer) {
                    self.saveBlob(new Blob([buffer], {type: XLSX_MIME}),
                        global.PhoneRegister.buildFileName(self.examYear, group));
                    self.showToast('已生成：' + (group.name || group.code), 'success');
                }).catch(function (err) {
                    self.showToast(err && err.message ? err.message : '生成 Excel 失败', 'error');
                }).then(function () {
                    self.downloadingCode = '';
                });
            },

            downloadAllZip: function () {
                if (this.zipping || !this.phoneGroups.length) return;

                var self = this;
                var groups = this.phoneGroups.slice();
                var files = [];

                this.zipping = true;
                this.zipProgress = '正在生成…';

                this.getTemplateBuffer().then(function (template) {
                    // 逐个生成，避免同时占用过多内存
                    var chain = Promise.resolve();
                    groups.forEach(function (group, index) {
                        chain = chain.then(function () {
                            self.zipProgress = '正在生成 ' + (index + 1) + '/' + groups.length + '…';
                            return global.PhoneRegister.buildWorkbook(template, {
                                year: self.examYear,
                                group: group
                            }).then(function (buffer) {
                                files.push({
                                    name: global.PhoneRegister.buildFileName(self.examYear, group),
                                    buffer: buffer
                                });
                            });
                        });
                    });
                    return chain;
                }).then(function () {
                    self.zipProgress = '正在打包…';
                    return global.PhoneRegister.buildZip(files, 'blob');
                }).then(function (blob) {
                    self.saveBlob(blob, self.examYear + '年浙江美术统考手机收缴登记-按报名点.zip');
                    self.showToast('已打包 ' + files.length + ' 个文件', 'success');
                }).catch(function (err) {
                    self.showToast(err && err.message ? err.message : '打包失败', 'error');
                }).then(function () {
                    self.zipping = false;
                    self.zipProgress = '';
                });
            },

            saveBlob: function (blob, fileName) {
                var url = URL.createObjectURL(blob);
                var link = document.createElement('a');
                link.href = url;
                link.download = fileName;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                setTimeout(function () {
                    URL.revokeObjectURL(url);
                }, 3000);
            },

            showToast: function (text, type) {
                this.toastText = text;
                this.toastType = type || 'info';
                if (this._toastTimer) clearTimeout(this._toastTimer);
                this._toastTimer = setTimeout(function () {
                    this.toastText = '';
                }.bind(this), 3000);
            },

            /* ---------------- 打印与重置 ---------------- */

            doPrint: function () {
                if (!this.pages.length) return;
                window.print();
            },

            reset: function () {
                this.view = 'overview';

                this.fileState = 'idle';
                this.fileName = '';
                this.errorMsg = '';
                this.records = [];
                this.pages = [];
                this.warnings = [];

                this.gkState = 'idle';
                this.gkFileName = '';
                this.gkErrorMsg = '';
                this.gkRecords = [];

                this.phoneGroups = [];
                this.phoneWarnings = [];
                this.phoneMatched = 0;
                this.phoneUnmatched = 0;
                this.phoneUnmatchedList = [];

                if (this.$refs.artPicker) this.$refs.artPicker.value = '';
                if (this.$refs.gkPicker) this.$refs.gkPicker.value = '';
            }
        }
    });

    // 暴露实例，方便在控制台核对数据（不影响正常使用）
    global.__artApp = app.mount('#app');

})(typeof window !== 'undefined' ? window : globalThis);
