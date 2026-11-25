/**
 * 1 月学选考 · 考场编排校验
 *
 * 文件 1（报名信息）：BMXH（报名序号）、SFZH、XM、BMD、KM_XX（科目_xx）、WYYZ、KD（考点）、WYKD（外语考点）
 * 文件 2（考务数据）：BMXH、XM、BMD、KD_XX（考点_xx）、KSZL_XX（考试种类_xx）、WYYZ
 *
 * 比对规则：
 *   1. 以 BMXH 关联两个文件；
 *   2. 文件 1 的 KD 与文件 2 的每个 KD_XX（xx ≠ yy）逐科目比对；
 *   3. 文件 1 的 WYKD 与文件 2 的 KD_YY 比对（外语科目）；
 *   4. 比对时去掉开头的 0（0012 与 12 视为相同），全角/半角空格与大小写差异忽略；
 *   5. 只比对考生实际报考的科目（文件 1 的 KM_XX 非空；外语另认 WYKD 非空）。
 *
 * 所有解析均在浏览器本地完成，不发起网络请求，也不写入本地存储。
 */
(function (global) {
    'use strict';

    /* ------------------------------------------------------------------ */
    /* 常量                                                                */
    /* ------------------------------------------------------------------ */

    /**
     * 科目定义：
     *   key       内部标识
     *   name      显示名
     *   kmAliases 报名信息文件中的科目列（默认 KM_<KEY>）
     *   kdAliases 考务数据文件中的考点列（默认 KD_<KEY>）
     * 政治在考务数据文件中的列名为 KD_ZZ（考试种类写作 KD_ZS 的历史格式也兼容）。
     */
    var SUBJECTS = [
        {key: 'yw', name: '语文'},
        {key: 'sx', name: '数学'},
        {key: 'yy', name: '外语'},
        {key: 'sz', name: '政治', kmAliases: ['KM_SZ', 'KM_ZZ'], kdAliases: ['KD_ZZ', 'KD_SZ']},
        {key: 'ls', name: '历史'},
        {key: 'dl', name: '地理'},
        {key: 'wl', name: '物理'},
        {key: 'hx', name: '化学'},
        {key: 'sw', name: '生物'},
        {key: 'js', name: '技术'}
    ];

    /** 报名信息文件的科目列名（按优先级） */
    function kmHeaders(subject) {
        return subject.kmAliases || ['KM_' + subject.key.toUpperCase()];
    }

    /** 考务数据文件的考点列名（按优先级） */
    function kdHeaders(subject) {
        return subject.kdAliases || ['KD_' + subject.key.toUpperCase()];
    }

    /** 对外展示的考点列名，如 kd_zz */
    function kdHeaderName(subject) {
        return kdHeaders(subject)[0];
    }

    function pickColumn(map, headers) {
        for (var i = 0; i < headers.length; i++) {
            if (map[headers[i]] !== undefined) return map[headers[i]];
        }
        return -1;
    }

    var HEADER_SCAN = 20;
    var DETAIL_LIMIT = 300;

    /* ------------------------------------------------------------------ */
    /* 文本处理                                                            */
    /* ------------------------------------------------------------------ */

    function cellText(value) {
        if (value === null || value === undefined) return '';
        if (typeof value === 'number') return String(value);
        return String(value).replace(/^[\s\u3000\uFEFF']+|[\s\u3000\uFEFF]+$/g, '');
    }

    /** 表头统一：去空格、去末尾下划线、转大写（表头不区分大小写） */
    function normalizeHeader(value) {
        return cellText(value)
            .replace(/[\s\u3000\uFEFF]/g, '')
            .replace(/_+$/, '')
            .toUpperCase();
    }

    /** 对比口径：去空格、纯数字去掉开头的 0、其余转大写 */
    function normalizeCode(value) {
        var text = cellText(value).replace(/[\s\u3000\uFEFF]/g, '');
        if (!text) return '';
        if (/^\d+$/.test(text)) return text.replace(/^0+(?=\d)/, '');
        return text.toUpperCase();
    }

    /* ------------------------------------------------------------------ */
    /* 读取 Excel                                                          */
    /* ------------------------------------------------------------------ */

    function readSheets(buffer) {
        var xlsx = global.XLSX;
        if (!xlsx) throw new Error('未加载 SheetJS 解析库，请检查 lib/sheetjs 路径是否正确。');

        var workbook = xlsx.read(new Uint8Array(buffer), {type: 'array'});
        return (workbook.SheetNames || []).map(function (name) {
            var sheet = workbook.Sheets[name];
            return {
                name: name,
                aoa: xlsx.utils.sheet_to_json(sheet, {
                    header: 1,
                    raw: true,
                    defval: '',
                    blankrows: false
                })
            };
        });
    }

    /** 表头行 → {标准化表头: 列下标}，重复表头只取第一列 */
    function buildIndex(row) {
        var map = {};
        (row || []).forEach(function (cell, i) {
            var key = normalizeHeader(cell);
            if (key && map[key] === undefined) map[key] = i;
        });
        return map;
    }

    /** 在前 20 行里找满足条件的表头行 */
    function findHeader(aoa, predicate) {
        for (var i = 0; i < aoa.length && i < HEADER_SCAN; i++) {
            var map = buildIndex(aoa[i] || []);
            if (predicate(map)) return {index: i, map: map, row: aoa[i]};
        }
        return null;
    }

    function columnOf(map, name) {
        return map[name] === undefined ? -1 : map[name];
    }

    function valueAt(line, index) {
        return index < 0 ? '' : cellText(line[index]);
    }

    /* ------------------------------------------------------------------ */
    /* 解析文件 1：报名信息                                                */
    /* ------------------------------------------------------------------ */

    function parseFile1(buffer) {
        var sheets = readSheets(buffer);
        var fallback = '';

        for (var s = 0; s < sheets.length; s++) {
            var aoa = sheets[s].aoa;
            var head = findHeader(aoa, function (map) {
                if (map.BMXH === undefined) return false;
                return hasAnyColumn(map, 'KM_');
            });
            if (!head) {
                if (!fallback && aoa.length && aoa[0]) {
                    fallback = (aoa[0] || []).map(cellText).filter(Boolean).slice(0, 12).join('、');
                }
                continue;
            }

            var idx = {
                bmxh: columnOf(head.map, 'BMXH'),
                xm: columnOf(head.map, 'XM'),
                bmd: columnOf(head.map, 'BMD'),
                kd: columnOf(head.map, 'KD'),
                wykd: columnOf(head.map, 'WYKD'),
                wyyz: columnOf(head.map, 'WYYZ')
            };

            var kmIndex = {};
            var kmColumns = [];
            SUBJECTS.forEach(function (subject) {
                var col = pickColumn(head.map, kmHeaders(subject));
                if (col >= 0) {
                    kmIndex[subject.key] = col;
                    kmColumns.push(subject.key);
                }
            });

            var records = [];
            var seen = {};
            var duplicates = 0;

            for (var r = head.index + 1; r < aoa.length; r++) {
                var line = aoa[r] || [];
                var bmxh = valueAt(line, idx.bmxh);
                if (!bmxh) continue;
                if (seen[bmxh]) {
                    duplicates++;
                    continue;
                }
                seen[bmxh] = true;

                var km = {};
                SUBJECTS.forEach(function (subject) {
                    var col = kmIndex[subject.key];
                    km[subject.key] = col === undefined ? '' : valueAt(line, col);
                });

                records.push({
                    bmxh: bmxh,
                    xm: valueAt(line, idx.xm),
                    bmd: valueAt(line, idx.bmd),
                    kd: valueAt(line, idx.kd),
                    wykd: valueAt(line, idx.wykd),
                    wyyz: valueAt(line, idx.wyyz),
                    km: km
                });
            }

            if (!records.length) continue;

            return {
                sheetName: sheets[s].name,
                records: records,
                duplicates: duplicates,
                hasKd: idx.kd >= 0,
                hasWykd: idx.wykd >= 0,
                kmColumns: kmColumns
            };
        }

        throw new Error('未识别出「报名信息」表头：需要 BMXH 列与 KM_* 科目列（如 KM_YW、KM_SX）。'
            + (fallback ? '当前文件首个表头行为：' + fallback : ''));
    }

    /* ------------------------------------------------------------------ */
    /* 解析文件 2：考务数据                                                */
    /* ------------------------------------------------------------------ */

    function parseFile2(buffer) {
        var sheets = readSheets(buffer);
        var fallback = '';

        for (var s = 0; s < sheets.length; s++) {
            var aoa = sheets[s].aoa;
            var head = findHeader(aoa, function (map) {
                if (map.BMXH === undefined) return false;
                return hasAnyColumn(map, 'KD_');
            });
            if (!head) {
                if (!fallback && aoa.length && aoa[0]) {
                    fallback = (aoa[0] || []).map(cellText).filter(Boolean).slice(0, 12).join('、');
                }
                continue;
            }

            var idx = {
                bmxh: columnOf(head.map, 'BMXH'),
                xm: columnOf(head.map, 'XM'),
                bmd: columnOf(head.map, 'BMD')
            };

            var kdIndex = {};
            var subjectColumns = [];
            SUBJECTS.forEach(function (subject) {
                var col = pickColumn(head.map, kdHeaders(subject));
                if (col >= 0) {
                    kdIndex[subject.key] = col;
                    subjectColumns.push(subject.key);
                }
            });

            var records = [];
            var seen = {};
            var duplicates = 0;

            for (var r = head.index + 1; r < aoa.length; r++) {
                var line = aoa[r] || [];
                var bmxh = valueAt(line, idx.bmxh);
                if (!bmxh) continue;
                if (seen[bmxh]) {
                    duplicates++;
                    continue;
                }
                seen[bmxh] = true;

                var kd = {};
                SUBJECTS.forEach(function (subject) {
                    var col = kdIndex[subject.key];
                    kd[subject.key] = col === undefined ? '' : valueAt(line, col);
                });

                records.push({
                    bmxh: bmxh,
                    xm: valueAt(line, idx.xm),
                    bmd: valueAt(line, idx.bmd),
                    kd: kd
                });
            }

            if (!records.length) continue;

            return {
                sheetName: sheets[s].name,
                records: records,
                duplicates: duplicates,
                subjectColumns: subjectColumns
            };
        }

        throw new Error('未识别出「考务数据」表头：需要 BMXH 列与 KD_* 科目列（如 KD_YW、KD_SX）。'
            + (fallback ? '当前文件首个表头行为：' + fallback : ''));
    }

    function hasAnyColumn(map, prefix) {
        for (var i = 0; i < SUBJECTS.length; i++) {
            if (map[prefix + SUBJECTS[i].key.toUpperCase()] !== undefined) return true;
        }
        return false;
    }

    /* ------------------------------------------------------------------ */
    /* 比对                                                                */
    /* ------------------------------------------------------------------ */

    /**
     * @returns {{rows:Array, stats:Object, warnings:Array, onlyIn2:Array}}
     */
    function compareArrange(file1, file2) {
        var lookup = {};
        var duplicated = {};
        var dupIn2 = 0;

        (file2.records || []).forEach(function (record) {
            if (lookup[record.bmxh]) {
                if (!duplicated[record.bmxh]) {
                    duplicated[record.bmxh] = true;
                    dupIn2++;
                }
                return;
            }
            lookup[record.bmxh] = record;
        });

        var firstSet = {};
        (file1.records || []).forEach(function (record) {
            firstSet[record.bmxh] = true;
        });

        var rows = [];
        var stats = {total: 0, same: 0, diff: 0, missing: 0, diffCells: 0, checkedCells: 0};

        (file1.records || []).forEach(function (record) {
            var other = lookup[record.bmxh];
            var cells = {};
            var diffCount = 0;

            // 外语语种为 0 表示未选择外语科目，不校验外语考点
            var noForeignLanguage = normalizeCode(record.wyyz) === '0';

            SUBJECTS.forEach(function (subject) {
                // 是否报考：普通科目看 KM_XX，外语额外认 WYKD
                var registered = subject.key === 'yy'
                    ? (!noForeignLanguage && !!(record.km.yy || normalizeCode(record.wykd)))
                    : !!record.km[subject.key];

                if (!registered) {
                    cells[subject.key] = {state: 'skip'};
                    return;
                }
                if (file2.subjectColumns.indexOf(subject.key) < 0) {
                    cells[subject.key] = {state: 'nocol'};
                    return;
                }
                if (!other) {
                    cells[subject.key] = {state: 'missing'};
                    return;
                }

                var value1 = subject.key === 'yy' ? record.wykd : record.kd;
                var value2 = other.kd[subject.key] || '';

                stats.checkedCells++;
                if (normalizeCode(value1) === normalizeCode(value2)) {
                    cells[subject.key] = {state: 'same', a: value1, b: value2};
                } else {
                    diffCount++;
                    stats.diffCells++;
                    cells[subject.key] = {state: 'diff', a: value1, b: value2};
                }
            });

            var status = diffCount > 0 ? 'diff' : (other ? 'same' : 'missing');
            stats.total++;
            stats[status]++;

            rows.push({
                bmxh: record.bmxh,
                xm: record.xm,
                bmd: record.bmd,
                cells: cells,
                diffCount: diffCount,
                status: status
            });
        });

        var onlyIn2 = (file2.records || []).filter(function (record) {
            return !firstSet[record.bmxh];
        }).map(function (record) {
            return {bmxh: record.bmxh, xm: record.xm};
        });

        var warnings = [];
        if (file1.duplicates) {
            warnings.push('报名信息文件中有 ' + file1.duplicates + ' 条重复的报名序号，已按首条记录比对。');
        }
        if (dupIn2) {
            warnings.push('考务数据文件中有 ' + dupIn2 + ' 个重复的报名序号，已按首条记录比对。');
        }
        if (!file1.hasKd) {
            warnings.push('报名信息文件中没有 KD（考点）列，普通科目无法比对。');
        }
        if (!file1.hasWykd) {
            warnings.push('报名信息文件中没有 WYKD（外语考点）列，外语科目无法比对。');
        }

        var missingColumns = SUBJECTS.filter(function (subject) {
            return file2.subjectColumns.indexOf(subject.key) < 0;
        }).map(function (subject) {
            return kdHeaderName(subject).toLowerCase();
        });
        if (missingColumns.length) {
            warnings.push('考务数据文件缺少列：' + missingColumns.join('、') + '，对应科目未参与比对。');
        }
        if (stats.missing) {
            warnings.push('有 ' + stats.missing + ' 名考生的报名序号在考务数据文件中找不到，已标为「未匹配」。');
        }
        if (onlyIn2.length) {
            warnings.push('考务数据文件中有 ' + onlyIn2.length + ' 个报名序号未出现在报名信息文件中。');
        }

        return {
            rows: rows,
            stats: stats,
            warnings: warnings,
            onlyIn2: onlyIn2.slice(0, DETAIL_LIMIT)
        };
    }

    /* 供调试与本地校验使用 */
    global.KcArrangeCheck = {
        SUBJECTS: SUBJECTS,
        normalizeCode: normalizeCode,
        normalizeHeader: normalizeHeader,
        parseFile1: parseFile1,
        parseFile2: parseFile2,
        compareArrange: compareArrange,
        kdHeaderName: kdHeaderName,
        kmHeaderName: function (subject) {
            return kmHeaders(subject)[0];
        }
    };

    if (typeof global.Vue === 'undefined') return; // Node/测试环境下只暴露纯函数

    /* ------------------------------------------------------------------ */
    /* 页面逻辑                                                            */
    /* ------------------------------------------------------------------ */

    var app = global.Vue.createApp({
        data: function () {
            return {
                subjects: SUBJECTS,
                files: {
                    f1: {status: 'idle', name: '', error: ''},
                    f2: {status: 'idle', name: '', error: ''}
                },
                parsed1: null,
                parsed2: null,
                dragState: {f1: false, f2: false},

                filter: 'all',
                keyword: '',

                toastText: '',
                toastClass: ''
            };
        },

        computed: {
            result: function () {
                if (!this.parsed1 || !this.parsed2) return null;
                return compareArrange(this.parsed1, this.parsed2);
            },

            stats: function () {
                return this.result ? this.result.stats : {total: 0, same: 0, diff: 0, missing: 0, diffCells: 0};
            },

            warnings: function () {
                return this.result ? this.result.warnings : [];
            },

            onlyIn2: function () {
                return this.result ? this.result.onlyIn2 : [];
            },

            parseWarnings: function () {
                var list = [];
                if (this.parsed1 && !this.parsed1.hasKd) list.push('报名信息文件中没有 KD（考点）列。');
                if (this.parsed1 && !this.parsed1.hasWykd) list.push('报名信息文件中没有 WYKD（外语考点）列。');
                return list;
            },

            /** 编排一致的记录不展示，只保留有问题（不一致 / 未匹配）的考生 */
            problemRows: function () {
                if (!this.result) return [];
                return this.result.rows.filter(function (row) {
                    return row.status !== 'same';
                });
            },

            filteredRows: function () {
                var keyword = this.keyword.toLowerCase();
                var mode = this.filter;

                return this.problemRows.filter(function (row) {
                    if (mode === 'diff' && row.status !== 'diff') return false;
                    if (mode === 'missing' && row.status !== 'missing') return false;
                    if (!keyword) return true;
                    var text = row.bmxh + ' ' + row.xm + ' ' + row.bmd;
                    return text.toLowerCase().indexOf(keyword) >= 0;
                });
            }
        },

        beforeUnmount: function () {
            if (this._toastTimer) clearTimeout(this._toastTimer);
        },

        methods: {
            cardClass: function (kind) {
                var file = this.files[kind];
                return {
                    'is-dragging': this.dragState[kind],
                    'is-ready': file.status === 'ready',
                    'is-error': file.status === 'error'
                };
            },

            columnCode: function (subject) {
                return kdHeaderName(subject).toLowerCase();
            },

            cellDisplay: function (cell) {
                if (!cell) return '';
                if (cell.state === 'same') return '✓';
                if (cell.state === 'diff') return (cell.a || '空') + ' / ' + (cell.b || '空');
                if (cell.state === 'missing') return '未匹配';
                return '—';
            },

            cellClass: function (cell) {
                if (!cell) return '';
                return 'cell-' + cell.state;
            },

            cellTitle: function (cell) {
                if (!cell || (cell.state !== 'same' && cell.state !== 'diff')) return '';
                return '报名信息：' + (cell.a || '空') + '\n考务数据：' + (cell.b || '空');
            },

            pickFile: function (kind) {
                var ref = kind === 'f1' ? 'f1Picker' : 'f2Picker';
                if (this.$refs[ref]) this.$refs[ref].click();
            },

            onPick: function (kind, event) {
                var file = event.target.files && event.target.files[0];
                if (file) this.handleFile(kind, file);
                event.target.value = '';
            },

            onDragOver: function (kind) {
                this.dragState[kind] = true;
            },

            onDragLeave: function (kind) {
                this.dragState[kind] = false;
            },

            onDrop: function (kind, event) {
                this.dragState[kind] = false;
                var file = event.dataTransfer.files && event.dataTransfer.files[0];
                if (file) this.handleFile(kind, file);
            },

            handleFile: function (kind, file) {
                if (!/\.(xlsx|xls)$/i.test(file.name || '')) {
                    this.fail(kind, '仅支持 .xlsx / .xls 文件。');
                    return;
                }

                var self = this;
                this.files[kind].status = 'reading';
                this.files[kind].name = file.name;
                this.files[kind].error = '';

                var reader = new FileReader();
                reader.onload = function () {
                    try {
                        self.load(kind, reader.result, file.name);
                    } catch (err) {
                        self.fail(kind, err && err.message ? err.message : '文件解析失败，请检查文件内容。');
                    }
                };
                reader.onerror = function () {
                    self.fail(kind, '文件读取失败，请重新选择。');
                };
                reader.readAsArrayBuffer(file);
            },

            load: function (kind, buffer, name) {
                var parsed = kind === 'f1' ? parseFile1(buffer) : parseFile2(buffer);

                if (kind === 'f1') this.parsed1 = parsed;
                else this.parsed2 = parsed;

                this.files[kind].status = 'ready';
                this.files[kind].name = name;
                this.files[kind].error = '';
                this.showToast(kind === 'f1'
                    ? '已读取报名信息 ' + parsed.records.length + ' 条'
                    : '已读取考务数据 ' + parsed.records.length + ' 条');
            },

            fail: function (kind, message) {
                this.files[kind].status = 'error';
                this.files[kind].error = message;
                if (kind === 'f1') this.parsed1 = null;
                else this.parsed2 = null;
            },

            reset: function () {
                this.files.f1 = {status: 'idle', name: '', error: ''};
                this.files.f2 = {status: 'idle', name: '', error: ''};
                this.parsed1 = null;
                this.parsed2 = null;
                this.filter = 'all';
                this.keyword = '';
                if (this.$refs.f1Picker) this.$refs.f1Picker.value = '';
                if (this.$refs.f2Picker) this.$refs.f2Picker.value = '';
            },

            showToast: function (text) {
                this.toastText = text;
                this.toastClass = '';
                if (this._toastTimer) clearTimeout(this._toastTimer);
                this._toastTimer = setTimeout(function () {
                    this.toastText = '';
                }.bind(this), 2600);
            }
        }
    });

    global.__kcArrangeApp = app.mount('#app');

})(typeof window !== 'undefined' ? window : globalThis);
