/**
 * 页面交互：上传、映射配置、统计类型选择、结果渲染。
 * 所有业务数据仅在浏览器内存中处理。
 */
(function () {
    "use strict";

    var DN = window.DataNormalizer;
    var ER = window.ExcelReader;

    var STORAGE_KEYS = {
        bmd: "exam-data-statistics:bmd-map:v1",
        county: "exam-data-statistics:county-map:v1"
    };

    var UPLOAD_CARDS = [
        {
            kind: "gaokao",
            title: "高考报名信息",
            extLabel: ".xlsx",
            accept: ".xlsx",
            hint: "将高考报名信息 Excel 文件拖动到此区域"
        },
        {
            kind: "payment",
            title: "高考报名缴费信息",
            extLabel: ".xls / .xlsx",
            accept: ".xls,.xlsx",
            hint: "将高考报名缴费 Excel 文件拖动到此区域"
        },
        {
            kind: "january",
            title: "1月学选考报名信息",
            extLabel: ".xlsx",
            accept: ".xlsx",
            hint: "将 1 月学选考报名信息 Excel 文件拖动到此区域"
        },
        {
            kind: "dandu",
            title: "单独考试报名信息",
            extLabel: ".xlsx",
            accept: ".xlsx",
            hint: "将单独考试报名信息 Excel 文件拖动到此区域"
        }
    ];

    var STAT_TYPES = [
        {
            key: "gaokao",
            title: "高考报名信息统计",
            requireText: "高考报名信息文件",
            desc: "按报名点及县区统计高考报名人数、各术科类别人数与特招生测试人数；若同时上传缴费文件，自动执行缴费交叉过滤。"
        },
        {
            key: "payment",
            title: "高考报名缴费统计",
            requireText: "高考报名缴费信息文件",
            desc: "按报名点统计总人数（普通类去重人数）及各缴费项目的已缴费、未缴费人数。"
        },
        {
            key: "january",
            title: "1月学选考报名统计",
            requireText: "1月学选考报名信息文件",
            desc: "按报名点及县区统计语文学考、数学学考、外语各语种人数及政治、历史、地理、物理、化学、生物、技术的学考/选考人数。"
        },
        {
            key: "dandu",
            title: "单独考试报名统计",
            requireText: "单独考试报名信息文件",
            desc: "按报名点及县区统计总人数与各报考类别（BKKL_）报名人数；报考类别 17 按其他类专业（QTLZY_）细分为 17-1 安全防范、17-2 体育等。"
        }
    ];

    var PARSERS = {
        gaokao: DN.parseGaokao,
        payment: DN.parsePayment,
        january: DN.parseJanuary,
        dandu: DN.parseDandu
    };

    // 工作簿与解析结果放在非响应式存储中，避免大对象被 Vue 代理
    var fileStore = {gaokao: null, payment: null, january: null, dandu: null};
    var saveTimer = null;
    var toastTimer = null;

    function createFileState() {
        return {
            status: "idle",
            fileName: "",
            fileSize: 0,
            errorMessage: "",
            sheetName: "",
            sheetNames: [],
            candidates: [],
            headerRowIndex: null,
            summary: null,
            diagnostics: [],
            distributions: []
        };
    }

    function formatSize(bytes) {
        var size = Number(bytes) || 0;
        if (size < 1024) return size + " B";
        if (size < 1024 * 1024) return (size / 1024).toFixed(1) + " KB";
        return (size / 1024 / 1024).toFixed(2) + " MB";
    }

    function duplicateCodes(rows) {
        var seen = new Set();
        var duplicate = new Set();
        (rows || []).forEach(function (row) {
            var code = String(row.code || "").trim();
            if (!code) return;
            if (seen.has(code)) duplicate.add(code);
            else seen.add(code);
        });
        return Array.from(duplicate);
    }

    function mergeDiagnostics(list) {
        var map = new Map();
        (list || []).forEach(function (item) {
            if (!item) return;
            var entry = map.get(item.code);
            if (!entry) {
                entry = {level: item.level, code: item.code, title: item.title, count: 0, samples: []};
                map.set(item.code, entry);
            }
            entry.count += item.count || 0;
            (item.samples || []).forEach(function (sample) {
                var text = String(sample);
                if (entry.samples.length < 5 && entry.samples.indexOf(text) === -1) entry.samples.push(text);
            });
        });
        return Array.from(map.values());
    }

    var app = Vue.createApp({
        data: function () {
            return {
                uploadCards: UPLOAD_CARDS,
                statTypes: STAT_TYPES,
                files: {
                    gaokao: createFileState(),
                    payment: createFileState(),
                    january: createFileState(),
                    dandu: createFileState()
                },
                dragState: {gaokao: false, payment: false, january: false, dandu: false},
                bmdRows: [],
                countyRows: [],
                // 有映射数据时默认收起，没有数据时展开以便录入
                mappingOpen: {bmd: true, county: true},
                batchInput: {bmd: "", county: ""},
                selection: {gaokao: false, payment: false, january: false, dandu: false},
                hideZeroColumns: {gaokao: false, payment: false, january: false, dandu: false},
                rawResults: [],
                statDiagnostics: [],
                busy: false,
                stale: false,
                statusText: "等待上传文件",
                generatedAt: "",
                toast: null
            };
        },

        computed: {
            canGenerate: function () {
                var self = this;
                return Object.keys(this.selection).some(function (key) {
                    return self.selection[key] && self.isTypeAvailable(key);
                });
            },

            results: function () {
                var self = this;
                return this.rawResults.map(function (table) {
                    return self.visibleTable(table);
                });
            },

            duplicateBmdCodes: function () {
                return duplicateCodes(this.bmdRows);
            },

            duplicateCountyCodes: function () {
                return duplicateCodes(this.countyRows);
            },

            uploadDiagnostics: function () {
                var list = [];
                var self = this;
                Object.keys(this.files).forEach(function (kind) {
                    var state = self.files[kind];
                    if (state.status === "ready") list = list.concat(state.diagnostics || []);
                });
                return list;
            },

            allDiagnostics: function () {
                var list = this.uploadDiagnostics.concat(this.statDiagnostics);

                if (this.duplicateBmdCodes.length) {
                    list.push({
                        level: "warning",
                        code: "MAP_BMD_DUPLICATE",
                        title: "报名点映射存在重复代码，仅第一条生效：" + this.duplicateBmdCodes.join("、"),
                        count: this.duplicateBmdCodes.length,
                        samples: []
                    });
                }
                if (this.duplicateCountyCodes.length) {
                    list.push({
                        level: "warning",
                        code: "MAP_COUNTY_DUPLICATE",
                        title: "县区映射存在重复代码，仅第一条生效：" + this.duplicateCountyCodes.join("、"),
                        count: this.duplicateCountyCodes.length,
                        samples: []
                    });
                }

                return mergeDiagnostics(list).sort(function (a, b) {
                    if (a.level !== b.level) return a.level === "error" ? -1 : 1;
                    return b.count - a.count;
                });
            },

            errorDiagnostics: function () {
                return this.allDiagnostics.filter(function (item) {
                    return item.level === "error";
                });
            },

            warningDiagnostics: function () {
                return this.allDiagnostics.filter(function (item) {
                    return item.level !== "error";
                });
            },

            distributions: function () {
                var list = [];
                var self = this;
                Object.keys(this.files).forEach(function (kind) {
                    var state = self.files[kind];
                    if (state.status !== "ready") return;
                    var card = UPLOAD_CARDS.filter(function (item) {
                        return item.kind === kind;
                    })[0];
                    (state.distributions || []).forEach(function (dist) {
                        list.push({
                            group: card ? card.title : kind,
                            title: dist.title,
                            items: dist.items
                        });
                    });
                });
                return list;
            },

            hasAnyDiagnostics: function () {
                return this.allDiagnostics.length > 0;
            }
        },

        watch: {
            bmdRows: {
                deep: true,
                handler: function () {
                    this.persistMappings();
                    this.markStale();
                }
            },
            countyRows: {
                deep: true,
                handler: function () {
                    this.persistMappings();
                    this.markStale();
                }
            }
        },

        mounted: function () {
            this.loadMappings();
            // 已有映射数据时默认收起，暂无数据时展开
            this.mappingOpen = {
                bmd: this.bmdRows.length === 0,
                county: this.countyRows.length === 0
            };
            window.addEventListener("dragover", this.preventWindowDrop);
            window.addEventListener("drop", this.preventWindowDrop);
        },

        beforeUnmount: function () {
            window.removeEventListener("dragover", this.preventWindowDrop);
            window.removeEventListener("drop", this.preventWindowDrop);
        },

        methods: {
            /* ---------------- 通用 ---------------- */
            preventWindowDrop: function (event) {
                event.preventDefault();
            },

            markStale: function () {
                if (this.rawResults.length) this.stale = true;
            },

            /**
             * 按统计类型的“隐藏合计为 0 的列”选项，生成用于展示的表格副本。
             * 代码/名称列始终保留；字段缺失（不可用）列也保留，避免误以为数据为 0。
             */
            visibleTable: function (table) {
                var hideZero = !!(table && table.statType && this.hideZeroColumns[table.statType]);
                if (!hideZero) return table;

                var hidden = [];
                var columns = table.columns.filter(function (column) {
                    if (column.type !== "number") return true;
                    if (column.available === false) return true;
                    var total = table.totalRow ? table.totalRow[column.key] : null;
                    var isZero = total === 0 || total === "0";
                    if (isZero) hidden.push(column.label);
                    return !isZero;
                });

                if (!hidden.length) return table;

                var copy = Object.assign({}, table, {columns: columns, hiddenColumns: hidden});
                if (Vue.markRaw) Vue.markRaw(copy);
                return copy;
            },

            showToast: function (message, type) {
                var self = this;
                this.toast = {message: message, type: type || "info"};
                window.clearTimeout(toastTimer);
                toastTimer = window.setTimeout(function () {
                    self.toast = null;
                }, 3200);
            },

            formatSize: formatSize,

            formatCell: function (row, column) {
                var value = row ? row[column.key] : null;
                if (value === null || value === undefined || value === "") return "—";
                if (column.type === "number") {
                    var number = Number(value);
                    return Number.isFinite(number) ? String(Math.round(number)) : "—";
                }
                return String(value);
            },

            isTypeAvailable: function (kind) {
                return this.files[kind].status === "ready" && !!(fileStore[kind] && fileStore[kind].parseResult);
            },

            /* ---------------- 复制为 Tab 分隔文本 ---------------- */
            tsvCell: function (row, column) {
                var value = row ? row[column.key] : null;
                if (value === null || value === undefined) return "";
                if (column.type === "number") {
                    var number = Number(value);
                    return Number.isFinite(number) ? String(Math.round(number)) : "";
                }
                // 单元格内不允许出现制表符或换行，避免粘贴到 Excel 后错列
                return String(value).replace(/[\t\r\n]+/g, " ").trim();
            },

            tableToTsv: function (table) {
                var lines = [];
                lines.push(table.columns.map(function (column) {
                    return String(column.label).replace(/[\t\r\n]+/g, " ").trim();
                }).join("\t"));

                table.rows.forEach(function (row) {
                    lines.push(table.columns.map(function (column) {
                        return this.tsvCell(row, column);
                    }, this).join("\t"));
                }, this);

                if (table.totalRow) {
                    lines.push(table.columns.map(function (column) {
                        return this.tsvCell(table.totalRow, column);
                    }, this).join("\t"));
                }

                return lines.join("\r\n");
            },

            copyText: async function (text) {
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    try {
                        // 剪贴板权限可能被挂起（例如等待用户授权），加超时避免按钮无响应
                        var result = await Promise.race([
                            navigator.clipboard.writeText(text).then(function () {
                                return true;
                            }, function () {
                                return false;
                            }),
                            new Promise(function (resolve) {
                                window.setTimeout(function () {
                                    resolve(null);
                                }, 1200);
                            })
                        ]);
                        if (result === true) return true;
                    } catch (err) {
                        // 继续尝试兼容方案
                    }
                }

                try {
                    var textarea = document.createElement("textarea");
                    textarea.value = text;
                    textarea.setAttribute("readonly", "");
                    textarea.style.position = "fixed";
                    textarea.style.top = "-1000px";
                    textarea.style.opacity = "0";
                    document.body.appendChild(textarea);
                    textarea.select();
                    var copied = document.execCommand("copy");
                    document.body.removeChild(textarea);
                    return copied;
                } catch (err) {
                    return false;
                }
            },

            copyTable: async function (table) {
                if (!table || !table.columns.length) return;
                var text = this.tableToTsv(table);
                var copied = await this.copyText(text);
                if (copied) {
                    this.showToast("已复制「" + table.title + "」（Tab 分隔，可直接粘贴到 Excel）", "success");
                } else {
                    this.showToast("复制失败：浏览器未授权访问剪贴板，请手动选中表格内容复制", "error");
                }
            },

            /* ---------------- 文件上传 ---------------- */
            pickFile: function (kind) {
                var refs = this.$refs["input-" + kind];
                var input = Array.isArray(refs) ? refs[0] : refs;
                if (input) input.click();
            },

            onFileChange: function (kind, event) {
                var input = event.target;
                var file = input.files && input.files[0];
                input.value = "";
                if (file) this.processFile(kind, file);
            },

            onDragOver: function (kind, event) {
                if (event) event.preventDefault();
                this.dragState[kind] = true;
            },

            onDragEnter: function (kind, event) {
                if (event) event.preventDefault();
                this.dragState[kind] = true;
            },

            onDragLeave: function (kind, event) {
                if (event && event.currentTarget && event.relatedTarget && event.currentTarget.contains(event.relatedTarget)) return;
                this.dragState[kind] = false;
            },

            onDrop: function (kind, event) {
                this.dragState[kind] = false;
                var files = event && event.dataTransfer ? event.dataTransfer.files : null;
                if (!files || !files.length) {
                    this.showToast("未获取到拖入的文件，请重新拖入或点击“选择文件”", "warning");
                    return;
                }
                if (files.length > 1) {
                    this.showToast("一次只能读取一个文件，已使用第一个文件", "warning");
                }
                this.processFile(kind, files[0]);
            },

            processFile: async function (kind, file) {
                var card = UPLOAD_CARDS.filter(function (item) {
                    return item.kind === kind;
                })[0];
                var state = this.files[kind];

                var extMatch = file.name.match(/\.[^.]+$/);
                var ext = extMatch ? extMatch[0].toLowerCase() : "";
                var allowed = card.accept.split(",").map(function (item) {
                    return item.trim().toLowerCase();
                }).filter(Boolean);

                if (allowed.indexOf(ext) === -1) {
                    Object.assign(state, createFileState(), {
                        status: "error",
                        fileName: file.name,
                        fileSize: file.size,
                        errorMessage: "文件类型不匹配：该区域只接受 " + allowed.join("、") + " 文件，当前文件为 " + (ext || "未知类型") + "，请拖入对应区域。"
                    });
                    fileStore[kind] = null;
                    this.selection[kind] = false;
                    this.markStale();
                    return;
                }

                Object.assign(state, createFileState(), {
                    status: "reading",
                    fileName: file.name,
                    fileSize: file.size
                });
                fileStore[kind] = null;
                this.selection[kind] = false;

                try {
                    var result = await ER.readWorkbook(file);
                    var candidates = ER.detectCandidates(result.workbook, ER.SCHEMAS[kind]);
                    state.sheetNames = result.sheetNames;

                    if (!candidates.length) {
                        state.status = "error";
                        state.errorMessage = "未识别到 " + card.title + " 所需的表头（缺少必要字段）。已扫描工作表：" +
                            result.sheetNames.join("、") + "。请确认文件类型是否放错区域。";
                        this.markStale();
                        return;
                    }

                    fileStore[kind] = {workbook: result.workbook, parseResult: null, candidates: candidates};
                    state.candidates = candidates.map(function (candidate) {
                        return {
                            sheetName: candidate.sheetName,
                            headerRowIndex: candidate.headerRowIndex,
                            label: candidate.sheetName + "（表头行 " + (candidate.headerRowIndex + 1) + "）"
                        };
                    });

                    this.applyCandidate(kind, candidates[0]);
                    state.status = "ready";
                    this.statusText = card.title + " 读取完成，请选择统计类型";
                } catch (err) {
                    fileStore[kind] = null;
                    state.status = "error";
                    state.errorMessage = err && err.message ? err.message : "文件读取失败";
                } finally {
                    this.markStale();
                }
            },

            applyCandidate: function (kind, candidate) {
                var store = fileStore[kind];
                if (!store) return;
                var matrix = ER.readMatrix(store.workbook, candidate.sheetName);
                var slice = ER.sliceRows(matrix, candidate.headerRowIndex);
                var parsed = PARSERS[kind]({
                    columns: candidate.columns,
                    dataRows: slice.dataRows,
                    firstDataRowNumber: slice.firstDataRowNumber
                });
                store.parseResult = parsed;

                var state = this.files[kind];
                state.sheetName = candidate.sheetName;
                state.headerRowIndex = candidate.headerRowIndex;
                state.summary = parsed.summary;
                state.diagnostics = parsed.diagnostics;
                state.distributions = DN.formatDistributions(kind, parsed.distributions);
            },

            onSheetChange: function (kind, event) {
                var sheetName = event.target.value;
                var store = fileStore[kind];
                if (!store) return;
                var candidate = (store.candidates || []).filter(function (item) {
                    return item.sheetName === sheetName;
                })[0];
                if (!candidate) return;
                this.applyCandidate(kind, candidate);
                this.statusText = "已切换工作表：" + sheetName;
                this.markStale();
            },

            removeFile: function (kind) {
                fileStore[kind] = null;
                this.files[kind] = createFileState();
                this.selection[kind] = false;
                this.markStale();
                this.statusText = "已移除文件";
            },

            summaryItems: function (kind) {
                var state = this.files[kind];
                var summary = state.summary;
                if (!summary) return [];

                var items = [
                    {label: "文件名", value: state.fileName},
                    {label: "文件大小", value: formatSize(state.fileSize)},
                    {label: "工作表", value: state.sheetName},
                    {label: "表头行", value: state.headerRowIndex === null ? "—" : "第 " + (state.headerRowIndex + 1) + " 行"},
                    {label: "有效数据行数", value: summary.totalRows}
                ];

                if (kind === "payment") {
                    items.push({label: "已识别报名点数", value: summary.bmdCount});
                    items.push({label: "无法解析缴款单号行数", value: summary.orderNoFail});
                } else if (kind === "dandu") {
                    var danduParse = fileStore.dandu && fileStore.dandu.parseResult;
                    items.push({label: "报名点数", value: summary.bmdCount});
                    items.push({label: "报考类别数", value: summary.categoryCount});
                    if (danduParse && danduParse.presence.auditStatus) {
                        items.push({label: "审核通过人数", value: summary.auditPassed});
                    }
                    if (summary.duplicateIds !== null && summary.duplicateIds !== undefined) {
                        items.push({label: "重复身份证数", value: summary.duplicateIds});
                    }
                } else {
                    items.push({label: "审核通过人数", value: summary.auditPassed});
                    items.push({label: "重复身份证数", value: summary.duplicateIds});
                }

                return items;
            },

            /* ---------------- 映射配置 ---------------- */
            loadMappings: function () {
                function read(key) {
                    try {
                        var raw = window.localStorage.getItem(key);
                        if (!raw) return [];
                        var parsed = JSON.parse(raw);
                        if (!Array.isArray(parsed)) return [];
                        return parsed.map(function (row) {
                            return {code: String(row.code || ""), name: String(row.name || "")};
                        });
                    } catch (err) {
                        return [];
                    }
                }

                this.bmdRows = read(STORAGE_KEYS.bmd);
                this.countyRows = read(STORAGE_KEYS.county);
            },

            persistMappings: function () {
                var self = this;
                window.clearTimeout(saveTimer);
                saveTimer = window.setTimeout(function () {
                    try {
                        window.localStorage.setItem(STORAGE_KEYS.bmd, JSON.stringify(self.bmdRows));
                        window.localStorage.setItem(STORAGE_KEYS.county, JSON.stringify(self.countyRows));
                    } catch (err) {
                        // localStorage 不可用时静默忽略，不影响统计
                    }
                }, 300);
            },

            addMappingRow: function (kind) {
                var rows = kind === "bmd" ? this.bmdRows : this.countyRows;
                rows.push({code: "", name: ""});
            },

            removeMappingRow: function (kind, index) {
                var rows = kind === "bmd" ? this.bmdRows : this.countyRows;
                rows.splice(index, 1);
            },

            isDupCode: function (kind, code) {
                var list = kind === "bmd" ? this.duplicateBmdCodes : this.duplicateCountyCodes;
                return list.indexOf(String(code || "").trim()) !== -1;
            },

            importBatch: function (kind) {
                var text = String(this.batchInput[kind] || "").trim();
                if (!text) {
                    this.showToast("请先在文本框中粘贴内容", "warning");
                    return;
                }

                var parsed = [];
                var self = this;

                try {
                    if (text.charAt(0) === "{" || text.charAt(0) === "[") {
                        var json = JSON.parse(text);
                        if (Array.isArray(json)) {
                            parsed = json.map(function (item) {
                                return {code: String(item.code !== undefined ? item.code : (item["代码"] || "")), name: String(item.name !== undefined ? item.name : (item["名称"] || ""))};
                            });
                        } else {
                            parsed = Object.keys(json).map(function (code) {
                                return {code: code, name: String(json[code])};
                            });
                        }
                    } else {
                        text.split(/\r?\n/).forEach(function (line) {
                            if (!line.trim()) return;
                            var cells = line.split("\t");
                            if (cells.length < 2) cells = line.split(",");
                            if (cells.length < 2) {
                                var matched = line.match(/^(\S+)\s+(.+)$/);
                                cells = matched ? [matched[1], matched[2]] : [line];
                            }
                            var code = String(cells[0] || "").trim();
                            var name = cells.slice(1).join(" ").trim();
                            if (code) parsed.push({code: code, name: name});
                        });
                    }
                } catch (err) {
                    this.showToast("批量导入内容无法解析，请检查是否为两列文本或合法 JSON", "error");
                    return;
                }

                if (!parsed.length) {
                    this.showToast("没有解析到有效数据", "warning");
                    return;
                }

                var rows = kind === "bmd" ? this.bmdRows : this.countyRows;
                var added = 0;
                var updated = 0;
                parsed.forEach(function (item) {
                    if (!item.code) return;
                    var existing = rows.filter(function (row) {
                        return String(row.code || "").trim() === item.code;
                    })[0];
                    if (existing) {
                        if (item.name) {
                            existing.name = item.name;
                            updated += 1;
                        }
                    } else {
                        rows.push({code: item.code, name: item.name});
                        added += 1;
                    }
                });

                this.batchInput[kind] = "";
                self.showToast("批量导入完成：新增 " + added + " 条，更新 " + updated + " 条", "success");
            },

            buildMapping: function () {
                var bmdMap = new Map();
                this.bmdRows.forEach(function (row) {
                    var code = DN.normalizeBmdCode(row.code).code;
                    var name = String(row.name || "").trim();
                    if (!code || !name) return;
                    if (!bmdMap.has(code)) bmdMap.set(code, name);
                });

                var countyMap = new Map();
                this.countyRows.forEach(function (row) {
                    var code = String(row.code || "").trim().replace(/[\s\u3000]/g, "");
                    if (/^\d$/.test(code)) code = "0" + code;
                    var name = String(row.name || "").trim();
                    if (!code || !name) return;
                    if (!countyMap.has(code)) countyMap.set(code, name);
                });

                return {bmdMap: bmdMap, countyMap: countyMap};
            },

            /* ---------------- 统计 ---------------- */
            generate: async function () {
                if (this.busy) return;

                var self = this;
                var selected = Object.keys(this.selection).filter(function (key) {
                    return self.selection[key];
                });

                if (!selected.length) {
                    this.showToast("请先勾选至少一种统计类型", "warning");
                    return;
                }

                var missing = selected.filter(function (key) {
                    return !self.isTypeAvailable(key);
                });
                if (missing.length) {
                    this.showToast("所选统计类型缺少文件，请先上传对应文件", "error");
                    return;
                }

                this.busy = true;
                this.rawResults = [];
                this.statDiagnostics = [];

                await new Promise(function (resolve) {
                    window.setTimeout(resolve, 20);
                });

                try {
                    var mapping = this.buildMapping();
                    var tables = [];
                    var diags = [];

                    if (this.selection.gaokao) {
                        var gaokaoData = Object.assign({}, fileStore.gaokao.parseResult, {fileName: this.files.gaokao.fileName});
                        var paymentData = fileStore.payment && fileStore.payment.parseResult
                            ? Object.assign({}, fileStore.payment.parseResult, {fileName: this.files.payment.fileName})
                            : null;
                        var gaokaoOut = window.GaokaoStats.build({
                            gaokao: gaokaoData,
                            payment: paymentData,
                            mapping: mapping
                        });
                        tables = tables.concat(gaokaoOut.tables);
                        diags = diags.concat(gaokaoOut.diagnostics);
                    }

                    if (this.selection.payment) {
                        var paymentOnly = Object.assign({}, fileStore.payment.parseResult, {fileName: this.files.payment.fileName});
                        var paymentOut = window.PaymentStats.build({payment: paymentOnly, mapping: mapping});
                        tables = tables.concat(paymentOut.tables);
                        diags = diags.concat(paymentOut.diagnostics);
                    }

                    if (this.selection.january) {
                        var januaryData = Object.assign({}, fileStore.january.parseResult, {fileName: this.files.january.fileName});
                        var januaryOut = window.JanuaryStats.build({january: januaryData, mapping: mapping});
                        tables = tables.concat(januaryOut.tables);
                        diags = diags.concat(januaryOut.diagnostics);
                    }

                    if (this.selection.dandu) {
                        var danduData = Object.assign({}, fileStore.dandu.parseResult, {fileName: this.files.dandu.fileName});
                        var danduOut = window.DanduStats.build({dandu: danduData, mapping: mapping});
                        tables = tables.concat(danduOut.tables);
                        diags = diags.concat(danduOut.diagnostics);
                    }

                    var generatedAt = new Date().toLocaleString("zh-CN", {hour12: false});
                    tables.forEach(function (table) {
                        table.generatedAt = generatedAt;
                        table.statType = String(table.key).split("-")[0];
                        // 结果表生成后不再变化，避免大量行数据被深度代理
                        if (Vue.markRaw) Vue.markRaw(table);
                    });

                    this.rawResults = tables;
                    this.statDiagnostics = diags;
                    this.generatedAt = generatedAt;
                    this.stale = false;
                    this.statusText = "已生成 " + tables.length + " 张统计表";
                } catch (err) {
                    if (window.console && window.console.error) window.console.error(err);
                    this.showToast("统计失败：" + (err && err.message ? err.message : "未知错误"), "error");
                } finally {
                    this.busy = false;
                }
            },

            clearResults: function () {
                this.rawResults = [];
                this.statDiagnostics = [];
                this.generatedAt = "";
                this.stale = false;
                this.statusText = "已清空统计结果（文件保留）";
            },

            resetAll: function () {
                if (this.busy) return;
                if (!window.confirm("重置全部：将清空已上传文件、统计结果和当前页面数据。是否继续？")) return;
                var clearMappings = window.confirm("是否同时清空报名点映射与县区映射？\n点击“确定”清空映射表；点击“取消”保留映射表。");

                var self = this;
                Object.keys(this.files).forEach(function (kind) {
                    fileStore[kind] = null;
                    self.files[kind] = createFileState();
                    self.selection[kind] = false;
                });
                this.rawResults = [];
                this.statDiagnostics = [];
                this.generatedAt = "";
                this.stale = false;
                this.batchInput = {bmd: "", county: ""};
                this.statusText = "已重置";

                if (clearMappings) {
                    this.bmdRows = [];
                    this.countyRows = [];
                    this.mappingOpen = {bmd: true, county: true};
                    try {
                        window.localStorage.removeItem(STORAGE_KEYS.bmd);
                        window.localStorage.removeItem(STORAGE_KEYS.county);
                    } catch (err) {
                        // 忽略
                    }
                    this.showToast("已清空文件、统计结果与映射表", "success");
                } else {
                    this.showToast("已清空文件与统计结果，映射表已保留", "success");
                }
            }
        }
    });

    app.mount("#app");
})();
