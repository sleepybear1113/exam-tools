/**
 * 字段标准化与原始行转换模块。
 * 统计层只依赖本模块输出的统一数据对象，不直接依赖 Excel 表头。
 */
(function (global) {
    "use strict";

    var TRIM_RE = /^[\s\u3000\uFEFF]+|[\s\u3000\uFEFF]+$/g;
    var ALL_SPACE_RE = /[\s\u3000\uFEFF]+/g;

    /* ------------------------------------------------------------------ */
    /* 常量口径                                                            */
    /* ------------------------------------------------------------------ */

    var ARTS_CATEGORIES = [
        {code: "30", label: "书法类"},
        {code: "31", label: "美术与设计类"},
        {code: "32", label: "音乐类"},
        {code: "33", label: "舞蹈类"},
        {code: "34", label: "表（导）演类"},
        {code: "36", label: "播音与主持类"},
        {code: "39", label: "戏曲类省际联考"},
        {code: "41", label: "体育类"}
    ];

    var ARTS_CODE_SET = {};
    ARTS_CATEGORIES.forEach(function (item) {
        ARTS_CODE_SET[item.code] = true;
    });

    var PAYMENT_CATEGORIES = [
        {code: "pt", label: "普通类"},
        {code: "cm", label: "传媒"},
        {code: "ms", label: "美术类"},
        {code: "sf", label: "书法类"},
        {code: "szby", label: "时装表演类"},
        {code: "ty", label: "体育类"},
        {code: "wd", label: "舞蹈类"},
        {code: "yy", label: "音乐类"}
    ];

    var PAYMENT_CATEGORY_SET = {};
    PAYMENT_CATEGORIES.forEach(function (item) {
        PAYMENT_CATEGORY_SET[item.code] = true;
    });

    // 术科代码 → 缴费项目类别；39（戏曲类省际联考）不需要缴费项目
    var ARTS_TO_PAYMENT = {
        "30": "sf",
        "31": "ms",
        "32": "yy",
        "33": "wd",
        "34": "szby",
        "36": "cm",
        "41": "ty"
    };

    var JANUARY_SUBJECTS = [
        {key: "sz", field: "KM_SZ", label: "政治"},
        {key: "ls", field: "KM_LS", label: "历史"},
        {key: "dl", field: "KM_DL", label: "地理"},
        {key: "wl", field: "KM_WL", label: "物理"},
        {key: "hx", field: "KM_HX", label: "化学"},
        {key: "sw", field: "KM_SW", label: "生物"},
        {key: "js", field: "KM_JS", label: "技术"}
    ];

    var LANGUAGES = [
        {code: "1", label: "英语"},
        {code: "2", label: "俄语"},
        {code: "3", label: "日语"},
        {code: "4", label: "德语"},
        {code: "5", label: "法语"},
        {code: "6", label: "西班牙语"}
    ];

    // 单独考试：报考类别（BKKL_）代码含义
    var DANDU_CATEGORIES = [
        {code: "01", label: "机械类"},
        {code: "02", label: "计算机类"},
        {code: "03", label: "文秘类"},
        {code: "04", label: "化工(环保)类"},
        {code: "05", label: "药学类"},
        {code: "06", label: "建筑类"},
        {code: "07", label: "烹饪类"},
        {code: "08", label: "旅游服务类"},
        {code: "09", label: "服装类"},
        {code: "10", label: "财会类"},
        {code: "11", label: "电子与电工类"},
        {code: "12", label: "商业类"},
        {code: "13", label: "外贸类"},
        {code: "14", label: "医学护理类"},
        {code: "15", label: "农艺类"},
        {code: "16", label: "艺术类"},
        {code: "17", label: "其他类"}
    ];
    var DANDU_CATEGORY_LABELS = {};
    DANDU_CATEGORIES.forEach(function (item) {
        DANDU_CATEGORY_LABELS[item.code] = item.label;
    });

    // 单独考试：报考类别（BKKL_）为 17 时，QTLZY_ 的取值含义
    var OTHER_MAJOR_LABELS = {
        "1": "安全防范",
        "2": "体育",
        "4": "学前教育",
        "5": "退役士兵高职招生",
        "6": "汽车专业",
        "7": "内地西藏新疆中职班"
    };

    // 单独考试：其他类专业（QTLZY_）仅在 BKKL_ = 17 时有意义
    var OTHER_MAJOR_PARENT = "17";

    /* ------------------------------------------------------------------ */
    /* 基础工具                                                            */
    /* ------------------------------------------------------------------ */

    function stripText(value) {
        if (value === null || value === undefined) return "";
        if (Object.prototype.toString.call(value) === "[object Date]") return "";
        return String(value).replace(TRIM_RE, "");
    }

    function isBlankRow(row) {
        if (!row || !row.length) return true;
        for (var i = 0; i < row.length; i++) {
            if (stripText(row[i]) !== "") return false;
        }
        return true;
    }

    function cellValue(row, map, key) {
        var index = map[key];
        if (index === undefined || index === null) return undefined;
        return row[index];
    }

    function bump(bucket, key) {
        var name = key === "" || key === null || key === undefined ? "(空白)" : String(key);
        bucket[name] = (bucket[name] || 0) + 1;
    }

    /**
     * 身份证号码标准化：不做数值转换，仅去首尾空格并把末位 x 转大写。
     */
    function normalizeIdCard(value) {
        if (value === null || value === undefined) return "";
        if (typeof value === "number") {
            if (!Number.isSafeInteger(value)) return "";
            return String(value).replace(TRIM_RE, "").toUpperCase();
        }
        return String(value).replace(TRIM_RE, "").replace(/x$/, "X");
    }

    /**
     * 数值单元格或已变成科学计数法的身份证，视为精度可能丢失。
     */
    function isCorruptedId(value) {
        if (typeof value === "number") return !Number.isSafeInteger(value);
        if (typeof value === "string") return /^\d(\.\d+)?e\+?\d+$/i.test(stripText(value));
        return false;
    }

    /**
     * 报名点代码标准化：纯数字不足 5 位左补零，超长不截断但标记异常。
     */
    function normalizeBmdCode(value) {
        var text = stripText(value).replace(ALL_SPACE_RE, "");
        if (!text) return {code: "", issue: "empty"};
        if (/^\d+$/.test(text)) {
            if (text.length < 5) return {code: text.padStart(5, "0"), issue: null};
            if (text.length > 5) return {code: text, issue: "too-long"};
            return {code: text, issue: null};
        }
        return {code: text, issue: text.length > 5 ? "too-long" : null};
    }

    /**
     * 报考类别代码标准化：纯数字统一补齐为两位（1 → 01），保证 01 与 1 不重复计数。
     */
    function normalizeDanduCategory(value) {
        var text = stripText(value).replace(ALL_SPACE_RE, "");
        if (/^\d$/.test(text)) return "0" + text;
        return text;
    }

    function countyCodeOf(bmdCode) {
        if (!bmdCode || bmdCode.length < 3) return "";
        return bmdCode.slice(1, 3);
    }

    function normalizeAuditStatus(value) {
        return stripText(value).replace(ALL_SPACE_RE, "");
    }

    function isAuditPassed(statusText) {
        return statusText === "3" || statusText === "审核通过";
    }

    function parseArtsCodes(value) {
        var text = stripText(value);
        if (!text) return [];
        var result = [];
        text.split(/[|｜]/).forEach(function (item) {
            var code = item.trim();
            if (code && result.indexOf(code) === -1) result.push(code);
        });
        return result;
    }

    /**
     * CJTZSCS_ 固定口径：去首尾空格后严格等于 "1"。
     */
    function isSpecialTestHit(value) {
        return stripText(value) === "1";
    }

    /**
     * 从缴款单号提取报名点代码（位置按 1 起始计数）。
     */
    function extractBmdFromOrderNo(value) {
        if (value === null || value === undefined) return {code: "", issue: "empty"};
        if (typeof value === "number") return {code: "", issue: "numeric"};
        var text = String(value).replace(TRIM_RE, "");
        if (!text) return {code: "", issue: "empty"};
        if (text.length < 23) return {code: "", issue: "too-short"};

        var suffix = text.slice(17, 21);
        var indicatorText = text.slice(21, 23);
        if (!/^\d{4}$/.test(suffix)) return {code: "", issue: "suffix"};
        if (!/^\d{2}$/.test(indicatorText)) return {code: "", issue: "indicator"};

        var indicator = Number(indicatorText);
        var first;
        if (indicator <= 12) first = "8";
        else if (indicator <= 50) first = "9";
        else first = "7";

        return {code: first + suffix, issue: null};
    }

    function parseProjectCategory(projectNo) {
        var text = stripText(projectNo);
        if (!text) return {code: "", raw: text, issue: "empty"};
        var index = text.lastIndexOf("_");
        if (index < 0) return {code: "", raw: text, issue: "format"};
        var code = text.slice(index + 1).trim().toLowerCase();
        if (!code) return {code: "", raw: text, issue: "format"};
        return {code: code, raw: text, issue: null};
    }

    function parsePayStatus(value) {
        var text = stripText(value).replace(ALL_SPACE_RE, "");
        if (text === "已缴款" || text === "已缴费") return "paid";
        if (text === "待缴款" || text === "未缴款" || text === "待缴费" || text === "未缴费") return "unpaid";
        return "unknown";
    }

    function maskIdCard(value) {
        var text = String(value || "");
        if (text.length <= 8) return text ? text.slice(0, 2) + "****" : "";
        return text.slice(0, 4) + new Array(text.length - 7).join("*") + text.slice(-4);
    }

    /* ------------------------------------------------------------------ */
    /* 异常收集器                                                          */
    /* ------------------------------------------------------------------ */

    function createCollector() {
        var entries = new Map();

        function ensure(level, code, title) {
            var entry = entries.get(code);
            if (!entry) {
                entry = {level: level, code: code, title: title, count: 0, samples: []};
                entries.set(code, entry);
            }
            return entry;
        }

        function pushSamples(entry, samples) {
            (samples || []).forEach(function (sample) {
                if (entry.samples.length >= 5) return;
                var text = String(sample);
                if (entry.samples.indexOf(text) === -1) entry.samples.push(text);
            });
        }

        return {
            add: function (level, code, title, sample) {
                var entry = ensure(level, code, title);
                entry.count += 1;
                if (sample !== undefined && sample !== null && sample !== "") pushSamples(entry, [sample]);
            },
            addMany: function (level, code, title, count, samples) {
                var entry = ensure(level, code, title);
                entry.count += count;
                pushSamples(entry, samples);
            },
            merge: function (list) {
                (list || []).forEach(function (item) {
                    var entry = ensure(item.level, item.code, item.title);
                    entry.count += item.count || 0;
                    pushSamples(entry, item.samples);
                });
            },
            toArray: function () {
                return Array.from(entries.values());
            }
        };
    }

    /**
     * 名称解析：未配置时返回“未配置名称”并记录警告。
     */
    function createNameResolver(mapping, diag) {
        var bmdMap = (mapping && mapping.bmdMap) || new Map();
        var countyMap = (mapping && mapping.countyMap) || new Map();

        return {
            bmd: function (code) {
                if (!code) {
                    diag.add("warning", "MAP_BMD_EMPTY", "报名点代码为空，无法匹配报名点名称", "");
                    return "未配置名称";
                }
                var name = bmdMap.get(code);
                if (!name) diag.add("warning", "MAP_BMD_MISSING", "报名点名称未配置，结果中显示为“未配置名称”", code);
                return name || "未配置名称";
            },
            county: function (code) {
                if (!code) {
                    diag.add("warning", "MAP_COUNTY_EMPTY", "报名点代码不足 3 位，无法提取县区代码", "");
                    return "未配置名称";
                }
                var name = countyMap.get(code);
                if (!name) diag.add("warning", "MAP_COUNTY_MISSING", "县区名称未配置，结果中显示为“未配置名称”", code);
                return name || "未配置名称";
            }
        };
    }

    /* ------------------------------------------------------------------ */
    /* 高考报名信息解析                                                    */
    /* ------------------------------------------------------------------ */

    function parseGaokao(ctx) {
        var map = (ctx.columns && ctx.columns.map) || {};
        var dataRows = ctx.dataRows || [];
        var firstRowNumber = ctx.firstDataRowNumber || 1;
        var diag = createCollector();

        var presence = {
            auditStatus: map.auditStatus !== undefined,
            bmdCode: map.bmdCode !== undefined,
            idCard: map.idCard !== undefined,
            name: map.name !== undefined,
            arts: map.arts !== undefined,
            specialTest: map.specialTest !== undefined
        };

        var distributions = {audit: {}, special: {}, artsUnknown: {}};
        var records = [];

        for (var i = 0; i < dataRows.length; i++) {
            var row = dataRows[i] || [];
            if (isBlankRow(row)) continue;
            var rowNumber = firstRowNumber + i;

            var rawId = cellValue(row, map, "idCard");
            var corrupted = isCorruptedId(rawId);
            var idCard = corrupted ? "" : normalizeIdCard(rawId);
            if (corrupted) {
                diag.add("warning", "GK_ID_CORRUPTED", "身份证号码为数值单元格且可能已丢失精度，该行不参与统计与跨文件匹配", "第 " + rowNumber + " 行");
            } else if (!idCard) {
                diag.add("warning", "GK_ID_EMPTY", "身份证号码为空，该行未纳入统计", "第 " + rowNumber + " 行");
            } else if (typeof rawId === "number") {
                diag.add("warning", "GK_ID_NUMERIC", "身份证号码来自数值单元格，可能存在前导零丢失", maskIdCard(idCard));
            }

            var auditStatus = presence.auditStatus ? normalizeAuditStatus(cellValue(row, map, "auditStatus")) : "";
            if (presence.auditStatus) bump(distributions.audit, auditStatus);

            var bmd = presence.bmdCode ? normalizeBmdCode(cellValue(row, map, "bmdCode")) : {code: "", issue: "empty"};
            if (bmd.issue === "too-long") {
                diag.add("warning", "GK_BMD_LONG", "报名点代码超过 5 位，已按原值保留", bmd.code);
            } else if (bmd.issue === "empty") {
                diag.add("warning", "GK_BMD_EMPTY", "报名点代码为空，该行未纳入分组统计", "第 " + rowNumber + " 行");
            }

            var artsCodes = presence.arts ? parseArtsCodes(cellValue(row, map, "arts")) : [];
            artsCodes.forEach(function (code) {
                if (!ARTS_CODE_SET[code]) distributions.artsUnknown[code] = (distributions.artsUnknown[code] || 0) + 1;
            });

            var specialRaw = presence.specialTest ? stripText(cellValue(row, map, "specialTest")) : "";
            if (presence.specialTest && specialRaw !== "") bump(distributions.special, specialRaw);

            records.push({
                rowNumber: rowNumber,
                idCard: idCard,
                idCorrupted: corrupted,
                idNumeric: typeof rawId === "number",
                name: presence.name ? stripText(cellValue(row, map, "name")) : "",
                bmdCode: bmd.code,
                bmdIssue: bmd.issue,
                auditStatus: auditStatus,
                artsCodes: artsCodes,
                specialHit: presence.specialTest && isSpecialTestHit(cellValue(row, map, "specialTest")),
                specialRaw: specialRaw,
                className: stripText(cellValue(row, map, "className")),
                signupNo: stripText(cellValue(row, map, "signupNo"))
            });
        }

        Object.keys(distributions.artsUnknown).forEach(function (code) {
            diag.addMany("warning", "GK_ARTS_UNKNOWN", "术科代码 " + code + " 不在本期统计范围内，未生成对应统计列", distributions.artsUnknown[code], [code]);
        });

        if (!presence.arts) {
            diag.add("warning", "GK_ARTS_FIELD_MISSING", "缺少艺术体育术科字段（YSTYSK_），术科各列显示为“字段缺失”");
        }
        if (!presence.specialTest) {
            diag.add("warning", "GK_SPECIAL_FIELD_MISSING", "缺少特招生测试字段（CJTZSCS_），该指标显示为“字段缺失”");
        }

        var passed = records.filter(function (record) {
            return isAuditPassed(record.auditStatus);
        });
        var passedIds = passed.filter(function (record) {
            return !!record.idCard;
        }).map(function (record) {
            return record.idCard;
        });
        var uniqueIds = new Set(passedIds);

        return {
            records: records,
            presence: presence,
            distributions: distributions,
            diagnostics: diag.toArray(),
            summary: {
                totalRows: records.length,
                auditPassed: passed.length,
                uniqueIds: uniqueIds.size,
                duplicateIds: passedIds.length - uniqueIds.size,
                filteredOut: records.length - passed.length
            }
        };
    }

    /* ------------------------------------------------------------------ */
    /* 高考报名缴费信息解析                                                */
    /* ------------------------------------------------------------------ */

    function parsePayment(ctx) {
        var map = (ctx.columns && ctx.columns.map) || {};
        var dataRows = ctx.dataRows || [];
        var firstRowNumber = ctx.firstDataRowNumber || 1;
        var diag = createCollector();

        var presence = {
            orderNo: map.orderNo !== undefined,
            idCard: map.idCard !== undefined,
            name: map.name !== undefined,
            payStatus: map.payStatus !== undefined,
            projectNo: map.projectNo !== undefined,
            auditStatus: map.auditStatus !== undefined
        };

        var distributions = {audit: {}, payStatus: {}, projectUnknown: {}};
        var records = [];
        var bmdSet = new Set();
        var orderNoFail = 0;

        for (var i = 0; i < dataRows.length; i++) {
            var row = dataRows[i] || [];
            if (isBlankRow(row)) continue;
            var rowNumber = firstRowNumber + i;

            var rawId = cellValue(row, map, "idCard");
            var corrupted = isCorruptedId(rawId);
            var idCard = corrupted ? "" : normalizeIdCard(rawId);
            if (corrupted) {
                diag.add("warning", "PAY_ID_CORRUPTED", "缴费记录身份证号码精度可能已丢失，该行不参与统计", "第 " + rowNumber + " 行");
            } else if (!idCard) {
                diag.add("warning", "PAY_ID_EMPTY", "缴费记录身份证号码为空，该行未纳入统计", "第 " + rowNumber + " 行");
            }

            var orderResult = presence.orderNo ? extractBmdFromOrderNo(cellValue(row, map, "orderNo")) : {code: "", issue: "empty"};
            if (orderResult.issue) {
                orderNoFail += 1;
                var reasonMap = {
                    "empty": "缴款单号为空",
                    "numeric": "缴款单号为数值单元格，精度可能丢失",
                    "too-short": "缴款单号长度不足 23 位",
                    "suffix": "缴款单号第 18–21 位不是 4 位有效字符",
                    "indicator": "缴款单号第 22–23 位不是数字"
                };
                diag.add("warning", "PAY_ORDER_" + orderResult.issue.toUpperCase(), "无法从缴款单号提取报名点代码：" + (reasonMap[orderResult.issue] || orderResult.issue), "第 " + rowNumber + " 行");
            } else {
                bmdSet.add(orderResult.code);
            }

            var project = presence.projectNo ? parseProjectCategory(cellValue(row, map, "projectNo")) : {code: "", raw: "", issue: "empty"};
            var categoryKnown = false;
            if (project.issue) {
                var projReason = project.issue === "empty" ? "项目编号为空" : "项目编号缺少下划线或后缀为空";
                diag.add("warning", "PAY_PROJECT_FORMAT", "无法解析缴费项目类别：" + projReason, project.raw || "第 " + rowNumber + " 行");
            } else if (!PAYMENT_CATEGORY_SET[project.code]) {
                if (!distributions.projectUnknown[project.code]) distributions.projectUnknown[project.code] = {count: 0, samples: []};
                distributions.projectUnknown[project.code].count += 1;
                if (distributions.projectUnknown[project.code].samples.length < 5) {
                    distributions.projectUnknown[project.code].samples.push(project.raw);
                }
            } else {
                categoryKnown = true;
            }

            var payState = presence.payStatus ? parsePayStatus(cellValue(row, map, "payStatus")) : "unknown";
            if (presence.payStatus) bump(distributions.payStatus, stripText(cellValue(row, map, "payStatus")));
            if (payState === "unknown") {
                diag.add("warning", "PAY_STATUS_UNKNOWN", "缴款状态无法识别，未归入已缴费或未缴费", "第 " + rowNumber + " 行：" + (stripText(cellValue(row, map, "payStatus")) || "(空白)"));
            }

            var auditStatus = presence.auditStatus ? normalizeAuditStatus(cellValue(row, map, "auditStatus")) : "";
            if (presence.auditStatus) bump(distributions.audit, auditStatus);

            records.push({
                rowNumber: rowNumber,
                idCard: idCard,
                idCorrupted: corrupted,
                name: presence.name ? stripText(cellValue(row, map, "name")) : "",
                orderNo: stripText(cellValue(row, map, "orderNo")),
                bmdCode: orderResult.code,
                bmdIssue: orderResult.issue,
                projectNo: project.raw,
                categoryCode: project.code,
                categoryKnown: categoryKnown,
                payState: payState,
                auditStatus: auditStatus
            });
        }

        Object.keys(distributions.projectUnknown).forEach(function (code) {
            var info = distributions.projectUnknown[code];
            diag.addMany("warning", "PAY_PROJECT_UNKNOWN", "缴费项目代码 " + code + " 未识别，未参与已知类别统计", info.count, info.samples);
        });

        return {
            records: records,
            presence: presence,
            distributions: distributions,
            diagnostics: diag.toArray(),
            summary: {
                totalRows: records.length,
                bmdCount: bmdSet.size,
                orderNoFail: orderNoFail
            }
        };
    }

    /* ------------------------------------------------------------------ */
    /* 1月学选考报名信息解析                                               */
    /* ------------------------------------------------------------------ */

    function parseJanuary(ctx) {
        var map = (ctx.columns && ctx.columns.map) || {};
        var dataRows = ctx.dataRows || [];
        var firstRowNumber = ctx.firstDataRowNumber || 1;
        var diag = createCollector();

        function has(key) {
            return map[key] !== undefined;
        }

        var presence = {
            idCard: has("idCard"),
            bmdCode: has("bmdCode"),
            name: has("name"),
            auditStatus: has("auditStatus"),
            languageCode: has("languageCode"),
            yw: has("yw"),
            sx: has("sx"),
            yy: has("yy"),
            subjects: {}
        };
        JANUARY_SUBJECTS.forEach(function (item) {
            presence.subjects[item.key] = has(item.key);
        });

        var distributions = {audit: {}, languages: {}, subjectUnknown: {}};
        var records = [];

        for (var i = 0; i < dataRows.length; i++) {
            var row = dataRows[i] || [];
            if (isBlankRow(row)) continue;
            var rowNumber = firstRowNumber + i;

            var rawId = cellValue(row, map, "idCard");
            var corrupted = isCorruptedId(rawId);
            var idCard = corrupted ? "" : normalizeIdCard(rawId);
            if (corrupted) {
                diag.add("warning", "JAN_ID_CORRUPTED", "身份证号码为数值单元格且可能已丢失精度，该行不参与统计", "第 " + rowNumber + " 行");
            } else if (!idCard) {
                diag.add("warning", "JAN_ID_EMPTY", "身份证号码为空，该行未纳入统计", "第 " + rowNumber + " 行");
            } else if (typeof rawId === "number") {
                diag.add("warning", "JAN_ID_NUMERIC", "身份证号码来自数值单元格，可能存在前导零丢失", maskIdCard(idCard));
            }

            var auditStatus = presence.auditStatus ? normalizeAuditStatus(cellValue(row, map, "auditStatus")) : "";
            if (presence.auditStatus) bump(distributions.audit, auditStatus);

            var bmd = presence.bmdCode ? normalizeBmdCode(cellValue(row, map, "bmdCode")) : {code: "", issue: "empty"};
            if (bmd.issue === "too-long") {
                diag.add("warning", "JAN_BMD_LONG", "报名点代码超过 5 位，已按原值保留", bmd.code);
            } else if (bmd.issue === "empty") {
                diag.add("warning", "JAN_BMD_EMPTY", "报名点代码为空，该行未纳入分组统计", "第 " + rowNumber + " 行");
            }

            var languageCode = presence.languageCode ? stripText(cellValue(row, map, "languageCode")).replace(ALL_SPACE_RE, "") : "";
            if (languageCode) bump(distributions.languages, languageCode);

            var subjects = {
                yw: presence.yw ? stripText(cellValue(row, map, "yw")) : "",
                sx: presence.sx ? stripText(cellValue(row, map, "sx")) : "",
                yy: presence.yy ? stripText(cellValue(row, map, "yy")) : ""
            };
            JANUARY_SUBJECTS.forEach(function (item) {
                var value = presence.subjects[item.key] ? stripText(cellValue(row, map, item.key)) : "";
                subjects[item.key] = value;
                if (value && value !== "学考" && value !== "选考") {
                    bump(distributions.subjectUnknown, item.label + "=" + value);
                }
            });

            records.push({
                rowNumber: rowNumber,
                idCard: idCard,
                idCorrupted: corrupted,
                name: presence.name ? stripText(cellValue(row, map, "name")) : "",
                bmdCode: bmd.code,
                bmdIssue: bmd.issue,
                auditStatus: auditStatus,
                languageCode: languageCode,
                subjects: subjects
            });
        }

        Object.keys(distributions.subjectUnknown).forEach(function (key) {
            diag.add("warning", "JAN_SUBJECT_UNKNOWN", "科目取值不在“学考/选考/空白”范围内，未计入对应列", key);
        });

        if (!presence.yy) {
            diag.add("warning", "JAN_FOREIGN_FIELD_MISSING", "缺少 KM_YY 字段，外语相关列显示为“字段缺失”");
        }

        var passed = records.filter(function (record) {
            return isAuditPassed(record.auditStatus);
        });
        var passedIds = passed.filter(function (record) {
            return !!record.idCard;
        }).map(function (record) {
            return record.idCard;
        });
        var uniqueIds = new Set(passedIds);

        return {
            records: records,
            presence: presence,
            distributions: distributions,
            diagnostics: diag.toArray(),
            summary: {
                totalRows: records.length,
                auditPassed: passed.length,
                uniqueIds: uniqueIds.size,
                duplicateIds: passedIds.length - uniqueIds.size,
                filteredOut: records.length - passed.length
            }
        };
    }

    /* ------------------------------------------------------------------ */
    /* 单独考试报名信息解析                                                */
    /* ------------------------------------------------------------------ */

    function parseDandu(ctx) {
        var map = (ctx.columns && ctx.columns.map) || {};
        var dataRows = ctx.dataRows || [];
        var firstRowNumber = ctx.firstDataRowNumber || 1;
        var diag = createCollector();

        var presence = {
            bmdCode: map.bmdCode !== undefined,
            category: map.category !== undefined,
            otherMajor: map.otherMajor !== undefined,
            idCard: map.idCard !== undefined,
            auditStatus: map.auditStatus !== undefined,
            name: map.name !== undefined
        };

        var distributions = {audit: {}, categories: {}, otherMajors: {}, otherUnknown: {}};
        var records = [];
        var bmdSet = new Set();
        var categorySet = new Set();

        for (var i = 0; i < dataRows.length; i++) {
            var row = dataRows[i] || [];
            if (isBlankRow(row)) continue;
            var rowNumber = firstRowNumber + i;

            var bmd = presence.bmdCode ? normalizeBmdCode(cellValue(row, map, "bmdCode")) : {code: "", issue: "empty"};
            if (bmd.issue === "too-long") {
                diag.add("warning", "DANDU_BMD_LONG", "报名点代码超过 5 位，已按原值保留", bmd.code);
            } else if (bmd.issue === "empty") {
                diag.add("warning", "DANDU_BMD_EMPTY", "报名点代码为空，该行未纳入分组统计", "第 " + rowNumber + " 行");
            }
            if (bmd.code) bmdSet.add(bmd.code);

            var categoryCode = presence.category
                ? normalizeDanduCategory(cellValue(row, map, "category"))
                : "";
            if (!categoryCode) {
                diag.add("warning", "DANDU_CATEGORY_EMPTY", "报考类别（BKKL_）为空，该行计入“(空白)”类别", "第 " + rowNumber + " 行");
            } else {
                categorySet.add(categoryCode);
                bump(distributions.categories, categoryCode);
            }

            var otherCode = presence.otherMajor
                ? stripText(cellValue(row, map, "otherMajor")).replace(ALL_SPACE_RE, "")
                : "";
            if (otherCode) {
                if (categoryCode !== OTHER_MAJOR_PARENT) {
                    diag.add(
                        "warning",
                        "DANDU_OTHER_IGNORED",
                        "QTLZY_（其他类专业）仅在报考类别为 17 时有值，该行其他取值未参与分类",
                        "第 " + rowNumber + " 行：" + categoryCode + "/" + otherCode
                    );
                } else if (!OTHER_MAJOR_LABELS[otherCode]) {
                    distributions.otherUnknown[otherCode] = (distributions.otherUnknown[otherCode] || 0) + 1;
                } else {
                    bump(distributions.otherMajors, otherCode + " " + OTHER_MAJOR_LABELS[otherCode]);
                }
            }

            var auditStatus = presence.auditStatus ? normalizeAuditStatus(cellValue(row, map, "auditStatus")) : "";
            if (presence.auditStatus) bump(distributions.audit, auditStatus);

            var rawId = presence.idCard ? cellValue(row, map, "idCard") : undefined;
            var corrupted = presence.idCard ? isCorruptedId(rawId) : false;
            var idCard = presence.idCard ? (corrupted ? "" : normalizeIdCard(rawId)) : "";
            if (presence.idCard && corrupted) {
                diag.add("warning", "DANDU_ID_CORRUPTED", "身份证号码为数值单元格且可能已丢失精度，该行不参与去重", "第 " + rowNumber + " 行");
            } else if (presence.idCard && !idCard) {
                diag.add("warning", "DANDU_ID_EMPTY", "身份证号码为空，该行不纳入去重统计", "第 " + rowNumber + " 行");
            }

            records.push({
                rowNumber: rowNumber,
                idCard: idCard,
                idCorrupted: corrupted,
                name: presence.name ? stripText(cellValue(row, map, "name")) : "",
                bmdCode: bmd.code,
                bmdIssue: bmd.issue,
                categoryCode: categoryCode,
                otherCode: otherCode,
                categoryKey: resolveDanduCategoryKey(categoryCode, otherCode),
                auditStatus: auditStatus
            });
        }

        Object.keys(distributions.otherUnknown).forEach(function (code) {
            diag.addMany(
                "warning",
                "DANDU_OTHER_UNKNOWN",
                "QTLZY_ 代码 " + code + " 未在已知范围（1/2/4/5/6/7）内，已并入“17”类别统计",
                distributions.otherUnknown[code],
                [code]
            );
        });

        if (!presence.otherMajor) {
            diag.add("warning", "DANDU_OTHER_FIELD_MISSING", "缺少 QTLZY_（其他类专业）字段，报考类别 17 无法细分");
        }

        var passed = presence.auditStatus ? records.filter(function (record) {
            return isAuditPassed(record.auditStatus);
        }) : records.slice();
        var passedIds = presence.idCard ? passed.filter(function (record) {
            return !!record.idCard;
        }).map(function (record) {
            return record.idCard;
        }) : [];
        var uniqueIds = presence.idCard ? new Set(passedIds) : null;

        return {
            records: records,
            presence: presence,
            distributions: distributions,
            diagnostics: diag.toArray(),
            summary: {
                totalRows: records.length,
                bmdCount: bmdSet.size,
                categoryCount: categorySet.size,
                auditPassed: passed.length,
                uniqueIds: uniqueIds ? uniqueIds.size : null,
                duplicateIds: presence.idCard ? passedIds.length - uniqueIds.size : null,
                filteredOut: records.length - passed.length
            }
        };
    }

    /**
     * 类别键：BKKL_ 为 17 且 QTLZY_ 命中已知代码时拆成 17-x，否则使用 BKKL_ 原值。
     */
    function resolveDanduCategoryKey(categoryCode, otherCode) {
        if (!categoryCode) return "";
        if (categoryCode === OTHER_MAJOR_PARENT && otherCode && OTHER_MAJOR_LABELS[otherCode]) {
            return categoryCode + "-" + otherCode;
        }
        return categoryCode;
    }

    function danduCategoryLabel(categoryKey) {
        var key = String(categoryKey === null || categoryKey === undefined ? "" : categoryKey);
        if (!key) return "(空白)";
        var parts = key.split("-");
        if (parts.length === 2 && parts[0] === OTHER_MAJOR_PARENT && OTHER_MAJOR_LABELS[parts[1]]) {
            return key + " " + OTHER_MAJOR_LABELS[parts[1]];
        }
        if (DANDU_CATEGORY_LABELS[key]) return key + " " + DANDU_CATEGORY_LABELS[key];
        return key;
    }

    /* ------------------------------------------------------------------ */
    /* 分布信息格式化                                                      */
    /* ------------------------------------------------------------------ */

    function formatDistributions(kind, distributions) {
        var result = [];
        if (!distributions) return result;

        function push(title, bucket) {
            var entries = Object.keys(bucket || {})
                .filter(function (key) {
                    return bucket[key] > 0;
                })
                .sort(function (a, b) {
                    return bucket[b] - bucket[a];
                });
            if (!entries.length) return;
            result.push({
                title: title,
                items: entries.map(function (key) {
                    return {label: key, count: bucket[key]};
                })
            });
        }

        push("审核状态取值分布", distributions.audit);

        if (kind === "gaokao") {
            push("CJTZSCS_ 非空取值分布", distributions.special);
            push("未识别术科代码出现次数", distributions.artsUnknown);
        } else if (kind === "payment") {
            push("缴款状态取值分布", distributions.payStatus);
            push("未识别项目代码出现次数", distributions.projectUnknown && (function () {
                var bucket = {};
                Object.keys(distributions.projectUnknown).forEach(function (code) {
                    bucket[code] = distributions.projectUnknown[code].count;
                });
                return bucket;
            })());
        } else if (kind === "january") {
            push("外语语种代码分布", distributions.languages);
            push("无法识别的科目取值", distributions.subjectUnknown);
        } else if (kind === "dandu") {
            var categoryBucket = {};
            Object.keys(distributions.categories || {}).forEach(function (code) {
                categoryBucket[danduCategoryLabel(code)] = distributions.categories[code];
            });
            push("报考类别（BKKL_）取值分布", categoryBucket);
            push("其他类专业（QTLZY_）细分分布", distributions.otherMajors);
            push("未识别的其他类专业代码", distributions.otherUnknown);
        }

        return result;
    }

    global.DataNormalizer = {
        ARTS_CATEGORIES: ARTS_CATEGORIES,
        ARTS_CODE_SET: ARTS_CODE_SET,
        PAYMENT_CATEGORIES: PAYMENT_CATEGORIES,
        PAYMENT_CATEGORY_SET: PAYMENT_CATEGORY_SET,
        ARTS_TO_PAYMENT: ARTS_TO_PAYMENT,
        JANUARY_SUBJECTS: JANUARY_SUBJECTS,
        LANGUAGES: LANGUAGES,
        OTHER_MAJOR_LABELS: OTHER_MAJOR_LABELS,
        OTHER_MAJOR_PARENT: OTHER_MAJOR_PARENT,
        DANDU_CATEGORIES: DANDU_CATEGORIES,
        DANDU_CATEGORY_LABELS: DANDU_CATEGORY_LABELS,
        stripText: stripText,
        isBlankRow: isBlankRow,
        normalizeIdCard: normalizeIdCard,
        isCorruptedId: isCorruptedId,
        normalizeBmdCode: normalizeBmdCode,
        normalizeDanduCategory: normalizeDanduCategory,
        countyCodeOf: countyCodeOf,
        normalizeAuditStatus: normalizeAuditStatus,
        isAuditPassed: isAuditPassed,
        parseArtsCodes: parseArtsCodes,
        isSpecialTestHit: isSpecialTestHit,
        extractBmdFromOrderNo: extractBmdFromOrderNo,
        parseProjectCategory: parseProjectCategory,
        parsePayStatus: parsePayStatus,
        maskIdCard: maskIdCard,
        createCollector: createCollector,
        createNameResolver: createNameResolver,
        parseGaokao: parseGaokao,
        parsePayment: parsePayment,
        parseJanuary: parseJanuary,
        parseDandu: parseDandu,
        resolveDanduCategoryKey: resolveDanduCategoryKey,
        danduCategoryLabel: danduCategoryLabel,
        formatDistributions: formatDistributions
    };
})(window);
