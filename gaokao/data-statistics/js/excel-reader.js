/**
 * Excel 读取与表头识别模块。
 * 依赖全局 XLSX（SheetJS 完整版，支持 .xls 与 .xlsx）。
 * 所有解析均在浏览器本地完成，不发起任何网络请求。
 */
(function (global) {
    "use strict";

    var MAX_HEADER_SCAN = 20;

    /* ------------------------------------------------------------------ */
    /* 表头规范化                                                          */
    /* ------------------------------------------------------------------ */

    function stripText(value) {
        if (value === null || value === undefined) return "";
        return String(value).replace(/^[\s\u3000\uFEFF]+|[\s\u3000\uFEFF]+$/g, "");
    }

    function normalizeHeaderCell(value) {
        if (value === null || value === undefined) return "";
        return String(value).replace(/[\s\u3000\uFEFF]+/g, "").toUpperCase();
    }

    /* ------------------------------------------------------------------ */
    /* 文件类型 Schema                                                     */
    /* ------------------------------------------------------------------ */

    var gaokaoSchema = {
        kind: "gaokao",
        label: "高考报名信息",
        matchMode: "code",
        // 出现 1 月学选考的科目字段（KM_*）或单独考试字段时，说明文件放错区域，不得误解析
        denyCodes: ["KM", "BKKL", "QTLZY"],
        fields: [
            {key: "auditStatus", code: "SHZT", required: true, label: "审核状态(SHZT_)"},
            {key: "className", code: "BJ", label: "班级(BJ_)"},
            {key: "arts", code: "YSTYSK", label: "艺术体育术科(YSTYSK_)"},
            {key: "voice", code: "SYCFDM", label: "声乐唱法代码(SYCFDM_)"},
            {key: "danceType", code: "WDWZ", label: "舞蹈类舞种(WDWZ_)"},
            {key: "signupNo", code: "BMXH", label: "报名序号(BMXH_)"},
            {key: "bmdCode", code: "BMD", required: true, label: "报名点(BMD_)"},
            {key: "name", code: "XM", label: "姓名(XM_)"},
            {key: "idCard", code: "SFZH", required: true, label: "身份证号码(SFZH_)"},
            {key: "inlandClass", code: "NDBXS", label: "内地班学生(NDBXS_)"},
            {key: "specialTest", code: "CJTZSCS", label: "参加特招生测试(CJTZSCS_)"}
        ]
    };

    var paymentSchema = {
        kind: "payment",
        label: "高考报名缴费信息",
        matchMode: "exact",
        allowPrefix: true,
        fields: [
            {key: "orderNo", header: "缴款单号", required: true, aliases: ["缴款单号", "缴款书号", "缴款单编号"]},
            {key: "name", header: "姓名", aliases: ["姓名", "考生姓名"]},
            {key: "idCard", header: "证件号码", required: true, aliases: ["证件号码", "身份证号码", "身份证号", "证件号"]},
            {key: "payStatus", header: "缴款状态", required: true, aliases: ["缴款状态", "缴费状态"]},
            {key: "amount", header: "缴款金额", aliases: ["缴款金额", "缴费金额"]},
            {key: "projectNo", header: "项目编号", required: true, aliases: ["项目编号", "缴费项目编号", "项目代码"]},
            {key: "auditStatus", header: "审核状态", aliases: ["审核状态", "SHZT"]}
        ]
    };

    var januarySchema = {
        kind: "january",
        label: "1月学选考报名信息",
        matchMode: "exact",
        denyCodes: ["YSTYSK", "CJTZSCS", "SYCFDM", "WDWZ", "SSMZSZXZ", "QYDM"],
        fields: [
            {key: "idCard", header: "SFZH", required: true, aliases: ["SFZH"]},
            {key: "name", header: "XM", aliases: ["XM"]},
            {key: "bmdCode", header: "BMD", required: true, aliases: ["BMD"]},
            {key: "grade", header: "NJ", aliases: ["NJ"]},
            {key: "className", header: "BJ", aliases: ["BJ"]},
            {key: "auditStatus", header: "SHZT", required: true, aliases: ["SHZT"]},
            {key: "languageCode", header: "WYYZ", aliases: ["WYYZ"]},
            {key: "yw", header: "KM_YW", aliases: ["KM_YW"]},
            {key: "sx", header: "KM_SX", aliases: ["KM_SX"]},
            {key: "yy", header: "KM_YY", aliases: ["KM_YY"]},
            {key: "sz", header: "KM_SZ", aliases: ["KM_SZ"]},
            {key: "ls", header: "KM_LS", aliases: ["KM_LS"]},
            {key: "dl", header: "KM_DL", aliases: ["KM_DL"]},
            {key: "wl", header: "KM_WL", aliases: ["KM_WL"]},
            {key: "hx", header: "KM_HX", aliases: ["KM_HX"]},
            {key: "sw", header: "KM_SW", aliases: ["KM_SW"]},
            {key: "js", header: "KM_JS", aliases: ["KM_JS"]}
        ]
    };

    var danduSchema = {
        kind: "dandu",
        label: "单独考试报名信息",
        matchMode: "code",
        denyCodes: ["KM", "YSTYSK", "CJTZSCS", "SYCFDM", "WDWZ"],
        fields: [
            {key: "bmdCode", code: "BMD", required: true, label: "报名点(BMD_)"},
            {key: "category", code: "BKKL", required: true, label: "报考类别(BKKL_)"},
            {key: "otherMajor", code: "QTLZY", label: "其他类专业(QTLZY_)"},
            // 以下字段仅在该文件中存在时使用：有审核状态则过滤，有身份证则去重
            {key: "idCard", code: "SFZH", label: "身份证号码(SFZH_)"},
            {key: "auditStatus", code: "SHZT", label: "审核状态(SHZT_)"},
            {key: "name", code: "XM", label: "姓名(XM_)"}
        ]
    };

    var SCHEMAS = {
        gaokao: gaokaoSchema,
        payment: paymentSchema,
        january: januarySchema,
        dandu: danduSchema
    };

    /* ------------------------------------------------------------------ */
    /* 表头匹配                                                            */
    /* ------------------------------------------------------------------ */

    function fieldMatches(header, field, schema) {
        var normalized = normalizeHeaderCell(header);
        if (!normalized) return false;

        if (field.aliases) {
            for (var i = 0; i < field.aliases.length; i++) {
                if (normalizeHeaderCell(field.aliases[i]) === normalized) return true;
            }
        }

        if (schema.matchMode === "code") {
            // 带下划线解释的字段按下划线前的字段代码匹配
            var codePart = normalized.split("_")[0];
            return codePart === String(field.code || "").toUpperCase();
        }

        if (schema.matchMode === "exact") {
            var target = normalizeHeaderCell(field.header || field.code || "");
            if (!target) return false;
            if (normalized === target) return true;
            if (schema.allowPrefix && normalized.indexOf(target) === 0) return true;
            return false;
        }

        return false;
    }

    function resolveColumns(headers, schema) {
        var map = {};
        var labels = {};
        var missingRequired = [];

        schema.fields.forEach(function (field) {
            var found = -1;
            for (var i = 0; i < headers.length; i++) {
                if (fieldMatches(headers[i], field, schema)) {
                    found = i;
                    break;
                }
            }
            if (found >= 0) {
                map[field.key] = found;
                labels[field.key] = stripText(headers[found]);
            } else if (field.required) {
                missingRequired.push(field.label || field.header || field.code);
            }
        });

        return {map: map, labels: labels, missingRequired: missingRequired};
    }

    /* ------------------------------------------------------------------ */
    /* 读取工作簿                                                          */
    /* ------------------------------------------------------------------ */

    function readWorkbook(file) {
        return new Promise(function (resolve, reject) {
            if (!file) {
                reject(new Error("未选择文件"));
                return;
            }
            if (typeof XLSX === "undefined") {
                reject(new Error("SheetJS 未加载，请检查 lib/sheetjs 依赖是否可用"));
                return;
            }
            var bufferPromise;
            try {
                bufferPromise = file.arrayBuffer();
            } catch (err) {
                reject(new Error("当前浏览器不支持读取本地文件，请使用最新版 Chrome 或 Edge"));
                return;
            }
            bufferPromise.then(function (buffer) {
                var workbook;
                try {
                    workbook = XLSX.read(buffer, {type: "array", cellDates: false});
                } catch (err) {
                    reject(new Error("文件解析失败，可能不是有效的 Excel 文件（" + (err && err.message ? err.message : "未知原因") + "）"));
                    return;
                }
                var sheetNames = (workbook && workbook.SheetNames) || [];
                if (!sheetNames.length) {
                    reject(new Error("文件中没有可读取的工作表"));
                    return;
                }
                resolve({workbook: workbook, sheetNames: sheetNames});
            }).catch(function () {
                reject(new Error("文件读取失败，请重新选择文件"));
            });
        });
    }

    function readMatrix(workbook, sheetName) {
        if (!workbook || !workbook.Sheets) return [];
        if (!workbook.__matrixCache) workbook.__matrixCache = {};
        if (workbook.__matrixCache[sheetName]) return workbook.__matrixCache[sheetName];

        var sheet = workbook.Sheets[sheetName];
        if (!sheet) return [];
        var matrix = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            defval: "",
            raw: true,
            blankrows: true
        });
        workbook.__matrixCache[sheetName] = matrix;
        return matrix;
    }

    /**
     * 表头是否命中其他文件类型的特征字段（用于阻止把文件误解析为另一种类型）。
     */
    function isDeniedRow(headers, schema) {
        if (!schema.denyCodes || !schema.denyCodes.length) return false;
        return headers.some(function (header) {
            if (!header) return false;
            var codePart = header.split("_")[0];
            return schema.denyCodes.indexOf(codePart) !== -1;
        });
    }

    /**
     * 扫描所有工作表前 20 行，返回可用的候选表头。
     */
    function detectCandidates(workbook, schema) {
        var candidates = [];
        var sheetNames = workbook.SheetNames || [];

        sheetNames.forEach(function (sheetName) {
            var matrix;
            try {
                matrix = readMatrix(workbook, sheetName);
            } catch (err) {
                return;
            }
            var limit = Math.min(matrix.length, MAX_HEADER_SCAN);
            for (var r = 0; r < limit; r++) {
                var row = matrix[r] || [];
                var headers = row.map(normalizeHeaderCell);
                var hasAny = headers.some(function (h) {
                    return !!h;
                });
                if (!hasAny) continue;
                if (isDeniedRow(headers, schema)) continue;

                var resolved = resolveColumns(headers, schema);
                if (!resolved.missingRequired.length) {
                    candidates.push({
                        sheetName: sheetName,
                        headerRowIndex: r,
                        columns: resolved,
                        labels: resolved.labels
                    });
                    break;
                }
            }
        });

        return candidates;
    }

    /**
     * 取表头行之后的数据行，保留原始行序号用于异常定位。
     */
    function sliceRows(matrix, headerRowIndex) {
        var start = headerRowIndex + 1;
        return {
            dataRows: matrix.slice(start),
            firstDataRowNumber: start + 1
        };
    }

    global.ExcelReader = {
        SCHEMAS: SCHEMAS,
        normalizeHeaderCell: normalizeHeaderCell,
        stripText: stripText,
        readWorkbook: readWorkbook,
        readMatrix: readMatrix,
        detectCandidates: detectCandidates,
        sliceRows: sliceRows
    };
})(window);
