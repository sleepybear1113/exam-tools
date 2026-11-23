/**
 * 高考报名缴费信息独立统计（按报名点）。
 */
(function (global) {
    "use strict";

    var DN = global.DataNormalizer;

    var AUDIT_RULE = "缴费文件不含审核状态字段时不做审核过滤，仅按缴款状态统计。";
    var PAYMENT_RULE = "同一身份证 + 项目类别 + 报名点去重；任一条“已缴款”即视为已缴款；总人数为普通类（pt）项目的去重身份证人数。";

    function sizeOf(map, key) {
        var set = map.get(key);
        return set ? set.size : 0;
    }

    function addTo(map, key, value) {
        var set = map.get(key);
        if (!set) {
            set = new Set();
            map.set(key, set);
        }
        set.add(value);
    }

    function compareCode(a, b) {
        return String(a || "").localeCompare(String(b || ""), "zh-CN", {numeric: true});
    }

    function buildColumns() {
        var columns = [
            {key: "countyCode", label: "县区代码", type: "text"},
            {key: "countyName", label: "县区名称", type: "text"},
            {key: "bmdCode", label: "报名点代码", type: "text"},
            {key: "bmdName", label: "报名点名称", type: "text"},
            {key: "total", label: "总人数", type: "number"}
        ];
        // 先排列所有类别的“已缴费”列，再排列“未缴费”列
        DN.PAYMENT_CATEGORIES.forEach(function (category) {
            columns.push({key: "paid_" + category.code, label: category.label + "_已缴费", type: "number"});
        });
        DN.PAYMENT_CATEGORIES.forEach(function (category) {
            columns.push({key: "unpaid_" + category.code, label: category.label + "_未缴费", type: "number"});
        });
        return columns;
    }

    function buildRow(group, resolver) {
        var row = {
            countyCode: group.countyCode,
            countyName: resolver.county(group.countyCode),
            bmdCode: group.bmdCode,
            bmdName: resolver.bmd(group.bmdCode),
            total: group.total.size
        };
        DN.PAYMENT_CATEGORIES.forEach(function (category) {
            row["paid_" + category.code] = sizeOf(group.paid, category.code);
            row["unpaid_" + category.code] = sizeOf(group.unpaid, category.code);
        });
        return row;
    }

    function build(ctx) {
        var payment = ctx.payment;
        var mapping = ctx.mapping;
        var diag = DN.createCollector();
        var resolver = DN.createNameResolver(mapping, diag);

        var auditFiltered = false;
        var records = payment.records;

        if (payment.presence.auditStatus) {
            auditFiltered = true;
            records = records.filter(function (record) {
                return DN.isAuditPassed(record.auditStatus);
            });
        }

        /* ---------------- 1. 按 身份证 + 类别 归并缴款状态 ---------------- */
        var cells = new Map();
        records.forEach(function (record) {
            if (!record.idCard) return;
            if (!record.bmdCode) return;
            if (!record.categoryKnown || !record.categoryCode) return;

            var key = record.idCard + "|" + record.categoryCode;
            var cell = cells.get(key);
            if (!cell) {
                cell = {
                    id: record.idCard,
                    category: record.categoryCode,
                    bmdCodes: new Set(),
                    state: "unknown"
                };
                cells.set(key, cell);
            }
            cell.bmdCodes.add(record.bmdCode);

            if (record.payState === "paid") {
                cell.state = "paid";
            } else if (record.payState === "unpaid") {
                if (cell.state !== "paid") cell.state = "unpaid";
            }
        });

        /* ---------------- 2. 分组汇总 ---------------- */
        var groups = new Map();
        var overall = {
            total: new Set(),
            paid: new Map(),
            unpaid: new Map()
        };
        var unknownCells = 0;
        var conflictCells = 0;

        cells.forEach(function (cell) {
            if (cell.bmdCodes.size > 1) {
                conflictCells += 1;
                diag.add(
                    "warning",
                    "PAY_BMD_CONFLICT",
                    "同一身份证、同一缴费类别对应多个报名点，该类别未计入任何报名点",
                    DN.maskIdCard(cell.id) + "（" + cell.category + "：" + Array.from(cell.bmdCodes).join("、") + "）"
                );
                return;
            }

            var bmdCode = Array.from(cell.bmdCodes)[0];
            var group = groups.get(bmdCode);
            if (!group) {
                group = {bmdCode: bmdCode, countyCode: DN.countyCodeOf(bmdCode), total: new Set(), paid: new Map(), unpaid: new Map()};
                groups.set(bmdCode, group);
            }

            // 总人数：普通类（pt）项目中去重身份证，不区分缴款状态
            if (cell.category === "pt") {
                group.total.add(cell.id);
                overall.total.add(cell.id);
            }

            if (cell.state === "paid") {
                addTo(group.paid, cell.category, cell.id);
                addTo(overall.paid, cell.category, cell.id);
            } else if (cell.state === "unpaid") {
                addTo(group.unpaid, cell.category, cell.id);
                addTo(overall.unpaid, cell.category, cell.id);
            } else {
                unknownCells += 1;
            }
        });

        if (unknownCells > 0) {
            diag.addMany(
                "warning",
                "PAY_CELL_UNKNOWN_STATUS",
                "存在“身份证 + 类别”组合的所有记录缴款状态均无法识别，未计入已缴费/未缴费列",
                unknownCells,
                []
            );
        }

        /* ---------------- 3. 普通类前置校验 ---------------- */
        var ptPaid = new Set();
        var ptKnown = new Set();
        cells.forEach(function (cell) {
            if (cell.bmdCodes.size > 1) return;
            if (cell.category !== "pt") return;
            ptKnown.add(cell.id);
            if (cell.state === "paid") ptPaid.add(cell.id);
        });

        cells.forEach(function (cell) {
            if (cell.bmdCodes.size > 1) return;
            if (cell.category === "pt") return;
            if (cell.state !== "paid") return;
            if (!ptKnown.has(cell.id)) {
                diag.add("warning", "PAY_NONPT_WITHOUT_PT", "非普通类项目已缴款但缺少普通类（pt）记录，原始缴款记录仍保留在统计中", DN.maskIdCard(cell.id));
            } else if (!ptPaid.has(cell.id)) {
                diag.add("warning", "PAY_NONPT_PT_UNPAID", "非普通类项目已缴款但普通类（pt）未缴款，原始缴款记录仍保留在统计中", DN.maskIdCard(cell.id));
            }
        });

        /* ---------------- 4. 生成表格 ---------------- */
        var rows = Array.from(groups.values())
            .sort(function (a, b) {
                return compareCode(a.countyCode, b.countyCode) || compareCode(a.bmdCode, b.bmdCode);
            })
            .map(function (group) {
                return buildRow(group, resolver);
            });

        var totalRow = {
            countyCode: "合计",
            countyName: "—",
            bmdCode: "—",
            bmdName: "合计",
            total: overall.total.size
        };
        DN.PAYMENT_CATEGORIES.forEach(function (category) {
            totalRow["paid_" + category.code] = sizeOf(overall.paid, category.code);
            totalRow["unpaid_" + category.code] = sizeOf(overall.unpaid, category.code);
        });

        var notes = [];
        notes.push({
            level: "info",
            text: auditFiltered
                ? "缴费文件中识别到审核状态列，已按审核状态过滤后再统计。"
                : "缴费文件未包含审核状态列，独立缴费统计未做审核过滤（不代表审核结果）。"
        });
        notes.push({level: "info", text: "有效缴费记录 " + records.length + " 条，识别到报名点 " + groups.size + " 个。"});
        if (conflictCells > 0) {
            notes.push({level: "warning", text: "有 " + conflictCells + " 个“身份证 + 类别”组合对应多个报名点，已按要求排除。"});
        }

        var table = {
            key: "payment-bmd",
            title: "高考报名缴费统计（按报名点）",
            source: payment.fileName || "高考报名缴费信息",
            auditRule: AUDIT_RULE,
            paymentRule: PAYMENT_RULE,
            notes: notes,
            columns: buildColumns(),
            rows: rows,
            totalRow: totalRow
        };

        return {tables: [table], diagnostics: diag.toArray()};
    }

    global.PaymentStats = {build: build};
})(window);
