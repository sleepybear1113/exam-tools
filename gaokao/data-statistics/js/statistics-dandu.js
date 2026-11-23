/**
 * 单独考试报名信息统计（按报名点、按县区汇总）。
 * 口径：总人数 + 每个报考类别（BKKL_）一列报名人数；
 * BKKL_ = 17 时按 QTLZY_（其他类专业）细分为 17-1 安全防范、17-2 体育 等。
 */
(function (global) {
    "use strict";

    var DN = global.DataNormalizer;

    function compareCode(a, b) {
        return String(a || "").localeCompare(String(b || ""), "zh-CN", {numeric: true});
    }

    function build(ctx) {
        var data = ctx.dandu;
        var mapping = ctx.mapping;
        var diag = DN.createCollector();
        var resolver = DN.createNameResolver(mapping, diag);

        var useId = !!data.presence.idCard;
        var filterAudit = !!data.presence.auditStatus;

        var auditRule = filterAudit
            ? "文件含审核状态字段，仅统计审核状态为 3 或“审核通过”的记录。"
            : "文件不含审核状态字段，未做审核过滤。";
        var dedupeRule = useId
            ? "文件含身份证号码字段，按身份证去重（同一人同一类别只计 1 人，跨报名点冲突排除）。"
            : "文件不含身份证号码字段，按记录行数统计，未做去重。";

        var records = filterAudit
            ? data.records.filter(function (record) {
                return DN.isAuditPassed(record.auditStatus);
            })
            : data.records;

        /* ---------------- 1. 类别列（按全部记录确定，保证两张表列一致） ---------------- */
        var categoryLabels = new Map();
        data.records.forEach(function (record) {
            var key = record.categoryKey || "";
            if (!categoryLabels.has(key)) categoryLabels.set(key, DN.danduCategoryLabel(key));
        });
        var categoryKeys = Array.from(categoryLabels.keys()).sort(compareCode);

        /* ---------------- 2. 统计单元 ---------------- */
        var units = [];

        if (useId) {
            var personMap = new Map();
            records.forEach(function (record) {
                if (!record.idCard) return; // 解析阶段已记录异常
                var person = personMap.get(record.idCard);
                if (!person) {
                    person = {id: record.idCard, bmdCodes: new Set(), categories: new Set()};
                    personMap.set(record.idCard, person);
                }
                if (record.bmdCode) person.bmdCodes.add(record.bmdCode);
                person.categories.add(record.categoryKey || "");
            });

            personMap.forEach(function (person) {
                if (person.bmdCodes.size > 1) {
                    diag.add(
                        "warning",
                        "DANDU_BMD_CONFLICT",
                        "同一身份证出现在多个报名点，已排除该人员（避免重复计数）",
                        DN.maskIdCard(person.id) + "（" + Array.from(person.bmdCodes).join("、") + "）"
                    );
                    return;
                }
                if (!person.bmdCodes.size) {
                    diag.add("warning", "DANDU_NO_BMD", "报名点代码为空，该人员未纳入统计", DN.maskIdCard(person.id));
                    return;
                }
                units.push({
                    bmdCode: Array.from(person.bmdCodes)[0],
                    categories: Array.from(person.categories),
                    id: person.id
                });
            });
        } else {
            records.forEach(function (record) {
                if (!record.bmdCode) return; // 解析阶段已记录异常
                units.push({bmdCode: record.bmdCode, categories: [record.categoryKey || ""]});
            });
        }

        /* ---------------- 3. 分组累加 ---------------- */
        function createBucket() {
            return {total: useId ? new Set() : 0, categories: new Map()};
        }

        function addTotal(bucket, unit) {
            if (useId) bucket.total.add(unit.id);
            else bucket.total += 1;
        }

        function addCategory(bucket, key, unit) {
            var entry = bucket.categories.get(key);
            if (!entry) {
                entry = {value: useId ? new Set() : 0};
                bucket.categories.set(key, entry);
            }
            if (useId) entry.value.add(unit.id);
            else entry.value += 1;
        }

        function totalOf(bucket) {
            return useId ? bucket.total.size : bucket.total;
        }

        function categoryOf(bucket, key) {
            var entry = bucket.categories.get(key);
            if (!entry) return 0;
            return useId ? entry.value.size : entry.value;
        }

        var bmdGroups = new Map();
        var countyGroups = new Map();
        var overall = createBucket();

        units.forEach(function (unit) {
            var countyCode = DN.countyCodeOf(unit.bmdCode);
            var bmdBucket = bmdGroups.get(unit.bmdCode);
            if (!bmdBucket) {
                bmdBucket = createBucket();
                bmdGroups.set(unit.bmdCode, bmdBucket);
            }
            var countyBucket = countyGroups.get(countyCode);
            if (!countyBucket) {
                countyBucket = createBucket();
                countyGroups.set(countyCode, countyBucket);
            }

            [bmdBucket, countyBucket, overall].forEach(function (bucket) {
                addTotal(bucket, unit);
                unit.categories.forEach(function (key) {
                    addCategory(bucket, key, unit);
                });
            });
        });

        /* ---------------- 4. 生成表格 ---------------- */
        function buildColumns(includeBmd) {
            var columns = [
                {key: "countyCode", label: "县区代码", type: "text"},
                {key: "countyName", label: "县区名称", type: "text"}
            ];
            if (includeBmd) {
                columns.push({key: "bmdCode", label: "报名点代码", type: "text"});
                columns.push({key: "bmdName", label: "报名点名称", type: "text"});
            }
            columns.push({key: "total", label: "总人数", type: "number"});
            categoryKeys.forEach(function (key) {
                columns.push({key: "cat_" + key, label: categoryLabels.get(key), type: "number"});
            });
            return columns;
        }

        function buildRow(bucket, countyCode, bmdCode, includeBmd) {
            var row = {
                countyCode: countyCode,
                countyName: resolver.county(countyCode),
                total: totalOf(bucket)
            };
            if (includeBmd) {
                row.bmdCode = bmdCode;
                row.bmdName = resolver.bmd(bmdCode);
            }
            categoryKeys.forEach(function (key) {
                row["cat_" + key] = categoryOf(bucket, key);
            });
            return row;
        }

        var bmdRows = Array.from(bmdGroups.keys())
            .sort(function (a, b) {
                return compareCode(DN.countyCodeOf(a), DN.countyCodeOf(b)) || compareCode(a, b);
            })
            .map(function (bmdCode) {
                return buildRow(bmdGroups.get(bmdCode), DN.countyCodeOf(bmdCode), bmdCode, true);
            });

        var countyRows = Array.from(countyGroups.keys())
            .sort(compareCode)
            .map(function (countyCode) {
                return buildRow(countyGroups.get(countyCode), countyCode, "", false);
            });

        var totalRowBmd = buildRow(overall, "", "合计", true);
        totalRowBmd.bmdCode = "合计";
        totalRowBmd.bmdName = "合计";
        totalRowBmd.countyName = "—";

        var totalRowCounty = buildRow(overall, "合计", "", false);
        totalRowCounty.countyCode = "合计";
        totalRowCounty.countyName = "—";

        var notes = [];
        notes.push({
            level: "info",
            text: "有效记录 " + records.length + " 条，涉及报名点 " + bmdGroups.size + " 个、报考类别 " + categoryKeys.length + " 类" +
                (useId ? "；按身份证去重后 " + units.length + " 人" : "（按记录行数统计）") + "。"
        });
        if (!useId) {
            notes.push({
                level: "warning",
                text: "未在文件中识别到身份证号码字段（SFZH_），按记录行数统计，若同一考生存在多条记录会被重复计数，请核对源数据。"
            });
        }
        if (!filterAudit) {
            notes.push({
                level: "info",
                text: "未在文件中识别到审核状态字段（SHZT_），统计包含全部记录。"
            });
        }
        var otherMajorCount = categoryKeys.filter(function (key) {
            return String(key).indexOf(DN.OTHER_MAJOR_PARENT + "-") === 0;
        }).length;
        if (otherMajorCount) {
            notes.push({
                level: "info",
                text: "报考类别 17（其他类）已按 QTLZY_ 细分为 " + otherMajorCount + " 列（17-1 安全防范、17-2 体育、17-4 学前教育、17-5 退役士兵高职招生、17-6 汽车专业、17-7 内地西藏新疆中职班）。"
            });
        }

        var baseMeta = {
            source: data.fileName || "单独考试报名信息",
            auditRule: auditRule,
            paymentRule: dedupeRule,
            paymentRuleLabel: "去重口径",
            notes: notes
        };

        var tables = [
            Object.assign({
                key: "dandu-bmd",
                title: "单独考试报名统计（按报名点）",
                columns: buildColumns(true),
                rows: bmdRows,
                totalRow: totalRowBmd
            }, baseMeta),
            Object.assign({
                key: "dandu-county",
                title: "单独考试报名统计（按县区汇总）",
                columns: buildColumns(false),
                rows: countyRows,
                totalRow: totalRowCounty
            }, baseMeta)
        ];

        return {tables: tables, diagnostics: diag.toArray()};
    }

    global.DanduStats = {build: build};
})(window);
