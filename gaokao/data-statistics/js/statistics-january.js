/**
 * 1月学选考报名信息统计（按报名点、按县区汇总）。
 */
(function (global) {
    "use strict";

    var DN = global.DataNormalizer;

    var AUDIT_RULE = "仅统计审核状态为 3 或“审核通过”的报名记录，并按身份证号码去重。";
    var PAYMENT_RULE = "1月学选考缴费信息本期不处理，统计不受缴费状态影响。";

    var LANGUAGE_CODE_SET = {};
    DN.LANGUAGES.forEach(function (item) {
        LANGUAGE_CODE_SET[item.code] = true;
    });

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

    function buildColumns(includeBmd, availability, hasUnknownLanguage) {
        var columns = [
            {key: "countyCode", label: "县区代码", type: "text"},
            {key: "countyName", label: "县区名称", type: "text"}
        ];
        if (includeBmd) {
            columns.push({key: "bmdCode", label: "报名点代码", type: "text"});
            columns.push({key: "bmdName", label: "报名点名称", type: "text"});
        }

        columns.push({key: "yw", label: "语文学考", type: "number", available: availability.yw});
        columns.push({key: "sx", label: "数学学考", type: "number", available: availability.sx});

        DN.LANGUAGES.forEach(function (item) {
            columns.push({key: "lang_" + item.code, label: item.label, type: "number", available: availability.language});
        });
        if (hasUnknownLanguage) {
            columns.push({key: "lang_unknown", label: "未知语种", type: "number", available: availability.language});
        }
        columns.push({key: "foreignTotal", label: "外语合计", type: "number", available: availability.foreign});

        DN.JANUARY_SUBJECTS.forEach(function (item) {
            columns.push({key: "xk_" + item.key, label: item.label + "学考", type: "number", available: availability.subjects[item.key]});
        });
        DN.JANUARY_SUBJECTS.forEach(function (item) {
            columns.push({key: "xx_" + item.key, label: item.label + "选考", type: "number", available: availability.subjects[item.key]});
        });

        return columns;
    }

    function buildRow(group, options) {
        var row = {
            countyCode: group.countyCode,
            countyName: options.resolver.county(group.countyCode),
            yw: options.availability.yw ? group.yw.size : null,
            sx: options.availability.sx ? group.sx.size : null,
            foreignTotal: options.availability.foreign ? group.foreign.size : null,
            lang_unknown: options.availability.language && options.hasUnknownLanguage ? sizeOf(group.languages, "unknown") : null
        };
        if (options.includeBmd) {
            row.bmdCode = group.bmdCode;
            row.bmdName = options.resolver.bmd(group.bmdCode);
        }
        DN.LANGUAGES.forEach(function (item) {
            row["lang_" + item.code] = options.availability.language ? sizeOf(group.languages, item.code) : null;
        });
        DN.JANUARY_SUBJECTS.forEach(function (item) {
            row["xk_" + item.key] = options.availability.subjects[item.key] ? sizeOf(group.xuekao, item.key) : null;
            row["xx_" + item.key] = options.availability.subjects[item.key] ? sizeOf(group.xuankao, item.key) : null;
        });
        return row;
    }

    function build(ctx) {
        var january = ctx.january;
        var mapping = ctx.mapping;
        var diag = DN.createCollector();
        var resolver = DN.createNameResolver(mapping, diag);

        var presence = january.presence;
        var availability = {
            yw: !!presence.yw,
            sx: !!presence.sx,
            foreign: !!presence.yy,
            language: !!(presence.languageCode && presence.yy),
            subjects: {}
        };
        DN.JANUARY_SUBJECTS.forEach(function (item) {
            availability.subjects[item.key] = !!presence.subjects[item.key];
        });

        /* ---------------- 1. 审核过滤 ---------------- */
        var passedRecords = january.records.filter(function (record) {
            return DN.isAuditPassed(record.auditStatus);
        });

        /* ---------------- 2. 按身份证去重并合并科目标记 ---------------- */
        var personMap = new Map();
        passedRecords.forEach(function (record) {
            if (!record.idCard) return;
            var person = personMap.get(record.idCard);
            if (!person) {
                person = {
                    id: record.idCard,
                    bmdCodes: new Set(),
                    yw: false,
                    sx: false,
                    foreign: false,
                    language: "",
                    xuekao: {},
                    xuankao: {}
                };
                personMap.set(record.idCard, person);
            }
            if (record.bmdCode) person.bmdCodes.add(record.bmdCode);

            if (record.subjects.yw) person.yw = true;
            if (record.subjects.sx) person.sx = true;
            if (record.subjects.yy) {
                person.foreign = true;
                if (!person.language && record.languageCode) person.language = record.languageCode;
            }

            DN.JANUARY_SUBJECTS.forEach(function (item) {
                var value = record.subjects[item.key];
                if (value === "学考") person.xuekao[item.key] = true;
                else if (value === "选考") person.xuankao[item.key] = true;
            });
        });

        /* ---------------- 3. 排除跨报名点冲突 ---------------- */
        var persons = [];
        personMap.forEach(function (person) {
            if (person.bmdCodes.size > 1) {
                diag.add(
                    "warning",
                    "JAN_BMD_CONFLICT",
                    "同一身份证出现在多个报名点，已排除该人员（避免重复计数）",
                    DN.maskIdCard(person.id) + "（" + Array.from(person.bmdCodes).join("、") + "）"
                );
                return;
            }
            if (person.bmdCodes.size === 0) {
                diag.add("warning", "JAN_NO_BMD", "报名点代码为空，该人员未纳入统计", DN.maskIdCard(person.id));
                return;
            }
            person.bmdCode = Array.from(person.bmdCodes)[0];
            persons.push(person);
        });

        /* ---------------- 4. 分组汇总 ---------------- */
        var bmdGroups = new Map();
        var countyGroups = new Map();
        var overall = createGroup("", "");

        var unknownLanguageCount = 0;

        function createGroup(bmdCode, countyCode) {
            return {
                bmdCode: bmdCode,
                countyCode: countyCode,
                yw: new Set(),
                sx: new Set(),
                foreign: new Set(),
                languages: new Map(),
                xuekao: new Map(),
                xuankao: new Map()
            };
        }

        function ensureGroup(store, key, bmdCode, countyCode) {
            var group = store.get(key);
            if (!group) {
                group = createGroup(bmdCode, countyCode);
                store.set(key, group);
            }
            return group;
        }

        persons.forEach(function (person) {
            var countyCode = DN.countyCodeOf(person.bmdCode);
            var bmdGroup = ensureGroup(bmdGroups, person.bmdCode, person.bmdCode, countyCode);
            var countyGroup = ensureGroup(countyGroups, countyCode, "", countyCode);

            var targets = [bmdGroup, countyGroup, overall];

            var languageKey = "";
            if (person.foreign) {
                if (person.language && LANGUAGE_CODE_SET[person.language]) languageKey = person.language;
                else if (person.language === "0") languageKey = "";
                else if (!person.language) languageKey = "";
                else languageKey = "unknown";
            }
            if (languageKey === "unknown") unknownLanguageCount += 1;

            targets.forEach(function (group) {
                if (person.yw) group.yw.add(person.id);
                if (person.sx) group.sx.add(person.id);
                if (person.foreign) {
                    group.foreign.add(person.id);
                    if (languageKey && languageKey !== "unknown") addTo(group.languages, languageKey, person.id);
                    else if (languageKey === "unknown") addTo(group.languages, "unknown", person.id);
                }
                DN.JANUARY_SUBJECTS.forEach(function (item) {
                    if (person.xuekao[item.key]) addTo(group.xuekao, item.key, person.id);
                    if (person.xuankao[item.key]) addTo(group.xuankao, item.key, person.id);
                });
            });
        });

        if (unknownLanguageCount > 0) {
            diag.addMany(
                "warning",
                "JAN_LANG_UNKNOWN",
                "KM_YY 非空但外语语种代码未知，已计入“未知语种”列并保留在外语合计中",
                unknownLanguageCount,
                []
            );
            if (!presence.languageCode) {
                diag.add("warning", "JAN_LANG_FIELD_MISSING", "缺少 WYYZ 字段，无法按语种分列，全部外语考生计入“未知语种”。");
            }
        }

        /* ---------------- 5. 生成表格 ---------------- */
        var hasUnknownLanguage = unknownLanguageCount > 0;

        var bmdRows = Array.from(bmdGroups.values())
            .sort(function (a, b) {
                return compareCode(a.countyCode, b.countyCode) || compareCode(a.bmdCode, b.bmdCode);
            })
            .map(function (group) {
                return buildRow(group, {
                    resolver: resolver,
                    includeBmd: true,
                    availability: availability,
                    hasUnknownLanguage: hasUnknownLanguage
                });
            });

        var countyRows = Array.from(countyGroups.values())
            .sort(function (a, b) {
                return compareCode(a.countyCode, b.countyCode);
            })
            .map(function (group) {
                return buildRow(group, {
                    resolver: resolver,
                    includeBmd: false,
                    availability: availability,
                    hasUnknownLanguage: hasUnknownLanguage
                });
            });

        var totalResolver = {
            bmd: function () { return "合计"; },
            county: function () { return "—"; }
        };

        var totalRowBmd = buildRow(overall, {
            resolver: totalResolver,
            includeBmd: true,
            availability: availability,
            hasUnknownLanguage: hasUnknownLanguage
        });
        totalRowBmd.bmdCode = "合计";

        var totalRowCounty = buildRow(overall, {
            resolver: totalResolver,
            includeBmd: false,
            availability: availability,
            hasUnknownLanguage: hasUnknownLanguage
        });
        totalRowCounty.countyCode = "合计";

        var notes = [];
        notes.push({
            level: "info",
            text: "审核通过记录 " + passedRecords.length + " 条，按身份证去重后 " + personMap.size + " 人；最终计入统计 " + persons.length + " 人。"
        });
        if (hasUnknownLanguage) {
            notes.push({level: "warning", text: "存在外语语种代码无法识别的考生，已单列“未知语种”，并计入外语合计。"});
        }

        var baseMeta = {
            source: january.fileName || "1月学选考报名信息",
            auditRule: AUDIT_RULE,
            paymentRule: PAYMENT_RULE,
            notes: notes
        };

        var tables = [
            Object.assign({
                key: "january-bmd",
                title: "1月学选考报名统计（按报名点）",
                columns: buildColumns(true, availability, hasUnknownLanguage),
                rows: bmdRows,
                totalRow: totalRowBmd
            }, baseMeta),
            Object.assign({
                key: "january-county",
                title: "1月学选考报名统计（按县区汇总）",
                columns: buildColumns(false, availability, hasUnknownLanguage),
                rows: countyRows,
                totalRow: totalRowCounty
            }, baseMeta)
        ];

        return {tables: tables, diagnostics: diag.toArray()};
    }

    global.JanuaryStats = {build: build};
})(window);
