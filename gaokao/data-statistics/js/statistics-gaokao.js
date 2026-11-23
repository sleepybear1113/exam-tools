/**
 * 高考报名信息统计（含与缴费信息的交叉过滤）。
 */
(function (global) {
    "use strict";

    var DN = global.DataNormalizer;

    var AUDIT_RULE = "仅统计审核状态为 3 或“审核通过”的报名记录，并按身份证号码去重。";
    var PAYMENT_RULE_ON = "已按缴费信息交叉过滤：普通类（pt）已缴款为纳入总人数前提；术科按对应缴费项目已缴款判断；戏曲类省际联考（39）不要求缴费项目。";
    var PAYMENT_RULE_OFF = "未提供缴费信息，仅按报名信息统计，结果可能包含未缴费考生。";

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

    function buildColumns(options) {
        var columns = [
            {key: "countyCode", label: "县区代码", type: "text"},
            {key: "countyName", label: "县区名称", type: "text"}
        ];
        if (options.withBmd) {
            columns.push({key: "bmdCode", label: "报名点代码", type: "text"});
            columns.push({key: "bmdName", label: "报名点名称", type: "text"});
        }
        columns.push({key: "total", label: "人数", type: "number"});
        DN.ARTS_CATEGORIES.forEach(function (category) {
            columns.push({key: "arts_" + category.code, label: category.label, type: "number", available: options.artsAvailable});
        });
        columns.push({key: "special", label: "特招生测试人数", type: "number", available: options.specialAvailable});
        return columns;
    }

    function buildRow(group, options) {
        var row = {
            countyCode: group.countyCode,
            countyName: options.resolver.county(group.countyCode),
            total: group.total.size,
            special: options.specialAvailable ? group.special.size : null
        };
        if (options.withBmd) {
            row.bmdCode = group.bmdCode;
            row.bmdName = options.resolver.bmd(group.bmdCode);
        }
        DN.ARTS_CATEGORIES.forEach(function (category) {
            row["arts_" + category.code] = options.artsAvailable ? sizeOf(group.arts, category.code) : null;
        });
        return row;
    }

    function build(ctx) {
        var gaokao = ctx.gaokao;
        var payment = ctx.payment;
        var mapping = ctx.mapping;
        var diag = DN.createCollector();
        var resolver = DN.createNameResolver(mapping, diag);

        var hasPayment = !!payment;
        var artsAvailable = !!gaokao.presence.arts;
        var specialAvailable = !!gaokao.presence.specialTest;

        /* ---------------- 1. 审核过滤 ---------------- */
        var passedRecords = gaokao.records.filter(function (record) {
            return DN.isAuditPassed(record.auditStatus);
        });

        /* ---------------- 2. 按身份证去重 ---------------- */
        var personMap = new Map();
        passedRecords.forEach(function (record) {
            if (!record.idCard) return;
            var person = personMap.get(record.idCard);
            if (!person) {
                person = {
                    id: record.idCard,
                    name: record.name || "",
                    bmdCodes: new Set(),
                    arts: new Set(),
                    special: false,
                    rows: 0
                };
                personMap.set(record.idCard, person);
            }
            if (!person.name && record.name) person.name = record.name;
            if (record.bmdCode) person.bmdCodes.add(record.bmdCode);
            record.artsCodes.forEach(function (code) {
                person.arts.add(code);
            });
            if (record.specialHit) person.special = true;
            person.rows += 1;
        });

        /* ---------------- 3. 排除跨报名点冲突 ---------------- */
        var persons = [];
        personMap.forEach(function (person) {
            if (person.bmdCodes.size > 1) {
                diag.add(
                    "warning",
                    "GK_BMD_CONFLICT",
                    "同一身份证出现在多个报名点，已排除该人员（避免重复计数）",
                    DN.maskIdCard(person.id) + "（" + Array.from(person.bmdCodes).join("、") + "）"
                );
                return;
            }
            if (person.bmdCodes.size === 0) {
                diag.add("warning", "GK_NO_BMD", "报名点代码为空，该人员未纳入统计", DN.maskIdCard(person.id));
                return;
            }
            person.bmdCode = Array.from(person.bmdCodes)[0];
            persons.push(person);
        });

        /* ---------------- 4. 缴费索引 ---------------- */
        var payIndex = new Map();
        if (hasPayment) {
            payment.records.forEach(function (record) {
                if (!record.idCard || record.idCorrupted) return;
                if (!record.categoryKnown) return;
                var entry = payIndex.get(record.idCard);
                if (!entry) {
                    entry = {categories: new Map()};
                    payIndex.set(record.idCard, entry);
                }
                var previous = entry.categories.get(record.categoryCode);
                if (previous === "paid") return;
                if (record.payState === "paid") entry.categories.set(record.categoryCode, "paid");
                else if (record.payState === "unpaid") entry.categories.set(record.categoryCode, "unpaid");
                else if (!previous) entry.categories.set(record.categoryCode, "unknown");
            });
        }

        /* ---------------- 4.1 缴费项目与报名术科对应关系核对 ---------------- */
        if (hasPayment) {
            var artsPresence = new Map();
            persons.forEach(function (person) {
                person.arts.forEach(function (code) {
                    artsPresence.set(code, (artsPresence.get(code) || 0) + 1);
                });
            });

            Object.keys(DN.ARTS_TO_PAYMENT).forEach(function (artsCode) {
                var payCategory = DN.ARTS_TO_PAYMENT[artsCode];
                if (artsPresence.get(artsCode)) return;

                var paidIds = new Set();
                payIndex.forEach(function (entry, id) {
                    if (entry.categories.get(payCategory) === "paid") paidIds.add(id);
                });
                if (!paidIds.size) return;

                var artsLabel = artsCode;
                DN.ARTS_CATEGORIES.forEach(function (item) {
                    if (item.code === artsCode) artsLabel = item.label + "（" + item.code + "）";
                });
                var payLabel = payCategory;
                DN.PAYMENT_CATEGORIES.forEach(function (item) {
                    if (item.code === payCategory) payLabel = item.label + "（" + item.code + "）";
                });

                diag.addMany(
                    "warning",
                    "GK_PAYMENT_WITHOUT_ARTS",
                    "缴费文件中有 " + paidIds.size + " 人已缴 " + payLabel + "，但报名信息中没有 " + artsLabel + " 考生，该缴费记录不影响任何术科人数，请核对缴费项目与术科对应关系",
                    paidIds.size,
                    [payCategory + " → " + artsCode]
                );
            });
        }

        /* ---------------- 5. 人员纳入规则 ---------------- */
        var included = [];
        var unpaidArtsExcluded = 0;

        persons.forEach(function (person) {
            var includeTotal = true;
            var payCategories = null;

            if (hasPayment) {
                var pay = payIndex.get(person.id);
                if (!pay) {
                    // 缴费文件中完全没有该身份证记录：视为已缴费
                    includeTotal = true;
                } else {
                    var ptState = pay.categories.get("pt");
                    if (!ptState) {
                        includeTotal = false;
                        diag.add("warning", "GK_PT_MISSING", "存在缴费记录但缺少普通类（pt）记录，不纳入报名总人数及术科人数", DN.maskIdCard(person.id));
                    } else if (ptState === "paid") {
                        includeTotal = true;
                        payCategories = pay.categories;
                    } else if (ptState === "unpaid") {
                        includeTotal = false;
                    } else {
                        includeTotal = false;
                        diag.add("warning", "GK_PT_UNKNOWN", "普通类（pt）缴费状态无法识别，不纳入报名总人数及术科人数", DN.maskIdCard(person.id));
                    }
                }
            }

            if (!includeTotal) return;

            var arts = new Set();
            person.arts.forEach(function (code) {
                if (code === "39") {
                    // 戏曲类省际联考不要求缴费项目
                    arts.add(code);
                    return;
                }
                var payCategory = DN.ARTS_TO_PAYMENT[code];
                if (!payCategory) return; // 未知术科代码已在解析阶段提示
                if (!payCategories) {
                    arts.add(code);
                    return;
                }
                var state = payCategories.get(payCategory);
                if (state === "paid") {
                    arts.add(code);
                } else {
                    unpaidArtsExcluded += 1;
                    diag.add(
                        "warning",
                        "GK_ARTS_PAYMENT_EXCLUDED",
                        "术科类别对应缴费项目未完成缴费，该类别不计入人数（不影响该考生总人数）",
                        code + " ← " + payCategory
                    );
                }
            });

            included.push({id: person.id, bmdCode: person.bmdCode, arts: arts, special: person.special});
        });

        /* ---------------- 6. 分组统计 ---------------- */
        var bmdGroups = new Map();
        var countyGroups = new Map();
        var overall = {total: new Set(), arts: new Map(), special: new Set()};

        function ensureGroup(store, key, countyCode) {
            var group = store.get(key);
            if (!group) {
                group = {bmdCode: key, countyCode: countyCode, total: new Set(), arts: new Map(), special: new Set()};
                store.set(key, group);
            }
            return group;
        }

        included.forEach(function (person) {
            var countyCode = DN.countyCodeOf(person.bmdCode);
            var bmdGroup = ensureGroup(bmdGroups, person.bmdCode, countyCode);
            bmdGroup.total.add(person.id);
            overall.total.add(person.id);
            if (person.special) {
                bmdGroup.special.add(person.id);
            }

            var countyGroup = ensureGroup(countyGroups, countyCode, countyCode);
            countyGroup.total.add(person.id);
            if (person.special) {
                countyGroup.special.add(person.id);
                overall.special.add(person.id);
            }

            person.arts.forEach(function (code) {
                addTo(bmdGroup.arts, code, person.id);
                addTo(countyGroup.arts, code, person.id);
                addTo(overall.arts, code, person.id);
            });
        });

        /* ---------------- 7. 生成表格 ---------------- */
        var bmdRows = Array.from(bmdGroups.values())
            .sort(function (a, b) {
                return compareCode(a.countyCode, b.countyCode) || compareCode(a.bmdCode, b.bmdCode);
            })
            .map(function (group) {
                return buildRow(group, {resolver: resolver, withBmd: true, artsAvailable: artsAvailable, specialAvailable: specialAvailable});
            });

        var countyRows = Array.from(countyGroups.values())
            .sort(function (a, b) {
                return compareCode(a.countyCode, b.countyCode);
            })
            .map(function (group) {
                return buildRow(group, {resolver: resolver, withBmd: false, artsAvailable: artsAvailable, specialAvailable: specialAvailable});
            });

        var totalRowBmd = buildRow(
            {bmdCode: "", countyCode: "", total: overall.total, arts: overall.arts, special: overall.special},
            {resolver: {bmd: function () { return "合计"; }, county: function () { return "—"; }}, withBmd: true, artsAvailable: artsAvailable, specialAvailable: specialAvailable}
        );
        totalRowBmd.countyName = "—";
        totalRowBmd.bmdName = "合计";

        var totalRowCounty = buildRow(
            {countyCode: "", total: overall.total, arts: overall.arts, special: overall.special},
            {resolver: {county: function () { return "—"; }}, withBmd: false, artsAvailable: artsAvailable, specialAvailable: specialAvailable}
        );
        totalRowCounty.countyCode = "合计";

        var notes = [];
        if (!hasPayment) {
            notes.push({
                level: "warning",
                text: "当前未上传高考报名缴费信息，艺术、体育及其他术科报名人数尚未按缴费情况过滤，统计结果可能包含未缴费考生。"
            });
        } else {
            notes.push({
                level: "info",
                text: "本表人数为缴费交叉过滤后的报名人员数，不等同于“高考报名缴费统计”中普通类的总人数（两者数据来源与统计对象不同）。"
            });
        }
        notes.push({
            level: "info",
            text: "审核通过记录 " + passedRecords.length + " 条，按身份证去重后 " + personMap.size + " 人；最终计入统计 " + included.length + " 人" +
                (unpaidArtsExcluded ? "；因对应缴费项目未缴费而未计入的术科类别记录 " + unpaidArtsExcluded + " 条" : "") + "。"
        });
        if (!artsAvailable) {
            notes.push({level: "warning", text: "源文件缺少 YSTYSK_ 字段，术科各列无法统计。"});
        }
        if (!specialAvailable) {
            notes.push({level: "warning", text: "源文件缺少 CJTZSCS_ 字段，特招生测试人数无法统计。"});
        }

        var baseMeta = {
            source: gaokao.fileName || "高考报名信息",
            auditRule: AUDIT_RULE,
            paymentRule: hasPayment ? PAYMENT_RULE_ON : PAYMENT_RULE_OFF,
            notes: notes
        };

        var tables = [
            Object.assign({
                key: "gaokao-bmd",
                title: "高考报名信息统计（按报名点）",
                columns: buildColumns({withBmd: true, artsAvailable: artsAvailable, specialAvailable: specialAvailable}),
                rows: bmdRows,
                totalRow: totalRowBmd
            }, baseMeta),
            Object.assign({
                key: "gaokao-county",
                title: "高考报名信息统计（按县区汇总）",
                columns: buildColumns({withBmd: false, artsAvailable: artsAvailable, specialAvailable: specialAvailable}),
                rows: countyRows,
                totalRow: totalRowCounty
            }, baseMeta)
        ];

        return {tables: tables, diagnostics: diag.toArray()};
    }

    global.GaokaoStats = {build: build};
})(window);
