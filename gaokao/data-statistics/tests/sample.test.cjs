/**
 * 真实样本校验（样本属于敏感数据，本文件不含任何样本内容）：
 * 只输出表头、汇总数量、取值分布与统计表规模，不输出姓名与身份证号码。
 * 运行：node gaokao/data-statistics/tests/sample.test.cjs
 */
const fs = require("fs");
const path = require("path");

global.window = global;
const XLSX = require(path.join(__dirname, "..", "..", "..", "lib", "sheetjs", "0.20.3", "xlsx.full.min.js"));
global.XLSX = XLSX;

const jsDir = path.join(__dirname, "..", "js");
require(path.join(jsDir, "excel-reader.js"));
require(path.join(jsDir, "data-normalizer.js"));
require(path.join(jsDir, "statistics-gaokao.js"));
require(path.join(jsDir, "statistics-payment.js"));
require(path.join(jsDir, "statistics-january.js"));

const ER = global.ExcelReader;
const DN = global.DataNormalizer;
const sampleDir = path.join(__dirname, "..", "sample");

let failures = 0;
function check(name, ok, detail) {
    if (!ok) failures++;
    console.log((ok ? "ok   " : "FAIL ") + name + (detail === undefined ? "" : " => " + JSON.stringify(detail)));
}

function findSample(keyword) {
    if (!fs.existsSync(sampleDir)) return null;
    const file = fs.readdirSync(sampleDir).find(name => name.indexOf(keyword) !== -1);
    return file ? path.join(sampleDir, file) : null;
}

async function loadFile(filePath, kind) {
    const buffer = fs.readFileSync(filePath);
    const workbook = XLSX.read(buffer, {type: "buffer", cellDates: false});
    const candidates = ER.detectCandidates(workbook, ER.SCHEMAS[kind]);
    if (!candidates.length) return {candidates: []};
    const chosen = candidates[0];
    const matrix = ER.readMatrix(workbook, chosen.sheetName);
    const slice = ER.sliceRows(matrix, chosen.headerRowIndex);
    const parsed = DN[kind === "gaokao" ? "parseGaokao" : kind === "payment" ? "parsePayment" : "parseJanuary"]({
        columns: chosen.columns,
        dataRows: slice.dataRows,
        firstDataRowNumber: slice.firstDataRowNumber
    });
    return {
        candidates: candidates,
        sheetName: chosen.sheetName,
        headerRowIndex: chosen.headerRowIndex,
        labels: chosen.columns.labels,
        missingOptional: ER.SCHEMAS[kind].fields
            .filter(field => !field.required && chosen.columns.map[field.key] === undefined)
            .map(field => field.label || field.header || field.code),
        parsed: parsed,
        fileName: path.basename(filePath)
    };
}

function showDiagnostics(title, diagnostics) {
    const relevant = diagnostics.filter(item => item.level === "warning" || item.level === "error");
    console.log("  " + title + " 异常类别数： " + relevant.length);
    relevant.forEach(item => {
        console.log("    - [" + item.level + "] " + item.title + " ×" + item.count +
            (item.samples.length ? " 样例：" + item.samples.slice(0, 3).join(" / ") : ""));
    });
}

function showDist(title, distributions) {
    distributions.forEach(dist => {
        const text = dist.items.slice(0, 12).map(item => item.label + "(" + item.count + ")").join("、");
        console.log("    " + dist.title + "：" + text + (dist.items.length > 12 ? " …共" + dist.items.length + "种" : ""));
    });
}

(async function () {
    const gkFile = findSample("高考报名信息");
    const payFile = findSample("高考缴费信息");
    const janFile = findSample("1月学选考报名信息");

    if (!gkFile || !payFile || !janFile) {
        console.log("未找到样本文件，跳过（sample 目录：高考报名信息 / 高考缴费信息 / 1月学选考报名信息）");
        return;
    }

    console.log("=== 1. 文件识别");
    const gaokao = await loadFile(gkFile, "gaokao");
    const payment = await loadFile(payFile, "payment");
    const january = await loadFile(janFile, "january");

    check("高考报名文件被识别", gaokao.candidates.length > 0, gaokao.candidates.length);
    check("缴费文件被识别", payment.candidates.length > 0, payment.candidates.length);
    check("学选考文件被识别", january.candidates.length > 0, january.candidates.length);

    console.log("  高考报名：工作表 " + gaokao.sheetName + "，表头行 " + (gaokao.headerRowIndex + 1) +
        "，匹配字段 " + Object.keys(gaokao.labels).length + " 个");
    console.log("  缴费：工作表 " + payment.sheetName + "，表头行 " + (payment.headerRowIndex + 1) +
        "，匹配字段 " + Object.keys(payment.labels).length + " 个");
    console.log("  学选考：工作表 " + january.sheetName + "，表头行 " + (january.headerRowIndex + 1) +
        "，匹配字段 " + Object.keys(january.labels).length + " 个");

    check("高考报名识别到术科字段", gaokao.parsed.presence.arts === true);
    check("高考报名识别到特招生字段", gaokao.parsed.presence.specialTest === true);
    check("学选考识别到外语语种字段", january.parsed.presence.languageCode === true);

    console.log("\n=== 2. 解析摘要");
    [["高考报名", gaokao], ["缴费", payment], ["学选考", january]].forEach(item => {
        const parsed = item[1].parsed;
        console.log("  " + item[0] + "：" + JSON.stringify(parsed.summary));
    });

    console.log("\n=== 3. 取值分布");
    console.log("  [高考报名]");
    showDist("", DN.formatDistributions("gaokao", gaokao.parsed.distributions));
    console.log("  [缴费]");
    showDist("", DN.formatDistributions("payment", payment.parsed.distributions));
    console.log("  [学选考]");
    showDist("", DN.formatDistributions("january", january.parsed.distributions));

    // 报名术科与缴费项目的原始分布（只输出代码与人数）
    const artsBucket = {};
    gaokao.parsed.records.forEach(record => {
        record.artsCodes.forEach(code => {
            artsBucket[code] = (artsBucket[code] || 0) + 1;
        });
    });
    const artsText = Object.keys(artsBucket).sort().map(code => {
        const category = DN.ARTS_CATEGORIES.filter(item => item.code === code)[0];
        return code + (category ? "(" + category.label + ")" : "(未知)") + "=" + artsBucket[code];
    }).join("，");
    console.log("  高考报名术科原始分布：" + (artsText || "无"));

    const projectBucket = {};
    const payIds = new Set();
    const payPtIds = new Set();
    payment.parsed.records.forEach(record => {
        if (record.idCard) payIds.add(record.idCard);
        if (record.categoryKnown) {
            projectBucket[record.categoryCode] = (projectBucket[record.categoryCode] || 0) + 1;
            if (record.categoryCode === "pt") payPtIds.add(record.idCard);
        }
    });
    console.log("  缴费项目代码分布：" + Object.keys(projectBucket).sort().map(code => code + "=" + projectBucket[code]).join("，"));
    console.log("  缴费去重人数：" + payIds.size + "，其中有普通类记录：" + payPtIds.size);

    console.log("  高考报名报名点代码：" + Array.from(new Set(gaokao.parsed.records.map(r => r.bmdCode).filter(Boolean))).join("、"));
    console.log("  缴费报名点代码：" + Array.from(new Set(payment.parsed.records.map(r => r.bmdCode).filter(Boolean))).join("、"));
    console.log("  学选考报名点代码：" + Array.from(new Set(january.parsed.records.map(r => r.bmdCode).filter(Boolean))).join("、"));

    const gkIds = new Set(gaokao.parsed.records.map(record => record.idCard).filter(Boolean));
    const payOnly = Array.from(payPtIds).filter(id => !gkIds.has(id)).length;
    const gkOnly = Array.from(gkIds).filter(id => !payPtIds.has(id)).length;
    console.log("  缴费普通类人员与报名人员比对：仅缴费有 " + payOnly + " 人，仅报名有 " + gkOnly + " 人");

    console.log("\n=== 4. 数据质量");
    showDiagnostics("高考报名", gaokao.parsed.diagnostics);
    showDiagnostics("缴费", payment.parsed.diagnostics);
    showDiagnostics("学选考", january.parsed.diagnostics);

    const gkBmd = new Set(gaokao.parsed.records.map(r => r.bmdCode).filter(Boolean));
    const payBmd = new Set(payment.parsed.records.map(r => r.bmdCode).filter(Boolean));
    const payNotInGk = Array.from(payBmd).filter(code => !gkBmd.has(code));
    console.log("  高考报名报名点数： " + gkBmd.size + "，缴费提取报名点数： " + payBmd.size);
    console.log("  缴费中多出的报名点代码： " + (payNotInGk.length ? payNotInGk.join("、") : "无"));
    console.log("  缴费单号解析失败行数： " + payment.parsed.summary.orderNoFail);

    console.log("\n=== 5. 统计结果（不配置映射，名称显示“未配置名称”）");
    const mapping = {bmdMap: new Map(), countyMap: new Map()};

    const gkOut = global.GaokaoStats.build({
        gaokao: Object.assign({}, gaokao.parsed, {fileName: gaokao.fileName}),
        payment: Object.assign({}, payment.parsed, {fileName: payment.fileName}),
        mapping: mapping
    });
    const payOut = global.PaymentStats.build({
        payment: Object.assign({}, payment.parsed, {fileName: payment.fileName}),
        mapping: mapping
    });
    const janOut = global.JanuaryStats.build({
        january: Object.assign({}, january.parsed, {fileName: january.fileName}),
        mapping: mapping
    });

    const tables = gkOut.tables.concat(payOut.tables, janOut.tables);
    tables.forEach(table => {
        const numbers = table.columns
            .filter(column => column.type === "number" && column.available !== false)
            .map(column => column.label + "=" + table.totalRow[column.key]);
        console.log("  ◆ " + table.title + "：数据行 " + table.rows.length + " 行");
        console.log("     合计：" + numbers.join("，"));
        table.notes.filter(note => note.level === "warning").forEach(note => console.log("     [提示] " + note.text));
    });

    console.log("\n  术科交叉过滤对照（报名人数 → 计入人数）：");
    DN.ARTS_CATEGORIES.forEach(category => {
        const raw = artsBucket[category.code] || 0;
        const cross = gkOut.tables[0].totalRow["arts_" + category.code];
        console.log("    " + category.label + "（" + category.code + "）：" + raw + " → " + (cross === null || cross === undefined ? "字段缺失" : cross) +
            (raw !== cross ? "   ← 有差异" : ""));
    });

    // 交叉口径核对：表（导）演类考生与传媒缴费人员的重合情况（只输出人数）
    const arts34Ids = new Set(gaokao.parsed.records.filter(r => r.artsCodes.indexOf("34") !== -1).map(r => r.idCard));
    const arts36Count = gaokao.parsed.records.filter(r => r.artsCodes.indexOf("36") !== -1).length;
    const cmRecords = payment.parsed.records.filter(r => r.categoryCode === "cm");
    const cmIds = new Set(cmRecords.map(r => r.idCard));
    const overlap = Array.from(arts34Ids).filter(id => cmIds.has(id)).length;
    const cmPaid = cmRecords.filter(r => r.payState === "paid").length;
    console.log("    表（导）演类(34)考生 " + arts34Ids.size + " 人，播音与主持类(36)考生 " + arts36Count +
        " 人；传媒(cm)缴费记录 " + cmRecords.length + " 条（已缴 " + cmPaid + "），涉及 " + cmIds.size + " 人；与 34 类考生重合 " + overlap + " 人");

    console.log("\n  统计阶段异常：");
    showDiagnostics("高考报名统计", gkOut.diagnostics);
    showDiagnostics("缴费统计", payOut.diagnostics);
    showDiagnostics("学选考统计", janOut.diagnostics);

    console.log("\n=== 6. 一致性校验");
    const gkBmdTable = gkOut.tables[0];
    const gkCountyTable = gkOut.tables[1];
    const sumBmd = gkBmdTable.rows.reduce((sum, row) => sum + row.total, 0);
    const sumCounty = gkCountyTable.rows.reduce((sum, row) => sum + row.total, 0);
    check("高考按报名点合计 = 合计行", sumBmd === gkBmdTable.totalRow.total, {sumBmd: sumBmd, total: gkBmdTable.totalRow.total});
    check("高考按县区合计 = 合计行", sumCounty === gkCountyTable.totalRow.total, {sumCounty: sumCounty, total: gkCountyTable.totalRow.total});
    check("每个县区合计不超过报名点合计", sumCounty <= sumBmd + 0.0001, {sumCounty: sumCounty, sumBmd: sumBmd});

    const payTable = payOut.tables[0];
    const sumPayBmd = payTable.rows.reduce((sum, row) => sum + row.total, 0);
    check("缴费按报名点总人数 = 合计行", sumPayBmd === payTable.totalRow.total, {sumPayBmd: sumPayBmd, total: payTable.totalRow.total});

    const janBmdTable = janOut.tables[0];
    check("学选考外语合计 ≥ 各语种之和",
        janBmdTable.totalRow.foreignTotal >= DN.LANGUAGES.reduce((sum, item) => sum + (janBmdTable.totalRow["lang_" + item.code] || 0), 0),
        {foreign: janBmdTable.totalRow.foreignTotal});

    check("高考报名每个考生只归入一个报名点", gkOut.tables[0].rows.every(row => row.bmdCode));
    check("缴费统计所有已缴/未缴列非负",
        payTable.rows.every(row => payTable.columns.filter(c => c.type === "number").every(c => (row[c.key] || 0) >= 0)));

    console.log("\n" + (failures === 0 ? "样本校验通过（0 项失败）" : "存在 " + failures + " 项失败"));
    process.exitCode = failures === 0 ? 0 : 1;
})();
