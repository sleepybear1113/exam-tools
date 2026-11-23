/* 自测：用真实生成的 xlsx / xls 验证表头识别与读取 */
const path = require("path");
global.window = global;
const base = path.join(__dirname, "..", "js");
const XLSX = require(path.join(__dirname, "..", "..", "..", "lib", "sheetjs", "0.20.3", "xlsx.full.min.js"));
global.XLSX = XLSX;
require(path.join(base, "excel-reader.js"));
require(path.join(base, "data-normalizer.js"));

const ER = global.ExcelReader;
const DN = global.DataNormalizer;
let failed = 0;

function check(name, ok, extra) {
    if (!ok) {
        failed++;
        console.log("FAIL", name, extra === undefined ? "" : JSON.stringify(extra));
    } else {
        console.log("ok  ", name);
    }
}

function makeBook(rows, sheetName) {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName);
    return wb;
}

function asFile(buffer) {
    return {
        name: "test.xlsx",
        size: buffer.byteLength,
        arrayBuffer: async () => buffer
    };
}

(async () => {
    /* ---------- 高考报名：表头在第 3 行 + 额外工作表 ---------- */
    const gkRows = [
        ["2026年高考报名信息导出", "", "", "", "", ""],
        ["导出时间：2026-01-01", "", "", "", "", ""],
        ["SHZT_审核状态", "YSTYSK_艺术体育术科", "BMD_报名点", "XM_姓名", "SFZH_身份证号码", "CJTZSCS_参加特招生测试"],
        ["3", "30|31", "80101", "张三", "330402199001011234", "1"],
        ["3", "", "80102", "李四", "330402199001011235", "0"],
        ["审核通过", "41", "80101", "王五", "330402199001011236", ""]
    ];
    let wb = makeBook(gkRows, "报名信息");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["说明"], ["本表为导出数据"]]), "说明");

    let buffer = XLSX.write(wb, {type: "array", bookType: "xlsx"});
    let result = await ER.readWorkbook(asFile(buffer));
    let candidates = ER.detectCandidates(result.workbook, ER.SCHEMAS.gaokao);
    check("xlsx 候选识别数量", candidates.length === 1, candidates.map(c => c.sheetName));
    check("xlsx 候选工作表", candidates[0].sheetName === "报名信息", candidates[0] && candidates[0].sheetName);
    check("xlsx 表头行索引", candidates[0].headerRowIndex === 2, candidates[0] && candidates[0].headerRowIndex);

    const matrix = ER.readMatrix(result.workbook, candidates[0].sheetName);
    const sliced = ER.sliceRows(matrix, candidates[0].headerRowIndex);
    const parsed = DN.parseGaokao({
        columns: candidates[0].columns,
        dataRows: sliced.dataRows,
        firstDataRowNumber: sliced.firstDataRowNumber
    });
    check("xlsx 有效行数", parsed.summary.totalRows === 3, parsed.summary);
    check("xlsx 审核通过数", parsed.summary.auditPassed === 3, parsed.summary);
    check("xlsx 首人术科", JSON.stringify(parsed.records[0].artsCodes) === JSON.stringify(["30", "31"]));
    check("xlsx 首人特招生", parsed.records[0].specialHit === true);
    check("xlsx 第三人审核通过文本", parsed.records[2].auditStatus === "审核通过");
    check("xlsx 数值型身份证识别", parsed.records.every(r => !r.idNumeric || r.idCorrupted));

    /* ---------- 缴费 .xls（BIFF8） ---------- */
    const payRows = [
        ["缴款单号", "姓名", "证件号码", "缴款状态", "缴款金额", "项目编号"],
        ["X".repeat(17) + "0101" + "12" + "ZZ", "张三", "330402199001011234", "已缴款", 100, "A_pt"]
    ];
    const payWb = makeBook(payRows, "缴费");
    let payBuffer;
    try {
        payBuffer = XLSX.write(payWb, {type: "array", bookType: "biff8"});
        check("biff8 写入成功", payBuffer && payBuffer.byteLength > 0);
    } catch (err) {
        payBuffer = XLSX.write(payWb, {type: "array", bookType: "xlsx"});
        console.log("skip  当前 SheetJS 构建不支持写 biff8，改用 xlsx 验证读取流程");
    }

    result = await ER.readWorkbook(asFile(payBuffer));
    candidates = ER.detectCandidates(result.workbook, ER.SCHEMAS.payment);
    check("缴费表头识别", candidates.length === 1, candidates.map(c => c.sheetName));
    const payMatrix = ER.readMatrix(result.workbook, candidates[0].sheetName);
    const paySlice = ER.sliceRows(payMatrix, candidates[0].headerRowIndex);
    const payment = DN.parsePayment({
        columns: candidates[0].columns,
        dataRows: paySlice.dataRows,
        firstDataRowNumber: paySlice.firstDataRowNumber
    });
    check("缴费 报名点提取", payment.records[0].bmdCode === "80101", payment.records[0].bmdCode);
    check("缴费 类别识别", payment.records[0].categoryCode === "pt" && payment.records[0].categoryKnown);
    check("缴费 状态识别", payment.records[0].payState === "paid");

    /* ---------- 1月学选考：表头在第 2 行 ---------- */
    const janRows = [
        ["1月学选考报名数据"],
        ["SFZH", "XM", "BMD", "SHZT", "WYYZ", "KM_YW", "KM_SX", "KM_YY", "KM_SZ", "KM_LS", "KM_DL", "KM_WL", "KM_HX", "KM_SW", "KM_JS"],
        ["330402199001011234", "张三", "80101", 3, 1, "学考", "学考", "高考", "学考", "选考", "", "", "", "", ""]
    ];
    result = await ER.readWorkbook(asFile(XLSX.write(makeBook(janRows, "学选考"), {type: "array", bookType: "xlsx"})));
    candidates = ER.detectCandidates(result.workbook, ER.SCHEMAS.january);
    check("1月表头识别", candidates.length === 1 && candidates[0].headerRowIndex === 1, candidates);
    const janMatrix = ER.readMatrix(result.workbook, candidates[0].sheetName);
    const janSlice = ER.sliceRows(janMatrix, candidates[0].headerRowIndex);
    const january = DN.parseJanuary({
        columns: candidates[0].columns,
        dataRows: janSlice.dataRows,
        firstDataRowNumber: janSlice.firstDataRowNumber
    });
    check("1月 审核通过（数值3）", january.summary.auditPassed === 1, january.summary);
    check("1月 语种", january.records[0].languageCode === "1", january.records[0].languageCode);

    /* ---------- 单独考试报名：表头在第 1 行 ---------- */
    const ddRows = [
        ["BMD_报名点", "BKKL_报考类别", "QTLZY_其他类专业", "SFZH_身份证号码", "SHZT_审核状态", "XM_姓名"],
        ["80101", "17", "2", "330402199001011234", "3", "张三"],
        ["80101", "1", "", "330402199001011235", "审核通过", "李四"]
    ];
    const ddResult = await ER.readWorkbook(asFile(XLSX.write(makeBook(ddRows, "单独考试"), {type: "array", bookType: "xlsx"})));
    const ddCandidates = ER.detectCandidates(ddResult.workbook, ER.SCHEMAS.dandu);
    check("单独表头识别", ddCandidates.length === 1 && ddCandidates[0].headerRowIndex === 0, ddCandidates.map(c => c.sheetName));
    const ddMatrix = ER.readMatrix(ddResult.workbook, ddCandidates[0].sheetName);
    const ddSlice = ER.sliceRows(ddMatrix, ddCandidates[0].headerRowIndex);
    const dandu = DN.parseDandu({
        columns: ddCandidates[0].columns,
        dataRows: ddSlice.dataRows,
        firstDataRowNumber: ddSlice.firstDataRowNumber
    });
    check("单独 有效行数", dandu.summary.totalRows === 2, dandu.summary);
    check("单独 审核通过数", dandu.summary.auditPassed === 2, dandu.summary);
    check("单独 报考类别细分", dandu.records[0].categoryKey === "17-2", dandu.records[0].categoryKey);
    check("单独 其余字段齐全", dandu.presence.idCard && dandu.presence.name && dandu.presence.auditStatus);
    check("单独文件不被高考区域识别", ER.detectCandidates(ddResult.workbook, ER.SCHEMAS.gaokao).length === 0);

    /* ---------- 放错区域的表头不应被识别 ---------- */
    candidates = ER.detectCandidates(result.workbook, ER.SCHEMAS.gaokao);
    check("错放区域不识别", candidates.length === 0, candidates.length);

    console.log(failed === 0 ? "\n全部通过" : "\n失败 " + failed + " 项");
    process.exit(failed === 0 ? 0 : 1);
})();
