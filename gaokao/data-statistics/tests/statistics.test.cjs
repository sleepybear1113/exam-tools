/* 自测：用构造数据验证统计口径与边界规则 */
const path = require("path");
global.window = global;
const base = path.join(__dirname, "..", "js");
require(path.join(base, "data-normalizer.js"));
require(path.join(base, "statistics-gaokao.js"));
require(path.join(base, "statistics-payment.js"));
require(path.join(base, "statistics-january.js"));
require(path.join(base, "statistics-dandu.js"));

const DN = global.DataNormalizer;
let failed = 0;

function eq(name, actual, expected) {
    const ok = JSON.stringify(actual) === JSON.stringify(expected);
    if (!ok) {
        failed++;
        console.log("FAIL", name, "=>", JSON.stringify(actual), "expected", JSON.stringify(expected));
    } else {
        console.log("ok  ", name, "=", JSON.stringify(actual));
    }
}

/* ---------- 基础函数 ---------- */
eq("bmd 不足5位补零", DN.normalizeBmdCode("101").code, "00101");
eq("bmd 超长保留", DN.normalizeBmdCode("123456"), {code: "123456", issue: "too-long"});
eq("县区代码", DN.countyCodeOf("80101"), "01");

function orderNo(suffix, indicator) {
    return "X".repeat(17) + suffix + indicator + "ZZ";
}
eq("单号 12→8", DN.extractBmdFromOrderNo(orderNo("0101", "12")).code, "80101");
eq("单号 13→9", DN.extractBmdFromOrderNo(orderNo("0101", "13")).code, "90101");
eq("单号 50→9", DN.extractBmdFromOrderNo(orderNo("0101", "50")).code, "90101");
eq("单号 51→7", DN.extractBmdFromOrderNo(orderNo("0101", "51")).code, "70101");
eq("单号 太短", DN.extractBmdFromOrderNo("123").issue, "too-short");
eq("术科拆分", DN.parseArtsCodes("30|31|41"), ["30", "31", "41"]);
eq("特招生 文本1", DN.isSpecialTestHit(" 1 "), true);
eq("特招生 数字1", DN.isSpecialTestHit(1), true);
eq("特招生 Y", DN.isSpecialTestHit("Y"), false);
eq("待缴款=未缴", DN.parsePayStatus("待缴款"), "unpaid");
eq("项目代码大小写", DN.parseProjectCategory("ABC_MS").code, "ms");
eq("身份证 x 转大写", DN.normalizeIdCard(" 33040219900101123x "), "33040219900101123X");
eq("身份证 数值精度", DN.normalizeIdCard(330402199001011234), "");
eq("身份证 数值精度标记", DN.isCorruptedId(330402199001011234), true);

/* ---------- 构造高考报名数据 ---------- */
const gkColumns = {
    map: {auditStatus: 0, className: 1, arts: 2, bmdCode: 3, name: 4, idCard: 5, specialTest: 6},
    labels: {}
};
const gkRows = [
    ["3", "1班", "30|31", "80101", "张三", "330402199001011234", "1"],
    ["3", "1班", "31", "80101", "李四", "330402199001011235", ""],
    ["3", "1班", "41", "80201", "王五", "330402199001011236", "1"],
    ["1", "1班", "31", "80101", "赵六", "330402199001011237", ""],
    ["3", "1班", "39", "80101", "张三", "330402199001011234", "1"],
    ["3", "1班", "32", "80301", "孙七", "330402199001011238", ""]
];
const gaokao = DN.parseGaokao({columns: gkColumns, dataRows: gkRows, firstDataRowNumber: 2});
eq("高考 有效行", gaokao.summary.totalRows, 6);
eq("高考 审核通过", gaokao.summary.auditPassed, 5);
eq("高考 重复身份证", gaokao.summary.duplicateIds, 1);
eq("高考 特招生分布", gaokao.distributions.special, {"1": 3});

/* ---------- 构造缴费数据 ---------- */
const payColumns = {
    map: {orderNo: 0, name: 1, idCard: 2, payStatus: 3, amount: 4, projectNo: 5},
    labels: {}
};
const payRows = [
    [orderNo("0101", "12"), "张三", "330402199001011234", "已缴款", "100", "A_pt"],
    [orderNo("0101", "12"), "张三", "330402199001011234", "未缴款", "50", "B_ms"],
    [orderNo("0201", "12"), "王五", "330402199001011236", "待缴款", "100", "C_pt"],
    [orderNo("0301", "12"), "孙七", "330402199001011238", "已缴款", "50", "D_yy"]
];
const payment = DN.parsePayment({columns: payColumns, dataRows: payRows, firstDataRowNumber: 2});
eq("缴费 报名点数", payment.summary.bmdCount, 3);
eq("缴费 单号失败数", payment.summary.orderNoFail, 0);

/* ---------- 交叉统计 ---------- */
const mapping = {
    bmdMap: new Map([["80101", "一中"], ["80201", "二中"], ["80301", "三中"]]),
    countyMap: new Map([["01", "南湖区"], ["02", "秀洲区"], ["03", "嘉善县"]])
};
const gkOut = global.GaokaoStats.build({gaokao, payment, mapping});
const bmdTable = gkOut.tables[0];
const row80101 = bmdTable.rows.find(r => r.bmdCode === "80101");
eq("交叉 80101 总人数", row80101.total, 2);
eq("交叉 80101 美术类(ms未缴→0)", row80101.arts_31, 1);
eq("交叉 80101 书法类(sf无记录→0)", row80101.arts_30, 0);
eq("交叉 80101 戏曲类(39无需缴费)", row80101.arts_39, 1);
eq("交叉 80101 特招生", row80101.special, 1);
eq("交叉 80201 总人数(pt未缴→排除)", bmdTable.rows.find(r => r.bmdCode === "80201"), undefined);
eq("交叉 80301 总人数(无pt记录)", (bmdTable.rows.find(r => r.bmdCode === "80301") || {}).total, undefined);
eq("交叉 合计人数", bmdTable.totalRow.total, 2);
eq("县区名称", bmdTable.rows.find(r => r.bmdCode === "80101").countyName, "南湖区");

/* ---------- 独立缴费统计 ---------- */
const payOut = global.PaymentStats.build({payment, mapping});
const payTable = payOut.tables[0];
const pay80101 = payTable.rows.find(r => r.bmdCode === "80101");
eq("缴费表 总人数", pay80101.total, 1);
eq("缴费表 pt已缴", pay80101.paid_pt, 1);
eq("缴费表 ms未缴", pay80101.unpaid_ms, 1);
eq("缴费表 已缴费列在前", payTable.columns[5].key, "paid_pt");
eq("缴费表 未缴列在后", payTable.columns[13].key, "unpaid_pt");
eq("缴费表 合计数", payTable.totalRow.total, 2);

/* ---------- 1月学选考 ---------- */
const janColumns = {
    map: {idCard: 0, name: 1, bmdCode: 2, auditStatus: 3, languageCode: 4, yw: 5, sx: 6, yy: 7, sz: 8, ls: 9, dl: 10, wl: 11, hx: 12, sw: 13, js: 14},
    labels: {}
};
const janRows = [
    ["330402199001011234", "张三", "80101", "3", "1", "学考", "学考", "高考", "学考", "选考", "", "", "", "", ""],
    ["330402199001011235", "李四", "80101", "3", "9", "学考", "", "高考", "学考", "", "", "", "", "", ""],
    ["330402199001011236", "王五", "80201", "3", "", "", "", "", "", "", "", "选考", "", "", ""],
    ["330402199001011234", "张三", "80101", "3", "1", "", "", "高考", "", "", "", "", "", "", ""]
];
const january = DN.parseJanuary({columns: janColumns, dataRows: janRows, firstDataRowNumber: 2});
const janOut = global.JanuaryStats.build({january, mapping});
const janTable = janOut.tables[0];
const j80101 = janTable.rows.find(r => r.bmdCode === "80101");
eq("1月 审核通过", january.summary.auditPassed, 4);
eq("1月 重复身份证", january.summary.duplicateIds, 1);
eq("1月 语文", j80101.yw, 2);
eq("1月 数学", j80101.sx, 1);
eq("1月 外语合计", j80101.foreignTotal, 2);
eq("1月 英语", j80101.lang_1, 1);
eq("1月 未知语种", j80101.lang_unknown, 1);
eq("1月 政治学考", j80101.xk_sz, 2);
eq("1月 历史选考", j80101.xx_ls, 1);
const janCounty = janOut.tables[1];
eq("1月 县区数", janCounty.rows.length, 2);

/* ---------- 单独考试 ---------- */
const ddColumns = {
    map: {bmdCode: 0, category: 1, otherMajor: 2, idCard: 3, auditStatus: 4, name: 5},
    labels: {}
};
const ddRows = [
    ["80101", "17", "2", "330402199001011234", "3", "张三"],
    ["80101", "17", "1", "330402199001011235", "3", "李四"],
    ["80101", "1", "", "330402199001011236", "3", "王五"],
    ["80201", "1", "", "330402199001011234", "3", "张三"],
    ["80101", "1", "", "330402199001011237", "1", "赵六"]
];
const dandu = DN.parseDandu({columns: ddColumns, dataRows: ddRows, firstDataRowNumber: 2});
eq("单独 有效行", dandu.summary.totalRows, 5);
eq("单独 报名点数", dandu.summary.bmdCount, 2);
eq("单独 报考类别数", dandu.summary.categoryCount, 2);
eq("单独 审核通过", dandu.summary.auditPassed, 4);
eq("单独 重复身份证", dandu.summary.duplicateIds, 1);
eq("单独 presence 审核字段", dandu.presence.auditStatus, true);
eq("单独 17-2 类别键", dandu.records[0].categoryKey, "17-2");
eq("单独 17-1 类别键", dandu.records[1].categoryKey, "17-1");
eq("单独 类别代码补零", dandu.records[2].categoryCode, "01");
eq("单独 类别标签", DN.danduCategoryLabel("17-2"), "17-2 体育");
eq("单独 类别标签 01", DN.danduCategoryLabel("01"), "01 机械类");
eq("单独 类别代码归一", DN.normalizeDanduCategory(1), "01");

const ddOut = global.DanduStats.build({dandu, mapping});
const ddBmdTable = ddOut.tables[0];
const dd80101 = ddBmdTable.rows.find(r => r.bmdCode === "80101");
eq("单独表 报名点数", ddBmdTable.rows.length, 1);
eq("单独表 80101 总人数", dd80101.total, 2);
eq("单独表 80101 类别01", dd80101.cat_01, 1);
eq("单独表 类别列标题", ddBmdTable.columns.find(c => c.key === "cat_01").label, "01 机械类");
eq("单独表 80101 类别17-1", dd80101["cat_17-1"], 1);
eq("单独表 80101 类别17-2", dd80101["cat_17-2"], 0);
eq("单独表 合计人数", ddBmdTable.totalRow.total, 2);
eq("单独表 跨报名点冲突排除", ddBmdTable.rows.find(r => r.bmdCode === "80201"), undefined);
eq("单独表 跨报名点冲突提示", ddOut.diagnostics.some(d => d.code === "DANDU_BMD_CONFLICT"), true);
eq("单独 县区表行数", ddOut.tables[1].rows.length, 1);

/* ---------- 缴费单号无法解析 ---------- */
const badPay = DN.parsePayment({
    columns: payColumns,
    dataRows: [["123", "x", "330402199001011234", "已缴款", "", "A_pt"]],
    firstDataRowNumber: 2
});
eq("缴费 单号失败统计", badPay.summary.orderNoFail, 1);

console.log(failed === 0 ? "\n全部通过" : "\n失败 " + failed + " 项");
process.exit(failed === 0 ? 0 : 1);
