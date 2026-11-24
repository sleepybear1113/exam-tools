/**
 * 手机收缴登记表 · 数据处理模块
 *
 * 职责：
 *   1. 读取「高考报名库」，取 BMXH_（报名序号）、BJ_（班级）、BMD_（报名点）；
 *   2. 以美术报名库为主表，按 美术库 KSH ↔ 高考库 BMXH_ 一对一匹配班级与报名点；
 *   3. 报名点名称取自「高考及学选考报名数据统计工具」保存在 localStorage 的映射表；
 *   4. 以 art/20xx年浙江美术统考手机收缴登记-模板.xlsx 为底，填数据生成 4 个工作表的 Excel。
 *
 * 依赖全局对象：XLSX（读表）、ExcelJS（写模板）、JSZip（打包）。
 * 所有解析与生成均在浏览器本地完成，不上传任何数据。
 */
(function (global) {
    'use strict';

    /* 与模板中的 4 个工作表一一对应 */
    var SUBJECTS = ['熟悉试场', '素描', '速写', '色彩'];

    /* 与「数据统计工具」共用同一份报名点映射（仅代码与名称，不含考生信息） */
    var BMD_STORAGE_KEY = 'exam-data-statistics:bmd-map:v1';

    var HEADER_SCAN = 20;
    var DATA_START_ROW = 4;   // 模板第 4 行起为数据行

    /* 高考报名库需要的列：key 为内部字段名，aliases 为可识别的表头写法 */
    var GK_FIELDS = [
        {key: 'bmxh', label: 'BMXH_（报名序号）', aliases: ['BMXH', '报名序号'], required: true},
        {key: 'bj', label: 'BJ_（班级）', aliases: ['BJ', '班级'], required: true},
        {key: 'bmd', label: 'BMD_（报名点）', aliases: ['BMD', '报名点', '报名点代码'], required: true}
    ];

    /* ------------------------------------------------------------------ */
    /* 通用文本处理                                                        */
    /* ------------------------------------------------------------------ */

    function normalizeHeader(value) {
        return String(value === null || value === undefined ? '' : value)
            .replace(/[\s\u3000\uFEFF]/g, '')
            .replace(/_+$/, '')
            .toUpperCase();
    }

    function cellText(value) {
        if (value === null || value === undefined) return '';
        if (typeof value === 'number') return String(value);
        return String(value).replace(/^[\s\u3000\uFEFF']+|[\s\u3000\uFEFF]+$/g, '');
    }

    /** 报名点代码：纯数字不足 5 位左补零（与数据统计工具口径一致） */
    function normalizeBmdCode(value) {
        var text = cellText(value).replace(/[\s\u3000\uFEFF]/g, '');
        if (!text) return '';
        if (/^\d+$/.test(text)) {
            return text.length < 5 ? ('00000' + text).slice(-5) : text;
        }
        return text;
    }

    function matchHeader(row) {
        var normalized = row.map(normalizeHeader);
        var map = {};
        var missing = [];

        GK_FIELDS.forEach(function (field) {
            var aliases = field.aliases.map(normalizeHeader);
            var hit = -1;
            for (var i = 0; i < normalized.length; i++) {
                if (normalized[i] && aliases.indexOf(normalized[i]) >= 0) {
                    hit = i;
                    break;
                }
            }
            if (hit >= 0) {
                map[field.key] = hit;
            } else if (field.required) {
                missing.push(field.label);
            }
        });

        return {map: map, missing: missing};
    }

    /* ------------------------------------------------------------------ */
    /* 读取高考报名库                                                      */
    /* ------------------------------------------------------------------ */

    function parseGaokaoBuffer(buffer) {
        var xlsx = global.XLSX;
        if (!xlsx) throw new Error('未加载 SheetJS 解析库，请检查 lib/sheetjs 路径是否正确。');

        var workbook = xlsx.read(new Uint8Array(buffer), {type: 'array'});
        var sheetNames = workbook.SheetNames || [];
        var fallbackHeaders = [];

        for (var s = 0; s < sheetNames.length; s++) {
            var sheet = workbook.Sheets[sheetNames[s]];
            if (!sheet) continue;

            var aoa = xlsx.utils.sheet_to_json(sheet, {
                header: 1,
                raw: true,
                defval: '',
                blankrows: false
            });

            var head = null;
            for (var i = 0; i < aoa.length && i < HEADER_SCAN; i++) {
                var row = aoa[i] || [];
                if (!fallbackHeaders.length) {
                    fallbackHeaders = row.map(cellText).filter(Boolean).slice(0, 12);
                }
                var result = matchHeader(row);
                if (result.missing.length === 0) {
                    head = {index: i, map: result.map};
                    break;
                }
            }
            if (!head) continue;

            var records = [];
            for (var r = head.index + 1; r < aoa.length; r++) {
                var line = aoa[r] || [];
                var bmxh = cellText(line[head.map.bmxh]);
                var bj = cellText(line[head.map.bj]);
                var bmd = normalizeBmdCode(line[head.map.bmd]);

                if (!bmxh && !bj && !bmd) continue;
                if (!bmxh) continue;

                records.push({bmxh: bmxh, bj: bj, bmd: bmd});
            }

            if (!records.length) {
                throw new Error('工作表「' + sheetNames[s] + '」中没有读取到有效的报名数据行。');
            }

            return {sheetName: sheetNames[s], records: records};
        }

        throw new Error('未找到所需的列：BMXH_（报名序号）、BJ_（班级）、BMD_（报名点）。'
            + (fallbackHeaders.length ? '当前文件首个表头行为：' + fallbackHeaders.join('、') : ''));
    }

    /* ------------------------------------------------------------------ */
    /* 报名点映射（与数据统计工具共用 localStorage）                        */
    /* ------------------------------------------------------------------ */

    function readBmdMap() {
        var map = {};
        var rows = [];
        try {
            var raw = global.localStorage.getItem(BMD_STORAGE_KEY);
            if (raw) {
                var parsed = JSON.parse(raw);
                if (Array.isArray(parsed)) rows = parsed;
            }
        } catch (err) {
            return {map: map, count: 0, available: false};
        }

        rows.forEach(function (item) {
            if (!item) return;
            var code = normalizeBmdCode(item.code);
            var name = cellText(item.name);
            if (!code || !name) return;
            if (!map[code]) map[code] = name;
        });

        return {map: map, count: Object.keys(map).length, available: true};
    }

    /* ------------------------------------------------------------------ */
    /* 主表关联                                                            */
    /* ------------------------------------------------------------------ */

    /** 登记表行序：先按班级（自然排序，如 高三(2)班 在 高三(10)班 之前），再按考生号 */
    function compareRow(a, b) {
        var byClass = String(a.bj || '').localeCompare(String(b.bj || ''), 'zh-CN', {numeric: true});
        if (byClass !== 0) return byClass;

        var na = Number(a.ksh);
        var nb = Number(b.ksh);
        if (isFinite(na) && isFinite(nb) && na !== nb) return na - nb;
        return String(a.ksh).localeCompare(String(b.ksh), 'zh-CN', {numeric: true});
    }

    function compareBmd(a, b) {
        var na = parseInt(a, 10);
        var nb = parseInt(b, 10);
        if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
        return String(a).localeCompare(String(b), 'zh-CN', {numeric: true});
    }

    /**
     * 以美术报名库为主表，按 KSH ↔ BMXH_ 关联高考报名库，再按报名点分组。
     * @param {Array} artRecords 美术库记录（含 zkzh / ksh / xm / kc）
     * @param {Array} gkRecords  高考库记录（含 bmxh / bj / bmd）
     * @param {Object} bmdMap    报名点代码 → 名称
     */
    function groupByBmd(artRecords, gkRecords, bmdMap) {
        var lookup = {};
        var duplicateMbxh = 0;

        (gkRecords || []).forEach(function (item) {
            if (lookup[item.bmxh]) {
                duplicateMbxh++;
                return;
            }
            lookup[item.bmxh] = item;
        });

        var groups = {};
        var order = [];
        var unmatched = [];
        var noBmd = 0;

        (artRecords || []).forEach(function (art) {
            var gk = lookup[art.ksh];
            if (!gk) {
                unmatched.push(art);
                return;
            }
            if (!gk.bmd) {
                noBmd++;
                return;
            }
            var group = groups[gk.bmd];
            if (!group) {
                group = {
                    code: gk.bmd,
                    name: (bmdMap && bmdMap[gk.bmd]) || '',
                    rows: []
                };
                groups[gk.bmd] = group;
                order.push(gk.bmd);
            }
            group.rows.push({
                zkzh: art.zkzh,
                ksh: art.ksh,
                xm: art.xm,
                kc: art.kc,
                seat: art.seat,
                bj: gk.bj
            });
        });

        order.sort(compareBmd);
        var list = order.map(function (code) {
            var group = groups[code];
            group.rows.sort(compareRow);
            return group;
        });

        var warnings = [];
        if (duplicateMbxh) {
            warnings.push('高考报名库中有 ' + duplicateMbxh + ' 条重复的报名序号（BMXH_），已按首条记录匹配。');
        }
        if (unmatched.length) {
            warnings.push('有 ' + unmatched.length + ' 名考生在高考报名库中找不到对应的报名序号（KSH ↔ BMXH_），未纳入手机收缴登记表。');
        }
        if (noBmd) {
            warnings.push('有 ' + noBmd + ' 名考生的报名点（BMD_）为空，未纳入手机收缴登记表。');
        }
        return {
            groups: list,
            unmatched: unmatched,
            warnings: warnings,
            matched: (artRecords || []).length - unmatched.length - noBmd
        };
    }

    /* ------------------------------------------------------------------ */
    /* 生成 Excel                                                          */
    /* ------------------------------------------------------------------ */

    function cloneStyle(style) {
        return style ? JSON.parse(JSON.stringify(style)) : {};
    }

    function safeFileName(text) {
        return String(text === null || text === undefined ? '' : text)
            .replace(/[\\\/:*?"<>|\r\n\t]/g, '_')
            .trim();
    }

    /**
     * 每个报名点一个文件，命名规则：前缀-报名点代码-报名点名称
     * 例：2027年浙江美术统考手机收缴登记-33040201-嘉兴市第五高级中学.xlsx
     */
    function buildFileName(year, group) {
        var head = year + '年浙江美术统考手机收缴登记-'
            + (group.code || '') + '-' + (group.name || '未配置名称');
        return safeFileName(head) + '.xlsx';
    }

    /** 表内报名点写法：名称（代码）；未配置名称时只写代码 */
    function buildBmdLabel(group) {
        if (!group.name) return group.code || '';
        return group.name + '（' + group.code + '）';
    }

    /** 把一行数据写入某个工作表的数据区（模板第 4 行起） */
    function fillSheet(ws, options) {
        var rows = options.rows || [];
        var year = String(options.year);
        var bmdLabel = options.bmdLabel || '';

        // 标题：YEAR → 年份
        var titleCell = ws.getCell('A1');
        var title = typeof titleCell.value === 'string' ? titleCell.value : (titleCell.text || '');
        if (title) titleCell.value = title.replace(/YEAR/g, year);

        // 副标题：BMD → 报名点名称（未配置名称时用代码）
        var subCell = ws.getCell('A2');
        var sub = typeof subCell.value === 'string' ? subCell.value : (subCell.text || '');
        if (sub) subCell.value = sub.replace(/BMD/g, bmdLabel);

        // 先取出模板第 4 行的样式与行高，作为后续数据行的样式来源
        var templateStyles = [];
        for (var c = 1; c <= 6; c++) {
            templateStyles.push(cloneStyle(ws.getCell(DATA_START_ROW, c).style));
        }
        var templateHeight = ws.getRow(DATA_START_ROW).height;

        for (var i = 0; i < rows.length; i++) {
            var rowNo = DATA_START_ROW + i;
            var row = ws.getRow(rowNo);
            row.height = templateHeight;

            var item = rows[i];
            var values = [
                i + 1,          // 序号
                item.zkzh,      // 准考证号
                item.ksh,       // 考生号
                item.xm,        // 姓名
                item.bj,        // 班级
                ''              // 收缴情况（留空手写/勾选）
            ];

            for (var col = 1; col <= 6; col++) {
                var cell = ws.getCell(rowNo, col);
                cell.style = cloneStyle(templateStyles[col - 1]);
                cell.value = values[col - 1];
                // 准考证号、考生号为长数字串，强制文本格式，避免被科学计数
                if (col === 2 || col === 3) cell.numFmt = '@';
            }
        }

        // 清掉模板数据区以外的残留（模板本身只有 3 行内容，这里仅作保护）
        if (rows.length === 0) {
            ws.getRow(DATA_START_ROW).height = templateHeight;
        }
    }

    /**
     * 以模板为底生成一个报名点的 Excel（4 个科目工作表）。
     * @returns {Promise<ArrayBuffer|Buffer>}
     */
    function buildWorkbook(templateBuffer, options) {
        var ExcelJS = global.ExcelJS;
        if (!ExcelJS) return Promise.reject(new Error('未加载 ExcelJS 库，请检查 lib/exceljs 路径是否正确。'));

        var workbook = new ExcelJS.Workbook();
        return workbook.xlsx.load(templateBuffer).then(function () {
            var label = buildBmdLabel(options.group);
            workbook.worksheets.forEach(function (ws) {
                fillSheet(ws, {
                    year: options.year,
                    bmdLabel: label,
                    rows: options.group.rows
                });
            });
            return workbook.xlsx.writeBuffer();
        });
    }

    /** 打包 ZIP：每个报名点一个 xlsx */
    function buildZip(files, type) {
        var JSZip = global.JSZip;
        if (!JSZip) return Promise.reject(new Error('未加载 JSZip 库，请检查 lib/jszip 路径是否正确。'));

        var zip = new JSZip();
        files.forEach(function (file) {
            zip.file(file.name, file.buffer);
        });
        return zip.generateAsync({
            type: type || 'blob',
            compression: 'DEFLATE',
            compressionOptions: {level: 6}
        });
    }

    global.PhoneRegister = {
        SUBJECTS: SUBJECTS,
        BMD_STORAGE_KEY: BMD_STORAGE_KEY,
        DATA_START_ROW: DATA_START_ROW,
        parseGaokaoBuffer: parseGaokaoBuffer,
        normalizeBmdCode: normalizeBmdCode,
        readBmdMap: readBmdMap,
        groupByBmd: groupByBmd,
        buildFileName: buildFileName,
        buildBmdLabel: buildBmdLabel,
        buildWorkbook: buildWorkbook,
        buildZip: buildZip
    };

})(typeof window !== 'undefined' ? window : globalThis);
