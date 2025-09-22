/* global XLSX */

/**
 * 读取Excel文件，返回数据数组<br>
 * 用法： const dataArray = await readExcel(file);
 * @param {File} file - 用户选择的文件对象
 * @returns {Promise<Array>} - Excel解析后的数组
 */
export function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = f => {
            try {
                const data = new Uint8Array(f.target.result);
                const workbook = XLSX.read(data, {type: 'array'});

                // 默认取第一个工作表
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const result = XLSX.utils.sheet_to_json(firstSheet, {header: 1});

                resolve(result);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}