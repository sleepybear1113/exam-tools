class Field {
    constructor(name, type, length) {
        this.name = name;
        this.type = type;
        this.length = length;
    }
}

function parseDBF(arrayBuffer, encoding) {
    const view = new DataView(arrayBuffer);
    if (view.byteLength < 32) {
        throw new Error('DBF 文件头不完整');
    }

    // 解析文件头
    const version = view.getUint8(0);
    const year = view.getUint8(1) + 1900;
    const month = view.getUint8(2);
    const day = view.getUint8(3);
    const recordCount = view.getUint32(4, true);  // Little-endian
    const headerLength = view.getUint16(8, true);
    const recordLength = view.getUint16(10, true);
    if (headerLength < 33 || headerLength > view.byteLength || recordLength < 1) {
        throw new Error('DBF 文件头长度无效');
    }

    // 解析字段信息
    const fields = [];
    let offset = 32; // 头部后的第一个字段描述

    // 读取字段描述，直到遇到终止符 0x0D
    while (offset < headerLength && view.getUint8(offset) !== 0x0D) {
        if (offset + 32 > headerLength) {
            throw new Error('DBF 字段定义不完整');
        }
        // 读取字段名（前11字节）
        let fieldName = '';
        for (let j = 0; j < 11; j++) {
            const byte = view.getUint8(offset + j);
            if (byte !== 0) { // 跳过空字节
                fieldName += String.fromCharCode(byte);
            }
        }

        const fieldType = String.fromCharCode(view.getUint8(offset + 11));
        const fieldLength = view.getUint8(offset + 16);

        fields.push(new Field(fieldName.trim(), fieldType, fieldLength));

        offset += 32; // 每个字段描述32字节
    }
    if (offset >= headerLength || view.getUint8(offset) !== 0x0D) {
        throw new Error('DBF 缺少字段结束标记');
    }
    if (fields.reduce((total, field) => total + field.length, 0) + 1 > recordLength) {
        throw new Error('DBF 记录长度无效');
    }

    // 获取适当的TextDecoder
    const decoder = getDecoder(encoding);

    // 解析记录数据
    const records = [];
    offset = headerLength;  // 记录起始位置

    for (let i = 0; i < recordCount; i++) {
        if (offset + recordLength > view.byteLength) {
            throw new Error('DBF 记录数据不完整');
        }
        // 检查删除标记（0x2A表示删除）
        if (view.getUint8(offset) === 0x2A) {
            offset += recordLength;
            continue;
        }

        // record 是一个对象，存储每个字段的值，每行数据对应一个对象
        const record = {};
        let fieldOffset = offset + 1;  // 跳过删除标志

        for (const field of fields) {
            const bytes = new Uint8Array(arrayBuffer, fieldOffset, field.length);
            let value = '';

            try {
                value = decoder.decode(bytes);
            } catch (e) {
                value = Array.from(bytes).map(b => String.fromCharCode(b)).join('');
            }

            record[field.name] = value.trim().replace(/\0/g, '');
            fieldOffset += field.length;
        }

        records.push(record);
        offset += recordLength;
    }

    return {fields, records};
}

function getDecoder(encoding) {
    try {
        // 对于中文编码使用gb18030（大多数现代浏览器支持）
        if (encoding === 'GBK' || encoding === 'GB2312') {
            return new TextDecoder('gb18030');
        } else {
            return new TextDecoder(encoding);
        }
    } catch (e) {
        console.warn(`不支持${encoding}编码，将使用ASCII解码`);
        // 降级处理：使用ASCII解码
        return {
            decode: bytes => Array.from(bytes).map(b => String.fromCharCode(b)).join('')
        };
    }
}
