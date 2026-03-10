class Student {
    constructor(data) {
        // 显式定义字段
        this.id = data.id || '';
        this.XM = data.XM || '';        // 姓名
        this.XBDM = data.XBDM || '';    // 性别代码
        this.BMXH = data.BMXH || '';    // 报名序号

        this.YK_LYSLY = data.YK_LYSLY || '';  // 右眼裸眼视力
        this.YK_LYSLZ = data.YK_LYSLZ || '';  // 左眼裸眼视力
        this.YK_JZSLY = data.YK_JZSLY || '';  // 右眼矫正视力
        this.YK_JZSLZ = data.YK_JZSLZ || '';  // 左眼矫正视力
        this.YK_JZDSY = data.YK_JZDSY || '';  // 右眼矫正度数
        this.YK_JZDSZ = data.YK_JZDSZ || '';  // 左眼矫正度数
        this.YK_SJJC = data.YK_SJJC || '';    // 色觉检查
        this.YK_YB = data.YK_YB || '';

        this.NK_XYSSY = data.NK_XYSSY || '';  // 收缩压
        this.NK_XYSZY = data.NK_XYSZY || '';  // 舒张压

        this.WK_SG = data.WK_SG || '';  // 身高
        this.WK_TZ = data.WK_TZ || '';  // 体重

        this.EB_ZETL = data.EB_ZETL || '';  // 左耳听力
        this.EB_YETL = data.EB_YETL || '';  // 右耳听力
        this.EB_XJ = data.EB_XJ || '';      // 嗅觉

        this.TJSXBZ = data.TJSXBZ || '';  // 体检受限标志
        this.TJYJBZ = data.TJYJBZ || '';  // 体检依据标志
        this.TJJLDM = data.TJJLDM || '';  // 体检结论代码

        // 错误相关属性
        this.errorList = [];

        // 计算属性
        this.eyes = [
            {
                name: "右",
                val: parseFloat(this.YK_LYSLY),
                jzsl: parseFloat(this.YK_JZSLY),
                jzds: parseInt(this.YK_JZDSY)
            },
            {name: "左", val: parseFloat(this.YK_LYSLZ), jzsl: parseFloat(this.YK_JZSLZ), jzds: parseInt(this.YK_JZDSZ)}
        ];
        this.ss = parseInt(this.NK_XYSSY);
        this.sz = parseInt(this.NK_XYSZY);
        this.sg = parseFloat(this.WK_SG);
        this.tz = parseFloat(this.WK_TZ);
        this.bmi = this.tz / ((this.sg / 100) ** 2);
        this.etlL = parseFloat(this.EB_ZETL);
        this.etlR = parseFloat(this.EB_YETL);
    }

    // 获取所有字段名，用于表格表头展示
    static getFields() {
        return [
            '序号',
            '可能存在异常的检查信息',
            'XM', 'XBDM', 'BMXH',
            'YK_LYSLY', 'YK_LYSLZ',
            'YK_JZSLY', 'YK_JZSLZ',
            'YK_JZDSY', 'YK_JZDSZ',
            'YK_SJJC', 'YK_YB',
            'NK_XYSSY', 'NK_XYSZY',
            'WK_SG', 'WK_TZ',
            'EB_ZETL', 'EB_YETL',
            'EB_XJ',
            'TJSXBZ', 'TJYJBZ',
            'TJJLDM'
        ];
    }

    checkValidate() {
        const pushError = (desc, val) => {
            this.errorList.push(`${desc} (实测值: ${val})`);
        };

        // 裸眼视力校验
        for (let eye of this.eyes) {
            const {name, val, jzsl, jzds} = eye;
            if (val === 0 || val < 4.0 || val > 5.4) pushError(`裸眼视力(${name}眼)偏离正常值(4.0-5.4)`, val);

            if (val < 5.0 && this.TJSXBZ[4] !== '1') pushError(`${name}眼视力低于5.0，但体检受限标志(第5位)不匹配`, this.TJSXBZ);
            if (val < 4.8 && this.TJSXBZ[5] !== '1') pushError(`${name}眼视力低于4.8，但体检受限标志(第6位)不匹配`, this.TJSXBZ);

            // 矫正视力校验
            if (val >= 4.8 && jzsl) pushError(`${name}眼视力正常，但包含矫正视力数据`, jzsl);
            if (val < 4.8 && jzsl !== 4.8) pushError(`${name}眼视力低于4.8，但矫正视力非标准值(4.8)`, jzsl);

            // 矫正度数校验
            if (jzds) {
                if (jzds < 0 || jzds > 1200) pushError(`${name}眼矫正度数偏离正常范围(0-1200)`, jzds);
                if (jzds < 100 && val < 4.5) pushError(`${name}眼矫正度数较小(${jzds})，但裸眼视力偏低`, val);
                if (jzds > 400 && val > 4.6) pushError(`${name}眼矫正度数较大(${jzds})，但裸眼视力较高`, val);
                if (jzds > 800 && val > 4.4) pushError(`${name}眼矫正度数较大(${jzds})，但裸眼视力较高`, val);
                if (jzds > 400 && this.TJSXBZ[10] !== '1') pushError(`${name}眼矫正度数超过400，但体检受限标志(第11位)不匹配`, this.TJSXBZ);
                if (jzds > 800 && this.TJSXBZ[11] !== '1') pushError(`${name}眼矫正度数超过800，但体检受限标志(第12位)不匹配`, this.TJSXBZ);
            }
        }

        // 矫正度数+裸眼为0校验
        if ((parseFloat(this.YK_LYSLY) === 0 || parseFloat(this.YK_LYSLZ) === 0) &&
            (parseInt(this.YK_JZDSY) > 400 || parseInt(this.YK_JZDSZ) > 400)) {
            if (this.TJSXBZ[12] !== '1') pushError(`一眼裸眼为0且另一眼矫正度数大于400，但体检受限标志(第13位)不匹配`, this.TJSXBZ);
        }

        // 色觉校验
        if (parseInt(this.YK_SJJC) === 2 && !this.TJSXBZ.slice(1, 4).includes("1")) {
            pushError(`色觉检查异常，但体检受限标志(第2-4位)未设置`, this.TJSXBZ);
        }

        // 血压校验
        if (this.ss < 80 || this.ss > 160) pushError("收缩压偏离正常值较大", this.ss);
        if (this.sz < 50 || this.sz > 110) pushError("舒张压偏离正常值较大", this.sz);

        // 身高体重校验
        let minBmi = 14.5, maxBmi = 40;
        let minSg = 140, maxSg = 200;
        let minTz = 40, maxTz = 120;

        // 根据性别调整阈值 (1:男, 2:女)
        if (this.XBDM === '1') {
            minSg = 145; // 男生身高下限稍高
            minTz = 45;  // 男生体重下限稍高
        } else if (this.XBDM === '2') {
            maxSg = 190; // 女生身高上限稍低
            maxTz = 100; // 女生体重上限稍低
        }

        if (this.bmi < minBmi || this.bmi > maxBmi) pushError("BMI指数偏离正常范围", this.bmi.toFixed(2) + " (身高:" + this.sg + ", 体重:" + this.tz + ")");
        if (this.sg < minSg || this.sg > maxSg) pushError("身高偏离正常值较大", this.sg);
        if (this.tz < minTz || this.tz > maxTz) pushError("体重偏离正常值过大", this.tz);

        // 听力检查
        if ((this.etlL < 3 && this.etlR < 3) || (this.etlL === 5 && this.etlR === 0) || (this.etlR === 5 && this.etlL === 0)) {
            if (this.TJSXBZ[13] !== '1') pushError("听力检查异常，但体检受限标志(第14位)不匹配", this.TJSXBZ);
        }

        // 嗅觉检查
        if (parseInt(this.EB_XJ) === 2 && this.TJSXBZ[14] !== '1') {
            pushError("嗅觉检查异常，但体检受限标志(第15位)不匹配", this.TJSXBZ);
        }

        // TJSXBZ逻辑校验
        const validFlag = validateTjsxbz(this.TJSXBZ, this.TJYJBZ);
        if (validFlag !== true) pushError("体检受限标志逻辑校验未通过", validFlag);

        // TJJLDM 为1时要求 TJSXBZ 第1位为1其余为0
        if (parseInt(this.TJJLDM) === 1) {
            if (this.TJSXBZ[0] !== '1' || this.TJSXBZ.slice(1).includes('1')) {
                pushError("体检结论为合格，但体检受限标志不符合规范(应仅第1位为1)", this.TJSXBZ);
            }
        }
    }
}

// 模拟外部函数：只提供框架
function validateTjsxbz(tjsxbz, tjyjbz) {
    // 校验长度和是否只包含0或1
    if (!/^[01]{22}$/.test(tjsxbz)) {
        return "格式错误，不是22位0或者1组成";
    }

    // 第1位为1，其余必须全为0
    if (tjsxbz[0] === '1' && /1/.test(tjsxbz.slice(1))) {
        return "合格学生有其他受限专业";
    }

    if (!tjyjbz) {
        return true;
    }

    // 分组：第1组(1位)、第2组(2-7位)、第3组(8-16位)、第4组(17-22位)
    const group2 = tjsxbz.slice(1, 7);  // 第2组，6位，对应标记为2
    const group3 = tjsxbz.slice(7, 16); // 第3组，9位，对应标记为3
    const group4 = tjsxbz.slice(16);    // 第4组，6位，对应标记为1

    let bz = '';

    // 遍历构建 bz 字符串
    const buildBz = (group, tag) => {
        for (let i = 0; i < group.length; i++) {
            if (group[i] === '1') {
                bz += `${tag}${i + 1}`;
            }
        }
    };

    buildBz(group2, 2); // 第2组，标记2
    buildBz(group3, 3); // 第3组，标记3
    buildBz(group4, 1); // 第4组，标记1

    return bz === tjyjbz;
}
