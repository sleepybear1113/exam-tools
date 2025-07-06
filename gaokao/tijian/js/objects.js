class Student {
    constructor(data) {
        // 显式定义字段
        this.id = data.id || '';
        this.XM = data.XM || '';        // 姓名
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
            '异常信息',
            'XM', 'BMXH',
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
            this.errorList.push(`${desc} 异常值: ${val}`);
        };

        // 裸眼视力校验
        for (let eye of this.eyes) {
            const {name, val, jzsl, jzds} = eye;
            if (val === 0 || val < 4.0 || val > 5.4) pushError(`裸眼视力(${name}眼)不在4.0-5.4范围内`, val);

            if (val < 5.0 && this.TJSXBZ[4] !== '1') pushError(`${name}眼视力低于5.0但TJSXBZ第5位不为1`, this.TJSXBZ);
            if (val < 4.8 && this.TJSXBZ[5] !== '1') pushError(`${name}眼视力低于4.8但TJSXBZ第6位不为1`, this.TJSXBZ);

            // 矫正视力校验
            if (val >= 4.8 && jzsl) pushError(`${name}眼视力正常但有矫正视力`, jzsl);
            if (val < 4.8 && jzsl !== 4.8) pushError(`${name}眼视力低于4.8但矫正视力不为4.8`, jzsl);

            // 矫正度数校验
            if (jzds) {
                if (jzds < 0 || jzds > 1200) pushError(`${name}眼矫正度数超范围`, jzds);
                if (jzds < 100 && val < 4.5) pushError(`${name}眼矫正度数<100但裸眼视力低于4.5`, val);
                if (jzds > 400 && val > 4.6) pushError(`${name}眼矫正度数>400但裸眼视力高于4.6`, val);
                if (jzds > 800 && val > 4.4) pushError(`${name}眼矫正度数>800但裸眼视力高于4.4`, val);
                if (jzds > 400 && this.TJSXBZ[10] !== '1') pushError(`${name}眼矫正度数>400但TJSXBZ第11位不为1`, this.TJSXBZ);
                if (jzds > 800 && this.TJSXBZ[11] !== '1') pushError(`${name}眼矫正度数>800但TJSXBZ第12位不为1`, this.TJSXBZ);
            }
        }

        // 矫正度数+裸眼为0校验
        if ((parseFloat(this.YK_LYSLY) === 0 || parseFloat(this.YK_LYSLZ) === 0) &&
            (parseInt(this.YK_JZDSY) > 400 || parseInt(this.YK_JZDSZ) > 400)) {
            if (this.TJSXBZ[12] !== '1') pushError(`一眼裸眼为0另一眼矫正度数>400，但TJSXBZ第13位不为1`, this.TJSXBZ);
        }

        // 色觉校验
        if (parseInt(this.YK_SJJC) === 2 && !this.TJSXBZ.slice(1, 4).includes("1")) {
            pushError(`色觉检查异常但TJSXBZ第2-4位未设置`, this.TJSXBZ);
        }

        // 血压校验
        if (this.ss < 80 || this.ss > 160) pushError("收缩压异常", this.ss);
        if (this.sz < 50 || this.sz > 110) pushError("舒张压异常", this.sz);

        // 身高体重校验
        if (this.bmi < 14.5 || this.bmi > 40) pushError("BMI异常", this.bmi.toFixed(2) + "身高:" + this.sg + "体重:" + this.tz);
        if (this.sg < 140 || this.sg > 200) pushError("身高不在正常范围", this.sg);
        if (this.tz < 40 || this.tz > 120) pushError("体重不在正常范围", this.tz);

        // 听力检查
        if ((this.etlL < 3 && this.etlR < 3) || (this.etlL === 5 && this.etlR === 0) || (this.etlR === 5 && this.etlL === 0)) {
            if (this.TJSXBZ[13] !== '1') pushError("听力异常但TJSXBZ第14位不为1", this.TJSXBZ);
        }

        // 嗅觉检查
        if (parseInt(this.EB_XJ) === 2 && this.TJSXBZ[14] !== '1') {
            pushError("嗅觉异常但TJSXBZ第15位不为1", this.TJSXBZ);
        }

        // TJSXBZ逻辑校验
        const validFlag = validateTjsxbz(this.TJSXBZ, this.TJYJBZ);
        if (validFlag !== true) pushError("体检受限标志校验失败，TJSXBZ与TJSXBZ无法对应", validFlag);

        // TJJLDM 为1时要求 TJSXBZ 第1位为1其余为0
        if (parseInt(this.TJJLDM) === 1) {
            if (this.TJSXBZ[0] !== '1' || this.TJSXBZ.slice(1).includes('1')) {
                pushError("体检结论为1但TJSXBZ不符合要求（第1位为1，其余为0）", this.TJSXBZ);
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
