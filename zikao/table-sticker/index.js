const REQUIRED_HEADERS = ["考生姓名", "身份证号", "准考证号", "课程代码", "课程名称", "考场序号", "合并考场号", "座位号"];

Vue.createApp({
    data() {
        return {
            records: [],
            pages: [],
            errorMessage: "",
            continuous: false,
            perPage: 30,
            enabledFields: ["name", "admission", "courseName", "seat"],
            fields: [
                { key: "name", label: "考生姓名" },
                { key: "admission", label: "准考证号" },
                { key: "courseName", label: "课程名称" },
                { key: "seat", label: "座位号" },
                { key: "courseCode", label: "课程代码" },
                { key: "mergedRoom", label: "合并考场号" }
            ]
        };
    },
    computed: {
        stickerDensityClass() {
            if (this.enabledFields.length >= 5) return "density-compact";
            if (this.enabledFields.length <= 2) return "density-spacious";
            return "";
        }
    },
    methods: {
        hasField(field) {
            return this.enabledFields.includes(field);
        },
        normalizeHeader(value) {
            return String(value || "").replace(/[\s\u3000]/g, "");
        },
        readExcel(event) {
            const file = event.target.files[0];
            this.errorMessage = "";
            this.records = [];
            this.pages = [];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = (loadEvent) => {
                try {
                    const workbook = XLSX.read(loadEvent.target.result, { type: "array" });
                    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", raw: false });
                    const headers = (rows[0] || []).map(this.normalizeHeader);
                    const missing = REQUIRED_HEADERS.filter(header => !headers.includes(header));
                    if (missing.length) {
                        this.errorMessage = "Excel 缺少表头：" + missing.join("、");
                        return;
                    }
                    const indexes = Object.fromEntries(REQUIRED_HEADERS.map(header => [header, headers.indexOf(header)]));
                    this.records = rows.slice(1)
                        .filter(row => row.some(cell => String(cell).trim() !== ""))
                        .map(row => {
                            const record = {};
                            REQUIRED_HEADERS.forEach(header => record[header] = String(row[indexes[header]] || "").trim());
                            record.合并号 = record.合并考场号.split("-")[0].trim();
                            return record;
                        });
                    if (!this.records.length) this.errorMessage = "Excel 中没有可读取的考生数据。";
                } catch (error) {
                    this.errorMessage = "Excel 读取失败，请确认文件格式正确。";
                }
            };
            reader.readAsArrayBuffer(file);
        },
        courseInfo(student) {
            return {
                key: [student.课程代码, student.课程名称, student.考场序号, student.合并考场号].join("|"),
                课程代码: student.课程代码,
                课程名称: student.课程名称,
                考场序号: student.考场序号,
                合并考场号: student.合并考场号
            };
        },
        groupKey(student) {
            const courseKey = [student.课程代码, student.课程名称, student.考场序号].join("|");
            if (!student.合并号) return "单独考场|" + courseKey;
            return this.continuous ? "合并考场|" + student.合并号 : "课程考场|" + student.合并考场号 + "|" + courseKey;
        },
        splitIntoPages(group, courses, keyPrefix) {
            const result = [];
            for (let start = 0; start < group.length; start += this.perPage) {
                result.push({
                    key: keyPrefix + "-" + start,
                    courses,
                    groupTotal: group.length,
                    students: group.slice(start, start + this.perPage)
                });
            }
            return result;
        },
        buildPages() {
            const pages = [];
            const groups = new Map();
            this.records.forEach(student => {
                const key = this.groupKey(student);
                if (!groups.has(key)) groups.set(key, []);
                groups.get(key).push({ ...student });
            });

            groups.forEach((students, key) => {
                if (this.continuous && students[0].合并号) {
                    students.forEach((student, index) => student.displaySeat = String(index + 1));
                    const courses = [];
                    const seen = new Set();
                    students.forEach(student => {
                        const course = this.courseInfo(student);
                        if (!seen.has(course.key)) {
                            seen.add(course.key);
                            courses.push(course);
                        }
                    });
                    pages.push(...this.splitIntoPages(students, courses, key));
                } else {
                    students.forEach(student => student.displaySeat = student.座位号);
                    pages.push(...this.splitIntoPages(students, [this.courseInfo(students[0])], key));
                }
            });
            this.pages = pages;
        }
    }
}).mount("#app");
