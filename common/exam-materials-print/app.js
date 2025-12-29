// noinspection JSUnresolvedReference
const {createApp, ref, computed, watch} = Vue;

// noinspection JSUnusedGlobalSymbols
const vm = createApp({
    setup() {
        const fileSelect = ref(null);
        const fileName = ref("");
        const rawData = ref([]);
        const processedData = ref([]);
        const examName = ref("");
        const sortMode = ref("venue"); // "venue" or "session"
        const currentModule = ref("handover"); // "handover", "envelope", "bundle"
        const dragOver = ref(false);

        // 从localStorage加载边距设置
        const loadEnvelopeMargins = () => {
            const savedLeft = localStorage.getItem('envelopeMarginLeft');
            const savedTop = localStorage.getItem('envelopeMarginTop');
            return {
                left: savedLeft ? parseInt(savedLeft) : 15,
                top: savedTop ? parseInt(savedTop) : 25
            };
        };
        const margins = loadEnvelopeMargins();
        const envelopeMarginLeft = ref(margins.left);
        const envelopeMarginTop = ref(margins.top);

        // 监听边距变化，实时保存
        watch(envelopeMarginLeft, (val) => {
            localStorage.setItem('envelopeMarginLeft', val.toString());
        });
        watch(envelopeMarginTop, (val) => {
            localStorage.setItem('envelopeMarginTop', val.toString());
        });

        const messages = ref([]);

        // 添加日志消息
        const addMessage = (type, text) => {
            messages.value.push({type, text, time: new Date().toLocaleTimeString()});
            console.log(`[${type}] ${text}`);
        };

        // 选择文件
        const selectFile = () => {
            fileSelect.value?.click();
        };

        // 处理文件（统一处理函数）
        const processFile = async (file) => {
            if (!file) {
                return;
            }

            // 检查文件类型
            const validTypes = ['.xlsx', '.xls'];
            const fileExtension = file.name.toLowerCase();
            const isValidType = validTypes.some(type => fileExtension.endsWith(type));

            if (!isValidType) {
                addMessage("danger", "请选择Excel文件（.xlsx或.xls）");
                return;
            }

            fileName.value = file.name;
            messages.value = [];
            processedData.value = []; // 清空已处理的数据
            rawData.value = []; // 清空原始数据

            try {
                const data = await readExcelFile(file);
                rawData.value = data.rows;
                examName.value = data.examName;
                processData();
                addMessage("success", `成功解析 ${data.rows.length} 行数据`);
            } catch (error) {
                addMessage("danger", `解析Excel出错：${error.message}`);
                fileName.value = "";
            }
        };

        // 处理文件选择
        const handleFileChange = async (event) => {
            const file = event.target.files[0];
            event.target.value = ""; // 清空，允许重复选择同一文件
            await processFile(file);
        };

        // 处理文件拖拽
        const handleFileDrop = async (event) => {
            dragOver.value = false;
            const file = event.dataTransfer.files[0];
            await processFile(file);
        };

        // 读取Excel文件
        const readExcelFile = (file) => {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, {type: "array"});
                        const firstSheetName = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheetName];

                        // 先读取所有数据（包括第一行）
                        const allData = XLSX.utils.sheet_to_json(worksheet, {
                            header: 1,
                            defval: ""
                        });

                        // 提取考试名称（第一行包含【】）
                        let examName = "";
                        if (allData.length > 0 && Array.isArray(allData[0])) {
                            const firstRow = allData[0];
                            const firstRowStr = firstRow.join("");
                            const match = firstRowStr.match(/【([^】]+)】/);
                            if (match) {
                                examName = match[1];
                            }
                        }

                        // 从第二行开始的数据（索引1为表头，索引2开始为数据）
                        const jsonData = allData.slice(1);

                        if (jsonData.length < 2) {
                            reject(new Error("Excel文件格式错误：至少需要表头行和数据行"));
                            return;
                        }

                        const headers = jsonData[0];
                        const rows = [];

                        // 字段映射
                        const fieldMap = {
                            "考点代码": "venueId",
                            "考点名称": "venueName",
                            "场次": "session",
                            "科目": "subject",
                            "考场起始号": "roomStart",
                            "考场结束号": "roomEnd",
                            "考场数": "roomCount",
                            "条形码数": "barcodeCount",
                            "试卷捆最大单捆数": "paperMaxBundle",
                            "试卷捆分捆策略": "paperBundleStrategy",
                            "试卷捆科目合并": "paperSubjectMerge",
                            "答卷捆最大单捆数": "answerMaxBundle",
                            "答卷捆分捆策略": "answerBundleStrategy",
                            "答卷捆科目合并": "answerSubjectMerge",
                            "备用卷数": "spareVolumeCount",
                            "听力盘数": "discCount",
                            "发送方": "sender",
                            "发卷人": "issuer",
                            "交接时间前缀": "handoverTimePrefix",
                            "考试时间": "examTime"
                        };

                        // 构建字段索引映射
                        const fieldIndexMap = {};
                        headers.forEach((header, idx) => {
                            const trimmedHeader = String(header).trim();
                            if (fieldMap[trimmedHeader]) {
                                fieldIndexMap[fieldMap[trimmedHeader]] = idx;
                            }
                        });

                        // 验证必需字段
                        const requiredFields = ["venueName", "session", "subject"];
                        const missingFields = requiredFields.filter(f => !fieldIndexMap[f]);
                        if (missingFields.length > 0) {
                            reject(new Error(`缺少必需字段：${missingFields.map(f => Object.keys(fieldMap).find(k => fieldMap[k] === f)).join(", ")}`));
                            return;
                        }

                        // 解析数据行
                        for (let i = 1; i < jsonData.length; i++) {
                            const row = jsonData[i];
                            const item = {};

                            Object.keys(fieldIndexMap).forEach(field => {
                                const idx = fieldIndexMap[field];
                                const value = row[idx];
                                if (value !== undefined && value !== null) {
                                    item[field] = String(value).trim();
                                } else {
                                    item[field] = "";
                                }
                            });

                            rows.push(item);
                        }

                        resolve({rows, examName});
                    } catch (error) {
                        reject(error);
                    }
                };
                reader.onerror = () => reject(new Error("读取文件失败"));
                reader.readAsArrayBuffer(file);
            });
        };

        // 计算考场数量
        const calculateRoomCount = (roomStart, roomEnd, roomCount) => {
            // 如果已有值且不是*，直接返回
            if (roomCount && roomCount !== "*" && roomCount !== "") {
                const num = parseInt(roomCount);
                if (!isNaN(num)) {
                    return num;
                }
            }

            // 自动计算
            const startStr = String(roomStart).trim();
            const endStr = String(roomEnd).trim();

            // 分离前缀和后缀
            const startMatch = startStr.match(/^([^0-9]*)(\d+)$/);
            const endMatch = endStr.match(/^([^0-9]*)(\d+)$/);

            if (!startMatch || !endMatch) {
                return "--";
            }

            const startPrefix = startMatch[1];
            const startSuffix = parseInt(startMatch[2]);
            const endPrefix = endMatch[1];
            const endSuffix = parseInt(endMatch[2]);

            // 前缀必须一致
            if (startPrefix !== endPrefix) {
                return "--";
            }

            const count = endSuffix - startSuffix + 1;
            return count > 0 ? count : "--";
        };

        // 数据清洗和处理
        const processData = () => {
            const processed = [];
            let invalidCount = 0;

            rawData.value.forEach((row, idx) => {
                // 验证必需字段
                if (!row.venueName || row.venueName.trim() === "") {
                    invalidCount++;
                    addMessage("warning", `第 ${idx + 3} 行：考点名称为空，已跳过`);
                    return;
                }

                const session = parseInt(row.session);
                if (isNaN(session)) {
                    invalidCount++;
                    addMessage("warning", `第 ${idx + 3} 行：场次不是有效数字，已跳过`);
                    return;
                }

                // 计算考场数
                const roomCount = calculateRoomCount(row.roomStart, row.roomEnd, row.roomCount);

                // 条形码数
                let barcodeCount = roomCount;
                if (row.barcodeCount && row.barcodeCount.trim() !== "") {
                    const bc = parseInt(row.barcodeCount);
                    if (!isNaN(bc) && bc > 0) {
                        barcodeCount = bc;
                    }
                }

                // 数值字段转换
                const paperMaxBundle = parseInt(row.paperMaxBundle) || 0;
                const paperBundleStrategy = parseInt(row.paperBundleStrategy) || 0;
                const answerMaxBundle = parseInt(row.answerMaxBundle) || 0;
                const answerBundleStrategy = parseInt(row.answerBundleStrategy) || 0;

                processed.push({
                    venueId: row.venueId ? row.venueId.trim() : "",
                    venueName: row.venueName.trim(),
                    session: session,
                    subject: row.subject ? row.subject.trim() : "",
                    roomStart: row.roomStart ? row.roomStart.trim() : "",
                    roomEnd: row.roomEnd ? row.roomEnd.trim() : "",
                    roomCount: roomCount,
                    barcodeCount: barcodeCount,
                    spareVolumeCount: row.spareVolumeCount ? row.spareVolumeCount.trim() : "",
                    discCount: row.discCount ? row.discCount.trim() : "",
                    paperMaxBundle: paperMaxBundle,
                    paperBundleStrategy: paperBundleStrategy,
                    paperSubjectMerge: row.paperSubjectMerge ? row.paperSubjectMerge.trim() : "",
                    answerMaxBundle: answerMaxBundle,
                    answerBundleStrategy: answerBundleStrategy,
                    answerSubjectMerge: row.answerSubjectMerge ? row.answerSubjectMerge.trim() : "",
                    sender: row.sender ? row.sender.trim() : "",
                    issuer: row.issuer ? row.issuer.trim() : "",
                    handoverTimePrefix: row.handoverTimePrefix ? row.handoverTimePrefix.trim() : "",
                    examTime: row.examTime ? row.examTime.trim() : "",
                    originalIndex: idx // 保留原始行号，用于同一场次内保持Excel顺序
                });
            });

            processedData.value = processed;
            if (invalidCount > 0) {
                addMessage("warning", `共跳过 ${invalidCount} 行无效数据`);
            }

            console.log(processedData)
        };

        // 排序后的数据
        const sortedData = computed(() => {
            const data = [...processedData.value];
            if (sortMode.value === "venue") {
                // 考点优先：考点名称 > 场次 > Excel原始顺序
                return data.sort((a, b) => {
                    if (a.venueId !== b.venueId) {
                        return a.venueId.localeCompare(b.venueId, "zh-CN");
                    }
                    if (a.venueName !== b.venueName) {
                        return a.venueName.localeCompare(b.venueName, "zh-CN");
                    }
                    if (a.session !== b.session) {
                        return a.session - b.session;
                    }
                    // 同一场次内，按Excel原始顺序排序
                    return a.originalIndex - b.originalIndex;
                });
            } else {
                // 场次优先：场次 > 考点名称 > Excel原始顺序
                return data.sort((a, b) => {
                    if (a.session !== b.session) {
                        return a.session - b.session;
                    }
                    if (a.venueId !== b.venueId) {
                        return a.venueId.localeCompare(b.venueId, "zh-CN");
                    }
                    if (a.venueName !== b.venueName) {
                        return a.venueName.localeCompare(b.venueName, "zh-CN");
                    }
                    // 同一场次内，按Excel原始顺序排序
                    return a.originalIndex - b.originalIndex;
                });
            }
        });

        // 分组数据（用于交接单）
        const groupedData = computed(() => {
            const groups = new Map();
            const groupOrder = []; // 保持顺序
            sortedData.value.forEach(item => {
                const key = `${item.venueId}_${item.venueName}_${item.session}`;
                if (!groups.has(key)) {
                    groups.set(key, {
                        venueName: item.venueName,
                        venueId: item.venueId,
                        session: item.session,
                        items: [],
                        hasDisc: false,
                        hasSpareVolume: false
                    });
                    groupOrder.push(key); // 记录顺序
                }
                groups.get(key).items.push(item);
                // 检查是否有光盘
                if (item.discCount && item.discCount.trim() !== "" && item.discCount !== "--") {
                    groups.get(key).hasDisc = true;
                }
                // 检查是否有备用卷
                if (item.spareVolumeCount && item.spareVolumeCount.trim() !== "" && item.spareVolumeCount !== "--") {
                    groups.get(key).hasSpareVolume = true;
                }
            });
            // 按照原始顺序返回，过滤掉没有items的组
            let filter = groupOrder.map(key => groups.get(key)).filter(group => group.items.length > 0);
            console.log(filter);
            return filter;
        });

        // 生成分捆数据
        const generateBundleData = (type) => {
            // type: "paper" or "answer"
            const maxBundleField = type === "paper" ? "paperMaxBundle" : "answerMaxBundle";
            const strategyField = type === "paper" ? "paperBundleStrategy" : "answerBundleStrategy";
            const mergeField = type === "paper" ? "paperSubjectMerge" : "answerSubjectMerge";

            const bundles = [];
            const groups = new Map();
            const data = sortedData.value; // 获取排序后的数据

            // 按考点+场次分组
            data.forEach(item => {
                const groupKey = `${item.venueName}_${item.session}`;
                if (!groups.has(groupKey)) {
                    groups.set(groupKey, []);
                }
                groups.get(groupKey).push(item);
            });

            // 处理每个考点+场次组
            groups.forEach((items) => {
                const venueName = items[0].venueName;
                // 计算整个组的起止号范围（取最小起始号和最大结束号）
                let minStartStr = null, maxEndStr = null;
                items.forEach(item => {
                    if (item.roomStart && item.roomEnd) {
                        if (minStartStr === null || item.roomStart < minStartStr) {
                            minStartStr = item.roomStart;
                        }
                        if (maxEndStr === null || item.roomEnd > maxEndStr) {
                            maxEndStr = item.roomEnd;
                        }
                    }
                });

                // 按合并标识再次分组
                const mergeGroups = new Map();
                items.forEach(item => {
                    const mergeKey = item[mergeField] || `_${item.subject}`; // 无合并标识时用科目区分
                    if (!mergeGroups.has(mergeKey)) {
                        mergeGroups.set(mergeKey, []);
                    }
                    mergeGroups.get(mergeKey).push(item);
                });

                // 处理每个合并组
                mergeGroups.forEach((mergeItems) => {
                    // 计算总数（考场数之和）- 用于判断是否需要分捆
                    let totalCount = 0;
                    mergeItems.forEach(item => {
                        if (item.roomCount !== "--" && !isNaN(item.roomCount)) {
                            totalCount += parseInt(item.roomCount);
                        }
                    });

                    if (totalCount === 0) {
                        return;
                    }

                    const maxBundle = mergeItems[0][maxBundleField];
                    const strategy = mergeItems[0][strategyField];
                    // 合并科目时去重
                    const subjectList = [...new Set(mergeItems.map(i => i.subject).filter(s => s))];

                    // 获取考试时间（从第一个item获取）
                    const examTime = mergeItems[0].examTime || '';

                    // 为每个科目计算数量和起止号
                    const subjectDataList = [];
                    subjectList.forEach(subject => {
                        // 计算该科目的数量
                        const subjectItems = mergeItems.filter(i => i.subject === subject);
                        let subjectCount = 0;
                        let subjectRoomStart = null;
                        let subjectRoomEnd = null;
                        let subjectSpareVolume = 0; // 计算该科目的备用卷数量

                        subjectItems.forEach(item => {
                            if (item.roomCount !== "--" && !isNaN(item.roomCount)) {
                                subjectCount += parseInt(item.roomCount);
                            }
                            // 获取该科目的起止号
                            if (item.roomStart && item.roomEnd) {
                                if (subjectRoomStart === null || item.roomStart < subjectRoomStart) {
                                    subjectRoomStart = item.roomStart;
                                }
                                if (subjectRoomEnd === null || item.roomEnd > subjectRoomEnd) {
                                    subjectRoomEnd = item.roomEnd;
                                }
                            }
                            // 收集该科目的备用卷数量
                            if (item.spareVolumeCount && item.spareVolumeCount.trim() !== "" && item.spareVolumeCount !== "--") {
                                const spareVol = parseInt(item.spareVolumeCount);
                                if (!isNaN(spareVol) && spareVol > 0) {
                                    subjectSpareVolume += spareVol;
                                }
                            }
                        });

                        if (subjectCount > 0) {
                            subjectDataList.push({
                                subject: subject,
                                totalCount: subjectCount,
                                roomStart: subjectRoomStart || '',
                                roomEnd: subjectRoomEnd || '',
                                fullRange: subjectRoomStart && subjectRoomEnd ? `${subjectRoomStart}-${subjectRoomEnd}` : '--',
                                examTime: examTime,
                                spareVolumeCount: subjectSpareVolume > 0 ? subjectSpareVolume : null // 该科目的备用卷数量
                            });
                        }
                    });

                    if (subjectDataList.length === 0) {
                        return;
                    }

                    // 计算每个科目的捆数
                    const subjectBundleCounts = subjectDataList.map(subjData => {
                        let bundleCount = 1;
                        if (subjData.totalCount > maxBundle && strategy > 0) {
                            bundleCount = Math.ceil(subjData.totalCount / strategy);
                        }
                        return {
                            ...subjData,
                            bundleCount: bundleCount,
                            processed: false
                        };
                    });

                    // 重新组织分捆逻辑：合并的科目放在上一个科目的最后一捆
                    // 如果只有一捆或剩一捆，都放在同一页
                    let bundleIndex = 0;

                    for (let subjIdx = 0; subjIdx < subjectBundleCounts.length; subjIdx++) {
                        const subjData = subjectBundleCounts[subjIdx];
                        if (subjData.processed) continue;

                        const isLastSubject = subjIdx === subjectBundleCounts.length - 1;
                        const isOnlyOneBundle = subjData.bundleCount === 1;

                        // 如果只有一捆，检查是否可以与其他只有一捆的科目合并
                        if (isOnlyOneBundle) {
                            // 收集所有只有一捆的科目（包括当前科目）
                            const singleBundleSubjects = [subjData];
                            for (let j = subjIdx + 1; j < subjectBundleCounts.length; j++) {
                                if (subjectBundleCounts[j].bundleCount === 1 && !subjectBundleCounts[j].processed) {
                                    singleBundleSubjects.push(subjectBundleCounts[j]);
                                    subjectBundleCounts[j].processed = true;
                                } else {
                                    break; // 遇到多捆的科目就停止
                                }
                            }

                            // 将所有只有一捆的科目放在同一页（这是尾数捆，需要添加备用卷）
                            const subjectsData = singleBundleSubjects.map(s => ({
                                subject: s.subject,
                                bundleNo: 1,
                                totalBundles: 1,
                                quantity: s.totalCount,
                                totalCount: s.totalCount,
                                fullRange: s.fullRange,
                                examTime: s.examTime,
                                spareVolumeCount: s.spareVolumeCount || null // 每个科目保留自己的备用卷数量
                            }));

                            bundles.push({
                                venueName: venueName,
                                subjects: subjectsData,
                                bundleNo: bundleIndex + 1,
                                totalBundles: 1
                            });

                            bundleIndex++;
                            subjData.processed = true;
                        } else {
                            // 多捆科目：前面的捆正常显示，最后一捆可以包含合并的科目
                            for (let i = 0; i < subjData.bundleCount; i++) {
                                bundleIndex++;
                                const isLastBundle = i === subjData.bundleCount - 1;
                                const subjectsData = [];

                                // 当前科目的当前捆
                                let bundleQuantity;
                                if (subjData.bundleCount === 1) {
                                    bundleQuantity = subjData.totalCount;
                                } else if (isLastBundle) {
                                    // 最后一捆：总数减去前面所有捆的数量
                                    bundleQuantity = subjData.totalCount - (subjData.bundleCount - 1) * strategy;
                                } else {
                                    bundleQuantity = strategy;
                                }

                                subjectsData.push({
                                    subject: subjData.subject,
                                    bundleNo: i + 1,
                                    totalBundles: subjData.bundleCount,
                                    quantity: bundleQuantity,
                                    totalCount: subjData.totalCount,
                                    fullRange: subjData.fullRange,
                                    examTime: subjData.examTime,
                                    spareVolumeCount: isLastBundle && subjData.spareVolumeCount ? subjData.spareVolumeCount : null // 只在尾数捆添加当前科目的备用卷
                                });

                                // 如果是最后一捆，且后面还有合并科目，将所有只有一捆的后续科目都加入
                                if (isLastBundle && !isLastSubject) {
                                    // 检查所有后续科目，将所有只有一捆的科目合并进来
                                    // 遇到多捆的科目时跳过，但继续检查后面的科目
                                    for (let j = subjIdx + 1; j < subjectBundleCounts.length; j++) {
                                        const nextSubject = subjectBundleCounts[j];
                                        if (!nextSubject || nextSubject.processed) {
                                            continue; // 跳过已处理的科目
                                        }
                                        if (nextSubject.bundleCount === 1) {
                                            // 只有一捆的科目，合并进来
                                            subjectsData.push({
                                                subject: nextSubject.subject,
                                                bundleNo: 1,
                                                totalBundles: 1,
                                                quantity: nextSubject.totalCount,
                                                totalCount: nextSubject.totalCount,
                                                fullRange: nextSubject.fullRange,
                                                examTime: nextSubject.examTime,
                                                spareVolumeCount: nextSubject.spareVolumeCount || null // 每个科目保留自己的备用卷数量
                                            });
                                            nextSubject.processed = true;
                                        }
                                        // 如果是多捆的科目，不合并但继续检查后面的科目
                                    }
                                }

                                bundles.push({
                                    venueName: venueName,
                                    subjects: subjectsData,
                                    bundleNo: bundleIndex,
                                    totalBundles: subjData.bundleCount
                                });
                            }
                            subjData.processed = true;
                        }
                    }
                });
            });

            return bundles;
        };

        // 生成捆扎标签页面数据
        const bundlePages = computed(() => {
            const paperBundles = generateBundleData("paper");
            const answerBundles = generateBundleData("answer");

            // 按考点分组，确保考点名称只出现一次
            const processBundles = (bundles) => {
                const grouped = new Map();
                bundles.forEach(bundle => {
                    const key = bundle.venueName;
                    if (!grouped.has(key)) {
                        grouped.set(key, {
                            venueName: key,
                            bundles: []
                        });
                    }
                    grouped.get(key).bundles.push(bundle);
                });
                return Array.from(grouped.values());
            };

            const paperGroups = processBundles(paperBundles);
            const answerGroups = processBundles(answerBundles);

            // 合并试卷和答卷，按考点匹配，保持排序顺序
            const pages = [];
            // 获取所有考点，按照 sortedData 的顺序
            const venueOrder = [];
            const seenVenues = new Set();
            sortedData.value.forEach(item => {
                if (!seenVenues.has(item.venueName)) {
                    venueOrder.push(item.venueName);
                    seenVenues.add(item.venueName);
                }
            });
            // 添加其他可能存在的考点（如果 sortedData 中没有）
            [...paperGroups.map(g => g.venueName), ...answerGroups.map(g => g.venueName)].forEach(venue => {
                if (!seenVenues.has(venue)) {
                    venueOrder.push(venue);
                    seenVenues.add(venue);
                }
            });

            venueOrder.forEach(venue => {
                const paperGroup = paperGroups.find(g => g.venueName === venue);
                const answerGroup = answerGroups.find(g => g.venueName === venue);

                const paperBundlesList = paperGroup ? paperGroup.bundles : [];
                const answerBundlesList = answerGroup ? answerGroup.bundles : [];

                const maxLength = Math.max(paperBundlesList.length, answerBundlesList.length);

                for (let i = 0; i < maxLength; i++) {
                    const paperBundle = paperBundlesList[i] || null;
                    const answerBundle = answerBundlesList[i] || null;

                    // 只有当至少有一个bundle时才创建页面，避免空白页
                    if (paperBundle || answerBundle) {
                        pages.push({
                            venueName: venue,
                            paperBundle: paperBundle,
                            answerBundle: answerBundle
                        });
                    }
                }
            });

            return pages;
        });

        // 切换模块
        const switchModule = (module) => {
            currentModule.value = module;
        };

        // 打印当前模块
        const printCurrent = () => {
            // 根据当前模块设置打印方向
            if (currentModule.value === 'envelope') {
                document.body.classList.add('envelope-print');
            } else {
                document.body.classList.remove('envelope-print');
            }

            // 延迟打印，确保样式已应用
            setTimeout(() => {
                window.print();
                // 打印后移除类
                setTimeout(() => {
                    document.body.classList.remove('envelope-print');
                }, 100);
            }, 50);
        };

        // 监听排序模式变化，重新处理数据
        watch(sortMode, () => {
            // 数据已通过computed自动更新
        });

        return {
            fileSelect,
            fileName,
            processedData: sortedData,
            examName,
            sortMode,
            currentModule,
            envelopeMarginLeft,
            envelopeMarginTop,
            messages,
            dragOver,
            selectFile,
            handleFileChange,
            handleFileDrop,
            switchModule,
            printCurrent,
            groupedData,
            bundlePages
        };
    }
});

vm.mount("#app");

