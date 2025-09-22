import {readExcel} from "./excel-utils.js";

const {createApp, ref, computed, nextTick} = Vue;

let vm = createApp({

    setup() {
        const emptyPng = ref("data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");

        const examName = ref("20xxxxx年嘉兴市职业技能操作考试");

        const fileSelect = ref(null);
        const fileName = ref("");
        const picFolderSelect = ref(null);
        const picFolderName = ref("");
        const picFileList = ref([]);
        const studentsRaw = ref([]);
        const roomsRaw = ref([]);
        const arrangedRooms = ref([]);
        const random = ref(false); // 是否随机打乱学生顺序

        const messages = ref([]);
        const logTextarea = ref(null);

        const tab = ref("核对单");

        const picShown = ref(true);
        const examNotice = ref("");

        // 准考证排序相关变量
        const admissionTicketSortKey = ref("className");
        const admissionTicketSecondarySortKey = ref("");

        const arrangeConfig = ref({
            title1: "",
            title2: "",
            hdd: {
                showKsh: true,
                showName: true,
                showIdCard: true,
                showSex: true,
                showSeatNo: true,
                withPic: {},
                hddNoPic: {},
            },

        });

        const switchTab = (newTab) => {
            tab.value = newTab;
        }

        const switchPicShown = (shown) => {
            picShown.value = shown;
        }

        const addMessage = (type, text) => {
            let msg = new Message(new Date(), type, text);
            messages.value.push(msg);

            // 等待DOM更新后，将日志滚动到最底部
            nextTick(() => {
                try {
                    const el = logTextarea.value;
                    if (el) {
                        el.scrollTop = el.scrollHeight;
                    }
                } catch (e) {
                    // ignore
                }
            });
        };

        // 匹配照片到学生信息
        const matchPhotosToStudents = () => {
            if (!picFileList.value || picFileList.value.length === 0) {
                addMessage("warning", "没有照片文件可匹配");
                return;
            }

            if (!arrangedRooms.value || arrangedRooms.value.length === 0) {
                addMessage("warning", "等待安排试场(顺序/随机)后再匹配照片");
                return;
            }

            // 创建照片文件名到File对象的映射，去除文件扩展名
            const photoMap = new Map();
            picFileList.value.forEach(file => {
                const fileName = file.name;
                const nameWithoutExt = fileName.replace(/\.[^/.]+$/, ""); // 去除扩展名
                photoMap.set(nameWithoutExt, file);
            });

            let matchedCount = 0;
            let unmatchedPhotos = [];

            // 遍历所有安排好的考场中的学生，尝试匹配照片
            arrangedRooms.value.forEach(room => {
                if (room.students && room.students.length > 0) {
                    room.students.forEach(student => {
                        if (student.photo) {
                            return; // 已经有照片了，跳过
                        }

                        let matchedPhoto = null;

                        // 尝试用准考证号匹配
                        if (student.ksh && photoMap.has(student.ksh)) {
                            matchedPhoto = photoMap.get(student.ksh);
                        }
                        // 尝试用身份证号匹配
                        else if (student.idCard && photoMap.has(student.idCard)) {
                            matchedPhoto = photoMap.get(student.idCard);
                        }

                        if (matchedPhoto) {
                            // 创建图片URL
                            student.photo = URL.createObjectURL(matchedPhoto);
                            matchedCount++;
                        }
                    });
                }
            });

            // 统计未匹配的照片
            const matchedPhotoNames = new Set();
            arrangedRooms.value.forEach(room => {
                if (room.students && room.students.length > 0) {
                    room.students.forEach(student => {
                        if (student.photo) {
                            // 从URL中提取文件名（这里简化处理）
                            const photoName = student.ksh && photoMap.has(student.ksh) ? student.ksh : 
                                            student.idCard && photoMap.has(student.idCard) ? student.idCard : null;
                            if (photoName) {
                                matchedPhotoNames.add(photoName);
                            }
                        }
                    });
                }
            });

            picFileList.value.forEach(file => {
                const nameWithoutExt = file.name.replace(/\.[^/.]+$/, "");
                if (!matchedPhotoNames.has(nameWithoutExt)) {
                    unmatchedPhotos.push(file.name);
                }
            });
            addMessage("success", `照片匹配完成: 成功匹配 ${matchedCount} 张照片，${unmatchedPhotos.length} 张未匹配`);
        };


        const selectFile = () => {
            fileSelect.value.click();
        };

        const selectPicFolder = () => {
            picFolderSelect.value.click();
        };

        const handlePicFolderChange = (event) => {
            const files = event.target.files;
            if (!files || files.length === 0) {
                picFolderName.value = "";
                picFileList.value = [];
                event.target.value = ""; // 清空，保证能重复选择同一个文件夹
                return;
            }

            // 获取文件夹名称（从第一个文件的路径中提取）
            const firstFile = files[0];
            const folderPath = firstFile.webkitRelativePath || firstFile.name;
            const folderName = folderPath.split('/')[0] || "未知文件夹";
            
            // 过滤出图片文件
            const imageFiles = Array.from(files).filter(file => {
                const fileName = file.name.toLowerCase();
                return fileName.endsWith('.jpg') || 
                       fileName.endsWith('.jpeg') || 
                       fileName.endsWith('.png') || 
                       fileName.endsWith('.gif') || 
                       fileName.endsWith('.bmp') || 
                       fileName.endsWith('.webp');
            });

            picFolderName.value = folderName;
            picFileList.value = imageFiles;
            event.target.value = ""; // 清空，保证能重复选择同一个文件夹

            addMessage("success", `成功选择文件夹: ${folderName}，包含 ${imageFiles.length} 个图片文件`);
            
            // 自动匹配照片到学生
            matchPhotosToStudents();
        };

        const handleFileChange = async (event) => {
            const file = event.target.files[0];
            if (!file) {
                fileName.value = "";
                event.target.value = ""; // 关键：清空，保证能重复选择同一个文件
                return;
            }

            fileName.value = file.name;
            let dataList = null;
            try {
                dataList = await readExcel(file);
            } catch (e) {
                console.error("解析Excel出错：", e);
                addMessage("danger", `解析Excel出错：${e.message}`);
                fileName.value = "";
                event.target.value = ""; // 关键：清空，保证能重复选择同一个文件
                return;
            }
            fileName.value = file.name;
            event.target.value = ""; // 关键：清空，保证能重复选择同一个文件

            let {students, rooms} = processExcelData(dataList);
            if (!rooms || rooms.length === 0) {
                console.warn("未解析到有效的考场数据");
                addMessage("warning", "未解析到有效的考场数据");
                return;
            }

            if (!students || students.length === 0) {
                console.warn("未解析到有效的学生数据");
                addMessage("warning", "未解析到有效的学生数据");
                return;
            }

            let arrangeRooms = processRoomsData(rooms);

            studentsRaw.value = students;
            roomsRaw.value = arrangeRooms;
        };

        const processRoomsData = (rooms) => {
            let arrangeRooms = []; // 最终用于安排的考场列表

            let roomIdMaxLength = 2;
            let roomsObj = {};
            rooms.forEach(room => {
                let key = `${room.schoolId}||${room.typeId}`;
                if (!roomsObj[key]) {
                    roomsObj[key] = [];
                }
                roomsObj[key].push(room);
                if (room.roomId && room.roomId.length > roomIdMaxLength) {
                    roomIdMaxLength = room.roomId.length;
                }
            });
            let roomsAll = {};
            for (let key in roomsObj) {
                let roomList = roomsObj[key];
                let currRoomsAll = [];
                roomList.forEach(room => {
                    let roomCountStr = room.count;
                    if (!roomCountStr || isNaN(roomCountStr)) {
                        roomCountStr = "30";
                    }
                    let roomCount = parseInt(roomCountStr);

                    let addCurrRoomsAll
                    let roomIdStr = room.roomId;
                    if (roomIdStr) {
                        let expandRangesRes = expandRanges(roomIdStr);
                        if (expandRangesRes.error) {
                            console.error(`考点 ${room.schoolName} (${room.schoolId})，类别 ${room.typeName} (${room.typeId})，试场 ${room.roomName} (${room.roomId}) 的试场号解析错误：${expandRangesRes.error}，跳过该考场`);
                            addMessage("warning", `考点 ${room.schoolName} (${room.schoolId})，类别 ${room.typeName} (${room.typeId})，试场 ${room.roomName} (${room.roomId}) 的试场号解析错误：${expandRangesRes.error}，跳过该考场`);
                            return;
                        }
                        let roomIds = expandRangesRes.result.map(Number);
                        addCurrRoomsAll = mergeRoomIds(currRoomsAll, {count: null, arr: roomIds});
                    } else {
                        addCurrRoomsAll = mergeRoomIds(currRoomsAll, {count: roomCount, arr: null});
                    }

                    roomsAll[room.id] = addCurrRoomsAll;
                    currRoomsAll.push(...addCurrRoomsAll);

                    for (let i = 0; i < addCurrRoomsAll.length; i++) {
                        let newRoom = room.copy();
                        newRoom.id = `${room.id}-${String(i + 1).padStart(roomIdMaxLength, '0')}`;
                        newRoom.roomId = String(addCurrRoomsAll[i]).padStart(roomIdMaxLength, '0');
                        arrangeRooms.push(newRoom);
                    }
                });
            }

            return arrangeRooms;
        };

        const processExcelData = (dataList) => {
            let result = {students: [], rooms: []};
            if (!dataList || dataList.length === 0) {
                console.warn("数据为空");
                addMessage("warning", "数据为空");
                return result;
            }
            
            examName.value = dataList.shift()?.[0] || "考试安排";

            let splitSymbol = "!分割列!";
            let splitSymbolIndex = -1;
            let header = dataList[0];

            for (let i = 0; i < header.length; i++) {
                if (header[i] === splitSymbol) {
                    splitSymbolIndex = i;
                    break;
                }
            }

            if (splitSymbolIndex === -1) {
                console.warn("未找到分割列");
                addMessage("danger", `未找到分割列，请确保第一行包含 "${splitSymbol}" 列`);
                return result;
            }

            let examNoticeIndex = -1;
            let studentHeaderIndexes = new Student();
            let roomHeaderIndexes = new Room();
            for (let i = 0; i < header.length; i++) {
                let colName = header[i]?.trim();
                if (!colName) {
                    continue;
                }

                if (i < splitSymbolIndex) {
                    if (colName === "考生号") {
                        studentHeaderIndexes.ksh = i;
                    } else if (colName === "姓名") {
                        studentHeaderIndexes.name = i;
                    } else if (colName === "类别代码") {
                        studentHeaderIndexes.typeId = i;
                    } else if (colName === "身份证号") {
                        studentHeaderIndexes.idCard = i;
                    } else if (colName === "考点代码") {
                        studentHeaderIndexes.schoolId = i;
                    } else if (colName === "考点名称") {
                        studentHeaderIndexes.schoolName = i;
                    } else if (colName === "试场号") {
                        studentHeaderIndexes.roomId = i;
                    } else if (colName === "试场名称") {
                        studentHeaderIndexes.roomName = i;
                    } else if (colName === "座位号") {
                        studentHeaderIndexes.seatNo = i;
                    } else if (colName === "考试时间") {
                        studentHeaderIndexes.time = i;
                    } else if (colName === "班级") {
                        studentHeaderIndexes.className = i;
                    }
                } else {
                    if (colName === "考点代码") {
                        roomHeaderIndexes.schoolId = i;
                    } else if (colName === "考点名称") {
                        roomHeaderIndexes.schoolName = i;
                    } else if (colName === "类别代码") {
                        roomHeaderIndexes.typeId = i;
                    } else if (colName === "类别名称") {
                        roomHeaderIndexes.typeName = i;
                    } else if (colName === "考试时间") {
                        roomHeaderIndexes.time = i;
                    } else if (colName === "试场号") {
                        roomHeaderIndexes.roomId = i;
                    } else if (colName === "试场名称") {
                        roomHeaderIndexes.roomName = i;
                    } else if (colName === "容纳人数") {
                        roomHeaderIndexes.capacity = i;
                    } else if (colName === "该考试时间下的试场数") {
                        roomHeaderIndexes.count = i;
                    } else if (colName === "备注") {
                        roomHeaderIndexes.remark = i;
                    }
                }

                if (header[i] === "准考证考生必读") {
                    examNoticeIndex = i;
                }
            }

            // 解析数据行
            
            if (examNoticeIndex !== -1) {
                let data = dataList[1][examNoticeIndex];
                if (data && typeof data === "string") {
                    examNotice.value = data.trim();
                }
            }

            let students = [];
            let rooms = [];
            for (let i = 1; i < dataList.length; i++) {
                let row = dataList[i];

                // 遍历每个单元格，当单元格以#开头时，将当前单元格设置为空字符串
                for (let j = 0; j < row.length; j++) {
                    if (typeof row[j] === "string" && row[j].startsWith("#")) {
                        row[j] = "";
                    }
                }

                for (let j = 0; j < row.length; j++) {
                    if (typeof row[j] === "string") {
                        row[j] = row[j].trim();
                    }
                }

                // 解析学生
                let student = new Student(
                    students.length,
                    row[studentHeaderIndexes.name],
                    row[studentHeaderIndexes.ksh],
                    row[studentHeaderIndexes.typeId],
                    row[studentHeaderIndexes.roomId],
                    row[studentHeaderIndexes.roomName],
                    null,
                    row[studentHeaderIndexes.seatNo],
                    row[studentHeaderIndexes.idCard],
                    row[studentHeaderIndexes.schoolId],
                    row[studentHeaderIndexes.schoolName],
                    row[studentHeaderIndexes.time],
                    row[studentHeaderIndexes.className]
                );

                if (student.validate()) {
                    students.push(student);
                } else {
                    addMessage("warning", `第${i + 1}行学生数据不完整，跳过：${JSON.stringify(row)}`);
                }

                // 解析考场
                let room = new Room(
                    rooms.length,
                    row[roomHeaderIndexes.schoolId],
                    row[roomHeaderIndexes.schoolName],
                    row[roomHeaderIndexes.typeId],
                    row[roomHeaderIndexes.typeName],
                    row[roomHeaderIndexes.time],
                    row[roomHeaderIndexes.roomId],
                    row[roomHeaderIndexes.roomName],
                    row[roomHeaderIndexes.count],
                    row[roomHeaderIndexes.capacity],
                    row[roomHeaderIndexes.remark]
                );
                if (room.validate()) {
                    rooms.push(room);
                } else if (!room.allEmptyProps()) {
                    addMessage("warning", `第${i + 1}行考场数据不完整，跳过：${JSON.stringify(row)}`);
                }
            }

            addMessage("info", `解析完成，共${students.length}个学生，${rooms.length}个考场`);
            return {students, rooms};
        }

        const expandRanges = (str, mustContinuous) => {
            const numPattern = /^[0-9]+$/;
            let result = [];

            // 分割输入
            let parts = str.split(/[,，]/).map(s => s.trim()).filter(Boolean);

            if (mustContinuous) {
                if (parts.length !== 1) {
                    return {result: [], error: "必须是单个数字或单个范围表达式"};
                }

                let part = parts[0];

                if (part.includes('-')) {
                    // 范围
                    let rangeParts = part.split('-').map(s => s.trim());
                    if (rangeParts.length !== 2) {
                        return {result: [], error: "范围表达式格式不正确"};
                    }

                    let [start, end] = rangeParts;
                    if (!numPattern.test(start) || !numPattern.test(end)) {
                        return {result: [], error: `范围 ${part} 含有非法字符`};
                    }

                    let startNum = parseInt(start, 10);
                    let endNum = parseInt(end, 10);

                    if (startNum < 1 || endNum < 1) {
                        return {result: [], error: `范围 ${part} 包含非自然数`};
                    }
                    if (startNum > endNum) {
                        return {result: [], error: `范围 ${part} 起始值大于结束值`};
                    }

                    let width = Math.max(start.length, end.length);
                    for (let i = startNum; i <= endNum; i++) {
                        result.push(i.toString().padStart(width, '0'));
                    }
                    return {result, error: null};

                } else {
                    // 单个数字
                    if (!numPattern.test(part)) {
                        return {result: [], error: `值 ${part} 含有非法字符`};
                    }
                    let num = parseInt(part, 10);
                    if (num < 1) {
                        return {result: [], error: `值 ${part} 不是自然数`};
                    }
                    return {result: [part], error: null};
                }
            }

            // mustContinuous = false 的情况
            for (let part of parts) {
                if (part.includes('-')) {
                    let rangeParts = part.split('-').map(s => s.trim());
                    if (rangeParts.length !== 2) {
                        return {result: [], error: `范围 ${part} 格式不正确`};
                    }

                    let [start, end] = rangeParts;
                    if (!numPattern.test(start) || !numPattern.test(end)) {
                        return {result: [], error: `范围 ${part} 含有非法字符`};
                    }

                    let startNum = parseInt(start, 10);
                    let endNum = parseInt(end, 10);

                    if (startNum < 1 || endNum < 1) {
                        return {result: [], error: `范围 ${part} 包含非自然数`};
                    }
                    if (startNum > endNum) {
                        return {result: [], error: `范围 ${part} 起始值大于结束值`};
                    }

                    let width = Math.max(start.length, end.length);
                    for (let i = startNum; i <= endNum; i++) {
                        result.push(i.toString().padStart(width, '0'));
                    }
                } else {
                    if (!numPattern.test(part)) {
                        return {result: [], error: `值 ${part} 含有非法字符`};
                    }
                    let num = parseInt(part, 10);
                    if (num < 1) {
                        return {result: [], error: `值 ${part} 不是自然数`};
                    }
                    result.push(part);
                }
            }

            return {result, error: null};
        }

        const mergeRoomIds = (roomIds, {count = null, arr = null} = {}) => {
            if (count != null && arr != null) {
                throw new Error("只能使用 count 或 arr，一个参数必须为空");
            }

            const maxId = roomIds.length > 0 ? Math.max(...roomIds) : 0;
            let added = [];

            if (count != null) {
                for (let i = 1; i <= count; i++) {
                    added.push(maxId + i);
                }
            }

            if (arr != null) {
                const minArr = Math.min(...arr);
                if (minArr > maxId) {
                    added = [...arr];
                } else {
                    throw new Error("arr 中的最小值必须大于 roomIds 最大值");
                }
            }

            return added;
        }

        const arrangeStuIntoRooms = (randomInput) => {
            // 如果已经有安排结果，提示用户确认是否重新安排
            if (arrangedRooms.value && arrangedRooms.value.length > 0) {
                if (!confirm("已经有安排结果，是否重新安排？")) {
                    return;
                }
            }
            // 按照考点、类别、时间分组学生
            let studentCopy = [];
            studentsRaw.value.forEach(student => studentCopy.push(student.copy()));
            if (randomInput) {
                studentCopy = studentCopy.sort(() => Math.random() - 0.5);
            }

            let roomsCopy = [];
            roomsRaw.value.forEach(room => {
                let roomClone = room.copy();
                roomClone.students = [];
                roomsCopy.push(roomClone);
            });

            let roomsObj = {};
            roomsCopy.forEach(room => {
                let key = `${room.schoolId}||${room.typeId}}`;
                if (!roomsObj[key]) {
                    roomsObj[key] = [];
                }
                roomsObj[key].push(room);
            });

            // 遍历学生，依次安排到对应的考场
            studentCopy.forEach(student => {
                let key = `${student.schoolId || ""}||${student.typeId || ""}}`;
                let candidateRooms = roomsObj[key];
                if (candidateRooms && candidateRooms.length > 0) {
                    // 找到一个未满员的考场
                    for (let room of candidateRooms) {
                        if (room.students.length < room.capacity) {
                            student.roomId = room.roomId;
                            student.roomName = room.roomName;
                            student.seatNo = room.students.length + 1;

                            room.students.push(student);
                            break;
                        }
                    }

                } else {
                    console.warn(`未找到考点 ${student.schoolName} (${student.schoolId})，类别 ${student.typeName} (${student.typeId}) 的考场，无法安排该学生 ${student.name} (${student.ksh})`);
                    addMessage("warning", `未找到考点 ${student.schoolName} (${student.schoolId})，类别 ${student.typeName} (${student.typeId}) 的考场，无法安排该学生 ${student.name} (${student.ksh})`);
                }
            });

            // 输出安排结果，仅包含有学生的考场
            let res = roomsCopy.filter(room => room.students.length > 0);
            arrangedRooms.value = res;
            
            // 自动匹配照片到学生
            matchPhotosToStudents();

        }

        const getRoomRows = (room) => {
            if (!room || !room.students) {
                return 0;
            }

            return Math.ceil(room.students.length / 5);
        }

        const chunkStudents = (room) => {
            if (!room || !Array.isArray(room.students)) {
                return [];
            }
            const chunkSize = 5;
            const chunks = [];
            for (let i = 0; i < room.students.length; i += chunkSize) {
                chunks.push(room.students.slice(i, i + chunkSize));
            }
            return chunks;
        }

        // 相同学生分组列表（使用Computed自动计算）
        const sameStudentGroupList = computed(() => {
            if (!arrangedRooms.value || arrangedRooms.value.length === 0) {
                return [];
            }

            const groupMap = new Map();

            // 遍历所有安排好的考场中的学生
            arrangedRooms.value.forEach(room => {
                if (room.students && Array.isArray(room.students)) {
                    room.students.forEach(student => {
                        // 使用考生号、姓名、身份证号作为唯一标识
                        const key = `${student.ksh}-${student.name}-${student.idCard}`;

                        if (!groupMap.has(key)) {
                            groupMap.set(key, new SameStudentGroup(student.id, student.className, student.ksh, student.name, student.idCard));
                        }

                        // 将学生添加到对应的分组中
                        let s = groupMap.get(key);
                        s.students.push(student || {});
                        s.rooms.push(room || {});
                    });
                }
            });

            // 转换为数组并过滤掉只有一个学生的分组
            const groups = Array.from(groupMap.values());

            return groups;
        });

        // 准考证排序（自动）：根据所选排序键自动返回排序后的分组
        const sortedSameStudentGroupList = computed(() => {
            if (!sameStudentGroupList.value || sameStudentGroupList.value.length === 0) {
                return [];
            }
            const sortedGroups = [...sameStudentGroupList.value].sort((a, b) => {
                // 首要排序
                let result = 0;
                if (admissionTicketSortKey.value) {
                    const aVal = a[admissionTicketSortKey.value] || '';
                    const bVal = b[admissionTicketSortKey.value] || '';
                    result = aVal.localeCompare(bVal, 'zh-CN');
                }
                // 次要排序
                if (result === 0 && admissionTicketSecondarySortKey.value) {
                    const aVal = a[admissionTicketSecondarySortKey.value] || '';
                    const bVal = b[admissionTicketSecondarySortKey.value] || '';
                    result = aVal.localeCompare(bVal, 'zh-CN');
                }
                return result;
            });
            return sortedGroups;
        });

        // 重置排序方法
        const resetAdmissionTicketSort = () => {
            admissionTicketSortKey.value = "className";
            admissionTicketSecondarySortKey.value = "";
            addMessage('info', '已重置准考证排序条件');
        };

        // 将messages转换为文本格式
        const messagesText = computed(() => {
            return messages.value.map(msg => {
                const timeStr = msg.msgTime.toLocaleTimeString();
                const typeStr = {
                    'success': '[成功]',
                    'info': '[信息]',
                    'warning': '[警告]',
                    'danger': '[错误]'
                }[msg.type] || '[未知]';
                return `${timeStr} ${typeStr} ${msg.text}`;
            }).join('\n');
        });

        // 桌贴页面数据（使用Computed自动计算）
        const deskLabelPages = computed(() => {
            if (!arrangedRooms.value || arrangedRooms.value.length === 0) {
                return [];
            }

            const pages = [];
            const studentsPerPage = 15; // 固定每页15个学生

            arrangedRooms.value.forEach(room => {
                if (room.students && room.students.length > 0) {
                    const totalStudents = room.students.length;
                    const totalPages = Math.ceil(totalStudents / studentsPerPage);
                    
                    for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
                        const startIndex = pageIndex * studentsPerPage;
                        const endIndex = Math.min(startIndex + studentsPerPage, totalStudents);
                        const pageStudents = room.students.slice(startIndex, endIndex);
                        
                        pages.push({
                            room: room,
                            pageIndex: pageIndex,
                            totalPages: totalPages,
                            students: pageStudents,
                            startIndex: startIndex,
                            endIndex: endIndex
                        });
                    }
                }
            });

            return pages;
        });

        return {
            examName,
            emptyPng,
            fileSelect,
            selectFile,
            fileName,
            logTextarea,
            picFolderSelect,
            selectPicFolder,
            picFolderName,
            picFileList,
            handlePicFolderChange,
            matchPhotosToStudents,
            picShown,
            switchPicShown,
            examNotice,
            messages,
            messagesText,
            addMessage,
            arrangeConfig,
            tab,
            switchTab,
            handleFileChange,
            studentsRaw,
            roomsRaw,
            arrangedRooms,
            random,
            arrangeStuIntoRooms,
            getRoomRows,
            chunkStudents,
            sameStudentGroupList,
            deskLabelPages,
            // 准考证排序相关
            admissionTicketSortKey,
            admissionTicketSecondarySortKey,
            sortedSameStudentGroupList,
            resetAdmissionTicketSort,
        };
    }
});
let vmm = vm.mount("#exam-rooms-arrange");
window.examRoomArrangeApp = vmm;


