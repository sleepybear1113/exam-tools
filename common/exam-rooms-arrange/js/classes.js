class Student {
    constructor(id, name, ksh, typeId, typeName, roomId, roomName, seatNo, idCard, schoolId, schoolName, time, className) {
        this.id = id;
        this.name = name;
        this.ksh = ksh;
        this.typeId = typeId;
        this.typeName = typeName;
        this.roomId = roomId;
        this.roomName = roomName;
        this.seatNo = seatNo;
        this.idCard = idCard;
        this.schoolId = schoolId;
        this.schoolName = schoolName;
        this.time = time;
        this.className = className; // 班级

        this.photo = "";

        /**
         * 性别，true为男，false为女，null为未知
         * @type {null}
         */
        this.sex = null;
        if (idCard && idCard.length === 18) {
            const genderCode = parseInt(idCard.charAt(16));
            this.sex = (genderCode % 2 !== 0);
        }
    }

    /**
     * 校验学生整个信息的完整性
     * @returns {boolean}
     */
    validate() {
        return !(!this.ksh || !this.name || !this.typeId || !this.idCard);
    }

    copy() {
        const newStudent = new Student(this.id, this.name, this.ksh, this.typeId, this.typeName, this.roomId, this.roomName, this.seatNo, this.idCard, this.schoolId, this.schoolName, this.time, this.className);
        newStudent.photo = this.photo; // 复制照片属性
        return newStudent;
    }
}

class SameStudentGroup {
    constructor(id, className, ksh, name, idCard) {
        this.id = id;
        this.className = className;
        this.ksh = ksh;
        this.name = name;
        this.idCard = idCard;

        this.students = []; // 学生列表
        this.rooms = [];
    }

    key() {
        return this.ksh + "-" + this.name + "-" + this.idCard;
    }
}

class Room {

    constructor(id, schoolId, schoolName, typeId, typeName, time, roomId, roomName, count, capacity, remark) {
        this.id = id;
        this.schoolId = schoolId;
        this.schoolName = schoolName;
        this.typeId = typeId;
        this.typeName = typeName;
        this.time = time;
        this.roomId = roomId;
        this.roomName = roomName;
        this.count = count;
        this.capacity = capacity; // 座位数
        this.remark = remark; // 备注

        this.students = []; // 安排的学生列表
    }

    validate() {
        return !(!this.typeId || !this.time || !this.count || !this.capacity);
    }

    allEmptyProps() {
        let notEmptyCount = 0;
        for (const key in this) {
            if (key === 'id' || key === 'students') {
                continue;
            }
            if (this.hasOwnProperty(key)) {
                if (this[key]) {
                    notEmptyCount++;
                }
            }
        }
        return notEmptyCount <= 0;
    }

    copy() {
        return new Room(null, this.schoolId, this.schoolName, this.typeId, this.typeName, this.time, this.roomId, this.roomName, this.count, this.capacity, this.remark);
    }
}

class Message {
    constructor(msgTime, type, text) {
        this.type = type; // success, info, warning, danger
        this.text = text;
        this.msgTime = msgTime;
    }
}
