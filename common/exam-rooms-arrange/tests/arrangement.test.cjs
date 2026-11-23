const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require('../../../lib/sheetjs/0.20.3/xlsx.full.min.js');

const root = path.join(__dirname, '..');

function createApp() {
    const revoked = [];
    let nextUrl = 0;
    const context = vm.createContext({
        Vue: {
            ref: value => ({value}),
            computed: get => ({get value() { return get(); }}),
            nextTick: callback => callback(),
            createApp: options => ({mount: () => options.setup()}),
        },
        readExcel: async file => structuredClone(file.rows),
        URL: {
            createObjectURL: () => `blob:test-${++nextUrl}`,
            revokeObjectURL: url => revoked.push(url),
        },
        confirm: () => true,
        window: {},
    });
    vm.runInContext(fs.readFileSync(path.join(root, 'js', 'classes.js'), 'utf8'), context);
    const source = fs.readFileSync(path.join(root, 'js', 'app.js'), 'utf8')
        .replace('import {readExcel} from "./excel-utils.js";', '');
    vm.runInContext(source, context);
    return {app: context.window.examRoomArrangeApp, revoked};
}

function rows(students, rooms) {
    const header = [
        '姓名', '考生号', '班级', '类别代码', '身份证号', '考点代码', '!分割列!',
        '考点代码', '考点名称', '类别代码', '类别名称', '考试时间', '试场号',
        '容纳人数', '该考试时间下的试场数', '备注', '准考证考生必读',
    ];
    const data = Array.from({length: Math.max(students.length, rooms.length)}, (_, i) => [
        ...(students[i] || Array(6).fill('')), '',
        ...(rooms[i] || Array(9).fill('')), '',
    ]);
    return [['考试'], header, ...data];
}

async function load(app, data) {
    await app.handleFileChange({target: {files: [{name: 'test.xlsx', rows: data}], value: 'test.xlsx'}});
}

function messages(app) {
    return app.messages.value.map(message => message.text).join('\n');
}

test('安排无照片考生并生成核对单、准考证和桌贴数据', async () => {
    const {app} = createApp();
    const students = Array.from({length: 31}, (_, i) => [
        `考生${i}`, 1000 + i, i + 1, 2, `3301012000010100${String(i).padStart(2, '0')}`, 10,
    ]);
    await load(app, rows(students, [[10, '考点', 2, '类别', '上午', '05-06', 31, 2, '']]));
    app.arrangeStuIntoRooms(false);
    const room = app.arrangedRooms.value[0];
    assert.equal(room.roomId, '05');
    assert.equal(room.roomName, '05');
    assert.equal(room.students.length, 31);
    assert.equal(room.students[0].ksh, '1000');
    assert.equal(room.students[0].seatNo, 1);
    assert.equal(room.students[0].photo, '');
    assert.deepEqual(Array.from(app.photoChecklistPages.value, page => page.students.length), [30, 1]);
    assert.equal(app.sortedSameStudentGroupList.value.length, 31);
    assert.equal(app.deskLabelPages.value.length, 2);
    assert.match(messages(app), /未选择照片，预览将显示空白照片/);
    app.switchPicShown(false);
    assert.equal(app.picShown.value, false);
});

test('随机编排可匹配照片、重新编排释放旧照片并保留试场编号', async () => {
    const {app, revoked} = createApp();
    await load(app, rows([
        ['甲', '101', '1', 'A', '330101200001010011', ''],
        ['乙', '102', '1', 'A', '330101200001010022', ''],
    ], [['', '考点', 'A', '类别', '上午', '08', 2, 1, '']]));
    app.handlePicFolderChange({target: {
        files: [{name: '101.jpg', webkitRelativePath: 'photos/101.jpg'}], value: 'photos',
    }});
    app.arrangeStuIntoRooms(true);
    assert.equal(app.arrangedRooms.value[0].roomName, '08');
    assert.equal(app.arrangedRooms.value[0].students.filter(student => student.photo).length, 1);
    const oldPhoto = app.arrangedRooms.value[0].students.find(student => student.photo).photo;
    app.arrangeStuIntoRooms(false);
    assert.ok(revoked.includes(oldPhoto));
    assert.equal(app.arrangedRooms.value[0].students[0].photo, 'blob:test-2');
    app.handlePicFolderChange({target: {
        files: [{name: 'readme.txt', webkitRelativePath: 'empty/readme.txt'}], value: 'empty',
    }});
    assert.equal(app.arrangedRooms.value[0].students[0].photo, '');
    assert.ok(revoked.includes('blob:test-2'));
});

test('不完整或无效输入不会复用旧编排结果，错误会写入日志', async () => {
    const {app} = createApp();
    app.arrangeStuIntoRooms(false);
    assert.match(messages(app), /请先导入/);
    const student = [['甲', '101', '1', 'A', '330101200001010011', '']];
    await load(app, rows(student, [['', '考点', 'A', '类别', '上午', '', 1, 1, '']]));
    app.arrangeStuIntoRooms(false);
    assert.equal(app.arrangedRooms.value.length, 1);
    await load(app, [['考试']]);
    assert.equal(app.arrangedRooms.value.length, 0);
    app.arrangeStuIntoRooms(false);
    assert.match(messages(app), /Excel缺少表头/);
    await load(app, rows(student, [
        ['', '考点', 'A', '类别', '上午', '01,01', 1, 1, ''],
        ['', '考点', 'A', '类别', '下午', '', '无效', 1, ''],
    ]));
    app.arrangeStuIntoRooms(false);
    assert.equal(app.arrangedRooms.value.length, 0);
    assert.match(messages(app), /试场号为空、重复/);
    assert.match(messages(app), /试场数和容纳人数必须为正整数/);
});

test('仓库自带的完整样例可以编排并生成所有预览数据', async () => {
    const {app} = createApp();
    const sample = fs.readdirSync(path.join(root, 'files')).find(name => name.includes('随机数据填充'));
    const workbook = XLSX.read(fs.readFileSync(path.join(root, 'files', sample)), {type: 'buffer'});
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1});
    await load(app, data);
    app.arrangeStuIntoRooms(false);
    assert.ok(app.arrangedRooms.value.length > 0, messages(app));
    assert.equal(
        app.arrangedRooms.value.reduce((sum, room) => sum + room.students.length, 0),
        app.studentsRaw.value.length,
        messages(app),
    );
    assert.ok(app.photoChecklistPages.value.length > 0);
    assert.ok(app.sortedSameStudentGroupList.value.length > 0);
    assert.ok(app.deskLabelPages.value.length > 0);
    assert.ok(app.arrangedRooms.value.every(room => room.roomName === room.roomId));
});
