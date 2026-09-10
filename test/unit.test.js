'use strict';
/* 智慧树课后题记录器 v1 —— 纯函数单元测试（Node 直接跑：node test/unit.test.js） */
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const Z = require('../zhihuishu-recorder.user.js');

let passed = 0;
function ok(name, fn) {
  try { fn(); passed++; console.log('  ✓ ' + name); }
  catch (e) { console.error('  ✗ ' + name + '\n    ' + e.message); process.exitCode = 1; }
}

console.log('== normalizeStem / stemKey ==');
ok('去掉开头题号', () => assert.strictEqual(Z.normalizeStem('1. 中国特色社会主义的本质特征是什么？'), '中国特色社会主义的本质特征是什么？'));
ok('去掉全角题号', () => assert.strictEqual(Z.normalizeStem('12、以下哪项正确'), '以下哪项正确'));
ok('折叠空白', () => assert.strictEqual(Z.normalizeStem('  a\t b  c '), 'a b c'));
ok('同题干 key 稳定', () => assert.strictEqual(Z.stemKey('1. 同一道题'), Z.stemKey('  同一道题 ')));

console.log('== mergeQuestions 去重 ==');
const base = [{
  id: 'k1', stem: '中国的首都是哪座城市？', type: 'single',
  options: [{ letter: 'A', text: '北京' }, { letter: 'B', text: '上海' }], answer: 'A.北京', chapter: '第一章'
}];
const r1 = Z.mergeQuestions(base, [{
  stem: '中国的首都是哪座城市？', type: 'single',
  options: [{ letter: 'A', text: '北京' }, { letter: 'B', text: '上海' }], answer: 'A.北京', chapter: '第一章'
}]);
ok('同题干同选项 → 跳过', () => {
  assert.strictEqual(r1.list.length, 1);
  assert.strictEqual(r1.skipped, 1);
  assert.strictEqual(r1.added.length, 0);
});
const r2 = Z.mergeQuestions(base, [{
  stem: '中国的首都是哪座城市？', type: 'single',
  options: [{ letter: 'A', text: '上海' }, { letter: 'B', text: '北京' }], answer: 'B.北京', chapter: '第一章'
}]);
ok('同题干选项乱序 → 变体保留', () => {
  assert.strictEqual(r2.list.length, 2);
  assert.strictEqual(r2.variants.length, 1);
  assert.strictEqual(r2.skipped, 0);
});
const r3 = Z.mergeQuestions(base, [{ stem: '完全不同的一道新题', type: 'judge', options: [], answer: '正确', chapter: '第二章' }]);
ok('新题 → 新增', () => {
  assert.strictEqual(r3.list.length, 2);
  assert.strictEqual(r3.added.length, 1);
});

console.log('== zipStore ==');
const zip = Z.zipStore([
  { name: 'hello.txt', data: Z.utf8('hello 智慧树') },
  { name: 'word/document.xml', data: Z.utf8('<w:document/>') }
]);
ok('zip 魔数 PK', () => {
  assert.strictEqual(String.fromCharCode(zip[0], zip[1]), 'PK');
});
ok('zip 内含文件名与 EOCD', () => {
  const s = Buffer.from(zip).toString('latin1');
  assert.ok(s.indexOf('hello.txt') >= 0);
  assert.ok(s.indexOf('word/document.xml') >= 0);
  assert.ok(s.indexOf('\x50\x4b\x05\x06') >= 0);
});

console.log('== buildDocx ==');
const docx = Z.buildDocx({
  course: '马克思主义基本原理',
  sections: [
    { name: '第一章 绪论', questions: [
      { stem: '马克思主义的三个组成部分是？', type: 'multi',
        options: [{ letter: 'A', text: '哲学' }, { letter: 'B', text: '政治经济学' }, { letter: 'C', text: '科学社会主义' }],
        answer: 'A.哲学；B.政治经济学；C.科学社会主义' },
      { stem: '物质决定意识。', type: 'judge', options: [], answer: '正确' }
    ] }
  ]
});
ok('docx 魔数 PK', () => assert.strictEqual(String.fromCharCode(docx[0], docx[1]), 'PK'));
const latin = Buffer.from(docx).toString('latin1');
ok('docx 含必需部件名', () => {
  ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml'].forEach(n => assert.ok(latin.indexOf(n) >= 0, '缺少部件: ' + n));
});

/* 落盘，供 PowerShell Expand-Archive 做端到端解压校验 */
const out = path.join(__dirname, 'out-test.docx');
fs.writeFileSync(out, Buffer.from(docx));
console.log('  ✓ 已写入 ' + out + '（' + docx.length + ' bytes，待 Expand-Archive 校验）');

ok('去掉“第N题”题号', () => assert.strictEqual(Z.normalizeStem('第1题 以下哪项属于马克思主义？'), '以下哪项属于马克思主义？'));
ok('去掉“第1题.”题号', () => assert.strictEqual(Z.normalizeStem('第12题. 二次函数定义'), '二次函数定义'));

console.log('== coursesToDoc 跨章分组 ==');
const grouped = Z.coursesToDoc({
  name: '测试课',
  questions: [
    { stem: 'q1', chapter: '第一章' },
    { stem: 'q2', chapter: '第二章' },
    { stem: 'q3', chapter: '第一章' }
  ]
});
ok('按章节分组且保序', () => {
  assert.strictEqual(grouped.sections.length, 2);
  assert.strictEqual(grouped.sections[0].name, '第一章');
  assert.strictEqual(grouped.sections[0].questions.length, 2);
  assert.strictEqual(grouped.sections[1].name, '第二章');
});
ok('全局去重入库可被 Word 正确分组', () => {
  const m = Z.mergeQuestions([], [
    { stem: '第1题 三要素？', type: 'single', options: [], answer: 'A', chapter: '第一章' },
    { stem: '三要素？', type: 'single', options: [], answer: 'A', chapter: '第二章' }
  ]);
  assert.strictEqual(m.list.length, 1, '同题跨章应只存 1 条');
  assert.strictEqual(Z.coursesToDoc({ name: 'x', questions: m.list }).sections[0].name, '第一章');
});

console.log('== parseLines（真实界面样张） ==');
const p1 = Z.parseLines('单选\n题\n1. Odontalgia is the pain in the ______. ( )\nA. teeth\nB. stomach\nC. liver\nD. abdomen\n回答正确✓\n参考答案：A\n答案解析：无');
ok('单选：题干去掉标签与题号', () => assert.strictEqual(p1.stem, 'Odontalgia is the pain in the ______. ( )'));
ok('单选：选项4个且识别字母', () => assert.strictEqual(p1.options.length, 4) && assert.strictEqual(p1.options[0].letter, 'A') && assert.strictEqual(p1.options[0].text, 'teeth'));
ok('单选：题型 single', () => assert.strictEqual(Z.typeOf(p1.stem, p1.options, 1, 0), 'single'));

const p2 = Z.parseLines('判断\n题\n2. Jennie complained of painful urination. The medical term for this is hematuria. ( )\n对\n错\n回答正确✓\n参考答案：错\n答案解析：无');
ok('判断：裸“对/错”也能识别为选项', () => assert.strictEqual(p2.options.length, 2) && assert.strictEqual(p2.options[0].text, '对') && assert.strictEqual(p2.options[1].text, '错'));
ok('判断：题干干净', () => assert.strictEqual(p2.stem, 'Jennie complained of painful urination. The medical term for this is hematuria. ( )'));
ok('判断：题型 judge', () => assert.strictEqual(Z.typeOf(p2.stem, p2.options, 2, 0), 'judge'));

/* 选项内容读取（v1.8.2 修复：真机上选项字母与文字常分行/无标点） */
console.log('== 选项内容读取（分行 / 无标点 / 一行多个） ==');
const o1 = Z.parseLines('单选\n题\n1. 线性表的地址（ ）\nA\n连续\nB\n不连续\nC\n部分连续\nD\n均有可能\n回答正确✓\n参考答案：A');
ok('字母一行、文字下一行 → 4 个选项都读到文字', () => {
  assert.strictEqual(o1.stem, '线性表的地址（ ）');
  assert.strictEqual(o1.options.length, 4);
  assert.strictEqual(o1.options.map(o => o.letter).join(''), 'ABCD');
  assert.strictEqual(o1.options.map(o => o.text).join('/'), '连续/不连续/部分连续/均有可能');
});

const o2 = Z.parseLines('1. 栈的特点（ ）\nA 后进先出\nB 先进先出\nC 随机存取\nD 顺序存取\n答案：A');
ok('字母+文字（无标点）→ 选项文字读到', () => {
  assert.strictEqual(o2.stem, '栈的特点（ ）');
  assert.strictEqual(o2.options.length, 4);
  assert.strictEqual(o2.options.map(o => o.text).join('/'), '后进先出/先进先出/随机存取/顺序存取');
});

const o3 = Z.parseLines('1. 队列的特点（ ）\nA. 先进先出    B. 后进先出    C. 随机    D. 顺序\n答案：A');
ok('一行塞多个选项 → 自动拆开', () => {
  assert.strictEqual(o3.stem, '队列的特点（ ）');
  assert.strictEqual(o3.options.length, 4);
  assert.strictEqual(o3.options.map(o => o.text).join('/'), '先进先出/后进先出/随机/顺序');
});

const o4 = Z.parseLines('1. A study shows the sky is blue.（ ）\n对\n错\n参考答案：对');
ok('题干以“A ”开头（英文）→ 不被误认成选项', () => {
  assert.strictEqual(o4.stem, 'A study shows the sky is blue.（ ）');
  assert.strictEqual(o4.options.length, 2);
  assert.strictEqual(o4.options.map(o => o.text).join('/'), '对/错');
});

/* 字母块 / 文字块 被拆开时的配对（v1.8.4：真机样张） */
console.log('== 字母块 / 文字块 配对 ==');
const o5 = Z.parseLines('11. The combining form “typhl/o” is used to denote ______. （ ） cecum appendix colon rectum\nA.\nB.\nC.\nD.\n参考答案：A');
ok('题干行尾夹带选项文字 + 字母块 → 一一配对，题干不含选项词', () => {
  assert.strictEqual(o5.stem, 'The combining form “typhl/o” is used to denote ______. （ ）');
  assert.strictEqual(o5.options.map(o => o.letter + '.' + o.text).join(' | '), 'A.cecum | B.appendix | C.colon | D.rectum');
});

const o6 = Z.parseLines('13. _____________ refers to the localized death of living cells. （ ） Necrosis Strangulation Sclerosis Emesis\nA.\nB.\nC.\nD.\n参考答案：A');
ok('同上的另一道真题 → 4 个选项都对上', () => {
  assert.strictEqual(o6.stem, '_____________ refers to the localized death of living cells. （ ）');
  assert.strictEqual(o6.options.map(o => o.text).join('/'), 'Necrosis/Strangulation/Sclerosis/Emesis');
});

const o7 = Z.parseLines('1. 患者的主诉是（ ）\nA.\nB.\nC.\nD.\n头痛\n头晕\n恶心\n呕吐\n参考答案：B');
ok('字母块在前、文字在后面的 4 行 → 按顺序配对', () => {
  assert.strictEqual(o7.stem, '患者的主诉是（ ）');
  assert.strictEqual(o7.options.map(o => o.letter + '.' + o.text).join(' | '), 'A.头痛 | B.头晕 | C.恶心 | D.呕吐');
});

const o8 = Z.parseLines('1. 下列哪项正确（ ）\nA. 甲\nB. 乙\nC. 丙\nD. 丁\n参考答案：C');
ok('常规“A. 文字”不受预处理影响', () => {
  assert.strictEqual(o8.options.map(o => o.text).join('/'), '甲/乙/丙/丁');
});

console.log(passed + ' 项通过');
if (process.exitCode) { console.error('存在失败用例'); }
