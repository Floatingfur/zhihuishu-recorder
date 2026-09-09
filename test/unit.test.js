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

console.log(passed + ' 项通过');
if (process.exitCode) { console.error('存在失败用例'); }
