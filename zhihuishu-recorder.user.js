// ==UserScript==
// @name         智慧树课后题记录器 v1
// @namespace    https://dsh.local/zhihuishu-recorder
// @version      1.0.0
// @description  在你做完智慧树章节测验并进入「本次成绩/查看答案解析」页后，点「开始记录」把本章题目+正确答案存入本地题库（跨章节累计、按题干去重、选项乱序变体保留），可另存为 Word(.docx)。纯本地运行，不联网、不自动答题。
// @author       you
// @match        https://*.zhihuishu.com/*
// @match        http://*.zhihuishu.com/*
// @run-at       document-idle
// @all-frames   true
// @grant        GM_setClipboard
// @grant        GM_addStyle
// @grant        GM_getValue
// @grant        GM_setValue
// ==/UserScript==

/*
 * =====================================================================
 *  智慧树课后题记录器 v1 —— 单文件自包含实现
 *  - 纯函数部分（存储/去重/docx/zip）可在 Node 中 require 做单元测试
 *  - 浏览器环境下自动挂 UI（右下角浮动面板）
 * =====================================================================
 */
'use strict';
var ZHR = (function () {
  var VERSION = '1.0.0';

  /* ---------------- 基础工具（纯函数） ---------------- */

  function utf8(s) { return new TextEncoder().encode(String(s)); }

  function djb2(str) {
    var h = 5381, i;
    for (i = 0; i < str.length; i++) { h = ((h << 5) + h + str.charCodeAt(i)) >>> 0; }
    return h.toString(36);
  }

  function normWs(s) { return String(s || '').replace(/\s+/g, ' ').trim(); }

  /* 题干规范化：去首部题号、折叠空白，用于去重键 */
  function normalizeStem(stem) {
    var s = normWs(stem);
    s = s.replace(/^第\s*[0-9一二三四五六七八九十百零]+\s*题[.、．)）:：]?\s*/, '');
    s = s.replace(/^\d+[.、．)）:：]\s*/, '');
    return s;
  }

  function xmlEscape(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  /* ---------------- 最小 ZIP(STORE) 写入器（纯函数） ---------------- */

  var CRC_TABLE = null;
  function makeCrcTable() {
    var t = new Uint32Array(256), c, n, k;
    for (n = 0; n < 256; n++) {
      c = n;
      for (k = 0; k < 8; k++) { c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); }
      t[n] = c >>> 0;
    }
    return t;
  }
  function crc32(bytes) {
    if (!CRC_TABLE) CRC_TABLE = makeCrcTable();
    var c = 0xFFFFFFFF, i;
    for (i = 0; i < bytes.length; i++) { c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8); }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /* entries: [{name:String, data:Uint8Array}]，返回完整 zip 的 Uint8Array */
  function zipStore(entries) {
    var locals = [], centrals = [];
    var offset = 0, total = 0;
    entries.forEach(function (e) {
      var nameB = utf8(e.name);
      var crc = crc32(e.data);
      var locLen = 30 + nameB.length;
      var cenLen = 46 + nameB.length;
      locals.push({ nameB: nameB, data: e.data, crc: crc, off: offset });
      centrals.push({ nameB: nameB, dataLen: e.data.length, crc: crc, off: offset, locLen: locLen, cenLen: cenLen });
      offset += locLen + e.data.length;
      total += locLen + e.data.length;
    });
    var cenStart = offset;
    centrals.forEach(function (c) { total += c.cenLen; });
    total += 22;

    var out = new Uint8Array(total);
    var dv = new DataView(out.buffer);

    function putStr(p, s) { for (var i = 0; i < s.length; i++) { out[p + i] = s.charCodeAt(i) & 0xFF; } return p + s.length; }
    function putBytes(p, b) { out.set(b, p); return p + b.length; }

    var p = 0;
    locals.forEach(function (l) {
      dv.setUint32(p, 0x04034b50, true); p += 4;   // sig
      dv.setUint16(p, 20, true); p += 2;           // version needed
      dv.setUint16(p, 0, true); p += 2;            // flags
      dv.setUint16(p, 0, true); p += 2;            // method = store
      dv.setUint16(p, 0, true); p += 2;            // mod time
      dv.setUint16(p, 0x21, true); p += 2;         // mod date (1980-01-01)
      dv.setUint32(p, l.crc, true); p += 4;
      dv.setUint32(p, l.data.length, true); p += 4; // comp size
      dv.setUint32(p, l.data.length, true); p += 4; // uncomp size
      dv.setUint16(p, l.nameB.length, true); p += 2;
      dv.setUint16(p, 0, true); p += 2;            // extra len
      p = putBytes(p, l.nameB);
      p = putBytes(p, l.data);
    });

    centrals.forEach(function (c) {
      dv.setUint32(p, 0x02014b50, true); p += 4;   // sig
      dv.setUint16(p, 20, true); p += 2;           // version made by
      dv.setUint16(p, 20, true); p += 2;           // version needed
      dv.setUint16(p, 0, true); p += 2;            // flags
      dv.setUint16(p, 0, true); p += 2;            // method
      dv.setUint16(p, 0, true); p += 2;            // time
      dv.setUint16(p, 0x21, true); p += 2;         // date
      dv.setUint32(p, c.crc, true); p += 4;
      dv.setUint32(p, c.dataLen, true); p += 4;   // comp size
      dv.setUint32(p, c.dataLen, true); p += 4;   // uncomp size
      dv.setUint16(p, c.nameB.length, true); p += 2;
      dv.setUint16(p, 0, true); p += 2;            // extra
      dv.setUint16(p, 0, true); p += 2;            // comment
      dv.setUint16(p, 0, true); p += 2;            // disk
      dv.setUint16(p, 0, true); p += 2;            // internal attr
      dv.setUint32(p, 0, true); p += 4;            // external attr
      dv.setUint32(p, c.off, true); p += 4;        // local header offset
      p = putBytes(p, c.nameB);
    });

    dv.setUint32(p, 0x06054b50, true); p += 4;     // EOCD sig
    dv.setUint16(p, 0, true); p += 2;              // disk
    dv.setUint16(p, 0, true); p += 2;
    dv.setUint16(p, centrals.length, true); p += 2;
    dv.setUint16(p, centrals.length, true); p += 2;
    dv.setUint32(p, total - 22 - cenStart, true); p += 4; // central size
    dv.setUint32(p, cenStart, true); p += 4;       // central offset
    dv.setUint16(p, 0, true); p += 2;              // comment len
    return out;
  }

  /* ---------------- .docx 生成器（纯函数，OOXML 最小子集） ---------------- */

  function wRun(text, opts) {
    opts = opts || {};
    var rpr = '';
    if (opts.bold) rpr += '<w:b/>';
    if (opts.size) rpr += '<w:sz w:val="' + opts.size + '"/><w:szCs w:val="' + opts.size + '"/>';
    if (opts.color) rpr += '<w:color w:val="' + opts.color + '"/>';
    return '<w:r>' + (rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + xmlEscape(text) + '</w:t></w:r>';
  }

  function wPara(runsXml, spacingAfter) {
    var ppr = spacingAfter ? '<w:pPr><w:spacing w:after="' + spacingAfter + '"/></w:pPr>' : '';
    return '<w:p>' + ppr + runsXml + '</w:p>';
  }

  function typeLabel(t) {
    return { single: '单选', multi: '多选', judge: '判断', fill: '填空', other: '其他' }[t] || (t || '');
  }

  function optionsLine(q) {
    var opts = q.options || [];
    var parts = [];
    for (var i = 0; i < opts.length; i++) {
      var o = opts[i];
      parts.push((o.letter || String.fromCharCode(65 + i)) + '. ' + o.text);
    }
    return parts.join('　');
  }

  /* doc: { course: String, sections: [{name:String, questions:[{stem,type,options,answer}]}] } */
  function buildDocx(doc) {
    var paraXml = '';
    paraXml += wPara(wRun(doc.course || '智慧树课程题库', { bold: true, size: 36 }), 240);
    var total = 0;
    (doc.sections || []).forEach(function (sec) {
      paraXml += wPara(wRun(sec.name || '未命名章节', { bold: true, size: 28 }), 200);
      (sec.questions || []).forEach(function (q, idx) {
        total++;
        var qn = (idx + 1) + '. ';
        var t = typeLabel(q.type);
        var body = qn + (t ? '【' + t + '】' : '') + (q.stem || '');
        paraXml += wPara(wRun(body, { size: 21 }), 40);
        var ol = optionsLine(q);
        if (ol) paraXml += wPara(wRun(ol, { size: 21 }), 40);
        paraXml += wPara(wRun('【答案】' + (q.answer || '未记录'), { bold: true, color: '1F6F43', size: 21 }), 160);
      });
    });
    if (!total) paraXml += wPara(wRun('（暂无题目记录）'), 0);

    var documentXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body>' + paraXml +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>' +
      '</w:body></w:document>';

    var stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/>' +
      '<w:sz w:val="21"/><w:szCs w:val="21"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '</w:styles>';

    var contentTypesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '</Types>';

    var rootRelsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>';

    var docRelsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>';

    return zipStore([
      { name: '[Content_Types].xml', data: utf8(contentTypesXml) },
      { name: '_rels/.rels', data: utf8(rootRelsXml) },
      { name: 'word/document.xml', data: utf8(documentXml) },
      { name: 'word/_rels/document.xml.rels', data: utf8(docRelsXml) },
      { name: 'word/styles.xml', data: utf8(stylesXml) }
    ]);
  }

  /* ---------------- 题库合并/去重（纯函数） ----------------
   * 规则(Q12=a)：规范化题干相同 => 同一题；选项相同则跳过，选项不同按变体另存
   */
  function stemKey(stem) { return djb2(normalizeStem(stem)); }
  function optsHash(options) {
    var s = (options || []).map(function (o) { return (o.letter || '') + (o.text || ''); }).join('|');
    return djb2(normWs(s));
  }

  function mergeQuestions(existing, incoming) {
    var byKey = {};
    existing.forEach(function (q, i) {
      var k = stemKey(q.stem);
      if (!byKey[k]) byKey[k] = [];
      byKey[k].push(i);
    });
    var added = [], variants = [], skipped = 0;
    var list = existing.slice();
    incoming.forEach(function (inc) {
      var k = stemKey(inc.stem);
      var hits = byKey[k] || [];
      if (!hits.length) {
        var q = {
          id: k + '_' + list.length,
          stem: normalizeStem(inc.stem),
          type: inc.type || 'other',
          options: (inc.options || []).slice(),
          answer: inc.answer || '',
          chapter: inc.chapter || '',
          firstSeen: Date.now(),
          updatedAt: Date.now()
        };
        list.push(q);
        byKey[k] = [list.length - 1];
        added.push(q);
        return;
      }
      var oh = optsHash(inc.options);
      var matched = null;
      for (var i2 = 0; i2 < hits.length; i2++) {
        var idx = hits[i2];
        if (optsHash(list[idx].options) === oh) { matched = idx; break; }
      }
      if (matched !== null) {
        skipped++;
        list[matched].updatedAt = Date.now();
        if (inc.answer && inc.answer !== list[matched].answer) { list[matched].answer = inc.answer; }
        return;
      }
      var v = {
        id: k + '_v' + (list.length + variants.length),
        stem: normalizeStem(inc.stem),
        type: inc.type || 'other',
        options: (inc.options || []).slice(),
        answer: inc.answer || '',
        chapter: inc.chapter || '',
        variantOf: true,
        firstSeen: Date.now(),
        updatedAt: Date.now()
      };
      list.push(v);
      byKey[k].push(list.length - 1);
      variants.push(v);
    });
    return { list: list, added: added, variants: variants, skipped: skipped };
  }

  function coursesToDoc(course) {
    /* 全局题库 → 按章节分组（保持首次出现顺序），供 Word 排版 */
    var order = [];
    var byChapter = {};
    (course.questions || []).forEach(function (q) {
      var ch = q.chapter || '未命名章节';
      if (!(ch in byChapter)) { byChapter[ch] = []; order.push(ch); }
      byChapter[ch].push({ stem: q.stem, type: q.type, options: q.options, answer: q.answer });
    });
    return { course: course.name, sections: order.map(function (n) { return { name: n, questions: byChapter[n] }; }) };
  }

  /* ================================================================
   *  以下为浏览器端（DOM 识别 / UI），Node 环境不执行
   * ================================================================ */

  var OPT_RE = /^[（(]?([A-Ha-hＡ-Ｈａ-ｈ])[)）.．、:：]\s*/;
  /* 全角/半角拉丁字母归一（Ａ→A） */
  function normLetter(ch) {
    var code = ch.charCodeAt(0);
    if (code >= 0xFF21 && code <= 0xFF3A) return String.fromCharCode(code - 0xFEE0);
    if (code >= 0xFF41 && code <= 0xFF5A) return String.fromCharCode(code - 0xFEE0);
    return ch;
  }
  var ANSWER_LINE_RE = /(正确答案|参考答案|标准答案|正确答案是)/;
  var ANSWER_LINE_ANY_RE = /答案[:：]?\s*([^\n]{0,60})/;
  var CORRECT_CLS_RE = /\b(correct|right|answer|daan|green|succ|pass|true)\w*/i;
  var MARKED_WRONG_RE = /\b(wrong|incorrect|false|error)\w*/i;
  var CONTAINER_CLS_RE = /(question|timu|topic|subject|exam|test|answer|result|choice|item)/i;

  function splitLines(t) {
    return String(t || '').split(/\r?\n/).map(function (s) { return normWs(s); }).filter(Boolean);
  }

  function clsOf(el) {
    if (typeof el.className === 'string') return el.className;
    if (el.classList && el.classList.length) return Array.prototype.join.call(el.classList, ' ');
    return '';
  }

  function innerTextOf(el) { return (el.innerText || el.textContent || ''); }

  function hasAnswerMarker(el) {
    var t = innerTextOf(el);
    if (ANSWER_LINE_RE.test(t)) return true;
    if (CORRECT_CLS_RE.test(clsOf(el))) return true;
    var marked = el.querySelectorAll('[class]');
    for (var i = 0; i < marked.length && i < 200; i++) {
      if (CORRECT_CLS_RE.test(clsOf(marked[i]))) return true;
    }
    return false;
  }

  /* 找“整章回看/解析”页里的题目容器：class/id 命中关键词 且 含答案标记 */
  function findQuestionContainers(doc) {
    var all = doc.querySelectorAll('div,li,section,tr,fieldset');
    var cands = [];
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      var idc = el.id || '';
      var clsc = clsOf(el);
      if (!(CONTAINER_CLS_RE.test(idc) || CONTAINER_CLS_RE.test(clsc))) continue;
      var txt = innerTextOf(el);
      if (txt.length < 8 || txt.length > 20000) continue;
      if (!hasAnswerMarker(el)) continue;
      cands.push(el);
    }
    /* 去嵌套：候选里含另一个“可用候选”的不保留（取最内层题目块） */
    var usable = cands.filter(function (a) {
      for (var j = 0; j < cands.length; j++) {
        if (cands[j] === a) continue;
        if (a.contains(cands[j]) && cands[j] !== a) return false;
      }
      return true;
    });
    return usable;
  }

  /* 从一个题目容器的行文本里解析题干/选项/答案（行文本模型，兼容多数布局） */
  function parseLines(containerText) {
    var lines = splitLines(containerText);
    var stemLines = [], options = [], ansText = '', ansLineIdx = -1;
    for (var i = 0; i < lines.length; i++) {
      var L = lines[i];
      if (ANSWER_LINE_RE.test(L)) {
        if (/我的答案|您的答案/.test(L) && !/正确答案|参考答案|标准答案/.test(L)) continue;
        ansLineIdx = i;
        ansText = L;
        continue;
      }
      var m = L.match(OPT_RE);
      if (m) {
        options.push({ letter: normLetter(m[1]).toUpperCase(), text: L.replace(OPT_RE, '').trim() });
        continue;
      }
      if (ansLineIdx === -1) stemLines.push(L);
    }
    var stem = normalizeStem(stemLines.join(' '));
    return { stem: stem, options: options, answerLine: ansText };
  }

  function typeOf(stem, options, radios, checks) {
    var s = stem || '';
    if (/多选|多项/.test(s) || checks > 0) return 'multi';
    var isJudge = /判断/.test(s) ||
      (options.length === 2 && /^(正确|对)/.test(options[0].text) && /^(错误|错)/.test(options[1].text));
    if (isJudge) return 'judge';
    if (radios > 0 || options.length) return 'single';
    if (/填空|___|＿+|（\s*）|\(\s*\)/.test(s)) return 'fill';
    return 'other';
  }

  function answerFromLine(lineText, fallback) {
    if (!lineText) return fallback || '';
    var m = lineText.match(ANSWER_LINE_ANY_RE);
    var rest = m ? m[1] : lineText;
    rest = rest.replace(/^(是|为|：|:|:)/, '').trim();
    /* 只取“答案”后开头连续的一段字母（A、B、C / AB 及全角），避免吞入后文解析文字 */
    var lead = rest.match(/^[A-Ha-hＡ-Ｈａ-ｈ\s、，,和]+/);
    var letters = [];
    if (lead) {
      (lead[0] || '').split('').forEach(function (ch) {
        var a = normLetter(ch).toUpperCase();
        if (/[A-H]/.test(a) && letters.indexOf(a) < 0) letters.push(a);
      });
    }
    if (letters.length) return letters.join('');
    return normWs(rest).slice(0, 80);
  }

  /* 高亮(绿色/correct 类)标记的正确选项 → 字母 */
  function lettersFromMarkedOptions(container) {
    var found = [];
    var nodes = container.querySelectorAll('li,label,p,div,span');
    for (var i = 0; i < nodes.length && i < 500; i++) {
      var el = nodes[i];
      var c = clsOf(el);
      if (!c || !CORRECT_CLS_RE.test(c)) continue;
      if (MARKED_WRONG_RE.test(c)) continue;
      var t = normWs(innerTextOf(el));
      var m = t.match(OPT_RE);
      if (m) { var l = normLetter(m[1]).toUpperCase(); if (found.indexOf(l) < 0) found.push(l); }
    }
    return found.join('');
  }

  /* 单个容器的完整解析 */
  function parseQuestion(container, chapterName) {
    var hasInputs = container.querySelectorAll('input[type=radio]').length;
    var checkInputs = container.querySelectorAll('input[type=checkbox]').length;
    var raw = innerTextOf(container);
    var parsed = parseLines(raw);
    var type = typeOf(parsed.stem, parsed.options, hasInputs, checkInputs);
    var marked = lettersFromMarkedOptions(container);
    var answer = '';
    if (marked) answer = marked;
    else if (parsed.answerLine) answer = answerFromLine(parsed.answerLine, '');
    else if (type === 'judge' && parsed.options.length) {
      /* 判断：无文字答案时尝试从“√/×”类行取 */
      var lines = splitLines(raw);
      for (var i = 0; i < lines.length; i++) {
        if (/^(正确|对|√|是)\s*$/.test(lines[i])) { answer = '正确'; break; }
        if (/^(错误|错|×|否)\s*$/.test(lines[i])) { answer = '错误'; break; }
      }
    }
    /* 单选/多选：把字母答案还原为可读形式（字母 + 对应选项文字） */
    if (/^[A-H]+$/.test(answer)) {
      var parts = [];
      for (var j = 0; j < answer.length; j++) {
        var letter = answer.charAt(j);
        var op = null;
        for (var k = 0; k < parsed.options.length; k++) {
          if (parsed.options[k].letter === letter) { op = parsed.options[k]; break; }
        }
        parts.push(letter + (op ? '.' + op.text : ''));
      }
      answer = parts.join('；');
    }
    return {
      stem: parsed.stem,
      options: parsed.options,
      answer: answer || '',
      type: type,
      chapter: chapterName
    };
  }

  /* 课程名 / 章节名猜测 */
  function guessCourseName() {
    var title = document.title || '';
    var m = title.match(/(.{1,40}?)[-_—|]?(智慧树|知到)/);
    if (m && m[1]) return normWs(m[1].replace(/^[\s\-—_|:：]+/, ''));
    var heads = document.querySelectorAll('h1,h2,h3,.courseName,.course-name,.name,.title');
    for (var i = 0; i < heads.length && i < 30; i++) {
      var t = normWs(heads[i].innerText || '');
      if (t && t.length <= 40) return t;
    }
    return '';
  }

  function guessChapterName() {
    var re = /第\s*[0-9一二三四五六七八九十百零]+\s*[章节讲单元课]/;
    var els = document.querySelectorAll('h1,h2,h3,h4,.chapter,.section,.unit,li,span,div,b,strong,p');
    var best = '';
    for (var i = 0; i < els.length && i < 200; i++) {
      var el = els[i];
      if (el.closest && el.closest('#zhr-root')) continue;
      var t = normWs(el.innerText || '');
      if (!t || t.length > 60) continue;
      if (re.test(t)) {
        if (!best || t.length < best.length) best = t;
      }
    }
    return best;
  }

  /* ---------------- UI ---------------- */

  function toast(msg, isError) {
    var wrap = document.getElementById('zhr-root');
    if (!wrap) return;
    var t = document.createElement('div');
    t.textContent = msg;
    t.style.cssText = 'position:fixed;left:16px;bottom:16px;z-index:2147483647;background:' + (isError ? '#c0392b' : '#27ae60') +
      ';color:#fff;padding:10px 16px;border-radius:6px;font:13px/1.5 "Microsoft YaHei",sans-serif;box-shadow:0 4px 14px rgba(0,0,0,.3);max-width:70vw;word-break:break-all;transition:opacity .4s';
    wrap.appendChild(t);
    setTimeout(function () {
      t.style.opacity = '0';
      setTimeout(function () { if (t.parentNode) t.parentNode.removeChild(t); }, 450);
    }, 3800);
  }

  function showModal(title, bodyHtml, textareaValue) {
    var mask = document.createElement('div');
    mask.style.cssText = 'position:fixed;left:0;top:0;right:0;bottom:0;z-index:2147483647;background:rgba(0,0,0,.4);display:flex;align-items:center;justify-content:center';
    var panel = document.createElement('div');
    panel.style.cssText = 'width:min(92vw,860px);max-width:94%;height:min(85vh,680px);background:#fff;color:#222;border-radius:8px;box-shadow:0 8px 40px rgba(0,0,0,.4);display:flex;flex-direction:column;overflow:hidden;font:13px/1.5 "Microsoft YaHei",sans-serif';
    var head = document.createElement('div');
    head.style.cssText = 'padding:8px 14px;background:#f2f5fa;border-bottom:1px solid #ddd;font-weight:700';
    head.textContent = title;
    var body = document.createElement('div');
    body.style.cssText = 'flex:1;overflow:auto;padding:10px 14px';
    body.innerHTML = bodyHtml || '';
    var foot = document.createElement('div');
    foot.style.cssText = 'padding:8px 14px;background:#f2f5fa;border-top:1px solid #ddd;text-align:right';
    var bClose = document.createElement('button');
    bClose.textContent = '关闭';
    bClose.style.cssText = 'border:0;border-radius:6px;padding:6px 14px;cursor:pointer;background:#eef1f6;color:#333';
    foot.appendChild(bClose);
    if (textareaValue != null) {
      var ta = document.createElement('textarea');
      ta.readOnly = true;
      ta.value = textareaValue;
      ta.style.cssText = 'width:100%;height:100%;min-height:380px;box-sizing:border-box;font:11px/1.5 Consolas,monospace;resize:none;padding:6px';
      body.appendChild(ta);
      var bCopy = document.createElement('button');
      bCopy.textContent = '复制到剪贴板';
      bCopy.style.cssText = bClose.style.cssText + ';background:#2f6fed;color:#fff;margin-right:8px';
      foot.insertBefore(bCopy, foot.firstChild);
      bCopy.addEventListener('click', function () {
        try { GM_setClipboard(ta.value); bCopy.textContent = '已复制 ✓'; } catch (e) {
          ta.select(); try { document.execCommand('copy'); } catch (e2) { /* ignore */ }
          bCopy.textContent = '请手动 Ctrl+C';
        }
        setTimeout(function () { bCopy.textContent = '复制到剪贴板'; }, 1500);
      });
    }
    panel.appendChild(head); panel.appendChild(body); panel.appendChild(foot); mask.appendChild(panel);
    document.documentElement.appendChild(mask);
    function close() { if (mask.parentNode) mask.parentNode.removeChild(mask); }
    bClose.addEventListener('click', close);
    mask.addEventListener('click', function (e) { if (e.target === mask) close(); });
  }

  /* ---------------- 存储（GM_* 跨子域共享） ---------------- */

  var K_INDEX = 'zhr.v1.index';
  function loadCourseList() { return JSON.parse(GM_getValue(K_INDEX, 'null')) || []; }
  function courseKeyFor(name) { return 'zhr.v1.course.' + djb2(name || '未命名课程'); }
  function loadCourse(name) {
    var c = JSON.parse(GM_getValue(courseKeyFor(name), 'null'));
    if (!c) return null;
    if (c.chapters) {
      /* 旧版 {chapters:{章:[题]}} 摊平成全局题库（按 chapter 字段归章） */
      var qs = [];
      Object.keys(c.chapters).forEach(function (k) {
        (c.chapters[k].questions || []).forEach(function (q) {
          if (!q.chapter) q.chapter = k;
          qs.push(q);
        });
      });
      c = { name: c.name, questions: qs };
    }
    if (!c.questions) c.questions = [];
    return c;
  }
  function saveCourse(name, data) {
    GM_setValue(courseKeyFor(name), JSON.stringify(data));
    var list = loadCourseList();
    var hit = null;
    for (var i = 0; i < list.length; i++) if (list[i].name === name) { hit = list[i]; break; }
    if (!hit) { hit = { name: name }; list.push(hit); }
    hit.updatedAt = Date.now();
    GM_setValue(K_INDEX, JSON.stringify(list));
  }

  /* ---------------- 核心动作 ---------------- */

  function currentCourseGuess() {
    var inp = document.getElementById('zhr-course-input');
    return normWs(inp ? inp.value : '') || guessCourseName();
  }

  function doRecord() {
    var containers = findQuestionContainers(document);
    var chapterGuess = guessChapterName();
    var chapterInput = document.getElementById('zhr-chapter-input');
    var courseName = currentCourseGuess();
    var chapterName = normWs(chapterInput ? chapterInput.value : '') || chapterGuess || '未命名章节';

    if (!containers.length) {
      toast('未识别到带答案的题目块。若本页确为“成绩/查看解析”页，请点「诊断」把快照发我，或点「重新识别」后重试。', true);
      return { ok: false, reason: 'no-containers' };
    }

    var incoming = [];
    var failed = 0;
    for (var i = 0; i < containers.length; i++) {
      try {
        var q = parseQuestion(containers[i], chapterName);
        if (q.stem) incoming.push(q); else failed++;
      } catch (err) { failed++; }
    }
    if (!incoming.length) {
      toast('识别到容器但未能解析出题目（请点「诊断」把快照发我）。', true);
      return { ok: false, reason: 'parse-failed' };
    }

    var course = loadCourse(courseName) || { name: courseName, questions: [] };
    var res = mergeQuestions(course.questions, incoming);
    course.questions = res.list;
    saveCourse(courseName, course);
    var totalAll = course.questions.length;

    toast('已记录：' + courseName + ' › ' + chapterName + '\n新增 ' + res.added.length + ' 题，变体 ' + res.variants.length + ' 题，重复跳过 ' + res.skipped + ' 题，解析失败 ' + failed + ' 题；该课累计 ' + totalAll + ' 题');
    return { ok: true, added: res.added.length, variants: res.variants.length, skipped: res.skipped, failed: failed, totalAll: totalAll };
  }

  function diagnosticText() {
    var out = [];
    out.push('== v1 诊断快照 ==');
    out.push('version: ' + VERSION);
    out.push('url: ' + location.href);
    out.push('title: ' + normWs(document.title));
    out.push('courseGuess: ' + guessCourseName());
    out.push('chapterGuess: ' + guessChapterName());
    var conts = findQuestionContainers(document);
    out.push('containers=' + conts.length);
    for (var i = 0; i < conts.length && i < 30; i++) {
      var c = conts[i];
      out.push('  [' + i + '] <' + c.tagName.toLowerCase() + '> class="' + clsOf(c) + '" id="' + c.id + '"');
      var parsed = { err: '' };
      try { parsed = parseQuestion(c, ''); } catch (e) { parsed = { err: e.message }; }
      if (parsed.err) out.push('      parseErr: ' + parsed.err);
      else out.push('      stem=' + (parsed.stem || '').slice(0, 60) + ' | type=' + parsed.type + ' | opts=' + parsed.options.length + ' | ans=' + (parsed.answer || '').slice(0, 40));
    }
    if (!conts.length) {
      var txt = normWs(document.body.innerText);
      out.push('-- body 文本预览（前 1200 字，用于人工定位）--');
      out.push(txt.slice(0, 1200));
      var tags = document.querySelectorAll('[class*="question" i],[class*="timu"],[class*="subject" i],[class*="item"],[id*="question" i],[id*="timu"]');
      out.push('keywordElems=' + tags.length + '（下列为前 25 个的 class）');
      for (var k = 0; k < tags.length && k < 25; k++) {
        var el = tags[k];
        out.push('  <' + el.tagName.toLowerCase() + '> class="' + clsOf(el).slice(0, 60) + '" text=' + normWs(el.innerText).slice(0, 30));
      }
    }
    return out.join('\n');
  }

  function doExportWord() {
    var courseName = currentCourseGuess();
    if (!courseName) { toast('未能确定课程名，请在面板课程框里手动填写。', true); return; }
    var course = loadCourse(courseName);
    if (!course) { toast('该课程还没有任何记录，先做一章点「开始记录」。', true); return; }
    var bytes = buildDocx(coursesToDoc(course));
    var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = (courseName || '智慧树题库') + '.docx';
    document.documentElement.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(url);
      if (a.parentNode) a.parentNode.removeChild(a);
    }, 800);
    toast('已生成 Word 文档：' + a.download + '（若浏览器拦截下载请允许）');
  }

  function doClearCourse() {
    var courseName = currentCourseGuess();
    if (!courseName) { toast('未确定课程名。', true); return; }
    if (!window.confirm('确定清空课程「' + courseName + '」的全部记录？此操作不可恢复。')) return;
    if (!window.confirm('再次确认：真的要清空吗？')) return;
    GM_setValue(courseKeyFor(courseName), 'null');
    var list = loadCourseList().filter(function (c) { return c.name !== courseName; });
    GM_setValue(K_INDEX, JSON.stringify(list));
    toast('已清空课程：' + courseName);
  }

  function initUI() {
    if (window.__zhrV1UI) return;
    if (!document.body) return;
    var bodyLen = (document.body.innerText || '').length;
    var showHere = (window.top === window) || bodyLen >= 1200;
    if (!showHere) return;
    window.__zhrV1UI = true;

    GM_addStyle(
      '#zhr-root{position:fixed;right:16px;bottom:16px;z-index:2147483646;font:12px/1.5 "Microsoft YaHei",sans-serif;color:#222;width:330px;border-radius:8px;box-shadow:0 6px 24px rgba(0,0,0,.25);background:#fff;overflow:hidden}' +
      '#zhr-head{background:#2f6fed;color:#fff;padding:8px 12px;cursor:move;font-weight:700;display:flex;align-items:center;user-select:none}' +
      '#zhr-head .t{flex:1}' +
      '#zhr-head button{background:transparent;border:0;color:#fff;cursor:pointer;font-size:14px;padding:0 2px 0 8px}' +
      '#zhr-body{padding:10px 12px}' +
      '#zhr-body label{display:block;margin-bottom:6px;color:#666}' +
      '#zhr-body input{width:100%;box-sizing:border-box;border:1px solid #ccc;border-radius:4px;padding:4px 6px;font-size:12px;margin-top:2px}' +
      '#zhr-btns{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}' +
      '#zhr-btns button{flex:1;min-width:90px;border:0;border-radius:6px;padding:7px 4px;cursor:pointer;font-size:12px}' +
      '#zhr-b1{background:#2f6fed;color:#fff}' +
      '#zhr-b2{background:#27ae60;color:#fff}' +
      '#zhr-b3{background:#eef1f6;color:#333}' +
      '#zhr-b4{background:#fdecea;color:#c0392b}' +
      '#zhr-tip{margin-top:8px;color:#999;font-size:11px}'
    );

    var root = document.createElement('div');
    root.id = 'zhr-root';
    var head = document.createElement('div');
    head.id = 'zhr-head';
    var t = document.createElement('span');
    t.className = 't';
    t.textContent = '智慧树课后题记录器';
    var bMin = document.createElement('button');
    bMin.textContent = '—';
    bMin.title = '最小化';
    head.appendChild(t);
    head.appendChild(bMin);
    var body = document.createElement('div');
    body.id = 'zhr-body';

    var lb1 = document.createElement('label');
    lb1.textContent = '课程名（自动识别，可改）';
    var courseInput = document.createElement('input');
    courseInput.id = 'zhr-course-input';
    courseInput.placeholder = '自动识别中…';
    lb1.appendChild(courseInput);
    var lb2 = document.createElement('label');
    lb2.textContent = '章节名（自动识别，可改）';
    var chapterInput = document.createElement('input');
    chapterInput.id = 'zhr-chapter-input';
    chapterInput.placeholder = '自动识别中…';
    lb2.appendChild(chapterInput);

    var btns = document.createElement('div');
    btns.id = 'zhr-btns';
    var bRec = document.createElement('button');
    bRec.id = 'zhr-b1'; bRec.textContent = '▶ 开始记录';
    var bWord = document.createElement('button');
    bWord.id = 'zhr-b2'; bWord.textContent = '存为 Word';
    var bDiag = document.createElement('button');
    bDiag.id = 'zhr-b3'; bDiag.textContent = '🔍 诊断';
    var bClear = document.createElement('button');
    bClear.id = 'zhr-b4'; bClear.textContent = '清空本课';
    btns.appendChild(bRec); btns.appendChild(bWord); btns.appendChild(bDiag); btns.appendChild(bClear);

    var tip = document.createElement('div');
    tip.id = 'zhr-tip';
    tip.textContent = '用法：做完一章→在“成绩/查看解析”页点「开始记录」；每章点一次，自动累计。快捷键 Ctrl+Shift+X 也可记录。';

    body.appendChild(lb1); body.appendChild(courseInput);
    body.appendChild(lb2); body.appendChild(chapterInput);
    body.appendChild(btns); body.appendChild(tip);
    root.appendChild(head); root.appendChild(body);
    document.documentElement.appendChild(root);

    courseInput.value = guessCourseName();
    chapterInput.value = guessChapterName();

    /* 拖动 */
    var dragging = false, dx = 0, dy = 0;
    head.addEventListener('mousedown', function (e) {
      dragging = true;
      dx = e.clientX - root.getBoundingClientRect().left;
      dy = e.clientY - root.getBoundingClientRect().top;
      e.preventDefault();
    });
    document.addEventListener('mousemove', function (e) {
      if (!dragging) return;
      root.style.left = Math.max(0, Math.min(window.innerWidth - root.offsetWidth, e.clientX - dx)) + 'px';
      root.style.top = Math.max(0, Math.min(window.innerHeight - root.offsetHeight, e.clientY - dy)) + 'px';
    });
    document.addEventListener('mouseup', function () { dragging = false; });

    var minimized = false;
    bMin.addEventListener('click', function () {
      minimized = !minimized;
      body.style.display = minimized ? 'none' : '';
      bMin.textContent = minimized ? '＋' : '—';
    });

    bRec.addEventListener('click', doRecord);
    bWord.addEventListener('click', doExportWord);
    bDiag.addEventListener('click', function () {
      showModal('v1 诊断快照（复制后发给开发者）', '', diagnosticText());
    });
    bClear.addEventListener('click', doClearCourse);

    document.addEventListener('keydown', function (e) {
      if (e.ctrlKey && e.shiftKey && e.code === 'KeyX') { e.preventDefault(); doRecord(); }
    });

    setTimeout(function () {
      if (!courseInput.value) courseInput.value = guessCourseName();
      if (!chapterInput.value) chapterInput.value = guessChapterName();
    }, 1500);
  }

  return {
    _version: VERSION,
    utf8: utf8, djb2: djb2, normalizeStem: normalizeStem, xmlEscape: xmlEscape,
    crc32: crc32, zipStore: zipStore, buildDocx: buildDocx,
    stemKey: stemKey, optsHash: optsHash, mergeQuestions: mergeQuestions,
    coursesToDoc: coursesToDoc,
    init: initUI
  };
})();

/* Node 单元测试环境下导出纯函数；浏览器环境自动挂 UI */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = ZHR;
} else {
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { ZHR.init(); });
  } else {
    ZHR.init();
  }
}
