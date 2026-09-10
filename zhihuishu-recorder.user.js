// ==UserScript==
// @name         智慧树课后题记录器 v1
// @namespace    https://dsh.local/zhihuishu-recorder
// @version      1.8.3
// @description  在你做完智慧树章节测验并进入「本次成绩/查看答案解析」页后，点「开始记录」把本章题目+正确答案存入本地题库（跨章节累计、按题干去重、选项乱序变体保留），可另存为 Word(.docx)。内置页面结构侦察/运行错误收集与「自动遍历」（仅自动打开解析并读取已展示内容，不答题、不提交）。纯本地运行，不联网。
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
  var VERSION = '1.8.3';
  /* ---------------- 运行期错误收集（供诊断报告展示） ---------------- */
  var ERRORS = [];
  function collectErr(ev) {
    try {
      var msg = ev && (ev.message || (ev.reason && ev.reason.message) || ev.reason) || 'unknown error';
      var tag = (ev && ev.filename) || (ev && ev.target && ev.target.nodeName) || '';
      var line = new Date().toLocaleTimeString('zh-CN') + '  ' + String(msg).slice(0, 260) + (tag && ev.lineno ? '  @' + tag + ':' + ev.lineno : (tag ? '  @' + tag : ''));
      if (ERRORS.indexOf(line) < 0) { ERRORS.push(line); if (ERRORS.length > 30) ERRORS.shift(); }
    } catch (x) { /* ignore */ }
  }
  if (typeof window !== 'undefined' && window.addEventListener) {
    try {
      window.addEventListener('error', collectErr);
      window.addEventListener('unhandledrejection', collectErr);
    } catch (x2) { /* ignore */ }
  }

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
          course: inc.course || '',
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
        if (!list[matched].course && inc.course) list[matched].course = inc.course;
        if (inc.answer && inc.answer !== list[matched].answer) { list[matched].answer = inc.answer; }
        return;
      }
      var v = {
        id: k + '_v' + (list.length + variants.length),
        stem: normalizeStem(inc.stem),
        type: inc.type || 'other',
        options: (inc.options || []).slice(),
        answer: inc.answer || '',
        course: inc.course || '',
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

  /* 单题库 → 按 课程 → 章节 分组（当前课程排最前），供 Word 排版 */
  function bankToDoc(bank, focus) {
    var order = [], by = {};
    (bank.questions || []).forEach(function (q) {
      var c = q.course || '未命名课程';
      if (!by[c]) { by[c] = []; order.push(c); }
      by[c].push(q);
    });
    order.sort(function (a, b) {
      if (a === focus && b !== focus) return -1;
      if (b === focus && a !== focus) return 1;
      return 0;
    });
    var sections = [];
    order.forEach(function (c) {
      var sub = coursesToDoc({ name: c, questions: by[c] });
      (sub.sections || []).forEach(function (s) {
        sections.push({ name: order.length > 1 ? ('【' + c + '】' + s.name) : s.name, questions: s.questions });
      });
    });
    var title = order.length > 1 ? ('智慧树题库（' + order.length + ' 门课）') : (order[0] || '智慧树题库');
    return { course: title, sections: sections };
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
  var ANSWER_LINE_RE = /(正确答案|参考答案|标准答案|答案[:：])/;
  var ANSWER_LINE_ANY_RE = /答案[:：]?\s*([^\n]{0,60})/;
  var CORRECT_CLS_RE = /\b(correct|right|answer|daan|green|succ|pass|true)\w*/i;
  var MARKED_WRONG_RE = /\b(wrong|incorrect|false|error)\w*/i;
  var CONTAINER_CLS_RE = /(question|timu|topic|subject|exam|test|answer|result|choice|item)/i;
  /* 导航/目录类容器：不是题目 */
  var NAV_CLS_RE = /(catalog|catalogue|menu|nav|sidebar|sider|aside|breadcrumb|crumb|tree|directory|toc|tab)/i;
  /* 按钮/图标类元素：即便 class 含 answer 也不作为“答案标记” */
  var UI_EL_RE = /(btn|button|icon|toolbar|ai-|advert|avatar|logo)/i;
  /* 答题卡/答案汇总区：不能被当成题目记录 */
  var CARD_RE = /(答题卡|答案卡|答题情况|答案汇总|全部答案|答案速览|答题记录|答案列表|答案一览)/;
  /* 标题/课程信息类文本：无选项且无答案时不当题目 */
  var NOISE_RE = /(学年|学期|智慧树|知到|精品课程|课程介绍|学习进度|课程中心|我的课程|目录|导航|题库总览)/;

  function splitLines(t) {
    return String(t || '').split(/\r?\n/).map(function (s) { return normWs(s); }).filter(Boolean);
  }

  function clsOf(el) {
    if (typeof el.className === 'string') return el.className;
    if (el.classList && el.classList.length) return Array.prototype.join.call(el.classList, ' ');
    return '';
  }

  function innerTextOf(el) { return (el.innerText || el.textContent || ''); }

  function markerByClass(el) {
    var c = clsOf(el);
    if (!c || !CORRECT_CLS_RE.test(c)) return false;
    if (UI_EL_RE.test(c)) return false;              /* 按钮/图标不算答案标记 */
    return true;
  }

  function hasAnswerMarker(el) {
    var t = innerTextOf(el);
    if (ANSWER_LINE_RE.test(t)) return true;
    if (markerByClass(el)) return true;
    var marked = el.querySelectorAll('[class]');
    for (var i = 0; i < marked.length && i < 200; i++) {
      if (markerByClass(marked[i])) return true;
    }
    return false;
  }

  /* 候选里可能有“只有答案行、没有题干”的碎块，它们会把父级题目块一起排掉 → 先剔除 */
  function topLevelContainers(cands) {
    if (!cands.length) return cands;
    var usable = [];
    for (var i = 0; i < cands.length; i++) {
      var p = null;
      try { p = parseLines(innerTextOf(cands[i])); } catch (e) { p = null; }
      if (p && p.stem) usable.push(cands[i]);
    }
    if (!usable.length) return [];
    /* 取最内层（去掉包含其它可用候选的父块） */
    return usable.filter(function (a) {
      for (var j = 0; j < usable.length; j++) {
        if (usable[j] !== a && a.contains(usable[j])) return false;
      }
      return true;
    });
  }

  /* 找“整章回看/解析”页里的题目容器：class/id 命中关键词 且 含答案标记 */
  function findQuestionContainers(doc) {
    var all = doc.querySelectorAll('div,li,section,tr,fieldset');
    var cands = [];
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
      var idc = el.id || '';
      var clsc = clsOf(el);
      if (NAV_CLS_RE.test(clsc) || NAV_CLS_RE.test(idc)) continue;   /* 导航/目录类不算题目 */
      if (!(CONTAINER_CLS_RE.test(idc) || CONTAINER_CLS_RE.test(clsc))) continue;
      var txt = innerTextOf(el);
      if (txt.length < 8 || txt.length > 20000) continue;
      if (CARD_RE.test(normWs(txt).slice(0, 80))) continue;   /* 答题卡/答案汇总 */
      if (!hasAnswerMarker(el)) continue;
      cands.push(el);
    }
    return topLevelContainers(cands);
  }

  /* 内容特征判定：答题卡/答案汇总块（题号紧跟答案字母的行占多数；或题干里混入多个题号） */
  function looksLikeAnswerCard(container, parsed) {
    var raw = innerTextOf(container);
    if (CARD_RE.test(normWs(raw).slice(0, 80))) return true;
    var lines = splitLines(raw);
    if (lines.length >= 4) {
      var n = 0;
      for (var i = 0; i < lines.length; i++) {
        if (/^\d{1,3}\s*[.、．)）:：]?\s*[A-Ha-h√×对错]/.test(lines[i])) n++;
      }
      if (n >= 4 && n / lines.length >= 0.5) return true;
    }
    var stem = (parsed && parsed.stem) || '';
    if (/^(答题卡|答案卡|答案汇总|全部答案|答题情况)/.test(stem)) return true;
    if ((stem.match(/\d{1,3}\s*[.、．]/g) || []).length >= 3) return true;
    return false;
  }

  /* 标题/导航/空壳块：既无选项又无答案的，不当题目记录 */
  function looksLikeNoise(container, q) {
    var stem = (q && q.stem) || '';
    var hasOpts = !!(q && q.options && q.options.length);
    var hasAns = !!(q && q.answer);
    if (!hasOpts && !hasAns) return true;                 /* 不可作答的空壳 */
    if (!hasOpts && NOISE_RE.test(stem) && stem.length <= 40) return true;  /* 课程/学期标题 */
    if (cleanText(stem).length < 2) return true;          /* 题干过短 */
    return false;
  }

  /* 题干收尾：去题型标签 + 去题号 */
  function finalizeStem(s) {
    var t = cleanText(s);
    t = t.replace(/^(单选|多选|判断|填空|简答|不定项|单项选择|多项选择)\s*题?\s*(?=\d)/, '');
    return normalizeStem(t);
  }

  /* 文本净化：去掉界面文案、表情、零宽字符（题目题干/选项里不该有的杂质） */
  function cleanText(s) {
    var t = String(s || '');
    t = t.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, ' ');                       /* emoji（代理对） */
    t = t.replace(/[\u2190-\u21FF\u2600-\u27BF\u2B00-\u2BFF\uFE0F\u200D\u25A0-\u25FF]/g, ' ');
    t = t.replace(/(AI答题辅导|AI解析|AI智能解析|收藏为薄弱题|加入薄弱题|加入错题|薄弱题|收藏|纠错|报错|举报|分享|添加笔记|记笔记|笔记|点赞)/g, ' ');
    return t.replace(/\s+/g, ' ').trim();
  }

  /* 是否是“无字母的选项文字行”（多个短词，无括号/句末标点），是则返回词列表 */
  function bareOptionWords(line) {
    var t = cleanText(line);
    if (!t || t.length > 200) return null;
    if (/[。.；;：:!?？]$/.test(t)) return null;
    if (/[（()）]/.test(t)) return null;
    var parts = t.split(/\s+/);
    if (parts.length < 2 || parts.length > 8) return null;
    for (var i = 0; i < parts.length; i++) {
      if (parts[i].length > 24) return null;
    }
    return parts;
  }

  /* 宽松的选项字母匹配：字母可带括号/标点，后面可以有文字也可以没文字 */
  var OPT_SOFT_RE = /^[（(]?\s*([A-Ha-hＡ-Ｈａ-ｈ])\s*[)）.．、:：]?\s*(\S.*)?$/;
  /* 状态行/自答行/解析/知识点：不进入题干/选项 */
  var STATUS_LINE_RE = /^(我的答案|您的答案|回答正确|回答错误|答对|答错|已作答|未作答|答案解析|解析|考查知识点|知识点)/;
  /* 下一题的题号行 */
  var QNUM_LINE_RE = /^\d{1,3}\s*[.、．)）]\s*/;

  /* 一行里塞了多个选项（如 "A. 甲  B. 乙  C. 丙"）→ 拆成多行（只当字母 A、B、C… 连续时才拆） */
  function splitInlineOptions(line) {
    var t = normWs(line);
    if (t.length < 12) return [line];
    var re = /[（(]?([A-Ha-hＡ-Ｈａ-ｈ])[)）.．、:：]\s*/g;
    var hits = [], m;
    while ((m = re.exec(t))) {
      hits.push({ at: m.index, letter: normLetter(m[1]).toUpperCase() });
      if (hits.length > 10) break;
      if (re.lastIndex <= m.index) re.lastIndex = m.index + 1;
    }
    if (hits.length < 2) return [line];
    for (var i = 0; i < hits.length; i++) {
      if (hits[i].letter !== String.fromCharCode(65 + i)) return [line];
    }
    var out = [];
    var head = t.slice(0, hits[0].at).trim();
    if (head) out.push(head);
    for (var j = 0; j < hits.length; j++) {
      var end = (j + 1 < hits.length) ? hits[j + 1].at : t.length;
      out.push(t.slice(hits[j].at, end).trim());
    }
    return out;
  }

  /* 宽松选项解析：兼容“字母单独一行 + 文字在下一行”“A 文字（无标点）”
     返回 {stem, options}；字母必须 A、B、C… 连续，否则不认（避免把题干误当选项） */
  function parseOptionsLoose(lines) {
    var strong = [], weak = [];
    for (var i = 0; i < lines.length; i++) {
      var L = lines[i];
      if (!L || ANSWER_LINE_RE.test(L)) continue;
      if (STATUS_LINE_RE.test(L) || QNUM_LINE_RE.test(L)) continue;
      if (L.indexOf('（ ）') >= 0 || /\(\s*\)/.test(L)) continue;         /* 带空括号的多半是题干 */
      var m = L.match(OPT_SOFT_RE);
      if (!m) continue;
      var letter = normLetter(m[1]).toUpperCase();
      var rest = (m[2] || '').trim();
      if (/[。．？?！!]$/.test(rest)) continue;                           /* 完整句子 → 当题干 */
      var hasSep = new RegExp('^[（(]?\\s*' + m[1] + '\\s*[)）.．、:：]').test(L);
      if (rest && !hasSep) weak.push({ i: i, letter: letter, rest: rest });
      else strong.push({ i: i, letter: letter, rest: rest });
    }
    function okSeq(arr) {
      if (arr.length < 2) return false;
      for (var k = 0; k < arr.length; k++) if (arr[k].letter !== String.fromCharCode(65 + k)) return false;
      return true;
    }
    var marks = okSeq(strong) ? strong : (okSeq(weak) ? weak : null);
    if (!marks) return null;
    var opts = [], stemLines = [];
    for (var s = 0; s < marks[0].i; s++) {
      var L3 = lines[s];
      if (!L3 || ANSWER_LINE_RE.test(L3) || STATUS_LINE_RE.test(L3)) continue;
      stemLines.push(L3);
    }
    for (var j = 0; j < marks.length; j++) {
      var buf = [];
      if (marks[j].rest) buf.push(marks[j].rest);
      else {
        var to = (j + 1 < marks.length) ? marks[j + 1].i : lines.length;
        for (var q = marks[j].i + 1; q < to; q++) {
          var L2 = lines[q];
          if (!L2 || ANSWER_LINE_RE.test(L2) || STATUS_LINE_RE.test(L2) || QNUM_LINE_RE.test(L2)) break;
          buf.push(L2);
        }
      }
      opts.push({ letter: marks[j].letter, text: cleanText(buf.join(' ')) });
    }
    return { stem: cleanText(stemLines.join(' ')), options: opts };
  }

  /* 从一个题目容器的行文本里解析题干/选项/答案（行文本模型，兼容多数布局） */
  function parseLines(containerText) {
    var raw = splitLines(containerText);
    var lines = [];
    for (var e = 0; e < raw.length; e++) {
      var parts = splitInlineOptions(raw[e]);
      for (var p = 0; p < parts.length; p++) lines.push(parts[p]);
    }
    var stemLines = [], options = [], bareOpts = [], ansText = '', ansLineIdx = -1, pendingOpts = null;
    for (var i = 0; i < lines.length; i++) {
      var L = lines[i];
      if (ANSWER_LINE_RE.test(L)) {
        if (/我的答案|您的答案/.test(L) && !/正确答案|参考答案|标准答案/.test(L)) continue;
        ansLineIdx = i;
        ansText = L;
        continue;
      }
      /* 状态行/自答行/解析/知识点：不进入题干 */
      if (STATUS_LINE_RE.test(L)) continue;
      var m = L.match(OPT_RE);
      if (m) {
        var letter = normLetter(m[1]).toUpperCase();
        var otext = cleanText(L.replace(OPT_RE, ''));
        /* 选项字母与文字分离的布局：字母行没文字时，把上一行的“裸词列表”拿来做选项文字 */
        if (!options.length && !otext && stemLines.length) {
          var words = bareOptionWords(stemLines[stemLines.length - 1]);
          if (words && words.length >= 2) { stemLines.pop(); pendingOpts = words; }
        }
        if (!otext && pendingOpts && options.length < pendingOpts.length) otext = pendingOpts[options.length];
        /* 字母行没文字 → 把随后的正文行当作这个选项的文字（智慧树常见：字母一行、文字下一行） */
        if (!otext) {
          var buf = [];
          for (var q2 = i + 1; q2 < lines.length; q2++) {
            var L4 = lines[q2];
            if (OPT_RE.test(L4) || OPT_SOFT_RE.test(L4)) break;
            if (ANSWER_LINE_RE.test(L4) || STATUS_LINE_RE.test(L4) || QNUM_LINE_RE.test(L4)) break;
            buf.push(L4);
            if (buf.join(' ').length > 120) break;
          }
          if (buf.length) otext = cleanText(buf.join(' '));
        }
        options.push({ letter: letter, text: otext });
        continue;
      }
      /* 无字母前缀的判断题短选项（对/错/正确/错误/是/否/√/×） */
      if (/^(对|错|正确|错误|是|否|√|×)$/.test(L)) { bareOpts.push(L); continue; }
      if (ansLineIdx === -1) stemLines.push(L);
    }
    var stem = finalizeStem(stemLines.join(' '));
    /* 严格解析没拿到选项文字（或选项少于两个）→ 走宽松解析（修复"只读到字母、读不到内容"） */
    var emptyOpts = options.length > 0;
    for (var z = 0; z < options.length; z++) { if (options[z].text) { emptyOpts = false; break; } }
    if (options.length < 2 || emptyOpts) {
      var loose = null;
      try { loose = parseOptionsLoose(lines); } catch (eL) { loose = null; }
      if (loose && loose.options.length >= 2) {
        var looseStem = finalizeStem(loose.stem || stemLines.join(' '));
        return { stem: looseStem || stem, options: loose.options, answerLine: ansText };
      }
    }
    /* 判断题：字母选项缺失时，用裸的对/错选项补上 */
    if (!options.length && bareOpts.length) {
      for (var j = 0; j < bareOpts.length; j++) {
        options.push({ letter: String.fromCharCode(65 + j), text: bareOpts[j] });
      }
    }
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
    /* 题型标签（“多选 题”）常被题干净化掉，所以再看一眼整块原文 */
    if (/(多选|多项选择|多项|不定项)/.test(raw) && type === 'single') type = 'multi';
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

  /* ---------------- 课程名 / 章节名识别 ----------------
   *  思路：优先读左侧目录树里“当前高亮项”，章节名 = 所在章（若高亮的是章则取它下面第一个单元），
   *  再依次回退到页面标题/面包屑/正文里的“第X章”。
   */
  var CHAP_RE = /(第\s*[0-9一二三四五六七八九十百零]+\s*[章节讲单元模块篇]|Chapter\s*\d+|Unit\s*\d+|Module\s*\d+)/i;
  var UNIT_RE = /^(\d+(\.\d+)+\s*|第\s*[0-9一二三四五六七八九十百零]+\s*[节讲课]|\d+\s*[-、．.]\s*)/;

  function visOk(el) {
    if (!el || !el.getBoundingClientRect) return false;
    var r = el.getBoundingClientRect();
    if (r.width < 2 || r.height < 2) return false;
    var w = (el.ownerDocument && el.ownerDocument.defaultView) || window;
    try {
      var cs = w.getComputedStyle(el);
      if (cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0') return false;
    } catch (e) { /* ignore */ }
    return true;
  }

  /* 收集侧边目录/章节树中的可见条目（按文档顺序），取“最像目录”的那个容器 */
  function collectCatalogItems() {
    var sel = '[class*="catalog" i],[class*="catalogue" i],[class*="chapter" i],[class*="tree" i],' +
              '[class*="directory" i],[class*="menu" i],[class*="nav" i],[class*="sider" i],' +
              '[class*="aside" i],[class*="knowledge" i],[class*="outline" i],[class*="catalogItem" i]';
    var docs = pilotDocs();
    var best = null;
    for (var d0 = 0; d0 < docs.length; d0++) {
      var scopes;
      try { scopes = docs[d0].querySelectorAll(sel); } catch (e0) { continue; }
      for (var i = 0; i < scopes.length; i++) {
        var sc = scopes[i];
        if (sc.closest && sc.closest('#zhr-root,#zhr-scout-root')) continue;
        if (/\b(question|timu|exam|test|answer|result|choice)\b/i.test(clsOf(sc))) continue;  /* 别把题目区当目录 */
        if (!visOk(sc)) continue;
        var els;
        try { els = sc.querySelectorAll('a,li'); } catch (e1) { continue; }
        var items = [];
        for (var j = 0; j < els.length && items.length < 200; j++) {
          var el = els[j];
          if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
          if (!visOk(el)) continue;
          var t = cleanText(el.innerText || '');
          if (!t || t.length > 60) continue;
          items.push({ el: el, t: t });
        }
        if (items.length >= 3 && (!best || items.length > best.length)) best = items;
      }
    }
    return best;
  }

  function isActiveItem(el) {
    if (!el) return false;
    var c = clsOf(el);
    if (/(active|current|selected|checked|focus|on$|\bon\b)/i.test(c)) return true;
    try {
      if (el.getAttribute('aria-selected') === 'true' || el.getAttribute('aria-current')) return true;
    } catch (e) { /* ignore */ }
    var p = el.parentElement;
    if (p) {
      var pc = clsOf(p);
      if (/(active|current|selected|checked|focus)/i.test(pc)) return true;
    }
    return false;
  }

  function joinLoc(chap, unit) {
    chap = cleanText(chap); unit = cleanText(unit);
    if (chap && unit && chap !== unit) return chap + ' · ' + unit;
    return chap || unit || '';
  }

  function guessCourseName() {
    /* 1) 优先取“页面左上角”那块（课程名通常就在左上角）
       内容常在子 iframe 里，而面板挂在顶层文档，所以必须跨同源 frame 一起扫 */
    var docs = pilotDocs();
    var cand = [];
    for (var d0 = 0; d0 < docs.length; d0++) {
      var doc = docs[d0];
      var vw = 1200, vh = 800;
      try {
        var w0 = doc.defaultView;
        if (w0) { vw = w0.innerWidth || vw; vh = w0.innerHeight || vh; }
      } catch (e0) { /* ignore */ }
      var els;
      try { els = doc.querySelectorAll('h1,h2,h3,h4,h5,div,span,a,p,strong,b'); } catch (e1) { continue; }
      for (var i = 0; i < els.length && i < 4000; i++) {
        var el = els[i];
        if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
        var t = cleanText(el.innerText || '');
        if (!t || t.length < 2 || t.length > 40) continue;
        if (/^(首页|主页|登录|智慧树|知到|我的|课程中心|学习中心|个人中心)$/.test(t)) continue;
        if (CHAP_RE.test(t)) continue;                       /* 章节标题不当课程名 */
        var r = el.getBoundingClientRect();
        if (r.width < 8 || r.height < 8) continue;
        if (r.left > vw * 0.35 || r.top > vh * 0.30) continue;  /* 只看左上角区域 */
        if (r.width > vw * 0.6) continue;
        /* 档次：含“学年/学期”的最像课程名，其次“课程/大学/学院”，最后其它 */
        var tier = /(学年|学期)/.test(t) ? 0 : (/(课程|大学|学院)/.test(t) ? 1 : 2);
        cand.push({ t: t, tier: tier, score: r.left * 2 + r.top });
      }
    }
    if (cand.length) {
      cand.sort(function (a, b) { return (a.tier - b.tier) || (a.score - b.score); });
      return cand[0].t;
    }
    var title = normWs(document.title || '');
    /* 2) 标题："课程名 - 智慧树" / "智慧树 - 课程名" */
    var m = title.match(/^(.*?)\s*[-_—|｜]\s*(智慧树|知到)/);
    if (m && m[1] && !/登录|首页|智慧树|知到/.test(m[1])) return cleanText(m[1]);
    var m2 = title.match(/(?:智慧树|知到)\s*[-_—|｜]\s*(.+)$/);
    if (m2 && m2[1]) return cleanText(m2[1]);
    /* 3) 明确的“课程名”元素 / 面包屑 */
    var sels = ['[class*="courseName" i]', '[class*="course-name" i]', '[class*="course_title" i]',
                '[class*="courseTitle" i]', '[class*="course" i] h1', '[class*="course" i] h2',
                '[class*="course" i] h3', '[class*="breadcrumb" i] a', '[class*="crumb" i] a',
                '[class*="breadcrumb" i] span', 'h1', 'h2'];
    for (var s = 0; s < sels.length; s++) {
      var es2 = document.querySelectorAll(sels[s]);
      for (var j = 0; j < es2.length && j < 10; j++) {
        if (es2[j].closest && es2[j].closest('#zhr-root,#zhr-scout-root')) continue;
        var t2 = cleanText(es2[j].innerText || es2[j].textContent || '');
        if (!t2 || t2.length < 2 || t2.length > 40) continue;
        if (/^(首页|主页|登录|智慧树|知到|我的|课程中心|学习中心|个人中心)$/.test(t2)) continue;
        if (CHAP_RE.test(t2)) continue;
        return t2;
      }
    }
    /* 4) 退一步用标题（去掉站点名） */
    if (title && title.length <= 40 && !/智慧树|知到|登录/.test(title)) return cleanText(title);
    return '';
  }

  function guessChapterName() {
    var items = collectCatalogItems();
    if (!items || !items.length) {
      /* class 关键词没命中时，用“通用目录识别”的结果 */
      try {
        var raw = pilotCatalogItems();
        if (raw && raw.length) {
          items = [];
          for (var ri = 0; ri < raw.length; ri++) items.push({ el: raw[ri], t: cleanText(raw[ri].innerText || '') });
        }
      } catch (e0) { /* ignore */ }
    }
    if (items && items.length) {
      var act = -1;
      for (var i = 0; i < items.length; i++) { if (isActiveItem(items[i].el)) { act = i; break; } }
      if (act >= 0) {
        var cur = items[act].t;
        if (CHAP_RE.test(cur)) {
          /* 当前就在章上：取该章下第一个单元名 */
          var first = (act + 1 < items.length) ? items[act + 1].t : '';
          return joinLoc(cur, first);
        }
        /* 当前在单元上：往前找最近的章 */
        var chap = '';
        for (var k = act - 1; k >= 0; k--) {
          if (CHAP_RE.test(items[k].t)) { chap = items[k].t; break; }
        }
        if (!chap) {
          for (var q = 0; q < items.length; q++) { if (CHAP_RE.test(items[q].t)) { chap = items[q].t; break; } }
        }
        return joinLoc(chap, cur);
      }
      /* 无高亮：找第一个“章”，并用它后面第一个单元 */
      for (var m = 0; m < items.length; m++) {
        if (CHAP_RE.test(items[m].t)) {
          var nx = (m + 1 < items.length) ? items[m + 1].t : '';
          return joinLoc(items[m].t, nx);
        }
      }
    }
    /* 回退：正文里找“第X章”，再找同容器里第一个单元行 */
    var re = /第\s*[0-9一二三四五六七八九十百零]+\s*[章节讲单元课]\s*[^\s]{0,20}/;
    var best = '', bestEl = null;
    var docs2 = pilotDocs();
    for (var d3 = 0; d3 < docs2.length; d3++) {
      var els;
      try { els = docs2[d3].querySelectorAll('h1,h2,h3,h4,.chapter,.section,.unit,li,span,div,b,strong,p'); } catch (e3) { continue; }
      for (var n = 0; n < els.length && n < 300; n++) {
        var el = els[n];
        if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
        var t = cleanText(el.innerText || '');
        if (!t || t.length > 40) continue;
        if (re.test(t)) {
          if (!best || t.length < best.length) { best = t; bestEl = el; }
        }
      }
    }
    if (best) {
      var unit = '';
      if (bestEl && bestEl.parentElement) {
        var sib = bestEl.nextElementSibling;
        if (sib) {
          var st = cleanText(sib.innerText || '');
          if (st && st.length <= 40) unit = st;
        }
      }
      return joinLoc(best, unit);
    }
    return '';
  }

  /* ---------------- UI ---------------- */

  function toast(msg, isError) {
    var wrap = uiDoc().getElementById('zhr-root');
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

  /* 弹窗挂到“最顶层可访问的文档”：智慧树学习页是多层 iframe，
     若把弹窗留在小 iframe 内，会看不见、或被随页面重建一起清掉。同源时提升到顶层窗口。 */
  function uiDoc() {
    var w = window;
    for (;;) {
      var p = null;
      try { p = w.parent; } catch (e) { break; }
      if (!p || p === w) break;
      var ok = false;
      try { ok = !!(p.document && p.document.documentElement); } catch (e2) { ok = false; }
      if (!ok) break;
      w = p;
    }
    try { return w.document; } catch (e3) { return document; }
  }
  function uiAttach(el) {
    var d = uiDoc();
    (d.body || d.documentElement).appendChild(el);
    return d;
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
    uiAttach(mask);
    function close() { if (mask.parentNode) mask.parentNode.removeChild(mask); }
    bClose.addEventListener('click', close);
    mask.addEventListener('click', function (e) { if (e.target === mask) close(); });
  }

  /* ---------------- 存储（GM_* 跨子域共享） ---------------- */

  var K_INDEX = 'zhr.v1.index';
  var K_UI_COURSE = 'zhr.v1.ui.course';
  var K_UI_CHAPTER = 'zhr.v1.ui.chapter';
  var K_UI_COURSE_MANUAL = 'zhr.v1.ui.course.manual';
  var K_UI_CHAPTER_MANUAL = 'zhr.v1.ui.chapter.manual';
  /* GM 权限缺失时降级到同源 localStorage（面板/记录仍可用；跨子域共享退化为同源） */
  function sGet(k) {
    try { if (typeof GM_getValue === 'function') return GM_getValue(k, null); } catch (e) { /* ignore */ }
    try { return localStorage.getItem('zhrgm.' + k); } catch (e2) { return null; }
  }
  function sSet(k, v) {
    try { if (typeof GM_setValue === 'function') { GM_setValue(k, v); return; } } catch (e) { /* ignore */ }
    try { localStorage.setItem('zhrgm.' + k, v); } catch (e2) { /* ignore */ }
  }
  function loadCourseList() { return JSON.parse(sGet(K_INDEX) || 'null') || []; }
  function courseKeyFor(name) { return 'zhr.v1.course.' + djb2(name || '未命名课程'); }
  function loadCourse(name) {
    var c = JSON.parse(sGet(courseKeyFor(name)) || 'null');
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
    sSet(courseKeyFor(name), JSON.stringify(data));
    var list = loadCourseList();
    var hit = null;
    for (var i = 0; i < list.length; i++) if (list[i].name === name) { hit = list[i]; break; }
    if (!hit) { hit = { name: name }; list.push(hit); }
    hit.updatedAt = Date.now();
    sSet(K_INDEX, JSON.stringify(list));
  }

  /* ---------------- 题库（v2：单一题库，不再按课程名分库） ----------------
   *  v1 把题目存在 "zhr.v1.course.<课程名哈希>" 里，课程名识别一变就开到另一个"空库"，
   *  看上去像"记录被清空了"。v2 改成全局一个库，课程名/章节名只是每题身上的**标签**，
   *  所以换视频 / 换页面 / 识别到的课程名变化，都不会丢题目。
   */
  var K_BANK = 'zhr.v2.bank';
  var K_MIGRATED = 'zhr.v2.migrated';

  function ensureBank() {
    var bank = null;
    try { bank = JSON.parse(sGet(K_BANK) || 'null'); } catch (e) { bank = null; }
    if (!bank || typeof bank !== 'object') bank = { questions: [] };
    if (!bank.questions) bank.questions = [];
    if (sGet(K_MIGRATED) !== '1') {
      /* 把 v1 的分库数据（含更早的 chapters 结构）搬进来；旧键保留不删，安全可回退 */
      try {
        var list = loadCourseList();
        for (var i = 0; i < list.length; i++) {
          var old = loadCourse(list[i].name);
          if (!old || !old.questions || !old.questions.length) continue;
          var qs = [];
          for (var j = 0; j < old.questions.length; j++) {
            var q = old.questions[j];
            qs.push({
              stem: q.stem, type: q.type, options: q.options, answer: q.answer,
              course: q.course || list[i].name || '', chapter: q.chapter || ''
            });
          }
          bank.questions = mergeQuestions(bank.questions, qs).list;
        }
      } catch (e2) { /* 迁移出错不影响使用 */ }
      sSet(K_MIGRATED, '1');
      sSet(K_BANK, JSON.stringify(bank));
    }
    return bank;
  }
  function saveBank(bank) { sSet(K_BANK, JSON.stringify(bank)); }
  function bankQuestions() { return ensureBank().questions || []; }

  /* ---------------- 核心动作 ---------------- */

  /* 面板状态记忆：课程名/章节名持久化，切换界面（页面重载）后自动恢复；
     只有点「清空本课」才清除（且只清当前课程标签的那些题）。
     题目存在**单一题库**里，课程名只是标签，所以换视频/换页面都不会丢记录。 */

  /* 某门课在题库里有多少题 */
  function courseRecordsCount(name) {
    if (!name) return 0;
    var qs = null;
    try { qs = bankQuestions(); } catch (e) { return 0; }
    var n = 0;
    for (var i = 0; i < qs.length; i++) if (qs[i].course === name) n++;
    return n;
  }

  /* 其它课程标签的记录概览 */
  function otherCoursesText(exceptName) {
    var qs = null;
    try { qs = bankQuestions(); } catch (e) { return ''; }
    var count = {}, order = [];
    for (var i = 0; i < qs.length; i++) {
      var c = qs[i].course || '未命名课程';
      if (c === exceptName) continue;
      if (!(c in count)) { count[c] = 0; order.push(c); }
      count[c]++;
    }
    var parts = [];
    for (var k = 0; k < order.length && parts.length < 4; k++) parts.push(order[k] + '(' + count[order[k]] + '题)');
    return parts.join('、');
  }

  function statText() {
    var qs = bankQuestions();
    var name = currentCourseGuess();
    if (!qs.length) return '题库暂无记录（做完一章在解析页点「开始记录」）';
    var mine = 0, chs = {};
    for (var i = 0; i < qs.length; i++) {
      if (qs[i].course !== name) continue;
      mine++;
      chs[qs[i].chapter || '未命名章节'] = 1;
    }
    var head = '题库共 ' + qs.length + ' 题';
    if (!mine) {
      var others = otherCoursesText(name);
      return head + '；本课（' + (name || '未命名课程') + '）0 题' + (others ? '，其它：' + others : '');
    }
    return head + '；本课已记录 ' + mine + ' 题，覆盖 ' + Object.keys(chs).length + ' 个章节';
  }
  function refreshStatBar() {
    var el = uiDoc().getElementById('zhr-stat');
    if (el) el.textContent = statText();
  }
  function recordedText() {
    var qs = bankQuestions();
    var focus = currentCourseGuess();
    if (!qs.length) return '（题库暂无记录）\n做完一章、进入「本次成绩/查看解析」页后点「▶ 开始记录」即可累积。';
    var order = [], by = {};
    for (var i = 0; i < qs.length; i++) {
      var c = qs[i].course || '未命名课程';
      if (!by[c]) { by[c] = []; order.push(c); }
      by[c].push(qs[i]);
    }
    order.sort(function (a, b) {
      if (a === focus && b !== focus) return -1;
      if (b === focus && a !== focus) return 1;
      return 0;
    });
    var out = [];
    out.push('题库共 ' + qs.length + ' 题，' + order.length + ' 个课程标签' + (focus ? '（当前：' + focus + '）' : ''));
    for (var k = 0; k < order.length; k++) {
      var name = order[k];
      var doc = coursesToDoc({ name: name, questions: by[name] });
      out.push('');
      out.push('════ ' + name + '（' + by[name].length + ' 题）════');
      doc.sections.forEach(function (sec) {
        out.push('');
        out.push('【' + sec.name + '】');
        sec.questions.forEach(function (q, j) {
          out.push((j + 1) + '. ' + q.stem);
          if (q.options && q.options.length) out.push('   ' + optionsLine(q));
          out.push('   答案：' + (q.answer || '未记录'));
        });
      });
    }
    return out.join('\n');
  }

  /* 课程名识别带缓存（resolveCourseName 会被频繁调用） */
  var GUESS_CACHE = { t: 0, v: '' };
  function guessCourseNameCached() {
    var now = Date.now();
    if (now - GUESS_CACHE.t < 1200) return GUESS_CACHE.v;
    var v = '';
    try { v = guessCourseName() || ''; } catch (e) { v = ''; }
    GUESS_CACHE = { t: now, v: v };
    return v;
  }

  /* 课程名解析（v1.8.0：课程名只是题库里的“标签”，不再决定题目存在哪）
     手改 → 以手改为准；否则识别到含「学年/学期」的名字（课程名典型形态）就用它；
     再否则用面板里现有的值（保持稳定不抽动），最后才用识别值。 */
  function resolveCourseName() {
    var manual = sGet(K_UI_COURSE_MANUAL) === '1';
    var saved = sGet(K_UI_COURSE) || '';
    var inp = uiDoc().getElementById('zhr-course-input');
    var typed = normWs(inp ? inp.value : '');
    if (manual) return typed || saved;
    var guess = guessCourseNameCached();
    if (guess && /(学年|学期)/.test(guess)) return guess;
    return typed || saved || guess;
  }

  function currentCourseGuess() {
    var n = resolveCourseName();
    if (n) return n;
    var inp = uiDoc().getElementById('zhr-course-input');
    return normWs(inp ? inp.value : '') || guessCourseNameCached();
  }

  function doRecord(silent, presetContainers) {
    var containers = presetContainers || findQuestionContainers(document);
    /* 严格识别（class/id 含 question/answer… 关键词）没命中时，走宽松识别（带“答案”行 + 最内层） */
    if (!containers.length) {
      try { containers = pilotContainers(); } catch (e0) { /* ignore */ }
    }
    var chapterGuess = guessChapterName();
    var chapterInput = uiDoc().getElementById('zhr-chapter-input');
    var courseName = currentCourseGuess();
    var chapterName = normWs(chapterInput ? chapterInput.value : '') || chapterGuess || '未命名章节';

    if (!containers.length) {
      if (!silent) toast('未识别到带答案的题目块。若本页确为“成绩/查看解析”页，请点「诊断」把快照发我，或点「重新识别」后重试。', true);
      return { ok: false, reason: 'no-containers' };
    }

    var incoming = [];
    var failed = 0, cardSkipped = 0, noiseSkipped = 0;
    for (var i = 0; i < containers.length; i++) {
      try {
        var q = parseQuestion(containers[i], chapterName);
        if (!q.stem) { failed++; continue; }
        if (looksLikeAnswerCard(containers[i], q)) { cardSkipped++; continue; }
        if (looksLikeNoise(containers[i], q)) { noiseSkipped++; continue; }
        incoming.push(q);
      } catch (err) { failed++; }
    }
    if (!incoming.length) {
      if (!silent) toast('识别到容器但未能解析出题目（请点「诊断」把快照发我）。', true);
      return { ok: false, reason: 'parse-failed' };
    }

    var bank = ensureBank();
    for (var n = 0; n < incoming.length; n++) incoming[n].course = courseName;
    var res = mergeQuestions(bank.questions, incoming);
    bank.questions = res.list;
    saveBank(bank);
    var totalAll = bank.questions.length;
    var mineAll = 0;
    for (var m = 0; m < bank.questions.length; m++) if (bank.questions[m].course === courseName) mineAll++;

    if (!silent) toast('已记录：' + courseName + ' › ' + chapterName + '\n新增 ' + res.added.length + ' 题，变体 ' + res.variants.length + ' 题，重复跳过 ' + res.skipped + ' 题，解析失败 ' + failed + ' 题' + (cardSkipped ? '，跳过答题卡 ' + cardSkipped + ' 块' : '') + (noiseSkipped ? '，跳过标题/无效块 ' + noiseSkipped + ' 块' : '') + '；题库累计 ' + totalAll + ' 题（本课 ' + mineAll + ' 题）');
    refreshStatBar();
    return { ok: true, added: res.added.length, variants: res.variants.length, skipped: res.skipped, failed: failed, totalAll: totalAll };
  }

  /* ---------------- 自动遍历 ----------------
   *  边界：只自动化“浏览 + 读取页面上已展示的题目/解析”，不答题、不提交。
   *  安全：只点击白名单按钮文本，黑名单（提交/交卷/放弃/结束/删除…）一律不点。
   */
  var DANGER_RE = /(提交|交卷|上交|放弃|结束|删除|清空|保存并提交|确认提交|立即提交)/;
  var BTN_VIEW_RE = /(查看解析|查看答案解析|查看答案|试题解析)/;
  var BTN_ENTRY_RE = /(去提升|开始提升|进入提升|提升训练|去练习|开始练习|进入练习|去答题)/;
  var BTN_EXIT_RE = /(退出|返回|关闭|回到视频|返回目录|返回课程)/;
  var BTN_NEXT_RE = /(下一题|下一页|下一道)/;
  var BTN_SECTION_RE = /(下一节|下一个|下一章|下一课|下一讲|下一个视频|继续学习)/;
  var CLOSE_SIG_RE = /(close|exit|back|关闭|退出|返回)/i;

  var PILOT = { running: false, timer: null, steps: 0, maxSteps: 400, logs: [], idle: 0, noop: 0, catalogIdx: null, catalogCache: null, videoDone: false, catalogExhausted: false, switchFails: 0, entered: false, viewClicked: false, boostIdle: 0, stuckSkip: false, lastAction: '', sameAction: 0 };

  /* 自动遍历状态持久化：点“下一节/下一个视频”常会整页刷新，
     刷新后脚本重建，运行状态丢了就会“断”。存下来就能自动续跑。 */
  var K_PILOT = 'zhr.pilot.auto';
  function pilotSaveState() {
    try {
      sSet(K_PILOT, JSON.stringify({
        on: 1,
        idx: (typeof PILOT.catalogIdx === 'number') ? PILOT.catalogIdx : null,
        videoDone: !!PILOT.videoDone,
        t: Date.now(),
        logs: PILOT.logs.slice(-10)
      }));
    } catch (e) { /* ignore */ }
  }

  function pilotLog(msg) {
    PILOT.logs.push(new Date().toLocaleTimeString('zh-CN') + ' ' + msg);
    if (PILOT.logs.length > 30) PILOT.logs.shift();
    var el = uiDoc().getElementById('zhr-pilotlog');
    if (el) el.textContent = PILOT.logs.slice(-3).join('\n');
    if (PILOT.running) pilotSaveState();      /* 心跳：刷新后能续跑 */
  }

  /* 当前文档 + 同源 iframe 文档（“去提升/退出”按钮与题目常在 iframe 里） */
  function pilotDocs() {
    var docs = [document];
    var ifs = document.querySelectorAll('iframe');
    for (var i = 0; i < ifs.length; i++) {
      var d = null;
      try { d = ifs[i].contentDocument; } catch (e) { d = null; }
      if (d && d !== document) docs.push(d);
    }
    return docs;
  }

  function pilotVisible(el) {
    if (!el || !el.getBoundingClientRect) return false;
    var r = el.getBoundingClientRect();
    if (r.width < 2 || r.height < 2) return false;
    var w = (el.ownerDocument && el.ownerDocument.defaultView) || window;
    try {
      var cs = w.getComputedStyle(el);
      if (cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0') return false;
    } catch (e) { /* ignore */ }
    return true;
  }

  /* 模拟真实点击：滚动到可见 + 完整鼠标事件序列（很多页面只认 pointerdown/mousedown） */
  function pilotFireClick(el) {
    try { el.scrollIntoView({ block: 'center', inline: 'nearest' }); } catch (e) { /* ignore */ }
    var w = (el.ownerDocument && el.ownerDocument.defaultView) || window;
    var opt = { bubbles: true, cancelable: true, view: w };
    var seq = ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'];
    for (var i = 0; i < seq.length; i++) {
      var type = seq[i];
      try {
        var Ctor = (type.indexOf('pointer') === 0 && w.PointerEvent) ? w.PointerEvent : w.MouseEvent;
        el.dispatchEvent(new Ctor(type, opt));
      } catch (e2) { /* ignore */ }
    }
  }

  /* 在所有文档里找：可见 + 文案匹配白名单 + 不含危险词 的元素
     （用 textContent 而非 innerText，避免大页面遍历时反复触发重排） */
  function pilotFind(re, maxLen) {
    var docs = pilotDocs();
    for (var d = 0; d < docs.length; d++) {
      var qs;
      try { qs = docs[d].querySelectorAll('button,a,li,span,div,em,i,p,label,[role="button"]'); } catch (e) { continue; }
      var n = Math.min(qs.length, 6000);
      for (var i = 0; i < n; i++) {
        var el = qs[i];
        if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
        if (!pilotVisible(el)) continue;
        var t = normWs(el.textContent || '');
        if (!t || t.length > (maxLen || 16)) continue;
        if (DANGER_RE.test(t)) continue;
        if (!re.test(t)) continue;
        var inner = el.querySelector('button,a,[role="button"]');
        if (inner && inner !== el) continue;
        return { el: el, text: t };
      }
    }
    return null;
  }

  function pilotClick(re, maxLen) {
    var hit = pilotFind(re, maxLen);
    if (!hit) return null;
    pilotFireClick(hit.el);
    PILOT.noop = 0;
    /* 同一按钮反复点但页面没进展 → 不再停止整个遍历，改为“放弃这个视频、跳下一个” */
    if (PILOT.lastAction === hit.text) PILOT.sameAction = (PILOT.sameAction || 0) + 1;
    else { PILOT.lastAction = hit.text; PILOT.sameAction = 1; }
    if (PILOT.sameAction >= 8) {
      PILOT.sameAction = 0;
      PILOT.lastAction = '';
      PILOT.stuckSkip = true;
      pilotLog('「' + hit.text + '」连点 8 次无效 → 放弃这个视频，切下一个');
    }
    return hit.text;
  }

  /* 无文字的图标按钮兔底：aria-label/title/class 里的词边界匹配 close/exit/back（不会被 background 误命中） */
  function pilotClickIcon(re) {
    var wordRe = /(?:^|\s)(close|closed|exit|back|return|goback|cancel)(?:$|\s)/i;
    var docs = pilotDocs();
    for (var d = 0; d < docs.length; d++) {
      var els;
      try { els = docs[d].querySelectorAll('button,a,i,span,div,svg,img,[class],[aria-label],[title]'); } catch (e) { continue; }
      var n = Math.min(els.length, 4000);
      for (var i = 0; i < n; i++) {
        var el = els[i];
        var tag = el.tagName;
        if (tag !== 'BUTTON' && tag !== 'A' && tag !== 'I' && tag !== 'SPAN' && tag !== 'DIV' && tag !== 'SVG' && tag !== 'IMG') continue;
        if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
        if (!pilotVisible(el)) continue;
        var aria = (el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '');
        var sig = (aria + ' ' + clsOf(el)).replace(/[-_]/g, ' ');
        if (!re.test(sig) && !wordRe.test(sig)) continue;
        if (DANGER_RE.test(normWs(el.textContent || ''))) continue;
        pilotFireClick(el);
        return normWs(sig).slice(0, 24);
      }
    }
    return null;
  }

  /* 宽松兜底：带“答案”行、体积不大、非导航/答题卡的块（智慧树「查看解析」后常见） */
  function looseContainers(doc) {
    var all;
    try { all = doc.querySelectorAll('div,li,section'); } catch (e) { return []; }
    var cands = [];
    var n = Math.min(all.length, 8000);
    for (var i = 0; i < n; i++) {
      var el = all[i];
      if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
      var c = clsOf(el), idc = el.id || '';
      if (NAV_CLS_RE.test(c) || NAV_CLS_RE.test(idc)) continue;
      var t = innerTextOf(el);
      if (t.length < 8 || t.length > 6000) continue;
      if (!ANSWER_LINE_RE.test(t)) continue;
      if (CARD_RE.test(normWs(t).slice(0, 80))) continue;
      cands.push(el);
    }
    /* 取最内层（先剔掉只有答案行的碎块，否则会把父级题目块一起排掉） */
    return topLevelContainers(cands);
  }

  /* 只取“屏幕上可见、且解析出真正题目”的容器（含同源 iframe）；
     空壳/标题/答题卡不算，避免在视频页被误当成“题目页” */
  function pilotContainers() {
    var docs = pilotDocs();
    var out = [];
    for (var d = 0; d < docs.length; d++) {
      var all = [];
      var loose = false;
      try { all = findQuestionContainers(docs[d]); } catch (e) { all = []; }
      if (!all.length) { loose = true; all = looseContainers(docs[d]); }
      for (var i = 0; i < all.length; i++) {
        var el = all[i];
        if (!pilotVisible(el)) continue;
        var p = null;
        try { p = parseQuestion(el, ''); } catch (e2) { p = null; }
        if (!p || !p.stem) continue;
        if (looksLikeAnswerCard(el, p)) continue;
        if (looksLikeNoise(el, p)) continue;
        /* 宽松通道额外要求：至少像一道题（有选项 或 判断/填空） */
        if (loose && !(p.options.length >= 2 || p.type === 'judge' || p.type === 'fill')) continue;
        out.push(el);
      }
    }
    return out;
  }

  /* 目录条目：跨同源 iframe 扫描；先走 class 关键词快路径，再走“通用列表”识别（不依赖 class 名、允许折叠项） */
  function pilotCatalogItems() {
    if (PILOT.catalogCache && PILOT.catalogCache.length) {
      var c0 = PILOT.catalogCache[0];
      if (c0 && c0.isConnected !== false) return PILOT.catalogCache;   /* 缓存仍有效 */
    }
    var fastSel = '[class*="catalog" i] li,[class*="catalog" i] a,[class*="chapter" i] li,' +
                  '[class*="menu" i] li,[class*="nav" i] li,[class*="tree" i] li,[class*="tree" i] a,' +
                  '[class*="sidebar" i] li,[class*="sider" i] li,[class*="list" i] li,[class*="list" i] a';
    var docs = pilotDocs();
    var fast = [];
    for (var d0 = 0; d0 < docs.length; d0++) {
      var qs;
      try { qs = docs[d0].querySelectorAll(fastSel); } catch (e0) { continue; }
      var n0 = Math.min(qs.length, 600);
      for (var i = 0; i < n0; i++) {
        var el0 = qs[i];
        if (el0.closest && el0.closest('#zhr-root,#zhr-scout-root')) continue;
        if (!pilotVisible(el0)) continue;
        var t0 = cleanText(el0.innerText || '');
        if (!t0 || t0.length > 60) continue;
        fast.push(el0);
      }
    }
    if (fast.length >= 2) { PILOT.catalogCache = fast; return fast; }

    /* 通用路径：找“在左边、条目多、每条都是短文本”的容器当目录 */
    var best = null, bestScore = null;
    for (var d = 0; d < docs.length; d++) {
      var doc = docs[d];
      var vw = 1200;
      try { if (doc.defaultView) vw = doc.defaultView.innerWidth || vw; } catch (e1) { /* ignore */ }
      var boxes;
      try { boxes = doc.querySelectorAll('ul,ol,div,section,nav,aside'); } catch (e2) { continue; }
      var nb = Math.min(boxes.length, 6000);
      for (var b = 0; b < nb; b++) {
        var box = boxes[b];
        if (box.closest && box.closest('#zhr-root,#zhr-scout-root')) continue;
        var kids = box.children;
        if (!kids || kids.length < 3 || kids.length > 400) continue;
        var rb = box.getBoundingClientRect();
        if (rb.width < 40 || rb.height < 24) continue;
        if (rb.left > vw * 0.5) continue;                    /* 目录一般在左侧 */
        var items = [], hint = 0, visCount = 0;
        for (var k = 0; k < kids.length; k++) {
          var kid = kids[k];
          if (!pilotVisible(kid)) continue;                  /* 折叠隐藏的子项不算失败 */
          var t = cleanText(kid.innerText || '');
          if (!t || t.length > 60) continue;
          visCount++;
          if (CHAP_RE.test(t) || UNIT_RE.test(t) || /必学|知识点|单元|模块|课|节|测验|视频/.test(t)) hint++;
          items.push(kid);
        }
        if (items.length < 3) continue;
        if (visCount < kids.length * 0.5 && hint < 3) continue;   /* 大多是隐藏项且无章节特征 */
        if (hint < 1) continue;
        var score = items.length * 10 + visCount * 5 - rb.left - rb.top / 10;
        if (bestScore === null || score > bestScore) { bestScore = score; best = items; }
      }
    }
    if (best) { PILOT.catalogCache = best; return best; }
    PILOT.catalogCache = null;
    return [];
  }

  /* 找不到目录时的现场诊断信息（写到日志里，便于不改代码定位） */
  function pilotCatalogStats() {
    var docs = [];
    try { docs = pilotDocs(); } catch (e) { docs = []; }
    var boxes = 0, maxKid = 0, maxVis = 0;
    for (var d = 0; d < docs.length; d++) {
      var doc = docs[d];
      var vw = 1200;
      try { if (doc.defaultView) vw = doc.defaultView.innerWidth || vw; } catch (e1) { /* ignore */ }
      var all;
      try { all = doc.querySelectorAll('ul,ol,div,section,nav,aside'); } catch (e2) { continue; }
      var n = Math.min(all.length, 6000);
      for (var i = 0; i < n; i++) {
        var box = all[i];
        if (box.closest && box.closest('#zhr-root,#zhr-scout-root')) continue;
        var kids = box.children;
        if (!kids || kids.length < 3) continue;
        var rb = box.getBoundingClientRect();
        if (rb.width < 40 || rb.height < 24 || rb.left > vw * 0.5) continue;
        var vis = 0;
        for (var k = 0; k < kids.length; k++) {
          if (!pilotVisible(kids[k])) continue;
          var t = cleanText(kids[k].innerText || '');
          if (t && t.length <= 60) vis++;
        }
        if (vis < 3) continue;
        boxes++;
        if (kids.length > maxKid) maxKid = kids.length;
        if (vis > maxVis) maxVis = vis;
      }
    }
    return '左栏候选容器 ' + boxes + '（最多子项 ' + maxKid + '、可见短文本子项 ' + maxVis + '）';
  }

  /* 点目录条目：优先点最里层的可点元素，并补上原生 click / 回车（不同树控件响应方式不一样） */
  function pilotClickItem(el) {
    var target = el, deep = null;
    try { deep = el.querySelector('a[href],button,[role="button"]'); } catch (e) { deep = null; }
    if (deep && pilotVisible(deep) && normWs(deep.innerText || '').length <= 60) target = deep;
    pilotFireClick(target);
    var isBlank = false;
    try { isBlank = /(^|,)\s*_blank/.test(target.getAttribute('target') || ''); } catch (e2) { /* ignore */ }
    if (!isBlank && (target.tagName === 'A' || target.tagName === 'BUTTON')) {
      try { target.click(); } catch (e3) { /* ignore */ }
    } else {
      /* 非链接/按钮：聚焦后补一个回车（部分树控件只认键盘） */
      try {
        if (target.focus) target.focus();
        var w = (target.ownerDocument && target.ownerDocument.defaultView) || window;
        var ev = { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true };
        target.dispatchEvent(new w.KeyboardEvent('keydown', ev));
        target.dispatchEvent(new w.KeyboardEvent('keyup', ev));
      } catch (e4) { /* ignore */ }
    }
    return target;
  }

  /* 顺序切到下一个视频/节点（首次从当前高亮项往下走，之后递增，不重复点同一项） */
  function pilotNextCatalog() {
    var list = pilotCatalogItems();
    if (!list.length) return null;
    var idx = PILOT.catalogIdx;
    if (idx === null || idx === undefined || idx < 0) {
      idx = -1;
      for (var i = 0; i < list.length; i++) { if (isActiveItem(list[i])) { idx = i; break; } }
    }
    var next = idx + 1;
    while (next < list.length && DANGER_RE.test(normWs(list[next].innerText || ''))) next++;
    if (next >= list.length) { PILOT.catalogExhausted = true; pilotLog('目录已到最后一节（共 ' + list.length + ' 项）'); return null; }
    var el = list[next];
    var t = normWs(el.innerText || '');
    PILOT.catalogIdx = next;
    PILOT.catalogExhausted = false;
    PILOT.noop = 0;
    pilotSaveState();          /* 点之前先存状态：点完可能整页刷新 */
    pilotClickItem(el);
    return t.slice(0, 24);
  }

  function pilotStep() {
    if (!PILOT.running) return;
    if (++PILOT.steps > PILOT.maxSteps) { pilotStop('已达步数上限'); return; }

    /* 0) 上个视频点不动 → 借道“切下一个”逻辑跳过，继续往下跑（不停止） */
    if (PILOT.stuckSkip) {
      PILOT.stuckSkip = false;
      PILOT.entered = false;
      PILOT.viewClicked = false;
      PILOT.boostIdle = 0;
      PILOT.videoDone = true;
    }

    /* 1) 本页有可见题目（已展开答案）→ 读取（一个视频只进一次提升） */
    var conts = pilotContainers();
    if (conts.length) {
      PILOT.noop = 0;
      var r = doRecord(true, conts) || {};
      var added = r.added || 0;
      pilotLog('读取本页：新增 ' + added + ' 题');
      PILOT.idle = 0;
      PILOT.viewClicked = false;          /* 下一题可能要再点一次「查看解析」 */
      if (added > 0) {
        PILOT.lastAction = '';
        PILOT.sameAction = 0;
        var nx = pilotClick(BTN_NEXT_RE, 16);
        if (nx) { pilotLog('→ 下一题（' + nx + '）'); return; }
      }
      /* 读完（或全是重复）→ 退出，回到视频页；本视频不再重复进入 */
      var ex = pilotClick(BTN_EXIT_RE, 20) || pilotClickIcon(CLOSE_SIG_RE);
      if (ex) {
        PILOT.videoDone = true;
        PILOT.entered = false;
        PILOT.viewClicked = false;
        PILOT.boostIdle = 0;
        pilotLog('→ 退出（' + ex + '），本视频完成，接着切下一个');
        return;
      }
      PILOT.idle = (PILOT.idle || 0) + 1;
      pilotLog('本页已读，但未找到「退出」按钮（' + PILOT.idle + '/3）');
      if (PILOT.idle >= 3) pilotStop('找不到退出按钮');
      return;
    }

    /* 2) 刚退出上一个提升 → 先切下一个视频/节点，再继续点它的「去提升」 */
    if (PILOT.videoDone) {
      var n0 = pilotNextCatalog();
      if (n0) {
        PILOT.videoDone = false;
        PILOT.idle = 0;
        PILOT.switchFails = 0;
        pilotLog('→ 下一个视频/节点（' + n0 + '）');
        return;
      }
      var ns = pilotClick(BTN_SECTION_RE, 20);          /* 兔底：页面上的“下一节/下一个”按钮 */
      if (ns) {
        PILOT.videoDone = false;
        PILOT.idle = 0;
        PILOT.switchFails = 0;
        pilotLog('→ 下一节（' + ns + '）');
        return;
      }
      if (PILOT.catalogExhausted) { pilotStop('已遍历完课程（最后一个视频已完成）'); return; }
      /* 切不到就继续重试（不停止、也不回同一个视频里反复进出） */
      PILOT.switchFails = (PILOT.switchFails || 0) + 1;
      PILOT.catalogCache = null;                        /* 清缓存，下轮重新识别目录 */
      if (PILOT.switchFails % 3 === 1) {
        pilotLog('切不到下一个（第 ' + PILOT.switchFails + ' 次，' + pilotCatalogStats() + '），继续重试…');
      }
      if (PILOT.switchFails >= 20) { pilotStop('连续 20 次无法切到下一个视频，已停止'); return; }
      return;
    }

    /* 3) 已经进了提升/练习界面 → 必须先点「查看解析」才能看到答案（智慧树顺序） */
    if (PILOT.entered && !PILOT.viewClicked) {
      var v = pilotClick(BTN_VIEW_RE, 20);
      if (v) {
        PILOT.idle = 0;
        PILOT.boostIdle = 0;
        PILOT.viewClicked = true;
        pilotLog('→ 查看解析（' + v + '）');
        return;
      }
    }

    /* 4) 进入提升（未进入时） */
    var en = pilotClick(BTN_ENTRY_RE, 20);
    if (en) {
      PILOT.idle = 0;
      PILOT.entered = true;
      PILOT.viewClicked = false;
      PILOT.boostIdle = 0;
      pilotLog('→ 进入（' + en + '），接着点「查看解析」');
      return;
    }

    /* 4b) 已进入却既没题目也点不到解析 → 试着退出来，换下一个，别卡死在里面 */
    if (PILOT.entered) {
      PILOT.boostIdle = (PILOT.boostIdle || 0) + 1;
      pilotLog('已进入提升，但未发现题目/「查看解析」（' + PILOT.boostIdle + '/4）');
      if (PILOT.boostIdle >= 4) {
        var back = pilotClick(BTN_EXIT_RE, 20) || pilotClickIcon(CLOSE_SIG_RE);
        PILOT.entered = false;
        PILOT.viewClicked = false;
        PILOT.boostIdle = 0;
        if (back) {
          PILOT.videoDone = true;
          pilotLog('→ 退出（' + back + '），换下一个视频');
          return;
        }
      }
    }

    /* 5) 切下一个目录项 */
    var n2 = pilotNextCatalog();
    if (n2) { PILOT.idle = 0; pilotLog('→ 下个视频/节点（' + n2 + '）'); return; }

    /* 6) 真的找不到可操作项：先试着跳过这个视频，连续多次才停 */
    PILOT.idle = (PILOT.idle || 0) + 1;
    PILOT.noop = (PILOT.noop || 0) + 1;
    pilotLog('未找到可操作项（' + PILOT.idle + '/3）');
    if (PILOT.idle >= 3) {
      PILOT.idle = 0;
      PILOT.catalogCache = null;
      if (PILOT.catalogExhausted) { pilotStop('已遍历完课程（最后一个视频已完成）'); return; }
      var sk = pilotNextCatalog();
      if (sk) { pilotLog('→ 跳过，切下一个节点（' + sk + '）'); return; }
    }
    if (PILOT.noop >= 24) pilotStop('连续 24 次找不到可操作项，已停止');
  }

  function pilotStart(resume) {
    if (PILOT.running) return;
    PILOT.running = true;
    PILOT.steps = 0;
    if (!resume) {
      PILOT.logs = [];
      PILOT.catalogIdx = null;
      PILOT.videoDone = false;
    }
    PILOT.idle = 0;
    PILOT.noop = 0;
    PILOT.catalogCache = null;
    PILOT.catalogExhausted = false;
    PILOT.switchFails = 0;
    PILOT.entered = false;
    PILOT.viewClicked = false;
    PILOT.boostIdle = 0;
    PILOT.stuckSkip = false;
    PILOT.lastAction = '';
    PILOT.sameAction = 0;
    pilotLog(resume ? '继续自动遍历（上次因页面跳转中断，已自动续跑）' : '开始自动遍历（不答题、不提交）');
    var b = uiDoc().getElementById('zhr-b6');
    if (b) b.textContent = '⏹ 停止遍历';
    if (PILOT.timer) clearInterval(PILOT.timer);
    PILOT.timer = setInterval(pilotStep, 1600);
    pilotSaveState();
  }
  function pilotStop(reason) {
    PILOT.running = false;
    if (PILOT.timer) { clearInterval(PILOT.timer); PILOT.timer = null; }
    pilotLog('已停止' + (reason ? '：' + reason : ''));
    var b = uiDoc().getElementById('zhr-b6');
    if (b) b.textContent = '🤖 自动遍历';
    try { sSet(K_PILOT, ''); } catch (e) { /* ignore */ }   /* 清掉续跑标记 */
    refreshStatBar();
  }
  function pilotToggle() { if (PILOT.running) pilotStop(''); else pilotStart(); }

  /* ---------------- 页面结构侦察（合并自原侦察脚本） ---------------- */

  var SCOUT_KW = /question|stem|topic|choice|option|answer|correct|right|wrong|result|parse|score|item|daan|chapter|section|unit|答题|题目|答案|解析|正确|错误|得分|测验|测试|判断|选择/i;

  function oneLine(s, n) { return normWs(s).slice(0, n || 80); }

  function scoutText() {
    var out = [];
    out.push('-- 页面结构侦察（candidates）--');
    var SELECTOR = 'div,li,label,p,span,section,tr,td,h1,h2,h3,h4,h5,ul,ol,dl';
    var nodes = document.querySelectorAll(SELECTOR);
    var hits = [];
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
      var id = el.id || '';
      var cls = clsOf(el);
      var txt = (el.innerText || '').trim();
      if (!txt) continue;
      if (SCOUT_KW.test(id) || SCOUT_KW.test(cls) || /^[（(]?[A-Ha-h][)）.、:]/.test(txt)) {
        hits.push('<' + el.tagName.toLowerCase() + '> id="' + id + '" class="' + oneLine(cls, 60) + '" :: ' + oneLine(txt, 70));
        if (hits.length >= 90) break;
      }
    }
    out.push('candidateElements=' + hits.length);
    for (var j = 0; j < hits.length; j++) out.push('  ' + hits[j]);

    var ca = [];
    var scanned = 0;
    var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT);
    while (walker.nextNode() && scanned < 20000 && ca.length < 15) {
      var e2 = walker.currentNode;
      scanned++;
      if (e2.closest && e2.closest('#zhr-root,#zhr-scout-root')) continue;
      var c2 = '';
      if (typeof e2.className === 'string') c2 = e2.className;
      if (c2 && /\b(correct|right|answer|daan|true|green|wrong|red)\w*/i.test(c2)) {
        ca.push(oneLine(c2, 60) + ' | ' + oneLine(e2.innerText || e2.textContent || '', 40));
      }
    }
    if (ca.length) {
      out.push('-- class 含 correct/right/answer/daan/true 等关键字的元素 --');
      for (var k = 0; k < ca.length; k++) out.push('  ' + ca[k]);
    }
    return out.join('\n');
  }

  /* 页面正文文本（排除本脚本与 scout 的面板、以及 script/style 等），用于诊断预览 */
  function pageTextOf() {
    var parts = [];
    var kids = (document.body && document.body.children) || [];
    for (var i = 0; i < kids.length; i++) {
      var k = kids[i];
      if (!k) continue;
      var tag = k.tagName;
      if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'TEMPLATE' || tag === 'NOSCRIPT') continue;
      if (k.id === 'zhr-root' || k.id === 'zhr-scout-root') continue;
      var t = (typeof k.innerText === 'string' && k.innerText) || '';
      if (t) parts.push(t);
    }
    return parts.join('\n');
  }

  function diagnosticText() {
    var out = [];
    out.push('== 智慧树页面诊断快照 ==');
    out.push('version: ' + VERSION);
    out.push('url: ' + location.href);
    out.push('title: ' + normWs(document.title));
    out.push('isTopFrame: ' + (window.top === window) + ' | readyState: ' + document.readyState);
    out.push('courseGuess: ' + guessCourseName());
    out.push('chapterGuess: ' + guessChapterName());
    out.push('radioInputs=' + document.querySelectorAll('input[type=radio]').length +
             ' checkboxInputs=' + document.querySelectorAll('input[type=checkbox]').length +
             ' bodyTextLen=' + normWs(pageTextOf()).length);

    /* 题库识别诊断 */
    var conts = findQuestionContainers(document);
    out.push('-- 题库识别 --');
    out.push('containers=' + conts.length);
    for (var i = 0; i < conts.length && i < 30; i++) {
      var c = conts[i];
      out.push('  [' + i + '] <' + c.tagName.toLowerCase() + '> class="' + clsOf(c) + '" id="' + c.id + '"');
      var parsed = { err: '' };
      try { parsed = parseQuestion(c, ''); } catch (e) { parsed = { err: e.message }; }
      if (parsed.err) out.push('      parseErr: ' + parsed.err);
      else if (looksLikeAnswerCard(c, parsed)) out.push('      → 判定为答题卡/无效块（将被跳过）');
      else if (looksLikeNoise(c, parsed)) out.push('      → 判定为标题/导航/空壳块（将被跳过）');
      else out.push('      stem=' + (parsed.stem || '').slice(0, 60) + ' | type=' + parsed.type + ' | opts=' + parsed.options.length + ' | ans=' + (parsed.answer || '').slice(0, 40));
    }

    /* 页面结构侦察 */
    out.push(scoutText());

    /* 正文预览 */
    out.push('-- body 文本预览（前 1500 字，已排除脚本面板）--');
    out.push(normWs(pageTextOf()).slice(0, 1500));

    /* 运行期错误 */
    if (ERRORS.length) {
      out.push('-- 运行期间捕获到的错误/警告（可能含页面自身脚本）--');
      for (var m = 0; m < ERRORS.length && m < 12; m++) out.push('  ' + ERRORS[m]);
    } else {
      out.push('-- 运行期间未捕获到错误 --');
    }
    return out.join('\n');
  }

  function doExportWord() {
    var courseName = currentCourseGuess();
    var bank = ensureBank();
    if (!bank.questions.length) { toast('题库还没有任何记录，先做一章点「开始记录」。', true); return; }
    var courses = {};
    for (var i = 0; i < bank.questions.length; i++) courses[bank.questions[i].course || '未命名课程'] = 1;
    var nCourse = Object.keys(courses).length;
    var doc = bankToDoc(bank, courseName);
    if (nCourse < 2 && courseName) doc.course = courseName;
    var bytes = buildDocx(doc);
    var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = (nCourse < 2 ? (courseName || '智慧树题库') : ('智慧树题库-' + nCourse + '门课')) + '.docx';
    document.documentElement.appendChild(a);
    a.click();
    setTimeout(function () {
      URL.revokeObjectURL(url);
      if (a.parentNode) a.parentNode.removeChild(a);
    }, 800);
    toast('已生成 Word 文档：' + a.download + '（若浏览器拦截下载请允许）');
  }

  /* 清空整个题库（v1.8.3）：把历史遗留/调试用的题目一次清掉，并阻止旧 v1 分库再被迁入 */
  function doClearBank() {
    var bank = ensureBank();
    var n = bank.questions.length;
    if (!n) { toast('题库已经空了。', true); return; }
    var courses = {};
    for (var i = 0; i < bank.questions.length; i++) courses[bank.questions[i].course || '未命名课程'] = 1;
    var nCourse = Object.keys(courses).length;
    if (!window.confirm('确定清空【整个题库】的 ' + n + ' 题（' + nCourse + ' 个课程标签）？此操作不可恢复。\n如需保留，请先点「📚 已记录」把内容复制备份。')) return;
    if (!window.confirm('再次确认：真的要全部清空吗？')) return;
    saveBank({ questions: [] });
    sSet(K_MIGRATED, '1');            /* 别再自动导入旧 v1 的分库数据（否则清完又回来） */
    sSet(K_UI_COURSE, '');
    sSet(K_UI_CHAPTER, '');
    sSet(K_UI_COURSE_MANUAL, '');
    sSet(K_UI_CHAPTER_MANUAL, '');
    var ci = uiDoc().getElementById('zhr-course-input');
    var hi = uiDoc().getElementById('zhr-chapter-input');
    if (ci) ci.value = '';
    if (hi) hi.value = '';
    refreshStatBar();
    toast('已清空整个题库（删了 ' + n + ' 题）');
  }

  function doClearCourse() {
    var courseName = currentCourseGuess();
    if (!courseName) { toast('未确定课程名。', true); return; }
    var bank = ensureBank();
    var kept = [], mine = 0;
    for (var i = 0; i < bank.questions.length; i++) {
      if (bank.questions[i].course === courseName) mine++;
      else kept.push(bank.questions[i]);
    }
    if (!mine) { toast('当前课程「' + courseName + '」在题库里没有记录。', true); return; }
    if (!window.confirm('确定清空课程「' + courseName + '」的 ' + mine + ' 题记录？此操作不可恢复。\n（题库里其它课程的 ' + kept.length + ' 题不受影响）')) return;
    if (!window.confirm('再次确认：真的要清空这 ' + mine + ' 题吗？')) return;
    bank.questions = kept;
    saveBank(bank);
    /* 面板记忆也一并清除，回到自动识别状态 */
    sSet(K_UI_COURSE, '');
    sSet(K_UI_CHAPTER, '');
    sSet(K_UI_COURSE_MANUAL, '');
    sSet(K_UI_CHAPTER_MANUAL, '');
    var ci = uiDoc().getElementById('zhr-course-input');
    var hi = uiDoc().getElementById('zhr-chapter-input');
    if (ci) ci.value = '';
    if (hi) hi.value = '';
    refreshStatBar();
    toast('已清空课程：' + courseName + '（删了 ' + mine + ' 题，题库还剩 ' + kept.length + ' 题）');
  }

  function addStyleTo(doc, css) {
    try {
      var st = doc.createElement('style');
      st.textContent = css;
      (doc.head || doc.documentElement).appendChild(st);
      return;
    } catch (e) { /* fall through */ }
    try { if (typeof GM_addStyle === 'function') GM_addStyle(css); } catch (e2) { /* ignore */ }
  }

  function initUI() {
    if (window.__zhrV1UI) return;
    if (!document.body) return;
    var udoc = uiDoc();
    var uwin = udoc.defaultView || window;
    /* 顶层去重：多个 frame 都注入时，只在最顶层可达文档里放一份面板 */
    if (uwin.__zhrV1UI) return;
    uwin.__zhrV1UI = true;
    window.__zhrV1UI = true;
    try { if (window.console && console.log) console.log('[ZHR v' + VERSION + '] init on ' + location.href + ' | top=' + (window.top === window)); } catch (e2) { /* ignore */ }

    addStyleTo(udoc,
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
      '#zhr-b5{background:#6c5ce7;color:#fff}' +
      '#zhr-b6{background:#e67e22;color:#fff}' +
      '#zhr-b7{background:#fdecea;color:#a93226}' +
      '#zhr-stat{margin-top:8px;color:#2f6fed;font-size:11px}' +
      '#zhr-pilotlog{margin-top:6px;color:#666;font-size:10px;line-height:1.4;white-space:pre-wrap;max-height:54px;overflow:auto}' +
      '#zhr-tip{margin-top:6px;color:#999;font-size:11px}'
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
    var bList = document.createElement('button');
    bList.id = 'zhr-b5'; bList.textContent = '📚 已记录';
    var bAuto = document.createElement('button');
    bAuto.id = 'zhr-b6'; bAuto.textContent = '🤖 自动遍历';
    var bWipe = document.createElement('button');
    bWipe.id = 'zhr-b7'; bWipe.textContent = '🧹 清空题库';
    btns.appendChild(bRec); btns.appendChild(bWord); btns.appendChild(bDiag); btns.appendChild(bList); btns.appendChild(bAuto); btns.appendChild(bClear); btns.appendChild(bWipe);

    var tip = document.createElement('div');
    tip.id = 'zhr-tip';
    tip.textContent = '自动遍历顺序：去提升 → 查看解析 → 读取 → 退出 → 切下一个视频；页面跳转（整页刷新）会自动续跑，卡住会跳过继续跑（只在你点「⏹ 停止」或课程跑完时才停）。题目全部存在同一个题库里（课程名/章节名只是标签），只有点「清空本课」才会删。';

    var stat = document.createElement('div');
    stat.id = 'zhr-stat';
    var plog = document.createElement('div');
    plog.id = 'zhr-pilotlog';

    body.appendChild(lb1); body.appendChild(courseInput);
    body.appendChild(lb2); body.appendChild(chapterInput);
    body.appendChild(btns); body.appendChild(stat); body.appendChild(plog); body.appendChild(tip);
    root.appendChild(head); root.appendChild(body);
    (udoc.body || udoc.documentElement).appendChild(root);

    /* 自动识别优先：未手改时就用当前页面的识别值（点进新章节会自动更新）；
       用户手改过则尊重手动值（存起来），只有「清空本课」会重置。 */
    var manualCourse = sGet(K_UI_COURSE_MANUAL) === '1';
    var manualChapter = sGet(K_UI_CHAPTER_MANUAL) === '1';
    var savedCourse = sGet(K_UI_COURSE) || '';
    var savedChapter = sGet(K_UI_CHAPTER) || '';
    /* 有记录的记忆课程名优先：避免这次识别波动 → 看上去"记录被清空" */
    var initCourse = savedCourse || guessCourseNameCached();
    courseInput.value = manualCourse ? (savedCourse || initCourse) : initCourse;
    chapterInput.value = manualChapter ? savedChapter : (guessChapterName() || savedChapter);
    if (!manualCourse && courseInput.value) sSet(K_UI_COURSE, courseInput.value);
    if (!manualChapter && chapterInput.value) sSet(K_UI_CHAPTER, chapterInput.value);
    courseInput.addEventListener('input', function () {
      sSet(K_UI_COURSE, normWs(courseInput.value));
      sSet(K_UI_COURSE_MANUAL, '1');
      refreshStatBar();
    });
    chapterInput.addEventListener('input', function () {
      sSet(K_UI_CHAPTER, normWs(chapterInput.value));
      sSet(K_UI_CHAPTER_MANUAL, '1');
    });
    refreshStatBar();

    /* 页内切章节（不重载页面）时，未手改的课程名/章节名自动跟随刷新；
       课程名走锚点规则：识别到的新名字若没记录，不会把当前有记录的库换掉 */
    setInterval(function () {
      try {
        if (document.activeElement === courseInput || document.activeElement === chapterInput) { refreshStatBar(); return; }
        if (sGet(K_UI_COURSE_MANUAL) !== '1') {
          var keep = resolveCourseName();
          if (keep && keep !== courseInput.value) { courseInput.value = keep; sSet(K_UI_COURSE, keep); }
        }
        if (sGet(K_UI_CHAPTER_MANUAL) !== '1') {
          var gh = guessChapterName();
          if (gh && gh !== chapterInput.value) { chapterInput.value = gh; sSet(K_UI_CHAPTER, gh); }
        }
        refreshStatBar();
      } catch (e) { /* ignore */ }
    }, 2500);

    /* 拖动 */
    var dragging = false, dx = 0, dy = 0;
    head.addEventListener('mousedown', function (e) {
      dragging = true;
      dx = e.clientX - root.getBoundingClientRect().left;
      dy = e.clientY - root.getBoundingClientRect().top;
      e.preventDefault();
    });
    udoc.addEventListener('mousemove', function (e) {
      if (!dragging) return;
      root.style.left = Math.max(0, Math.min(window.innerWidth - root.offsetWidth, e.clientX - dx)) + 'px';
      root.style.top = Math.max(0, Math.min(window.innerHeight - root.offsetHeight, e.clientY - dy)) + 'px';
    });
    udoc.addEventListener('mouseup', function () { dragging = false; });

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
    bList.addEventListener('click', function () {
      showModal('本课已记录题目（可复制保存）', '', recordedText());
    });
    bAuto.addEventListener('click', pilotToggle);
    bClear.addEventListener('click', doClearCourse);
    bWipe.addEventListener('click', doClearBank);

    udoc.addEventListener('keydown', function (e) {
      if (e.ctrlKey && e.shiftKey && e.code === 'KeyX') { e.preventDefault(); doRecord(); }
    });

    setTimeout(function () {
      /* 页面晚渲染时补一次识别：仅在为空时填入并记下，不覆盖用户已有内容 */
      if (!courseInput.value) { courseInput.value = guessCourseName(); if (courseInput.value) sSet(K_UI_COURSE, courseInput.value); }
      if (!chapterInput.value) { chapterInput.value = guessChapterName(); if (chapterInput.value) sSet(K_UI_CHAPTER, chapterInput.value); }
      refreshStatBar();
    }, 1500);

    /* 自动续跑：上次正在自动遍历，但点“下一节/下一个视频”把页面整页刷新了 → 接着跑，不要断 */
    try {
      var st = JSON.parse(sGet(K_PILOT) || 'null');
      if (st && st.on) {
        if (st.t && (Date.now() - st.t) < 120000) {
          PILOT.logs = (st.logs || []).slice(-10);
          PILOT.catalogIdx = (typeof st.idx === 'number') ? st.idx : null;
          PILOT.videoDone = !!st.videoDone;
          pilotLog('检测到上次的自动遍历，正在续跑…');
          setTimeout(function () { pilotStart(true); }, 1200);
        } else {
          sSet(K_PILOT, '');        /* 太久之前（>2 分钟）的标记不再自动续跑 */
        }
      }
    } catch (e0) { /* ignore */ }
  }

  return {
    _version: VERSION,
    utf8: utf8, djb2: djb2, normalizeStem: normalizeStem, xmlEscape: xmlEscape,
    crc32: crc32, zipStore: zipStore, buildDocx: buildDocx,
    stemKey: stemKey, optsHash: optsHash, mergeQuestions: mergeQuestions,
    coursesToDoc: coursesToDoc,
    parseLines: parseLines, typeOf: typeOf,
    init: initUI
  };
})();

/* Node 单元测试环境下导出纯函数；浏览器环境自动挂 UI */
if (typeof module !== 'undefined' && module.exports) {
  module.exports = ZHR;
} else {
  /* 学习/考试页常是异步渲染的多层 iframe：轮询等 body 出现再注入，别错过时机 */
  (function boot() {
    if (!document.body) {
      if ((boot.tries = (boot.tries || 0) + 1) < 60) setTimeout(boot, 250);
      return;
    }
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', function () { ZHR.init(); });
    } else {
      ZHR.init();
    }
  })();
}
