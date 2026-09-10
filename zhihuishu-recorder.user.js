// ==UserScript==
// @name         智慧树课后题记录器 v1
// @namespace    https://dsh.local/zhihuishu-recorder
// @version      1.7.1
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
  var VERSION = '1.7.1';

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

  /* 从一个题目容器的行文本里解析题干/选项/答案（行文本模型，兼容多数布局） */
  function parseLines(containerText) {
    var lines = splitLines(containerText);
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
      if (/^(我的答案|您的答案|回答正确|回答错误|答对|答错|已作答|未作答|答案解析|解析|考查知识点|知识点)/.test(L)) continue;
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
        options.push({ letter: letter, text: otext });
        continue;
      }
      /* 无字母前缀的判断题短选项（对/错/正确/错误/是/否/√/×） */
      if (/^(对|错|正确|错误|是|否|√|×)$/.test(L)) { bareOpts.push(L); continue; }
      if (ansLineIdx === -1) stemLines.push(L);
    }
    var stem = cleanText(stemLines.join(' '));
    /* 去掉开头的题型标签（单选 题 / 判断 题 ...），其后的题号由 normalizeStem 剥离 */
    stem = stem.replace(/^(单选|多选|判断|填空|简答|不定项|单项选择|多项选择)\s*题?\s*(?=\d)/, '');
    stem = normalizeStem(stem);
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
    var scopes = document.querySelectorAll(sel);
    var best = null;
    for (var i = 0; i < scopes.length; i++) {
      var sc = scopes[i];
      if (sc.closest && sc.closest('#zhr-root,#zhr-scout-root')) continue;
      if (/\b(question|timu|exam|test|answer|result|choice)\b/i.test(clsOf(sc))) continue;  /* 别把题目区当目录 */
      if (!visOk(sc)) continue;
      var els = sc.querySelectorAll('a,li');
      var items = [];
      for (var j = 0; j < els.length && items.length < 200; j++) {
        var el = els[j];
        if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
        if (!visOk(el)) continue;
        var t = cleanText(el.innerText || '');
        if (!t || t.length > 40) continue;
        items.push({ el: el, t: t });
      }
      if (items.length >= 3 && (!best || items.length > best.length)) best = items;
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
    var title = normWs(document.title || '');
    /* 1) 标题："课程名 - 智慧树" / "智慧树 - 课程名" */
    var m = title.match(/^(.*?)\s*[-_—|｜]\s*(智慧树|知到)/);
    if (m && m[1] && !/登录|首页|智慧树|知到/.test(m[1])) return cleanText(m[1]);
    var m2 = title.match(/(?:智慧树|知到)\s*[-_—|｜]\s*(.+)$/);
    if (m2 && m2[1]) return cleanText(m2[1]);
    /* 2) 明确的“课程名”元素 / 面包屑 */
    var sels = ['[class*="courseName" i]', '[class*="course-name" i]', '[class*="course_title" i]',
                '[class*="courseTitle" i]', '[class*="course" i] h1', '[class*="course" i] h2',
                '[class*="course" i] h3', '[class*="breadcrumb" i] a', '[class*="crumb" i] a',
                '[class*="breadcrumb" i] span', 'h1', 'h2'];
    for (var i = 0; i < sels.length; i++) {
      var els = document.querySelectorAll(sels[i]);
      for (var j = 0; j < els.length && j < 10; j++) {
        if (els[j].closest && els[j].closest('#zhr-root,#zhr-scout-root')) continue;
        var t = cleanText(els[j].innerText || els[j].textContent || '');
        if (!t || t.length < 2 || t.length > 40) continue;
        if (/^(首页|主页|登录|智慧树|知到|我的|课程中心|学习中心|个人中心)$/.test(t)) continue;
        if (CHAP_RE.test(t)) continue;               /* 像章节名的不当课程名 */
        return t;
      }
    }
    /* 3) 退一步用标题（去掉站点名） */
    if (title && title.length <= 40 && !/智慧树|知到|登录/.test(title)) return cleanText(title);
    return '';
  }

  function guessChapterName() {
    var items = collectCatalogItems();
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
    var els = document.querySelectorAll('h1,h2,h3,h4,.chapter,.section,.unit,li,span,div,b,strong,p');
    var best = '', bestEl = null;
    for (var n = 0; n < els.length && n < 300; n++) {
      var el = els[n];
      if (el.closest && el.closest('#zhr-root,#zhr-scout-root')) continue;
      var t = cleanText(el.innerText || '');
      if (!t || t.length > 40) continue;
      if (re.test(t)) {
        if (!best || t.length < best.length) { best = t; bestEl = el; }
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

  /* ---------------- 核心动作 ---------------- */

  /* 面板状态记忆：课程名/章节名持久化，切换界面（页面重载）后自动恢复；
     只有点「清空本课」才清除。 */
  function statText() {
    var c = loadCourse(currentCourseGuess());
    if (!c || !c.questions.length) return '本课暂无记录（做完一章在解析页点「开始记录」）';
    var chs = {};
    c.questions.forEach(function (q) { chs[q.chapter || '未命名章节'] = 1; });
    return '本课已记录 ' + c.questions.length + ' 题，覆盖 ' + Object.keys(chs).length + ' 个章节';
  }
  function refreshStatBar() {
    var el = uiDoc().getElementById('zhr-stat');
    if (el) el.textContent = statText();
  }
  function recordedText() {
    var c = loadCourse(currentCourseGuess());
    if (!c || !c.questions.length) return '（本课暂无记录）\n做完一章、进入「本次成绩/查看解析」页后点「▶ 开始记录」即可累积。';
    var doc = coursesToDoc(c);
    var out = [];
    out.push('课程：' + (c.name || '未命名课程') + '　合计 ' + c.questions.length + ' 题');
    doc.sections.forEach(function (sec) {
      out.push('');
      out.push('【' + sec.name + '】');
      sec.questions.forEach(function (q, i) {
        out.push((i + 1) + '. ' + q.stem);
        if (q.options && q.options.length) out.push('   ' + optionsLine(q));
        out.push('   答案：' + (q.answer || '未记录'));
      });
    });
    return out.join('\n');
  }

  function currentCourseGuess() {
    var inp = uiDoc().getElementById('zhr-course-input');
    return normWs(inp ? inp.value : '') || guessCourseName();
  }

  function doRecord(silent, presetContainers) {
    var containers = presetContainers || findQuestionContainers(document);
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

    var course = loadCourse(courseName) || { name: courseName, questions: [] };
    var res = mergeQuestions(course.questions, incoming);
    course.questions = res.list;
    saveCourse(courseName, course);
    var totalAll = course.questions.length;

    if (!silent) toast('已记录：' + courseName + ' › ' + chapterName + '\n新增 ' + res.added.length + ' 题，变体 ' + res.variants.length + ' 题，重复跳过 ' + res.skipped + ' 题，解析失败 ' + failed + ' 题' + (cardSkipped ? '，跳过答题卡 ' + cardSkipped + ' 块' : '') + (noiseSkipped ? '，跳过标题/无效块 ' + noiseSkipped + ' 块' : '') + '；该课累计 ' + totalAll + ' 题');
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
  var CLOSE_SIG_RE = /(close|exit|back|关闭|退出|返回)/i;

  var PILOT = { running: false, timer: null, steps: 0, maxSteps: 400, logs: [], idle: 0, catalogIdx: null, emptyInThisVideo: false, emptyRounds: 0, catalogExhausted: false, lastAction: '', sameAction: 0 };

  function pilotLog(msg) {
    PILOT.logs.push(new Date().toLocaleTimeString('zh-CN') + ' ' + msg);
    if (PILOT.logs.length > 30) PILOT.logs.shift();
    var el = uiDoc().getElementById('zhr-pilotlog');
    if (el) el.textContent = PILOT.logs.slice(-3).join('\n');
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
    /* 同一按钮反复点但页面没进展 → 自动停，避免无效循环 */
    if (PILOT.lastAction === hit.text) PILOT.sameAction = (PILOT.sameAction || 0) + 1;
    else { PILOT.lastAction = hit.text; PILOT.sameAction = 1; }
    if (PILOT.sameAction >= 8) pilotStop('同一按钮连续 8 次无效：' + hit.text);
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

  /* 只取“屏幕上可见、且解析出真正题目”的容器（含同源 iframe）；
     空壳/标题/答题卡不算，避免在视频页被误当成“题目页” */
  function pilotContainers() {
    var docs = pilotDocs();
    var out = [];
    for (var d = 0; d < docs.length; d++) {
      var all = [];
      try { all = findQuestionContainers(docs[d]); } catch (e) { all = []; }
      for (var i = 0; i < all.length; i++) {
        var el = all[i];
        if (!pilotVisible(el)) continue;
        var p = null;
        try { p = parseQuestion(el, ''); } catch (e2) { p = null; }
        if (!p || !p.stem) continue;
        if (looksLikeAnswerCard(el, p)) continue;
        if (looksLikeNoise(el, p)) continue;
        out.push(el);
      }
    }
    return out;
  }

  function pilotCatalogItems() {
    var sel = '[class*="catalog" i] li,[class*="chapter" i] li,[class*="section" i] li,' +
              '[class*="menu" i] li,[class*="nav" i] li,[class*="list" i] li,' +
              '[class*="catalog" i] a,[class*="chapter" i] a,[class*="menu" i] a,[class*="nav" i] a';
    var items = document.querySelectorAll(sel);
    var list = [];
    for (var i = 0; i < items.length; i++) if (pilotVisible(items[i])) list.push(items[i]);
    return list;
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
    if (next >= list.length) { PILOT.catalogExhausted = true; pilotLog('目录已到最后一节'); return null; }
    var el = list[next];
    var t = normWs(el.innerText || '');
    if (DANGER_RE.test(t)) { PILOT.catalogIdx = next; return null; }
    PILOT.catalogIdx = next;
    PILOT.catalogExhausted = false;
    pilotFireClick(el);
    return t.slice(0, 24);
  }

  function pilotStep() {
    if (!PILOT.running) return;
    if (++PILOT.steps > PILOT.maxSteps) { pilotStop('已达步数上限'); return; }

    /* 1) 本页有可见题目 → 读取 */
    var conts = pilotContainers();
    if (conts.length) {
      var r = doRecord(true, conts) || {};
      var added = r.added || 0;
      pilotLog('读取本页：新增 ' + added + ' 题');
      if (added > 0) {
        PILOT.emptyInThisVideo = false;
        PILOT.emptyRounds = 0;
        PILOT.idle = 0;
        PILOT.lastAction = '';
        PILOT.sameAction = 0;
        var nx = pilotClick(BTN_NEXT_RE, 16);
        if (nx) { pilotLog('→ 下一题（' + nx + '）'); return; }
      } else {
        /* 新增 0 题：本视频的提升已读完，退出后不再重复进入 */
        PILOT.emptyInThisVideo = true;
        PILOT.emptyRounds = (PILOT.emptyRounds || 0) + 1;
        if (PILOT.emptyRounds >= 3) { pilotStop('连续 3 个提升均为空，已停止'); return; }
      }
      var ex = pilotClick(BTN_EXIT_RE, 20) || pilotClickIcon(CLOSE_SIG_RE);
      if (ex) { PILOT.idle = 0; pilotLog('→ 退出（' + ex + '）'); return; }
      PILOT.idle = (PILOT.idle || 0) + 1;
      pilotLog('本页已读，但未找到「退出」按钮（' + PILOT.idle + '/3）');
      if (PILOT.idle >= 3) pilotStop('找不到退出按钮');
      return;
    }

    /* 2) 刚读完一个视频的提升 → 直接切下一个视频/节点 */
    if (PILOT.emptyInThisVideo) {
      PILOT.emptyInThisVideo = false;
      var n0 = pilotNextCatalog();
      if (n0) { PILOT.idle = 0; pilotLog('→ 下个视频/节点（' + n0 + '）'); return; }
      if (PILOT.catalogExhausted) { pilotStop('已遍历到最后一个视频'); return; }
    }

    /* 3) 打开解析 */
    var v = pilotClick(BTN_VIEW_RE, 16);
    if (v) { PILOT.idle = 0; pilotLog('→ 查看解析（' + v + '）'); return; }

    /* 4) 进入提升 */
    var en = pilotClick(BTN_ENTRY_RE, 16);
    if (en) { PILOT.idle = 0; pilotLog('→ 进入（' + en + '）'); return; }

    /* 5) 切下一个目录项 */
    var n2 = pilotNextCatalog();
    if (n2) { PILOT.idle = 0; pilotLog('→ 下个视频/节点（' + n2 + '）'); return; }

    PILOT.idle = (PILOT.idle || 0) + 1;
    pilotLog('未找到可操作项（' + PILOT.idle + '/3）');
    if (PILOT.idle >= 3) pilotStop('连续未找到可操作项');
  }

  function pilotStart() {
    if (PILOT.running) return;
    PILOT.running = true;
    PILOT.steps = 0;
    PILOT.logs = [];
    PILOT.idle = 0;
    PILOT.catalogIdx = null;
    PILOT.emptyInThisVideo = false;
    PILOT.emptyRounds = 0;
    PILOT.catalogExhausted = false;
    PILOT.lastAction = '';
    PILOT.sameAction = 0;
    pilotLog('开始自动遍历（不答题、不提交）');
    var b = uiDoc().getElementById('zhr-b6');
    if (b) b.textContent = '⏹ 停止遍历';
    PILOT.timer = setInterval(pilotStep, 1600);
  }
  function pilotStop(reason) {
    PILOT.running = false;
    if (PILOT.timer) { clearInterval(PILOT.timer); PILOT.timer = null; }
    pilotLog('已停止' + (reason ? '：' + reason : ''));
    var b = uiDoc().getElementById('zhr-b6');
    if (b) b.textContent = '🤖 自动遍历';
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
    sSet(courseKeyFor(courseName), 'null');
    var list = loadCourseList().filter(function (c) { return c.name !== courseName; });
    sSet(K_INDEX, JSON.stringify(list));
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
    toast('已清空课程：' + courseName);
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
    btns.appendChild(bRec); btns.appendChild(bWord); btns.appendChild(bDiag); btns.appendChild(bList); btns.appendChild(bAuto); btns.appendChild(bClear);

    var tip = document.createElement('div');
    tip.id = 'zhr-tip';
    tip.textContent = '课程名/章节名自动识别（点进章节会更新为该章+第一个单元名），手改后以你改的为准；「清空本课」恢复自动。用法：把解析页点「开始记录」，快捷键 Ctrl+Shift+X。';

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
    courseInput.value = manualCourse ? savedCourse : (guessCourseName() || savedCourse);
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

    /* 页内切章节（不重载页面）时，未手改的课程名/章节名自动跟随刷新 */
    setInterval(function () {
      try {
        if (document.activeElement === courseInput || document.activeElement === chapterInput) return;
        if (sGet(K_UI_COURSE_MANUAL) !== '1') {
          var gc = guessCourseName();
          if (gc && gc !== courseInput.value) { courseInput.value = gc; sSet(K_UI_COURSE, gc); }
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

    udoc.addEventListener('keydown', function (e) {
      if (e.ctrlKey && e.shiftKey && e.code === 'KeyX') { e.preventDefault(); doRecord(); }
    });

    setTimeout(function () {
      /* 页面晚渲染时补一次识别：仅在为空时填入并记下，不覆盖用户已有内容 */
      if (!courseInput.value) { courseInput.value = guessCourseName(); if (courseInput.value) sSet(K_UI_COURSE, courseInput.value); }
      if (!chapterInput.value) { chapterInput.value = guessChapterName(); if (chapterInput.value) sSet(K_UI_CHAPTER, chapterInput.value); }
      refreshStatBar();
    }, 1500);
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
