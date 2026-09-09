// ==UserScript==
// @name         智慧树课后题记录器 · 结构侦察版
// @namespace    https://dsh.local/zhihuishu-recorder
// @version      0.1.0
// @description  开发期侦察工具：分析智慧树「本次成绩/查看答案解析」页的 DOM 结构（纯本地，不联网、不上传任何数据）。配合「智慧树课后题记录器」主脚本开发使用，开发完成后可卸载。
// @author       you
// @match        https://*.zhihuishu.com/*
// @match        http://*.zhihuishu.com/*
// @run-at       document-idle
// @all-frames   true
// @grant        GM_setClipboard
// @grant        GM_addStyle
// ==/UserScript==

(function () {
  'use strict';
  if (window.__zhrScoutInjected) return;
  if (!document.body || !(document.body.innerText || '').trim()) return;
  window.__zhrScoutInjected = true;

  var KW = /question|stem|topic|choice|option|answer|correct|right|wrong|result|parse|score|item|daan|chapter|section|unit|答题|题目|答案|解析|正确|错误|得分|测验|测试|判断|选择/i;

  function oneLine(s, n) {
    return String(s || '').replace(/\s+/g, ' ').trim().slice(0, n || 80);
  }

  GM_addStyle(
    '#zhr-scout-root{position:fixed;right:14px;bottom:14px;z-index:2147483646;font:12px/1.4 "Segoe UI","Microsoft YaHei",sans-serif}' +
    '#zhr-scout-pill{cursor:pointer;background:#2f6fed;color:#fff;border:0;border-radius:14px;padding:6px 12px;box-shadow:0 2px 8px rgba(0,0,0,.25);opacity:.85}' +
    '#zhr-scout-pill:hover{opacity:1}' +
    '#zhr-scout-mask{position:fixed;left:0;top:0;right:0;bottom:0;z-index:2147483647;background:rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center}' +
    '#zhr-scout-panel{width:min(92vw,900px);max-width:94%;height:min(86vh,700px);background:#fff;color:#222;border-radius:8px;box-shadow:0 8px 40px rgba(0,0,0,.4);display:flex;flex-direction:column;overflow:hidden}' +
    '#zhr-scout-head{padding:8px 14px;background:#f2f5fa;border-bottom:1px solid #ddd;display:flex;align-items:center;gap:8px;flex-wrap:wrap}' +
    '#zhr-scout-head b{font-size:13px}' +
    '#zhr-scout-head .u{margin-left:auto;color:#888;font-size:11px;max-width:45%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}' +
    '#zhr-scout-body{flex:1;overflow:auto;padding:10px 14px}' +
    '#zhr-scout-body textarea{width:100%;height:100%;min-height:420px;border:1px solid #ccc;border-radius:4px;font:11px/1.5 Consolas,monospace;resize:none;padding:6px;box-sizing:border-box}' +
    '#zhr-scout-tip{color:#888;font-size:11px;padding:2px 14px 8px}' +
    '.zhr-btn{border:0;border-radius:6px;padding:6px 12px;cursor:pointer;font-size:12px;background:#2f6fed;color:#fff;margin-right:6px}' +
    '.zhr-btn.ghost{background:#eef1f6;color:#333}'
  );

  var root = document.createElement('div');
  root.id = 'zhr-scout-root';
  var pill = document.createElement('button');
  pill.id = 'zhr-scout-pill';
  pill.type = 'button';
  pill.textContent = '🔍 侦察本页';
  pill.title = '分析当前页面 DOM 结构' + (window.top === window ? '' : '（iframe 内）') + '\n' + location.href;
  root.appendChild(pill);
  document.documentElement.appendChild(root);

  function analyze() {
    var out = [];
    out.push('== 智慧树侦察报告 ==');
    out.push('time: ' + new Date().toLocaleString('zh-CN'));
    out.push('isTopFrame: ' + (window.top === window));
    out.push('url: ' + location.href);
    out.push('title: ' + oneLine(document.title, 160));
    out.push('readyState: ' + document.readyState);
    out.push('radioInputs=' + document.querySelectorAll('input[type=radio]').length +
             ' checkboxInputs=' + document.querySelectorAll('input[type=checkbox]').length +
             ' bodyTextLen=' + (document.body.innerText || '').length);

    var SELECTOR = 'div,li,label,p,span,section,tr,td,h1,h2,h3,h4,h5,ul,ol,dl';
    var nodes = document.querySelectorAll(SELECTOR);
    var hits = [];
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      if (el.closest && el.closest('#zhr-scout-root')) continue;
      var id = el.id || '';
      var cls = '';
      if (typeof el.className === 'string') cls = el.className;
      else if (el.classList && el.classList.length) cls = Array.prototype.join.call(el.classList, ' ');
      var txt = (el.innerText || '').trim();
      if (!txt) continue;
      if (KW.test(id) || KW.test(cls) || /^[（(]?[A-Ha-h][)）.、:]/.test(txt)) {
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
      if (e2.closest && e2.closest('#zhr-scout-root')) continue;
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

    out.push('-- body 文本预览（前 1500 字）--');
    out.push(oneLine(document.body.innerText, 1500));
    return out.join('\n');
  }

  function openModal(report) {
    var mask = document.createElement('div');
    mask.id = 'zhr-scout-mask';
    var panel = document.createElement('div');
    panel.id = 'zhr-scout-panel';

    var head = document.createElement('div');
    head.id = 'zhr-scout-head';
    var b = document.createElement('b');
    b.textContent = '智慧树侦察报告';
    var u = document.createElement('span');
    u.className = 'u';
    u.textContent = location.href;
    head.appendChild(b);
    head.appendChild(u);

    var body = document.createElement('div');
    body.id = 'zhr-scout-body';
    var ta = document.createElement('textarea');
    ta.readOnly = true;
    ta.value = report;
    body.appendChild(ta);

    var tip = document.createElement('div');
    tip.id = 'zhr-scout-tip';
    tip.textContent = '点「复制报告」后直接粘贴发给开发者即可。若此页没有题目内容（内容在另一个弹窗/页面里），请切到那个页面再点「🔍 侦察本页」。';

    var foot = document.createElement('div');
    foot.id = 'zhr-scout-head';
    var bCopy = document.createElement('button');
    bCopy.className = 'zhr-btn';
    bCopy.textContent = '复制报告';
    var bAgain = document.createElement('button');
    bAgain.className = 'zhr-btn ghost';
    bAgain.textContent = '重新扫描';
    var bClose = document.createElement('button');
    bClose.className = 'zhr-btn ghost';
    bClose.textContent = '关闭';
    foot.appendChild(bCopy);
    foot.appendChild(bAgain);
    foot.appendChild(bClose);

    panel.appendChild(head);
    panel.appendChild(body);
    panel.appendChild(tip);
    panel.appendChild(foot);
    mask.appendChild(panel);
    document.documentElement.appendChild(mask);

    function close() {
      if (mask.parentNode) mask.parentNode.removeChild(mask);
    }
    bClose.addEventListener('click', close);
    mask.addEventListener('click', function (ev) { if (ev.target === mask) close(); });
    bAgain.addEventListener('click', function () { ta.value = analyze(); });
    bCopy.addEventListener('click', function () {
      try {
        GM_setClipboard(ta.value);
        bCopy.textContent = '已复制 ✓';
      } catch (err) {
        ta.select();
        try { document.execCommand('copy'); } catch (e2) { /* ignore */ }
        bCopy.textContent = '请手动 Ctrl+C';
      }
      setTimeout(function () { bCopy.textContent = '复制报告'; }, 1500);
    });
  }

  pill.addEventListener('click', function () {
    root.style.visibility = 'hidden';
    var report = '';
    try { report = analyze(); } catch (err) { report = '分析出错：' + err.message + '\n' + err.stack; }
    root.style.visibility = 'visible';
    openModal(report);
  });
})();
