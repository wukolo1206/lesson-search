/* 課堂講義：主題闖關畫面（左側欄、闖關地圖、單元關卡、完成關）
 * 依賴 class.html 的全域：ALL、SHOWN、TAUGHT、applyFilter、unitHtml、renderResults、saveTaught、esc、$
 * 純邏輯在 class-topics.js（ClassTopicsCore）。
 */
(function (root) {
  'use strict';
  var C = root.ClassTopicsCore;
  var TOPICS = [], STATE = { mode: 'home' }, inited = false;
  var LAST_KEY = 'ziyin_topic_last_v1';

  function h(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function island(cat) { return TOPICS.filter(function (t) { return t.cat === cat; })[0] || null; }
  function unitOf(st) { var isl = island(st.cat); return isl ? isl.units.filter(function (u) { return u.grade === st.grade; })[0] || null : null; }
  function taughtCount(qs) { return qs.filter(function (q) { return TAUGHT.has(String(q['題目ID'] || q['題號'])); }).length; }
  function gradeShort(g) { return g.replace('年級', ''); }

  function saveLast() {
    if (STATE.mode !== 'unit') return;
    try { root.localStorage.setItem(LAST_KEY, JSON.stringify(STATE)); } catch (e) { /* 存不了沒關係 */ }
  }
  function loadLast() {
    try { var v = JSON.parse(root.localStorage.getItem(LAST_KEY) || 'null'); return v && v.mode === 'unit' && unitOf(v) ? v : null; } catch (e) { return null; }
  }

  /* 切換畫面：寫進網址（可上一頁、重新整理）、套用篩選、重畫 */
  function go(st, push) {
    if (st.mode === 'unit' && !unitOf(st)) st = { mode: 'home' };
    if (st.mode === 'unit') {
      var n = unitOf(st).questions.length;
      st.idx = Math.max(0, Math.min(st.idx || 0, n));
    }
    var prevMode = STATE.mode;
    STATE = st;
    var q = C.stateQuery(st) || location.pathname.split('/').pop();
    try {
      if (push) history.pushState(st, '', q); else history.replaceState(st, '', q);
    } catch (e) { /* file:// 等情況 */ }
    document.body.classList.toggle('topic-all', st.mode === 'all');
    document.body.classList.toggle('topic-mode', st.mode !== 'all');
    document.body.classList.remove('topic-nav-open');
    if (st.mode === 'unit') {
      ['fYear', 'fDiff', 'fTaught'].forEach(function (id) { $(id).value = ''; });
      $('fKaoguti').checked = false; $('fKw').value = '';
      $('fCat').value = st.cat; $('fGrade').value = st.grade;
      saveLast();
    } else if (st.mode === 'all' && prevMode !== 'all') {
      ['fGrade', 'fYear', 'fCat', 'fDiff', 'fTaught'].forEach(function (id) { $(id).value = ''; });
      $('fKaoguti').checked = false; $('fKw').value = '';
    }
    renderNav();
    applyFilter();                   // 會呼叫 renderList → render()
    if (push) window.scrollTo(0, 0);
  }

  /* ── 左側欄 ── */
  function renderNav() {
    var nav = $('topicNav');
    if (!nav) return;
    var html = '<button class="tn-home' + (STATE.mode === 'home' ? ' on' : '') + '" data-tn="home">🗺️ 闖關地圖</button>';
    TOPICS.forEach(function (isl) {
      var open = STATE.mode === 'unit' && STATE.cat === isl.cat;
      html += '<div class="tn-isl' + (open ? ' open' : '') + '" style="--cat:' + isl.color + '">' +
        '<button class="tn-isl-head" data-tn="isl" data-cat="' + h(isl.cat) + '"><span class="tn-ico">' + isl.icon + '</span><span class="tn-name">' + h(isl.cat) + '</span><span class="tn-cnt">' + isl.total + '</span></button>' +
        '<div class="tn-units">' + isl.units.map(function (u) {
          var on = open && STATE.grade === u.grade;
          var t = taughtCount(u.questions);
          return '<button class="tn-unit' + (on ? ' on' : '') + '" data-tn="unit" data-cat="' + h(isl.cat) + '" data-grade="' + h(u.grade) + '">' +
            h(u.grade) + '<small>' + u.questions.length + ' 題' + (t ? '・已教 ' + t : '') + '</small></button>';
        }).join('') + '</div></div>';
    });
    html += '<button class="tn-all' + (STATE.mode === 'all' ? ' on' : '') + '" data-tn="all">📋 全部題目（篩選、列印、快速測驗）</button>';
    nav.innerHTML = html;
  }

  /* ── 闖關地圖 ── */
  function homeHtml() {
    var last = loadLast();
    var lastBtn = '';
    if (last) {
      var lu = unitOf(last), lq = lu.questions[last.idx];
      lastBtn = '<button class="tp-continue" data-tn="unit" data-cat="' + h(last.cat) + '" data-grade="' + h(last.grade) + '" data-idx="' + last.idx + '">▶ 繼續上次：' +
        h(last.cat) + '・' + h(last.grade) + (lq ? '　第 ' + (last.idx + 1) + ' 關（' + h(C.shortLabel(lq)) + '）' : '　完成關') + '</button>';
    }
    return '<div class="tp-hero"><div><h2>字音形・闖關地圖</h2><p>選一座島、選年級出發：每一題是一關，看比較表 → 一起動腦 → 換個句子 → 寫寫看。</p>' + lastBtn + '</div></div>' +
      '<div class="tp-map">' + TOPICS.map(function (isl) {
        return '<div class="tp-isl" style="--cat:' + isl.color + '"><div class="tp-isl-head"><span class="tp-ico">' + isl.icon + '</span><div><b>' + h(isl.cat) + '</b><small>' + h(isl.tag) + '</small></div></div>' +
          isl.units.map(function (u) {
            var t = taughtCount(u.questions), n = u.questions.length;
            return '<button class="tp-unit" data-tn="unit" data-cat="' + h(isl.cat) + '" data-grade="' + h(u.grade) + '"><span class="tp-g">' + gradeShort(u.grade) + '</span><span class="tp-un">' + h(u.grade) + '</span>' +
              '<span class="tp-prog"><i style="width:' + Math.round(t / n * 100) + '%"></i></span><small>' + (t ? t + '/' : '') + n + ' 題</small></button>';
          }).join('') + '</div>';
      }).join('') + '</div>';
  }

  /* ── 單元：橫幅＋關卡＋這一關 ── */
  function unitHtmlTop(isl, u) {
    var n = u.questions.length, idx = STATE.idx;
    var steps = u.questions.map(function (q, i) {
      var done = TAUGHT.has(String(q['題目ID'] || q['題號']));
      return '<button class="tp-step' + (i === idx ? ' on' : '') + (done ? ' done' : '') + '" data-tn="step" data-idx="' + i + '" title="' + h(q['題目重點'] || '') + '">' +
        '<span class="tp-dot">' + (done ? '✓' : (i + 1)) + '</span><span class="tp-lab">' + h(C.shortLabel(q)) + '</span></button>';
    }).join('') + '<button class="tp-step fin' + (idx === n ? ' on' : '') + '" data-tn="step" data-idx="' + n + '"><span class="tp-dot">🏁</span><span class="tp-lab">完成</span></button>';
    return '<div class="tp-banner" style="--cat:' + isl.color + '"><span class="tp-ico big">' + isl.icon + '</span><div class="tp-btext">' +
      '<span class="tp-pill">' + h(isl.cat) + '</span><h2>' + h(isl.cat) + '・' + h(u.grade) + '</h2><p>' + h(isl.tag) + '</p>' +
      '<button class="tp-back" data-tn="home">← 回闖關地圖</button><button class="tp-navbtn" data-tn="navtoggle">☰ 主題</button></div></div>' +
      '<div class="tp-steps" style="--cat:' + isl.color + '">' + steps + '</div>';
  }

  function stepFoot(n) {
    var idx = STATE.idx;
    return '<div class="tp-foot">' +
      (idx > 0 ? '<button class="tp-prev" data-tn="step" data-idx="' + (idx - 1) + '">← 上一關</button>' : '<span></span>') +
      (idx < n ? '<button class="tp-next" data-tn="step" data-idx="' + (idx + 1) + '">' + (idx === n - 1 ? '完成這個單元 →' : '下一關 →') + '</button>' : '') + '</div>';
  }

  function doneHtml(isl, u) {
    var rows = u.questions.map(function (q, i) {
      var id = String(q['題目ID'] || q['題號']);
      var line = root.ClassResults ? ClassResults.line(id) : '';
      return '<li><button class="tp-dlink" data-tn="step" data-idx="' + i + '">' + (TAUGHT.has(id) ? '✅' : '○') + ' 第 ' + (i + 1) + ' 關　' + h(C.shortLabel(q)) +
        '<small>' + h(q['年度']) + ' 年第 ' + h(q['題號']) + ' 題</small></button>' + (line ? '<div class="tp-dres">📊 ' + h(line) + '</div>' : '') + '</li>';
    }).join('');
    var t = taughtCount(u.questions), n = u.questions.length;
    return '<div class="tp-done"><h3>🏁 ' + h(isl.cat) + '・' + h(u.grade) + '：已教 ' + t + ' / ' + n + ' 題</h3>' +
      '<div class="tp-actions">' +
      '<button class="tp-act" data-tn="markall">✅ 這個單元全部標記已教</button>' +
      '<a class="tp-act" href="words.html?taught=1">📝 產生複習短文（已教過的題目）</a>' +
      '<button class="tp-act" data-tn="proj">▶ 投影這個單元</button>' +
      '<button class="tp-act" data-tn="print">🖨️ 列印這個單元講義</button></div>' +
      '<p class="tp-note">隔幾天後回到這個單元，打開題目的「一起動腦」做「換個句子」，📊 會顯示「第一次 → 隔天」的對照。</p>' +
      '<ul class="tp-dlist">' + rows + '</ul></div>';
  }

  /* class.html 的 renderList 先問這裡：回傳 true 表示已畫好（闖關地圖／單元），false 用原本的全部列表 */
  function render() {
    if (!inited || STATE.mode === 'all') return false;
    var grid = $('grid');
    if (STATE.mode === 'home') { grid.innerHTML = homeHtml(); return true; }
    var isl = island(STATE.cat), u = unitOf(STATE);
    if (!u) { go({ mode: 'home' }, false); return true; }
    var n = u.questions.length;
    var body = STATE.idx < n ? unitHtml(u.questions[STATE.idx]) : doneHtml(isl, u);
    grid.innerHTML = unitHtmlTop(isl, u) + '<div class="tp-card">' + body + '</div>' + stepFoot(n);
    if (root.renderResults) renderResults();
    var on = grid.querySelector('.tp-step.on');
    if (on && on.scrollIntoView) { try { on.scrollIntoView({ block: 'nearest', inline: 'center' }); } catch (e) { /* 舊瀏覽器 */ } }
    return true;
  }

  function onClick(e) {
    var b = e.target.closest('[data-tn]');
    if (!b) return;
    var a = b.getAttribute('data-tn');
    if (a === 'home') return go({ mode: 'home' }, true);
    if (a === 'all') return go({ mode: 'all' }, true);
    if (a === 'navtoggle') { document.body.classList.toggle('topic-nav-open'); return; }
    if (a === 'isl') {
      var isl = island(b.getAttribute('data-cat'));
      if (isl) go({ mode: 'unit', cat: isl.cat, grade: isl.units[0].grade, idx: 0 }, true);
      return;
    }
    if (a === 'unit') return go({ mode: 'unit', cat: b.getAttribute('data-cat'), grade: b.getAttribute('data-grade'), idx: Number(b.getAttribute('data-idx')) || 0 }, true);
    if (a === 'step') return go({ mode: 'unit', cat: STATE.cat, grade: STATE.grade, idx: Number(b.getAttribute('data-idx')) }, true);
    if (a === 'markall') {
      var u = unitOf(STATE);
      if (!u || !confirm('把「' + STATE.cat + '・' + STATE.grade + '」' + u.questions.length + ' 題全部標記為已教過？')) return;
      u.questions.forEach(function (q) { TAUGHT.add(String(q['題目ID'] || q['題號'])); });
      saveTaught(); renderNav(); render();
      return;
    }
    if (a === 'proj') { if (root.enterProjector) enterProjector(); return; }
    if (a === 'print') return printUnit();
  }

  /* 列印這個單元：暫時把單元所有題目卡畫出來再印，印完回到原畫面 */
  function printUnit() {
    var grid = $('grid');
    grid.innerHTML = SHOWN.map(function (q) { return unitHtml(q); }).join('');
    if (root.renderResults) renderResults();
    document.body.classList.add('topic-printing');
    var restore = function () { document.body.classList.remove('topic-printing'); render(); };
    window.addEventListener('afterprint', restore, { once: true });
    window.print();
  }

  /* 從資料載入完成後呼叫：建島、讀網址（含 ?guided= ?q= 直接進那題的單元） */
  function init() {
    TOPICS = C.buildTopics(ALL);
    inited = true;
    var st = C.parseState(location.search), p;
    try { p = new URLSearchParams(location.search); } catch (e) { p = null; }
    var deep = p && (p.get('guided') || p.get('q'));
    if (deep) {
      var loc = C.locate(TOPICS, deep);
      if (loc) st = { mode: 'unit', cat: loc.cat, grade: loc.grade, idx: loc.idx };
    }
    if (!document.getElementById('topicNav')) return false;
    var bar = document.querySelector('.action-bar');
    if (bar && !bar.querySelector('.tn-backmap')) bar.insertAdjacentHTML('afterbegin', '<button class="btn btn-reset tn-backmap" data-tn="home">🗺️ 闖關地圖</button>');
    document.addEventListener('click', onClick);
    window.addEventListener('popstate', function (e) { go(e.state || C.parseState(location.search), false); });
    // 已教過（另一個分頁或卡片上的按鈕）變了 → 側欄與關卡打勾更新
    window.addEventListener('storage', function (e) { if (e.key === 'taught_questions_v1' && STATE.mode !== 'all') renderNav(); });
    // 保留 ?guided= 讓 openFromUrl 接著打開一起動腦
    go(st, false);
    if (deep) { try { history.replaceState(STATE, '', C.stateQuery(STATE) + (p.get('guided') ? '&guided=' + encodeURIComponent(deep) : '')); } catch (e) { /* 略 */ } }
    return true;
  }

  root.ClassTopics = { init: init, render: render, go: go, state: function () { return STATE; }, refreshNav: function () { if (inited) renderNav(); }, printUnit: printUnit };
})(typeof window !== 'undefined' ? window : globalThis);
