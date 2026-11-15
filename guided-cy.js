/* 一起動腦：詞義、成語（依學扶教材 4-1-3《語詞一家人（二）》：
 * 注意前後文 → 換詞讀一讀 → 查字典 → 依意思分一分）
 * 先猜 → 找線索 → 換詞讀一讀 → 查字典 → 想法改了嗎 → 原題挑戰
 * 資料：guided-cy-data.js（scripts/build_guided_cy_data.py 產生；線索經老師確認）
 * 設計文件：字音形互動引導設計_教材依據.md（題型 D）
 * 樣式沿用 guided-py.css 的 gp-*。
 */
(function (root) {
  'use strict';

  /* ═══ 核心（可在 Node 測試） ═══════════════════════════ */
  var DATA = null;
  function data() {
    if (DATA) return DATA;
    return root.GUIDED_CY_DATA || {};
  }
  function setData(d) { DATA = d; }

  function norm(t) { return String(t || '').replace(/[「」【】□\s，,。！!？?、；;：:]/g, ''); }

  /* 把選項放進句子：□ 換成選項，或「X」換成選項 → { before, word, after } */
  function swapIn(m, opt) {
    var s = m.sentence, i, len;
    if (m.kind === 'fill') {
      var mm = s.match(/□+/);
      if (!mm) return null;
      i = mm.index; len = mm[0].length;
    } else {
      var key = '「' + m.target + '」';
      i = s.indexOf(key); len = key.length;
      if (i < 0) return null;
    }
    return { before: s.slice(0, i), word: opt, after: s.slice(i + len) };
  }

  /* 句子哪幾個字可以點：「」裡被考的語詞與 □ 鎖住（lock），標點不能點。回傳 [{ch, tap, lock}] */
  function chars(text) {
    var s = String(text || ''), out = [], lock = {}, i, k, mm;
    var re = /□+|「[^」]*」/g;
    while ((mm = re.exec(s))) for (k = mm.index; k < mm.index + mm[0].length; k++) lock[k] = 1;
    for (i = 0; i < s.length; i++) {
      var c = s[i];
      out.push({ ch: c, lock: !!lock[i], tap: !lock[i] && !/[\s，,。！!？?、；;：:「」]/.test(c) });
    }
    return out;
  }

  /* 線索在句子裡的字位置 → [[start,end), ...] */
  function clueRanges(text, clues) {
    return (clues || []).map(function (c) {
      var i = String(text).indexOf(c);
      return i < 0 ? null : [i, i + c.length];
    });
  }

  /* 學生點的字（物件 {idx:true}）碰到幾個線索 */
  function clueHits(text, clues, picked) {
    var n = 0;
    clueRanges(text, clues).forEach(function (r) {
      if (!r) return;
      for (var k = r[0]; k < r[1]; k++) if (picked && picked[k]) { n++; return; }
    });
    return n;
  }

  /* 分兩堆：assign[i] 是 0/1；和老師的分法相同（不管哪一堆叫第一堆）就算對 */
  function samePiles(assign, groups) {
    if (!assign || assign.length !== groups.length) return false;
    for (var i = 0; i < groups.length; i++) if (assign[i] !== 0 && assign[i] !== 1) return false;
    var a = groups.every(function (g, i) { return g === assign[i]; });
    var b = groups.every(function (g, i) { return g === 1 - assign[i]; });
    return a || b;
  }

  function buildModel(q) {
    if (!q) return null;
    var uid = String(q['題目ID'] || q['題號'] || '');
    var d = data()[uid];
    if (!d) return null;
    var official = String(q['答案'] || '').trim();
    var qText = norm(q['完整題目與選項']);
    if (!official || official !== d.answer || !qText) return null;           // 答案和確認過的資料不同
    if (!d.options || d.options.length !== 4) return null;
    var parts = (d.kind === 'odd' || d.kind === 'usage') ? d.options : [d.sentence].concat(d.options);
    for (var i = 0; i < parts.length; i++) if (qText.indexOf(norm(parts[i])) < 0) return null; // 題目文字改過
    var m = {
      uid: uid, grade: q['年級'] || '', kind: d.kind, q: d.q, sentence: d.sentence, target: d.target,
      options: d.options, answer: d.answer, ans: Number(d.answer) - 1,
      clues: d.clues || [], why: d.why || '', sentences: d.sentences || [], groups: d.groups || [], dict: d.dict || []
    };
    if (m.kind === 'odd' || m.kind === 'usage') {
      if (m.sentences.length !== 4) return null;
      for (var j = 0; j < 4; j++) {
        var cs = m.sentences[j].clues || [];
        for (var k = 0; k < cs.length; k++) if (m.options[j].indexOf(cs[k]) < 0) return null;
      }
      if (m.kind === 'odd') {
        var g = m.sentences.map(function (s) { return s.group; });
        if (g.filter(function (x) { return x === g[m.ans]; }).length !== 1) return null;  // 答案要單獨一堆
        m.pileKey = g;
      }
    } else {
      if (!swapIn(m, m.options[0])) return null;
      for (var c = 0; c < m.clues.length; c++) if (m.sentence.indexOf(m.clues[c]) < 0) return null;
    }
    return m;
  }

  function qualifies(q) { return !!buildModel(q); }

  function pickNext(all, current) {
    var pool = (all || []).filter(function (q) {
      return String(q['題目ID'] || q['題號']) !== String(current.uid) && qualifies(q);
    });
    if (!pool.length) return null;
    var same = pool.filter(function (q) { return (q['年級'] || '') === current.grade; });
    var list = same.length ? same : pool;
    return list[Math.floor(Math.random() * list.length)];
  }

  var core = {
    buildModel: buildModel, qualifies: qualifies, pickNext: pickNext, swapIn: swapIn,
    chars: chars, clueHits: clueHits, samePiles: samePiles, setData: setData
  };
  if (typeof module !== 'undefined' && module.exports) { module.exports = core; return; }

  /* ═══ 畫面 ═══════════════════════════════════════════ */
  var STEPS = ['先猜猜看', '找線索', '換詞讀一讀', '查字典', '想法改了嗎', '原題挑戰'];
  var SOURCE = '學扶國語教材 四年級《語詞一家人（二）》（4-1-3）';
  var STEP_SRC = [
    '準備活動「念一念、想一想」：先讓學生說自己的想法，不急著給答案',
    '發展活動「注意前後文」：從句子裡找出可以幫忙判斷意思的線索',
    '發展活動「換詞讀一讀」：把語詞換進句子讀一讀，意思通順嗎？',
    '發展活動「查找字典」：查到詞義後，再放回句子確認最合適的意思',
    '「你對這個語詞的意思，要修正或補充嗎？」把想法說清楚',
    '統整活動「選一選，是哪一個意思」：自己判斷語境中的詞義'
  ];
  var KIND_TITLE = { mean: '詞義', swap: '詞義', fill: '詞義', odd: '一詞多義', usage: '用詞' };
  var RATE = { know: '我知道', half: '有點懂', no: '還不知道' };
  var S = null, audio = null;
  var cfg = { getAll: function () { return []; } };
  function configure(o) { for (var k in o) cfg[k] = o[k]; }
  function getAll() { return cfg.getAll() || []; }
  function isMulti() { return S.m.kind === 'odd' || S.m.kind === 'usage'; }

  function h(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  /* 「X」、□ 改成醒目樣式 */
  function mark(s) {
    return h(s).replace(/「(.+?)」/g, '<span class="gp-mark">$1</span>').replace(/(□+)/g, '<span class="gp-mark">$1</span>');
  }
  function swapped(i) {
    var p = swapIn(S.m, S.m.options[i]);
    return h(p.before) + '<span class="gp-mark gp-swap">' + h(p.word) + '</span>' + h(p.after);
  }
  function quote(list) { return list.map(function (c) { return '「' + h(c) + '」'; }).join(''); }

  function tone(ok) {
    try {
      audio = audio || new (window.AudioContext || window.webkitAudioContext)();
      var o = audio.createOscillator(), g = audio.createGain();
      o.type = 'sine'; o.frequency.value = ok ? 880 : 330;
      g.gain.setValueAtTime(0.12, audio.currentTime);
      g.gain.exponentialRampToValueAtTime(0.001, audio.currentTime + (ok ? 0.25 : 0.35));
      o.connect(g); g.connect(audio.destination); o.start(); o.stop(audio.currentTime + 0.4);
    } catch (e) { /* 沒有音效也能用 */ }
  }

  function overlay() {
    var el = document.getElementById('gp-overlay');
    if (!el) {
      el = document.createElement('div');
      el.id = 'gp-overlay';
      el.setAttribute('role', 'dialog');
      el.setAttribute('aria-modal', 'true');
      document.body.appendChild(el);
    }
    if (!el._cyBound) {
      el.addEventListener('click', onClick);
      el.addEventListener('input', onInput);
      el._cyBound = true;
    }
    return el;
  }

  function findQ(uid) {
    var all = getAll();
    for (var i = 0; i < all.length; i++) if (String(all[i]['題目ID'] || all[i]['題號']) === String(uid)) return all[i];
    return null;
  }

  function open(uid) {
    var q = findQ(uid), m = q && buildModel(q);
    if (!m) return;
    if (root.GuidedPY && root.GuidedPY.close) root.GuidedPY.close();
    if (root.GuidedCZ && root.GuidedCZ.close) root.GuidedCZ.close();
    if (root.GuidedXJ && root.GuidedXJ.close) root.GuidedXJ.close();
    S = {
      q: q, m: m, step: 0,
      rate: null, guessText: '', guessPick: null, guessUse: {},
      picked: [{}, {}, {}, {}, {}], revealed: false,
      judge: {}, jTries: {}, jDone: {}, jFirst: {}, jHint: {},
      piles: [null, null, null, null], pileSel: null, pileTries: 0, pileDone: false, pileFirst: false, pileHint: '',
      reviseText: '',
      finalPick: null, finalOk: null, finalFirst: null
    };
    document.body.classList.add('gp-open');
    render();
  }

  function close() {
    var el = document.getElementById('gp-overlay');
    if (el && S) el.innerHTML = '';
    document.body.classList.remove('gp-open');
    S = null;
  }

  function stepDone() {
    var m = S.m;
    switch (S.step) {
      case 0:
        if (m.kind === 'odd') return S.guessPick != null;
        if (m.kind === 'usage') return [0, 1, 2, 3].every(function (i) { return S.guessUse[i]; });
        return !!S.rate;
      case 1: return S.revealed;
      case 2:
        if (m.kind === 'odd') return S.pileDone;
        return [0, 1, 2, 3].every(function (i) { return S.jDone[i]; });
      case 3: case 4: return true;
      case 5: return S.finalOk === true;
    }
    return false;
  }

  /* ── ① 先猜 ── */
  function viewGuess() {
    var m = S.m;
    if (m.kind === 'odd') {
      return '<p class="gp-ask">先憑感覺讀一讀：哪一句的<span class="gp-mark">' + h(m.target) + '</span>，意思和另外三句不一樣？</p>' +
        '<div class="gp-choices col">' + m.options.map(function (o, i) {
          return '<button class="gp-choice' + (S.guessPick === i ? ' is-on' : '') + '" data-act="gpick" data-i="' + i + '">(' + (i + 1) + ') ' + mark(o) + '</button>';
        }).join('') + '</div>' +
        (stepDone() ? '<p class="gp-note">好，先記住你的想法，等一下我們找線索來驗證。</p>' : '');
    }
    if (m.kind === 'usage') {
      return '<p class="gp-ask">先自己讀讀看。每一句裡<span class="gp-mark">標出來的語詞</span>，用得對嗎？覺得怪怪的就按「怪怪的」。</p>' +
        m.options.map(function (o, i) {
          var g = S.guessUse[i];
          return '<div class="gp-srow"><div class="gp-stext">(' + (i + 1) + ') ' + mark(o) + '</div><div class="gp-choices">' +
            '<button class="gp-choice sm' + (g === 'ok' ? ' is-on' : '') + '" data-act="guse" data-i="' + i + '" data-v="ok">用得對</button>' +
            '<button class="gp-choice sm' + (g === 'odd' ? ' is-on' : '') + '" data-act="guse" data-i="' + i + '" data-v="odd">怪怪的</button></div></div>';
        }).join('') +
        (stepDone() ? '<p class="gp-note">好，先記住你的想法，等一下我們來驗證。</p>' : '');
    }
    var ask = m.kind === 'fill'
      ? '先不看選項。□ 裡應該放什麼意思的詞？用自己的話說說看。'
      : '先不看選項。<span class="gp-mark">' + h(m.target) + '</span> 在這句話裡是什麼意思？用自己的話說說看。';
    return '<div class="gp-q gp-sent">' + mark(m.sentence) + '</div>' +
      '<p class="gp-ask">' + ask + '</p>' +
      '<textarea class="gp-say" data-in="guess" rows="2" placeholder="可以寫下來，也可以直接說給同學聽（不寫也可以）">' + h(S.guessText) + '</textarea>' +
      '<p class="gp-ask" style="margin-top:12px">你覺得自己知道它的意思嗎？</p>' +
      '<div class="gp-choices">' + Object.keys(RATE).map(function (k) {
        return '<button class="gp-choice sm' + (S.rate === k ? ' is-on' : '') + '" data-act="rate" data-v="' + k + '">' + RATE[k] + '</button>';
      }).join('') + '</div>' +
      (stepDone() ? '<p class="gp-note">不管知不知道都沒關係，接下來我們從句子裡找線索。</p>' : '');
  }

  /* ── ② 找線索 ── */
  function tapLine(text, si, clues) {
    var ranges = S.revealed ? clueRanges(text, clues) : [];
    function inClue(k) { return ranges.some(function (r) { return r && k >= r[0] && k < r[1]; }); }
    return chars(text).map(function (c, k) {
      if (c.lock && (c.ch === '「' || c.ch === '」')) return '';
      if (c.lock) return '<span class="gp-mark gp-lock">' + h(c.ch) + '</span>';
      if (!c.tap) return h(c.ch);
      var on = S.picked[si][k];
      var cls = 'gp-ch' + (on ? ' is-picked' : '') + (inClue(k) ? ' is-clue' : '');
      return '<button class="' + cls + '" data-act="tap" data-s="' + si + '" data-k="' + k + '"' + (S.revealed ? ' disabled' : '') + '>' + h(c.ch) + '</button>';
    }).join('');
  }

  function viewClue() {
    var m = S.m, body, total = 0, hit = 0;
    if (isMulti()) {
      body = m.options.map(function (o, i) {
        var cs = m.sentences[i].clues;
        total += cs.length; hit += clueHits(o, cs, S.picked[i]);
        return '<div class="gp-srow"><div class="gp-stext gp-taps">(' + (i + 1) + ') ' + tapLine(o, i, cs) + '</div></div>';
      }).join('');
    } else {
      total = m.clues.length; hit = clueHits(m.sentence, m.clues, S.picked[4]);
      body = '<div class="gp-q gp-sent gp-taps">' + tapLine(m.sentence, 4, m.clues) + '</div>';
    }
    var ask = isMulti()
      ? '每一句裡，哪些字詞能幫你判斷<span class="gp-mark">' + h(m.target || '標出來的語詞') + '</span>的意思？點一點（可以點好幾個字，再點一次取消）。'
      : '句子裡哪些字詞能幫你猜出意思？點一點（可以點好幾個字，再點一次取消）。';
    var after = '';
    if (S.revealed) {
      after = '<div class="gp-good">老師找的線索用<span class="gp-ch is-clue" style="display:inline">綠底</span>標出來了。你點到了 ' + hit + ' / ' + total + ' 個線索' +
        (hit === total ? '，全部找到了！' : hit ? '，很棒！比比看還漏了哪些。' : '。沒關係，看看綠底的字為什麼是線索。') + '</div>' +
        '<p class="gp-note">你點的字如果不是綠底，也可以說說看為什麼你覺得它有幫助。</p>';
    } else {
      after = '<button class="gp-next glow" data-act="reveal">看看老師找的線索</button>';
    }
    return '<p class="gp-ask">' + ask + '</p>' + body + after;
  }

  /* ── ③ 換詞讀一讀 ── */
  function viewJudge() {
    var m = S.m;
    if (m.kind === 'odd') return viewPiles();
    var ask = {
      mean: '把每個選項放進句子，換掉「' + h(m.target) + '」，讀讀看合不合適。',
      swap: '把每個詞換進句子，換掉「' + h(m.target) + '」，讀讀看意思有沒有變、通不通順。',
      fill: '把每個詞填進 □，讀讀看合不合適。',
      usage: '再讀一次每一句，配合線索判斷：標出來的語詞用得對嗎？'
    }[m.kind];
    var yes = m.kind === 'usage' ? '用得對' : '合適', no = m.kind === 'usage' ? '怪怪的' : '不合適';
    return '<p class="gp-ask">' + ask + '</p>' + m.options.map(function (o, i) {
      var done = S.jDone[i], v = S.judge[i], right = i === m.ans ? 'ok' : 'no';
      var text = m.kind === 'usage' ? mark(o) : swapped(i);
      function btn(val, label) {
        var cls = v === val ? (val === right ? ' is-ok' : ' is-no') : '';
        return '<button class="gp-choice sm' + cls + '" data-act="judge" data-i="' + i + '" data-v="' + val + '"' + (done ? ' disabled' : '') + '>' + label + '</button>';
      }
      return '<div class="gp-srow"><div class="gp-stext">(' + (i + 1) + ') ' + text + '</div>' +
        '<div class="gp-choices">' + btn('ok', yes) + btn('no', no) + '</div>' +
        (S.jHint[i] ? '<div class="gp-hint">' + S.jHint[i] + '</div>' : '') + '</div>';
    }).join('');
  }

  function viewPiles() {
    var m = S.m;
    var loose = m.options.map(function (o, i) {
      if (S.piles[i] != null) return '';
      return '<button class="gp-card gp-scard' + (S.pileSel === i ? ' is-sel' : '') + '" data-act="pcard" data-i="' + i + '">(' + (i + 1) + ') ' + mark(o) + '</button>';
    }).join('');
    var bins = [0, 1].map(function (b) {
      var name = S.pileDone ? pileName(b) : (b === 0 ? '第一堆' : '第二堆');
      var inside = m.options.map(function (o, i) {
        if (S.piles[i] !== b) return '';
        return '<button class="gp-card gp-scard is-in" data-act="pback" data-i="' + i + '"' + (S.pileDone ? ' disabled' : '') + '>(' + (i + 1) + ') ' + mark(o) + '</button>';
      }).join('');
      return '<div class="gp-bin' + (S.pileSel != null ? ' is-target' : '') + '" data-act="pbin" data-b="' + b + '">' +
        '<div class="gp-bin-head">' + h(name) + '</div><div class="gp-bin-body">' + inside + '</div></div>';
    }).join('');
    var full = S.piles.every(function (x) { return x != null; });
    return '<p class="gp-ask">讀讀看每一句，配合線索想一想「' + h(m.target) + '」在這句是什麼意思。<br>意思一樣的放同一堆（先點句子，再點要放的堆；放錯了點一下可以拿回來）。</p>' +
      (loose ? '<div class="gp-pile">' + loose + '</div>' : '') +
      '<div class="gp-bins">' + bins + '</div>' +
      (full && !S.pileDone ? '<button class="gp-next glow" data-act="pcheck">分好了，檢查看看</button>' : '') +
      (S.pileHint ? '<div class="' + (S.pileDone ? 'gp-good' : 'gp-hint') + '">' + S.pileHint + '</div>' : '');
  }

  /* 學生的第 b 堆對應老師的哪一組意思 */
  function pileName(b) {
    var m = S.m;
    for (var i = 0; i < 4; i++) if (S.piles[i] === b) return '意思：' + m.groups[m.pileKey[i]];
    return '';
  }

  /* ── ④ 查字典 ── */
  function dictText(t) {
    return h(t).replace(/\n/g, '<br>').replace(/\[例\]/g, '<span class="gp-eg">例</span>');
  }
  function viewDict() {
    var m = S.m;
    var cards = m.dict.map(function (e) {
      return '<div class="gp-dict"><div class="gp-dict-head">' + h(e.head) + '</div>' +
        '<div class="gp-dict-text">' + dictText(e.text) + '</div>' +
        '<div class="gp-dict-src">' + h(e.source) + (e.note ? '｜' + h(e.note) : '') + '</div></div>';
    }).join('');
    var why = '';
    if (m.kind === 'usage') {
      why = '<ul class="gp-rule-list">' + m.options.map(function (o, i) {
        return '<li>' + (i === m.ans ? '✅ ' : '') + '(' + (i + 1) + ') ' + mark(o) + '<div class="gp-ex">' + h(m.sentences[i].why) + '</div></li>';
      }).join('') + '</ul>';
    } else if (m.kind === 'odd') {
      why = [0, 1].map(function (g) {
        return '<div class="gp-tip" style="margin-bottom:8px"><b>' + h(m.groups[g]) + '</b>：' +
          m.options.map(function (o, i) { return m.pileKey[i] === g ? '(' + (i + 1) + ')' : ''; }).join(' ') + '</div>';
      }).join('') + '<p class="gp-note">' + h(m.why) + '</p>';
    } else {
      why = '<div class="gp-q gp-sent">' + swapped(m.ans) + '</div>' +
        '<p class="gp-note">🔍 線索' + quote(m.clues) + '：' + h(m.why) + '</p>';
    }
    return '<p class="gp-ask">查查字典，看看這些語詞的意思，再放回句子想一想。</p>' + cards + why;
  }

  /* ── ⑤ 想法改了嗎 ── */
  function viewRevise() {
    var m = S.m;
    if (m.kind === 'usage') {
      var right = 0;
      var rows = m.options.map(function (o, i) {
        var said = S.guessUse[i] === 'ok' ? '用得對' : '怪怪的', truth = i === m.ans ? '用得對' : '用得不對';
        var ok = (S.guessUse[i] === 'ok') === (i === m.ans);
        if (ok) right++;
        return '<tr><td>(' + (i + 1) + ') ' + mark(o) + '</td><td>' + said + '</td><td>' + truth + (ok ? ' ✓' : '') + '</td></tr>';
      }).join('');
      var bad = m.options.map(function (o, i) { return i; }).filter(function (i) { return i !== m.ans; })[0];
      var bw = (m.options[bad].match(/「(.+?)」/) || [])[1] || '';
      return '<p class="gp-ask">一開始你的判斷，和現在比比看：</p>' +
        '<table class="gp-tbl"><thead><tr><th>句子</th><th>你一開始覺得</th><th>其實</th></tr></thead><tbody>' + rows + '</tbody></table>' +
        '<p class="gp-note">' + (right === 4 ? '你一開始就全部看出來了，很厲害！現在你還能說出為什麼。' : '想法改變了也很好，這就是在學習。') + '</p>' +
        frame('第 ' + (bad + 1) + ' 句的「' + h(bw) + '」用得不對，因為句子裡有' + quote(m.sentences[bad].clues.slice(0, 1)) + '，「' + h(bw) + '」的意思是<u>　　　　</u>，配不起來。');
    }
    if (m.kind === 'odd') {
      var same = S.guessPick === m.ans;
      return '<p class="gp-ask">一開始你選第 ' + (S.guessPick + 1) + ' 句，其實意思不一樣的是第 ' + (m.ans + 1) + ' 句。</p>' +
        '<p class="gp-note">' + (same ? '你一開始就猜對了，現在還能說出為什麼，很厲害！' : '想法改變了也很好，這就是在學習。') + '</p>' +
        frame('第 ' + (m.ans + 1) + ' 句的「' + h(m.target) + '」是「' + h(m.groups[m.pileKey[m.ans]]) + '」的意思；其他三句是「' +
          h(m.groups[1 - m.pileKey[m.ans]]) + '」的意思。我是看' + quote(m.sentences[m.ans].clues.slice(0, 1)) + '知道的。');
    }
    var said = S.guessText.trim();
    var ansWord = m.options[m.ans];
    return '<p class="gp-ask">一開始你的想法：</p>' +
      '<div class="gp-q">' + (said ? '「' + h(said) + '」' : '（沒有寫下來）') + '　<span class="gp-ex">你覺得自己：' + h(RATE[S.rate] || '') + '</span></div>' +
      '<p class="gp-ask">現在知道了：' + (m.kind === 'fill' ? '□ 要填「' + h(ansWord) + '」' : '「' + h(m.target) + '」在這句的意思是「' + h(ansWord) + '」') + '。想法要修正或補充嗎？</p>' +
      '<textarea class="gp-say" data-in="revise" rows="2" placeholder="用自己的話再說一次（不寫也可以）">' + h(S.reviseText) + '</textarea>' +
      frame(m.kind === 'fill'
        ? '□ 要填「' + h(ansWord) + '」，因為句子裡有' + quote(m.clues.slice(0, 1)) + '，意思是<u>　　　　</u>。'
        : '「' + h(m.target) + '」在這句的意思是<u>　　　　</u>，因為句子裡有' + quote(m.clues.slice(0, 1)) + '。');
  }
  function frame(t) {
    return '<div class="gp-frame"><div class="gp-frame-lbl">🗣️ 用這句話說一次給同學聽</div>' + t + '</div>';
  }

  /* ── ⑥ 原題挑戰 ── */
  function viewFinal() {
    var m = S.m;
    var stem = m.q + (m.sentence && !isMulti() ? '<div class="gp-sent" style="margin-top:6px">' + mark(m.sentence) + '</div>' : '');
    var btns = m.options.map(function (o, i) {
      var on = S.finalPick === i;
      var cls = on ? (i === m.ans ? ' is-ok' : ' is-no') : '';
      return '<button class="gp-choice' + cls + '" data-act="final" data-i="' + i + '"' + (S.finalOk ? ' disabled' : '') + '>(' + (i + 1) + ') ' + mark(o) + '</button>';
    }).join('');
    var msg = '';
    if (S.finalOk) msg = '<div class="gp-good">答對了！你用「找線索、換詞讀一讀」解出了考古題。</div>';
    else if (S.finalPick != null) msg = '<div class="gp-hint">再想想：回頭看看句子裡的線索，把這個選項放進句子讀一讀，合適嗎？</div>';
    return '<p class="gp-ask">現在自己做做看原題：</p><div class="gp-q">' + stem + '</div>' +
      '<div class="gp-choices col">' + btns + '</div>' + msg + (S.finalOk ? viewSummary() : '');
  }

  function viewSummary() {
    var m = S.m, guessLine, clueLine, judgeLine;
    if (m.kind === 'odd') guessLine = '一開始猜：' + (S.guessPick === m.ans ? '猜對了' : '猜錯，後來改正');
    else if (m.kind === 'usage') guessLine = '一開始就判斷對：' + [0, 1, 2, 3].filter(function (i) { return (S.guessUse[i] === 'ok') === (i === m.ans); }).length + ' / 4 句';
    else guessLine = '一開始覺得自己：' + (RATE[S.rate] || '');
    var total = 0, hit = 0;
    if (isMulti()) m.options.forEach(function (o, i) { total += m.sentences[i].clues.length; hit += clueHits(o, m.sentences[i].clues, S.picked[i]); });
    else { total = m.clues.length; hit = clueHits(m.sentence, m.clues, S.picked[4]); }
    clueLine = '找到線索：' + hit + ' / ' + total + ' 個';
    if (m.kind === 'odd') judgeLine = '分兩堆：' + (S.pileFirst ? '第一次就分對' : '試了不只一次');
    else judgeLine = '換詞讀一讀一次就判斷對：' + [0, 1, 2, 3].filter(function (i) { return S.jFirst[i]; }).length + ' / 4 個';
    var next = pickNext(getAll(), m);
    return '<div class="gp-sum"><div class="gp-frame-lbl">📋 這一題的學習紀錄（給老師看）</div>' +
      '<div>' + guessLine + '</div><div>' + clueLine + '</div><div>' + judgeLine + '</div>' +
      '<div>原題：' + (S.finalFirst ? '第一次就答對' : '試了不只一次才答對') + '</div></div>' +
      (next ? '<button class="gp-next glow" data-act="another" data-id="' + h(next['題目ID'] || next['題號']) + '">換一題再試試 →</button>' : '');
  }

  function render() {
    if (!S) return;
    var el = overlay();
    var bar = STEPS.map(function (t, i) {
      return '<span class="gp-step' + (i === S.step ? ' is-now' : i < S.step ? ' is-done' : '') + '">' + (i < S.step ? '✓' : (i + 1)) + ' ' + t + '</span>';
    }).join('');
    var body = [viewGuess, viewClue, viewJudge, viewDict, viewRevise, viewFinal][S.step]();
    var done = stepDone();
    var nextBtn = S.step < 5 ? '<button class="gp-next' + (done ? ' glow' : '') + '" data-act="next"' + (done ? '' : ' disabled') + '>下一步 →</button>' : '';
    el.innerHTML =
      '<div class="gp-panel" data-kind="cy">' +
        '<div class="gp-top"><div class="gp-title">✏️ 一起動腦：' + h(KIND_TITLE[S.m.kind]) + (S.m.target ? '「<b>' + h(S.m.target) + '</b>」' : '') + '</div>' +
          '<button class="gp-close" data-act="close" aria-label="離開">✕ 離開</button></div>' +
        '<div class="gp-steps">' + bar + '</div>' +
        '<div class="gp-src"><b>📘 教學依據</b>' + h(SOURCE) + '<span class="gp-src-step">這一步：' + h(STEP_SRC[S.step]) + '</span></div>' +
        '<div class="gp-body">' + body + '</div>' +
        '<div class="gp-foot">' +
          (S.step > 0 && S.step < 5 ? '<button class="gp-back" data-act="back">← 上一步</button>' : '<span></span>') +
          nextBtn + '</div>' +
      '</div>';
  }

  function onInput(ev) {
    if (!S) return;
    var k = ev.target.getAttribute && ev.target.getAttribute('data-in');
    if (k === 'guess') S.guessText = ev.target.value;
    if (k === 'revise') S.reviseText = ev.target.value;
  }

  function judgeHint(i, wrongTimes) {
    var m = S.m;
    var cs = m.kind === 'usage' ? m.sentences[i].clues : m.clues;
    if (wrongTimes === 1) return '再讀一次整句，和線索' + quote(cs) + '對得上嗎？';
    var ok = i === m.ans;
    if (m.kind === 'usage') return ok ? '這一句用得對，老師幫你標好了。' : '這一句用得不對，老師幫你標好了。等一下查字典看看為什麼。';
    return '「' + h(m.options[i]) + '」放進句子' + (ok ? '合適' : '不合適') + '，老師幫你標好了。等一下查字典看看為什麼。';
  }

  function onClick(ev) {
    if (!S) return;                               // 其他一起動腦開著時交給它們
    var b = ev.target.closest('[data-act]');
    if (!b || b.disabled) return;
    var act = b.getAttribute('data-act'), v = b.getAttribute('data-v');
    var i = Number(b.getAttribute('data-i'));

    if (act === 'close') return close();
    if (act === 'next' && stepDone()) { S.step++; return render(); }
    if (act === 'back') { S.step--; return render(); }
    if (act === 'another') { var nid = b.getAttribute('data-id'); close(); return open(nid); }
    if (act === 'rate') { S.rate = v; return render(); }
    if (act === 'gpick') { S.guessPick = i; return render(); }
    if (act === 'guse') { S.guessUse[i] = v; return render(); }

    if (act === 'tap') {
      var si = Number(b.getAttribute('data-s')), k = Number(b.getAttribute('data-k'));
      if (S.picked[si][k]) delete S.picked[si][k]; else S.picked[si][k] = true;
      return render();
    }
    if (act === 'reveal') { S.revealed = true; tone(true); return render(); }

    if (act === 'judge') {
      if (S.jDone[i]) return;
      var right = i === S.m.ans ? 'ok' : 'no';
      S.judge[i] = v;
      if (v === right) {
        S.jDone[i] = true;
        if (!S.jTries[i]) S.jFirst[i] = true;
        S.jHint[i] = '';
        tone(true);
      } else {
        S.jTries[i] = (S.jTries[i] || 0) + 1;
        tone(false);
        S.jHint[i] = judgeHint(i, S.jTries[i]);
        if (S.jTries[i] >= 2) { S.jDone[i] = true; S.judge[i] = right; }
      }
      return render();
    }

    if (act === 'pcard') { S.pileSel = S.pileSel === i ? null : i; return render(); }
    if (act === 'pback') {
      if (S.pileSel != null) {                     // 手上有句子時，點到堆裡的句子＝放進這一堆
        S.piles[S.pileSel] = S.piles[i]; S.pileSel = null; S.pileHint = '';
      } else { S.piles[i] = null; S.pileHint = ''; }
      return render();
    }
    if (act === 'pbin') {
      if (S.pileSel == null) { S.pileHint = '先點一個句子，再點要放的堆。'; return render(); }
      S.piles[S.pileSel] = Number(b.getAttribute('data-b')); S.pileSel = null; S.pileHint = '';
      return render();
    }
    if (act === 'pcheck') {
      if (samePiles(S.piles, S.m.pileKey)) {
        S.pileDone = true; if (!S.pileTries) S.pileFirst = true;
        S.pileHint = '分對了！兩堆的「' + h(S.m.target) + '」意思不一樣，堆名已經寫出來了。';
        tone(true);
      } else {
        S.pileTries++; tone(false);
        if (S.pileTries === 1) {
          S.pileHint = '有句子放錯堆了。讀讀看每句的線索：' + S.m.options.map(function (o, j) {
            return '(' + (j + 1) + ')' + quote(S.m.sentences[j].clues);
          }).join('　');
        } else {
          S.piles = S.m.pileKey.slice(); S.pileDone = true;
          S.pileHint = '老師幫你分好了。看看堆名，想一想為什麼這樣分。';
        }
      }
      return render();
    }

    if (act === 'final') {
      var ok = i === S.m.ans;
      if (S.finalPick == null) S.finalFirst = ok;
      S.finalPick = i; S.finalOk = ok; tone(ok);
      return render();
    }
  }

  document.addEventListener('keydown', function (e) { if (S && e.key === 'Escape') close(); });

  root.GuidedCY = { open: open, close: close, configure: configure, qualifies: qualifies, buildModel: buildModel };
})(typeof window !== 'undefined' ? window : globalThis);
