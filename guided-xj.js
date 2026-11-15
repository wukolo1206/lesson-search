/* 一起動腦：形近字（依學扶教材 3-1-5《國字會分工》：「一邊是意思，一邊是聲音」）
 * 先猜猜看 → 找相同的地方 → 哪邊管意思 → 哪邊管聲音 → 想法改了嗎 → 原題挑戰
 * 資料：guided-xj-data.js（scripts/build_guided_xj_data.py 產生；部首、讀音、釋義依簡編本，拆字依 cjkvi-ids，全部自動查證）
 * 設計文件：字音形互動引導設計_教材依據.md（題型 B）
 * 樣式沿用 guided-py.css 的 gp-*。
 */
(function (root) {
  'use strict';

  /* ═══ 核心（可在 Node 測試） ═══════════════════════════ */
  var DATA = null;
  function data() { return DATA || root.GUIDED_XJ_DATA || {}; }
  function setData(d) { DATA = d; }

  /* 少見偏旁寫法改成平板一定顯示得出來的字 */
  var SHOW = { '⺼': '月', '𥫗': '竹', '㣺': '心', '糹': '糸', '⻌': '辶' };
  function show(p) { return SHOW[p] || p; }

  function parseOptions(text) {
    var out = [], re = /[（(]([1-4１-４])[)）]([^（(]*)/g, m;
    while ((m = re.exec(String(text || '')))) out.push({ no: String('１２３４'.indexOf(m[1]) >= 0 ? '１２３４'.indexOf(m[1]) + 1 : m[1]), text: m[2].trim() });
    return out;
  }

  function buildModel(q) {
    if (!q) return null;
    var uid = String(q['題目ID'] || q['題號'] || '');
    var d = data()[uid];
    if (!d) return null;
    var official = String(q['答案'] || '').trim();
    if (!official || official !== d.answer) return null;
    var text = String(q['完整題目與選項'] || '');
    var opts = parseOptions(text);
    if (opts.length !== 4) return null;
    var pairs = d.pairs.filter(function (p) {
      return p.shared && p[p.a] && p[p.b] && p[p.a].read && p[p.b].read && (p[p.a].part || p[p.b].part);
    });
    if (!pairs.length) return null;
    /* 題目裡至少要出現每組的一個字（另一字可能是用來對照的形近字） */
    for (var i = 0; i < pairs.length; i++) if (text.indexOf(pairs[i].a) < 0 && text.indexOf(pairs[i].b) < 0) return null;
    var stem = text.split(/[（(][1-4１-４][)）]/)[0].replace(/^\s*\d+\.\s*/, '');
    return { uid: uid, grade: q['年級'] || '', pairs: pairs, answer: official, options: opts, stem: stem };
  }

  function qualifies(q) { return !!buildModel(q); }

  /* 第 2 步選項：共同部件＋兩邊不同的部件 */
  function partChoices(p) {
    var c = [p.shared];
    [p.a, p.b].forEach(function (ch) { var x = p[ch].part; if (x && c.indexOf(x) < 0) c.push(x); });
    return c;
  }

  /* 第 3 步：有偏旁的字（整字就是共同部件的不問） */
  function meaningItems(m) {
    var out = [];
    m.pairs.forEach(function (p, pi) {
      [p.a, p.b].forEach(function (ch) { if (p[ch].part) out.push({ pi: pi, ch: ch, info: p[ch], shared: p.shared }); });
    });
    return out;
  }

  function pickNext(all, current) {
    var pool = (all || []).filter(function (q) { return String(q['題目ID'] || q['題號']) !== String(current.uid) && qualifies(q); });
    if (!pool.length) return null;
    var same = pool.filter(function (q) { return (q['年級'] || '') === current.grade; });
    var list = same.length ? same : pool;
    return list[Math.floor(Math.random() * list.length)];
  }

  var core = { buildModel: buildModel, qualifies: qualifies, partChoices: partChoices, meaningItems: meaningItems, parseOptions: parseOptions, show: show, setData: setData, pickNext: pickNext };
  if (typeof module !== 'undefined' && module.exports) { module.exports = core; return; }

  /* ═══ 畫面 ═══════════════════════════════════════════ */
  var STEPS = ['先猜猜看', '找相同的地方', '哪邊管意思', '哪邊管聲音', '想法改了嗎', '原題挑戰'];
  var SOURCE = '學扶國語教材 三年級《國字會分工》（3-1-5）';
  var STEP_SRC = [
    '準備活動：先讓學生觀察、說出自己的發現，不急著給答案',
    '發展活動（一）：圈出字裡相同的地方（教材「馬」系列字的做法）',
    '核心口訣「一邊是意思」：部首表示字的意思；查字典確認',
    '核心口訣「一邊是聲音」：相同的部件常讓讀音很像，但不一定完全一樣',
    '發展活動（二）：「查完字典後，你要修正之前的想法或補充嗎？」',
    '統整活動：用學到的方法，判斷其他國字'
  ];
  var S = null, audio = null;
  var cfg = { getAll: function () { return []; } };
  function configure(o) { for (var k in o) cfg[k] = o[k]; }
  function getAll() { return cfg.getAll() || []; }

  function h(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function inWord(word, ch) { return h(word).replace(h(ch), '<span class="gp-mark">' + h(ch) + '</span>'); }
  function mark(s) { return h(s).replace(/「([^「」]+)」/g, '<span class="gp-mark">$1</span>').replace(/(□+)/g, '<span class="gp-mark">$1</span>'); }
  function dictText(t) { return h(t).replace(/\n/g, '<br>').replace(/\[例\]/g, '<span class="gp-eg">例</span>'); }
  function shuffle(a) { a = a.slice(); for (var i = a.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)); var t = a[i]; a[i] = a[j]; a[j] = t; } return a; }

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
      el = document.createElement('div'); el.id = 'gp-overlay';
      el.setAttribute('role', 'dialog'); el.setAttribute('aria-modal', 'true');
      document.body.appendChild(el);
    }
    if (!el._xjBound) { el.addEventListener('click', onClick); el._xjBound = true; }
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
    ['GuidedPY', 'GuidedCZ', 'GuidedCY', 'GuidedApply', 'GuidedSD'].forEach(function (k) { if (root[k] && root[k].close) root[k].close(); });
    var mi = meaningItems(m);
    S = {
      q: q, m: m, step: 0, mi: mi,
      guess: {},
      sPick: {}, sTries: {}, sDone: {}, sFirst: {}, sHint: {}, sOpts: m.pairs.map(function (p) { return shuffle(partChoices(p)); }),
      mPick: {}, mTries: {}, mDone: {}, mFirst: {}, mHint: {}, mOpts: mi.map(function (x) { return shuffle([x.info.part, x.shared]); }),
      rPick: {}, rTries: {}, rDone: {}, rFirst: {}, rHint: {},
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

  function all(n, obj) { for (var i = 0; i < n; i++) if (!obj[i]) return false; return true; }
  function stepDone() {
    var n = S.m.pairs.length;
    switch (S.step) {
      case 0: return all(n, S.guess);
      case 1: return all(n, S.sDone);
      case 2: return all(S.mi.length, S.mDone);
      case 3: return all(n, S.rDone);
      case 4: return true;
      case 5: return S.finalOk === true;
    }
    return false;
  }

  function pairHead(p) {
    return '<div class="gp-xj-pair"><span class="gp-xj-w">' + inWord(p[p.a].word, p.a) + '</span><span class="gp-xj-sep">／</span><span class="gp-xj-w">' + inWord(p[p.b].word, p.b) + '</span></div>';
  }

  /* ── ① 先猜 ── */
  function viewGuess() {
    return '<p class="gp-ask">先自己讀讀看。每一組<span class="gp-mark">標出來的兩個字</span>長得很像，它們的讀音一樣嗎？憑感覺選就好。</p>' +
      S.m.pairs.map(function (p, i) {
        var g = S.guess[i];
        return '<div class="gp-srow">' + pairHead(p) + '<div class="gp-choices">' +
          '<button class="gp-choice sm' + (g === 'same' ? ' is-on' : '') + '" data-act="guess" data-i="' + i + '" data-v="same">讀音一樣</button>' +
          '<button class="gp-choice sm' + (g === 'diff' ? ' is-on' : '') + '" data-act="guess" data-i="' + i + '" data-v="diff">讀音不一樣</button></div></div>';
      }).join('') + (stepDone() ? '<p class="gp-note">好，先記住你的想法，等一下我們來驗證。</p>' : '');
  }

  /* ── ② 找相同的地方 ── */
  function viewShared() {
    return '<p class="gp-ask">把兩個字放在一起看：它們<b>哪裡長得一樣</b>？點出兩個字都有的那一塊。</p>' +
      S.m.pairs.map(function (p, i) {
        var done = S.sDone[i];
        var btns = S.sOpts[i].map(function (c) {
          var on = S.sPick[i] === c, cls = on ? (c === p.shared ? ' is-ok' : ' is-no') : '';
          return '<button class="gp-choice gp-big' + cls + '" data-act="shared" data-i="' + i + '" data-v="' + h(c) + '"' + (done ? ' disabled' : '') + '>' + h(show(c)) + '</button>';
        }).join('');
        return '<div class="gp-srow"><div class="gp-xj-big">' + h(p.a) + '<small>' + h(p[p.a].word) + '</small></div><div class="gp-xj-big">' + h(p.b) + '<small>' + h(p[p.b].word) + '</small></div>' +
          '<div class="gp-choices">' + btns + '</div>' + (S.sHint[i] ? '<div class="gp-hint">' + S.sHint[i] + '</div>' : '') + '</div>';
      }).join('');
  }

  /* ── ③ 哪邊管意思 ── */
  function viewMeaning() {
    return '<p class="gp-ask">一個字常常分成兩邊：<b>一邊是意思，一邊是聲音</b>。剛才找到的「兩個字都有的那一塊」常常是管聲音的；那<b>管意思的是哪一邊</b>？</p>' +
      S.mi.map(function (x, i) {
        var done = S.mDone[i];
        var btns = S.mOpts[i].map(function (c) {
          var on = S.mPick[i] === c, cls = on ? (c === x.info.part ? ' is-ok' : ' is-no') : '';
          return '<button class="gp-choice gp-big' + cls + '" data-act="mean" data-i="' + i + '" data-v="' + h(c) + '"' + (done ? ' disabled' : '') + '>' + h(show(c)) + '</button>';
        }).join('');
        var info = done ? '<div class="gp-good"><b>' + h(x.ch) + '</b>是<b>' + h(x.info.rad) + '部</b>：' + h(x.info.radMean) +
          (x.info.radDef ? '<div class="gp-ex">簡編本「' + h(x.info.rad) + '」：' + h(x.info.radDef) + '</div>' : '') +
          (x.info.wordDef ? '<div class="gp-ex">簡編本「' + h(x.info.word) + '」：' + dictText(x.info.wordDef.split('\n')[0]) + '</div>' : '') + '</div>' : '';
        return '<div class="gp-srow"><div class="gp-stext"><span class="gp-xj-big sm">' + h(x.ch) + '</span>　' + inWord(x.info.word, x.ch) + '</div>' +
          '<div class="gp-choices">' + btns + '</div>' + (S.mHint[i] ? '<div class="gp-hint">' + S.mHint[i] + '</div>' : '') + info + '</div>';
      }).join('');
  }

  /* ── ④ 哪邊管聲音 ── */
  function viewSound() {
    return '<p class="gp-ask">兩個字都有一樣的部件，讀音會一樣嗎？看看字典的注音，選「一樣」或「不一樣」。</p>' +
      S.m.pairs.map(function (p, i) {
        var done = S.rDone[i], right = p.sameRead ? 'same' : 'diff';
        function btn(v, label) {
          var on = S.rPick[i] === v, cls = on ? (v === right ? ' is-ok' : ' is-no') : '';
          return '<button class="gp-choice sm' + cls + '" data-act="sound" data-i="' + i + '" data-v="' + v + '"' + (done ? ' disabled' : '') + '>' + label + '</button>';
        }
        return '<div class="gp-srow"><div class="gp-xj-read">' +
          '<span>' + inWord(p[p.a].word, p.a) + ' <b class="gp-bopo">' + h(p[p.a].read) + '</b></span>' +
          '<span>' + inWord(p[p.b].word, p.b) + ' <b class="gp-bopo">' + h(p[p.b].read) + '</b></span></div>' +
          '<div class="gp-ex">共同部件「' + h(show(p.shared)) + '」</div>' +
          '<div class="gp-choices">' + btn('same', '讀音一樣') + btn('diff', '讀音不一樣') + '</div>' +
          (S.rHint[i] ? '<div class="gp-hint">' + S.rHint[i] + '</div>' : '') + '</div>';
      }).join('') +
      (stepDone() ? '<div class="gp-rule">一邊是意思，一邊是聲音。</div><p class="gp-note">一樣的部件常常讓讀音很像，但<b>不一定完全一樣</b>。所以讀音要查字典確認，寫字要先想意思、看部首。</p>' : '') +
      '<div class="gp-dict-src" style="margin-top:8px">注音：教育部《國語辭典簡編本》（以語詞的注音為準）</div>';
  }

  /* ── ⑤ 想法改了嗎 ── */
  function viewRevise() {
    var right = 0;
    var rows = S.m.pairs.map(function (p, i) {
      var said = S.guess[i] === 'same' ? '讀音一樣' : '讀音不一樣';
      var truth = p.sameRead ? '讀音一樣' : '讀音不一樣';
      var ok = (S.guess[i] === 'same') === p.sameRead; if (ok) right++;
      return '<tr><td>' + h(p.a) + '（' + h(p[p.a].read) + '）／' + h(p.b) + '（' + h(p[p.b].read) + '）</td><td>' + said + '</td><td>' + truth + (ok ? ' ✓' : '') + '</td></tr>';
    }).join('');
    var p = S.m.pairs.filter(function (x) { return x[x.a].part && x[x.b].part; })[0] || S.m.pairs[0];
    var A = p[p.a], B = p[p.b];
    var line = '「' + h(p.a) + '」和「' + h(p.b) + '」都有「' + h(show(p.shared)) + '」，讀音' + (p.sameRead ? '一樣' : '很像但不一樣') + '；';
    [[p.a, A], [p.b, B]].forEach(function (z) { if (z[1].part) line += '「' + h(z[0]) + '」是<u>' + h(z[1].rad) + '部</u>，' + h(z[1].radMean) + '；'; });
    return '<p class="gp-ask">一開始你的想法，和現在比比看：</p>' +
      '<table class="gp-tbl"><thead><tr><th>字</th><th>你一開始覺得</th><th>字典說</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<p class="gp-note">' + (right === S.m.pairs.length ? '你一開始就全部猜對了，很厲害！現在你還能說出為什麼。' : '想法改變了也很好，這就是在學習。') + '</p>' +
      '<div class="gp-frame"><div class="gp-frame-lbl">🗣️ 用這句話說一次給同學聽</div>' + line.replace(/；$/, '。') + '</div>';
  }

  /* ── ⑥ 原題挑戰 ── */
  function viewFinal() {
    var btns = S.m.options.map(function (o) {
      var on = S.finalPick === o.no, cls = on ? (o.no === S.m.answer ? ' is-ok' : ' is-no') : '';
      return '<button class="gp-choice' + cls + '" data-act="final" data-v="' + h(o.no) + '"' + (S.finalOk ? ' disabled' : '') + '>(' + h(o.no) + ') ' + mark(o.text) + '</button>';
    }).join('');
    var msg = '';
    if (S.finalOk) msg = '<div class="gp-good">答對了！你用「一邊是意思，一邊是聲音」解出了考古題。</div>';
    else if (S.finalPick != null) msg = '<div class="gp-hint">再想想：這幾個字哪一邊管意思？讀音真的一樣嗎？回頭看看字典的注音。</div>';
    return '<p class="gp-ask">現在自己做做看原題：</p><div class="gp-q">' + mark(S.m.stem) + '</div>' +
      '<div class="gp-choices col">' + btns + '</div>' + msg + (S.finalOk ? viewSummary() : '');
  }

  function count(obj, n) { var c = 0; for (var i = 0; i < n; i++) if (obj[i]) c++; return c; }
  function viewSummary() {
    var n = S.m.pairs.length;
    var guessRight = S.m.pairs.filter(function (p, i) { return (S.guess[i] === 'same') === p.sameRead; }).length;
    var next = pickNext(getAll(), S.m);
    return '<div class="gp-sum"><div class="gp-frame-lbl">📋 這一題的學習紀錄（給老師看）</div>' +
      '<div>一開始讀音就猜對：' + guessRight + ' / ' + n + ' 組</div>' +
      '<div>相同的地方一次就找對：' + count(S.sFirst, n) + ' / ' + n + ' 組</div>' +
      '<div>管意思的一邊一次就選對：' + count(S.mFirst, S.mi.length) + ' / ' + S.mi.length + ' 個字</div>' +
      '<div>原題：' + (S.finalFirst ? '第一次就答對' : '試了不只一次才答對') + '</div></div>' +
      applyBtn(next) +
      (next ? '<button class="' + (hasApply() ? 'gp-back' : 'gp-next glow') + '" data-act="another" data-id="' + h(next['題目ID'] || next['題號']) + '">換一題再試試 →</button>' : '');
  }

  function hasApply() { return !!(root.GuidedApply && root.GuidedApply.has(S.m.uid)); }
  function applyBtn(next) {
    return hasApply() ? '<button class="gp-next glow" data-act="apply" data-next="' + h(next ? (next['題目ID'] || next['題號']) : '') + '">換個句子用用看 →</button> ' : '';
  }

  function render() {
    if (!S) return;
    var el = overlay();
    var bar = STEPS.map(function (t, i) {
      return '<span class="gp-step' + (i === S.step ? ' is-now' : i < S.step ? ' is-done' : '') + '">' + (i < S.step ? '✓' : (i + 1)) + ' ' + t + '</span>';
    }).join('');
    var body = [viewGuess, viewShared, viewMeaning, viewSound, viewRevise, viewFinal][S.step]();
    var done = stepDone();
    var nextBtn = S.step < 5 ? '<button class="gp-next' + (done ? ' glow' : '') + '" data-act="next"' + (done ? '' : ' disabled') + '>下一步 →</button>' : '';
    el.innerHTML =
      '<div class="gp-panel" data-kind="xj">' +
        '<div class="gp-top"><div class="gp-title">✏️ 一起動腦：形近字</div>' +
          '<button class="gp-close" data-act="close" aria-label="離開">✕ 離開</button></div>' +
        '<div class="gp-steps">' + bar + '</div>' +
        '<div class="gp-src"><b>📘 教學依據</b>' + h(SOURCE) + '<span class="gp-src-step">這一步：' + h(STEP_SRC[S.step]) + '</span></div>' +
        '<div class="gp-body">' + body + '</div>' +
        '<div class="gp-foot">' +
          (S.step > 0 && S.step < 5 ? '<button class="gp-back" data-act="back">← 上一步</button>' : '<span></span>') +
          nextBtn + '</div>' +
      '</div>';
  }

  /* 選擇題共用：對了記一次就對；錯一次給提示；錯兩次幫忙選好 */
  function answer(kind, i, v, right, hint1, hint2) {
    var P = S[kind + 'Pick'], T = S[kind + 'Tries'], D = S[kind + 'Done'], F = S[kind + 'First'], H = S[kind + 'Hint'];
    if (D[i]) return;
    P[i] = v;
    if (v === right) { D[i] = true; if (!T[i]) F[i] = true; H[i] = ''; tone(true); }
    else {
      T[i] = (T[i] || 0) + 1; tone(false);
      if (T[i] === 1) H[i] = hint1;
      else { D[i] = true; P[i] = right; H[i] = hint2; }
    }
  }

  function onClick(ev) {
    if (!S) return;
    var b = ev.target.closest('[data-act]');
    if (!b || b.disabled) return;
    var act = b.getAttribute('data-act'), v = b.getAttribute('data-v'), i = Number(b.getAttribute('data-i'));
    if (act === 'close') return close();
    if (act === 'next' && stepDone()) { S.step++; return render(); }
    if (act === 'back') { S.step--; return render(); }
    if (act === 'another') { var nid = b.getAttribute('data-id'); close(); return open(nid); }
    if (act === 'apply') { var au = S.m.uid, an = b.getAttribute('data-next'); close(); return root.GuidedApply.open(au, { mod: 'GuidedXJ', next: an }); }
    if (act === 'guess') { S.guess[i] = v; return render(); }
    if (act === 'shared') {
      var p = S.m.pairs[i];
      answer('s', i, v, p.shared, '再看一次：哪一塊是「' + h(p.a) + '」有、「' + h(p.b) + '」也有的？',
        '兩個字都有「' + h(show(p.shared)) + '」，老師幫你選好了。');
      return render();
    }
    if (act === 'mean') {
      var x = S.mi[i];
      answer('m', i, v, x.info.part, '「' + h(show(x.shared)) + '」是兩個字都有的，它管聲音；那另一邊呢？',
        '「' + h(show(x.info.part)) + '」這一邊管意思（' + h(x.info.rad) + '部，' + h(x.info.radMean) + '），老師幫你選好了。');
      return render();
    }
    if (act === 'sound') {
      var q = S.m.pairs[i];
      answer('r', i, v, q.sameRead ? 'same' : 'diff', '仔細比一比兩個注音，聲調也要看喔。',
        '「' + h(q.a) + '」讀' + h(q[q.a].read) + '，「' + h(q.b) + '」讀' + h(q[q.b].read) + '，' + (q.sameRead ? '一樣' : '不一樣') + '。老師幫你選好了。');
      return render();
    }
    if (act === 'final') {
      var ok = v === S.m.answer;
      if (S.finalPick == null) S.finalFirst = ok;
      if (S.finalPick == null && root.ClassResults) root.ClassResults.add(S.m.uid, 'guided', ok ? 1 : 0, 1);
      S.finalPick = v; S.finalOk = ok; tone(ok);
      return render();
    }
  }

  document.addEventListener('keydown', function (e) { if (S && e.key === 'Escape') close(); });

  root.GuidedXJ = { open: open, close: close, configure: configure, qualifies: qualifies, buildModel: buildModel };
})(typeof window !== 'undefined' ? window : globalThis);
