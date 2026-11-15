/* 一起動腦：讀音比一比（多音字、同音字的「字音前後相同」）＋看意思選字（同音字填字、注音寫國字）
 * 讀音比一比：先猜 → 查讀音 → 比意思 → 想法改了嗎 → 原題挑戰
 *   依據：3-1-3《語詞一家人（一）》音隨義轉；不同字讀音一樣時，依 3-1-5《國字會分工》看意思選字
 * 看意思選字：先猜 → 讀讀看 → 想意思選字 → 部首和意思 → 原題挑戰
 *   依據：3-1-5《國字會分工》「一邊是意思，一邊是聲音」：讀音一樣，就要看意思和部首選字
 * 資料：guided-sound-data.js（scripts/build_guided_sound_data.py：題目文字解析＋簡編本查證＋官方答案核對）
 * 樣式沿用 guided-py.css 的 gp-*。
 */
(function (root) {
  'use strict';

  var DATA = null;
  function data() { return DATA || root.GUIDED_SOUND_DATA || {}; }
  function setData(d) { DATA = d; }
  function qid(q) { return String((q && (q['題目ID'] || q['題號'])) || ''); }
  function norm(t) { return String(t || '').replace(/[「」□\s，,。！!？?、；;：:／╱/]/g, ''); }

  function parseOptions(text) {
    var out = [], re = /[（(]([1-4１-４])[)）]([^（(]*)/g, m;
    while ((m = re.exec(String(text || '')))) out.push({ no: String('１２３４'.indexOf(m[1]) >= 0 ? '１２３４'.indexOf(m[1]) + 1 : m[1]), text: m[2].trim() });
    return out;
  }

  /* 這題用哪一種、資料是否和題庫一致（答案一樣、句子對得上） */
  function buildModel(q) {
    if (!q) return null;
    var id = qid(q), D = data(), official = String(q['答案'] || '').trim();
    var text = String(q['完整題目與選項'] || ''), nt = norm(text);
    var opts = parseOptions(text);
    if (!official || opts.length !== 4) return null;
    var stem = text.split(/[（(][1-4１-４][)）]/)[0].replace(/^\s*\d+\.\s*/, '');
    var c = D.cmp && D.cmp[id];
    if (c) {
      if (c.answer !== official) return null;
      for (var i = 0; i < c.pairs.length; i++) {
        var p = c.pairs[i];
        if (nt.indexOf(norm(p.a.ctx)) < 0 || nt.indexOf(norm(p.b.ctx)) < 0) return null;
      }
      return { type: 'cmp', uid: id, grade: q['年級'] || '', cat: q['類別'] || '', pairs: c.pairs, mode: c.mode, ask: c.ask, answer: official, options: opts, stem: stem };
    }
    var k = D.pick && D.pick[id];
    if (k) {
      if (k.answer !== official) return null;
      if (k.mode === 'fill') {
        for (var j = 0; j < k.slots.length; j++) {
          for (var t = 0; t < k.slots[j].cands.length; t++) if (text.indexOf(k.slots[j].cands[t]) < 0) return null;
        }
      } else if (text.indexOf('「' + k.target + '」') < 0) return null;
      return { type: k.mode, uid: id, grade: q['年級'] || '', cat: q['類別'] || '', d: k, answer: official, options: opts, stem: stem };
    }
    return null;
  }
  function qualifies(q) { return !!buildModel(q); }

  function pickNext(all, current) {
    var pool = (all || []).filter(function (q) { return qid(q) !== String(current.uid) && (q['類別'] || '') === current.cat && qualifies(q); });
    if (!pool.length) pool = (all || []).filter(function (q) { return qid(q) !== String(current.uid) && qualifies(q); });
    if (!pool.length) return null;
    var same = pool.filter(function (q) { return (q['年級'] || '') === current.grade; });
    var list = same.length ? same : pool;
    return list[Math.floor(Math.random() * list.length)];
  }

  /* 一組讀音的道理（給「比意思」那一步） */
  function pairRule(p) {
    var sameChar = p.a.ch === p.b.ch;
    if (sameChar && !p.same) return '同一個字，意思不同，讀音就不同（音隨義轉）。';
    if (sameChar && p.same) return '同一個字，這兩個意思讀音一樣。';
    if (!sameChar && p.same) return '兩個不同的字讀音一樣（同音字），所以要看意思選字。';
    return '兩個字長得像或聽起來像，但讀音不一樣。';
  }

  var core = { buildModel: buildModel, qualifies: qualifies, parseOptions: parseOptions, pairRule: pairRule, setData: setData, pickNext: pickNext };
  if (typeof module !== 'undefined' && module.exports) { module.exports = core; return; }

  /* ═══ 畫面 ═══════════════════════════════════════════ */
  var S = null, audio = null;
  var cfg = { getAll: function () { return []; } };
  function configure(o) { for (var k in o) cfg[k] = o[k]; }
  function getAll() { return cfg.getAll() || []; }
  var STEPS = {
    cmp: ['先猜猜看', '查讀音', '比意思', '想法改了嗎', '原題挑戰'],
    fill: ['先猜猜看', '讀讀看', '想意思選字', '部首和意思', '原題挑戰'],
    which: ['先猜猜看', '讀讀看', '看意思判斷', '部首和意思', '原題挑戰']
  };
  var SRC = {
    cmp: ['準備活動：先讓學生試讀、說自己的想法，不急著糾正', '查字典：用辭典的注音確認讀音', '發現「字音和字義有關」：意思不同，讀音可能不同；讀音一樣，就要看意思', '「查完字典後，你要修正之前的想法或補充嗎？」', '統整：用學到的方法判斷原題'],
    fill: ['準備活動：先讓學生觀察、說出自己的想法', '這些字讀音一樣（或很像），只聽聲音分不出來', '《國字會分工》：先想這個詞是什麼意思，再選字', '「一邊是意思，一邊是聲音」：部首表示字的意思', '統整：用學到的方法判斷原題'],
    which: ['準備活動：先讓學生觀察、說出自己的想法', '這些注音讀音一樣，只聽聲音分不出來', '《國字會分工》：想想每個詞的意思，這裡該用哪個字', '「一邊是意思，一邊是聲音」：部首表示字的意思', '統整：用學到的方法判斷原題']
  };

  function h(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  /* 只標「」裡有字的部分；題幹常有空的「」（如「下面句子「」中的注音…是「故」」），不可一路連到下一個 」 */
  function mark(s) { return h(s).replace(/「([^「」]+)」/g, '<span class="gp-mark">$1</span>').replace(/(□+)/g, '<span class="gp-mark">$1</span>'); }
  function wordMark(w, ch) { return h(w).replace(h(ch), '<span class="gp-mark">' + h(ch) + '</span>'); }
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
    if (!el) { el = document.createElement('div'); el.id = 'gp-overlay'; el.setAttribute('role', 'dialog'); el.setAttribute('aria-modal', 'true'); document.body.appendChild(el); }
    if (!el._sdBound) { el.addEventListener('click', onClick); el._sdBound = true; }
    return el;
  }
  function findQ(uid) { var a = getAll(); for (var i = 0; i < a.length; i++) if (qid(a[i]) === String(uid)) return a[i]; return null; }
  function shuffle(a) { a = a.slice(); for (var i = a.length - 1; i > 0; i--) { var j = Math.floor(Math.random() * (i + 1)); var t = a[i]; a[i] = a[j]; a[j] = t; } return a; }

  function open(uid) {
    var q = findQ(uid), m = q && buildModel(q);
    if (!m) return;
    ['GuidedPY', 'GuidedCZ', 'GuidedCY', 'GuidedXJ', 'GuidedApply'].forEach(function (k) { if (root[k] && root[k].close) root[k].close(); });
    var n = m.type === 'cmp' ? m.pairs.length : m.type === 'fill' ? m.d.slots.length : m.d.rows.length;
    S = { q: q, m: m, step: 0, n: n, guess: {}, pick: {}, tries: {}, done: {}, first: {}, hint: {},
          opts: m.type === 'fill' ? m.d.slots.map(function (s) { return shuffle(s.cands); }) : null,
          finalPick: null, finalOk: null, finalFirst: null };
    document.body.classList.add('gp-open');
    render();
  }
  function close() { var el = document.getElementById('gp-overlay'); if (el && S) el.innerHTML = ''; document.body.classList.remove('gp-open'); S = null; }

  function all(obj) { for (var i = 0; i < S.n; i++) if (!obj[i]) return false; return true; }
  function stepDone() {
    var t = S.m.type;
    if (S.step === 0) return all(S.guess);
    if (S.step === 1) return t === 'cmp' ? all(S.done) : true;
    if (S.step === 2) return t === 'cmp' ? true : all(S.done);
    if (S.step === 3) return true;
    return S.finalOk === true;
  }

  function pairHead(p) {
    var A = p.a.word && p.a.word !== p.a.ch ? wordMark(p.a.word, p.a.ch) : mark(p.a.ctx);
    var B = p.b.word && p.b.word !== p.b.ch ? wordMark(p.b.word, p.b.ch) : mark(p.b.ctx);
    return '<div class="gp-xj-pair"><span class="gp-xj-w">' + A + '</span><span class="gp-xj-sep">／</span><span class="gp-xj-w">' + B + '</span></div>';
  }
  function twoBtns(act, i, v, done, right, labA, labB) {
    function b(val, lab) {
      var cls = v === val ? (right == null ? ' is-on' : (val === right ? ' is-ok' : ' is-no')) : '';
      return '<button class="gp-choice sm' + cls + '" data-act="' + act + '" data-i="' + i + '" data-v="' + val + '"' + (done ? ' disabled' : '') + '>' + lab + '</button>';
    }
    return '<div class="gp-choices">' + b('same', labA || '讀音一樣') + b('diff', labB || '讀音不一樣') + '</div>';
  }

  /* ── 讀音比一比 ── */
  function cmpView() {
    var m = S.m, ps = m.pairs;
    if (S.step === 0) {
      return '<p class="gp-ask">先自己讀讀看。每一組<span class="gp-mark">標出來的字</span>讀音一樣嗎？憑感覺選就好。</p>' +
        (m.mode === 'stem' ? '<div class="gp-q gp-sent">' + mark(ps[0].a.ctx) + '</div>' : '') +
        ps.map(function (p, i) { return '<div class="gp-srow">' + pairHead(p) + twoBtns('guess', i, S.guess[i], false, null) + '</div>'; }).join('');
    }
    if (S.step === 1) {
      return '<p class="gp-ask">查字典看注音，讀音一樣嗎？</p>' + ps.map(function (p, i) {
        return '<div class="gp-srow">' + pairHead(p) + '<div class="gp-xj-read"><span>' + h(p.a.word || p.a.ch) + ' <b class="gp-bopo">' + h(p.a.read) + '</b></span><span>' + h(p.b.word || p.b.ch) + ' <b class="gp-bopo">' + h(p.b.read) + '</b></span></div>' +
          twoBtns('read', i, S.pick[i], S.done[i], p.same ? 'same' : 'diff') + (S.hint[i] ? '<div class="gp-hint">' + S.hint[i] + '</div>' : '') + '</div>';
      }).join('') + '<div class="gp-dict-src" style="margin-top:8px">注音：教育部《國語辭典簡編本》（以語詞的注音為準）</div>';
    }
    if (S.step === 2) {
      return '<p class="gp-ask">讀音為什麼一樣或不一樣？看看兩個語詞的意思。</p>' + ps.map(function (p) {
        return '<div class="gp-srow">' + pairHead(p) +
          '<div class="gp-ex"><b>' + h(p.a.word || p.a.ch) + '</b>（' + h(p.a.read) + '）：' + h(p.a.def || '—') + '</div>' +
          '<div class="gp-ex"><b>' + h(p.b.word || p.b.ch) + '</b>（' + h(p.b.read) + '）：' + h(p.b.def || '—') + '</div>' +
          '<div class="gp-good" style="margin-top:6px">' + h(pairRule(p)) + '</div></div>';
      }).join('') + '<div class="gp-rule">字音和字義有關：先想意思，再判斷讀音或選字。</div>';
    }
    if (S.step === 3) {
      var right = 0;
      var rows = ps.map(function (p, i) {
        var said = S.guess[i] === 'same' ? '讀音一樣' : '讀音不一樣', truth = p.same ? '讀音一樣' : '讀音不一樣';
        var ok = (S.guess[i] === 'same') === p.same; if (ok) right++;
        return '<tr><td>' + h(p.a.word || p.a.ch) + '（' + h(p.a.read) + '）／' + h(p.b.word || p.b.ch) + '（' + h(p.b.read) + '）</td><td>' + said + '</td><td>' + truth + (ok ? ' ✓' : '') + '</td></tr>';
      }).join('');
      var p0 = ps.filter(function (p) { return p.a.ch === p.b.ch && !p.same; })[0] || ps[0];
      return '<p class="gp-ask">一開始你的想法，和現在比比看：</p><table class="gp-tbl"><thead><tr><th>語詞</th><th>你一開始覺得</th><th>字典說</th></tr></thead><tbody>' + rows + '</tbody></table>' +
        '<p class="gp-note">' + (right === ps.length ? '你一開始就全部猜對了，很厲害！現在你還能說出為什麼。' : '想法改變了也很好，這就是在學習。') + '</p>' +
        '<div class="gp-frame"><div class="gp-frame-lbl">🗣️ 用這句話說一次給同學聽</div>「' + h(p0.a.word || p0.a.ch) + '」讀 ' + h(p0.a.read) + '，「' + h(p0.b.word || p0.b.ch) + '」讀 ' + h(p0.b.read) + '，因為' + h(pairRule(p0).replace(/。$/, '')) + '。</div>';
    }
    return finalView();
  }

  /* ── 看意思選字 ── */
  function fillView() {
    var d = S.m.d, sl = d.slots;
    function sentence(showAns) {
      var k = -1;
      return h(d.sentence).replace(/□/g, function () { k++; return '<span class="gp-mark' + (showAns && S.done[k] ? ' gp-swap' : '') + '">' + (showAns && S.done[k] ? h(sl[k].ch) : '□<sub>' + (k + 1) + '</sub>') + '</span>'; });
    }
    if (S.step === 0) {
      return '<div class="gp-q gp-sent">' + sentence(false) + '</div><p class="gp-ask">先憑感覺猜猜看：每個 □ 要填哪個字？</p>' +
        sl.map(function (s, i) {
          return '<div class="gp-srow"><div class="gp-stext">□<sub>' + (i + 1) + '</sub></div><div class="gp-choices">' + S.opts[i].map(function (c) {
            return '<button class="gp-choice gp-big' + (S.guess[i] === c ? ' is-on' : '') + '" data-act="guess" data-i="' + i + '" data-v="' + h(c) + '">' + h(c) + '</button>';
          }).join('') + '</div></div>';
        }).join('');
    }
    if (S.step === 1) {
      var chars = []; sl.forEach(function (s) { s.cands.forEach(function (c) { if (chars.indexOf(c) < 0) chars.push(c); }); });
      return '<p class="gp-ask">讀讀看這些字：</p><div class="gp-xj-read" style="flex-wrap:wrap">' + chars.map(function (c) {
        return '<span class="gp-word">' + h(c) + ' <b class="gp-bopo">' + h((d.reads[c] || []).join('／')) + '</b></span>';
      }).join('') + '</div><div class="gp-rule" style="margin-top:12px">讀音一樣（或很像），只聽聲音分不出來！</div><p class="gp-note">那要怎麼選？下一步：先想這個詞是什麼意思。</p>' +
        '<div class="gp-dict-src">注音：教育部《國語辭典簡編本》</div>';
    }
    if (S.step === 2) {
      return '<div class="gp-q gp-sent">' + sentence(true) + '</div><p class="gp-ask">讀整句，想想這裡要說什麼意思，再選字。</p>' + sl.map(function (s, i) {
        return '<div class="gp-srow"><div class="gp-stext">□<sub>' + (i + 1) + '</sub>　「' + h(s.word.replace(s.ch, '？')) + '」</div><div class="gp-choices">' + S.opts[i].map(function (c) {
          var on = S.pick[i] === c, cls = on ? (c === s.ch ? ' is-ok' : ' is-no') : '';
          return '<button class="gp-choice gp-big' + cls + '" data-act="pick" data-i="' + i + '" data-v="' + h(c) + '"' + (S.done[i] ? ' disabled' : '') + '>' + h(c) + '</button>';
        }).join('') + '</div>' + (S.hint[i] ? '<div class="gp-hint">' + S.hint[i] + '</div>' : '') + '</div>';
      }).join('');
    }
    if (S.step === 3) {
      return '<p class="gp-ask">看看部首和意思的關係：</p><ul class="gp-rule-list">' + sl.map(function (s) {
        return '<li><b>' + h(s.ch) + '</b>（' + h(s.word) + '）' + (s.rad ? '：' + h(s.rad) + '部' + (s.radMean ? '，' + h(s.radMean) : '') : '') + '<div class="gp-ex">簡編本「' + h(s.word) + '」：' + h(s.def || '—') + '</div></li>';
      }).join('') + '</ul><div class="gp-rule">讀音一樣時，先想意思，再看部首選字。</div>';
    }
    return finalView();
  }

  function whichView() {
    var d = S.m.d, rows = d.rows, X = d.target;
    function rowText(r, reveal) { return h(r.text).replace('「' + h(r.bopo) + '」', '<span class="gp-mark">' + (reveal && r.isX ? h(X) : '「' + h(r.bopo) + '」') + '</span>'); }
    if (S.step === 0) {
      return '<p class="gp-ask">先猜猜看：哪一句的注音要寫成「<span class="gp-mark">' + h(X) + '</span>」？每一句選「是」或「不是」。</p>' + rows.map(function (r, i) {
        return '<div class="gp-srow"><div class="gp-stext">(' + (i + 1) + ') ' + rowText(r, false) + '</div>' + twoBtns('guess', i, S.guess[i], false, null, '是「' + X + '」', '不是') + '</div>';
      }).join('');
    }
    if (S.step === 1) {
      return '<p class="gp-ask">讀讀看：「' + h(X) + '」讀 <b class="gp-bopo">' + h(d.read) + '</b>，四句的注音都一樣。</p><div class="gp-rule">讀音一樣，只聽聲音分不出來！</div><p class="gp-note">下一步：想想每一句的詞是什麼意思，這裡該不該用「' + h(X) + '」。</p>';
    }
    if (S.step === 2) {
      return '<p class="gp-ask">想想每一句的意思：這裡要寫「' + h(X) + '」嗎？</p>' + rows.map(function (r, i) {
        return '<div class="gp-srow"><div class="gp-stext">(' + (i + 1) + ') ' + rowText(r, S.done[i]) + '</div>' + twoBtns('pick', i, S.pick[i], S.done[i], r.isX ? 'same' : 'diff', '是「' + X + '」', '不是') +
          (S.hint[i] ? '<div class="gp-hint">' + S.hint[i] + '</div>' : '') + '</div>';
      }).join('');
    }
    if (S.step === 3) {
      return '<p class="gp-ask">「' + h(X) + '」的意思和部首：</p><div class="gp-good">' + h(X) + (d.rad ? '：' + h(d.rad) + '部' + (d.radMean ? '，' + h(d.radMean) : '') : '') +
        '<div class="gp-ex">簡編本「' + h(d.word) + '」：' + h(d.def || '—') + '</div></div><div class="gp-rule" style="margin-top:10px">讀音一樣時，先想意思，再看部首選字。</div>';
    }
    return finalView();
  }

  function finalView() {
    var m = S.m;
    var btns = m.options.map(function (o) {
      var on = S.finalPick === o.no, cls = on ? (o.no === m.answer ? ' is-ok' : ' is-no') : '';
      return '<button class="gp-choice' + cls + '" data-act="final" data-v="' + h(o.no) + '"' + (S.finalOk ? ' disabled' : '') + '>(' + h(o.no) + ') ' + mark(o.text) + '</button>';
    }).join('');
    var msg = '';
    if (S.finalOk) msg = '<div class="gp-good">答對了！先想意思，再判斷讀音或選字。</div>';
    else if (S.finalPick != null) msg = '<div class="gp-hint">再想想：回頭看看字典的注音和意思。</div>';
    return '<p class="gp-ask">現在自己做做看原題：</p><div class="gp-q">' + mark(m.stem) + '</div><div class="gp-choices col">' + btns + '</div>' + msg + (S.finalOk ? summary() : '');
  }
  function summary() {
    var n = S.n, g = 0, f = 0, m = S.m;
    for (var i = 0; i < n; i++) {
      if (S.first[i]) f++;
      if (m.type === 'cmp' ? ((S.guess[i] === 'same') === m.pairs[i].same) : m.type === 'fill' ? S.guess[i] === m.d.slots[i].ch : ((S.guess[i] === 'same') === m.d.rows[i].isX)) g++;
    }
    var next = pickNext(getAll(), m);
    return '<div class="gp-sum"><div class="gp-frame-lbl">📋 這一題的學習紀錄（給老師看）</div><div>一開始就猜對：' + g + ' / ' + n + '</div>' +
      '<div>' + (m.type === 'cmp' ? '查讀音一次就判斷對' : '一次就選對') + '：' + f + ' / ' + n + '</div><div>原題：' + (S.finalFirst ? '第一次就答對' : '試了不只一次才答對') + '</div></div>' +
      (hasApply() ? '<button class="gp-next glow" data-act="apply" data-next="' + h(next ? qid(next) : '') + '">換個句子用用看 →</button> ' : '') +
      (next ? '<button class="' + (hasApply() ? 'gp-back' : 'gp-next glow') + '" data-act="another" data-id="' + h(qid(next)) + '">換一題再試試 →</button>' : '');
  }
  function hasApply() { return !!(root.GuidedApply && root.GuidedApply.has(S.m.uid)); }

  function render() {
    if (!S) return;
    var el = overlay(), t = S.m.type, steps = STEPS[t];
    var bar = steps.map(function (x, i) { return '<span class="gp-step' + (i === S.step ? ' is-now' : i < S.step ? ' is-done' : '') + '">' + (i < S.step ? '✓' : (i + 1)) + ' ' + x + '</span>'; }).join('');
    var body = t === 'cmp' ? cmpView() : t === 'fill' ? fillView() : whichView();
    var done = stepDone(), last = steps.length - 1;
    var title = t === 'cmp' ? '讀音比一比' : '看意思選字';
    var src = t === 'cmp' ? '學扶國語教材《語詞一家人（一）》（3-1-3）、《國字會分工》（3-1-5）' : '學扶國語教材 三年級《國字會分工》（3-1-5）';
    el.innerHTML = '<div class="gp-panel" data-kind="sd">' +
      '<div class="gp-top"><div class="gp-title">✏️ 一起動腦：' + title + '</div><button class="gp-close" data-act="close" aria-label="離開">✕ 離開</button></div>' +
      '<div class="gp-steps">' + bar + '</div>' +
      '<div class="gp-src"><b>📘 教學依據</b>' + h(src) + '<span class="gp-src-step">這一步：' + h(SRC[t][S.step]) + '</span></div>' +
      '<div class="gp-body">' + body + '</div>' +
      '<div class="gp-foot">' + (S.step > 0 && S.step < last ? '<button class="gp-back" data-act="back">← 上一步</button>' : '<span></span>') +
      (S.step < last ? '<button class="gp-next' + (done ? ' glow' : '') + '" data-act="next"' + (done ? '' : ' disabled') + '>下一步 →</button>' : '') + '</div></div>';
  }

  function answer(i, v, right, hint1, hint2) {
    if (S.done[i]) return;
    S.pick[i] = v;
    if (v === right) { S.done[i] = true; if (!S.tries[i]) S.first[i] = true; S.hint[i] = ''; tone(true); }
    else {
      S.tries[i] = (S.tries[i] || 0) + 1; tone(false);
      if (S.tries[i] === 1) S.hint[i] = hint1; else { S.done[i] = true; S.pick[i] = right; S.hint[i] = hint2; }
    }
  }

  function onClick(ev) {
    if (!S) return;
    var b = ev.target.closest('[data-act]');
    if (!b || b.disabled) return;
    var act = b.getAttribute('data-act'), v = b.getAttribute('data-v'), i = Number(b.getAttribute('data-i')), m = S.m;
    if (act === 'close') return close();
    if (act === 'next' && stepDone()) { S.step++; return render(); }
    if (act === 'back') { S.step--; return render(); }
    if (act === 'another') { var nid = b.getAttribute('data-id'); close(); return open(nid); }
    if (act === 'apply') { var au = S.m.uid, an = b.getAttribute('data-next'); close(); return root.GuidedApply.open(au, { mod: 'GuidedSD', next: an }); }
    if (act === 'guess') { S.guess[i] = v; return render(); }
    if (act === 'read') {
      var p = m.pairs[i];
      answer(i, v, p.same ? 'same' : 'diff', '仔細比一比兩個注音，聲調也要看喔。', '「' + h(p.a.word || p.a.ch) + '」讀 ' + h(p.a.read) + '，「' + h(p.b.word || p.b.ch) + '」讀 ' + h(p.b.read) + '，' + (p.same ? '一樣' : '不一樣') + '。老師幫你選好了。');
      return render();
    }
    if (act === 'pick' && m.type === 'fill') {
      var s = m.d.slots[i];
      answer(i, v, s.ch, '想想「' + h(s.word.replace(s.ch, '？')) + '」是什麼意思？哪個字的部首和這個意思有關？', '這裡要寫「' + h(s.ch) + '」（' + h(s.word) + '），老師幫你選好了。');
      return render();
    }
    if (act === 'pick' && m.type === 'which') {
      var r = m.d.rows[i];
      answer(i, v, r.isX ? 'same' : 'diff', '想想這一句的詞是什麼意思？和「' + h(m.d.target) + '」的意思有關嗎？', r.isX ? '這一句要寫「' + h(m.d.target) + '」（' + h(m.d.word) + '），老師幫你選好了。' : '這一句不是「' + h(m.d.target) + '」，老師幫你選好了。');
      return render();
    }
    if (act === 'final') {
      var ok = v === m.answer;
      if (S.finalPick == null) { S.finalFirst = ok; if (root.ClassResults) root.ClassResults.add(m.uid, 'guided', ok ? 1 : 0, 1); }
      S.finalPick = v; S.finalOk = ok; tone(ok);
      return render();
    }
  }

  document.addEventListener('keydown', function (e) { if (S && e.key === 'Escape') close(); });
  root.GuidedSD = { open: open, close: close, configure: configure, qualifies: qualifies, buildModel: buildModel };
})(typeof window !== 'undefined' ? window : globalThis);
