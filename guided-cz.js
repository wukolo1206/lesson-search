/* 一起動腦：錯別字（依學扶教材 3-1-5《國字會分工》：部首表示字的意思；
 * 第 2 步借用 4-1-5《國字會合作》猜燈謎的解題步驟：找關鍵詞 → 推測用字）
 * 先找找看（不評對錯）→ 這裡要寫哪個字 → 字和意思配對 → 找到規則 → 想法改了嗎 → 原題挑戰
 * 設計文件：字音形互動引導設計_教材依據.md
 * 樣式沿用 guided-py.css 的 gp-*。
 */
(function (root) {
  'use strict';

  /* ═══ 核心（可在 Node 測試） ═══════════════════════════ */
  function bracketChar(s) {
    var m = String(s || '').match(/【(.+?)】/);
    return m ? m[1] : '';
  }

  /* 「應為「求救」—救(攴部，搭救)；就是…」→ { ch:'救', radical:'攴部', meaning:'搭救', contrast:'' } */
  function parseExplain(text) {
    var s = String(text || '');
    var parts = s.split('；');
    var main = parts[0];
    var contrast = parts.slice(1).join('；').trim();
    if (/[A-Za-z]/.test(contrast)) contrast = '';
    var re = /([^\s—：「」，,；;()（）])[（(]([^）)]*)[）)]/g;
    var m, last = null;
    while ((m = re.exec(main))) last = m;
    if (!last) return null;
    var inside = last[2].split(/[，,]/);
    var radical = '', meaning = '';
    if (/部$/.test(inside[0].trim())) {
      radical = inside[0].trim();
      meaning = inside.slice(1).join('，').trim();
    } else {
      meaning = last[2].trim();
    }
    if (/部/.test(meaning)) meaning = '';      // 如「祕是示部」不是意思
    return { ch: last[1], radical: radical, meaning: meaning, contrast: contrast };
  }

  function norm(t) { return String(t || '').replace(/[「」【】\s，,。！!？?、；;：:]/g, ''); }

  /* 細目這一句對應原題第幾個選項：用整句比對（同一個字可能出現在兩個選項，如第 173 題的驕／嬌） */
  function optionIndexOf(qText, sentence, ch) {
    var re = /[（(]([1-4１-４])[)）]([^(（]*)/g, m, target = norm(sentence), hit = null;
    while ((m = re.exec(String(qText || '')))) {
      var o = m[2];
      if (o.indexOf('「' + ch + '」') < 0) continue;
      var no = norm(o);
      if (no && (target.indexOf(no) >= 0 || no.indexOf(target) >= 0)) {
        var k = '１２３４'.indexOf(m[1]);
        hit = String(k >= 0 ? k + 1 : m[1]);
        break;
      }
    }
    return hit;
  }

  function buildModel(q) {
    var rows = (q && q._cz) || [];
    var qText = String((q && q['完整題目與選項']) || '');
    var official = String((q && q['答案']) || '').trim();
    if (rows.length < 3 || !qText || !official) return null;
    var items = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var written = bracketChar(r['錯誤寫法']);
      if (!written) return null;
      var opt = optionIndexOf(qText, r['錯誤寫法'], written);
      if (!opt) return null;                                   // 細目和原題對不上
      var isRight = String(r['正誤'] || '').indexOf('✓') >= 0;
      if (isRight !== (opt === official)) return null;        // 細目對錯和官方答案不一致
      var ex = parseExplain(r['字義說明']);
      if (!ex) return null;
      var correct = isRight ? written : ex.ch;
      if (!isRight && (correct === written || String(r['正確詞語'] || '').indexOf(correct) < 0)) return null;
      items.push({
        id: i, opt: opt, sentence: String(r['錯誤寫法']), written: written, correct: correct,
        isRight: isRight, word: String(r['正確詞語'] || ''),
        radical: ex.radical, meaning: ex.meaning, contrast: ex.contrast
      });
    }
    var wrong = items.filter(function (x) { return !x.isRight; });
    if (wrong.length < 2) return null;
    var matchable = items.filter(function (x) { return x.meaning; });
    var seen = {};
    matchable = matchable.filter(function (x) { if (seen[x.correct] || seen['m' + x.meaning]) return false; seen[x.correct] = seen['m' + x.meaning] = 1; return true; });
    return {
      uid: q['題目ID'] || q['題號'],
      grade: q['年級'] || '',
      items: items,
      wrong: wrong,
      matchable: matchable.length >= 2 ? matchable : [],
      answer: official,
      tip: String(q['記憶訣竅'] || '').trim()
    };
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

  function shuffle(a) {
    a = a.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var t = a[i]; a[i] = a[j]; a[j] = t;
    }
    return a;
  }

  var core = { parseExplain: parseExplain, buildModel: buildModel, qualifies: qualifies, pickNext: pickNext };
  if (typeof module !== 'undefined' && module.exports) { module.exports = core; return; }

  /* ═══ 畫面 ═══════════════════════════════════════════ */
  var STEPS = ['先找找看', '這裡要寫哪個字', '字和意思配對', '找到規則', '想法改了嗎', '原題挑戰'];
  var SOURCE = '學扶國語教材 三年級《國字會分工》（3-1-5）';
  var STEP_SRC = [
    '準備活動：先讓學生觀察、說出自己的發現，不急著給答案',
    '四年級《國字會合作》（4-1-5）猜燈謎的步驟：找出關鍵詞 → 推測用字',
    '發展活動（一）（四）：查字典，說一說部首和字的意思有什麼關係',
    '核心口訣「一邊是意思，一邊是聲音」：部首表示字的意思',
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
  /* 句子裡【X】改成醒目的字 */
  function sent(s, ch) {
    return h(s).replace(/【(.+?)】/, '<span class="gp-mark">' + h(ch == null ? '$1' : ch) + '</span>');
  }

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
    if (!el._czBound) { el.addEventListener('click', onClick); el._czBound = true; }
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
    if (root.GuidedCY && root.GuidedCY.close) root.GuidedCY.close();
    if (root.GuidedXJ && root.GuidedXJ.close) root.GuidedXJ.close();
    S = {
      q: q, m: m, step: 0,
      guess: {},                 // item.id -> 'ok' | 'odd'
      pick: {}, tries: {}, firstTry: {}, done: {}, hint: {},
      opts: {},
      mSel: null, mDone: {}, mTries: {}, mHint: '',
      mOrder: shuffle(m.matchable), mMeanings: shuffle(m.matchable.map(function (x) { return x.meaning; })),
      finalPick: null, finalOk: null, finalFirst: null
    };
    m.wrong.forEach(function (x) { S.opts[x.id] = shuffle([x.written, x.correct]); });
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
    switch (S.step) {
      case 0: return S.m.items.every(function (x) { return S.guess[x.id]; });
      case 1: return S.m.wrong.every(function (x) { return S.done[x.id]; });
      case 2: return !S.m.matchable.length || S.m.matchable.every(function (x) { return S.mDone[x.id]; });
      case 3: case 4: return true;
      case 5: return S.finalOk === true;
    }
    return false;
  }

  function viewGuess() {
    return '<p class="gp-ask">先自己讀讀看。每一句裡<span class="gp-mark">標出來的字</span>，用得對嗎？覺得怪怪的就按「怪怪的」。</p>' +
      S.m.items.map(function (x) {
        var g = S.guess[x.id];
        return '<div class="gp-srow"><div class="gp-stext">(' + h(x.opt) + ') ' + sent(x.sentence) + '</div>' +
          '<div class="gp-choices">' +
          '<button class="gp-choice sm' + (g === 'ok' ? ' is-on' : '') + '" data-act="guess" data-id="' + x.id + '" data-v="ok">用得對</button>' +
          '<button class="gp-choice sm' + (g === 'odd' ? ' is-on' : '') + '" data-act="guess" data-id="' + x.id + '" data-v="odd">怪怪的</button>' +
          '</div></div>';
      }).join('') +
      (stepDone() ? '<p class="gp-note">好，先記住你的想法，等一下我們來驗證。</p>' : '');
  }

  function viewPickChar() {
    return '<p class="gp-ask">這幾句的字其實寫錯了。先讀整句，找出它要說的意思，再選出要寫哪個字。</p>' +
      S.m.wrong.map(function (x) {
        var done = S.done[x.id];
        var btns = S.opts[x.id].map(function (c) {
          var on = S.pick[x.id] === c;
          var cls = on ? (c === x.correct ? ' is-ok' : ' is-no') : '';
          return '<button class="gp-choice gp-big' + cls + '" data-act="char" data-id="' + x.id + '" data-v="' + h(c) + '"' + (done ? ' disabled' : '') + '>' + h(c) + '</button>';
        }).join('');
        return '<div class="gp-srow"><div class="gp-stext">' + sent(x.sentence, done ? x.correct : '？') + '</div>' +
          '<div class="gp-choices">' + btns + '</div>' +
          (S.hint[x.id] ? '<div class="gp-hint">' + S.hint[x.id] + '</div>' : '') + '</div>';
      }).join('');
  }

  function viewMatch() {
    if (!S.m.matchable.length) {
      return '<p class="gp-ask">說一說：每個正確的字，部首和它的意思有什麼關係？</p>' + viewRadicalList() +
        '<p class="gp-note">這一題的資料沒有寫出每個字的意思，請老師帶大家口頭討論。</p>';
    }
    var chars = S.mOrder.map(function (x) {
      var done = S.mDone[x.id];
      return '<button class="gp-card' + (S.mSel === x.id ? ' is-sel' : '') + (done ? ' is-in' : '') + '" data-act="mchar" data-id="' + x.id + '"' + (done ? ' disabled' : '') + '>' +
        h(x.correct) + (done && x.radical ? '<small class="gp-rad">' + h(x.radical) + '</small>' : '') + '</button>';
    }).join('');
    var means = S.mMeanings.map(function (mn) {
      var owner = S.m.matchable.filter(function (x) { return x.meaning === mn; })[0];
      var done = S.mDone[owner.id];
      return '<button class="gp-choice sm' + (done ? ' is-ok' : '') + '" data-act="mmean" data-v="' + h(mn) + '"' + (done ? ' disabled' : '') + '>' +
        (done ? h(owner.correct) + '：' : '') + h(mn) + '</button>';
    }).join('');
    return '<p class="gp-ask">點一個字，再點它的意思。配對成功會出現它的部首，想一想：部首和意思有什麼關係？</p>' +
      '<div class="gp-pile">' + chars + '</div>' +
      '<div class="gp-choices col">' + means + '</div>' +
      (S.mHint ? '<div class="gp-hint">' + S.mHint + '</div>' : '');
  }

  function viewRadicalList() {
    return '<ul class="gp-rule-list">' + S.m.items.map(function (x) {
      return '<li><b>' + h(x.correct) + '</b>' + (x.radical ? '（' + h(x.radical) + '）' : '') +
        (x.meaning ? '：' + h(x.meaning) : '') + (x.word ? '<span class="gp-ex">（' + h(x.word) + '）</span>' : '') +
        (x.contrast ? '<div class="gp-ex">對照：' + h(x.contrast) + '</div>' : '') + '</li>';
    }).join('') + '</ul>';
  }

  function viewRule() {
    return '<div class="gp-rule">部首會告訴我們字的意思。<br>寫字前，先想這個詞是什麼意思。</div>' +
      viewRadicalList() +
      (S.m.tip ? '<div class="gp-tip">🧠 記憶訣竅：' + h(S.m.tip) + '</div>' : '') +
      '<p class="gp-note">很多字讀音很像，不能只聽聲音選字。像教材說的「一邊是意思，一邊是聲音」，看部首就知道它管什麼意思。</p>';
  }

  function viewRevise() {
    var right = 0;
    var rows = S.m.items.map(function (x) {
      var said = S.guess[x.id] === 'ok' ? '用得對' : '怪怪的';
      var truth = x.isRight ? '用得對' : '寫錯了';
      var ok = (S.guess[x.id] === 'ok') === x.isRight;
      if (ok) right++;
      return '<tr><td>(' + h(x.opt) + ') ' + sent(x.sentence) + '</td><td>' + said + '</td><td>' + truth + (ok ? ' ✓' : '') + '</td></tr>';
    }).join('');
    var ex = S.m.wrong[0];
    return '<p class="gp-ask">一開始你的判斷，和現在比比看：</p>' +
      '<table class="gp-tbl"><thead><tr><th>句子</th><th>你一開始覺得</th><th>其實</th></tr></thead><tbody>' + rows + '</tbody></table>' +
      '<p class="gp-note">' + (right === S.m.items.length ? '你一開始就全部看出來了，很厲害！現在你還能說出為什麼。' : '想法改變了也很好，這就是在學習。') + '</p>' +
      '<div class="gp-frame"><div class="gp-frame-lbl">🗣️ 用這句話說一次給同學聽</div>' +
      '「' + h(ex.word) + '」的「' + h(ex.correct) + '」' + (ex.radical ? '是<u>' + h(ex.radical) + '</u>，' : '') +
      (ex.meaning ? '意思是<u>' + h(ex.meaning) + '</u>，' : '') + '所以不能寫成「' + h(ex.written) + '」。</div>';
  }

  function viewFinal() {
    var text = S.q['完整題目與選項'] || '';
    var stem = text.split(/[（(][1-4１-４][)）]/)[0];
    var btns = S.m.items.slice().sort(function (a, b) { return a.opt - b.opt; }).map(function (x) {
      var on = String(S.finalPick) === String(x.opt);
      var cls = on ? (String(x.opt) === S.m.answer ? ' is-ok' : ' is-no') : '';
      return '<button class="gp-choice' + cls + '" data-act="final" data-v="' + h(x.opt) + '"' + (S.finalOk ? ' disabled' : '') + '>(' + h(x.opt) + ') ' + sent(x.sentence) + '</button>';
    }).join('');
    var msg = '';
    if (S.finalOk) msg = '<div class="gp-good">答對了！你用「先想意思、再看部首」解出了考古題。</div>';
    else if (S.finalPick != null) msg = '<div class="gp-hint">再想想：每一句要說什麼意思？那個意思該用哪個部首的字？</div>';
    return '<p class="gp-ask">現在自己做做看原題：</p><div class="gp-q">' + h(stem) + '</div>' +
      '<div class="gp-choices col">' + btns + '</div>' + msg + (S.finalOk ? viewSummary() : '');
  }

  function viewSummary() {
    var n = S.m.wrong.length;
    var first = S.m.wrong.filter(function (x) { return S.firstTry[x.id]; }).length;
    var guessRight = S.m.items.filter(function (x) { return (S.guess[x.id] === 'ok') === x.isRight; }).length;
    var next = pickNext(getAll(), S.m);
    return '<div class="gp-sum"><div class="gp-frame-lbl">📋 這一題的學習紀錄（給老師看）</div>' +
      '<div>一開始就判斷對：' + guessRight + ' / ' + S.m.items.length + ' 句</div>' +
      '<div>選字一次就選對：' + first + ' / ' + n + ' 句</div>' +
      '<div>原題：' + (S.finalFirst ? '第一次就答對' : '試了不只一次才答對') + '</div></div>' +
      (next ? '<button class="gp-next glow" data-act="another" data-id="' + h(next['題目ID'] || next['題號']) + '">換一題再試試 →</button>' : '');
  }

  function render() {
    if (!S) return;
    var el = overlay();
    var bar = STEPS.map(function (t, i) {
      return '<span class="gp-step' + (i === S.step ? ' is-now' : i < S.step ? ' is-done' : '') + '">' + (i < S.step ? '✓' : (i + 1)) + ' ' + t + '</span>';
    }).join('');
    var body = [viewGuess, viewPickChar, viewMatch, viewRule, viewRevise, viewFinal][S.step]();
    var done = stepDone();
    var nextBtn = S.step < 5 ? '<button class="gp-next' + (done ? ' glow' : '') + '" data-act="next"' + (done ? '' : ' disabled') + '>下一步 →</button>' : '';
    el.innerHTML =
      '<div class="gp-panel" data-kind="cz">' +
        '<div class="gp-top"><div class="gp-title">✏️ 一起動腦：錯別字</div>' +
          '<button class="gp-close" data-act="close" aria-label="離開">✕ 離開</button></div>' +
        '<div class="gp-steps">' + bar + '</div>' +
        '<div class="gp-src"><b>📘 教學依據</b>' + h(SOURCE) + '<span class="gp-src-step">這一步：' + h(STEP_SRC[S.step]) + '</span></div>' +
        '<div class="gp-body">' + body + '</div>' +
        '<div class="gp-foot">' +
          (S.step > 0 && S.step < 5 ? '<button class="gp-back" data-act="back">← 上一步</button>' : '<span></span>') +
          nextBtn + '</div>' +
      '</div>';
  }

  function onClick(ev) {
    if (!S) return;                               // 破音字開著時交給 GuidedPY
    var b = ev.target.closest('[data-act]');
    if (!b || b.disabled) return;
    var act = b.getAttribute('data-act'), v = b.getAttribute('data-v');
    var id = Number(b.getAttribute('data-id'));

    if (act === 'close') return close();
    if (act === 'next' && stepDone()) { S.step++; return render(); }
    if (act === 'back') { S.step--; return render(); }
    if (act === 'another') { var nid = b.getAttribute('data-id'); close(); return open(nid); }
    if (act === 'guess') { S.guess[id] = v; return render(); }

    if (act === 'char') {
      var x = S.m.items[id];
      S.pick[x.id] = v;
      if (v === x.correct) {
        S.done[x.id] = true;
        if (!S.tries[x.id]) S.firstTry[x.id] = true;
        S.hint[x.id] = (x.radical || x.meaning)
          ? '對！「' + h(x.correct) + '」' + (x.radical ? '是' + h(x.radical) : '') + (x.meaning ? '，意思是「' + h(x.meaning) + '」' : '') + '。'
          : '';
        tone(true);
      } else {
        S.tries[x.id] = (S.tries[x.id] || 0) + 1;
        tone(false);
        if (S.tries[x.id] === 1) {
          S.hint[x.id] = '再讀一次整句，這裡想說的詞是什麼？找出關鍵的意思再選。';
        } else if (S.tries[x.id] === 2) {
          S.hint[x.id] = x.meaning ? '提示：這裡要的是「' + h(x.meaning) + '」的意思，哪個字和這個意思有關？' : '提示：這個詞是「' + h(x.word) + '」。';
        } else {
          S.done[x.id] = true; S.pick[x.id] = x.correct;
          S.hint[x.id] = '這裡要寫「' + h(x.correct) + '」（' + h(x.word) + '），老師幫你選好了。等一下說說看為什麼。';
        }
      }
      return render();
    }

    if (act === 'mchar') { if (S.mSel !== id) S.mHint = ''; S.mSel = id; return render(); }
    if (act === 'mmean') {
      if (S.mSel == null) { S.mHint = '先點一個字，再點它的意思。'; return render(); }
      var it = S.m.items[S.mSel];
      if (it.meaning === v) {
        S.mDone[it.id] = true; S.mSel = null; tone(true);
        S.mHint = '「' + h(it.correct) + '」' + (it.radical ? '是' + h(it.radical) + '，' : '') + '和「' + h(it.meaning) + '」有關。';
      } else {
        S.mTries[it.id] = (S.mTries[it.id] || 0) + 1; tone(false);
        if (S.mTries[it.id] === 1) S.mHint = '想想有「' + h(it.correct) + '」的語詞「' + h(it.word) + '」，它是什麼意思？';
        else if (S.mTries[it.id] === 2) S.mHint = '提示：' + (it.radical ? '「' + h(it.correct) + '」是' + h(it.radical) + '，' : '') + '哪個意思和這個部首有關？';
        else { S.mDone[it.id] = true; S.mSel = null; S.mHint = '「' + h(it.correct) + '」的意思是「' + h(it.meaning) + '」，老師幫你配好了。'; }
      }
      return render();
    }

    if (act === 'final') {
      var ok = String(v) === S.m.answer;
      if (S.finalPick == null) S.finalFirst = ok;
      S.finalPick = v; S.finalOk = ok; tone(ok);
      return render();
    }
  }

  document.addEventListener('keydown', function (e) { if (S && e.key === 'Escape') close(); });

  root.GuidedCZ = { open: open, close: close, configure: configure, qualifies: qualifies, buildModel: buildModel };
})(typeof window !== 'undefined' ? window : globalThis);
