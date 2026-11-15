/* 一起動腦：破音字（依學扶教材 3-1-3《語詞一家人（一）》的引導順序）
 * 先念（不評對錯）→ 分堆 → 說共同點 → 命名規則 → 修正想法 → 原題 → 換一個字
 * 設計文件：字音形互動引導設計_教材依據.md
 * 不用語音朗讀：電腦語音常把破音字唸錯，會教錯讀音。
 */
(function (root) {
  'use strict';

  /* ═══ 核心（可在 Node 測試） ═══════════════════════════ */
  var BOPO_RE = /[ㄅ-ㄩˊˇˋ˙]/g;
  var HAS_BOPO = /[ㄅ-ㄩˊˇˋ˙]/;

  function cleanMeaning(text) {
    var s = String(text || '');
    s = s.split('→')[0];                       // 去掉「→宰相、相片」例詞
    s = s.replace(/[（(][^）)]*[）)]/g, '');   // 去掉括號（內含讀音）
    s = s.replace(BOPO_RE, '');
    s = s.replace(/[一二三四輕]聲/g, '');
    s = s.replace(/[，,、\s]+$/, '').replace(/：\s*，/, '：').trim();
    return s;
  }

  function examplesOf(text) {
    var parts = String(text || '').split('→');
    return parts.length > 1 ? parts[1].trim() : '';
  }

  /* 題目能不能用：同一個字、至少兩種讀音、至少 3 個語詞 */
  /* 卡片上不可出現讀音或說明（如「得（必須義ㄉㄟˇ）」會洩漏答案） */
  function goodWord(w, ch) {
    return !!w && !HAS_BOPO.test(w) && !/[（(]/.test(w) && w.indexOf(ch) >= 0;
  }

  /* 從原題取出某個選項的文字：(1)…(2)… */
  function optionText(qText, n) {
    var re = new RegExp('[（(]' + n + '[)）]\s*([^(（]+)');
    var m = String(qText || '').match(re);
    return m ? m[1].trim() : '';
  }

  /* 從選項句子取出含「字」的那一小句，去掉「」 */
  function clauseWith(text, ch) {
    var parts = String(text || '').split(/[！!？?。，,；;]/);
    for (var i = 0; i < parts.length; i++) {
      if (parts[i].indexOf('「' + ch + '」') >= 0) return parts[i].replace(/[「」]/g, '').trim();
    }
    return '';
  }

  /* 卡片用的語詞：細目語詞能用就用；不能用（含讀音、括號、或不含目標字，
     如「得（必須義ㄉㄟˇ）」「二月份」）就改用原題該選項的句子；都不行才排除 */
  function usableRows(q) {
    var qText = (q && q['完整題目與選項']) || '';
    var out = [];
    ((q && q._py) || []).forEach(function (r) {
      var ch = r['破音字'], w = String(r['詞語'] || '');
      if (goodWord(w, ch)) { out.push(r); return; }
      var m = String(r['來源'] || '').match(/選項\((\d)\)/);
      var alt = m ? clauseWith(optionText(qText, m[1]), ch) : '';
      if (goodWord(alt, ch)) {
        var copy = {}; for (var k in r) copy[k] = r[k];
        copy['詞語'] = alt;
        out.push(copy);
      }
    });
    return out;
  }

  function joinMeanings(list) {
    var order = [], by = {}, plain = [];
    list.forEach(function (m) {
      var i = m.indexOf('：');
      if (i > 0) {
        var lab = m.slice(0, i), d = m.slice(i + 1);
        if (!by[lab]) { by[lab] = []; order.push(lab); }
        if (d && by[lab].indexOf(d) < 0) by[lab].push(d);
      } else if (plain.indexOf(m) < 0) plain.push(m);
    });
    return order.map(function (lab) { return by[lab].length ? lab + '：' + by[lab].join('、') : lab; })
      .concat(plain).join('／');
  }

  function qualifies(q) {
    var rows = usableRows(q);
    if (rows.length < 3) return false;
    var chars = {}, reads = {};
    rows.forEach(function (r) { chars[r['破音字']] = 1; reads[r['注音']] = 1; });
    var ch = Object.keys(chars)[0];
    // 原題要真的在考這個字（第 210 題細目誤放「得」的資料，原題考「喝」）
    var qText = String((q && q['完整題目與選項']) || '');
    if (qText && qText.indexOf('「' + ch + '」') < 0) return false;
    return Object.keys(chars).length === 1 && Object.keys(reads).length >= 2;
  }

  function buildModel(q) {
    if (!qualifies(q)) return null;
    var rows = usableRows(q);
    var readings = [];
    var byReading = {};
    var cards = [];
    rows.forEach(function (r, i) {
      var rd = r['注音'];
      if (!byReading[rd]) {
        byReading[rd] = { reading: rd, meanings: [], examples: [] };
        readings.push(rd);
      }
      var m = cleanMeaning(r['字義說明']);
      if (m && byReading[rd].meanings.indexOf(m) < 0) byReading[rd].meanings.push(m);
      var ex = examplesOf(r['字義說明']);
      if (ex && byReading[rd].examples.indexOf(ex) < 0) byReading[rd].examples.push(ex);
      cards.push({ id: i, word: r['詞語'], reading: rd, meaning: m });
    });
    return {
      uid: q['題目ID'] || q['題號'],
      char: rows[0]['破音字'],
      grade: q['年級'] || '',
      cards: cards,
      groups: readings.map(function (rd) {
        var g = byReading[rd];
        return { reading: rd, meaning: joinMeanings(g.meanings), examples: g.examples.join('、') };
      }),
      canMeaning: (function () {
        var ms = readings.map(function (rd) { return joinMeanings(byReading[rd].meanings); });
        return ms.every(function (m) { return m; }) && ms.filter(function (m, i) { return ms.indexOf(m) === i; }).length === ms.length;
      })(),
      tip: String(q['記憶訣竅'] || '').trim(),
      summary: String(q['讀音區分摘要'] || '').trim()
    };
  }

  /* 換一個字：優先同年級、不同字 */
  function pickNext(all, current) {
    var pool = (all || []).filter(function (q) {
      return qualifies(q) && usableRows(q)[0]['破音字'] !== current.char;
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

  var core = { cleanMeaning: cleanMeaning, qualifies: qualifies, buildModel: buildModel, pickNext: pickNext };
  if (typeof module !== 'undefined' && module.exports) { module.exports = core; return; }

  /* ═══ 畫面 ═══════════════════════════════════════════ */
  var STEPS = ['先念念看', '分成兩堆', '說說意思', '找到規則', '想法改了嗎', '原題挑戰'];
  /* 給老師看的教材依據（學扶國語三年級教師手冊 3-1-3《語詞一家人（一）》） */
  var SOURCE = '學扶國語教材 三年級《語詞一家人（一）》（3-1-3）';
  var STEP_SRC = [
    '準備活動「念一念」：先讓學生試讀，手冊提醒「無須立即糾正讀錯的學生」',
    '發展活動（一）「來玩字音的遊戲」：念出語詞，把語詞卡放到對的字音卡上',
    '發展活動（二）「發現字音的祕密」：觀察兩堆牌卡，說說念不同音時是什麼意思',
    '發展活動（二）羊博士的話：字音會改變，是因為意思不同（音隨義轉）',
    '準備活動手冊提示「先讓學生提出自己的想法」，加上「抽語詞說句子」',
    '發展活動（三）（五）換一個字再玩一次、自己找多音字（漸進式釋放責任）'
  ];
  var S = null;          // 目前狀態
  var audio = null;
  var cfg = { getAll: function () { return []; }, parseOptions: null, getAnswer: null };
  function configure(o) { for (var k in o) cfg[k] = o[k]; }
  function getAll() { return cfg.getAll() || []; }
  function ans(q) { return cfg.getAnswer ? cfg.getAnswer(q) : null; }

  function h(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function tone(ok) {
    try {
      audio = audio || new (window.AudioContext || window.webkitAudioContext)();
      var o = audio.createOscillator(), g = audio.createGain();
      o.type = 'sine';
      o.frequency.value = ok ? 880 : 330;
      g.gain.setValueAtTime(0.12, audio.currentTime);
      g.gain.exponentialRampToValueAtTime(0.001, audio.currentTime + (ok ? 0.25 : 0.35));
      o.connect(g); g.connect(audio.destination);
      o.start(); o.stop(audio.currentTime + 0.4);
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
    if (!el._pyBound) { el.addEventListener('click', onClick); el._pyBound = true; }  // 錯別字模組可能先建立視窗
    return el;
  }

  function findQ(uid) {
    var all = getAll();
    for (var i = 0; i < all.length; i++) {
      if (String(all[i]['題目ID'] || all[i]['題號']) === String(uid)) return all[i];
    }
    return null;
  }

  function open(uid) {
    var q = findQ(uid);
    var model = q && buildModel(q);
    if (!model) return;
    if (root.GuidedCZ && root.GuidedCZ.close) root.GuidedCZ.close();
    if (root.GuidedCY && root.GuidedCY.close) root.GuidedCY.close();
    if (root.GuidedXJ && root.GuidedXJ.close) root.GuidedXJ.close();
    S = {
      q: q, m: model, step: 0,
      guess: null,
      order: shuffle(model.cards),
      placed: {},          // cardId -> reading
      tries: {},           // cardId -> 錯幾次
      firstTry: {},        // cardId -> 一次就對
      selected: null,
      hint: '',
      meaningPick: {},     // reading -> 選到的意思
      meaningOk: {},
      meaningOpts: {},
      finalPick: null, finalOk: null, finalFirst: null
    };
    model.groups.forEach(function (g) {
      S.meaningOpts[g.reading] = shuffle(model.groups.map(function (x) { return x.meaning; }));
    });
    document.body.classList.add('gp-open');
    render();
  }

  function close() {
    var el = document.getElementById('gp-overlay');
    if (el && S) el.innerHTML = '';
    document.body.classList.remove('gp-open');
    S = null;
  }

  function allPlaced() { return S.m.cards.every(function (c) { return S.placed[c.id]; }); }
  function allMeaningsOk() { return S.m.groups.every(function (g) { return S.meaningOk[g.reading]; }); }

  function stepDone() {
    switch (S.step) {
      case 0: return S.guess != null;
      case 1: return allPlaced();
      case 2: return !S.m.canMeaning || allMeaningsOk();
      case 3: case 4: return true;
      case 5: return S.finalOk === true;
    }
    return false;
  }

  /* ── 各步驟內容 ── */
  function viewGuess() {
    var c = S.m.char;
    var words = S.order.map(function (k) { return '<span class="gp-word">' + h(k.word) + '</span>'; }).join('');
    var opts = [['same', '都一樣'], ['diff', '有的不一樣'], ['unsure', '我還不確定']];
    return '<p class="gp-ask">先自己小聲唸唸看。這些語詞裡的「<b>' + h(c) + '</b>」，讀音都一樣嗎？</p>' +
      '<div class="gp-words">' + words + '</div>' +
      '<div class="gp-choices">' + opts.map(function (o) {
        return '<button class="gp-choice' + (S.guess === o[0] ? ' is-on' : '') + '" data-act="guess" data-v="' + o[0] + '">' + o[1] + '</button>';
      }).join('') + '</div>' +
      (S.guess ? '<p class="gp-note">好，先記住你的想法，等一下我們來驗證。</p>' : '');
  }

  function viewSort() {
    var c = S.m.char;
    var pile = S.order.filter(function (k) { return !S.placed[k.id]; }).map(function (k) {
      return '<button class="gp-card' + (S.selected === k.id ? ' is-sel' : '') + '" data-act="pick" data-id="' + k.id + '">' + h(k.word) + '</button>';
    }).join('');
    var bins = S.m.groups.map(function (g) {
      var inside = S.m.cards.filter(function (k) { return S.placed[k.id] === g.reading; }).map(function (k) {
        return '<span class="gp-card is-in">' + h(k.word) + '</span>';
      }).join('');
      return '<div class="gp-bin' + (S.selected != null ? ' is-target' : '') + '" data-act="drop" data-r="' + h(g.reading) + '">' +
        '<div class="gp-bin-head">「' + h(c) + '」念 <span class="gp-bopo">' + h(g.reading) + '</span></div>' +
        '<div class="gp-bin-body">' + inside + '</div></div>';
    }).join('');
    return '<p class="gp-ask">點一張語詞卡，唸出來，再點它的讀音，把卡片放進去。</p>' +
      '<div class="gp-pile">' + (pile || '<span class="gp-note">全部放好了！</span>') + '</div>' +
      '<div class="gp-bins">' + bins + '</div>' +
      (S.hint ? '<div class="gp-hint">' + S.hint + '</div>' : '');
  }

  function viewMeaning() {
    var c = S.m.char;
    if (!S.m.canMeaning) {
      return '<p class="gp-ask">看看每一堆的語詞，和同學說說：「' + h(c) + '」念這個音的時候，是什麼意思？</p>' +
        S.m.groups.map(function (g) {
          return '<div class="gp-mrow"><div class="gp-mhead"><span class="gp-bopo">' + h(g.reading) + '</span>' +
            S.m.cards.filter(function (k) { return k.reading === g.reading; }).map(function (k) { return '<span class="gp-word sm">' + h(k.word) + '</span>'; }).join('') + '</div></div>';
        }).join('') + '<p class="gp-note">這一題的資料沒有分開寫兩種意思，請老師帶大家口頭討論。</p>';
    }
    return '<p class="gp-ask">看看每一堆的語詞。「' + h(c) + '」念這個音的時候，是什麼意思？</p>' +
      S.m.groups.map(function (g) {
        var words = S.m.cards.filter(function (k) { return k.reading === g.reading; })
          .map(function (k) { return '<span class="gp-word sm">' + h(k.word) + '</span>'; }).join('');
        var opts = S.meaningOpts[g.reading].map(function (m) {
          var on = S.meaningPick[g.reading] === m;
          var cls = on ? (S.meaningOk[g.reading] ? ' is-ok' : ' is-no') : '';
          return '<button class="gp-choice sm' + cls + '" data-act="mean" data-r="' + h(g.reading) + '" data-v="' + h(m) + '"' +
            (S.meaningOk[g.reading] ? ' disabled' : '') + '>' + h(m) + '</button>';
        }).join('');
        return '<div class="gp-mrow"><div class="gp-mhead"><span class="gp-bopo">' + h(g.reading) + '</span>' + words + '</div>' +
          '<div class="gp-choices">' + opts + '</div>' +
          (S.meaningPick[g.reading] && !S.meaningOk[g.reading] ? '<div class="gp-hint">再看一次這一堆的語詞，它們有什麼共同的地方？</div>' : '') +
          '</div>';
      }).join('');
  }

  function viewRule() {
    var c = S.m.char;
    var lines = S.m.groups.map(function (g) {
      return '<li>念 <span class="gp-bopo">' + h(g.reading) + '</span>：' + h(g.meaning) +
        (g.examples ? '<span class="gp-ex">（還有：' + h(g.examples) + '）</span>' : '') + '</li>';
    }).join('');
    return '<div class="gp-rule">字的意思不一樣，讀音就跟著不一樣。</div>' +
      '<ul class="gp-rule-list">' + lines + '</ul>' +
      (S.m.tip ? '<div class="gp-tip">🧠 記憶訣竅：' + h(S.m.tip) + '</div>' : '') +
      '<p class="gp-note">這就是「音隨義轉」：同一個「' + h(c) + '」，要看它在語詞裡是什麼意思，才知道怎麼念。</p>';
  }

  function viewRevise() {
    var c = S.m.char;
    var label = { same: '都一樣', diff: '有的不一樣', unsure: '還不確定' }[S.guess];
    var right = S.guess === 'diff';
    var frame = S.m.groups.map(function (g) {
      return '「' + h(c) + '」念 <span class="gp-bopo">' + h(g.reading) + '</span> 的時候，意思是<u>' + h(g.meaning) + '</u>';
    }).join('；');
    return '<p class="gp-ask">一開始你覺得讀音「' + h(label) + '」。</p>' +
      '<p class="gp-note">' + (right
        ? '你一開始就發現有的不一樣，很厲害！現在你還能說出為什麼。'
        : '現在我們知道，這些「' + h(c) + '」有的念法不一樣。想法改變了也很好，這就是在學習。') + '</p>' +
      '<div class="gp-frame"><div class="gp-frame-lbl">🗣️ 用這句話說一次給同學聽</div>' + frame + '。</div>';
  }

  function viewFinal() {
    var text = S.q['完整題目與選項'] || '';
    var opts = (cfg.parseOptions && cfg.parseOptions(text)) || null;
    var answer = ans(S.q);
    if (!opts || !answer) {
      return '<p class="gp-note">這一題沒有可以自動對答案的選項，請老師帶大家一起看原題：</p><div class="gp-q">' + h(text) + '</div>';
    }
    var stem = text.split(/[（(][1-4１-４][)）]/)[0];
    var btns = opts.map(function (o) {
      var on = String(S.finalPick) === String(o.key);
      var cls = on ? (String(o.key) === String(answer) ? ' is-ok' : ' is-no') : '';
      return '<button class="gp-choice' + cls + '" data-act="final" data-v="' + h(o.key) + '"' + (S.finalOk ? ' disabled' : '') + '>(' + h(o.key) + ') ' + h(o.text) + '</button>';
    }).join('');
    var msg = '';
    if (S.finalOk) msg = '<div class="gp-good">答對了！你用剛才找到的規則解出了考古題。</div>';
    else if (S.finalPick != null) msg = '<div class="gp-hint">再想想：把每個選項的「' + h(S.m.char) + '」放回剛才的兩堆，哪一個和其他三個不同堆？</div>';
    return '<p class="gp-ask">現在用剛才的規則，自己做做看原題：</p><div class="gp-q">' + h(stem) + '</div>' +
      '<div class="gp-choices col">' + btns + '</div>' + msg + (S.finalOk ? viewSummary() : '');
  }

  function viewSummary() {
    var n = S.m.cards.length;
    var first = S.m.cards.filter(function (c) { return S.firstTry[c.id]; }).length;
    var next = pickNext(getAll(), S.m);
    return '<div class="gp-sum"><div class="gp-frame-lbl">📋 這一題的學習紀錄（給老師看）</div>' +
      '<div>一開始的猜測：' + h({ same: '都一樣', diff: '有的不一樣', unsure: '還不確定' }[S.guess]) + '</div>' +
      '<div>分堆一次就放對：' + first + ' / ' + n + ' 張</div>' +
      '<div>原題：' + (S.finalFirst ? '第一次就答對' : '試了不只一次才答對') + '</div></div>' +
      (next ? '<button class="gp-next glow" data-act="another" data-id="' + h(next['題目ID'] || next['題號']) + '">換一個字再試試：「' + h(next._py[0]['破音字']) + '」 →</button>' : '');
  }

  function render() {
    if (!S) return;
    var el = overlay();
    var bar = STEPS.map(function (t, i) {
      return '<span class="gp-step' + (i === S.step ? ' is-now' : i < S.step ? ' is-done' : '') + '">' + (i < S.step ? '✓' : (i + 1)) + ' ' + t + '</span>';
    }).join('');
    var body = [viewGuess, viewSort, viewMeaning, viewRule, viewRevise, viewFinal][S.step]();
    var done = stepDone();
    var nextBtn = S.step < 5
      ? '<button class="gp-next' + (done ? ' glow' : '') + '" data-act="next"' + (done ? '' : ' disabled') + '>下一步 →</button>'
      : '';
    el.innerHTML =
      '<div class="gp-panel">' +
        '<div class="gp-top"><div class="gp-title">✏️ 一起動腦：破音字「<b>' + h(S.m.char) + '</b>」</div>' +
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
    var b = ev.target.closest('[data-act]');
    if (!b || !S || b.disabled) return;
    var act = b.getAttribute('data-act');
    var v = b.getAttribute('data-v');

    if (act === 'close') return close();
    if (act === 'next' && stepDone()) { S.step++; S.hint = ''; return render(); }
    if (act === 'back') { S.step--; S.hint = ''; return render(); }
    if (act === 'another') { var id = b.getAttribute('data-id'); close(); return open(id); }

    if (act === 'guess') { S.guess = v; return render(); }

    if (act === 'pick') {
      var cid = Number(b.getAttribute('data-id'));
      if (S.selected !== cid) S.hint = '';   // 換別張卡才清掉提示；同一張卡維持選取與提示
      S.selected = cid;
      return render();
    }

    if (act === 'drop') {
      if (S.selected == null) { S.hint = '先點一張語詞卡，再點讀音。'; return render(); }
      var card = S.m.cards[S.selected];
      var r = b.getAttribute('data-r');
      if (card.reading === r) {
        S.placed[card.id] = r;
        if (!S.tries[card.id]) S.firstTry[card.id] = true;
        S.selected = null; S.hint = '';
        tone(true);
      } else {
        S.tries[card.id] = (S.tries[card.id] || 0) + 1;
        tone(false);
        if (S.tries[card.id] === 1) {
          S.hint = '再唸一次「' + h(card.word) + '」，想想這個語詞是什麼意思。';
        } else if (S.tries[card.id] === 2) {
          S.hint = '提示：在「' + h(card.word) + '」裡，「' + h(S.m.char) + '」是「' + h(card.meaning) + '」的意思。哪一堆的語詞也是這個意思？';
        } else {
          S.placed[card.id] = card.reading;
          S.hint = '「' + h(card.word) + '」的「' + h(S.m.char) + '」念 <span class="gp-bopo">' + h(card.reading) + '</span>，老師幫你放好了。等一下說說看為什麼。';
          S.selected = null;
        }
      }
      return render();
    }

    if (act === 'mean') {
      var rd = b.getAttribute('data-r');
      var g = S.m.groups.filter(function (x) { return x.reading === rd; })[0];
      S.meaningPick[rd] = v;
      S.meaningOk[rd] = (g.meaning === v);
      tone(S.meaningOk[rd]);
      return render();
    }

    if (act === 'final') {
      var ok = String(v) === String(ans(S.q));
      if (S.finalPick == null) S.finalFirst = ok;
      S.finalPick = v; S.finalOk = ok;
      tone(ok);
      return render();
    }
  }

  document.addEventListener('keydown', function (e) { if (S && e.key === 'Escape') close(); });

  root.GuidedPY = { open: open, close: close, configure: configure, qualifies: qualifies, buildModel: buildModel };
})(typeof window !== 'undefined' ? window : globalThis);
