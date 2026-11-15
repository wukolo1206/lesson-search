/* 一起動腦第 7 步：換個句子用用看（四種一起動腦共用）
 * 原題答對後，把剛學到的規則用在沒看過的句子：破音字選讀音、錯別字／形近字選字、詞義成語選詞或選意思。
 * 教材依據：各教材統整活動「回到新的句子再用一次」（3-1-3、3-1-5 統整；4-1-3「選一選，是哪一個意思」「我會完成句子」）
 * 資料：guided-apply-data.js（scripts/build_guided_apply.py 產生；句子由 Claude 撰寫，讀音、詞條、解釋以簡編本查證）
 * 樣式沿用 guided-py.css 的 gp-*。
 */
(function (root) {
  'use strict';

  var DATA = null;
  function data() { return DATA || root.GUIDED_APPLY_DATA || {}; }
  function setData(d) { DATA = d; }
  function items(uid) { return data()[String(uid)] || []; }
  function has(uid) { return items(uid).length > 0; }

  /* 句子顯示用的片段：read 把「詞」框出並標出字；char 的 □；word 的 ＿＿；mean 的「詞」 */
  function parts(it) {
    var s = it.s;
    if (it.t === 'char') { var i = s.indexOf('□'); return { before: s.slice(0, i), blank: '□', after: s.slice(i + 1) }; }
    if (it.t === 'word') { var j = s.indexOf('＿＿'); return { before: s.slice(0, j), blank: '＿＿', after: s.slice(j + 2) }; }
    var m = s.match(/「(.+?)」/);
    return { before: s.slice(0, m.index), word: m[1], after: s.slice(m.index + m[0].length) };
  }

  var core = { has: has, items: items, parts: parts, setData: setData };
  if (typeof module !== 'undefined' && module.exports) { module.exports = core; return; }

  var S = null, audio = null;
  var SOURCE = '各教材統整活動：把學到的方法用在新的句子（4-1-3「我會完成句子」、3-1-5 統整活動）';
  var ASK = {
    read: function (it) { return '「' + h(it.w) + '」的「' + h(it.ch) + '」在這裡讀什麼？'; },
    char: function () { return '□ 裡要填哪一個字？'; },
    word: function () { return '＿＿ 裡放哪一個最合適？'; },
    mean: function (it) { return '「' + h(parts(it).word) + '」在這句話裡是什麼意思？'; }
  };
  var HINT = {
    read: '想想這個語詞的意思，它和哪一個讀音的意思一樣？',
    char: '先想這個詞是什麼意思，再看哪一個字的部首和意思有關。',
    word: '找找句子裡的線索，把每個選項放進去讀讀看。',
    mean: '找找句子裡的線索，把每個意思放回句子讀讀看。'
  };

  function h(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function shuffleIdx(n) { var a = []; for (var i = 0; i < n; i++) a.push(i); for (var k = n - 1; k > 0; k--) { var j = Math.floor(Math.random() * (k + 1)); var t = a[k]; a[k] = a[j]; a[j] = t; } return a; }

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
    if (!el._apBound) { el.addEventListener('click', onClick); el._apBound = true; }
    return el;
  }

  /* opts：{ mod: 'GuidedPY'…, next: 下一題的題目ID } */
  function open(uid, opts) {
    var list = items(uid);
    if (!list.length) return;
    ['GuidedPY', 'GuidedCZ', 'GuidedCY', 'GuidedXJ'].forEach(function (k) { if (root[k] && root[k].close) root[k].close(); });
    S = { uid: String(uid), list: list, opts: opts || {}, pick: {}, tries: {}, done: {}, first: {}, hint: {},
          order: list.map(function (it) { return shuffleIdx(it.c.length); }) };
    document.body.classList.add('gp-open');
    render();
  }

  function close() {
    var el = document.getElementById('gp-overlay');
    if (el && S) el.innerHTML = '';
    document.body.classList.remove('gp-open');
    S = null;
  }

  function sentence(it, i) {
    var p = parts(it);
    if (it.t === 'read') {
      var w = h(p.word).replace(h(it.ch), '<span class="gp-mark">' + h(it.ch) + '</span>');
      return h(p.before) + '<u class="gp-ap-w">' + w + '</u>' + h(p.after);
    }
    if (it.t === 'mean') return h(p.before) + '<span class="gp-mark">' + h(p.word) + '</span>' + h(p.after);
    var fill = S.done[i] ? '<span class="gp-mark gp-swap">' + h(it.c[it.a]) + '</span>' : '<span class="gp-mark">' + h(p.blank) + '</span>';
    return h(p.before) + fill + h(p.after);
  }

  function allDone() { for (var i = 0; i < S.list.length; i++) if (!S.done[i]) return false; return true; }

  function render() {
    if (!S) return;
    var el = overlay();
    var body = '<p class="gp-ask">剛才學到的方法，換一個沒看過的句子也用得上嗎？</p>' + S.list.map(function (it, i) {
      var done = S.done[i];
      var big = it.t === 'char' ? ' gp-big' : '';
      var btns = S.order[i].map(function (k) {
        var on = S.pick[i] === k, cls = on ? (k === it.a ? ' is-ok' : ' is-no') : '';
        var lab = it.t === 'read' ? '<b class="gp-bopo">' + h(it.c[k]) + '</b>' : h(it.c[k]);
        return '<button class="gp-choice' + big + (it.t === 'mean' || it.t === 'word' ? ' sm' : '') + cls + '" data-act="ap" data-i="' + i + '" data-k="' + k + '"' + (done ? ' disabled' : '') + '>' + lab + '</button>';
      }).join('');
      var info = done && it.y ? '<div class="gp-good">' + (it.w ? '簡編本「' + h(it.yw || it.w) + '」：' : '簡編本：') + h(it.y) + '</div>' : '';
      return '<div class="gp-srow"><div class="gp-stext gp-sent">' + (i + 1) + '. ' + sentence(it, i) + '</div>' +
        '<p class="gp-ask" style="margin:6px 0 8px;font-size:.9em">' + ASK[it.t](it) + '</p>' +
        '<div class="gp-choices' + (it.t === 'mean' ? ' col' : '') + '">' + btns + '</div>' +
        (S.hint[i] ? '<div class="gp-hint">' + S.hint[i] + '</div>' : '') + info + '</div>';
    }).join('');
    if (allDone()) {
      var n = S.list.length, f = 0;
      for (var i = 0; i < n; i++) if (S.first[i]) f++;
      var nx = S.opts.next, mod = S.opts.mod;
      body += '<div class="gp-sum"><div class="gp-frame-lbl">📋 換個句子的學習紀錄（給老師看）</div>' +
        '<div>新句子一次就答對：' + f + ' / ' + n + ' 句</div></div>' +
        '<p class="gp-note">' + (f === n ? '全部一次就答對，你真的學會了！' : '錯了也沒關係，記住剛才的提示，下次就會了。') + '</p>' +
        (nx && mod && root[mod] ? '<button class="gp-next glow" data-act="apnext">換一題再試試 →</button>' : '');
    }
    el.innerHTML =
      '<div class="gp-panel" data-kind="apply">' +
        '<div class="gp-top"><div class="gp-title">✏️ 換個句子用用看</div>' +
          '<button class="gp-close" data-act="close" aria-label="離開">✕ 離開</button></div>' +
        '<div class="gp-src"><b>📘 教學依據</b>' + h(SOURCE) + '<span class="gp-src-step">讀音、詞語與解釋：教育部《國語辭典簡編本》</span></div>' +
        '<div class="gp-body">' + body + '</div>' +
        '<div class="gp-foot"><span></span><button class="gp-back" data-act="close">回到講義</button></div>' +
      '</div>';
  }

  function onClick(ev) {
    if (!S) return;
    var b = ev.target.closest('[data-act]');
    if (!b || b.disabled) return;
    var act = b.getAttribute('data-act');
    if (act === 'close') return close();
    if (act === 'apnext') { var mod = S.opts.mod, nx = S.opts.next; close(); return root[mod].open(nx); }
    if (act === 'ap') {
      var i = Number(b.getAttribute('data-i')), k = Number(b.getAttribute('data-k')), it = S.list[i];
      if (S.done[i]) return;
      S.pick[i] = k;
      if (k === it.a) { S.done[i] = true; if (!S.tries[i]) S.first[i] = true; S.hint[i] = ''; tone(true); }
      else {
        S.tries[i] = (S.tries[i] || 0) + 1; tone(false);
        if (S.tries[i] === 1) S.hint[i] = HINT[it.t];
        else { S.done[i] = true; S.pick[i] = it.a; S.hint[i] = '答案是「' + h(it.c[it.a]) + '」，老師幫你選好了。看看下面的解釋。'; }
      }
      return render();
    }
  }

  document.addEventListener('keydown', function (e) { if (S && e.key === 'Escape') close(); });

  root.GuidedApply = { open: open, close: close, has: has };
})(typeof window !== 'undefined' ? window : globalThis);
