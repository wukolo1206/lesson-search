// similar-dict-gen.js — 多音字類似題：不靠 AI，直接從教育部簡編本讀音表產題
//
// 2026-10-12 實測：校內 gemma4:31b 與 Gemini 換新多音字正解時讀音常錯（「轉動」標 ㄓㄨㄢ、
// 「便利／便宜」同音），驗證全擋。改由程式從 data/concised-bopo.json 挑詞組題，讀音
// 全部取自辭典，保證正確。支援三種題型：
//   odd    「哪個字音和其他三者不同」：3 個讀音 A 的詞＋1 個讀音 B 的詞
//   anchor 「和哪個選項的「X」字音相同」＋題幹句子：題幹沿用，1 個同音詞＋3 個不同音詞
//   pair   「哪個字音前後相同」且每組同一字：正解那組兩詞同音，其他組兩詞不同音
var SimilarDictGen = (function () {
  'use strict';

  var RC = (typeof SimilarReadingCheck !== 'undefined') ? SimilarReadingCheck
    : (typeof require === 'function' ? require('./similar-reading-check.js') : null);

  function shuffle(arr, rand) {
    var a = arr.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(rand() * (i + 1));
      var t = a[i]; a[i] = a[j]; a[j] = t;
    }
    return a;
  }

  // 不適合國小學生的詞或字（辭典產題時排除）
  var BLOCK_WORDS = { '掃黃': 1, '倒會': 1, '以身相許': 1 };
  var BLOCK_CHARS = /[賭嫖娼淫屍姦妓毒殺]/;

  function commonSet(dict) {
    if (dict._common === undefined) {
      var s = null;
      if (dict.common) {
        s = {};
        String(dict.common).split('\n').forEach(function (w) { if (w) s[w] = true; });
      }
      dict._common = s;
    }
    return dict._common;
  }

  // 含 ch 恰一次、2～4 字、讀音唯一、課本常見的詞，依 ch 的讀音分組
  function groupsFor(ch, dict, learnedHas, exclude) {
    var groups = {};
    var common = commonSet(dict);
    Object.keys(dict.words || {}).forEach(function (w) {
      if (common && !common[w]) return;
      if (BLOCK_WORDS[w] || BLOCK_CHARS.test(w)) return;
      if (w.length < 2 || w.length > 4) return;
      var at = w.indexOf(ch);
      if (at < 0 || w.indexOf(ch, at + 1) >= 0) return;
      if (exclude[w]) return;
      if (!/^[㐀-鿿]+$/.test(w)) return;
      if (learnedHas && !w.split('').every(learnedHas)) return;
      var rs = {};
      dict.words[w].forEach(function (r) { rs[r.split(' ')[at]] = true; });
      var keys = Object.keys(rs);
      if (keys.length !== 1) return;
      (groups[keys[0]] = groups[keys[0]] || []).push(w);
    });
    return groups;
  }

  function mark(w, ch) {
    var i = w.indexOf(ch);
    return w.slice(0, i) + '「' + ch + '」' + w.slice(i + 1);
  }

  function excludeOf(stem, opts) {
    var ex = {};
    [stem].concat(opts).forEach(function (t) {
      var plain = String(t).replace(/[「」]/g, '');
      for (var len = 2; len <= 4; len++) {
        for (var i = 0; i + len <= plain.length; i++) ex[plain.substr(i, len)] = true;
      }
    });
    return ex;
  }

  function targetChar(opts) {
    var chars = {};
    opts.forEach(function (o) { (String(o).match(/「([^」])」/g) || []).forEach(function (m) { chars[m.charAt(1)] = true; }); });
    var keys = Object.keys(chars);
    return keys.length === 1 ? keys[0] : null;
  }

  function genOdd(stem, opts, dict, learnedHas, rand) {
    var ch = targetChar(opts);
    if (!ch) return null;
    var g = groupsFor(ch, dict, learnedHas, excludeOf(stem, opts));
    var many = shuffle(Object.keys(g).filter(function (r) { return g[r].length >= 3; }), rand);
    for (var i = 0; i < many.length; i++) {
      var a = many[i];
      var others = shuffle(Object.keys(g).filter(function (r) { return r !== a; }), rand);
      if (!others.length) continue;
      var b = others[0];
      var odd = shuffle(g[b], rand)[0];
      var same = shuffle(g[a], rand).slice(0, 3);
      return {
        stem: stem,
        options: [mark(odd, ch)].concat(same.map(function (w) { return mark(w, ch); })),
        answer: mark(odd, ch),
        explanation: '「' + odd + '」的「' + ch + '」讀' + b + '，其餘三個讀' + a + '。'
      };
    }
    return null;
  }

  function genAnchor(stem, opts, dict, learnedHas, rand) {
    var ch = targetChar(opts);
    if (!ch) return null;
    var sentence = String(stem).replace(/^[^？?]*[？?]/, '');
    var m = RC.marks(sentence);
    if (!m.marks.length) return null;
    var want = RC.readingAt(m.plain, m.marks[m.marks.length - 1].idx, dict);
    if (!want) return null;
    var g = groupsFor(ch, dict, learnedHas, excludeOf(stem, opts));
    if (!g[want] || !g[want].length) return null;
    var pool = [];
    Object.keys(g).forEach(function (r) { if (r !== want) g[r].forEach(function (w) { pool.push(w); }); });
    if (pool.length < 3) return null;
    var hit = shuffle(g[want], rand)[0];
    var wrong = shuffle(pool, rand).slice(0, 3);
    return {
      stem: stem,
      options: [mark(hit, ch)].concat(wrong.map(function (w) { return mark(w, ch); })),
      answer: mark(hit, ch),
      explanation: '題目中的「' + ch + '」讀' + want + '，「' + hit + '」的「' + ch + '」也讀' + want + '。'
    };
  }

  // 配對題：每組只有一個字（同字兩詞），正解組兩詞同音，其他組兩詞不同音
  function genPair(stem, opts, dict, learnedHas, rand) {
    var chars = [];
    for (var i = 0; i < opts.length; i++) {
      var ms = String(opts[i]).match(/「([^」])」/g) || [];
      if (ms.length !== 2 || ms[0] !== ms[1]) return null;
      chars.push(ms[0].charAt(1));
    }
    var exclude = excludeOf(stem, opts);
    var order = shuffle(chars.map(function (_, k) { return k; }), rand);
    for (var t = 0; t < order.length; t++) {
      var ansAt = order[t];
      var options = [];
      var answer = null;
      var expl = '';
      var ok = true;
      for (var k = 0; k < chars.length && ok; k++) {
        var ch = chars[k];
        var g = groupsFor(ch, dict, learnedHas, exclude);
        var rs = shuffle(Object.keys(g), rand);
        if (k === ansAt) {
          var two = rs.filter(function (r) { return g[r].length >= 2; })[0];
          if (!two) { ok = false; break; }
          var ws = shuffle(g[two], rand).slice(0, 2);
          answer = mark(ws[0], ch) + '／' + mark(ws[1], ch);
          options.push(answer);
          expl = '「' + ws[0] + '」和「' + ws[1] + '」的「' + ch + '」都讀' + two + '。';
        } else {
          if (rs.length < 2) { ok = false; break; }
          var w1 = shuffle(g[rs[0]], rand)[0];
          var w2 = shuffle(g[rs[1]], rand)[0];
          options.push(mark(w1, ch) + '／' + mark(w2, ch));
        }
      }
      if (ok && answer) return { stem: stem, options: options, answer: answer, explanation: expl };
    }
    return null;
  }

  // 回傳 { stem, options, answer, explanation } 或 null（題型不支援或字典湊不出）
  function generate(stem, opts, answerIdx, dict, learnedHas, rand) {
    rand = rand || Math.random;
    if (!RC || !RC.isReadingQuestion(stem, opts)) return null;
    var s = String(stem);
    var marks = opts.map(function (o) { return (String(o).match(/「[^」]+」/g) || []).length; });
    var out = null;
    if (marks.every(function (n) { return n === 2; }) && /相同/.test(s)) out = genPair(s, opts, dict, learnedHas, rand);
    else if (marks.every(function (n) { return n === 1; }) && /不同/.test(s)) out = genOdd(s, opts, dict, learnedHas, rand);
    else if (marks.every(function (n) { return n === 1; }) && /相同/.test(s)) out = genAnchor(s, opts, dict, learnedHas, rand);
    if (!out) return null;
    // 產完再用同一套驗證確認一次
    return RC.verify(out.stem, out.options, out.options.indexOf(out.answer), dict) === '' ? out : null;
  }

  var api = { generate: generate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
