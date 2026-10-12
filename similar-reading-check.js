// similar-reading-check.js — 用教育部簡編本讀音表驗證「字音類」類似題
//
// class.html「出類似題」對多音字／字音題允許 AI 換掉正解，但換上來的詞讀音對不對
// 必須自動驗證：用 data/concised-bopo.json（scripts/build_concised_bopo.py 由簡編本
// 離線版建置）查每個「」字在該詞裡的讀音，確認讀音關係只有正解成立。
// 查不到、讀音有歧義、或關係不成立一律擋下（前端改用原題）。
var SimilarReadingCheck = (function () {
  'use strict';

  var MARK = /「([^」]+)」/g;

  function isReadingQuestion(stem, opts) {
    if (!/字音|讀音|注音/.test(String(stem || ''))) return false;
    if (!opts || opts.length < 2) return false;
    return opts.every(function (o) { return /「[^」]+」/.test(String(o)); });
  }

  // 去掉「」後的純文字，以及每個被標記字在純文字中的位置
  function marks(text) {
    var s = String(text || '');
    var plain = '';
    var out = [];
    var i = 0;
    while (i < s.length) {
      var c = s.charAt(i);
      if (c === '「') {
        var j = s.indexOf('」', i + 1);
        if (j < 0) { i++; continue; }
        var inner = s.slice(i + 1, j);
        for (var k = 0; k < inner.length; k++) {
          out.push({ ch: inner.charAt(k), idx: plain.length });
          plain += inner.charAt(k);
        }
        i = j + 1;
      } else {
        plain += c;
        i++;
      }
    }
    return { plain: plain, marks: out };
  }

  // 取涵蓋 idx 的最長詞條讀音；同長度讀音不一致回 null；都查不到時單音字直接用字的讀音
  function readingAt(plain, idx, dict) {
    for (var len = Math.min(8, plain.length); len >= 2; len--) {
      var found = {};
      for (var start = idx - len + 1; start <= idx; start++) {
        if (start < 0 || start + len > plain.length) continue;
        var rs = dict.words[plain.substr(start, len)];
        if (!rs) continue;
        rs.forEach(function (r) { found[r.split(' ')[idx - start]] = true; });
      }
      var keys = Object.keys(found);
      if (keys.length === 1) return keys[0];
      if (keys.length > 1) return null;
    }
    var cr = dict.chars[plain.charAt(idx)];
    return cr && cr.length === 1 ? cr[0] : null;
  }

  function readingsOf(text, dict) {
    var m = marks(text);
    var rs = m.marks.map(function (x) { return readingAt(m.plain, x.idx, dict); });
    return { plain: m.plain, readings: rs, unknown: rs.indexOf(null) >= 0 };
  }

  function allSame(arr) {
    for (var i = 1; i < arr.length; i++) if (arr[i] !== arr[0]) return false;
    return true;
  }

  // 回傳 '' 通過；字串＝擋下原因；null＝題型看不出來（呼叫端自行決定）
  function check(stem, opts, answerIdx, dict) {
    var info = opts.map(function (o) { return readingsOf(o, dict); });
    var counts = info.map(function (x) { return x.readings.length; });
    var pairs = counts.every(function (n) { return n >= 2; });
    var singles = counts.every(function (n) { return n === 1; });
    var stemText = String(stem || '');
    // 題幹裡的考題句：最後一個「」所在的句子（「和哪個選項的「得」字音相同？我「得」在…」）
    var anchorSentence = stemText.replace(/^[^？?]*[？?]/, '');
    var anchorMarks = marks(anchorSentence).marks;
    var mode = null;
    if (pairs && /相同/.test(stemText)) mode = 'pairSame';
    else if (pairs && /不同/.test(stemText)) mode = 'pairDiff';
    else if (singles && /不同/.test(stemText)) mode = 'odd';
    else if (singles && /相同/.test(stemText) && anchorMarks.length) mode = 'anchor';
    if (!mode) return null;

    for (var i = 0; i < info.length; i++) {
      if (info[i].unknown) return '讀音查不到：' + info[i].plain;
    }
    var flags;
    if (mode === 'pairSame') flags = info.map(function (x) { return allSame(x.readings); });
    if (mode === 'pairDiff') flags = info.map(function (x) { return !allSame(x.readings); });
    if (mode === 'odd') {
      var rs = info.map(function (x) { return x.readings[0]; });
      flags = rs.map(function (r, k) {
        var others = rs.filter(function (_, j) { return j !== k; });
        return allSame(others) && others[0] !== r;
      });
    }
    if (mode === 'anchor') {
      var a = readingsOf(anchorSentence, dict);
      var anchor = a.readings[a.readings.length - 1];
      if (!anchor) return '讀音查不到：' + a.plain;
      flags = info.map(function (x) { return x.readings[0] === anchor; });
    }
    var hits = flags.filter(Boolean).length;
    if (!flags[answerIdx]) return '依簡編本讀音，正解不成立';
    if (hits !== 1) return '依簡編本讀音，有 ' + hits + ' 個選項都成立';
    return '';
  }

  // ── 其他可驗證題型：錯別字、□填字、注音排列、部首 ──────────────────

  function wordSet(dict) {
    if (!dict._set) {
      var s = {};
      String(dict.all || '').split('\n').forEach(function (w) { if (w) s[w] = true; });
      dict._set = s;
    }
    return dict._set;
  }

  // plain 第 idx 字和前後文能不能組成簡編本裡的詞（2～6 字）
  function formsWord(plain, idx, dict) {
    var set = wordSet(dict);
    for (var len = 2; len <= 6; len++) {
      for (var start = idx - len + 1; start <= idx; start++) {
        if (start < 0 || start + len > plain.length) continue;
        if (set[plain.substr(start, len)]) return true;
      }
    }
    return false;
  }

  function afterQuestion(stem) {
    return String(stem || '').replace(/^[^？?]*[？?]/, '');
  }

  function splitChars(opt) {
    return String(opt || '').split(/[／╱\/]/).map(function (s) { return s.trim(); });
  }

  var CJK1 = /^[㐀-鿿]$/;
  var ZHUYIN = /^[ㄅ-ㄩ˙ˊˇˋ\s／╱\/]+$/;

  function verifyMode(stem, opts) {
    var s = String(stem || '');
    opts = (opts || []).map(String);
    if (opts.length < 2) return '';
    if (/部首/.test(s) && opts.every(function (o) { return CJK1.test(o.trim()); })) return 'radical';
    if (/字音|讀音|注音/.test(s) && opts.every(function (o) { return ZHUYIN.test(o); }) && marks(afterQuestion(s)).marks.length) return 'zhuyin';
    var sentence = afterQuestion(s);
    if (sentence.indexOf('□') >= 0 && opts.every(function (o) {
      return splitChars(o).every(function (c) { return CJK1.test(c); });
    })) return 'fill';
    if (isReadingQuestion(s, opts)) return 'reading';
    if (/國字|錯字|錯別字/.test(s) && /正確|錯誤/.test(s) && opts.every(function (o) { return /「[^」]+」/.test(o); })) return 'spelling';
    return '';
  }

  function pick(flags, answerIdx, what) {
    var hits = flags.filter(Boolean).length;
    if (!flags[answerIdx]) return '依簡編本' + what + '，正解不成立';
    if (hits !== 1) return '依簡編本' + what + '，有 ' + hits + ' 個選項都成立';
    return '';
  }

  function checkSpelling(stem, opts, answerIdx, dict) {
    var wantWrong = /錯誤|寫錯/.test(stem);
    var flags = opts.map(function (o) {
      var m = marks(o);
      var ok = m.marks.every(function (x) { return formsWord(m.plain, x.idx, dict); });
      return wantWrong ? !ok : ok;
    });
    return pick(flags, answerIdx, '詞語');
  }

  function checkFill(stem, opts, answerIdx, dict) {
    var sentence = afterQuestion(stem);
    var blanks = [];
    for (var i = 0; i < sentence.length; i++) if (sentence.charAt(i) === '□') blanks.push(i);
    var flags = [];
    for (var k = 0; k < opts.length; k++) {
      var cs = splitChars(opts[k]);
      if (cs.length !== blanks.length) return null;
      var filled = sentence.split('');
      blanks.forEach(function (b, j) { filled[b] = cs[j]; });
      var plain = filled.join('');
      flags.push(blanks.every(function (b) { return formsWord(plain, b, dict); }));
    }
    return pick(flags, answerIdx, '詞語');
  }

  function normZhuyin(s) {
    return String(s || '').replace(/[╱\/]/g, '／').replace(/\s+/g, '');
  }

  function checkZhuyin(stem, opts, answerIdx, dict) {
    var r = readingsOf(afterQuestion(stem), dict);
    if (r.unknown) return '讀音查不到：' + r.plain;
    var want = normZhuyin(r.readings.join('／'));
    return pick(opts.map(function (o) { return normZhuyin(o) === want; }), answerIdx, '讀音');
  }

  function checkRadical(stem, opts, answerIdx, dict) {
    var rads = opts.map(function (o) { return (dict.radicals || {})[o.trim()] || null; });
    var miss = opts.filter(function (o, i) { return !rads[i]; });
    if (miss.length) return '部首查不到：' + miss.join('、');
    var anchor = (String(stem).match(/「([㐀-鿿])」/) || [])[1];
    if (anchor) {
      var ar = (dict.radicals || {})[anchor];
      if (!ar) return '部首查不到：' + anchor;
      // 「珠」的部首是哪個？選項本身就是部首字
      if (/是它的部首|哪個是.*部首|部首是/.test(stem)) {
        return pick(opts.map(function (o) { return o.trim() === ar; }), answerIdx, '部首');
      }
      var same = /相同/.test(stem);
      return pick(rads.map(function (r) { return same ? r === ar : r !== ar; }), answerIdx, '部首');
    }
    if (/不同/.test(stem)) {
      return pick(rads.map(function (r, k) {
        var others = rads.filter(function (_, j) { return j !== k; });
        return allSame(others) && others[0] !== r;
      }), answerIdx, '部首');
    }
    return null;
  }

  // 依題型驗證：'' 通過；字串＝擋下原因；null＝題型無法自動驗證
  function verify(stem, opts, answerIdx, dict) {
    var mode = verifyMode(stem, opts);
    if (mode === 'radical') return checkRadical(stem, opts, answerIdx, dict);
    if (mode === 'zhuyin') return checkZhuyin(stem, opts, answerIdx, dict);
    if (mode === 'fill') return checkFill(stem, opts, answerIdx, dict);
    if (mode === 'reading') return check(stem, opts, answerIdx, dict);
    if (mode === 'spelling') return checkSpelling(stem, opts, answerIdx, dict);
    return null;
  }

  var api = { isReadingQuestion: isReadingQuestion, readingAt: readingAt, check: check, marks: marks,
    verifyMode: verifyMode, verify: verify, formsWord: formsWord };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
