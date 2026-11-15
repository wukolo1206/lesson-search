/* 複習短文：把 AI 寫的短文轉成學生版（words.html 用；Node 可測試）
 * 依語詞來自哪一種題目決定怎麼練（第一性原理：練的是在新句子裡判斷音、字、義）：
 *   read  多音字 → 保留語詞、標出考點字，後面留（　）寫注音（整詞挖空練不到讀音）
 *   char  同音字／形近字／錯別字 → 考點字挖成 □
 *   word  詞義、成語 → 整個詞挖空，附詞語盒
 */
(function (root) {
  'use strict';
  var FORM_CATS = ['同音字辨析', '形近字辨析', '常見錯別字', '部首辨識'];

  function kindOf(d) {
    if (String(d.w || '').length < 2) return 'word';            // 單字詞挖字＝挖整個詞，放進詞語盒
    var hasChar = d.c && d.w.indexOf(d.c) >= 0;
    if (d.p && hasChar) return 'read';
    if (hasChar && (d.cats || []).some(function (c) { return FORM_CATS.indexOf(c) >= 0; })) return 'char';
    return 'word';
  }

  /* 依「先比長詞」把文字切成一般文字與語詞片段 */
  function segment(text, words) {
    text = String(text || '');
    var list = (words || []).slice().sort(function (a, b) { return b.w.length - a.w.length; });
    var out = [], buf = '', i = 0;
    while (i < text.length) {
      var hit = null;
      for (var k = 0; k < list.length; k++) if (text.substr(i, list[k].w.length) === list[k].w) { hit = list[k]; break; }
      if (hit) {
        if (buf) { out.push({ t: 'text', s: buf }); buf = ''; }
        out.push({ t: 'term', d: hit }); i += hit.w.length;
      } else { buf += text[i]; i++; }
    }
    if (buf) out.push({ t: 'text', s: buf });
    return out;
  }

  /* 學生版：每次出現都編號挖空；回傳片段、答案、詞語盒 */
  function studentSheet(text, words, shuffle) {
    var parts = [], answers = [], bank = [], n = 0;
    segment(text, words).forEach(function (p) {
      if (p.t === 'text') { parts.push(p); return; }
      var d = p.d, kind = kindOf(d); n++;
      parts.push({ t: 'blank', n: n, kind: kind, d: d });
      answers.push({ n: n, kind: kind, word: d.w, ans: kind === 'read' ? d.p : kind === 'char' ? d.c : d.w });
      if (kind === 'word' && bank.indexOf(d.w) < 0) bank.push(d.w);
    });
    if (shuffle) bank = shuffle(bank);
    return { parts: parts, answers: answers, bank: bank };
  }

  /* 年段：依語詞年級，3～4 年級＝中年級、5～6＝高年級；寫得淺白再降一級 */
  function levelFor(words, easy) {
    var g = 3;
    (words || []).forEach(function (d) { g = Math.max(g, (d.g || [3])[0]); });
    var lv = g >= 5 ? 'high' : 'middle';
    if (easy) lv = lv === 'high' ? 'middle' : 'low';
    return lv;
  }

  /* 講義「已教過」的題目ID → 這些題目的語詞 */
  function taughtFilter(words, taughtIds) {
    var set = {}; (taughtIds || []).forEach(function (id) { set[String(id)] = 1; });
    return (words || []).filter(function (d) { return (d.id || []).some(function (id) { return set[String(id)]; }); });
  }

  /* 詞語盒太少（少於 3 個）就從同類語詞補干擾詞，避免一看就知道答案 */
  function padBank(bank, pool, used, rand) {
    var need = bank.length && bank.length < 3 ? 2 : 0, out = bank.slice(), cand = [];
    (pool || []).forEach(function (d) {
      if (used.indexOf(d.w) < 0 && out.indexOf(d.w) < 0 && d.w.length >= 2 && (d.cats || []).some(function (c) { return c === '詞義理解' || c === '成語'; })) cand.push(d.w);
    });
    while (need > 0 && cand.length) { out.push(cand.splice(Math.floor((rand || Math.random)() * cand.length), 1)[0]); need--; }
    return { bank: out, padded: out.length > bank.length };
  }

  var api = { kindOf: kindOf, padBank: padBank, segment: segment, studentSheet: studentSheet, levelFor: levelFor, taughtFilter: taughtFilter, MAX_TERMS: 10 };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.WordsReview = api;
})(typeof window !== 'undefined' ? window : globalThis);
