// zhu-passage-parse.js — 解析老師從外部 AI 貼回來的短文
//
// 窄容錯：只認得出明確標記的格式，找不到標記就要求重貼，
// 絕不猜「哪一段是短文、哪一段是答案」。猜錯會產出一份看起來正常、
// 答案卻錯位的卷子，比要求重貼嚴重得多。
var ZhuPassageParse = (function () {
  'use strict';

  // 標記一律要求在行首（前面只允許空白）。裸標記「短文：」「答案：」若用
  // indexOf 全文搜尋，正文裡的「老師說出答案：小明贏了」會被當成標記，
  // 短文被腰斬、答案抓到「小明贏了」——而且不會報任何錯。
  var PASSAGE_RE = /^[ \t　]*(?:【短文】|\[短文\]|短文[：:])/m;
  var ANSWER_RE = /^[ \t　]*(?:【答案】|\[答案\]|答案[：:])/m;

  function problem(code, message) {
    return {
      code: code, paragraph: null, sentence: null, blank: null,
      chars: [], severity: 'blocking', message: message
    };
  }

  function fail(code, message) {
    return { ok: false, 段落: [], 答案: [], problems: [problem(code, message)] };
  }

  function stripFence(text) {
    return String(text == null ? '' : text)
      .replace(/^\s*```[a-zA-Z]*\s*\n/, '')
      .replace(/\n\s*```\s*$/, '');
  }

  // 回傳最先出現的行首標記；找不到給 index -1。
  // index = 標記所在位置（含前置空白，用來比順序）
  // end   = 標記結束位置（用來切內容）
  function findMarker(text, re) {
    var m = re.exec(text);
    if (!m) { return { index: -1, end: -1 }; }
    return { index: m.index, end: m.index + m[0].length };
  }

  // 數全文有幾個行首標記。只檢查短文區內部不夠：AI 生成兩篇完整的
  // 「【短文】…【答案】…」時，第二組整個落在第一個答案標記之後，
  // 短文區裡乾乾淨淨，會安靜地只吃掉第一篇。
  function countMarker(text, re) {
    var g = new RegExp(re.source, 'gm');
    var n = 0;
    while (g.exec(text)) { n++; }
    return n;
  }

  // 空括號才正規化成 {}；括號內有字的不動——那是 AI 把答案寫進正文，
  // 要留給洩題門禁擋，解析器擅自剝除等於幫忙藏證據。
  function normalizeBlanks(s) {
    return String(s)
      .replace(/[｛]\s*[｝]/g, '{}')
      .replace(/\{\s*\}/g, '{}')
      .replace(/[（(][\s　　]*[）)]/g, '{}');
  }

  function splitAnswers(line) {
    return String(line || '')
      .replace(/\d+\s*[.、)）]\s*/g, ' ')
      .split(/[、,，\s　]+/)
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s.length > 0; });
  }

  function parse(text) {
    var raw = stripFence(text);
    if (!raw.trim()) {
      return fail('parse.missingPassageMarker',
        '內容是空的，請貼上外部 AI 的完整回覆');
    }

    var p = findMarker(raw, PASSAGE_RE);
    var a = findMarker(raw, ANSWER_RE);

    if (p.index < 0) {
      return fail('parse.missingPassageMarker',
        '找不到「【短文】」標記，請確認格式後重貼');
    }
    if (a.index < 0) {
      return fail('parse.missingAnswerMarker',
        '找不到「【答案】」標記，請確認格式後重貼');
    }
    if (a.index < p.index) {
      return fail('parse.markerOrder',
        '「【答案】」出現在「【短文】」之前，請確認格式後重貼');
    }

    // 全文只該有一組標記。多出來的一組代表 AI 生成了兩篇，
    // 而兩篇的空格數與答案數可能剛好對得上，後面四道門禁全都抓不到。
    if (countMarker(raw, PASSAGE_RE) > 1 || countMarker(raw, ANSWER_RE) > 1) {
      return fail('parse.duplicateMarker',
        '短文區裡又出現了一組標記，可能是 AI 生成了兩篇。請只貼一篇。');
    }

    var body = raw.substring(p.end, a.index);
    var 段落 = body.split(/\n\s*\n/)
      .map(function (s) { return normalizeBlanks(s.trim()); })
      .filter(function (s) { return s.length > 0; });

    if (!段落.length) {
      return fail('parse.emptyPassage', '「【短文】」底下沒有內容');
    }

    // 標記同行剩餘文字優先；同行沒東西才往下找第一個非空行
    var after = raw.substring(a.end);
    var lines = after.split('\n');
    var answerLine = '';
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].trim()) { answerLine = lines[i].trim(); break; }
    }
    var 答案 = splitAnswers(answerLine);
    if (!答案.length) {
      return fail('parse.emptyAnswer', '「【答案】」底下沒有內容');
    }

    return { ok: true, 段落: 段落, 答案: 答案, problems: [] };
  }

  var api = { parse: parse };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  return api;
})();
