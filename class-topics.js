/* 課堂講義：主題闖關（2026-11-15）
 * 比照學扶國語句段系統：7 座島（題型）→ 三～六年級單元 → 一題一關 → 完成關。
 * 進入單元＝套用「類別＋年級」篩選，所以投影、題目模式、列印講義自動只用這個單元的題目。
 * 「全部題目」保留原本的篩選列表（view=all）。
 */
(function (root) {
  'use strict';

  var ISLANDS = [
    { cat: '多音字', icon: '🔊', color: '#a78bfa', tag: '音隨義轉：意思不同，讀音就不同' },
    { cat: '同音字辨析', icon: '👂', color: '#60a5fa', tag: '讀音一樣，看意思選字' },
    { cat: '形近字辨析', icon: '🔍', color: '#34d399', tag: '一邊是意思，一邊是聲音' },
    { cat: '常見錯別字', icon: '✏️', color: '#fb7185', tag: '先想意思，再看部首' },
    { cat: '部首辨識', icon: '🧩', color: '#fbbf24', tag: '部首表示字的意思' },
    { cat: '詞義理解與應用', icon: '💬', color: '#f97316', tag: '找線索，換詞讀一讀' },
    { cat: '成語', icon: '📜', color: '#2dd4bf', tag: '放回句子讀讀看' }
  ];
  var GRADES = ['三年級', '四年級', '五年級', '六年級'];

  function qid(q) { return String(q['題目ID'] || q['題號'] || ''); }

  /* 題目依島、年級分組；單元內依年度新到舊、題號小到大 */
  function buildTopics(all) {
    return ISLANDS.map(function (isl) {
      var units = GRADES.map(function (g) {
        var qs = (all || []).filter(function (q) { return (q['類別'] || '').trim() === isl.cat && (q['年級'] || '').trim() === g; });
        qs.sort(function (a, b) { return (Number(b['年度']) - Number(a['年度'])) || (Number(a['題號']) - Number(b['題號'])); });
        return { grade: g, questions: qs };
      }).filter(function (u) { return u.questions.length; });
      return { cat: isl.cat, icon: isl.icon, color: isl.color, tag: isl.tag, units: units, total: units.reduce(function (n, u) { return n + u.questions.length; }, 0) };
    }).filter(function (isl) { return isl.units.length; });
  }

  /* 關卡上的短標籤：「多音字 (相)」→ 相；「同音字辨析 (提/題)」→ 提/題 */
  function shortLabel(q) {
    var t = String(q['題目重點'] || '').trim();
    var m = t.match(/[（(]([^（()）]{1,12})[)）]/);
    if (m) return m[1];
    return (t || ('第' + (q['題號'] || '') + '題')).slice(0, 6);
  }

  /* 某一題在哪個島、哪個單元、第幾關 */
  function locate(topics, id) {
    for (var i = 0; i < topics.length; i++) {
      for (var j = 0; j < topics[i].units.length; j++) {
        var qs = topics[i].units[j].questions;
        for (var k = 0; k < qs.length; k++) if (qid(qs[k]) === String(id)) return { cat: topics[i].cat, grade: topics[i].units[j].grade, idx: k };
      }
    }
    return null;
  }

  /* 網址狀態：?view=all｜?t=題型&g=年級&i=關（從 0 起；等於題數＝完成關） */
  function parseState(search) {
    var p;
    try { p = new URLSearchParams(String(search || '')); } catch (e) { return { mode: 'home' }; }
    if (p.get('view') === 'all') return { mode: 'all' };
    var t = p.get('t'), g = p.get('g');
    if (t && g) return { mode: 'unit', cat: t, grade: g, idx: Math.max(0, parseInt(p.get('i'), 10) || 0) };
    return { mode: 'home' };
  }
  function stateQuery(st) {
    if (st.mode === 'all') return '?view=all';
    if (st.mode === 'unit') return '?t=' + encodeURIComponent(st.cat) + '&g=' + encodeURIComponent(st.grade) + '&i=' + (st.idx || 0);
    return '';
  }

  var core = { ISLANDS: ISLANDS, GRADES: GRADES, buildTopics: buildTopics, shortLabel: shortLabel, locate: locate, parseState: parseState, stateQuery: stateQuery };
  if (typeof module !== 'undefined' && module.exports) { module.exports = core; return; }
  root.ClassTopicsCore = core;
})(typeof window !== 'undefined' ? window : globalThis);
