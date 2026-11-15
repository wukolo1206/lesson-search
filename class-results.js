/* 全班成效紀錄（2026-11-15）
 * 記「全班這一輪」在一起動腦原題、換個句子的第一次答對情形，存在這台電腦（localStorage）。
 * 講義題目卡顯示「第一次上課 → 最近一次」的對照，老師一眼看出哪幾題要再教。
 * 注意：這是老師帶全班時的結果，不是每個學生的成績。
 */
(function (root) {
  'use strict';
  var KEY = 'ziyin_class_results_v1';
  var KINDS = { guided: '原題', apply: '換個句子' };

  function today(d) {
    d = d || new Date();
    var m = d.getMonth() + 1, day = d.getDate();
    return d.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (day < 10 ? '0' : '') + day;
  }

  function makeStore(ls) {
    function read() {
      try { var v = JSON.parse((ls && ls.getItem(KEY)) || '{}'); return v && typeof v === 'object' ? v : {}; } catch (e) { return {}; }
    }
    function write(all) { try { ls.setItem(KEY, JSON.stringify(all)); return true; } catch (e) { return false; } }
    return { read: read, write: write };
  }

  /* 新增一筆：kind＝guided／apply，first＝第一次就答對的數量，total＝題數 */
  function add(store, qid, kind, first, total, date) {
    if (!qid || !KINDS[kind] || !(total > 0)) return false;
    var all = store.read(), id = String(qid);
    (all[id] = all[id] || []).push({ d: date || today(), k: kind, f: Math.max(0, Math.min(first, total)), t: total });
    if (all[id].length > 30) all[id] = all[id].slice(-30);
    return store.write(all);
  }

  /* 每一種紀錄：第一次的日期與結果、最近一次（不同天才算）的日期與結果 */
  function summary(list) {
    var out = {};
    (list || []).forEach(function (r) {
      var s = out[r.k] || (out[r.k] = { first: r, last: null, n: 0 });
      s.n++;
      if (r.d !== s.first.d) s.last = r;
    });
    return out;
  }

  function md(d) { var p = String(d).split('-'); return p.length === 3 ? Number(p[1]) + '/' + Number(p[2]) : d; }
  function score(r) { return r.t === 1 ? (r.f ? '第一次就對' : '試了才對') : r.f + '/' + r.t; }

  /* 題目卡上那一行：「原題：11/15 第一次就對｜換個句子：11/15 1/2 → 11/18 2/2」 */
  function line(list) {
    var s = summary(list), parts = [];
    ['guided', 'apply'].forEach(function (k) {
      if (!s[k]) return;
      var t = KINDS[k] + '：' + md(s[k].first.d) + ' ' + score(s[k].first);
      if (s[k].last) t += ' → ' + md(s[k].last.d) + ' ' + score(s[k].last);
      parts.push(t);
    });
    return parts.join('｜');
  }

  var core = { add: add, summary: summary, line: line, today: today, makeStore: makeStore, KEY: KEY };
  if (typeof module !== 'undefined' && module.exports) { module.exports = core; return; }

  var store = makeStore(root.localStorage);
  root.ClassResults = {
    KEY: KEY,
    add: function (qid, kind, first, total) {
      var ok = add(store, qid, kind, first, total);
      try { root.dispatchEvent(new CustomEvent('class-results-change', { detail: { qid: String(qid) } })); } catch (e) { /* 舊瀏覽器 */ }
      return ok;
    },
    list: function (qid) { return store.read()[String(qid)] || []; },
    line: function (qid) { return line(store.read()[String(qid)] || []); },
    clear: function () { return store.write({}); }
  };
})(typeof window !== 'undefined' ? window : globalThis);
