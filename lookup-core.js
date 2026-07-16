// 解釋卡前端純邏輯層（無 DOM、無 fetch；node --test 可測）
var LookupCore = (function () {
  'use strict';

  function normalizeTerm(raw) {
    var s = String(raw || '');
    if (s.normalize) s = s.normalize('NFC');
    s = s.replace(/[\u200B-\u200F\u2060\uFEFF]/g, '');
    var punct = '[\\s\\u3000\\u3001\\u3002\\uFF0C\\uFF01\\uFF1F\\uFF1B\\uFF1A\\u300C\\u300D\\u300E\\u300F\\uFF08\\uFF09()!?,.;:\'\"]+';
    s = s.replace(new RegExp('^' + punct + '|' + punct + '$', 'g'), '');
    return s;
  }

  function isValidTerm(term) {
    if (!term) return false;
    var cps = Array.from(term);
    if (cps.length < 1 || cps.length > 10) return false;
    return cps.every(function (c) {
      var cp = c.codePointAt(0);
      return (cp >= 0x4E00 && cp <= 0x9FFF) || (cp >= 0x3400 && cp <= 0x4DBF) ||
             (cp >= 0xF900 && cp <= 0xFAFF) || (cp >= 0x20000 && cp <= 0x2EBEF);
    });
  }

  // 例句標色：切成 [{text, mark}] 片段，由 UI 層轉成文字節點與 <mark>
  function splitHighlight(sentence, term) {
    var out = [];
    var s = String(sentence || ''), t = String(term || '');
    if (!s) return out;
    if (!t) return [{ text: s, mark: false }];
    var idx, from = 0;
    while ((idx = s.indexOf(t, from)) !== -1) {
      if (idx > from) out.push({ text: s.slice(from, idx), mark: false });
      out.push({ text: t, mark: true });
      from = idx + t.length;
    }
    if (from < s.length) out.push({ text: s.slice(from), mark: false });
    return out;
  }

  function fileIdOk(id) { return /^[A-Za-z0-9_-]{20,60}$/.test(String(id || '')); }
  function thumbUrl(id) { return 'https:/' + '/drive.google.com/thumbnail?id=' + id + '&sz=w800'; }

  var TTL = 7 * 24 * 60 * 60 * 1000;
  function makeCache(storage, nowFn) {
    function key(term) { return 'lk.v1.' + term; }
    return {
      get: function (term) {
        try {
          var raw = storage.getItem(key(term));
          if (!raw) return null;
          var obj = JSON.parse(raw);
          if (!obj || !obj.exp || obj.exp < nowFn()) { storage.removeItem(key(term)); return null; }
          return obj.data;
        } catch (e) { return null; }
      },
      put: function (term, data) {
        try { storage.setItem(key(term), JSON.stringify({ exp: nowFn() + TTL, data: data })); }
        catch (e) { /* 空間滿：略過，不影響功能 */ }
      },
      remove: function (term) { try { storage.removeItem(key(term)); } catch (e) {} }
    };
  }

  function pollDelay(retryAfterMs, rand) {
    var base = Number(retryAfterMs) > 0 ? Number(retryAfterMs) : 2000;
    return base + Math.floor((rand || 0) * 500);
  }

  return { normalizeTerm: normalizeTerm, isValidTerm: isValidTerm,
    splitHighlight: splitHighlight, fileIdOk: fileIdOk, thumbUrl: thumbUrl,
    makeCache: makeCache, pollDelay: pollDelay, TTL: TTL };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = LookupCore;
