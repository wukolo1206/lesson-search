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

  function queryTerm(search) {
    var raw = '';
    try {
      raw = new URLSearchParams(String(search || '')).get('q') || '';
    } catch (e) {
      return '';
    }
    var term = normalizeTerm(raw);
    return isValidTerm(term) ? term : '';
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

  function listValue(value) {
    if (Array.isArray(value)) return value;
    return value === undefined || value === null || value === '' ? [] : [value];
  }

  function uniqueStrings(values) {
    var out = [];
    (values || []).forEach(function (value) {
      var s = String(value || '').trim();
      if (s && out.indexOf(s) === -1) out.push(s);
    });
    return out;
  }

  function definitionGloss(definition) {
    var type = String(definition.type || '').trim();
    var def = String(definition.def || '').trim();
    if (!def) return '';
    return type ? '（' + type + '）' + def : def;
  }

  function normalizeMoedict(raw, term) {
    if (!raw || !Array.isArray(raw.heteronyms)) return null;
    var heteronyms = raw.heteronyms.filter(function (item) {
      return item && Array.isArray(item.definitions) && item.definitions.length;
    });
    var allDefinitions = [];
    heteronyms.forEach(function (heteronym) {
      heteronym.definitions.forEach(function (definition) {
        if (definition && definition.def) allDefinitions.push(definition);
      });
    });
    if (!allDefinitions.length) return null;

    var readings = uniqueStrings(heteronyms.map(function (item) { return item.bopomofo; }));
    var examples = [];
    allDefinitions.forEach(function (definition) {
      var candidates = listValue(definition.example);
      if (!candidates.length) candidates = listValue(definition.quote).slice(0, 1);
      candidates.forEach(function (example) {
        var s = String(example || '').trim();
        if (s && examples.indexOf(s) === -1 && examples.length < 4) examples.push(s);
      });
    });

    var extraReadings = heteronyms.slice(1).map(function (heteronym) {
      return {
        zhuyin: String(heteronym.bopomofo || '').trim(),
        gloss: heteronym.definitions.map(function (definition) {
          return String(definition.def || '').trim();
        }).filter(Boolean).join('；')
      };
    }).filter(function (reading) { return reading.zhuyin || reading.gloss; });

    var normalizedTerm = String(term || raw.title || '').trim();
    return {
      term: normalizedTerm,
      zhuyin: readings.join('｜'),
      explanation: allDefinitions.map(definitionGloss).filter(Boolean).join('\n'),
      examples: examples,
      extraReadings: extraReadings,
      imageSuitable: null,
      imageStatus: 'none',
      imageFileId: '',
      source: 'moedict',
      sourceUrl: 'https://www.moedict.tw/' + encodeURIComponent(normalizedTerm)
    };
  }

  return { normalizeTerm: normalizeTerm, isValidTerm: isValidTerm, queryTerm: queryTerm,
    splitHighlight: splitHighlight, fileIdOk: fileIdOk, thumbUrl: thumbUrl,
    makeCache: makeCache, pollDelay: pollDelay, normalizeMoedict: normalizeMoedict, TTL: TTL };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = LookupCore;
