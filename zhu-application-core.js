(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
  } else {
    root.ZhuApplicationCore = factory();
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var LEVELS = ['low', 'middle', 'high'];
  var MODES = ['passage', 'sentences'];
  var FORMAT_ERROR = '回應格式不正確';
  var CONFIG_ERROR = '回應設定不一致';
  var TERMS_ERROR = '回應詞語不一致';
  var CONTENT_ERROR = '內容未完整包含所選詞語';

  function normalizeTerms(items) {
    var source = Array.isArray(items) ? items : (items === null || items === undefined ? [] : [items]);
    var terms = [];

    source.forEach(function (item) {
      var raw = typeof item === 'string' ? item : item && typeof item === 'object' ? item.word : null;
      if (typeof raw !== 'string') return;

      var term = raw.trim().replace(/\s+/g, '');
      if (!term || terms.indexOf(term) !== -1) return;
      terms.push(term);
    });

    return terms;
  }

  function validateConfig(config) {
    var terms = normalizeTerms(config && config.terms);
    if (terms.length < 1) return '請至少選擇一個詞';
    if (terms.length > 10) return '一次最多選擇 10 個詞';
    if (terms.join('').length > 60) return '所選詞語總長度過長';
    if (!config || LEVELS.indexOf(config.level) === -1) return '年段設定不正確';
    if (MODES.indexOf(config.mode) === -1) return '文章形式不正確';
    return null;
  }

  function configSignature(config) {
    config = config || {};
    return normalizeTerms(config.terms).join('\u001f') + '\u001e' + valueOf(config.level) + '\u001e' + valueOf(config.mode);
  }

  function valueOf(value) {
    return value === null || value === undefined ? '' : String(value);
  }

  function sameArray(left, right) {
    if (left.length !== right.length) return false;
    for (var i = 0; i < left.length; i += 1) {
      if (left[i] !== right[i]) return false;
    }
    return true;
  }

  function isObject(value) {
    return value !== null && typeof value === 'object' && !Array.isArray(value);
  }

  function hasOnlyKeys(value, allowedKeys) {
    return Object.keys(value).every(function (key) {
      return allowedKeys.indexOf(key) !== -1;
    });
  }

  function validateResult(config, result) {
    if (!isObject(result) || result.schemaVersion !== 1) return FORMAT_ERROR;
    if (!hasOnlyKeys(result, ['schemaVersion', 'terms', 'level', 'mode', 'passage', 'sentences', 'generatedAt', 'model'])) {
      return FORMAT_ERROR;
    }
    if (!Array.isArray(result.terms) || result.terms.some(function (term) { return typeof term !== 'string'; })) {
      return FORMAT_ERROR;
    }
    if (typeof result.level !== 'string' || typeof result.mode !== 'string') return FORMAT_ERROR;
    if (Object.prototype.hasOwnProperty.call(result, 'generatedAt') && typeof result.generatedAt !== 'string') return FORMAT_ERROR;
    if (Object.prototype.hasOwnProperty.call(result, 'model') && typeof result.model !== 'string') return FORMAT_ERROR;
    if (typeof result.passage !== 'string' || !Array.isArray(result.sentences)) return FORMAT_ERROR;

    config = config || {};
    if (result.level !== config.level || result.mode !== config.mode) return CONFIG_ERROR;

    var terms = normalizeTerms(config.terms);
    if (!sameArray(result.terms, terms)) return TERMS_ERROR;

    if (result.mode === 'passage') {
      if (result.sentences.length !== 0) return FORMAT_ERROR;
      if (!result.passage.trim()) return CONTENT_ERROR;
      for (var passageIndex = 0; passageIndex < terms.length; passageIndex += 1) {
        if (result.passage.indexOf(terms[passageIndex]) === -1) return CONTENT_ERROR;
      }
      return null;
    }

    if (result.mode !== 'sentences') return FORMAT_ERROR;
    if (result.passage !== '') return FORMAT_ERROR;
    if (result.sentences.length !== terms.length) return TERMS_ERROR;

    var generatedText = '';
    for (var sentenceIndex = 0; sentenceIndex < result.sentences.length; sentenceIndex += 1) {
      var sentence = result.sentences[sentenceIndex];
      if (!isObject(sentence) || typeof sentence.term !== 'string' || typeof sentence.text !== 'string') {
        return FORMAT_ERROR;
      }
      if (!hasOnlyKeys(sentence, ['term', 'text'])) return FORMAT_ERROR;
      if (sentence.term !== terms[sentenceIndex]) return TERMS_ERROR;
      if (sentence.text.indexOf(sentence.term) === -1) return CONTENT_ERROR;
      generatedText += sentence.text;
    }

    for (var termIndex = 0; termIndex < terms.length; termIndex += 1) {
      if (generatedText.indexOf(terms[termIndex]) === -1) return CONTENT_ERROR;
    }
    return null;
  }

  function splitTerms(text, terms) {
    text = valueOf(text);
    var candidates = normalizeTerms(terms).sort(function (left, right) {
      return right.length - left.length;
    });
    var parts = [];
    var ordinaryText = '';
    var position = 0;

    function flushOrdinaryText() {
      if (!ordinaryText) return;
      parts.push({ text: ordinaryText, term: '' });
      ordinaryText = '';
    }

    while (position < text.length) {
      var matchedTerm = null;
      for (var termIndex = 0; termIndex < candidates.length; termIndex += 1) {
        var candidate = candidates[termIndex];
        if (text.substr(position, candidate.length) === candidate) {
          matchedTerm = candidate;
          break;
        }
      }

      if (matchedTerm) {
        flushOrdinaryText();
        parts.push({ text: matchedTerm, term: matchedTerm });
        position += matchedTerm.length;
      } else {
        ordinaryText += text.charAt(position);
        position += 1;
      }
    }

    flushOrdinaryText();
    return parts;
  }

  function toStudentText(parts) {
    return (Array.isArray(parts) ? parts : []).map(function (part) {
      if (!part) return '';
      if (part.term) return '＿＿';
      return valueOf(part.text);
    }).join('');
  }

  function matchesStored(stored, config) {
    return isObject(stored) && stored.configSignature === configSignature(config);
  }

  function errorMessage(error) {
    var code = error && error.code || '';
    if (code === 'QUOTA_EXCEEDED') return '今日 AI 額度已用完，請稍後再試';
    if (code === 'RATE_LIMITED') return '目前使用人數較多，請稍後再試';
    if (code === 'SAFETY_BLOCKED') return '這組詞語無法產生合適內容，請調整後再試';
    if (code === 'INVALID_INPUT') return error.message || '所選詞語設定不正確';
    if (code === 'CONTENT_INVALID') return '內容未完整包含所選詞語，請重新生成';
    if (code === 'TIMEOUT') return '生成時間較久，請稍後再試';
    return '目前無法完成生成，請檢查網路後再試';
  }

  return {
    LEVELS: LEVELS,
    MODES: MODES,
    normalizeTerms: normalizeTerms,
    validateConfig: validateConfig,
    configSignature: configSignature,
    validateResult: validateResult,
    splitTerms: splitTerms,
    toStudentText: toStudentText,
    matchesStored: matchesStored,
    errorMessage: errorMessage
  };
}));
