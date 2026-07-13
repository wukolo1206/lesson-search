var ZhuData = (function () {
  'use strict';

  function charsOfShapeRow(row) {
    var c = (row['字'] || '').trim();
    return c ? [c] : [];
  }

  function charsOfSoundRow(row) {
    var c = (row['破音字'] || '').trim();
    return c ? [c] : [];
  }

  function parseMistakeRow(row) {
    var bracket = /【(.)】/.exec(row['錯誤寫法'] || '');
    var bracketChar = bracket ? bracket[1] : null;
    var meaning = row['字義說明'] || '';
    var match = /—\s*(.)\s*\(/.exec(meaning) || /正確用字：\s*(.)\s*\(/.exec(meaning);
    var parsedCorrect = match ? match[1] : null;

    if (/^✓/.test((row['正誤'] || '').trim())) {
      return { wrongChar: null, correctChar: bracketChar };
    }
    return { wrongChar: bracketChar, correctChar: parsedCorrect };
  }

  function charsOfMistakeRow(row) {
    var parsed = parseMistakeRow(row);
    var chars = [];
    if (parsed.wrongChar) chars.push(parsed.wrongChar);
    if (parsed.correctChar) chars.push(parsed.correctChar);
    return chars;
  }

  function parseRadical(text) {
    var match = /([一-鿿])部/.exec(text || '');
    if (!match) return null;
    return match[1] === '氵' ? '水' : match[1];
  }

  function buildRadicalIndex(shapeRows, mistakeRows) {
    var seen = {};

    function record(char, radical) {
      if (!char || !radical) return;
      if (!seen[char]) seen[char] = [];
      if (seen[char].indexOf(radical) === -1) seen[char].push(radical);
    }

    (shapeRows || []).forEach(function (row) {
      record((row['字'] || '').trim(), parseRadical(row['解釋']));
    });
    (mistakeRows || []).forEach(function (row) {
      var parsed = parseMistakeRow(row);
      record(parsed.correctChar, parseRadical(row['字義說明']));
    });

    return {
      radicalOf: function (char) {
        var radicals = seen[char];
        return radicals && radicals.length === 1 ? radicals[0] : null;
      },
      conflicts: function () {
        return Object.keys(seen)
          .filter(function (char) { return seen[char].length > 1; })
          .map(function (char) { return { char: char, radicals: seen[char] }; });
      }
    };
  }

  function buildIndex(tables) {
    var questions = tables.questions || [];
    var questionsById = {};
    var wordsByChar = {};
    var questionIdsByChar = {};
    questions.forEach(function (question) { questionsById[String(question['題目ID'])] = question; });

    function gradeOf(questionId) {
      var question = questionsById[String(questionId)];
      return question ? (question['年級'] || '') : '';
    }

    function addWord(char, entry) {
      if (!char || !entry.word) return;
      if (!wordsByChar[char]) wordsByChar[char] = [];
      wordsByChar[char].push(entry);
    }

    function addQuestion(char, questionId) {
      if (!char || !questionId) return;
      var id = String(questionId);
      if (!questionIdsByChar[char]) questionIdsByChar[char] = [];
      if (questionIdsByChar[char].indexOf(id) === -1) questionIdsByChar[char].push(id);
    }

    function addEntries(rows, charsOf, makeEntry) {
      (rows || []).forEach(function (row) {
        var questionId = String(row['題目ID']);
        charsOf(row).forEach(function (char) {
          addQuestion(char, questionId);
          addWord(char, makeEntry(row, char, questionId));
        });
      });
    }

    addEntries(tables.sound, charsOfSoundRow, function (row, char, questionId) {
      return { word: (row['詞語'] || '').trim(), char: char, bopomofo: (row['注音'] || '').trim(), gloss: (row['字義說明'] || '').trim(), questionId: questionId, grade: gradeOf(questionId), source: 'sound', optionOf: (row['來源'] || '').trim(), wrongChar: null, correctChar: null };
    });
    addEntries(tables.shape, charsOfShapeRow, function (row, char, questionId) {
      return { word: (row['語詞'] || '').trim(), char: char, bopomofo: '', gloss: (row['解釋'] || '').trim(), questionId: questionId, grade: gradeOf(questionId), source: 'shape', optionOf: '', wrongChar: null, correctChar: null };
    });
    addEntries(tables.mistake, charsOfMistakeRow, function (row, char, questionId) {
      var parsed = parseMistakeRow(row);
      return { word: (row['正確詞語'] || '').trim(), char: char, bopomofo: '', gloss: (row['字義說明'] || '').trim(), questionId: questionId, grade: gradeOf(questionId), source: 'mistake', optionOf: '', wrongChar: parsed.wrongChar, correctChar: parsed.correctChar };
    });

    return {
      wordsOf: function (char) { return (wordsByChar[char] || []).slice(); },
      questionsOf: function (char) {
        return (questionIdsByChar[char] || []).map(function (id) { return questionsById[id]; }).filter(Boolean);
      },
      allChars: function () { return Object.keys(wordsByChar); }
    };
  }

  var GRADE_DIGITS = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6 };

  function gradeNum(gradeText) {
    var match = /^([一二三四五六])年級/.exec((gradeText || '').trim());
    return match ? GRADE_DIGITS[match[1]] : null;
  }

  function optionWords(index, state) {
    if (!state.contextQuestionId) return [];
    return index.wordsOf(state.char).filter(function (word) {
      return String(word.questionId) === String(state.contextQuestionId);
    });
  }

  function supplementWords(index, state) {
    var excluded = {};
    var seen = {};
    var currentGrade = gradeNum(state.grade);
    optionWords(index, state).forEach(function (word) { excluded[word.word] = true; });

    return index.wordsOf(state.char)
      .filter(function (word) {
        if (excluded[word.word] || seen[word.word]) return false;
        seen[word.word] = true;
        return true;
      })
      .map(function (word, order) {
        var wordGrade = gradeNum(word.grade);
        return { word: word, order: order, distance: currentGrade === null || wordGrade === null ? 99 : Math.abs(wordGrade - currentGrade) };
      })
      .sort(function (left, right) { return left.distance - right.distance || left.order - right.order; })
      .map(function (item) { return item.word; });
  }

  var SUPPLEMENT_VISIBLE = 6;
  var basketOps = {
    keyOf: function (item) { return item.word + '@' + item.char; },
    add: function (list, item) { var key = basketOps.keyOf(item); return list.some(function (x) { return basketOps.keyOf(x) === key; }) ? list.slice() : list.concat([item]); },
    remove: function (list, word, char) { var key = word + '@' + char; return list.filter(function (item) { return basketOps.keyOf(item) !== key; }); },
    sortByGrade: function (list) { return list.map(function (item, order) { return { item: item, grade: gradeNum(item.grade), order: order }; }).sort(function (a, b) { if (a.grade === null) return b.grade === null ? a.order - b.order : 1; if (b.grade === null) return -1; return a.grade - b.grade || a.order - b.order; }).map(function (entry) { return entry.item; }); }
  };

  var SHEET_ID = '1MoZ0VhgJ9fSWktZUjunX1IMJwDuFBP_qyZFuEnH6uGo';
  var TABS = { questions: '題目主表', sound: '破音字細目', mistake: '錯別字細目', shape: '形近字細目', apply: '語詞應用展開' };
  var OPTIONAL_TABS = ['apply'];
  function csvUrl(name) { return 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/gviz/tq?tqx=out:csv&sheet=' + encodeURIComponent(name); }
  function fetchTable(name) {
    return new Promise(function (resolve, reject) {
      Papa.parse(csvUrl(name), { download: true, header: true, skipEmptyLines: true, complete: function (result) { resolve(result.data); }, error: reject });
    });
  }
  function loadAll() {
    var names = Object.keys(TABS);
    return Promise.all(names.map(function (key) { return fetchTable(TABS[key]).catch(function (error) { if (OPTIONAL_TABS.indexOf(key) !== -1) { console.warn('[zhu] optional table unavailable', error); return []; } throw new Error('分頁「' + TABS[key] + '」載入失敗'); }); })).then(function (results) { var tables = {}; names.forEach(function (key, i) { tables[key] = results[i]; }); return tables; });
  }

  return {
    charsOfShapeRow: charsOfShapeRow,
    charsOfSoundRow: charsOfSoundRow,
    parseMistakeRow: parseMistakeRow,
    charsOfMistakeRow: charsOfMistakeRow,
    parseRadical: parseRadical,
    buildRadicalIndex: buildRadicalIndex,
    buildIndex: buildIndex,
    gradeNum: gradeNum,
    optionWords: optionWords,
    supplementWords: supplementWords,
    SUPPLEMENT_VISIBLE: SUPPLEMENT_VISIBLE,
    basketOps: basketOps,
    loadAll: loadAll,
    TABS: TABS,
    csvUrl: csvUrl
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = ZhuData;
