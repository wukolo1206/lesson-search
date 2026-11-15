// zhu-core.js — 三模式共用：開機、索引、boardState、詞籃
var ZhuCore = (function () {
  'use strict';

  var store = null;
  var index = null;
  var radicals = null;
  var tables = null;
  var basket = [];
  var selectedChars = [];
  var taughtChars = [];          // 字主板手動標記的已教字（zhu.taught.v1）
  var activeSelectedIndex = 0;
  var lsImpl = null;
  // 講義「已教過」（以題記）。全站共用這一份：字主板讀它推算已教字，標記時有題目情境也寫回它。
  var CLASS_TAUGHT_KEY = 'taught_questions_v1';

  var state = {
    char: null,
    contextQuestionId: null,
    grade: '三年級',
    publisher: '康軒',
  };

  function init(localStorageImpl) {
    lsImpl = localStorageImpl || (typeof window !== 'undefined' ? window.localStorage : null);
    store = ZhuStore.create(localStorageImpl);
    state.grade = (store.get('prefs', {}).grade) || '三年級';
    basket = store.get('basket', []);
    selectedChars = normalizeSelectedChars(store.get('selection', []));
    taughtChars = normalizeTaughtChars(store.get('taught', []));
    activeSelectedIndex = 0;
    return store;
  }

  function normalizeSelectedChars(queue) {
    var out = [];
    (Array.isArray(queue) ? queue : []).forEach(function (entry) {
      if (!entry || !entry.char || out.some(function (item) { return item.char === entry.char; })) return;
      out.push({
        char: String(entry.char),
        contextQuestionId: entry.contextQuestionId === null || entry.contextQuestionId === undefined ? null : String(entry.contextQuestionId),
      });
    });
    return out;
  }

  function persistSelectedChars() {
    return store.set('selection', selectedChars);
  }

  function normalizeTaughtChars(chars) {
    var out = [];
    (Array.isArray(chars) ? chars : []).forEach(function (char) {
      char = String(char || '').trim();
      if (!/^[\u3400-\u9fff]$/.test(char) || out.indexOf(char) !== -1) return;
      out.push(char);
    });
    return out;
  }

  function boot(onReady, onError) {
    ZhuData.loadAll().then(function (loaded) {
      tables = loaded;
      index = ZhuData.buildIndex(loaded);
      radicals = ZhuData.buildRadicalIndex(loaded.shape, loaded.mistake);
      onReady();
    }).catch(onError);
  }

  function getIndex() { return index; }
  function getRadicals() { return radicals; }
  function getTables() { return tables; }
  function getBasket() { return basket.slice(); }
  function getStore() { return store; }
  function getSelectedChars() { return selectedChars.map(function (entry) { return { char: entry.char, contextQuestionId: entry.contextQuestionId }; }); }
  function getActiveSelectedIndex() { return activeSelectedIndex; }
  function readClassTaught() {
    try {
      var v = JSON.parse((lsImpl && lsImpl.getItem(CLASS_TAUGHT_KEY)) || '[]');
      return Array.isArray(v) ? v.map(String) : [];
    } catch (e) { return []; }
  }
  function writeClassTaught(ids) {
    try { lsImpl.setItem(CLASS_TAUGHT_KEY, JSON.stringify(ids)); return true; } catch (e) { return false; }
  }
  function setQuestionTaught(questionId, taught) {
    if (questionId === null || questionId === undefined || questionId === '') return true;
    var id = String(questionId), ids = readClassTaught(), at = ids.indexOf(id);
    if (taught && at === -1) ids.push(id);
    if (!taught && at !== -1) ids.splice(at, 1);
    return writeClassTaught(ids);
  }
  /* 這個字考過的題目裡，講義標了已教過的 */
  function taughtQuestionsOf(char) {
    if (!index) return [];
    var set = readClassTaught();
    return index.questionsOf(char).filter(function (q) { return set.indexOf(String(q['題目ID'])) !== -1; });
  }
  function isManualTaughtChar(char) { return taughtChars.indexOf(char) !== -1; }
  function isTaughtChar(char) { return isManualTaughtChar(char) || taughtQuestionsOf(char).length > 0; }
  function getTaughtChars() {
    var out = taughtChars.slice();
    if (index) index.allChars().forEach(function (c) { if (out.indexOf(c) === -1 && taughtQuestionsOf(c).length) out.push(c); });
    return out;
  }

  /* 標記／取消已教：字主板手動標記；有題目情境時同一題也寫回講義。
     取消時若這個字還有其他題在講義標了已教過，仍會顯示已教（要到講義取消）。 */
  function setTaughtChar(char, taught, contextQuestionId) {
    char = String(char || '').trim();
    if (!/^[\u3400-\u9fff]$/.test(char)) return false;
    var next = taughtChars.slice();
    var indexValue = next.indexOf(char);
    if (taught && indexValue === -1) next.push(char);
    if (!taught && indexValue !== -1) next.splice(indexValue, 1);
    if (!store.set('taught', next)) return false;
    taughtChars = next;
    return setQuestionTaught(contextQuestionId, taught);
  }

  function resetTaughtChars() {
    if (!store.set('taught', [])) return false;
    taughtChars = [];
    return true;
  }

  function setSelectedChars(queue) {
    selectedChars = normalizeSelectedChars(queue);
    activeSelectedIndex = Math.min(activeSelectedIndex, Math.max(0, selectedChars.length - 1));
    return persistSelectedChars();
  }

  function addSelectedChar(char, contextQuestionId) {
    if (!char || selectedChars.some(function (entry) { return entry.char === char; })) return true;
    selectedChars.push({
      char: String(char),
      contextQuestionId: contextQuestionId === null || contextQuestionId === undefined ? null : String(contextQuestionId),
    });
    return persistSelectedChars();
  }

  function removeSelectedChar(char) {
    var removedIndex = selectedChars.findIndex(function (entry) { return entry.char === char; });
    if (removedIndex === -1) return true;
    selectedChars.splice(removedIndex, 1);
    if (removedIndex < activeSelectedIndex) activeSelectedIndex -= 1;
    activeSelectedIndex = Math.min(activeSelectedIndex, Math.max(0, selectedChars.length - 1));
    return persistSelectedChars();
  }

  function setActiveSelectedIndex(indexValue) {
    var max = Math.max(0, selectedChars.length - 1);
    activeSelectedIndex = Math.max(0, Math.min(max, Number(indexValue) || 0));
    return activeSelectedIndex;
  }

  function setGrade(grade) {
    state.grade = grade;
    store.set('prefs', { grade: grade });
  }

  function addToBasket(entry) {
    basket = ZhuData.basketOps.add(basket, entry);
    return store.set('basket', basket);
  }

  function removeFromBasket(word, char) {
    basket = ZhuData.basketOps.remove(basket, word, char);
    return store.set('basket', basket);
  }

  function isInBasket(word, char) {
    var key = word + '@' + char;
    return basket.some(function (x) { return x.word + '@' + x.char === key; });
  }

  function getWordApplicationResult() {
    return store.get('wordApplication', null);
  }

  function setWordApplicationResult(result) {
    return store.set('wordApplication', result);
  }

  function clearAll() {
    store.clearAll();
    basket = [];
    selectedChars = [];
    activeSelectedIndex = 0;
  }

  return {
    state: state,
    init: init,
    boot: boot,
    getIndex: getIndex,
    getRadicals: getRadicals,
    getTables: getTables,
    getBasket: getBasket,
    getStore: getStore,
    getSelectedChars: getSelectedChars,
    getTaughtChars: getTaughtChars,
    isTaughtChar: isTaughtChar,
    isManualTaughtChar: isManualTaughtChar,
    taughtQuestionsOf: taughtQuestionsOf,
    setQuestionTaught: setQuestionTaught,
    setTaughtChar: setTaughtChar,
    resetTaughtChars: resetTaughtChars,
    setSelectedChars: setSelectedChars,
    addSelectedChar: addSelectedChar,
    removeSelectedChar: removeSelectedChar,
    getActiveSelectedIndex: getActiveSelectedIndex,
    setActiveSelectedIndex: setActiveSelectedIndex,
    setGrade: setGrade,
    addToBasket: addToBasket,
    removeFromBasket: removeFromBasket,
    isInBasket: isInBasket,
    getWordApplicationResult: getWordApplicationResult,
    setWordApplicationResult: setWordApplicationResult,
    clearAll: clearAll,
  };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = ZhuCore;
