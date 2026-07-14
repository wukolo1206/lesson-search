// zhu-core.js — 三模式共用：開機、索引、boardState、詞籃
var ZhuCore = (function () {
  'use strict';

  var store = null;
  var index = null;
  var radicals = null;
  var tables = null;
  var basket = [];

  var state = {
    char: null,
    contextQuestionId: null,
    grade: '三年級',
    publisher: '康軒',
  };

  function init(localStorageImpl) {
    store = ZhuStore.create(localStorageImpl);
    state.grade = (store.get('prefs', {}).grade) || '三年級';
    basket = store.get('basket', []);
    return store;
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

  function clearAll() {
    store.clearAll();
    basket = [];
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
    setGrade: setGrade,
    addToBasket: addToBasket,
    removeFromBasket: removeFromBasket,
    isInBasket: isInBasket,
    clearAll: clearAll,
  };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = ZhuCore;
