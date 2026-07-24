var ZhuStore = (function () {
  'use strict';
  var VERSION = 1;
  var NAMES = ['basket', 'ink', 'prefs', 'selection', 'wordApplication'];
  var SOFT_LIMIT_BYTES = 4 * 1024 * 1024;
  function keyOf(name) { return 'zhu.' + name + '.v' + VERSION; }
  function create(storage) {
    function get(name, fallback) {
      try { var raw = storage.getItem(keyOf(name)); var parsed = raw === null ? null : JSON.parse(raw); return parsed && parsed.v === VERSION ? parsed.data : fallback; } catch (e) { return fallback; }
    }
    function set(name, data) {
      try { var payload = JSON.stringify({ v: VERSION, data: data }); if (payload.length > SOFT_LIMIT_BYTES) return false; storage.setItem(keyOf(name), payload); return true; } catch (e) { return false; }
    }
    function clearAll() { NAMES.forEach(function (name) { try { storage.removeItem(keyOf(name)); } catch (e) {} }); }
    return { get: get, set: set, clearAll: clearAll };
  }
  return { create: create, VERSION: VERSION, SOFT_LIMIT_BYTES: SOFT_LIMIT_BYTES };
})();
if (typeof module !== 'undefined' && module.exports) module.exports = ZhuStore;
