/* 課堂講義版面切換（2026-11-15）
 * 原本／a 闖關風／b 大字投影／c 闖關＋類別標籤；記在這台瀏覽器（localStorage，讀不到也能用）。
 * 題目卡依「類別」上色加圖示；c 的類別標籤一按＝設定原本的「類別」篩選（#fCat）並觸發 change。
 * 樣式：class-themes.css
 */
(function (root) {
  'use strict';
  var KEY = 'ziyin-class-theme';
  var THEMES = [['', '原本'], ['a', '闖關風'], ['b', '大字投影'], ['c', '闖關＋標籤']];
  var CAT = {
    '多音字': ['#a78bfa', '🔊'], '同音字辨析': ['#60a5fa', '👂'], '形近字辨析': ['#34d399', '🔍'],
    '常見錯別字': ['#fb7185', '✏️'], '部首辨識': ['#fbbf24', '🧩'], '詞義理解與應用': ['#f97316', '💬'], '成語': ['#2dd4bf', '📜']
  };

  function catOf(name) { return CAT[String(name || '').trim()] || null; }

  if (typeof module !== 'undefined' && module.exports) { module.exports = { THEMES: THEMES, CAT: CAT, catOf: catOf }; return; }

  function load() { try { return root.localStorage.getItem(KEY) || ''; } catch (e) { return ''; } }
  function save(v) { try { root.localStorage.setItem(KEY, v); } catch (e) { /* 私密視窗等情況存不了也沒關係 */ } }
  function valid(v) { return THEMES.some(function (t) { return t[0] === v; }) ? v : ''; }

  function apply(v) {
    v = valid(v);
    if (v) document.documentElement.setAttribute('data-theme', v);
    else document.documentElement.removeAttribute('data-theme');
    var sw = document.getElementById('thSwitch');
    if (sw) Array.prototype.forEach.call(sw.querySelectorAll('button'), function (b) { b.classList.toggle('on', b.getAttribute('data-th') === v); });
  }
  apply(load());                                  // 盡早套用，避免先閃一下原本版面

  function mountSwitch() {
    if (document.getElementById('thSwitch')) return;
    var box = document.querySelector('.header-links');
    if (!box) return;
    var el = document.createElement('div');
    el.id = 'thSwitch'; el.className = 'th-switch'; el.setAttribute('role', 'group'); el.setAttribute('aria-label', '版面');
    el.innerHTML = '<span class="th-lbl">版面</span>' + THEMES.map(function (t) {
      return '<button type="button" data-th="' + t[0] + '">' + t[1] + '</button>';
    }).join('');
    el.addEventListener('click', function (e) {
      var b = e.target.closest('button[data-th]');
      if (!b) return;
      var v = b.getAttribute('data-th'); save(v); apply(v);
    });
    box.insertBefore(el, box.firstChild);
    apply(load());
  }

  function decorate() {
    Array.prototype.forEach.call(document.querySelectorAll('.lesson-unit'), function (u) {
      if (u.getAttribute('data-th-done')) return;
      var b = u.querySelector('.b-ct'), c = b && catOf(b.textContent);
      if (c) u.style.setProperty('--cat', c[0]);
      var t = u.querySelector('.unit-title');
      if (t && c && !t.querySelector('.th-ico')) t.insertAdjacentHTML('afterbegin', '<span class="th-ico" aria-hidden="true">' + c[1] + '</span>');
      u.setAttribute('data-th-done', '1');
    });
    syncChips();
  }

  function mountChips() {
    var sel = document.getElementById('fCat'), fb = document.querySelector('.filter-bar');
    if (!sel || !fb || sel.options.length < 2) return;
    var fg = sel.closest('.fg'); if (fg) fg.classList.add('fg-cat');
    var bar = document.getElementById('chipbar');
    if (!bar) { bar = document.createElement('div'); bar.id = 'chipbar'; fb.parentNode.insertBefore(bar, fb); }
    var html = '<button type="button" class="chip" data-v="" style="--cat:#64748b">全部</button>';
    Array.prototype.forEach.call(sel.options, function (o) {
      if (!o.value) return;
      var c = catOf(o.value) || ['#94a3b8', '📘'];
      html += '<button type="button" class="chip" data-v="' + o.value.replace(/"/g, '&quot;') + '" style="--cat:' + c[0] + '">' + c[1] + ' ' + o.value.replace(/</g, '&lt;') + '</button>';
    });
    if (bar.getAttribute('data-n') === String(sel.options.length)) return syncChips();
    bar.innerHTML = html; bar.setAttribute('data-n', String(sel.options.length));
    if (!bar._bound) {
      bar.addEventListener('click', function (e) {
        var b = e.target.closest('.chip'); if (!b) return;
        sel.value = b.getAttribute('data-v');
        sel.dispatchEvent(new Event('change'));
        syncChips();
      });
      sel.addEventListener('change', syncChips);
      bar._bound = true;
    }
    syncChips();
  }

  function syncChips() {
    var sel = document.getElementById('fCat'), bar = document.getElementById('chipbar');
    if (!sel || !bar) return;
    Array.prototype.forEach.call(bar.querySelectorAll('.chip'), function (b) { b.classList.toggle('on', b.getAttribute('data-v') === sel.value); });
  }

  function start() {
    mountSwitch();
    decorate(); mountChips();
    var grid = document.getElementById('grid');
    if (grid && root.MutationObserver) new MutationObserver(function () { decorate(); mountChips(); }).observe(grid, { childList: true });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();

  root.ClassThemes = { apply: apply, decorate: decorate };
})(typeof window !== 'undefined' ? window : globalThis);
