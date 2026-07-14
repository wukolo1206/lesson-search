// zhu-prep.js — 備課台：篩選＋題目→考點字
var ZhuPrep = (function () {
  'use strict';

  function filterQuestions(questions, filters) {
    filters = filters || {};
    return (questions || []).filter(function (q) {
      if (filters.grade && q['年級'] !== filters.grade) return false;
      if (filters.year && String(q['年度']) !== String(filters.year)) return false;
      if (filters.category && q['類別'] !== filters.category) return false;
      return true;
    });
  }

  function charsOfQuestion(index, questionId) {
    var id = String(questionId);
    return index.allChars().filter(function (char) {
      return index.questionsOf(char).some(function (q) { return String(q['題目ID']) === id; });
    });
  }

  // ── 以下是 DOM，Node 測試不會走到 ──────────────────────
  var GRADES = ['三年級', '四年級', '五年級', '六年級'];
  var CATEGORIES = ['同音字辨析', '常見錯別字', '形聲字與形近字讀音', '成語', '破音字', '詞義理解與應用', '部首辨識'];
  var currentFilters = { grade: '', year: '', category: '' };
  var expandedId = null;

  function render(container, opts) {
    opts = opts || {};
    var questions = (ZhuCore.getTables() || {}).questions || [];
    var years = Array.from(new Set(questions.map(function (q) { return String(q['年度']); }))).sort();
    container.innerHTML = '';

    var heading = document.createElement('div');
    heading.className = 'prep-heading';
    heading.innerHTML = '<strong>備課台</strong><span>先篩題目，再點考點字進入字主板。</span>';
    container.appendChild(heading);

    var filterBar = document.createElement('div');
    filterBar.className = 'panel prep-filters';
    filterBar.appendChild(makeSelect('年級（全部）', GRADES, currentFilters.grade, function (v) { currentFilters.grade = v; renderList(); }));
    filterBar.appendChild(makeSelect('年度（全部）', years, currentFilters.year, function (v) { currentFilters.year = v; renderList(); }));
    filterBar.appendChild(makeSelect('類別（全部）', CATEGORIES, currentFilters.category, function (v) { currentFilters.category = v; renderList(); }));
    container.appendChild(filterBar);

    var listPanel = document.createElement('div');
    listPanel.className = 'panel';
    listPanel.id = 'prepList';
    container.appendChild(listPanel);

    function renderList() {
      var filtered = filterQuestions(questions, {
        grade: currentFilters.grade || null,
        year: currentFilters.year || null,
        category: currentFilters.category || null,
      });
      listPanel.innerHTML = '';

      var count = document.createElement('div');
      count.className = 'prep-count';
      count.textContent = '共 ' + filtered.length + ' 題';
      listPanel.appendChild(count);

      if (!filtered.length) {
        var empty = document.createElement('p');
        empty.textContent = '（沒有符合篩選條件的題目）';
        listPanel.appendChild(empty);
        return;
      }

      filtered.forEach(function (q) {
        var id = String(q['題目ID']);
        var row = document.createElement('div');
        row.className = 'qcard' + (expandedId === id ? ' active' : '');
        row.setAttribute('tabindex', '0');
        var meta = document.createElement('div');
        meta.className = 'qmeta';
        meta.textContent = q['年級'] + '　' + q['年度'] + ' 年　第 ' + q['題號'] + ' 題　' + q['類別'];
        var body = document.createElement('div');
        body.textContent = q['完整題目與選項'] || '';
        row.appendChild(meta);
        row.appendChild(body);

        if (expandedId === id) {
          var chars = charsOfQuestion(ZhuCore.getIndex(), id);
          var charBox = document.createElement('div');
          charBox.className = 'prep-chars';
          chars.forEach(function (char) {
            var b = document.createElement('button');
            b.className = 'chip prep-char';
            b.textContent = char;
            b.onclick = function (ev) {
              ev.stopPropagation();
              if (opts.onPickChar) opts.onPickChar(char, id);
            };
            charBox.appendChild(b);
          });
          if (!chars.length) charBox.textContent = '（這題抓不到考點字）';
          row.appendChild(charBox);
        }

        function toggle() {
          expandedId = expandedId === id ? null : id;
          renderList();
        }
        row.onclick = toggle;
        row.onkeydown = function (ev) {
          if (ev.key === 'Enter' || ev.key === ' ') { ev.preventDefault(); toggle(); }
        };
        listPanel.appendChild(row);
      });
    }

    renderList();
  }

  function makeSelect(placeholder, options, current, onChange) {
    var sel = document.createElement('select');
    sel.setAttribute('aria-label', placeholder);
    var blank = document.createElement('option');
    blank.value = '';
    blank.textContent = placeholder;
    sel.appendChild(blank);
    options.forEach(function (v) {
      var o = document.createElement('option');
      o.value = v;
      o.textContent = v;
      sel.appendChild(o);
    });
    sel.value = current || '';
    sel.onchange = function () { onChange(sel.value); };
    return sel;
  }

  var api = { filterQuestions: filterQuestions, charsOfQuestion: charsOfQuestion, render: render };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
