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

  function charsFromQuestions(index, questions) {
    var ids = {};
    (questions || []).forEach(function (q) { ids[String(q['題目ID'])] = true; });
    return index.allChars().map(function (char) {
      var matched = index.questionsOf(char).filter(function (q) { return ids[String(q['題目ID'])]; });
      return { char: char, questionCount: matched.length };
    }).filter(function (entry) { return entry.questionCount > 0; });
  }

  function addSelectedChar(queue, char, contextQuestionId) {
    if (!char || (queue || []).some(function (entry) { return entry.char === char; })) return (queue || []).slice();
    return (queue || []).concat([{
      char: char,
      contextQuestionId: contextQuestionId === null || contextQuestionId === undefined ? null : String(contextQuestionId),
    }]);
  }

  function removeSelectedChar(queue, char) {
    return (queue || []).filter(function (entry) { return entry.char !== char; });
  }

  // ── 以下是 DOM，Node 測試不會走到 ──────────────────────
  var GRADES = ['三年級', '四年級', '五年級', '六年級'];
  var CATEGORIES = ['同音字辨析', '常見錯別字', '形近字辨析', '成語', '多音字', '詞義理解與應用', '部首辨識'];
  var currentFilters = { grade: '', year: '', category: '' };
  var expandedId = null;
  var viewMode = 'questions';

  function render(container, opts) {
    opts = opts || {};
    var questions = (ZhuCore.getTables() || {}).questions || [];
    var years = Array.from(new Set(questions.map(function (q) { return String(q['年度']); }))).sort();
    container.innerHTML = '';

    var heading = document.createElement('div');
    heading.className = 'prep-heading';
    heading.innerHTML = '<strong>備課台</strong><span>先篩題目，再多選考點字，最後逐字進入字主板。</span>';
    container.appendChild(heading);

    var viewTabs = document.createElement('div');
    viewTabs.className = 'prep-view-tabs';
    [['questions', '按考題'], ['chars', '類別字庫']].forEach(function (item) {
      var tab = document.createElement('button');
      tab.className = 'btn' + (viewMode === item[0] ? ' btn-primary' : '');
      tab.textContent = item[1];
      tab.setAttribute('aria-pressed', viewMode === item[0] ? 'true' : 'false');
      tab.onclick = function () { viewMode = item[0]; render(container, opts); };
      viewTabs.appendChild(tab);
    });
    container.appendChild(viewTabs);

    var filterBar = document.createElement('div');
    filterBar.className = 'panel prep-filters';
    filterBar.appendChild(makeSelect('年級（全部）', GRADES, currentFilters.grade, function (v) { currentFilters.grade = v; render(container, opts); }));
    filterBar.appendChild(makeSelect('年度（全部）', years, currentFilters.year, function (v) { currentFilters.year = v; render(container, opts); }));
    filterBar.appendChild(makeSelect('類別（全部）', CATEGORIES, currentFilters.category, function (v) { currentFilters.category = v; render(container, opts); }));
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

      if (viewMode === 'chars') {
        var charGrid = document.createElement('div');
        charGrid.className = 'prep-char-grid';
        charsFromQuestions(ZhuCore.getIndex(), filtered).forEach(function (entry) {
          var selected = isSelected(entry.char);
          var taught = isTaught(entry.char);
          var card = document.createElement('div');
          card.className = 'prep-char-card' + (selected ? ' selected' : '') + (taught ? ' taught' : '');

          var selectButton = document.createElement('button');
          selectButton.className = 'prep-char-select';
          selectButton.setAttribute('aria-label', '選取字：' + entry.char);
          selectButton.setAttribute('aria-pressed', selected ? 'true' : 'false');
          selectButton.innerHTML = '<strong>' + entry.char + '</strong><small>' + entry.questionCount + ' 題</small>';
          selectButton.onclick = function () { toggleSelected(entry.char, null); };
          card.appendChild(selectButton);

          card.appendChild(makeTaughtButton(entry.char, taught));
          charGrid.appendChild(card);
        });
        listPanel.appendChild(charGrid);
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
            var wrap = document.createElement('span');
            var taught = isTaught(char);
            wrap.className = 'prep-char-wrap' + (taught ? ' taught' : '');
            var b = document.createElement('button');
            var selected = isSelected(char);
            b.className = 'chip prep-char' + (selected ? ' selected' : '') + (taught ? ' taught' : '');
            b.textContent = char + (selected ? ' ✓' : ' ＋');
            b.setAttribute('aria-pressed', selected ? 'true' : 'false');
            b.onclick = function (ev) {
              ev.stopPropagation();
              toggleSelected(char, id);
            };
            wrap.appendChild(b);
            wrap.appendChild(makeTaughtButton(char, taught));
            charBox.appendChild(wrap);
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

    var selectedPanel = document.createElement('div');
    selectedPanel.className = 'panel prep-selection';
    var selectedHeading = document.createElement('div');
    selectedHeading.className = 'prep-selection-heading';
    selectedHeading.innerHTML = '<strong>已選考點字</strong><span>' + getSelected().length + ' 字，依序逐字備課</span>';
    selectedPanel.appendChild(selectedHeading);

    var selectedBox = document.createElement('div');
    selectedBox.className = 'prep-selection-chars';
    getSelected().forEach(function (entry, i) {
      var chip = document.createElement('button');
      chip.className = 'chip prep-selected-chip';
      chip.textContent = (i + 1) + '. ' + entry.char + ' ×';
      chip.title = '移除「' + entry.char + '」';
      chip.onclick = function () {
        if (opts.onToggleChar) opts.onToggleChar(entry.char, null);
      };
      selectedBox.appendChild(chip);
    });
    if (!getSelected().length) {
      var selectedEmpty = document.createElement('span');
      selectedEmpty.className = 'prep-selection-empty';
      selectedEmpty.textContent = '尚未選字，請從上方題目或類別字庫點選。';
      selectedBox.appendChild(selectedEmpty);
    }
    selectedPanel.appendChild(selectedBox);

    var taughtProgress = document.createElement('div');
    taughtProgress.className = 'prep-taught-progress';
    var taughtCount = opts.getTaughtChars ? opts.getTaughtChars().length : 0;
    var taughtSummary = document.createElement('span');
    taughtSummary.textContent = '已教 ' + taughtCount + ' 字';
    taughtProgress.appendChild(taughtSummary);
    var resetTaughtButton = document.createElement('button');
    resetTaughtButton.className = 'btn';
    resetTaughtButton.textContent = '重設已教紀錄';
    resetTaughtButton.disabled = taughtCount === 0;
    resetTaughtButton.onclick = function () {
      if (opts.onResetTaught) opts.onResetTaught();
    };
    taughtProgress.appendChild(resetTaughtButton);
    selectedPanel.appendChild(taughtProgress);

    var selectedActions = document.createElement('div');
    selectedActions.className = 'prep-selection-actions';
    var clearButton = document.createElement('button');
    clearButton.className = 'btn';
    clearButton.textContent = '清除選字';
    clearButton.disabled = !getSelected().length;
    clearButton.onclick = function () { if (opts.onClearSelection) opts.onClearSelection(); };
    selectedActions.appendChild(clearButton);
    var startButton = document.createElement('button');
    startButton.className = 'btn btn-primary';
    startButton.textContent = '開始逐字備課 →';
    startButton.disabled = !getSelected().length;
    startButton.onclick = function () { if (opts.onStart) opts.onStart(); };
    selectedActions.appendChild(startButton);
    selectedPanel.appendChild(selectedActions);
    container.appendChild(selectedPanel);

    function getSelected() {
      return opts.getSelectedChars ? opts.getSelectedChars() : [];
    }

    function isSelected(char) {
      return getSelected().some(function (entry) { return entry.char === char; });
    }

    function isTaught(char) {
      return opts.isTaughtChar ? opts.isTaughtChar(char) : false;
    }

    function makeTaughtButton(char, taught) {
      var button = document.createElement('button');
      button.className = 'taught-action' + (taught ? ' is-taught' : '');
      button.textContent = taught ? '✓ 已教' : '標記已教';
      button.setAttribute('aria-label', (taught ? '取消已教：' : '標記已教：') + char);
      button.setAttribute('aria-pressed', taught ? 'true' : 'false');
      button.onclick = function (ev) {
        ev.preventDefault();
        ev.stopPropagation();
        if (opts.onToggleTaught) opts.onToggleTaught(char);
      };
      return button;
    }

    function toggleSelected(char, questionId) {
      if (opts.onToggleChar) opts.onToggleChar(char, questionId);
    }
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

  var api = {
    filterQuestions: filterQuestions,
    charsOfQuestion: charsOfQuestion,
    charsFromQuestions: charsFromQuestions,
    addSelectedChar: addSelectedChar,
    removeSelectedChar: removeSelectedChar,
    render: render,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
