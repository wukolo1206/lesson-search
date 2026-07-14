// zhu-print.js — 詞籃整理、挖空模型、A4 分頁與列印預覽
var ZhuPrint = (function () {
  'use strict';

  function groupBasketByWord(basket) {
    var groups = [];
    var byWord = {};
    (basket || []).forEach(function (entry) {
      if (!entry || !entry.word) return;
      var group = byWord[entry.word];
      if (!group) {
        group = { word: entry.word, targets: [], entries: 0 };
        byWord[entry.word] = group;
        groups.push(group);
      }
      group.entries += 1;
      if (entry.char && group.targets.indexOf(entry.char) === -1) group.targets.push(entry.char);
    });
    return groups;
  }

  function lookupSyllables(bopo, word) {
    if (!bopo || typeof bopo.lookup !== 'function') return [];
    var result = bopo.lookup(word);
    return Array.isArray(result) ? result : [];
  }

  function buildPracticeCells(item, mode, selectedChars, bopo) {
    var word = item && item.word ? item.word : '';
    var selected = selectedChars || [];
    var syllables = lookupSyllables(bopo, word);
    var cells = [];
    for (var i = 0; i < word.length; i++) {
      var char = word[i];
      var isTarget = (item.targets || []).indexOf(char) !== -1 || selected.indexOf(char) !== -1;
      cells.push({
        char: char,
        blank: mode !== 'answer' && isTarget,
        bopomofo: syllables.length === word.length ? (syllables[i] || '') : '',
      });
    }
    return cells;
  }

  function createPage(columns) {
    var page = [];
    for (var i = 0; i < columns; i++) page.push([]);
    return page;
  }

  function paginatePracticeItems(items, options) {
    options = options || {};
    var columns = options.columns || 6;
    var rows = options.rows || 10;
    var pages = [];
    var page = createPage(columns);
    var columnIndex = columns - 1;
    var used = 0;

    (items || []).forEach(function (item) {
      var length = (item.word || '').length;
      if (!length) return;
      if (length > rows) return;
      if (used && used + length > rows) {
        columnIndex -= 1;
        used = 0;
      }
      if (columnIndex < 0) {
        pages.push({ columns: page });
        page = createPage(columns);
        columnIndex = columns - 1;
      }
      page[columnIndex].push(item);
      used += length;
    });

    var hasItems = page.some(function (column) { return column.length > 0; });
    if (hasItems) pages.push({ columns: page });
    return pages;
  }

  function buildPreviewState(options) {
    options = options || {};
    var mode = options.mode || 'fill';
    var grouped = groupBasketByWord(options.basket || []);
    var items = grouped.map(function (item) {
      var cells = buildPracticeCells(item, mode, options.selectedChars || [], options.bopo);
      return {
        word: item.word,
        targets: item.targets.slice(),
        entries: item.entries,
        cells: cells,
      };
    });
    var missingBopo = items.filter(function (item) {
      return item.cells.some(function (cell) { return !cell.bopomofo; });
    }).map(function (item) { return item.word; });
    var pages = paginatePracticeItems(items, options.pagination);
    return {
      mode: mode,
      items: items,
      pages: pages,
      pageCount: pages.length,
      missingBopo: missingBopo,
    };
  }

  function createElement(tag, className, text) {
    var el = document.createElement(tag);
    if (className) el.className = className;
    if (text !== undefined) el.textContent = text;
    return el;
  }

  function renderBopomofo(container, value) {
    container.innerHTML = '';
    if (!value) return;
    var light = value.charAt(0) === '˙';
    var toneMatch = light ? null : value.match(/[ˊˇˋ˙]$/);
    var sound = light ? value.slice(1) : (toneMatch ? value.slice(0, -1) : value);
    var tone = light ? '˙' : (toneMatch ? toneMatch[0] : '');
    var inner = createElement('span', 'practice-bopo-inner');
    var stack = createElement('span', 'practice-bopo-stack');
    Array.from(sound).forEach(function (symbol) {
      stack.appendChild(createElement('span', 'practice-bopo-symbol', symbol));
    });
    if (light) {
      inner.appendChild(createElement('span', 'practice-bopo-light', '˙'));
      inner.appendChild(stack);
    } else {
      inner.appendChild(stack);
      if (tone) inner.appendChild(createElement('span', 'practice-bopo-tone', tone));
    }
    container.setAttribute('aria-label', value);
    container.appendChild(inner);
  }

  function renderPracticeCell(cell) {
    var slot = createElement('div', 'practice-slot');
    var box = createElement('div', 'practice-box' + (cell.blank ? ' blank' : ''), cell.blank ? '' : cell.char);
    var bopo = createElement('div', 'practice-bopo' + (cell.bopomofo ? '' : ' empty'));
    renderBopomofo(bopo, cell.bopomofo);
    slot.appendChild(box);
    slot.appendChild(bopo);
    return slot;
  }

  function renderPracticePage(page, pageIndex, totalPages, meta) {
    var pageEl = createElement('section', 'practice-page');
    pageEl.style.setProperty('--page-num', pageIndex + 1);
    pageEl.style.setProperty('--total-pages', totalPages);

    var header = createElement('div', 'practice-header');
    header.appendChild(createElement('h1', 'practice-title', meta.title || '語詞練習'));
    var fields = createElement('div', 'practice-fields');
    fields.appendChild(createElement('span', '', '課別：'));
    fields.appendChild(createElement('span', 'practice-field-value', meta.lesson || ''));
    fields.appendChild(createElement('span', '', '班級：'));
    fields.appendChild(createElement('span', 'practice-field-value short', meta.className || ''));
    fields.appendChild(createElement('span', '', '姓名：'));
    fields.appendChild(createElement('span', 'practice-field-blank name', ''));
    fields.appendChild(createElement('span', '', '座號：'));
    fields.appendChild(createElement('span', 'practice-field-blank number', ''));
    header.appendChild(fields);
    pageEl.appendChild(header);

    var grid = createElement('div', 'practice-grid');
    page.columns.forEach(function (column) {
      var columnEl = createElement('div', 'practice-column');
      column.forEach(function (item) {
        var itemEl = createElement('div', 'practice-item');
        item.cells.forEach(function (cell) { itemEl.appendChild(renderPracticeCell(cell)); });
        columnEl.appendChild(itemEl);
      });
      grid.appendChild(columnEl);
    });
    pageEl.appendChild(grid);
    return pageEl;
  }

  function renderSheet(container, state, meta) {
    if (!container) return;
    container.innerHTML = '';
    (state.pages || []).forEach(function (page, i) {
      container.appendChild(renderPracticePage(page, i, state.pageCount, meta || {}));
    });
  }

  function openPreview(options) {
    options = options || {};
    var basket = options.basket || (window.ZhuCore ? ZhuCore.getBasket() : []);
    if (!basket.length) {
      alert('詞籃是空的，先在字主板或備課台加幾個詞。');
      return;
    }
    var overlay = document.getElementById('printPreview');
    var preview = document.getElementById('printPreviewPages');
    var printSheet = document.getElementById('printSheet');
    if (!overlay || !preview || !printSheet) return;

    var mode = options.mode || 'fill';
    var meta = { title: '語詞練習', lesson: '', className: '' };
    var selected = options.selectedChars || (window.ZhuCore ? ZhuCore.getSelectedChars().map(function (entry) { return entry.char; }) : []);
    var bopo = options.bopo || window.ZhuBopo || null;
    var modeSelect = document.getElementById('printMode');
    var lessonInput = document.getElementById('printLesson');
    var classInput = document.getElementById('printClass');
    var summary = document.getElementById('printSummary');
    var warning = document.getElementById('printWarning');
    if (modeSelect) modeSelect.value = mode;

    function render() {
      mode = modeSelect ? modeSelect.value : mode;
      meta.lesson = lessonInput ? lessonInput.value.trim() : '';
      meta.className = classInput ? classInput.value.trim() : '';
      var state = buildPreviewState({ basket: basket, selectedChars: selected, mode: mode, bopo: bopo });
      renderSheet(preview, state, meta);
      renderSheet(printSheet, state, meta);
      if (summary) summary.textContent = state.items.length + ' 詞／' + state.pageCount + ' 頁' + (mode === 'fill' ? '／填空版' : '／答案版');
      if (warning) {
        warning.textContent = state.missingBopo.length ? '有 ' + state.missingBopo.length + ' 個詞缺少逐字注音：' + state.missingBopo.join('、') : '逐字注音已載入';
        warning.classList.toggle('warning', state.missingBopo.length > 0);
      }
    }

    if (modeSelect) modeSelect.onchange = render;
    if (lessonInput) lessonInput.oninput = render;
    if (classInput) classInput.oninput = render;
    var printButton = document.getElementById('btnPrintPreview');
    if (printButton) printButton.onclick = function () { render(); window.print(); };
    var closeButton = document.getElementById('btnClosePrintPreview');
    if (closeButton) closeButton.onclick = function () { overlay.classList.add('hidden'); };
    render();
    overlay.classList.remove('hidden');
  }

  var api = {
    groupBasketByWord: groupBasketByWord,
    buildPracticeCells: buildPracticeCells,
    paginatePracticeItems: paginatePracticeItems,
    buildPreviewState: buildPreviewState,
    renderSheet: renderSheet,
    openPreview: openPreview,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
