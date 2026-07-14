// zhu-write.js — 手寫格元件（描紅格）＋ 列印練習紙
var ZhuWrite = (function () {
  'use strict';

  // 描紅格三段：描紅 → 淡描 → 空白
  function buildCells(word) {
    if (!word) return [];
    var chars = word.split('');
    var out = [];
    ['trace', 'faint'].forEach(function (style) {
      chars.forEach(function (c) { out.push({ char: c, style: style }); });
    });
    chars.forEach(function () { out.push({ char: '', style: 'blank' }); });
    return out;
  }

  function inkKeyFor(word, char, cellIndex) {
    return word + '@' + char + '#' + cellIndex;
  }

  function loadInk(store, word, char, cellIndex) {
    var all = store.get('ink', {});
    var value = all[inkKeyFor(word, char, cellIndex)];
    return value === undefined ? null : value;
  }

  function saveInk(store, word, char, cellIndex, dataUrl) {
    var all = store.get('ink', {});
    all[inkKeyFor(word, char, cellIndex)] = dataUrl;
    return store.set('ink', all);
  }

  function clearInkForWord(store, word, char) {
    var all = store.get('ink', {});
    var prefix = word + '@' + char + '#';
    Object.keys(all).forEach(function (key) {
      if (key.indexOf(prefix) === 0) delete all[key];
    });
    return store.set('ink', all);
  }

  // ── 以下是 DOM，Node 測試不會走到 ──────────────────────
  function renderRow(item) {
    var row = document.createElement('div');
    row.className = 'wrow';

    var label = document.createElement('div');
    label.className = 'wlabel';
    label.textContent = item.word;
    row.appendChild(label);

    var cells = document.createElement('div');
    cells.className = 'wcells';
    buildCells(item.word).forEach(function (cell, i) {
      var d = document.createElement('div');
      d.className = 'wcell ' + cell.style;
      // 每一段之間留空隙
      if (i > 0 && i % item.word.length === 0) d.classList.add('wgap');
      d.textContent = cell.char;
      cells.appendChild(d);
    });
    row.appendChild(cells);
    return row;
  }

  var CELL_PX = 96;
  var MODAL_PX = 640;

  function openWriteModal(item, cellIndex, cell, previewCanvas, store, onQuotaFull) {
    var backdrop = document.createElement('div');
    backdrop.className = 'write-modal-backdrop';

    var card = document.createElement('div');
    card.className = 'write-modal-card';
    backdrop.appendChild(card);

    var title = document.createElement('div');
    title.className = 'write-modal-title';
    title.textContent = item.word + '（第 ' + (cellIndex + 1) + ' 格）';
    card.appendChild(title);

    var wrap = document.createElement('div');
    wrap.className = 'write-modal-canvas-wrap ' + cell.style;
    card.appendChild(wrap);

    var guide = document.createElement('span');
    guide.className = 'wguide';
    guide.textContent = cell.char;
    wrap.appendChild(guide);

    var canvas = document.createElement('canvas');
    canvas.width = MODAL_PX;
    canvas.height = MODAL_PX;
    canvas.className = 'write-modal-canvas';
    wrap.appendChild(canvas);

    var ctx = canvas.getContext('2d');
    var saved = loadInk(store, item.word, item.char, cellIndex);
    if (saved) {
      var img = new Image();
      img.onload = function () { ctx.drawImage(img, 0, 0, canvas.width, canvas.height); };
      img.src = saved;
    }

    var drawing = false;
    function pos(e) {
      var rect = canvas.getBoundingClientRect();
      return {
        x: (e.clientX - rect.left) * (canvas.width / rect.width),
        y: (e.clientY - rect.top) * (canvas.height / rect.height),
      };
    }
    canvas.addEventListener('pointerdown', function (e) {
      e.preventDefault();
      drawing = true;
      if (canvas.setPointerCapture) canvas.setPointerCapture(e.pointerId);
      var point = pos(e);
      ctx.beginPath();
      ctx.moveTo(point.x, point.y);
    });
    canvas.addEventListener('pointermove', function (e) {
      if (!drawing) return;
      e.preventDefault();
      var point = pos(e);
      ctx.lineTo(point.x, point.y);
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.lineWidth = 8;
      ctx.strokeStyle = '#1e293b';
      ctx.stroke();
    });
    function finishStroke(e) {
      if (!drawing) return;
      e.preventDefault();
      drawing = false;
    }
    canvas.addEventListener('pointerup', finishStroke);
    canvas.addEventListener('pointercancel', finishStroke);

    function commitAndClose() {
      var previewCtx = previewCanvas.getContext('2d');
      previewCtx.clearRect(0, 0, previewCanvas.width, previewCanvas.height);
      previewCtx.drawImage(canvas, 0, 0, previewCanvas.width, previewCanvas.height);
      var ok = saveInk(store, item.word, item.char, cellIndex, previewCanvas.toDataURL('image/webp', 0.6));
      if (!ok && onQuotaFull) onQuotaFull();
      document.body.removeChild(backdrop);
      document.removeEventListener('keydown', onKeydown);
    }

    var actions = document.createElement('div');
    actions.className = 'write-modal-actions';
    card.appendChild(actions);

    var clearBtn = document.createElement('button');
    clearBtn.className = 'btn write-clear';
    clearBtn.textContent = '清除';
    clearBtn.onclick = function () { ctx.clearRect(0, 0, canvas.width, canvas.height); };
    actions.appendChild(clearBtn);

    var doneBtn = document.createElement('button');
    doneBtn.className = 'btn';
    doneBtn.textContent = '完成';
    doneBtn.onclick = commitAndClose;
    actions.appendChild(doneBtn);

    function onKeydown(e) {
      if (e.key === 'Escape') commitAndClose();
    }
    document.addEventListener('keydown', onKeydown);

    document.body.appendChild(backdrop);
  }

  function renderCanvasRow(item, store, onQuotaFull) {
    var row = document.createElement('div');
    row.className = 'wrow writable';

    var heading = document.createElement('div');
    heading.className = 'write-row-heading';
    var label = document.createElement('div');
    label.className = 'wlabel';
    label.textContent = item.word;
    heading.appendChild(label);

    var clearBtn = document.createElement('button');
    clearBtn.className = 'btn write-clear';
    clearBtn.textContent = '清除本列';
    heading.appendChild(clearBtn);
    row.appendChild(heading);

    var cellsBox = document.createElement('div');
    cellsBox.className = 'wcells screen-cells';
    row.appendChild(cellsBox);

    var cells = buildCells(item.word);
    var canvases = cells.map(function (cell, i) {
      var wrap = document.createElement('div');
      wrap.className = 'wcanvasWrap ' + cell.style;
      if (i > 0 && i % item.word.length === 0) wrap.classList.add('wgap');

      var guide = document.createElement('span');
      guide.className = 'wguide';
      guide.textContent = cell.char;
      wrap.appendChild(guide);

      var canvas = document.createElement('canvas');
      canvas.width = CELL_PX;
      canvas.height = CELL_PX;
      canvas.className = 'wcanvas';
      canvas.setAttribute('aria-label', item.word + ' 第 ' + (i + 1) + ' 格手寫區');
      wrap.appendChild(canvas);
      cellsBox.appendChild(wrap);

      var ctx = canvas.getContext('2d');
      var saved = loadInk(store, item.word, item.char, i);
      if (saved) {
        var img = new Image();
        img.onload = function () { ctx.drawImage(img, 0, 0); };
        img.src = saved;
      }

      wrap.addEventListener('click', function () {
        openWriteModal(item, i, cell, canvas, store, onQuotaFull);
      });
      return canvas;
    });

    clearBtn.onclick = function () {
      clearInkForWord(store, item.word, item.char);
      canvases.forEach(function (canvas) {
        canvas.getContext('2d').clearRect(0, 0, canvas.width, canvas.height);
      });
    };

    return row;
  }

  function renderProjectorWriteRow(item, store, onQuotaFull) {
    var row = document.createElement('div');
    row.className = 'projwrite';

    buildCells(item.word).forEach(function (cell, i) {
      var wrap = document.createElement('div');
      wrap.className = 'projcell-wrap ' + cell.style;
      if (i > 0 && i % item.word.length === 0) wrap.classList.add('projgap');

      var guide = document.createElement('span');
      guide.className = 'projcell-guide';
      guide.textContent = cell.char || '';
      wrap.appendChild(guide);

      var canvas = document.createElement('canvas');
      canvas.width = CELL_PX;
      canvas.height = CELL_PX;
      canvas.className = 'projcell-canvas';
      wrap.appendChild(canvas);

      var ctx = canvas.getContext('2d');
      var saved = loadInk(store, item.word, item.char, i);
      if (saved) {
        var img = new Image();
        img.onload = function () { ctx.drawImage(img, 0, 0); };
        img.src = saved;
      }

      wrap.addEventListener('click', function () {
        openWriteModal(item, i, cell, canvas, store, onQuotaFull);
      });
      row.appendChild(wrap);
    });

    return row;
  }

  function renderWriteBoard(container, basket, store, onQuotaFull) {
    container.innerHTML = '';
    if (!basket.length) {
      container.className = 'write-empty';
      container.textContent = '詞籃是空的，先在字主板或備課台加幾個詞。';
      return;
    }
    container.className = '';
    ZhuData.basketOps.sortByGrade(basket).forEach(function (item) {
      container.appendChild(renderCanvasRow(item, store, onQuotaFull));
    });
  }

  function printSheet(basket) {
    if (!basket || !basket.length) {
      alert('詞籃是空的，先點幾個語詞丟進來。');
      return;
    }
    var sheet = document.getElementById('printSheet');
    sheet.innerHTML = '';

    var h = document.createElement('h1');
    h.className = 'wtitle';
    h.textContent = '語詞練習';
    sheet.appendChild(h);

    ZhuData.basketOps.sortByGrade(basket).forEach(function (item) {
      sheet.appendChild(renderRow(item));
    });

    window.print();
  }

  var api = {
    buildCells: buildCells,
    printSheet: printSheet,
    inkKeyFor: inkKeyFor,
    loadInk: loadInk,
    saveInk: saveInk,
    clearInkForWord: clearInkForWord,
    renderWriteBoard: renderWriteBoard,
    renderProjectorWriteRow: renderProjectorWriteRow,
    openWriteModal: openWriteModal,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
