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

  var api = { buildCells: buildCells, printSheet: printSheet };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
