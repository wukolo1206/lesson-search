// zhu-projector.js — 投影：步驟建構＋全螢幕推進
var ZhuProjector = (function () {
  'use strict';

  function buildSteps(question, optionWords, supplementWords) {
    optionWords = optionWords || [];
    supplementWords = supplementWords || [];
    var explain = question ? [question['記憶訣竅'], question['教學策略']].filter(Boolean).join('\n\n') : '';
    return [
      { type: 'question', label: '題目', text: question ? (question['完整題目與選項'] || '') : '（沒有教學情境）' },
      { type: 'explain', label: '解析', text: explain },
      { type: 'words', label: '選項詞', words: optionWords },
      { type: 'words', label: '補充詞', words: supplementWords },
      { type: 'write', label: '手寫格', words: optionWords.concat(supplementWords) },
    ];
  }

  // ── 以下是 DOM，Node 測試不會走到 ──────────────────────
  var stepIndex = 0;
  var keyHandler = null;

  function enter(container, boardState, index) {
    exit();
    stepIndex = 0;
    render(container, boardState, index);
    keyHandler = function (e) {
      if (e.key === 'ArrowRight') {
        e.preventDefault();
        move(1, container, boardState, index);
      }
      if (e.key === 'ArrowLeft') {
        e.preventDefault();
        move(-1, container, boardState, index);
      }
    };
    document.addEventListener('keydown', keyHandler);
  }

  function exit() {
    if (keyHandler) document.removeEventListener('keydown', keyHandler);
    keyHandler = null;
  }

  function move(delta, container, boardState, index) {
    stepIndex = Math.max(0, Math.min(4, stepIndex + delta));
    render(container, boardState, index);
  }

  function render(container, boardState, index) {
    container.innerHTML = '';
    if (!boardState.char || !index) {
      var empty = document.createElement('p');
      empty.className = 'projempty';
      empty.textContent = '先在字主板或備課台選一個字。';
      container.appendChild(empty);
      return;
    }

    var question = null;
    if (boardState.contextQuestionId) {
      question = index.questionsOf(boardState.char).find(function (q) {
        return String(q['題目ID']) === String(boardState.contextQuestionId);
      }) || null;
    }
    var optionWords = ZhuData.optionWords(index, boardState);
    var supplementWords = ZhuData.supplementWords(index, boardState);
    var steps = buildSteps(question, optionWords, supplementWords);
    var step = steps[stepIndex];

    var big = document.createElement('div');
    big.className = 'projchar';
    big.textContent = boardState.char;
    container.appendChild(big);

    var label = document.createElement('div');
    label.className = 'projlabel';
    label.textContent = step.label + '（' + (stepIndex + 1) + '/5，← → 推進）';
    container.appendChild(label);

    var body = document.createElement('div');
    body.className = 'projbody';
    if (step.type === 'question') {
      body.textContent = step.text || '（沒有內容）';
    } else if (step.type === 'explain') {
      body.classList.add('projexplain');
      var parts = step.text ? step.text.split('\n\n') : [];
      if (!parts.length) {
        body.textContent = '（沒有內容）';
      } else {
        parts.forEach(function (part) {
          var p = document.createElement('p');
          p.className = 'projexplain-part';
          p.textContent = part.replace(/(?<!^)([0-9]+\.\s?)/g, '\n$1');
          body.appendChild(p);
        });
      }
    } else if (step.type === 'words') {
      step.words.forEach(function (w) {
        var chip = document.createElement('span');
        chip.className = 'chip';
        chip.textContent = w.word + (w.bopomofo ? '（' + w.bopomofo + '）' : '');
        body.appendChild(chip);
      });
      if (!step.words.length) body.textContent = '（沒有）';
    } else if (step.type === 'write') {
      step.words.forEach(function (w) {
        var row = document.createElement('div');
        row.className = 'projwrite';
        ZhuWrite.buildCells(w.word).forEach(function (cell, i) {
          var d = document.createElement('span');
          d.className = 'projcell ' + cell.style;
          if (i > 0 && i % w.word.length === 0) d.classList.add('projgap');
          d.textContent = cell.char || '　';
          row.appendChild(d);
        });
        body.appendChild(row);
      });
      if (!step.words.length) body.textContent = '（沒有可練習的詞）';
    }
    container.appendChild(body);

    var controls = document.createElement('div');
    controls.className = 'projcontrols';
    controls.appendChild(projectorButton('← 上一步', stepIndex === 0, function () { move(-1, container, boardState, index); }));
    controls.appendChild(projectorButton('下一步 →', stepIndex === 4, function () { move(1, container, boardState, index); }));
    container.appendChild(controls);
  }

  function projectorButton(label, disabled, onClick) {
    var button = document.createElement('button');
    button.className = 'projbtn';
    button.textContent = label;
    button.disabled = disabled;
    button.onclick = onClick;
    return button;
  }

  var api = { buildSteps: buildSteps, enter: enter, exit: exit };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  return api;
})();
