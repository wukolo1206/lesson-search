// zhu-board.js — 字主板 UI ＋ boardState ＋ 詞籃
(function () {
  'use strict';

  ZhuCore.init(window.localStorage);
  var boardState = ZhuCore.state;
  var supExpanded = false;

  var $ = function (id) { return document.getElementById(id); };
  var index, radicals;

  function renderPrep() {
    if (!window.ZhuPrep) return;
    window.ZhuPrep.render($('view-prep'), {
      getSelectedChars: ZhuCore.getSelectedChars,
      onToggleChar: function (char, questionId) {
        var selected = ZhuCore.getSelectedChars().some(function (entry) { return entry.char === char; });
        if (selected) ZhuCore.removeSelectedChar(char);
        else ZhuCore.addSelectedChar(char, questionId);
        renderPrep();
      },
      onClearSelection: function () {
        ZhuCore.setSelectedChars([]);
        renderPrep();
      },
      onStart: function () {
        var selected = ZhuCore.getSelectedChars();
        if (!selected.length) return;
        var active = Math.min(ZhuCore.getActiveSelectedIndex(), selected.length - 1);
        ZhuCore.setActiveSelectedIndex(active);
        switchMode('board');
        showChar(selected[active].char, selected[active].contextQuestionId);
      },
    });
  }

  function switchMode(mode) {
    document.querySelectorAll('#modeTabs .tab').forEach(function (btn) {
      var active = btn.getAttribute('data-mode') === mode;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-selected', active ? 'true' : 'false');
    });
    $('view-prep').classList.toggle('hidden', mode !== 'prep');
    $('view-board').classList.toggle('hidden', mode !== 'board');
    $('view-projector').classList.toggle('hidden', mode !== 'projector');
    if (mode === 'prep') renderPrep();
    if (mode === 'projector' && window.ZhuProjector) {
      window.ZhuProjector.enter($('view-projector'), boardState, index, ZhuCore.getStore(), function () {
        alert('筆跡空間已滿，請先清除全部資料。');
      });
    }
    if (mode !== 'projector' && window.ZhuProjector) window.ZhuProjector.exit();
  }

  function enterBoardFromPrep(char, questionId) {
    switchMode('board');
    showChar(char, questionId);
  }

  document.querySelectorAll('#modeTabs .tab').forEach(function (btn) {
    btn.onclick = function () { switchMode(btn.getAttribute('data-mode')); };
  });

  // ── 載入 ────────────────────────────────────────────────
  function boot() {
    $('status').textContent = '載入題庫中…';
    $('status').classList.remove('hidden');
    $('app').classList.add('hidden');

    ZhuCore.boot(function () {
      index = ZhuCore.getIndex();
      radicals = ZhuCore.getRadicals();
      $('status').classList.add('hidden');
      $('app').classList.remove('hidden');
      $('basket').classList.remove('hidden');
      renderBasket();
      render();
    }, function (err) {
      $('status').innerHTML = '';
      var msg = document.createElement('p');
      msg.textContent = '載入失敗：' + err.message + '（請確認網路連線）';
      var retry = document.createElement('button');
      retry.className = 'btn btn-primary';
      retry.textContent = '重試';
      retry.onclick = boot;
      $('status').appendChild(msg);
      $('status').appendChild(retry);
    });
  }

  // ── 字主板渲染 ──────────────────────────────────────────
  function showChar(char, contextQuestionId) {
    boardState.char = char;
    boardState.contextQuestionId = contextQuestionId || null;
    var selected = ZhuCore.getSelectedChars();
    var selectedIndex = selected.findIndex(function (entry) { return entry.char === char; });
    if (selectedIndex >= 0) ZhuCore.setActiveSelectedIndex(selectedIndex);
    supExpanded = false;
    render();
  }

  function renderQueueNav() {
    var nav = $('queueNav');
    if (!nav) return;
    var selected = ZhuCore.getSelectedChars();
    nav.innerHTML = '';
    nav.classList.toggle('hidden', selected.length === 0);
    if (!selected.length) return;

    var active = Math.min(ZhuCore.getActiveSelectedIndex(), selected.length - 1);
    if (active !== ZhuCore.getActiveSelectedIndex()) ZhuCore.setActiveSelectedIndex(active);

    var heading = document.createElement('div');
    heading.className = 'queue-heading';
    heading.innerHTML = '<strong>逐字備課</strong><span>' + (active + 1) + '／' + selected.length + '</span>';
    nav.appendChild(heading);

    var chars = document.createElement('div');
    chars.className = 'queue-chars';
    selected.forEach(function (entry, i) {
      var button = document.createElement('button');
      button.className = 'queue-char' + (i === active ? ' active' : '');
      button.textContent = (i + 1) + '. ' + entry.char;
      button.setAttribute('aria-label', '第 ' + (i + 1) + ' 個考點字：' + entry.char);
      button.setAttribute('aria-current', i === active ? 'true' : 'false');
      button.onclick = function () {
        ZhuCore.setActiveSelectedIndex(i);
        showChar(entry.char, entry.contextQuestionId);
      };
      chars.appendChild(button);
    });
    nav.appendChild(chars);

    var controls = document.createElement('div');
    controls.className = 'queue-controls';
    var previous = document.createElement('button');
    previous.className = 'btn';
    previous.textContent = '← 上一字';
    previous.disabled = active === 0;
    previous.onclick = function () {
      if (active === 0) return;
      var next = active - 1;
      ZhuCore.setActiveSelectedIndex(next);
      showChar(selected[next].char, selected[next].contextQuestionId);
    };
    var nextButton = document.createElement('button');
    nextButton.className = 'btn btn-primary';
    nextButton.textContent = '下一字 →';
    nextButton.disabled = active >= selected.length - 1;
    nextButton.onclick = function () {
      if (active >= selected.length - 1) return;
      var next = active + 1;
      ZhuCore.setActiveSelectedIndex(next);
      showChar(selected[next].char, selected[next].contextQuestionId);
    };
    controls.appendChild(previous);
    controls.appendChild(nextButton);
    nav.appendChild(controls);
  }

  function render() {
    renderQueueNav();
    ZhuWrite.renderWriteBoard($('writeBoard'), ZhuCore.getBasket(), ZhuCore.getStore(), function () {
      alert('筆跡空間已滿，請先清除全部資料。');
    });
    if (!boardState.char) return;
    renderBigChar();
    renderQuestions();
    renderOptionWords();
    renderSupWords();
  }

  function renderBigChar() {
    $('bigChar').textContent = boardState.char;

    var bits = [];
    var rad = radicals.radicalOf(boardState.char);
    if (rad) bits.push(rad + '部');   // 衝突或抓不到 → 不顯示，不硬掰

    var bpmf = firstBopomofo(boardState.char);
    if (bpmf) bits.push(bpmf);

    $('charMeta').innerHTML = '';
    if (bits.length) {
      $('charMeta').appendChild(document.createTextNode(bits.join('　')));
    }

    // 以錯字進入時要標出來，並顯示對應正字
    var asWrong = wrongCharInfo(boardState.char);
    if (asWrong) {
      var b = document.createElement('span');
      b.className = 'badge-wrong';
      b.textContent = '✗ 常見錯字，正確是「' + asWrong + '」';
      $('charMeta').appendChild(document.createElement('br'));
      $('charMeta').appendChild(b);
    }
  }

  function firstBopomofo(char) {
    var w = index.wordsOf(char).filter(function (x) { return x.bopomofo; });
    return w.length ? w[0].bopomofo : '';
  }

  // 這個字是否曾以「錯字」身分出現？是的話回它對應的正字。
  // 正字直接讀索引帶過來的 correctChar —— 不要試圖從語詞回推，推不出來。
  function wrongCharInfo(char) {
    var hit = index.wordsOf(char).find(function (w) {
      return w.source === 'mistake' && w.wrongChar === char && w.correctChar;
    });
    return hit ? hit.correctChar : null;
  }

  function renderQuestions() {
    var box = $('questions');
    box.innerHTML = '';
    var qs = index.questionsOf(boardState.char);
    if (!qs.length) {
      box.textContent = '（題庫裡沒有這個字）';
      return;
    }
    qs.forEach(function (q) {
      var d = document.createElement('div');
      d.className = 'qcard' + (String(q['題目ID']) === String(boardState.contextQuestionId) ? ' active' : '');
      var meta = document.createElement('div');
      meta.className = 'qmeta';
      meta.textContent = q['年級'] + '　' + q['年度'] + ' 年　第 ' + q['題號'] + ' 題　' + q['類別'];
      var body = document.createElement('div');
      body.textContent = q['完整題目與選項'] || '';
      d.appendChild(meta);
      d.appendChild(body);
      d.onclick = function () {
        // 再點一次同一題＝取消教學情境
        var same = String(q['題目ID']) === String(boardState.contextQuestionId);
        showChar(boardState.char, same ? null : q['題目ID']);
      };
      box.appendChild(d);
    });
  }

  function renderOptionWords() {
    var words = ZhuData.optionWords(index, boardState);
    // 直接查字（沒有教學情境）時整區不顯示 —— 沒有「本題」這回事
    $('optionPanel').classList.toggle('hidden', words.length === 0);
    var box = $('optionWords');
    box.innerHTML = '';
    words.forEach(function (w) { box.appendChild(chipFor(w, false)); });
  }

  function renderSupWords() {
    var publisherWords = typeof ZhuPublisherWords !== 'undefined' ? ZhuPublisherWords.wordsOf(boardState.char) : [];
    var all = ZhuData.mergeSupplementWords(index, boardState, publisherWords);
    var limit = ZhuData.SUPPLEMENT_VISIBLE;
    var shown = supExpanded ? all : all.slice(0, limit);

    var box = $('supWords');
    box.innerHTML = '';
    if (!shown.length) box.textContent = '（沒有補充詞）';
    shown.forEach(function (w) { box.appendChild(chipFor(w, true)); });

    var btn = $('btnMore');
    var rest = all.length - limit;
    if (rest > 0 && !supExpanded) {
      btn.textContent = '還有 ' + rest + ' 個 ▾';
      btn.classList.remove('hidden');
      btn.onclick = function () { supExpanded = true; renderSupWords(); };
    } else {
      btn.classList.add('hidden');
    }
  }

  function sourceLabel(sources) {
    var groups = {};
    var publisherOrder = ['康軒', '翰林', '南一'];
    (sources || []).forEach(function (source) {
      if (!source || !source.publisher || !source.volume) return;
      if (!groups[source.publisher]) groups[source.publisher] = {};
      if (!groups[source.publisher][source.volume]) groups[source.publisher][source.volume] = [];
      if (source.source && groups[source.publisher][source.volume].indexOf(source.source) === -1) {
        groups[source.publisher][source.volume].push(source.source);
      }
    });
    return publisherOrder.filter(function (publisher) { return groups[publisher]; }).map(function (publisher) {
      var volumes = Object.keys(groups[publisher]).map(function (volume) {
        var kinds = groups[publisher][volume];
        return volume + (kinds.length ? '（' + kinds.join('、') + '）' : '');
      });
      return publisher + volumes.join('、');
    }).join('｜');
  }

  function displayBopomofo(entry) {
    if (entry.bopomofo) return entry.bopomofo;
    if (typeof ZhuBopo === 'undefined' || !ZhuBopo.lookup) return '';
    var syllables = ZhuBopo.lookup(entry.word);
    return syllables.length === entry.word.length ? syllables.join(' ') : '';
  }

  function chipFor(entry, isSup) {
    var c = document.createElement('span');
    c.className = 'chip' + (isSup ? ' sup' : '');
    var inBasket = ZhuCore.isInBasket(entry.word, entry.char);
    if (inBasket) c.classList.add('in-basket');
    var main = document.createElement('span');
    main.className = 'chip-main';
    var bopomofo = displayBopomofo(entry);
    main.textContent = entry.word + (bopomofo ? '（' + bopomofo + '）' : '');
    var action = document.createElement('span');
    action.className = 'chip-action';
    action.textContent = inBasket ? '✓' : '＋';
    main.appendChild(action);
    c.appendChild(main);

    var sourceText = sourceLabel(entry.publisherSources);
    if (sourceText) {
      var source = document.createElement('small');
      source.className = 'publisher-source';
      source.textContent = sourceText;
      c.appendChild(source);
    }
    var sourceTitle = (entry.publisherSources || []).map(function (item) {
      return item.publisher + item.volume + (item.source ? '（' + item.source + '）' : '');
    }).join('、');
    c.title = [entry.gloss, sourceTitle].filter(Boolean).join('\n');
    c.setAttribute('aria-label', entry.word + (sourceText ? '，來源：' + sourceText : ''));
    c.onclick = function () {
      if (inBasket) { ZhuCore.removeFromBasket(entry.word, entry.char); }
      else if (!ZhuCore.addToBasket(entry)) { alert('儲存空間已滿，詞籃沒存進去。請按「清除全部資料」後再試。'); }
      render();
      renderBasket();
    };
    return c;
  }

  // ── 詞籃 ────────────────────────────────────────────────
  function renderBasket() {
    var basket = ZhuCore.getBasket();
    $('basketCount').textContent = basket.length;
    var box = $('basketItems');
    box.innerHTML = '';
    ZhuData.basketOps.sortByGrade(basket).forEach(function (item) {
      var s = document.createElement('span');
      s.className = 'chip';
      s.textContent = item.word + ' ×';
      s.onclick = function () {
        ZhuCore.removeFromBasket(item.word, item.char);
        render();
        renderBasket();
      };
      box.appendChild(s);
    });
  }

  // ── 事件 ────────────────────────────────────────────────
  $('btnSearch').onclick = function () {
    var c = $('kw').value.trim();
    if (c) showChar(c, null);   // 從搜尋進入 → 沒有教學情境
  };
  $('kw').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') $('btnSearch').click();
  });
  $('grade').value = boardState.grade;
  $('grade').onchange = function () {
    ZhuCore.setGrade(this.value);
    render();
  };
  $('btnClear').onclick = function () {
    if (!confirm('詞籃、筆跡、偏好都會清掉，確定嗎？')) return;
    ZhuCore.clearAll();
    renderBasket();
    render();
  };
  $('btnPrint').onclick = function () {
    if (window.ZhuPrint) ZhuPrint.openPreview({ basket: ZhuCore.getBasket() });
    else ZhuWrite.printSheet(ZhuCore.getBasket());
  };

  boot();
})();
