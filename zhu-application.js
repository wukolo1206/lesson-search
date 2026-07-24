var ZhuApplication = (function () {
  'use strict';

  var GAS_URL = 'https:/' + '/script.google.com/macros/s/AKfycbwHoHUCmxwsZyMl7TZdFVOMAYZZixJIp_JaMfsQQOpRDzHklArqfmX2nAiFW2NQwVrN/exec';
  var currentBasket = [];
  var currentResult = null;
  var currentVersion = 'teacher';
  var running = false;
  var dictionaryLoaded = false;
  var lastFocus = null;
  var requestSeq = 0;
  var $ = function (id) { return document.getElementById(id); };

  function selectedTerms() {
    return Array.from(document.querySelectorAll('#waTermChoices input:checked'))
      .map(function (input) { return input.value; });
  }

  function config() {
    return {
      terms: selectedTerms(),
      level: $('waLevel').value,
      mode: $('waMode').value
    };
  }

  function updateGenerateState() {
    var error = ZhuApplicationCore.validateConfig(config());
    $('waTermError').textContent = error || '';
    $('waGenerate').disabled = running || Boolean(error);
    $('waRegenerate').disabled = running || Boolean(error);
  }

  function renderTermChoices() {
    var box = $('waTermChoices');
    box.textContent = '';
    ZhuApplicationCore.normalizeTerms(currentBasket).forEach(function (term, index) {
      var label = document.createElement('label');
      label.className = 'wa-term-choice';
      var input = document.createElement('input');
      input.type = 'checkbox';
      input.value = term;
      input.checked = true;
      input.id = 'waTerm' + index;
      input.addEventListener('change', function () {
        currentResult = null;
        $('waResult').classList.add('hidden');
        updateGenerateState();
      });
      label.appendChild(input);
      label.appendChild(document.createTextNode(term));
      box.appendChild(label);
    });
  }

  function open(basket, trigger) {
    requestSeq += 1;
    currentBasket = Array.isArray(basket) ? basket.slice() : [];
    lastFocus = trigger || document.activeElement;
    currentVersion = 'teacher';
    dictionaryLoaded = false;
    renderTermChoices();
    $('waDictionary').open = false;
    $('waDictionaryBody').textContent = '';
    $('waStatus').textContent = '';
    $('waOverlay').classList.remove('hidden');
    $('waOverlay').setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';
    restoreStored();
    updateGenerateState();
    $('waTitle').focus();
  }

  function close() {
    $('waOverlay').classList.add('hidden');
    $('waOverlay').setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
    if (lastFocus && lastFocus.focus) lastFocus.focus();
  }

  function restoreStored() {
    var stored = ZhuCore.getWordApplicationResult();
    if (ZhuApplicationCore.matchesStored(stored, config())) {
      currentResult = stored;
      renderResult();
    } else {
      currentResult = null;
      $('waResult').classList.add('hidden');
    }
  }

  function renderDictionaryRow(data, term) {
    var row = document.createElement('section');
    row.className = 'wa-dictionary-row';
    var heading = document.createElement('h3');
    heading.textContent = term;
    row.appendChild(heading);
    if (!data) {
      row.appendChild(document.createTextNode('萌典找不到這個詞'));
      return row;
    }
    var reading = document.createElement('p');
    reading.textContent = data.zhuyin || '未提供注音';
    var explanation = document.createElement('p');
    explanation.textContent = data.explanation || '未提供詞義';
    row.appendChild(reading);
    row.appendChild(explanation);
    return row;
  }

  function loadDictionary() {
    if (dictionaryLoaded) return;
    dictionaryLoaded = true;
    var terms = selectedTerms();
    var body = $('waDictionaryBody');
    body.textContent = '查詢萌典中…';
    Promise.all(terms.map(function (term) {
      return fetch('https:/' + '/www.moedict.tw/uni/' + encodeURIComponent(term))
        .then(function (response) { return response.ok ? response.json() : null; })
        .then(function (raw) { return { term: term, data: LookupCore.normalizeMoedict(raw, term) }; })
        .catch(function () { return { term: term, data: null }; });
    })).then(function (rows) {
      body.textContent = '';
      rows.forEach(function (item) {
        body.appendChild(renderDictionaryRow(item.data, item.term));
      });
    });
  }

  function setBusy(value) {
    running = value;
    $('waGenerate').textContent = value ? '生成中…' : '生成應用內容';
    $('waRegenerate').textContent = value ? '生成中…' : '重新生成';
    updateGenerateState();
  }

  function post(payload) {
    return fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(payload)
    }).then(function (response) {
      if (!response.ok) throw new Error('HTTP_' + response.status);
      return response.json();
    });
  }

  function generate() {
    var requestConfig = config();
    var configError = ZhuApplicationCore.validateConfig(requestConfig);
    if (configError || running) {
      $('waTermError').textContent = configError || '';
      return;
    }
    var mySeq = ++requestSeq;
    setBusy(true);
    $('waStatus').textContent = '正在產生適合學生的內容…';
    post({
      action: 'generateWordApplication',
      v: 1,
      terms: requestConfig.terms,
      level: requestConfig.level,
      mode: requestConfig.mode
    }).then(function (response) {
      if (mySeq !== requestSeq) return;
      if (!response || !response.ok) {
        $('waStatus').textContent = ZhuApplicationCore.errorMessage(response && response.error);
        return;
      }
      var validationError = ZhuApplicationCore.validateResult(requestConfig, response.data);
      if (validationError) {
        $('waStatus').textContent = validationError + '，請重新生成';
        return;
      }
      var next = Object.assign({}, response.data, {
        configSignature: ZhuApplicationCore.configSignature(requestConfig)
      });
      currentResult = next;
      currentVersion = 'teacher';
      renderResult();
      if (ZhuCore.setWordApplicationResult(next)) {
        $('waStatus').textContent = '內容已生成並保留在這台裝置。';
      } else {
        $('waStatus').textContent = '內容無法保留，關閉或重新整理後可能消失。';
      }
    }).catch(function () {
      if (mySeq === requestSeq) $('waStatus').textContent = '目前無法連線，請檢查網路後再試';
    }).finally(function () {
      if (mySeq === requestSeq) setBusy(false);
    });
  }

  function printCurrent() {
    if (!currentResult) return;
    var sheet = $('waPrintSheet');
    sheet.textContent = '';
    var title = document.createElement('h1');
    title.textContent = '語詞應用｜' + (currentVersion === 'teacher' ? '教師版' : '學生版');
    sheet.appendChild(title);
    var meta = document.createElement('p');
    meta.textContent = '年段：' +
      ({ low: '低年級', middle: '中年級', high: '高年級' }[currentResult.level]) +
      '　形式：' + (currentResult.mode === 'passage' ? '短文' : '造句');
    sheet.appendChild(meta);
    var content = document.createElement('div');
    if (currentResult.mode === 'passage') {
      appendVersionText(content, currentResult.passage, currentResult.terms, currentVersion === 'student');
    } else {
      currentResult.sentences.forEach(function (item, index) {
        var line = document.createElement('p');
        line.appendChild(document.createTextNode((index + 1) + '. '));
        appendVersionText(line, item.text, [item.term], currentVersion === 'student');
        content.appendChild(line);
      });
    }
    sheet.appendChild(content);
    document.body.classList.add('wa-print-mode');
    sheet.setAttribute('aria-hidden', 'false');
    window.print();
  }

  function finishPrint() {
    document.body.classList.remove('wa-print-mode');
    $('waPrintSheet').setAttribute('aria-hidden', 'true');
  }

  function mount() {
    $('waClose').addEventListener('click', close);
    $('waOverlay').addEventListener('click', function (event) {
      if (event.target === $('waOverlay')) close();
    });
    $('waLevel').addEventListener('change', function () {
      currentResult = null;
      $('waResult').classList.add('hidden');
      updateGenerateState();
    });
    $('waMode').addEventListener('change', function () {
      currentResult = null;
      $('waResult').classList.add('hidden');
      updateGenerateState();
    });
    $('waDictionary').addEventListener('toggle', function () {
      if ($('waDictionary').open) loadDictionary();
    });
    $('waGenerate').addEventListener('click', generate);
    $('waRegenerate').addEventListener('click', generate);
    $('waTeacherTab').addEventListener('click', function () {
      currentVersion = 'teacher';
      renderResult();
    });
    $('waStudentTab').addEventListener('click', function () {
      currentVersion = 'student';
      renderResult();
    });
    $('waPrint').addEventListener('click', printCurrent);
    window.addEventListener('afterprint', finishPrint);
    document.addEventListener('keydown', function (event) {
      if (event.key === 'Escape' && !$('waOverlay').classList.contains('hidden')) close();
      if (event.key === 'Tab' && !$('waOverlay').classList.contains('hidden')) trapFocus(event);
    });
  }

  function trapFocus(event) {
    var controls = Array.from($('waDialog').querySelectorAll(
      'button:not([disabled]),select:not([disabled]),input:not([disabled]),summary,[tabindex="0"]'
    )).filter(function (node) { return node.offsetParent !== null; });
    if (!controls.length) return;
    var first = controls[0];
    var last = controls[controls.length - 1];
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  }

  function appendVersionText(container, text, terms, student) {
    var parts = ZhuApplicationCore.splitTerms(text, terms);
    parts.forEach(function (part) {
      if (part.term && student) {
        container.appendChild(document.createTextNode('＿＿'));
      } else if (part.term) {
        var mark = document.createElement('mark');
        mark.textContent = part.text;
        container.appendChild(mark);
      } else {
        container.appendChild(document.createTextNode(part.text));
      }
    });
  }

  function renderResult() {
    if (!currentResult) {
      $('waResult').classList.add('hidden');
      return;
    }
    $('waResult').classList.remove('hidden');
    $('waTeacherTab').setAttribute('aria-selected', currentVersion === 'teacher' ? 'true' : 'false');
    $('waStudentTab').setAttribute('aria-selected', currentVersion === 'student' ? 'true' : 'false');
    var output = $('waOutput');
    output.textContent = '';
    var student = currentVersion === 'student';
    if (currentResult.mode === 'passage') {
      appendVersionText(output, currentResult.passage, currentResult.terms, student);
    } else {
      currentResult.sentences.forEach(function (item, index) {
        var line = document.createElement('p');
        line.appendChild(document.createTextNode((index + 1) + '. '));
        appendVersionText(line, item.text, [item.term], student);
        output.appendChild(line);
      });
    }
  }

  return {
    mount: mount,
    open: open,
    close: close
  };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = ZhuApplication;
