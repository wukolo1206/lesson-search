// 解釋卡 UI 元件：萌典優先、手動 AI、手動圖片、stale-while-revalidate 與輪詢
// 依賴 lookup-core.js（LookupCore）。模型內容只走節點 API。
var LookupCard = (function () {
  'use strict';

  var GAS_URL = 'https:/' + '/script.google.com/macros/s/AKfycbwHoHUCmxwsZyMl7TZdFVOMAYZZixJIp_JaMfsQQOpRDzHklArqfmX2nAiFW2NQwVrN/exec';
  var MAX_POLLS = 15;

  var cache = null;
  var seq = 0;          // requestSeq：每次 open/close 遞增，舊回應一律作廢
  var activeTerm = '';
  var currentMoedictData = null;
  var els = null;
  var lastFocus = null;

  function getCache() {
    if (!cache) cache = LookupCore.makeCache(window.localStorage, function () { return Date.now(); });
    return cache;
  }

  function post(payload) {
    return fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain' },
      body: JSON.stringify(payload)
    }).then(function (r) { return r.json(); });
  }

  function actionButton(label) {
    var btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = label;
    btn.style.cssText = 'border:1px solid #2b6cb0;border-radius:999px;background:#fff;color:#1f4e79;padding:7px 13px;margin:4px 6px 4px 0;font-size:.95em;cursor:pointer;';
    return btn;
  }

  // ── modal（建一次；lookup.html 以 <body data-lk-size="normal"> 用學生字級） ──
  function ensureModal() {
    if (els) return els;
    var compact = document.body.getAttribute('data-lk-size') === 'normal';
    var overlay = document.createElement('div');
    overlay.id = 'lkOverlay';
    overlay.style.cssText = 'display:none;position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:9000;align-items:center;justify-content:center;padding:16px;';
    var dlg = document.createElement('div');
    dlg.setAttribute('role', 'dialog');
    dlg.setAttribute('aria-modal', 'true');
    dlg.setAttribute('aria-label', '解釋卡');
    dlg.style.cssText = 'background:#fff;border-radius:16px;max-width:640px;width:100%;max-height:90vh;overflow-y:auto;padding:24px;position:relative;';
    var close = document.createElement('button');
    close.textContent = '✕';
    close.setAttribute('aria-label', '關閉');
    close.style.cssText = 'position:absolute;top:10px;right:12px;border:none;background:none;font-size:1.4em;cursor:pointer;color:#666;';
    var title = document.createElement('div');
    title.style.cssText = 'font-weight:bold;line-height:1.2;font-size:' + (compact ? '1.8em' : '3em') + ';';
    var zhuyin = document.createElement('div');
    zhuyin.style.cssText = 'color:#c62828;margin:2px 0 10px;font-size:' + (compact ? '1.2em' : '1.6em') + ';';
    var source = document.createElement('div');
    source.style.cssText = 'color:#1f4e79;margin:0 0 10px;font-size:.9em;font-weight:bold;';
    var extra = document.createElement('div');
    extra.style.cssText = 'color:#555;margin-bottom:8px;font-size:' + (compact ? '1em' : '1.2em') + ';';
    var expl = document.createElement('div');
    expl.style.cssText = 'line-height:1.7;white-space:pre-line;margin-bottom:10px;font-size:' + (compact ? '1.1em' : '1.5em') + ';';
    var exBox = document.createElement('div');
    exBox.style.cssText = 'line-height:1.8;font-size:' + (compact ? '1.05em' : '1.4em') + ';';
    var actions = document.createElement('div');
    actions.style.cssText = 'margin-top:12px;display:flex;flex-wrap:wrap;align-items:center;';
    var status = document.createElement('div');
    status.style.cssText = 'font-size:1.1em;color:#666;margin:10px 0;';
    var imgWrap = document.createElement('div');
    imgWrap.style.cssText = 'margin-top:12px;text-align:center;';
    var note = document.createElement('div');
    note.textContent = '';
    note.style.cssText = 'font-size:.75em;color:#aaa;margin-top:14px;text-align:right;';
    dlg.appendChild(close); dlg.appendChild(title); dlg.appendChild(zhuyin);
    dlg.appendChild(source); dlg.appendChild(extra); dlg.appendChild(expl); dlg.appendChild(exBox);
    dlg.appendChild(actions);
    dlg.appendChild(status); dlg.appendChild(imgWrap); dlg.appendChild(note);
    overlay.appendChild(dlg);
    document.body.appendChild(overlay);

    close.onclick = closeModal;
    overlay.addEventListener('click', function (ev) { if (ev.target === overlay) closeModal(); });
    document.addEventListener('keydown', function (ev) {
      if (ev.key === 'Escape' && overlay.style.display !== 'none') closeModal();
    });

    els = { overlay: overlay, close: close, title: title, zhuyin: zhuyin,
      source: source, extra: extra, expl: expl, exBox: exBox, actions: actions,
      status: status, imgWrap: imgWrap, note: note };
    return els;
  }

  function openModal(term) {
    ensureModal();
    lastFocus = document.activeElement;
    els.overlay.style.display = 'flex';
    els.title.textContent = term;
    els.close.focus();
  }

  function closeModal() {
    activeTerm = '';
    seq++;   // 進行中的輪詢與圖片等待全部作廢
    if (els) els.overlay.style.display = 'none';
    if (lastFocus && lastFocus.focus) lastFocus.focus();
  }

  function clearBody() {
    els.zhuyin.textContent = ''; els.source.textContent = ''; els.extra.textContent = '';
    els.expl.textContent = ''; els.exBox.textContent = '';
    els.actions.textContent = ''; els.status.textContent = '';
    els.imgWrap.textContent = ''; els.note.textContent = '';
  }

  function renderLoading(message) { clearBody(); els.status.textContent = message || '查詢中…'; }

  function renderMoedictUnavailable(term, message) {
    clearBody();
    els.title.textContent = term;
    els.status.textContent = message;
    var ai = actionButton(message === '萌典查詢失敗' ? '改用 AI 查詢' : 'AI 查詢');
    ai.onclick = function () { fetchAiText(term, seq, 0); };
    els.actions.appendChild(ai);
  }

  function renderActions(term, data) {
    if ((data.source || 'ai') !== 'ai') {
      var ai = actionButton('AI 查詢');
      ai.onclick = function () { fetchAiText(term, seq, 0); };
      els.actions.appendChild(ai);
    }
    if ((data.source || 'ai') === 'ai' && currentMoedictData) {
      var moedict = actionButton('查看萌典');
      moedict.onclick = function () {
        if (!isStale(term, seq)) renderData(term, currentMoedictData);
      };
      els.actions.appendChild(moedict);
    }
  }

  function renderData(term, data) {
    data = data || {};
    if (!data.source) data.source = 'ai';
    clearBody();
    els.title.textContent = term;
    els.zhuyin.textContent = data.zhuyin || '';
    els.source.textContent = data.source === 'moedict' ? '來源：萌典' : '來源：AI';
    els.extra.textContent = (data.extraReadings || [])
      .map(function (r) { return '又讀 ' + r.zhuyin + '：' + r.gloss; }).join('　');
    els.expl.textContent = data.explanation || '';
    (data.examples || []).forEach(function (ex, i) {
      var line = document.createElement('div');
      line.appendChild(document.createTextNode((i + 1) + '. '));
      LookupCore.splitHighlight(ex, term).forEach(function (seg) {
        if (seg.mark) {
          var m = document.createElement('mark');
          m.textContent = seg.text;
          line.appendChild(m);
        } else {
          line.appendChild(document.createTextNode(seg.text));
        }
      });
      els.exBox.appendChild(line);
    });
    renderActions(term, data);
    renderImage(term, data);
    els.note.textContent = data.source === 'moedict'
      ? '資料來源：萌典；例句為辭典原文，僅供教學參考'
      : '內容由 AI 生成，僅供教學參考';
  }

  function renderImage(term, data) {
    els.imgWrap.textContent = '';
    var state = data.imageStatus || 'none';
    var fileId = data.imageFileId || '';
    if (state === 'ready' && LookupCore.fileIdOk(fileId)) {
      var img = document.createElement('img');
      img.src = LookupCore.thumbUrl(fileId);
      img.alt = '';
      img.style.cssText = 'max-width:100%;border-radius:12px;opacity:0;transition:opacity .4s;';
      img.onload = function () { img.style.opacity = '1'; };
      els.imgWrap.appendChild(img);
      return;
    }
    if (state === 'pending') {
      var p = document.createElement('div');
      p.textContent = '圖片生成中…';
      p.style.cssText = 'color:#888;';
      els.imgWrap.appendChild(p);
      return;
    }
    if (state === 'quota') {
      var q = document.createElement('div');
      q.textContent = '今日圖片額度已用完';
      q.style.cssText = 'color:#888;font-size:.9em;';
      els.imgWrap.appendChild(q);
      return;
    }
    if (state === 'failed') {
      var f = document.createElement('div');
      f.textContent = '圖片生成失敗，請稍後再試';
      f.style.cssText = 'color:#888;font-size:.9em;';
      els.imgWrap.appendChild(f);
      return;
    }
    if (state === 'none' && data.imageSuitable !== false) {
      var image = actionButton('生成圖片');
      image.onclick = function () { requestImage(term, data, 0); };
      els.imgWrap.appendChild(image);
    }
  }

  function isStale(term, mySeq) { return mySeq !== seq || term !== activeTerm; }

  function fetchMoedict(term, mySeq) {
    return fetch('https://www.moedict.tw/uni/' + encodeURIComponent(term), {
      headers: { 'Accept': 'application/json' }
    }).then(function (r) {
      if (!r.ok) throw new Error('NOT_FOUND');
      return r.json();
    }).then(function (raw) {
      var data = LookupCore.normalizeMoedict(raw, term);
      if (!data) throw new Error('NOT_FOUND');
      if (isStale(term, mySeq)) return null;
      currentMoedictData = data;
      renderData(term, data);
      return data;
    });
  }

  function renderAiError(term, message) {
    if (currentMoedictData) {
      renderData(term, currentMoedictData);
      els.status.textContent = message;
      return;
    }
    clearBody();
    els.title.textContent = term;
    els.status.textContent = message;
    var retry = actionButton('AI 查詢');
    retry.onclick = function () { fetchAiText(term, seq, 0); };
    els.actions.appendChild(retry);
  }

  function fetchAiText(term, mySeq, polls) {
    var local = getCache().get(term);
    if (local && !isStale(term, mySeq)) {
      local.source = 'ai';
      renderData(term, local);
      els.status.textContent = 'AI 查詢中…';
    } else if (!currentMoedictData) {
      renderLoading('AI 查詢中…');
    } else {
      els.status.textContent = 'AI 查詢中…';
    }
    post({ action: 'lookup', term: term, v: 1 }).then(function (res) {
      if (isStale(term, mySeq)) return;
      if (res.ok) {
        getCache().put(term, res.data);          // revalidate：一律以伺服器為準
        var aiData = Object.assign({}, res.data, { source: 'ai' });
        renderData(term, aiData);
        return;
      }
      var err = res.error || {};
      if (err.code === 'PENDING' || err.code === 'RATE_LIMITED') {
        if (polls >= MAX_POLLS) { renderAiError(term, '等太久了，請再點一次'); return; }
        setTimeout(function () {
          if (!isStale(term, mySeq)) fetchAiText(term, mySeq, polls + 1);
        }, LookupCore.pollDelay(err.retryAfterMs, Math.random()));
        return;
      }
      if (err.code === 'DISABLED') getCache().remove(term);   // 治理：清本機
      renderAiError(term, err.message || '查詢失敗');
    }).catch(function () {
      if (isStale(term, mySeq)) return;
      if (!local) renderAiError(term, '網路不通，請檢查連線');
    });
  }

  function requestImage(term, data, polls) {
    if (isStale(term, seq)) return;
    if (data.imageSuitable === false) { renderImage(term, data); return; }
    if (data.imageStatus === 'ready') { renderImage(term, data); return; }
    data.imageStatus = 'pending';
    renderImage(term, data);
    var action = data.source === 'moedict' ? 'lookupImageMoedict' : 'lookupImage';
    var payload = { action: action, term: term, v: 1 };
    if (action === 'lookupImageMoedict') payload.explanation = data.explanation || '';
    post(payload).then(function (res) {
      if (isStale(term, seq)) return;
      if (res.ok && res.data.imageStatus === 'ready') {
        data.imageStatus = 'ready';
        data.imageFileId = res.data.imageFileId;
        var cached = data.source === 'ai' ? getCache().get(term) : null;
        if (cached) {
          cached.imageStatus = 'ready'; cached.imageFileId = res.data.imageFileId;
          getCache().put(term, cached);
        }
        renderImage(term, data);
        return;
      }
      var code = res.ok ? res.data.imageStatus : (res.error && res.error.code);
      if (code === 'PENDING') {
        if (polls >= MAX_POLLS) { data.imageStatus = 'failed'; renderImage(term, data); return; }
        setTimeout(function () {
          if (!isStale(term, seq)) requestImage(term, data, polls + 1);
        }, LookupCore.pollDelay(res.error && res.error.retryAfterMs, Math.random()));
        return;
      }
      data.imageStatus = code === 'QUOTA_EXCEEDED' ? 'quota' : (code || 'failed');
      renderImage(term, data);
    }).catch(function () {
      if (!isStale(term, seq)) { data.imageStatus = 'failed'; renderImage(term, data); }
    });
  }

  function open(rawTerm) {
    var term = LookupCore.normalizeTerm(rawTerm);
    if (!LookupCore.isValidTerm(term)) return;
    var mySeq = ++seq;
    activeTerm = term;
    currentMoedictData = null;
    openModal(term);
    renderLoading('先查萌典…');
    fetchMoedict(term, mySeq).catch(function (err) {
      if (isStale(term, mySeq)) return;
      renderMoedictUnavailable(term, err && err.message === 'NOT_FOUND'
        ? '萌典找不到這個詞' : '萌典查詢失敗');
    });
  }

  // ── 🔍 啟動器（右下角固定鈕＋展開輸入框） ──
  function mountLauncher() {
    if (document.getElementById('lkLauncher')) return;
    var wrap = document.createElement('div');
    wrap.id = 'lkLauncher';
    wrap.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:8999;display:flex;gap:6px;align-items:center;';
    var input = document.createElement('input');
    input.type = 'text';
    input.maxLength = 10;
    input.placeholder = '輸入國字詞語（最多 10 字）';
    input.style.cssText = 'display:none;padding:8px 10px;border:1px solid #bbb;border-radius:8px;font-size:1em;width:200px;';
    var btn = document.createElement('button');
    btn.textContent = '🔍';
    btn.setAttribute('aria-label', '查詢解釋卡');
    btn.style.cssText = 'width:44px;height:44px;border-radius:50%;border:none;background:#1F4E79;color:#fff;font-size:1.2em;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.25);';
    function submit() { if (input.value.trim()) { open(input.value); input.value = ''; } }
    btn.onclick = function () {
      if (input.style.display === 'none') { input.style.display = 'block'; input.focus(); }
      else submit();
    };
    input.addEventListener('keydown', function (ev) { if (ev.key === 'Enter') submit(); });
    wrap.appendChild(input); wrap.appendChild(btn);
    document.body.appendChild(wrap);
  }

  return { open: open, close: closeModal, mountLauncher: mountLauncher };
})();

if (typeof module !== 'undefined' && module.exports) module.exports = LookupCard;
