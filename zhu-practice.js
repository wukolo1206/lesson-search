// zhu-practice.js — 備課台「產練習卷」：勾考古題 → 產草稿 → 編輯 → 下載 docx
//
// 排版規則（段落樣式、空格寬度、編號）與 scripts/build_practice_docx.py 的
// paragraphs() 是等價實作，兩邊必須逐字一致，由 tests/test_renderer_parity.py 強制比對。
// 改動任一邊都要同步改另一邊。
var ZhuPractice = (function () {
  'use strict';

  var SCHEMA = 'practice-sheet/1';
  var BLANK = '{}';
  var CN_NUM = '一二三四五六七八九十';

  var BLANK_WIDTH = {
    '短文填國字': 3, '填字成詞': 3, '二選一': 3, '改錯': 3,
    '多音應用': 5, '多音填注音': 5
  };
  var DEFAULT_WIDTH = 4;
  var BOPOMOFO = /[ㄅ-ㄩˇˊˋ˙]/;

  // ───────────────────────────────────── 排版（對應 Python 的 paragraphs()）

  function blankWidth(answer, fallback) {
    if (answer && BOPOMOFO.test(answer)) { return 5; }
    if (answer) { return Math.max(2, answer.length + 2); }
    return fallback;
  }

  function fill(text, answers, showAnswer, fallbackWidth) {
    var pieces = String(text).split(BLANK);
    var out = '';
    for (var i = 0; i < pieces.length; i++) {
      out += pieces[i];
      if (i < answers.length) {
        out += showAnswer
          ? '（' + answers[i] + '）'
          : '（' + repeat('　', blankWidth(answers[i], fallbackWidth)) + '）';
      }
    }
    return out;
  }

  function repeat(s, n) {
    var out = '';
    for (var i = 0; i < n; i++) { out += s; }
    return out;
  }

  function asArray(ans) {
    if (ans === undefined || ans === null) { return []; }
    return (typeof ans === 'string') ? [ans] : ans;
  }

  function countBlanks(text) {
    return String(text).split(BLANK).length - 1;
  }

  function paragraphs(sheet, showAnswer) {
    var out = [];
    var head = sheet['頁首'] || '姓名：____________';
    if (showAnswer) { head += '　　【答案版】'; }
    out.push({ style: '卷_頁首', text: head });

    (sheet['大題'] || []).forEach(function (sec, idx) {
      var si = idx + 1;
      var t = sec['題型'];
      var w = BLANK_WIDTH[t] === undefined ? DEFAULT_WIDTH : BLANK_WIDTH[t];
      var num = si <= CN_NUM.length ? CN_NUM.charAt(si - 1) : String(si);
      out.push({ style: '卷_大題', text: num + '、' + sec['標題'] });

      if (sec['說明']) { out.push({ style: '卷_說明', text: sec['說明'] }); }
      (sec['提示'] || []).forEach(function (h) {
        out.push({ style: '卷_提示', text: h });
      });

      if (t === '短文填國字') {
        var answers = (sec['答案'] || []).slice();
        (sec['段落'] || []).forEach(function (p) {
          var n = countBlanks(p);
          out.push({ style: '卷_短文', text: fill(p, answers.slice(0, n), showAnswer, w) });
          answers = answers.slice(n);
        });
        return;
      }

      var numbered = !!sec['編號'];
      if (sec['題組']) {
        sec['題組'].forEach(function (g, gi) {
          if (g['標題']) {
            out.push({ style: '卷_題組', text: (gi + 1) + '. ' + g['標題'] });
          }
          (g['小題'] || []).forEach(function (item) {
            out.push({
              style: '卷_小題',
              text: fill(item['句'], asArray(item['答']), showAnswer, w)
            });
          });
        });
      } else {
        (sec['小題'] || []).forEach(function (item, ii) {
          var line = fill(item['句'], asArray(item['答']), showAnswer, w);
          out.push({ style: '卷_小題', text: numbered ? (ii + 1) + '. ' + line : line });
        });
      }
    });

    if ((sheet['來源考古題'] || []).length) {
      out.push({
        style: '卷_出處',
        text: '（出處：考古題 ' + sheet['來源考古題'].map(function (q) {
          return '#' + q;
        }).join('、') + '）'
      });
    }
    return out;
  }

  // ───────────────────────────────────── 卷稿驗證（與 Python validate() 對齊）

  function validate(sheet) {
    var errs = [];
    if (sheet['schema'] !== SCHEMA) { errs.push('schema 應為 ' + SCHEMA); }
    if (sheet['卷型'] !== 'A' && sheet['卷型'] !== 'B') { errs.push('卷型應為 A 或 B'); }
    if (!(sheet['大題'] || []).length) { errs.push('沒有任何大題'); }

    (sheet['大題'] || []).forEach(function (sec, idx) {
      var where0 = '大題[' + (idx + 1) + ']';
      if (!sec['標題']) { errs.push(where0 + ' 缺少標題'); }
      if (sec['題型'] === '短文填國字') {
        var total = (sec['段落'] || []).reduce(function (a, p) { return a + countBlanks(p); }, 0);
        if (total !== (sec['答案'] || []).length) {
          errs.push(where0 + ' 短文空格 ' + total + ' 個，答案 ' + (sec['答案'] || []).length + ' 個');
        }
        return;
      }
      var buckets = [];
      (sec['題組'] || []).forEach(function (g, gi) {
        (g['小題'] || []).forEach(function (it, ii) {
          buckets.push(['題組[' + (gi + 1) + '] 小題[' + (ii + 1) + ']', it]);
        });
      });
      (sec['小題'] || []).forEach(function (it, ii) {
        buckets.push(['小題[' + (ii + 1) + ']', it]);
      });
      buckets.forEach(function (pair) {
        var n = countBlanks(pair[1]['句']);
        var a = asArray(pair[1]['答']).length;
        if (n !== a) {
          errs.push(where0 + ' ' + pair[0] + ' 空格 ' + n + ' 個、答案 ' + a + ' 個');
        } else if (n === 0) {
          errs.push(where0 + ' ' + pair[0] + ' 沒有填空標記');
        }
      });
    });
    return errs;
  }

  // ───────────────────────────────────── 由勾選的考古題產生草稿

  function blankAt(word, ch) {
    var i = word.indexOf(ch);
    if (i < 0) { return null; }
    return word.slice(0, i + 1) + BLANK + word.slice(i + 1);
  }

  function replaceWithBlank(word, ch) {
    var i = word.indexOf(ch);
    if (i < 0) { return null; }
    return word.slice(0, i) + BLANK + word.slice(i + 1);
  }

  // 句庫（zhu-variant-bank.js）裡的句子是完整句、已過六道門禁，優先用。
  // 還沒建庫的字組才退回骨架版：拿細目表的代表語詞挖空，老師自行改寫。
  function banked(g, qtype) {
    if (typeof ZhuVariantBank === 'undefined') { return []; }
    var entry = ZhuVariantBank.get(g['類型'], g['字']);
    if (!entry) { return []; }
    return entry['句'].filter(function (s) { return s['題型'] === qtype; });
  }

  function hasBank(g) {
    return typeof ZhuVariantBank !== 'undefined'
      && !!ZhuVariantBank.get(g['類型'], g['字']);
  }

  function collect(ids) {
    // 同一組字可能掛在多題底下（如「捨／舍」同時出現在 #80 與 #99），
    // 不去重就會在同一張卷子上出兩次一樣的題。
    var groups = [];
    var seen = {};
    ids.forEach(function (id) {
      var q = ZhuPracticeData.byId(id);
      if (!q) { return; }
      q['字組'].forEach(function (g) {
        var key = g['類型'] + '|' + g['字'].join('/');
        if (seen[key]) { return; }
        seen[key] = true;
        groups.push({ q: q, g: g });
      });
    });
    return groups;
  }

  function soundSection(groups) {
    var hints = [], items = [];
    groups.forEach(function (e) {
      if (e.g['類型'] !== '多音單字') { return; }
      var ch = e.g['字'][0];
      var fromBank = banked(e.g, '多音填注音');
      if (fromBank.length) {
        var bt = [];
        fromBank.forEach(function (s) {
          if (bt.indexOf(s['答'][0]) < 0) { bt.push(s['答'][0]); }
          items.push({ '句': s['句'], '答': s['答'], '_來源': '#' + e.q.id, '_句庫': true });
        });
        if (bt.length >= 2) { hints.push(ch + '（' + bt.join(' / ') + '）'); }
        return;
      }
      var pairs = (e.g['讀音'] || {})[ch] || [];
      var tones = [];
      pairs.forEach(function (p) { if (tones.indexOf(p['音']) < 0) { tones.push(p['音']); } });
      // 只有一個讀音就不成題：提示列出唯一的音，等於把答案直接寫給學生看
      if (tones.length < 2) { return; }
      hints.push(ch + '（' + tones.join(' / ') + '）');
      pairs.forEach(function (p) {
        var sent = blankAt(p['詞'], ch);
        if (sent) { items.push({ '句': sent, '答': [p['音']], '_來源': '#' + e.q.id }); }
      });
    });
    if (!items.length) { return null; }
    return {
      '題型': '多音填注音',
      '標題': '多音字綜合測驗（填入正確的注音）',
      '提示': hints.length ? ['提示字：' + hints.join('、')] : [],
      '小題': items
    };
  }

  function pickSection(groups) {
    var items = [], pairs = [];
    groups.forEach(function (e) {
      if (e.g['類型'] !== '對比字組' || e.g['字'].length < 2) { return; }
      var bankPick = banked(e.g, '二選一');
      var words = e.g['字語詞'] || {};
      var used = [];
      e.g['字'].forEach(function (ch) {
        var hit = null;
        for (var i = 0; i < bankPick.length; i++) {
          if (bankPick[i]['答'].join('') === ch) { hit = bankPick[i]; break; }
        }
        if (hit) {
          items.push({ '句': hit['句'], '答': hit['答'], '_來源': '#' + e.q.id, '_句庫': true });
          used.push(ch);
          return;
        }
        var w = (words[ch] || [])[0];
        if (!w) { return; }
        var sent = replaceWithBlank(w, ch);
        if (sent) { items.push({ '句': sent, '答': [ch], '_來源': '#' + e.q.id }); used.push(ch); }
      });
      if (used.length >= 2) { pairs.push(used.join('／')); }
    });
    if (!items.length) { return null; }
    return {
      '題型': '二選一',
      '標題': '同音字與形近字辨析',
      '說明': pairs.length ? ('字組：' + pairs.join('、')) : '',
      '小題': items
    };
  }

  function fixSection(groups) {
    var items = [];
    groups.forEach(function (e) {
      if (e.g['類型'] !== '對比字組' || e.g['字'].length < 2) { return; }
      var fromBank2 = banked(e.g, '改錯');
      if (fromBank2.length) {
        fromBank2.forEach(function (s2) {
          items.push({ '句': s2['句'], '答': s2['答'], '_來源': '#' + e.q.id, '_句庫': true });
        });
        return;
      }
      var words = e.g['字語詞'] || {};
      e.g['字'].forEach(function (ch, i) {
        var w = (words[ch] || [])[0];
        var foil = e.g['字'][(i + 1) % e.g['字'].length];
        if (!w || !foil || foil === ch) { return; }
        var wrong = w.replace(ch, foil);
        items.push({ '句': wrong + '　→　' + BLANK, '答': [ch], '_來源': '#' + e.q.id });
      });
    });
    if (!items.length) { return null; }
    return {
      '題型': '改錯',
      '標題': '改錯',
      '說明': '下列語詞各有一個字寫錯，把正確的字填在括號裡',
      '小題': items.slice(0, 12)
    };
  }

  function groupSection(groups) {
    var blocks = [];
    groups.forEach(function (e) {
      if (e.g['類型'] !== '對比字組' || e.g['字'].length < 2) { return; }
      var gb = banked(e.g, '二選一');
      var words = e.g['字語詞'] || {};
      var items = [];
      e.g['字'].forEach(function (ch) {
        var hit = null;
        for (var i = 0; i < gb.length; i++) {
          if (gb[i]['答'].join('') === ch) { hit = gb[i]; break; }
        }
        if (hit) { items.push({ '句': hit['句'], '答': hit['答'], '_句庫': true }); return; }
        var w = (words[ch] || [])[0];
        var sent = w ? replaceWithBlank(w, ch) : null;
        if (sent) { items.push({ '句': sent, '答': [ch] }); }
      });
      if (items.length >= 2) {
        blocks.push({ '標題': '「' + e.g['字'].join('／') + '」辨析', '小題': items });
      }
    });
    if (!blocks.length) { return null; }
    return { '題型': '多音應用', '標題': '字音字形辨析', '題組': blocks };
  }

  function idiomSection(groups) {
    var items = [];
    groups.forEach(function (e) {
      if (e.g['類型'] !== '成語') { return; }
      var idiom = e.g['字'][0];
      if (!idiom || idiom.length < 2) { return; }
      items.push({ '句': BLANK + idiom.slice(1), '答': [idiom.charAt(0)], '_來源': '#' + e.q.id });
    });
    if (!items.length) { return null; }
    return { '題型': '填字成詞', '標題': '語文進階：填字成詞', '小題': items };
  }

  /**
   * 由勾選的考古題產生卷稿草稿（骨架版）。
   * 目前句子直接用細目表的代表語詞挖空，老師要在編輯區改寫成完整句子。
   * 等 P4 變式句庫建好，這裡改成從句庫抽句，介面不用動。
   */
  function buildDraft(ids, type) {
    var groups = collect(ids);
    var sections = [];
    var notes = [];

    if (type === 'A') {
      pushIf(sections, groupSection(groups));
      pushIf(sections, idiomSection(groups));
      notes.push('卷型 A 的第一大題「短文填國字」需要整篇情境短文，'
        + '骨架版無法自動生成（要等變式句庫）。目前先出後兩個大題，短文請自行加寫。');
    } else {
      pushIf(sections, soundSection(groups));
      pushIf(sections, pickSection(groups));
      pushIf(sections, fixSection(groups));
    }

    var dirty = [];
    groups.forEach(function (e) {
      if (e.g['髒污']) { dirty.push('#' + e.q.id + ' ' + e.g['字'].join('/') + '：' + e.g['髒污']); }
    });
    if (dirty.length) {
      notes.push('以下字組的字義說明在源頭資料就有問題，出題時不要沿用：' + dirty.join('；'));
    }

    return {
      sheet: {
        'schema': SCHEMA,
        '卷型': type,
        '頁首': '姓名：____________',
        '來源考古題': ids.slice(),
        '大題': sections
      },
      notes: notes
    };
  }

  function pushIf(arr, x) { if (x) { arr.push(x); } }

  // ───────────────────────────────────── 介面

  var state = { picked: {}, type: 'B', draft: null, notes: [], filters: {} };
  var chrome = {};   // 需要局部更新的 DOM：計數列與動作按鈕
  var tplCache = null;

  function updateChrome() {
    var n = pickedIds().length;
    if (chrome.count) {
      chrome.count.innerHTML = '<strong>可出題的考古題</strong><span>'
        + chrome.total + ' 題（已勾 ' + n + ' 題）</span>';
    }
    if (chrome.gen) { chrome.gen.disabled = !n; }
    if (chrome.clear) { chrome.clear.disabled = !n; }
    if (chrome.step) {
      chrome.step.innerHTML = n
        ? ('已勾 <b>' + n + '</b> 題　→　選卷型，再按「產生草稿」')
        : '第一步：在上面勾要出題的考古題（可跨年級年度多勾）';
    }
  }

  function dropDraft(container) {
    state.draft = null;
    state.notes = [];
    if (chrome.draft && chrome.draft.parentNode) {
      chrome.draft.parentNode.removeChild(chrome.draft);
      chrome.draft = null;
    }
  }

  // 模板是內嵌的（zhu-practice-template.js），不用 fetch：
  // file:// 下 fetch 會被瀏覽器擋掉，老師直接雙擊開 zhu.html 就會下載失敗。
  function loadTemplate() {
    if (tplCache) { return tplCache; }
    if (typeof ZhuPracticeTemplate === 'undefined') {
      throw new Error('模板未載入（zhu-practice-template.js），'
        + '請執行 scripts/make_practice_template.py 重新產生');
    }
    tplCache = ZhuPracticeTemplate.buffer();
    return tplCache;
  }

  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) { n.className = cls; }
    if (text !== undefined) { n.textContent = text; }
    return n;
  }

  function pickedIds() {
    return Object.keys(state.picked).filter(function (k) { return state.picked[k]; })
      .sort(function (a, b) { return Number(a) - Number(b); });
  }

  function render(container) {
    container.innerHTML = '';
    if (typeof ZhuPracticeData === 'undefined') {
      container.appendChild(el('div', 'prep-selection-empty',
        '練習卷資料尚未載入（zhu-practice-data.js）。'));
      return;
    }

    var head = el('div', 'prep-heading');
    head.innerHTML = '<strong>產練習卷</strong><span>勾考古題 → 產草稿 → 改句子 → 下載 Word。'
      + '考點與答案照抄考古題，題目重寫。</span>';
    container.appendChild(head);

    chrome = {};
    container.appendChild(buildFilterBar(container));
    container.appendChild(buildQuestionList(container));
    container.appendChild(buildActionBar(container));
    if (state.draft) {
      chrome.draft = buildDraftPanel(container);
      container.appendChild(chrome.draft);
    }
    updateChrome();
  }

  function buildFilterBar(container) {
    var bar = el('div', 'panel prep-filters');
    var qs = ZhuPracticeData.questions();
    var grades = uniq(qs.map(function (q) { return q['年級']; }));
    var years = uniq(qs.map(function (q) { return String(q['年度']); })).sort();
    var cats = uniq(qs.map(function (q) { return q['類別']; }));

    [['grade', '年級（全部）', grades], ['year', '年度（全部）', years],
     ['category', '類別（全部）', cats]].forEach(function (spec) {
      var sel = el('select');
      var blank = el('option', null, spec[1]);
      blank.value = '';
      sel.appendChild(blank);
      spec[2].forEach(function (v) {
        var o = el('option', null, v);
        o.value = v;
        sel.appendChild(o);
      });
      sel.value = state.filters[spec[0]] || '';
      sel.onchange = function () {
        state.filters[spec[0]] = sel.value;
        dropDraft(container);
        render(container);
      };
      bar.appendChild(sel);
    });
    return bar;
  }

  function uniq(arr) {
    var seen = {}, out = [];
    arr.forEach(function (v) { if (v && !seen[v]) { seen[v] = 1; out.push(v); } });
    return out;
  }

  function filtered() {
    var f = state.filters;
    return ZhuPracticeData.questions().filter(function (q) {
      if (f.grade && q['年級'] !== f.grade) { return false; }
      if (f.year && String(q['年度']) !== f.year) { return false; }
      if (f.category && q['類別'] !== f.category) { return false; }
      return true;
    });
  }

  function buildQuestionList(container) {
    var panel = el('div', 'panel pgen-list');
    var rows = filtered();
    var hint = el('div', 'prep-selection-heading');
    panel.appendChild(hint);
    chrome.count = hint;
    chrome.total = rows.length;

    rows.forEach(function (q) {
      var row = el('label', 'pgen-q');
      var cb = el('input');
      cb.type = 'checkbox';
      cb.checked = !!state.picked[q.id];
      // 勾選不整頁重繪：重繪會換掉整份清單的 DOM，捲動位置會跳回頂端
      cb.onchange = function () {
        state.picked[q.id] = cb.checked;
        if (state.draft) { dropDraft(container); }
        updateChrome();
      };
      row.appendChild(cb);

      var body = el('span', 'pgen-q-body');
      body.appendChild(el('span', 'pgen-q-title',
        '#' + q.id + '　' + q['年級'] + q['年度'] + '年第' + q['題號'] + '題　' + q['重點']));

      var chips = el('span', 'pgen-q-chips');
      q['字組'].forEach(function (g) {
        var ok = g['素材'] === '細目表';
        var chip = el('span', 'chip' + (ok ? '' : ' pgen-chip-thin'),
          g['字'].join('／') + (ok ? '' : '（缺素材）'));
        if (g['髒污']) { chip.title = '源頭資料有問題：' + g['髒污']; chip.textContent += ' ⚠'; }
        chips.appendChild(chip);
      });
      body.appendChild(chips);
      row.appendChild(body);
      panel.appendChild(row);
    });

    if (!rows.length) {
      panel.appendChild(el('div', 'prep-selection-empty', '這個篩選條件下沒有可出題的考古題。'));
    }
    return panel;
  }

  function buildActionBar(container) {
    var bar = el('div', 'panel prep-selection pgen-actions');
    var step = el('div', 'pgen-step');
    bar.appendChild(step);
    chrome.step = step;
    var actions = el('div', 'prep-selection-actions');

    [['B', '操練卷（多音／二選一／改錯）'], ['A', '情境卷（辨析題組／填字成詞）']]
      .forEach(function (spec) {
        var b = el('button', 'btn' + (state.type === spec[0] ? ' btn-primary' : ''), spec[1]);
        b.setAttribute('aria-pressed', state.type === spec[0] ? 'true' : 'false');
        b.onclick = function () {
          state.type = spec[0];
          dropDraft(container);
          render(container);
        };
        actions.appendChild(b);
      });

    var gen = el('button', 'btn btn-primary', '產生草稿 →');
    gen.onclick = function () {
      var r = buildDraft(pickedIds(), state.type);
      state.draft = r.sheet;
      state.notes = r.notes;
      if (!r.sheet['大題'].length) {
        state.notes = state.notes.concat(['勾選的題目湊不出這個卷型的任何大題，'
          + '請改選卷型或多勾幾題（尤其是有細目素材的題）。']);
      }
      if (chrome.draft && chrome.draft.parentNode) {
        chrome.draft.parentNode.removeChild(chrome.draft);
      }
      chrome.draft = buildDraftPanel(container);
      container.appendChild(chrome.draft);
      chrome.draft.scrollIntoView({ behavior: 'smooth', block: 'start' });
    };
    actions.appendChild(gen);
    chrome.gen = gen;

    var clear = el('button', 'btn', '清除勾選');
    clear.onclick = function () {
      state.picked = {};
      dropDraft(container);
      render(container);
    };
    actions.appendChild(clear);
    chrome.clear = clear;

    bar.appendChild(actions);
    return bar;
  }

  function buildDraftPanel(container) {
    var panel = el('div', 'panel pgen-draft');
    var nBank = 0, nStub = 0;
    state.draft['大題'].forEach(function (sec) {
      var all = (sec['小題'] || []).concat(
        (sec['題組'] || []).reduce(function (a, g) { return a.concat(g['小題'] || []); }, []));
      all.forEach(function (it) { if (it['_句庫']) { nBank++; } else { nStub++; } });
    });
    panel.appendChild(el('div', 'prep-selection-heading')).innerHTML =
      '<strong>草稿</strong><span>'
      + '<b>✓ ' + nBank + ' 句</b>來自句庫（完整句、已過品質門禁，可直接用）；'
      + '<b>✎ ' + nStub + ' 句</b>還是骨架（只把語詞挖空，請改寫成完整句子）。'
      + '<code>{}</code> 是填空位置，不要刪掉。</span>';

    state.notes.forEach(function (n) {
      panel.appendChild(el('div', 'pgen-note', '※ ' + n));
    });

    state.draft['大題'].forEach(function (sec, si) {
      panel.appendChild(el('div', 'pgen-sec-title',
        CN_NUM.charAt(si) + '、' + sec['標題']));
      if (sec['說明']) { panel.appendChild(el('div', 'pgen-note', sec['說明'])); }
      (sec['提示'] || []).forEach(function (h) {
        panel.appendChild(el('div', 'pgen-note', h));
      });

      var lists = [];
      (sec['題組'] || []).forEach(function (g) {
        lists.push([g['標題'], g['小題']]);
      });
      if (sec['小題']) { lists.push([null, sec['小題']]); }

      lists.forEach(function (pair) {
        if (pair[0]) { panel.appendChild(el('div', 'pgen-group-title', pair[0])); }
        pair[1].forEach(function (item) {
          var row = el('div', 'pgen-item');
          var ta = el('textarea', 'pgen-input');
          ta.value = item['句'];
          ta.rows = 1;
          ta.oninput = function () {
            item['句'] = ta.value;
            refreshWarn();
          };
          row.appendChild(ta);
          row.appendChild(el('span', 'pgen-ans' + (item['_句庫'] ? ' pgen-ans-bank' : ''),
            (item['_句庫'] ? '✓ ' : '✎ ') + '答：' + asArray(item['答']).join('／')
            + (item['_來源'] ? '　' + item['_來源'] : '')));
          panel.appendChild(row);
        });
      });
    });

    var warn = el('div', 'pgen-warn');
    panel.appendChild(warn);

    var actions = el('div', 'prep-selection-actions');
    var dl = el('button', 'btn btn-primary', '⬇ 下載 Word（填空版＋答案版）');
    dl.onclick = function () { downloadBoth(warn, dl); };
    actions.appendChild(dl);

    var dj = el('button', 'btn', '⬇ 下載卷稿 JSON');
    dj.onclick = function () { downloadJson(); };
    actions.appendChild(dj);
    panel.appendChild(actions);

    function refreshWarn() {
      var errs = validate(state.draft);
      warn.textContent = errs.length
        ? ('還不能下載：' + errs.slice(0, 3).join('；')
           + (errs.length > 3 ? ' …等 ' + errs.length + ' 項' : ''))
        : '';
      dl.disabled = !!errs.length;
    }
    refreshWarn();
    return panel;
  }

  function downloadBoth(warn, btn) {
    var errs = validate(state.draft);
    if (errs.length) { warn.textContent = '還不能下載：' + errs.join('；'); return; }
    btn.disabled = true;
    warn.textContent = '產生中…';
    try {
      var buf = loadTemplate();
      var base = '練習卷_' + state.draft['卷型'] + '卷_' + stamp();
      [[false, '填空版'], [true, '答案版']].forEach(function (spec) {
        var bytes = ZhuDocx.build(buf, paragraphs(state.draft, spec[0]));
        ZhuDocx.download(bytes, base + '_' + spec[1] + '.docx');
      });
      warn.textContent = '已下載兩個檔案。';
    } catch (e) {
      warn.textContent = '產生失敗：' + e.message;
    }
    btn.disabled = false;
  }

  function downloadJson() {
    var clean = JSON.parse(JSON.stringify(state.draft));
    var blob = new Blob([JSON.stringify(clean, null, 2)], { type: 'application/json' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '卷稿_' + state.draft['卷型'] + '卷_' + stamp() + '.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  function stamp() {
    var d = new Date();
    function p(n) { return (n < 10 ? '0' : '') + n; }
    return String(d.getFullYear()) + p(d.getMonth() + 1) + p(d.getDate())
      + '_' + p(d.getHours()) + p(d.getMinutes());
  }

  var api = {
    paragraphs: paragraphs,
    validate: validate,
    buildDraft: buildDraft,
    blankWidth: blankWidth,
    fill: fill,
    render: render,
    _state: state
  };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  return api;
})();
