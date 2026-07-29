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

  // 填國字：一個字一格（「　」），與學生的書寫習慣一致。
  // 填注音：一個音最多三個符號加聲調（如 ㄓㄨㄤˋ），單格寫不下，內部留寬。
  function blankWidth(answer, fallback) {
    if (answer && BOPOMOFO.test(answer)) { return 5; }
    if (answer) { return Math.max(1, answer.length); }
    return fallback;
  }

  function fill(text, answers, showAnswer, fallbackWidth) {
    var pieces = String(text).split(BLANK);
    var out = '';
    for (var i = 0; i < pieces.length; i++) {
      out += pieces[i];
      if (i < answers.length) {
        out += showAnswer
          ? '「' + answers[i] + '」'
          : '「' + repeat('　', blankWidth(answers[i], fallbackWidth)) + '」';
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
        // 去重鍵必須排序：資料裡「提/題」與「題/提」是同一組字的兩種寫法，
        // 不排序會讓 4 組重複字組逃過去重，同一張卷子出兩次一樣的辨析題。
        var key = g['類型'] + '|' + ZhuPassageSlots.canonicalKey(g['字']);
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

  // 規格 §8.2：放行只對「當次檢查結果」有效。正文、答案、slots、年級
  // 任一改變就清掉放行與檢查結果——否則老師放行了超綱字、接著改了正文，
  // 舊放行會沿用到沒檢查過的新內容。
  function invalidatePassage() {
    state.passage.problems = [];
    state.passage.overridden = false;
    state.passage.confirmed = false;
    state.passage.checked = false;
  }

  function passageBlockers() {
    return state.passage.problems.filter(function (p) { return p.severity === 'blocking'; });
  }

  // 可併入草稿／可下載的條件
  function passageReady() {
    var p = state.passage;
    if (!p.doc) { return false; }
    if (!p.checked) { return false; }
    if (passageBlockers().length) { return false; }
    if (!p.confirmed) { return false; }
    var soft = p.problems.filter(function (x) { return x.severity !== 'blocking'; });
    return soft.length === 0 || p.overridden;
  }

  function rebuildSlots() {
    var entries = [];
    Object.keys(state.picked).forEach(function (id) {
      var q = ZhuPracticeData.byId(id);
      if (!q) { return; }
      q['字組'].forEach(function (g) { entries.push({ qid: q.id, g: g }); });
    });
    var out = ZhuPassageSlots.build(entries, {
      bank: (typeof ZhuVariantBank !== 'undefined') ? ZhuVariantBank : null,
      words: (typeof ZhuBopo !== 'undefined') ? ZhuBopo : null,
      learned: (typeof ZhuLearnedChars !== 'undefined') ? ZhuLearnedChars : null,
      grade: gradeNumber(state.passage.grade),
      overrides: state.passage.overrides
    });
    var before = JSON.stringify(state.passage.slots);
    state.passage.slots = out.slots;
    state.passage.ok = out.ok;
    state.passage.conflicts = out.conflicts;
    state.passage.excluded = out.excluded;
    state.passage.shortBy = out.shortBy;
    // slots 變了才清放行（§8.2）。每次 render 都無條件清會讓老師剛勾好的
    // 確認項立刻被抹掉，畫面上看起來像按鈕壞了。
    if (JSON.stringify(out.slots) !== before) { invalidatePassage(); }
    return out;
  }

  // ZhuPassageSlots.build() 吃數字年級、ZhuPassagePrompt/ZhuPassageGates 吃中文年級字串。
  var GRADE_NUMBER = { '三年級': 3, '四年級': 4, '五年級': 5, '六年級': 6 };
  function gradeNumber(grade) { return GRADE_NUMBER[grade] || null; }

  function runPassageGates() {
    var p = state.passage;
    var parsed = ZhuPassageParse.parse(p.raw);
    if (!parsed.ok) {
      p.doc = null;
      p.problems = parsed.problems;
      p.checked = true;
      return;
    }
    p.doc = { '段落': parsed['段落'], '答案': parsed['答案'] };
    p.problems = ZhuPassageGates.check(p.doc, {
      slots: p.slots,
      grade: p.grade,
      bopo: (typeof ZhuBopo !== 'undefined') ? ZhuBopo : null,
      learned: (typeof ZhuLearnedChars !== 'undefined') ? ZhuLearnedChars : null
    });
    p.checked = true;
  }

  function passageSection() {
    var p = state.passage;
    if (!passageReady()) { return null; }
    return {
      '題型': '短文填國字',
      '標題': '填入適當的國字',
      '段落': p.doc['段落'].slice(),
      '答案': p.doc['答案'].slice()
    };
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
      // 規格 §4.1：後兩大題只用「本篇採用」的最多 12 組，否則三大題閉環會破——
      // 第一大題練的字與第二、三大題不是同一批。
      var keys = {};
      state.passage.slots.forEach(function (s) { keys[s.groupKey] = true; });
      var scoped = state.passage.slots.length
        ? groups.filter(function (e) {
            return e.g['類型'] !== '對比字組'
              || keys[ZhuPassageSlots.canonicalKey(e.g['字'])];
          })
        : groups;

      var passage = passageSection();
      if (passage) {
        sections.push(passage);
      } else {
        notes.push('第一大題「短文填國字」尚未完成——請先在上方產提示詞、'
          + '貼回外部 AI 的短文並通過檢查。目前先出後兩個大題。');
      }
      pushIf(sections, groupSection(scoped));
      pushIf(sections, idiomSection(scoped));
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

  var state = {
    picked: {}, type: 'B', draft: null, notes: [], filters: {},
    // 短文（P5）
    passage: {
      grade: '',          // 老師指定的出卷年級，空字串＝未指定
      overrides: {},      // groupKey → 選定的 target
      slots: [],
      ok: false,
      conflicts: [],      // target 撞字（§4.1 第 4 步）
      excluded: [],       // 超過 12 組被截掉的（§4.1）
      shortBy: 0,         // 不足 4 組時還差幾組
      promptText: '',
      raw: '',            // 老師貼回的原始文字
      doc: null,          // 解析後的 { 段落, 答案 }
      problems: [],
      overridden: false,  // 老師是否已放行
      confirmed: false,   // 老師是否已勾「故事連貫」
      checked: false      // 是否跑過門禁且結果對應目前內容
    }
  };
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

    if (!state.passage.grade) {
      try { state.passage.grade = localStorage.getItem('zhuPassageGrade') || ''; } catch (e) {}
    }

    var head = el('div', 'prep-heading');
    head.innerHTML = '<strong>產練習卷</strong><span>勾考古題 → 產草稿 → 改句子 → 下載 Word。'
      + '考點與答案照抄考古題，題目重寫。</span>';
    container.appendChild(head);

    chrome = {};
    container.appendChild(buildFilterBar(container));
    container.appendChild(buildQuestionList(container));
    container.appendChild(buildActionBar(container));
    // 短文面板有自己的錨點，勾選考古題時只重繪這個錨點（見 refreshPassagePanel），
    // 不整頁重繪——整頁重繪會換掉整份考古題清單的 DOM，捲動位置跳回頂端，
    // 而勾題→看字組湊到幾組正是這個面板最常互動的動作。
    var passageAnchor = el('div', 'pgen-passage-anchor');
    container.appendChild(passageAnchor);
    chrome.passageAnchor = passageAnchor;
    refreshPassagePanel(container);
    if (state.draft) {
      chrome.draft = buildDraftPanel(container);
      container.appendChild(chrome.draft);
    }
    updateChrome();
  }

  function refreshPassagePanel(container) {
    if (!chrome.passageAnchor) { return; }
    chrome.passageAnchor.innerHTML = '';
    if (state.type === 'A' && pickedIds().length) {
      rebuildSlots();
      var pp = buildPassagePanel(container);
      if (pp) { chrome.passageAnchor.appendChild(pp); }
    }
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
      // 勾選不整頁重繪：重繪會換掉整份清單的 DOM，捲動位置會跳回頂端。
      // 短文面板獨立重繪（見 refreshPassagePanel）——卷型 A 時勾題就是為了
      // 湊對比字組，面板必須跟著每次勾選即時更新，否則老師會以為勾選沒反應。
      cb.onchange = function () {
        state.picked[q.id] = cb.checked;
        if (state.draft) { dropDraft(container); }
        updateChrome();
        refreshPassagePanel(container);
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

  function buildPassagePanel(container) {
    var p = state.passage;
    var panel = el('div', 'panel pgen-passage');
    panel.appendChild(el('div', 'prep-selection-heading')).innerHTML =
      '<strong>第一大題：短文填國字</strong>'
      + '<span>系統產提示詞 → 你貼去 ChatGPT／Gemini → 把回覆貼回來 → 自動檢查</span>';

    // 出卷年級（規格 §4.3：必選，不從考古題猜——老師可能跨年級勾題）
    var gradeRow = el('div', 'pgen-step');
    gradeRow.appendChild(document.createTextNode('學生程度（出卷年級）：'));
    var sel = document.createElement('select');
    var blankOpt = document.createElement('option');
    blankOpt.value = ''; blankOpt.textContent = '請選擇';
    sel.appendChild(blankOpt);
    ['三年級', '四年級', '五年級', '六年級'].forEach(function (g) {
      var o = document.createElement('option');
      o.value = g; o.textContent = g;
      if (p.grade === g) { o.selected = true; }
      sel.appendChild(o);
    });
    sel.onchange = function () {
      p.grade = sel.value;
      try { localStorage.setItem('zhuPassageGrade', sel.value); } catch (e) {}
      invalidatePassage();
      refreshPassagePanel(container);
    };
    gradeRow.appendChild(sel);
    panel.appendChild(gradeRow);

    // target 衝突（§4.1 第 4 步）必須在產提示詞前解決，否則送出的提示詞注定失敗
    if (p.conflicts && p.conflicts.length) {
      p.conflicts.forEach(function (c) {
        panel.appendChild(el('div', 'pgen-warn',
          '「' + c.char + '」同時被 ' + c.groupKeys.join('、')
          + ' 選為要挖的字。請切換其中一組要挖的字，或取消勾選該題。'));
      });
    }

    // 注意：`build()` 在組數不足時仍會回傳 slots（讓畫面能顯示已經湊到哪些字組），
    // 門檻只反映在 ok／shortBy。所以這裡不可以用 `!p.slots.length` 判斷不足。
    if (!p.ok) {
      var got = p.slots.length;
      var pickedN = pickedIds().length;
      var ratio = got > 0 ? (pickedN / got) : 0;
      var suggestMore = ratio > 0
        ? Math.max(1, Math.ceil(ratio * p.shortBy))
        : null;
      var msg = p.shortBy
        ? ('已湊到 ' + got + ' 組對比字組，還差 ' + p.shortBy + ' 組才能產生短文（至少 4 組）。'
           + (suggestMore
              ? ('依目前比例（已勾 ' + pickedN + ' 題 ÷ 已得 ' + got + ' 組），'
                 + '建議再多勾約 ' + suggestMore + ' 題有對比字組的考古題。')
              : '請多勾幾題有對比字組的考古題。'))
        : '請先解決上面的字組衝突。';
      panel.appendChild(el('div', 'pgen-note', msg));
      return panel;
    }

    if (p.excluded && p.excluded.length) {
      panel.appendChild(el('div', 'pgen-note',
        '本篇採用 ' + p.slots.length + ' 組，另有 ' + p.excluded.length + ' 組未納入：'
        + p.excluded.map(function (e) { return e.chars.join('／'); }).join('、')
        + '。第二、三大題也只會用本篇採用的這批字組。'));
    }

    // 每組的目標字切換
    var slotRow = el('div', 'pgen-step');
    slotRow.appendChild(document.createTextNode('要挖的字（可切換）：'));
    p.slots.forEach(function (s) {
      var wrap = el('span', 'pgen-slot');
      var ss = document.createElement('select');
      var all = [s.target].concat(s.distractors.map(function (d) { return d.char; }));
      all.sort().forEach(function (c) {
        var o = document.createElement('option');
        o.value = c; o.textContent = c;
        if (c === s.target) { o.selected = true; }
        ss.appendChild(o);
      });
      // 只重繪面板：年級與目標字會改變 slots，但 slots 只影響這個面板，
      // 題目清單不受影響。整頁重繪會把畫面彈回題目清單頂端，老師勾一次就要
      // 重捲一次，還看不出剛剛的操作有沒有生效。
      ss.onchange = function () {
        p.overrides[s.groupKey] = ss.value;
        refreshPassagePanel(container);
      };
      wrap.appendChild(ss);
      wrap.appendChild(el('span', 'pgen-note', s.groupKey
        + (s.wordSource === '無' ? '（無參考語詞）' : '')));
      slotRow.appendChild(wrap);
    });
    panel.appendChild(slotRow);

    // 產提示詞
    var actions = el('div', 'prep-selection-actions');
    var btn = el('button', 'btn', '① 產短文提示詞');
    btn.disabled = !p.grade;
    btn.onclick = function () {
      var text = ZhuPassagePrompt.build(p.slots, p.grade);
      copyToClipboard(text);
      p.promptText = text;
      refreshPassagePanel(container);
    };
    actions.appendChild(btn);
    panel.appendChild(actions);
    if (!p.grade) {
      panel.appendChild(el('div', 'pgen-note', '※ 請先選出卷年級才能產提示詞。'));
    }
    if (p.promptText) {
      panel.appendChild(el('div', 'pgen-note', '提示詞已複製到剪貼簿，貼給外部 AI：'));
      panel.appendChild(el('div', 'pgen-prompt-box', p.promptText));
    }

    // 貼回
    panel.appendChild(el('div', 'pgen-step', '② 把 AI 的回覆整段貼在這裡：'));
    var ta = document.createElement('textarea');
    ta.value = p.raw;
    ta.oninput = function () { p.raw = ta.value; invalidatePassage(); };
    panel.appendChild(ta);

    var act2 = el('div', 'prep-selection-actions');
    var chk = el('button', 'btn', '③ 檢查');
    // 檢查後最需要就地看結果，整頁重繪會把老師彈離剛產生的問題清單
    chk.onclick = function () { runPassageGates(); refreshPassagePanel(container); };
    act2.appendChild(chk);
    panel.appendChild(act2);

    // 檢查結果
    if (p.checked) {
      if (!p.problems.length) {
        panel.appendChild(el('div', 'pgen-ok', '✓ 四道門禁全過'));
      } else {
        // 摘要放最上面：10～12 個空、好幾條軟性警告時面板很長，
        // 沒有摘要老師得整個捲完才知道自己卡在哪、能不能用。
        var nBlock = 0, nConfirm = 0, nOver = 0;
        p.problems.forEach(function (x) {
          if (x.severity === 'blocking') { nBlock++; }
          else if (x.severity === 'confirm') { nConfirm++; }
          else { nOver++; }
        });
        var parts = [];
        if (nBlock) { parts.push(nBlock + ' 個必須修正'); }
        if (nConfirm) { parts.push(nConfirm + ' 個需確認'); }
        if (nOver) { parts.push(nOver + ' 個可放行'); }
        panel.appendChild(el('div',
          nBlock ? 'pgen-summary pgen-summary-blocking' : 'pgen-summary',
          '檢查結果：' + parts.join('、') + (nBlock ? '（還不能用）' : '')));

        p.problems.forEach(function (x) {
          var where = [];
          if (x.paragraph) { where.push('第 ' + x.paragraph + ' 段'); }
          if (x.sentence) { where.push('第 ' + x.sentence + ' 句'); }
          if (x.blank) { where.push('第 ' + x.blank + ' 空'); }
          panel.appendChild(el('div', 'pgen-issue pgen-issue-' + x.severity,
            (where.length ? where.join('') + '：' : '') + x.message));
        });
        var fix = el('button', 'btn', '複製修正提示詞');
        fix.onclick = function () {
          copyToClipboard(ZhuPassagePrompt.buildFix(p.problems, p.raw));
        };
        panel.appendChild(fix);

        var soft = p.problems.filter(function (x) { return x.severity !== 'blocking'; });
        if (!passageBlockers().length && soft.length) {
          // 各自獨立成行：這一項是「放棄門禁警告」，與下面「聲明自己讀過」
          // 是語意完全不同的兩個決定，並排會被當成同一個控制項的兩半一起掃過去勾掉。
          var ov = el('label', 'pgen-confirm');
          var cb = document.createElement('input');
          cb.type = 'checkbox'; cb.checked = p.overridden;
          cb.onchange = function () { p.overridden = cb.checked; refreshPassagePanel(container); };
          ov.appendChild(cb);
          ov.appendChild(el('span', 'pgen-confirm-text',
            '我確認以上問題可以接受，仍要使用這篇短文'));
          panel.appendChild(ov);
        }
        if (passageBlockers().length) {
          panel.appendChild(el('div', 'pgen-note',
            '※ 上面標紅的問題無法放行——答案錯位或洩題的卷子沒有發下去的價值，請修正後重貼。'));
        }
      }

      // 敘事結構是人工覆核項，門禁查不到（規格 §4）。
      // 視覺上明顯區隔（分隔線＋說明小字）：整個功能只靠這一個勾選框保證有人
      // 真的讀過短文，它若看起來像上面放行框的附屬品，這道人工防線就形同虛設。
      if (p.doc && !passageBlockers().length) {
        var human = el('div', 'pgen-human');
        human.appendChild(el('div', 'pgen-human-hint',
          '下面這一項程式檢查不到，需要你自己讀過：'));
        var lb = el('label', 'pgen-confirm');
        var cb2 = document.createElement('input');
        cb2.type = 'checkbox'; cb2.checked = p.confirmed;
        cb2.onchange = function () { p.confirmed = cb2.checked; refreshPassagePanel(container); };
        lb.appendChild(cb2);
        lb.appendChild(el('span', 'pgen-confirm-text',
          '我已確認這篇短文是連貫的故事（有起因、經過、結果）'));
        human.appendChild(lb);
        panel.appendChild(human);
      }
    }

    return panel;
  }

  function copyToClipboard(text) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        var p = navigator.clipboard.writeText(text);
        // writeText 失敗時是回傳 rejected Promise，不是 throw——try/catch 接不到。
        // 少了這個 fallback，老師在 file:// 開檔或未授權剪貼簿時按「產提示詞」，
        // 畫面寫著「已複製到剪貼簿」但其實什麼都沒複製到，貼過去是空的。
        if (p && typeof p['catch'] === 'function') {
          p['catch'](function () { legacyCopy(text); });
        }
        return;
      }
    } catch (e) {}
    legacyCopy(text);
  }

  function legacyCopy(text) {
    var t = document.createElement('textarea');
    t.value = text;
    document.body.appendChild(t);
    t.select();
    try { document.execCommand('copy'); } catch (e) {}
    document.body.removeChild(t);
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
      + '<code>{}</code> 是填空位置，不要刪掉——'
      + '下載的 Word 裡會變成「　」格子，答案版則直接印出答案。</span>';

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
    invalidatePassage: invalidatePassage,
    passageReady: passageReady,
    passageSection: passageSection,
    _state: state
  };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  return api;
})();
