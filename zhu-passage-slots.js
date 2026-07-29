// zhu-passage-slots.js — 短文填國字的空位映射
//
// 從勾選的字組算出「第幾空挖哪個字、干擾字是誰、參考語詞從哪來」。
// 順序不可調換：去重 → 選 target → 衝突檢查 → 截斷 → 下限檢查。
// 先截斷再去重會發生「取了 12 組其實只有 10 組不同」。
var ZhuPassageSlots = (function () {
  'use strict';

  var MAX_SLOTS = 12;
  var MIN_SLOTS = 4;

  // canonical key 必須排序且不含題號：實測 19 種字組重複出現，
  // 其中「提/題 vs 題/提」等 4 組只是順序相反，不排序就會被當成兩組。
  function canonicalKey(chars) {
    return uniq(chars).sort().join('/');
  }

  function uniq(list) {
    var seen = {}, out = [];
    (list || []).forEach(function (x) {
      if (x && !seen[x]) { seen[x] = true; out.push(x); }
    });
    return out;
  }

  // 候選字＝字 ∪ 干擾字。實測 165/166 組的「干擾字」是空的，混淆字全在「字」欄，
  // 只讀「干擾字」會漏掉 99%。
  function candidates(g) {
    return uniq((g['字'] || []).concat(g['干擾字'] || []));
  }

  function wordsOf(g, ch) {
    var m = g['字語詞'] || {};
    return (m[ch] || []).slice();
  }

  var CJK = /^[一-鿿]+$/;
  var MIN_WORD_LEN = 2;
  var MAX_WORD_LEN = 4;

  // 參考語詞的每個字都必須在該年級的累計已學字內。
  //
  // 實測發現我們自己餵超綱字給 AI：「到/倒」的參考語詞是「到達」，
  // 而「達」是四年級才教的字——AI 照著用，然後被難度門禁擋下，
  // 老師還要多跑一輪 buildFix。餵超綱詞比不餵更糟。
  //
  // 目標字與干擾字本身豁免：它們正是這一題要考的字，本來就可能超綱。
  // 沒有傳 learned／grade 時整道檢查跳過，不硬性依賴。
  function gradeFilter(list, ctxObj) {
    var learned = ctxObj.learned, grade = ctxObj.grade, exempt = ctxObj.exempt;
    if (!learned || typeof learned.has !== 'function' ||
        grade === null || grade === undefined || grade === '') {
      return list;
    }
    return list.filter(function (w) {
      for (var i = 0; i < w.length; i++) {
        var c = w.charAt(i);
        if (exempt[c]) { continue; }
        if (!learned.has(grade, c)) { return false; }
      }
      return true;
    });
  }

  // 從例句中含目標字的位置，枚舉所有包含該字的 2～4 字子字串，
  // 用詞典驗證哪些是真的詞，取最長者。
  //
  // 舊版取「答案字前後各一字」的固定字元窗口，不看詞界，抽出來的是
  // 「太重了」「又清又」「心跌了」這種贅字或破碎片段——145 組裡有 48 組中招，
  // 而這些會原封不動當參考語詞寫進提示詞，等於教外部 AI 造怪詞。
  // 寧可回空陣列讓它落到階梯 3 標「無」，也不要給一個假詞。
  function realWordsAt(words, sentence, idx) {
    var out = [];
    if (!words || typeof words.lookup !== 'function') { return out; }
    // 由長到短：同一位置若「公園裡」與「公園」都成詞，取較長的無妨，
    // 但實務上只有真詞會過關，長詞優先能保住四字成語。
    for (var len = MAX_WORD_LEN; len >= MIN_WORD_LEN; len--) {
      // start 需讓子字串涵蓋 idx，且不越界
      for (var start = idx - len + 1; start <= idx; start++) {
        if (start < 0 || start + len > sentence.length) { continue; }
        var cand = sentence.substr(start, len);
        if (!CJK.test(cand)) { continue; }
        if (!words.lookup(cand).length) { continue; }
        if (out.indexOf(cand) < 0) { out.push(cand); }
      }
      if (out.length) { return out; }
    }
    return out;
  }

  // 句庫的 key 用原始字序，故兩種排列都試。
  function bankWords(bank, words, chars, ch) {
    if (!bank || !words) { return []; }
    var entry = bank.get('對比字組', chars) ||
                bank.get('對比字組', chars.slice().reverse());
    if (!entry) { return []; }
    var out = [];
    (entry['句'] || []).forEach(function (item) {
      var ans = item['答'];
      ans = (typeof ans === 'string') ? [ans] : ans;
      if (!ans || ans.join('') !== ch) { return; }
      var filled = String(item['句'] || '').replace('{}', ch);
      // 同一句可能多處出現該字（改錯題把答案字附在句尾），逐處都試
      for (var idx = filled.indexOf(ch); idx >= 0; idx = filled.indexOf(ch, idx + 1)) {
        realWordsAt(words, filled, idx).forEach(function (w) {
          if (out.indexOf(w) < 0) { out.push(w); }
        });
      }
    });
    return out;
  }

  // 單一候選字的完整階梯：字語詞 → 變式句庫 → 無。兩層都先過年級檢查。
  function resolveWords(g, ch, cands, ctxObj) {
    var w = gradeFilter(wordsOf(g, ch), ctxObj);
    if (w.length) { return { words: w, source: '字語詞' }; }
    var bw = gradeFilter(bankWords(ctxObj.bank, ctxObj.words, g['字'] || cands, ch), ctxObj);
    if (bw.length) { return { words: bw, source: '變式句庫' }; }
    return { words: [], source: '無' };
  }

  function pickTarget(g, cands, ctxObj, override) {
    // override 也要跑完整階梯：指定的字若其實沒有「字語詞」，
    // 硬標成 '字語詞' 卻回空 targetWords 會讓下游的來源標籤失真。
    if (override && cands.indexOf(override) >= 0) {
      var r = resolveWords(g, override, cands, ctxObj);
      return { target: override, words: r.words, source: r.source };
    }
    // 階梯 1：有「字語詞」的候選字優先（實測 59 組的 字[0] 沒有語詞）。
    // 超綱的語詞會被 gradeFilter 剔掉，剔完沒剩就換下一個候選字。
    for (var i = 0; i < cands.length; i++) {
      var w = gradeFilter(wordsOf(g, cands[i]), ctxObj);
      if (w.length) { return { target: cands[i], words: w, source: '字語詞' }; }
    }
    // 階梯 2：退到變式句庫例句（須經詞典驗證與年級檢查）
    for (var j = 0; j < cands.length; j++) {
      var bw = gradeFilter(bankWords(ctxObj.bank, ctxObj.words, g['字'] || cands, cands[j]), ctxObj);
      if (bw.length) { return { target: cands[j], words: bw, source: '變式句庫' }; }
    }
    // 階梯 3：兩層都沒有。不可假設語詞必然存在。
    return { target: cands[0], words: [], source: '無' };
  }

  /**
   * entries: [{ qid: '4', g: 字組物件 }]，順序即老師的勾選順序
   * opts: {
   *   bank: ZhuVariantBank,
   *   words: 詞典物件（需有 lookup(word)，如 ZhuBopo）——階梯 2 用它驗證抽出的是真詞，
   *          不傳就跳過階梯 2 直接落到「無」。刻意不在模組內取全域變數，
   *          以免瀏覽器載入順序決定行為。
   *   learned: 已學字物件（需有 has(grade, ch)，如 ZhuLearnedChars）
   *   grade: 出卷年級（3/4/5/6）——參考語詞的每個字都須在該年級累計已學字內。
   *          learned 或 grade 缺一就跳過這道檢查。
   *   overrides: { groupKey: 選定的字 }
   * }
   * 回傳 { ok, slots, excluded, conflicts, shortBy }
   */
  function build(entries, opts) {
    opts = opts || {};
    var bank = opts.bank || null;
    var words = opts.words || null;
    var learned = opts.learned || null;
    var grade = (opts.grade === 0) ? 0 : (opts.grade || null);
    var overrides = opts.overrides || {};

    // 1＋2：只留對比字組，並以 canonical key 去重（題號合併進 sourceQuestionIds）
    var order = [];
    var byKey = {};
    (entries || []).forEach(function (e) {
      var g = e.g;
      if (!g || g['類型'] !== '對比字組') { return; }
      var cands = candidates(g);
      if (cands.length < 2) { return; }
      var key = canonicalKey(cands);
      if (!byKey[key]) {
        byKey[key] = { key: key, g: g, cands: cands, qids: [] };
        order.push(key);
      }
      if (byKey[key].qids.indexOf(String(e.qid)) < 0) {
        byKey[key].qids.push(String(e.qid));
      }
    });

    // 3：逐組選 target
    var picked = order.map(function (key) {
      var rec = byKey[key];
      // 豁免名單逐組算：這一組的目標字與干擾字正是要考的字，不受年級限制
      var exempt = {};
      rec.cands.forEach(function (c) { exempt[c] = true; });
      var ctxObj = { bank: bank, words: words, learned: learned, grade: grade, exempt: exempt };
      var p = pickTarget(rec.g, rec.cands, ctxObj, overrides[key]);
      return { key: key, rec: rec, pick: p };
    });

    // 4：target 互異檢查。不同字組可能選到同一個字（種/重 與 重/量 都選「重」），
    //    而格式門禁把重複 target 判為 format.duplicateTarget——
    //    不在這裡擋住，就會送出一份注定失敗的提示詞。
    var byTarget = {};
    picked.forEach(function (p) {
      (byTarget[p.pick.target] = byTarget[p.pick.target] || []).push(p.key);
    });
    var conflicts = [];
    Object.keys(byTarget).forEach(function (ch) {
      if (byTarget[ch].length > 1) {
        conflicts.push({ char: ch, groupKeys: byTarget[ch].slice() });
      }
    });
    if (conflicts.length) {
      return { ok: false, slots: [], excluded: [], conflicts: conflicts, shortBy: 0 };
    }

    // 5：截斷
    var kept = picked.slice(0, MAX_SLOTS);
    var excluded = picked.slice(MAX_SLOTS).map(function (p) {
      return { groupKey: p.key, chars: p.rec.cands.slice() };
    });

    var slots = kept.map(function (p, i) {
      var target = p.pick.target;
      return {
        blank: i + 1,
        groupKey: p.key,
        sourceQuestionIds: p.rec.qids.slice(),
        target: target,
        targetWords: p.pick.words,
        wordSource: p.pick.source,
        distractors: p.rec.cands.filter(function (c) { return c !== target; })
          .map(function (c) { return { char: c, words: wordsOf(p.rec.g, c) }; })
      };
    });

    // 6：下限。注意：slots 一律照實回傳（就算未達門檻），只有 ok/shortBy
    // 反映是否達標——呼叫端才能在「還差幾組」的提示裡照樣列出已選的空格預覽。
    var shortBy = slots.length >= MIN_SLOTS ? 0 : (MIN_SLOTS - slots.length);
    return {
      ok: shortBy === 0,
      slots: slots,
      excluded: excluded,
      conflicts: [],
      shortBy: shortBy
    };
  }

  function charCountRange(n) {
    return { min: n * 25, max: n * 35 };
  }

  var api = {
    build: build,
    canonicalKey: canonicalKey,
    charCountRange: charCountRange,
    MAX_SLOTS: MAX_SLOTS,
    MIN_SLOTS: MIN_SLOTS
  };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  return api;
})();
