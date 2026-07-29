// zhu-passage-checks.js — P5「短文填國字」的五道品質門禁
//
// 呼叫順序固定：format → leak → unique → answer → difficulty（見 check()）。
// 規則沿用 scripts/variant_gates.py，差別是套用於整篇短文而非單句。
// 生成與把關刻意分開：短文由外部 AI（ChatGPT 等）貼回，本模組只負責檢查，
// 是純函式、可重跑、可測試，不是黑箱。
//
// 一律回結構化問題物件，不回組好的中文句子：跨語言一致測試（P5 下一步的
// Python 鏡像）、修正提示詞、畫面定位三者共用同一份資料。若比對中文訊息，
// 任何措辭調整都會弄壞測試。
var ZhuPassageGates = (function () {
  'use strict';

  var BLANK = '{}';
  var CJK_RE = /[一-鿿]/g;

  // 句長上限：學扶生閱讀能力普遍落後 2～3 個年級，句子太長會變成考閱讀而非考字
  var MAX_LEN = { '三年級': 25, '四年級': 25, '五年級': 35, '六年級': 35 };
  var MAX_CLAUSES = 3;
  var MIN_SENTENCES = 3;

  // 生字表只收「課本生字」，不含標點與常見功能字；這些不該被當成超綱
  var ALWAYS_OK = '的了是在有和就不人我你他她它們這那個之與或而且但也還很都才把被讓從對到於為以及並';

  var GRADE_NUM = { '三年級': 3, '四年級': 4, '五年級': 5, '六年級': 6 };

  // ------------------------------------------------------------------ 小工具

  function cjkOnly(s) {
    var m = String(s === null || s === undefined ? '' : s).match(CJK_RE);
    return m ? m.join('') : '';
  }

  function uniq(list) {
    var seen = {}, out = [];
    (list || []).forEach(function (x) {
      if (x && !seen[x]) { seen[x] = true; out.push(x); }
    });
    return out;
  }

  function plainText(para) {
    return String(para === null || para === undefined ? '' : para).split(BLANK).join('');
  }

  function countBlanks(para) {
    var s = String(para === null || para === undefined ? '' : para);
    var n = 0, idx = s.indexOf(BLANK);
    while (idx >= 0) { n++; idx = s.indexOf(BLANK, idx + BLANK.length); }
    return n;
  }

  // 逐段切句（以「。！？」），保留段落/句子索引，供難度門禁定位訊息。
  function sentencesOf(plain) {
    return plain.split(/[。！？]/).filter(function (s) { return s.trim().length > 0; });
  }

  // 全篇的空格位置：以「{}」切開每一段，第 j 個空的「前文」是第 j 段（同段內
  // 前一個空之後的文字），「後文」是第 j+1 段（下一個空之前的文字）。
  // 空格編號跨段連續累加，須與 ctx.slots 的 blank 順序一致（上游依老師勾選
  // 順序、段落先後產生，兩者順序理當相同）。
  function locateBlanks(paragraphs) {
    var out = [], counter = 0;
    (paragraphs || []).forEach(function (para, pIdx) {
      var parts = String(para === null || para === undefined ? '' : para).split(BLANK);
      for (var j = 0; j < parts.length - 1; j++) {
        counter++;
        out.push({ blank: counter, paragraph: pIdx + 1, before: parts[j], after: parts[j + 1] });
      }
    });
    return out;
  }

  function problem(code, extra) {
    var base = {
      code: code,
      paragraph: null, sentence: null, blank: null,
      chars: [],
      severity: severityOf(code),
      message: ''
    };
    for (var k in extra) { if (Object.prototype.hasOwnProperty.call(extra, k)) { base[k] = extra[k]; } }
    return base;
  }

  function severityOf(code) {
    if (code.indexOf('parse.') === 0) { return 'blocking'; }
    if (code.indexOf('format.') === 0) { return 'blocking'; }
    if (code.indexOf('leak.') === 0) { return 'blocking'; }
    if (code.indexOf('unique.') === 0) { return 'confirm'; }
    if (code.indexOf('answer.') === 0) { return 'confirm'; }
    if (code.indexOf('difficulty.') === 0) { return 'overridable'; }
    return 'blocking';
  }

  // ------------------------------------------------------------------ 格式

  function gateFormat(doc, ctx) {
    doc = doc || {};
    ctx = ctx || {};
    var paragraphs = doc['段落'] || [];
    var answers = doc['答案'] || [];
    var slots = ctx.slots || [];
    var grade = ctx.grade;
    var problems = [];

    // 內部不變量：slots 的 blank 必須是 1..n 的連續遞增序列。
    // gateUnique 用 locs[s.blank - 1] 取上下文，這個假設一旦破了，會安靜地
    // 從錯的空格取前後文——不當機、不報錯，照樣說「唯一性通過」。
    // format.targetMismatch 補不了：它只比 target 序列，blank 被打亂而
    // targets 剛好對得上就抓不到。正常情況永遠不會觸發；一旦觸發代表
    // gateUnique 的取樣不可信，必須硬擋。
    for (var si = 0; si < slots.length; si++) {
      if (slots[si].blank !== si + 1) {
        problems.push(problem('format.slotOrder', {
          message: '空位編號不連續或順序錯亂（第 ' + (si + 1) + ' 個 slot 的 blank 是 ' +
            slots[si].blank + '），這是程式內部問題，請回報'
        }));
        break;
      }
    }

    var totalBlanks = 0;
    paragraphs.forEach(function (p) { totalBlanks += countBlanks(p); });

    if (totalBlanks !== answers.length) {
      problems.push(problem('format.countMismatch', {
        message: '空格 ' + totalBlanks + ' 個但答案 ' + answers.length + ' 個'
      }));
    } else {
      var targets = slots.map(function (s) { return s.target; });
      var mismatch = targets.length !== answers.length;
      if (!mismatch) {
        for (var i = 0; i < answers.length; i++) {
          if (answers[i] !== targets[i]) { mismatch = true; break; }
        }
      }
      if (mismatch) {
        problems.push(problem('format.targetMismatch', {
          chars: answers.slice(),
          message: '答案陣列與 slots 的目標字序列不符（含順序）'
        }));
      }
    }

    var seenAns = {}, dupChars = [];
    answers.forEach(function (a) {
      if (seenAns[a]) { if (dupChars.indexOf(a) < 0) { dupChars.push(a); } }
      seenAns[a] = true;
    });
    if (dupChars.length) {
      problems.push(problem('format.duplicateTarget', {
        chars: dupChars,
        message: '同一目標字被挖了兩次：' + dupChars.join('、')
      }));
    }

    var plainAll = paragraphs.map(plainText).join('');
    var totalCJK = cjkOnly(plainAll).length;
    var n = slots.length || answers.length || totalBlanks;
    if (n > 0) {
      var minLen = n * 25, maxLenLimit = n * 35;
      if (totalCJK < minLen || totalCJK > maxLenLimit) {
        problems.push(problem('format.length', {
          message: '篇幅 ' + totalCJK + ' 字，不在 ' + minLen + '～' + maxLenLimit + ' 字範圍（n=' + n + '）'
        }));
      }
    }

    var totalSentences = 0;
    paragraphs.forEach(function (p) { totalSentences += sentencesOf(plainText(p)).length; });
    if (totalSentences < MIN_SENTENCES) {
      problems.push(problem('format.length', {
        message: '句數不足：' + totalSentences + ' 句（至少 ' + MIN_SENTENCES + ' 句）'
      }));
    }

    return problems;
  }

  // ------------------------------------------------------------------ 洩題

  // 逐段逐句掃描，一句一個問題——不是一個 slot 一個問題。
  //
  // 舊版逐 slot 掃描時，paragraph 取自「第一個含該字的段落」、blank 取自
  // slot 編號，兩者來自不同位置。實測產出「第 3 段第 1 個空：「到」出現在
  // 挖空以外」——但「到」的空格在第 1 段，洩題點在第 3 段第 8 句，
  // 老師照著訊息去看第 3 段第一個空，那裡是乾淨的 {}，什麼也找不到。
  //
  // 同一個字在三句裡出現就回報三次，老師才知道要改幾個地方。
  // blank 一律 null：洩題是「某句有問題」，不是「某個空有問題」，
  // blank 欄位留給 unique.ambiguous（那個確實在問第幾個空）。
  function gateLeak(doc, ctx) {
    doc = doc || {};
    ctx = ctx || {};
    var paragraphs = doc['段落'] || [];
    var slots = ctx.slots || [];
    var problems = [];

    // 全部 slots 的目標字＋干擾字合起來當作監看集合
    var watch = [];
    slots.forEach(function (s) {
      watch.push(s.target);
      (s.distractors || []).forEach(function (d) { watch.push(d.char); });
    });
    watch = uniq(watch).filter(function (ch) { return !!ch; });
    if (!watch.length) { return problems; }

    paragraphs.forEach(function (para, pIdx) {
      var sentences = sentencesOf(plainText(para));
      sentences.forEach(function (sent, sIdx) {
        var hit = watch.filter(function (ch) { return sent.indexOf(ch) >= 0; });
        if (!hit.length) { return; }
        hit = hit.slice().sort();
        problems.push(problem('leak.inText', {
          paragraph: pIdx + 1,
          sentence: sIdx + 1,
          blank: null,
          chars: hit,
          message: '「' + hit.join('、') + '」出現在挖空以外的位置'
        }));
      });
    });

    return problems;
  }

  // ------------------------------------------------------------------ 唯一性

  function gateUnique(doc, ctx) {
    doc = doc || {};
    ctx = ctx || {};
    var paragraphs = doc['段落'] || [];
    var slots = ctx.slots || [];
    var bopo = ctx.bopo;
    var problems = [];
    if (!bopo || typeof bopo.lookup !== 'function') { return problems; }

    var locs = locateBlanks(paragraphs);

    slots.forEach(function (s) {
      var loc = locs[s.blank - 1];
      if (!loc) { return; }
      var beforeCjk = cjkOnly(loc.before);
      var afterCjk = cjkOnly(loc.after);
      var before = beforeCjk.slice(-2);
      var after = afterCjk.slice(0, 2);
      var others = (s.distractors || []).map(function (d) { return d.char; });
      var ambiguous = [];
      var hitChars = [];

      others.forEach(function (other) {
        if (!other || other === s.target) { return; }
        for (var nb = 0; nb <= 2; nb++) {
          for (var na = 0; na <= 2; na++) {
            if (nb === 0 && na === 0) { continue; }
            var w = (nb ? before.slice(before.length - nb) : '') + other + (na ? after.slice(0, na) : '');
            if (w.length >= 2 && bopo.lookup(w).length) {
              if (ambiguous.indexOf(w) < 0) { ambiguous.push(w); }
              if (hitChars.indexOf(other) < 0) { hitChars.push(other); }
            }
          }
        }
      });

      if (ambiguous.length) {
        problems.push(problem('unique.ambiguous', {
          paragraph: loc.paragraph,
          blank: s.blank,
          chars: hitChars,
          message: '代入干擾字後也成詞：' + ambiguous.join('、')
        }));
      }
    });

    return problems;
  }

  // ------------------------------------------------------------------ 答案成詞

  // gateUnique 的鏡像：那道問「代入干擾字會不會**也**成詞」（防兩個答案都對），
  // 這道問「代入目標字到底**成不成**詞」。兩件事對稱，過去只做了一半——
  // AI 因此能造出避開所有陷阱、卻沒真的組成詞的空格：參考語詞給「發明」，
  // 但它寫成「自己{}出好東西」，填回去是「自己明出」，「明」在那裡什麼都不是。
  //
  // severity 用 confirm 而非 blocking：詞庫 21,905 詞不可能涵蓋所有合理搭配，
  // 硬擋會誤殺；而這種問題老師一眼就看得出來，判斷權給老師比較合適。
  function gateAnswer(doc, ctx) {
    doc = doc || {};
    ctx = ctx || {};
    var paragraphs = doc['段落'] || [];
    var slots = ctx.slots || [];
    var bopo = ctx.bopo;
    var problems = [];
    if (!bopo || typeof bopo.lookup !== 'function') { return problems; }

    var locs = locateBlanks(paragraphs);

    slots.forEach(function (s) {
      var loc = locs[s.blank - 1];
      if (!loc || !s.target) { return; }
      var before = cjkOnly(loc.before).slice(-2);
      var after = cjkOnly(loc.after).slice(0, 2);

      // 前後都沒有中文字就湊不出任何候選詞。這時回報「不成詞」是誣賴，
      // 不是判斷——無從判斷就不報。
      if (!before && !after) { return; }

      for (var nb = 0; nb <= 2; nb++) {
        for (var na = 0; na <= 2; na++) {
          if (nb === 0 && na === 0) { continue; }
          var w = (nb ? before.slice(before.length - nb) : '') + s.target + (na ? after.slice(0, na) : '');
          if (w.length >= 2 && bopo.lookup(w).length) { return; }   // 任何一個成詞就過關
        }
      }

      problems.push(problem('answer.notAWord', {
        paragraph: loc.paragraph,
        blank: s.blank,
        chars: [s.target],
        message: '填入「' + s.target + '」後不成詞（' + before + s.target + after +
          '）——這個位置可能不適合挖這個字'
      }));
    });

    return problems;
  }

  // ------------------------------------------------------------------ 難度

  function gateDifficulty(doc, ctx) {
    doc = doc || {};
    ctx = ctx || {};
    var paragraphs = doc['段落'] || [];
    var slots = ctx.slots || [];
    var grade = ctx.grade;
    var learned = ctx.learned;
    var problems = [];

    var plainByPara = paragraphs.map(plainText);
    var fullText = cjkOnly(plainByPara.join(''));

    // 考點字與答案字一律豁免超綱檢查——那正是要教的字；干擾字也豁免（同一組要對比的字）
    var exempt = {};
    slots.forEach(function (s) {
      exempt[s.target] = true;
      (s.distractors || []).forEach(function (d) { exempt[d.char] = true; });
    });
    (doc['答案'] || []).forEach(function (a) {
      if (typeof a === 'string') {
        for (var i = 0; i < a.length; i++) { exempt[a.charAt(i)] = true; }
      }
    });

    // 超綱字逐段回報。整篇合成一筆時，老師拿到「達、隊、頁」卻要在 278 字裡
    // 自己找——跟洩題只報段落不報句子是同一種痛：判斷沒錯，但老師找不到。
    //
    // 段落層級而非句子層級：超綱字往往一句裡好幾個，切到句子會把
    // 「達、隊、頁」拆成三筆散在不同句，面板變吵；段落層級足以定位。
    var gnum = GRADE_NUM[grade];
    if (gnum && learned && typeof learned.has === 'function') {
      paragraphs.forEach(function (para, pIdx) {
        var text = cjkOnly(plainByPara[pIdx]);
        var seen = {}, over = [];
        for (var i = 0; i < text.length; i++) {
          var ch = text.charAt(i);
          if (seen[ch]) { continue; }
          seen[ch] = true;
          if (exempt[ch]) { continue; }
          if (ALWAYS_OK.indexOf(ch) >= 0) { continue; }
          if (learned.has(gnum, ch)) { continue; }
          over.push(ch);
        }
        if (over.length) {
          over.sort();
          problems.push(problem('difficulty.unlearned', {
            paragraph: pIdx + 1,
            chars: over,
            message: '超出' + grade + '累計已學字：' + over.join('')
          }));
        }
      });
    }

    var maxLen = MAX_LEN[grade] || 30;
    paragraphs.forEach(function (para, pIdx) {
      var plain = plainByPara[pIdx];
      var sentences = sentencesOf(plain);
      sentences.forEach(function (sent, sIdx) {
        var len = cjkOnly(sent).length;
        if (len > maxLen) {
          problems.push(problem('difficulty.sentenceTooLong', {
            paragraph: pIdx + 1,
            sentence: sIdx + 1,
            message: '句子過長：' + len + ' 字（' + grade + '上限 ' + maxLen + '）'
          }));
        }
        var clauses = sent.split(/[，、；]/).filter(function (p) { return p.trim().length > 0; });
        if (clauses.length > MAX_CLAUSES) {
          problems.push(problem('difficulty.tooManyClauses', {
            paragraph: pIdx + 1,
            sentence: sIdx + 1,
            message: '分句過多：' + clauses.length + ' 個（上限 ' + MAX_CLAUSES + '）'
          }));
        }
      });
    });

    return problems;
  }

  // ------------------------------------------------------------------ 總入口

  function check(doc, ctx) {
    doc = doc || {};
    ctx = ctx || {};
    var problems = [];
    problems = problems.concat(gateFormat(doc, ctx));
    problems = problems.concat(gateLeak(doc, ctx));
    problems = problems.concat(gateUnique(doc, ctx));
    problems = problems.concat(gateAnswer(doc, ctx));
    problems = problems.concat(gateDifficulty(doc, ctx));
    return problems;
  }

  var api = {
    check: check,
    severityOf: severityOf,
    gateFormat: gateFormat,
    gateLeak: gateLeak,
    gateUnique: gateUnique,
    gateAnswer: gateAnswer,
    gateDifficulty: gateDifficulty
  };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  return api;
})();
