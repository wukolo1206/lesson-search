// zhu-passage-prompt.js — 產給外部 AI 的短文提示詞
//
// 本模組無外部依賴：參考語詞在 slot 裡已由 zhu-passage-slots.js 備妥。
// 四條硬規則與 zhu-passage-checks.js 的四道門禁一一對應，用詞要一致——
// 提示詞說的與門禁擋的不同調，老師會反覆修卻修不過。
var ZhuPassagePrompt = (function () {
  'use strict';

  function wordList(words) {
    return (words && words.length) ? words.join('、') : '';
  }

  function slotLine(s) {
    var line = '  ' + s.blank + '. 挖「' + s.target + '」';
    var tw = wordList(s.targetWords);
    line += tw ? ('（參考語詞：' + tw + '）')
               : '（無參考語詞，請自行造出該年級看得懂的詞）';
    var ds = (s.distractors || []).map(function (d) {
      var w = wordList(d.words);
      return d.char + (w ? '（' + w + '）' : '');
    });
    if (ds.length) { line += '　易混淆、必須避開的字：' + ds.join('、'); }
    return line;
  }

  // 規則 3 的例子必須取自當次 slots：寫死一組無關的字（早期版本用「導/道」），
  // 外部 AI 會把它當成本次要避開的具體案例，或根本沒意識到要套用到自己那幾組。
  // 挑第一個「干擾字帶語詞」的 slot——那種例子最貼近規則本意（真的換字就成詞）。
  function ruleThreeExample(slots) {
    for (var i = 0; i < slots.length; i++) {
      var s = slots[i];
      var ds = s.distractors || [];
      for (var j = 0; j < ds.length; j++) {
        if (ds[j].words && ds[j].words.length) {
          // 例子本身含一個干擾字組成的詞，與規則 2 咬到——不標明的話
          // AI 會困惑（你自己不是說不能出現？），更糟的是直接拿去當素材寫進短文。
          return ['例如挖「' + s.target + '」的地方若換成「' + ds[j].char
            + '」會變成「' + ds[j].words[0] + '」，那個位置就不能用。',
            '（這個詞只是拿來說明，不要寫進短文。）'];
        }
      }
    }
    // 整批都沒有干擾字語詞時（實測常見），寧可純敘述也不要舉無關的字。
    // 敘述式沒有具體詞，不必加免責句。
    return ['例如某個空的答案是甲字、易混淆字是乙字，'
      + '若該處填乙字也能組成一個常見的詞，那個位置就不能用，請換一個位置或換個說法。'];
  }

  function build(slots, grade) {
    slots = slots || [];
    var n = slots.length;
    var min = n * 25, max = n * 35;
    var targets = slots.map(function (s) { return s.target; }).join('、');

    var lines = [];
    lines.push('請幫我寫一篇國小' + grade + '程度的情境短文，用來出「填國字」練習題。');
    lines.push('');
    lines.push('讀者是學習扶助的學生，閱讀能力比同年級落後約 2～3 個年級，請寫得淺白。');
    lines.push('');
    lines.push('【要挖空的字】共 ' + n + ' 個：' + targets);
    slots.forEach(function (s) { lines.push(slotLine(s)); });
    lines.push('');
    lines.push('【硬性規則】以下每一條都會被程式檢查，違反就要重寫：');
    lines.push('1. 每個目標字只挖一次，挖空的位置寫成 {}（半形大括號，中間不留空白）。');
    lines.push('2. 上面列出的字不可以出現在挖空以外的任何位置，整篇都算，包括其他句子。');
    lines.push('　 目標字：除了各自那唯一一個挖空處，其餘任何地方都不能再出現；'
      + '即使是別的意思、別的詞，只要含這個字就算違規。');
    lines.push('　 易混淆字：整篇任何地方都不能出現，它只是拿來對照的。');
    lines.push('3. 每個挖空處，換成該組的易混淆字之後不可以也成詞。');
    ruleThreeExample(slots).forEach(function (t) { lines.push('　 ' + t); });
    // 規則 4 刻意不寫成硬規則：AI 手上沒有字表，「不可超出」它執行不了，
    // 只會亂猜或空口保證。真正的精準修正在 buildFix()——門禁會指名超綱字，
    // 而「把指定的幾個字換掉」是 AI 做得到的指令。
    lines.push('4. 請盡量只用國小' + grade + '以前學過的字，寫得越淺白越好。');
    lines.push('　 這一項我會用' + grade + '累計已學字表檢查（上面的目標字本身除外），'
      + '如果有超出的字，我會把它們列出來請你替換——');
    lines.push('　 所以先寫得淺白一點，可以省下來回修改的次數。');
    lines.push('');
    lines.push('【篇幅與結構】');
    lines.push('* ' + min + '～' + max + ' 個中文字（標點與 {} 不算）。');
    lines.push('* 至少 3 句，要有起因、經過、結果，是一個完整的小故事，'
      + '不要只是幾句描寫拼起來。');
    lines.push('* 情境要自足，不可以依賴任何課文背景知識。');
    lines.push('');
    lines.push('【輸出格式】請完全照這個格式回覆，不要加任何額外說明：');
    lines.push('');
    lines.push('【短文】');
    lines.push('（第一段）');
    lines.push('');
    lines.push('（第二段，若有）');
    lines.push('');
    lines.push('【答案】' + targets);
    lines.push('');
    lines.push('答案請照挖空出現的先後順序寫，並且必須剛好是上面指定的 ' + n + ' 個字。');
    return lines.join('\n');
  }

  function buildFix(problems, previous) {
    var lines = [];
    lines.push('你剛才給的短文沒有通過檢查，請修正以下問題：');
    lines.push('');
    (problems || []).forEach(function (p, i) {
      var where = [];
      if (p.paragraph) { where.push('第 ' + p.paragraph + ' 段'); }
      if (p.sentence) { where.push('第 ' + p.sentence + ' 句'); }
      if (p.blank) { where.push('第 ' + p.blank + ' 個空'); }
      lines.push((i + 1) + '. ' + (where.length ? where.join('') + '：' : '') + p.message);
      // 只複述 message 等於把判斷丟回給 AI。要給可執行的動作，
      // 而且動作依問題類型而不同：超綱字是「換字」，洩題是「改寫該處」。
      var cs = (p.chars && p.chars.length) ? p.chars.join('、') : '';
      if (cs) {
        var code = String(p.code || '');
        if (code.indexOf('difficulty') === 0) {
          lines.push('   請把這些字換掉：' + cs + '（改用更常見、更好懂的說法）。');
        } else if (code.indexOf('leak') === 0) {
          lines.push('   請把該處改寫掉，讓這些字不再出現：' + cs + '。');
        } else {
          lines.push('   牽涉到的字：' + cs + '。');
        }
      }
    });
    lines.push('');
    lines.push('請只修正這些地方，其餘保持不變——不要重寫整篇、不要換情境、'
      + '不要改動挖空的位置與答案字。');
    lines.push('回覆格式與上次相同（【短文】…【答案】…）。');
    lines.push('');
    lines.push('以下是你上次的內容：');
    lines.push('');
    lines.push(previous || '');
    return lines.join('\n');
  }

  var api = { build: build, buildFix: buildFix };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; }
  return api;
})();
