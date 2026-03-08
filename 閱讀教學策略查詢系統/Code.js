// Code.gs

// 多頁路由：?page=quiz → quiz.html, ?page=teacher → teacher.html, 預設 → home.html
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';
  var pageMap = {
    'home': { file: 'home', title: '學習扶助閱讀策略小幫手' },
    'index': { file: 'index', title: '閱讀教學策略查詢系統' },
    'calibration': { file: 'calibration_guide', title: '內部協作與資料優化指南' },
    'quiz': { file: 'quiz', title: '閱讀遷移測驗 — 學生作答' },
    'teacher': { file: 'teacher', title: '閱讀遷移測驗 — 老師管理頁' },
    'self_practice': { file: 'student_practice', title: '閱讀測驗練習 — 學生自主' }
  };
  var config = pageMap[page] || pageMap['home'];
  var template = HtmlService.createTemplateFromFile(config.file);
  template.scriptUrl = ScriptApp.getService().getUrl();
  
  return template.evaluate()
    .setTitle(config.title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 修改後的讀取函式：一次讀取多個工作表
function getData() {
  // 使用明確的試算表 ID 確保連線穩定
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');

  // 1. 指定要讀取的工作表名稱清單 (請確認名稱跟下方標籤一模一樣)
  var targetSheets = ['三年級', '四年級', '五年級', '六年級'];

  var combinedData = [];

  // 2. 使用迴圈，一個一個工作表進去抓資料
  targetSheets.forEach(function (sheetName) {
    var sheet = ss.getSheetByName(sheetName);

    // 確保工作表存在才讀取，避免報錯
    if (sheet) {
      // 讀取該工作表所有資料
      var data = sheet.getDataRange().getValues();

      // 確保資料不只一行（避免只有標題沒有內容）
      if (data.length > 1) {
        var headers = data[0]; // 第一列是標題
        var rows = data.slice(1); // 第二列之後是內容

        // 將每一列資料轉換成物件 (Key-Value)
        var sheetData = rows.map(function (row) {
          var obj = {};
          headers.forEach(function (header, index) {
            // 移除標題前後空白，確保對應準確
            var key = String(header).trim();
            obj[key] = row[index];
          });
          return obj;
        });

        // 將處理好的資料合併到大陣列中
        combinedData = combinedData.concat(sheetData);
      }
    }
  });

  // 讀取學習遷移題目（只讀關鍵欄位，避免讀入大量長文字造成執行超時）
  var transferSheet = ss.getSheetByName('學習遷移題目');
  var transferData = [];
  // 只撈首頁需要的欄位，略過可能非常長的「遷移文本內容」
  var TRANSFER_SKIP_COLS = ['遷移文本內容'];
  if (transferSheet) {
    var data = transferSheet.getDataRange().getValues();
    if (data.length > 1) {
      var headers = data[0];
      var rows = data.slice(1);
      transferData = rows.map(function (row) {
        var obj = {};
        headers.forEach(function (header, index) {
          var key = String(header).trim();
          if (TRANSFER_SKIP_COLS.indexOf(key) === -1) {
            obj[key] = row[index];
          }
        });
        return obj;
      });
    }
  }

  // 讀取教材 PDF 連結對應表
  var materialLinks = {};
  var linkSheet = ss.getSheetByName('教材PDF連結');
  if (linkSheet) {
    var linkData = linkSheet.getDataRange().getValues();
    for (var i = 1; i < linkData.length; i++) {
      var fname = String(linkData[i][0] || '');
      var fid = String(linkData[i][1] || '');
      if (!fname || !fid || fname.indexOf('.pdf') === -1) continue;
      if (fname.indexOf(' (1).pdf') !== -1 || fname.indexOf(' (2).pdf') !== -1) continue;
      var m = fname.match(/^(\d+-\d+-\d+)/);
      if (m && !materialLinks[m[1]]) materialLinks[m[1]] = fid;
    }
  }

  // 讀取考古題 PDF 連結對應表
  var examLinks = {};
  var gradeKeyMap = {'\u4e09': 'G3', '\u56db': 'G4', '\u4e94': 'G5', '\u516d': 'G6'};
  var examSheet = ss.getSheetByName('\u8003\u53e4\u984cPDF\u9023\u7d50');
  if (examSheet) {
    var examData = examSheet.getDataRange().getValues();
    for (var i = 1; i < examData.length; i++) {
      var efname = String(examData[i][0] || '');
      var efid   = String(examData[i][1] || '');
      var efolder = String(examData[i][2] || '');
      if (!efname || !efid) continue;
      var ym = efname.match(/^(\d{3})/);
      if (!ym) continue;
      var eyear = ym[1];
      var egrade = '';
      for (var g in gradeKeyMap) {
        if (efolder.indexOf(g) !== -1) { egrade = gradeKeyMap[g]; break; }
      }
      if (!egrade) continue;
      var ekey = eyear + '_' + egrade;
      if (!examLinks[ekey]) examLinks[ekey] = efid;
    }
  }

  // 讀取考古題頁碼對應表
  var examPages = {};
  var pageSheet = ss.getSheetByName('\u8003\u53e4\u984c\u9801\u78bc');
  if (pageSheet) {
    var pageData = pageSheet.getDataRange().getValues();
    for (var i = 1; i < pageData.length; i++) {
      var py = String(pageData[i][0] || '');
      var pg = String(pageData[i][1] || '');
      var pt = String(pageData[i][2] || '');
      var pp = pageData[i][3];
      if (py && pg && pt && pp) examPages[py + '_' + pg + '_' + pt] = Number(pp);
    }
  }

  // 回傳整合後的所有年級資料與遷移資料
  return {
    curriculum: combinedData,
    transfer: transferData,
    materialLinks: materialLinks,
    examLinks: examLinks,
    examPages: examPages
  };
}

// 按需讀取單一文章的遷移文本（供簡報模式使用）
function getTransferText(origTitle) {
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');
  var sheet = ss.getSheetByName('學習遷移題目');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var rows = data.slice(1);

  // 找各欄位索引，避免欄位名稱中有隱藏字元時對應失敗
  var idxTitle    = headers.indexOf('原始案例標題');
  var idxOrigQ    = headers.indexOf('原始題目');
  var idxContent  = headers.indexOf('遷移文本內容');
  var idxQ        = headers.indexOf('遷移題目');
  var idxStrategy = headers.indexOf('教學策略');
  var idxProcess  = headers.indexOf('認知歷程');

  var result = [];
  rows.forEach(function (row) {
    if (String(row[idxTitle] || '').trim() !== String(origTitle).trim()) return;
    result.push({
      '原始案例標題': String(row[idxTitle]    || ''),
      '原始題目':     String(row[idxOrigQ]    || ''),
      '遷移文本內容': String(row[idxContent]  || ''),
      '遷移題目':     String(row[idxQ]        || ''),
      '教學策略':     String(row[idxStrategy] || ''),
      '認知歷程':     String(row[idxProcess]  || '')
    });
  });
  return result;
}

// =============================================
// 一次性文章分段更新（執行後可刪除）
// =============================================

function updateArticleText() {
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');

  var updates = [
    {
      sheet: '三年級',
      titleCol: '文本標題',
      contentCol: '文章全文',
      title: '小光與雜貨店',
      text: '小光每天在雜貨店幫阿婆的忙。有天他在超商看到一臺彈珠遊戲機，非常喜歡，但是沒錢買。所以，小光就趁阿婆不注意時，從抽屜拿了一張鈔票。傍晚，阿婆要關門休息，小光看到她著急的四處翻找，嘴裡念著：「錢怎麼少了？」\n後來，小光走進超商。他盯著遊戲機，想起了阿婆著急的樣子，怎麼也高興不起來。\n從那時候開始，小光就沒去雜貨店了。\n小光生日那天，託人送一塊蛋糕給阿婆。阿婆發現蛋糕底下有張紙條和五百元鈔票。紙條上寫著：\n阿婆：五百元是我拿的，本來想去買遊戲機，現在決定還給您。對不起！小光 敬上\n過了幾天，小光經過雜貨店，發現店門關著。鄰居跟他說：「阿婆搬到兒子家了，她有一封信要給你。」小光打開一看，有張五百元鈔票。信紙上寫著：「小光，祝你生日快樂！」'
    }
  ];

  var results = [];
  updates.forEach(function(u) {
    var sheet = ss.getSheetByName(u.sheet);
    if (!sheet) { results.push(u.title + ': 找不到工作表'); return; }
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var titleIdx = headers.indexOf(u.titleCol);
    var contentIdx = headers.indexOf(u.contentCol);
    if (titleIdx < 0 || contentIdx < 0) { results.push(u.title + ': 找不到欄位'); return; }
    var updated = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][titleIdx]).trim() === u.title) {
        sheet.getRange(i + 1, contentIdx + 1).setValue(u.text);
        updated++;
      }
    }
    results.push(u.title + ': 更新 ' + updated + ' 列');
  });

  return results.join('\n');
}

// =============================================
// 學生測驗系統 API
// =============================================

var SHEET_ID = '1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI';

/**
 * 初始化：自動建立「今日指派」與「作答紀錄」工作表（只需執行一次）
 */
function initSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // 建立「今日指派」
  if (!ss.getSheetByName('今日指派')) {
    var s1 = ss.insertSheet('今日指派');
    s1.appendRow(['文章標題']);
  }

  // 建立「作答紀錄」
  if (!ss.getSheetByName('作答紀錄')) {
    var s2 = ss.insertSheet('作答紀錄');
    s2.appendRow(['時間戳', '座號', '姓名', '文章標題', '題號', '學生答案', '正確答案', '是否正確']);
  }

  return '初始化完成！';
}

/**
 * 老師指派文章（儲存到「今日指派」）
 */
function setAssignment(title) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('今日指派');
  if (!sheet) return { error: '找不到「今日指派」工作表，請先執行 initSheets()' };
  // 清空舊指派並寫入新指派
  sheet.clearContents();
  sheet.appendRow(['文章標題']);
  sheet.appendRow([title]);
  return { success: true, title: title };
}

/**
 * 取得今日指派的遷移文章 + 所有對應的遷移題
 * 今日指派工作表儲存 JSON 格式：{"orig":"...","key":"..."}
 */
function getAssignment() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var assignSheet = ss.getSheetByName('今日指派');
  if (!assignSheet) return { error: '尚未初始化，請老師先執行 initSheets()' };
  var assignData = assignSheet.getDataRange().getValues();
  if (assignData.length < 2 || !assignData[1][0]) {
    return { assigned: false, message: '今天還沒有指派測驗，請等待老師指派。' };
  }
  var stored = String(assignData[1][0]).trim();

  // 解析儲存值（新格式 JSON，相容舊格式字串）
  var assignedOrig = '', assignedKey = '';
  try {
    var parsed = JSON.parse(stored);
    assignedOrig = parsed.orig || '';
    assignedKey = parsed.key || '';
  } catch (e) {
    assignedOrig = stored; // 舊格式相容
  }

  var transferSheet = ss.getSheetByName('學習遷移題目');
  if (!transferSheet) return { error: '找不到學習遷移題目工作表' };
  var data = transferSheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);

  var questions = [];
  var articleText = '';
  var transferTitle = '';

  rows.forEach(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[String(h).trim()] = row[i]; });
    var rowOrig = String(obj['原始案例標題'] || '').trim();
    var rowText = String(obj['遷移文本內容'] || '').trim();
    var rowKey = rowOrig + '|||' + rowText.substring(0, 50);

    // 比對：若有 key 則精確比對，否則用原始標題比對
    var matched = assignedKey ? (rowKey === assignedKey) : (rowOrig === assignedOrig);
    if (!matched) return;

    // 記錄文章全文（只記一次）
    if (!articleText && rowText) {
      articleText = rowText;
      // 提取遷移文章標題（格式：【1. xxx】）
      var titleMatch = rowText.match(/^[【\[](.+?)[】\]]/);
      transferTitle = titleMatch ? titleMatch[0] : rowOrig;
    }

    var rawQ = String(obj['修補後題目'] || obj['遷移題目'] || '').trim();
    if (!rawQ) return;
    var answerMatch = rawQ.match(/答案[：:][（(]([ABCD1234])[)）]/);
    var answer = answerMatch ? answerMatch[1] : '';
    questions.push({ id: questions.length + 1, question: rawQ, answer: answer });
  });

  return {
    assigned: true,
    orig: assignedOrig,
    title: transferTitle || assignedOrig,
    articleText: articleText,
    questions: questions
  };
}

/**
 * 學生送出作答：批改並儲存（覆蓋當天同一學生同一文章的舊紀錄）
 * payload: { seatNum, name, title, answers: [{id, selected}] }
 */
function submitAnswers(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('作答紀錄');
  if (!sheet) return { error: '找不到作答紀錄工作表' };

  var seatNum = String(payload.seatNum);
  var name = String(payload.name);
  var title = String(payload.title);
  var answers = payload.answers; // [{id, question, selected, correct}]
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  // 刪除同學生今日同文章的舊紀錄
  var data = sheet.getDataRange().getValues();
  var rowsToDelete = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), 'Asia/Taipei', 'yyyy-MM-dd') : '';
    if (String(data[i][1]) === seatNum && rowDate === today && String(data[i][3]) === title) {
      rowsToDelete.push(i + 1);
    }
  }
  rowsToDelete.forEach(function (r) { sheet.deleteRow(r); });

  // 批改並寫入新紀錄
  var results = [];
  var timestamp = new Date();
  answers.forEach(function (a) {
    var isCorrect = String(a.selected).trim().toUpperCase() === String(a.correct).trim().toUpperCase();
    sheet.appendRow([timestamp, seatNum, name, title, a.id, a.selected, a.correct, isCorrect ? '✅' : '❌']);
    results.push({
      id: a.id,
      question: a.question,
      selected: a.selected,
      correct: a.correct,
      isCorrect: isCorrect
    });
  });

  var score = results.filter(function (r) { return r.isCorrect; }).length;
  return { success: true, score: score, total: results.length, results: results };
}

/**
 * 老師讀取作答紀錄
 * dateFilter: 'today' 或 'all'
 */
function getResults(dateFilter) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('作答紀錄');
  if (!sheet) return { error: '找不到作答紀錄工作表' };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { records: [] };

  var headers = data[0];
  var rows = data.slice(1);
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  var records = rows.map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });

  if (dateFilter === 'today') {
    records = records.filter(function (r) {
      var d = r['時間戳'] ? Utilities.formatDate(new Date(r['時間戳']), 'Asia/Taipei', 'yyyy-MM-dd') : '';
      return d === today;
    });
  }

  return { records: records };
}

/**
 * 老師端取得兩層文章清單
 * 回傳：{ originals: [ { orig, transfers: [{label, key}] } ] }
 */
function getArticleList() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('學習遷移題目');
  if (!sheet) return { originals: [] };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);

  var origMap = {}; // orig -> [ {label, key} ]
  var seenKeys = {};

  rows.forEach(function (row) {
    var obj = {};
    headers.forEach(function (h, i) { obj[String(h).trim()] = row[i]; });
    var orig = String(obj['原始案例標題'] || '').trim();
    var transferText = String(obj['遷移文本內容'] || '').trim();
    if (!orig || !transferText) return;

    var key = orig + '|||' + transferText.substring(0, 50);
    if (seenKeys[key]) return;
    seenKeys[key] = true;

    // 提取遷移文章標題（格式：【1. xxx】）
    var titleMatch = transferText.match(/^[【\[](.+?)[】\]]/);
    var label = titleMatch ? titleMatch[0] : transferText.substring(0, 20) + '…';

    if (!origMap[orig]) origMap[orig] = [];
    origMap[orig].push({ label: label, key: key });
  });

  var originals = Object.keys(origMap).map(function (orig) {
    return { orig: orig, transfers: origMap[orig] };
  });
  return { originals: originals };
}

function getMaterialLinks() {
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');
  var sheet = ss.getSheetByName('教材PDF連結');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0] || '');
    var fileId = String(data[i][1] || '');
    if (!name || !fileId || name.indexOf('.pdf') === -1) continue;
    var match = name.match(/^(\d+-\d+-\d+)/);
    if (match && !map[match[1]]) {
      map[match[1]] = fileId;
    }
  }
  return map;
}

function listPdfFiles() {
  var folderId = '1Pcn4rT1z9ECEOAcG51jI-vf1a-SwNElg';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');
  var sheetName = '教材PDF連結';
  var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.appendRow(['檔案名稱', 'Drive檔案ID', '教材編號（請手動填）']);
  while (files.hasNext()) {
    var f = files.next();
    sheet.appendRow([f.getName(), f.getId(), '']);
  }
  SpreadsheetApp.flush();
  Logger.log('完成，請到試算表「教材PDF連結」工作表查看結果');
}

function updateTextbookHyperlinks() {
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');
  
  // 1. 建立教材對應表
  var linkSheet = ss.getSheetByName('教材PDF連結');
  if (!linkSheet) return;
  var linkData = linkSheet.getDataRange().getValues();
  var urlMap = {};
  for (var i = 1; i < linkData.length; i++) {
    var fname = String(linkData[i][0] || '');
    var fid = String(linkData[i][1] || '');
    var manualCode = String(linkData[i][2] || '').trim();
    if (!fname || !fid || fname.indexOf('.pdf') === -1) continue;
    
    var url = 'https://drive.google.com/file/d/' + fid + '/view';
    if (manualCode) {
      if (!urlMap[manualCode]) urlMap[manualCode] = url;
    } else {
      var match = fname.match(/^(\d+-\d+-\d+)/);
      if (match && !urlMap[match[1]]) urlMap[match[1]] = url;
    }
  }

  // 2. 寫入各年級
  var targetSheets = ['三年級', '四年級', '五年級', '六年級'];
  targetSheets.forEach(function(sName) {
    var sheet = ss.getSheetByName(sName);
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    var headers = data[0];
    var targetColIdx = -1;
    for (var j = 0; j < headers.length; j++) {
      if (String(headers[j]).trim() === '對應教材編號') {
        targetColIdx = j;
        break;
      }
    }
    if (targetColIdx === -1) return;
    // 預先將整欄設為純文字，避免後續老師輸入時被轉成日期
    sheet.getRange(2, targetColIdx + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    
    // 收集所有需要加上超連結的儲存格
    for (var r = 1; r < data.length; r++) {
      var rawValue = data[r][targetColIdx];
      var code = String(rawValue).trim();
      
      // 如果被 Google Sheets 自動轉成了日期物件 (例如 3-2-12 被轉成 2003-02-12)
      if (rawValue instanceof Date) {
        var y = rawValue.getFullYear() % 100;
        var m = rawValue.getMonth() + 1;
        var d = rawValue.getDate();
        code = y + '-' + m + '-' + d;
      } else if (code.match(/^2\d{3}[-\/]\d{1,2}[-\/]\d{1,2}$/)) {
        // 如果拿到的是字串格式的誤判日期，例如 "2003-2-15" 或 "2003/2/15"
        var parts = code.split(/[-\/]/);
        var y = parseInt(parts[0], 10) % 100;
        code = y + '-' + parts[1] + '-' + parts[2];
      }
      
      var cell = sheet.getRange(r + 1, targetColIdx + 1);
      
      // 無論有沒有找到對應網址，都把這格修正後的純文字 code 寫回去，把日期救回來
      if (rawValue instanceof Date || code !== String(rawValue).trim()) {
         cell.setValue(code);
      }
      
      // 檢查是否已經是超連結公式
      var existingFormula = cell.getFormula();
      if (existingFormula && existingFormula.toUpperCase().indexOf('HYPERLINK') !== -1) {
         continue; // 已經有超連結公式，跳過
      }
      // 避免覆蓋已經帶有 Rich Text Link 的內容
      var existingRtv = cell.getRichTextValue();
      if (existingRtv && existingRtv.getLinkUrl()) {
         continue; // 已經有富文本超連結，跳過
      }
      
      if (code && urlMap[code]) {
         // 使用 HYPERLINK 公式，並用雙引號包住字串，防止再次被轉為日期
         cell.setFormula('=HYPERLINK("' + urlMap[code] + '", "' + code + '")');
      }
    }
  });
  Logger.log('教材超連結更新完成！');
}

function updateExamHyperlinks() {
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');
  
  // 1. 建立考古題對應表 (Year_Grade -> File ID)
  var examLinks = {};
  var gradeKeyMap = {'\u4e09': 'G3', '\u56db': 'G4', '\u4e94': 'G5', '\u516d': 'G6'};
  var examSheet = ss.getSheetByName('\u8003\u53e4\u984cPDF\u9023\u7d50');
  if (!examSheet) return;
  
  var examData = examSheet.getDataRange().getValues();
  for (var i = 1; i < examData.length; i++) {
    var efname = String(examData[i][0] || '');
    var efid   = String(examData[i][1] || '');
    var efolder = String(examData[i][2] || '');
    if (!efname || !efid) continue;
    var ym = efname.match(/^(\d{3})/);
    if (!ym) continue;
    var eyear = ym[1];
    var egrade = '';
    for (var g in gradeKeyMap) {
      if (efolder.indexOf(g) !== -1) { egrade = gradeKeyMap[g]; break; }
    }
    if (!egrade) continue;
    var ekey = eyear + '_' + egrade;
    if (!examLinks[ekey]) examLinks[ekey] = 'https://drive.google.com/file/d/' + efid + '/view';
  }

  // 2. 寫入各年級
  var targetSheets = ['三年級', '四年級', '五年級', '六年級'];
  targetSheets.forEach(function(sName) {
    var sheet = ss.getSheetByName(sName);
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    var headers = data[0];
    var yearColIdx = -1;
    var gradeColIdx = -1;
    for (var j = 0; j < headers.length; j++) {
      var h = String(headers[j]).trim();
      if (h === '年度') yearColIdx = j;
      if (h === '年級') gradeColIdx = j;
    }
    if (yearColIdx === -1 || gradeColIdx === -1) return;
    
    // 預先將整欄設為純文字，避免自動轉型問題
    sheet.getRange(2, yearColIdx + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    sheet.getRange(2, gradeColIdx + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    
    for (var r = 1; r < data.length; r++) {
      var yearCode = String(data[r][yearColIdx]).trim();
      var gradeCode = String(data[r][gradeColIdx]).trim();
      
      var ekey = yearCode + '_' + gradeCode;
      var url = examLinks[ekey];
      
      if (url) {
        var yearCell = sheet.getRange(r + 1, yearColIdx + 1);
        var yearFormula = yearCell.getFormula();
        if (!yearFormula || yearFormula.toUpperCase().indexOf('HYPERLINK') === -1) {
            yearCell.setFormula('=HYPERLINK("' + url + '", "' + yearCode + '")');
        }
        
        var gradeCell = sheet.getRange(r + 1, gradeColIdx + 1);
        var gradeFormula = gradeCell.getFormula();
        if (!gradeFormula || gradeFormula.toUpperCase().indexOf('HYPERLINK') === -1) {
            gradeCell.setFormula('=HYPERLINK("' + url + '", "' + gradeCode + '")');
        }
      }
    }
  });
  Logger.log('考古題超連結更新完成！');
}

function listExamPdfs() {
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');
  var sheet = ss.getSheetByName('\u8003\u53e4\u984cPDF\u9023\u7d50') || ss.insertSheet('\u8003\u53e4\u984cPDF\u9023\u7d50');
  sheet.clearContents();
  sheet.appendRow(['\u6a94\u6848\u540d\u7a31', 'Drive\u6a94\u6848ID', '\u8cc7\u6599\u5937\u540d\u7a31']);
  var rows = [];
  function scanFolder(folder) {
    var folderName = folder.getName();
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      rows.push([f.getName(), f.getId(), folderName]);
    }
    var subs = folder.getFolders();
    while (subs.hasNext()) { scanFolder(subs.next()); }
  }
  scanFolder(DriveApp.getFolderById('15lsaGiVsuj6QGQyiYooacRwRvbZuX-sM'));
  rows.forEach(function(r) { sheet.appendRow(r); });
  SpreadsheetApp.flush();
  Logger.log('\u5171\u627e\u5230 ' + rows.length + ' \u500b\u6a94\u6848');
}

