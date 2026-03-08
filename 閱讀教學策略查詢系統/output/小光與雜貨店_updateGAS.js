// 一次性文章分段更新 — 小光與雜貨店（執行後可刪除）
function updateArticleText() {
  var ss = SpreadsheetApp.openById('1IDu-J5luPJsKA5O7UPtiXhnueQ_JNZJ2ccdnldCCpdI');
  var updates = [{
    sheet: '三年級',
    titleCol: '文本標題',
    contentCol: '文章全文',
    title: '小光與雜貨店',
    text: '小光每天在雜貨店幫阿婆的忙。有天他在超商看到一臺彈珠遊戲機，非常喜歡，但是沒錢買。所以，小光就趁阿婆不注意時，從抽屜拿了一張鈔票。傍晚，阿婆要關門休息，小光看到她急的四處翻找，嘴裡念著：「錢怎麼少了？」\n後來，小光走進超商。他盯著遊戲機，想起了阿婆急的樣子，怎麼也高興不起來。\n從那時候開始，小光就沒去雜貨店了。\n小光生日那天，託人送一塊蛋糕給阿婆。阿婆發現蛋糕底下有張紙條和五百元鈔票。紙條上寫著：阿婆：\n五百元是我拿的，本來想去買遊戲機，現在決定還給您。對不起！\n小光敬上\n過了幾天，小光經過雜貨店，發現店門關著。鄰居跟他說：「阿婆搬到兒子家了，她有一封信要給你。」小光打開一看，有張五百元鈔票。信紙上寫著：「小光，祝你生日快樂！」'
  }];
  var results = [];
  updates.forEach(function(u) {
    var sheet = ss.getSheetByName(u.sheet);
    if (!sheet) { results.push(u.title + ': 找不到工作表'); return; }
    var data = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var ti = headers.indexOf(u.titleCol), ci = headers.indexOf(u.contentCol);
    if (ti < 0 || ci < 0) { results.push(u.title + ': 找不到欄位'); return; }
    var updated = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][ti]).trim() === u.title) {
        sheet.getRange(i + 1, ci + 1).setValue(u.text);
        updated++;
      }
    }
    results.push(u.title + ': 更新 ' + updated + ' 列');
  });
  Logger.log(results.join('\n'));
  return results.join('\n');
}
