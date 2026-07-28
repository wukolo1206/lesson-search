// zhu-docx.js — 在瀏覽器裡直接產生 Word .docx
//
// docx 就是一個 zip。模板 templates/練習卷範本.docx 已改為「不壓縮儲存」
// （ZIP_STORED，見 scripts/make_practice_template.py），所以這裡只需要一個
// 極小的 zip 讀寫器，不必依賴任何函式庫，也不必用 DecompressionStream。
//
// 版面（紙張、邊距、字型、樣式、分頁保護）全部沿用模板裡的 styles.xml 與
// sectPr——模板仍是唯一的版面來源，本檔只負責產生段落。
var ZhuDocx = (function () {
  'use strict';

  var DOC_PATH = 'word/document.xml';

  // ---------------------------------------------------------------- CRC32
  var CRC_TABLE = (function () {
    var table = new Uint32Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      table[n] = c >>> 0;
    }
    return table;
  })();

  function crc32(bytes) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < bytes.length; i++) {
      c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // ------------------------------------------------------------ zip 讀取
  // 只支援 STORED（未壓縮）。模板以外的 zip 不保證讀得動，這是刻意的：
  // 讀不動就該在開發階段就爆掉，而不是產出壞掉的 docx。
  function readZip(buffer) {
    var dv = new DataView(buffer);
    var bytes = new Uint8Array(buffer);

    var eocd = -1;
    for (var i = bytes.length - 22; i >= 0 && i > bytes.length - 65558; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) { throw new Error('模板不是有效的 zip：找不到 EOCD'); }

    var count = dv.getUint16(eocd + 10, true);
    var cdOffset = dv.getUint32(eocd + 16, true);

    var entries = [];
    var p = cdOffset;
    for (var n = 0; n < count; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) {
        throw new Error('模板 zip 中央目錄損壞');
      }
      var method = dv.getUint16(p + 10, true);
      var size = dv.getUint32(p + 24, true);
      var nameLen = dv.getUint16(p + 28, true);
      var extraLen = dv.getUint16(p + 30, true);
      var commentLen = dv.getUint16(p + 32, true);
      var localOffset = dv.getUint32(p + 42, true);
      var name = utf8Decode(bytes.subarray(p + 46, p + 46 + nameLen));
      if (method !== 0) {
        throw new Error('模板必須以未壓縮方式儲存，但 ' + name + ' 是壓縮的。'
          + '請重跑 scripts/make_practice_template.py');
      }
      // 本地檔頭的 extra 長度可能與中央目錄不同，要各自讀
      var lnameLen = dv.getUint16(localOffset + 26, true);
      var lextraLen = dv.getUint16(localOffset + 28, true);
      var dataStart = localOffset + 30 + lnameLen + lextraLen;
      entries.push({ name: name, data: bytes.subarray(dataStart, dataStart + size) });
      p += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }

  // ------------------------------------------------------------ zip 寫出
  function writeZip(entries) {
    var chunks = [];
    var central = [];
    var offset = 0;

    entries.forEach(function (e) {
      var nameBytes = utf8Encode(e.name);
      var crc = crc32(e.data);

      var local = new Uint8Array(30 + nameBytes.length);
      var ldv = new DataView(local.buffer);
      ldv.setUint32(0, 0x04034b50, true);
      ldv.setUint16(4, 20, true);          // version needed
      ldv.setUint16(6, 0, true);           // flags
      ldv.setUint16(8, 0, true);           // method = STORED
      ldv.setUint16(10, 0, true);          // time
      ldv.setUint16(12, 0x21, true);       // date（1980-01-01，固定值讓輸出可重現）
      ldv.setUint32(14, crc, true);
      ldv.setUint32(18, e.data.length, true);
      ldv.setUint32(22, e.data.length, true);
      ldv.setUint16(26, nameBytes.length, true);
      ldv.setUint16(28, 0, true);
      local.set(nameBytes, 30);

      chunks.push(local, e.data);

      var cd = new Uint8Array(46 + nameBytes.length);
      var cdv = new DataView(cd.buffer);
      cdv.setUint32(0, 0x02014b50, true);
      cdv.setUint16(4, 20, true);
      cdv.setUint16(6, 20, true);
      cdv.setUint16(8, 0, true);
      cdv.setUint16(10, 0, true);
      cdv.setUint16(12, 0, true);
      cdv.setUint16(14, 0x21, true);
      cdv.setUint32(16, crc, true);
      cdv.setUint32(20, e.data.length, true);
      cdv.setUint32(24, e.data.length, true);
      cdv.setUint16(28, nameBytes.length, true);
      cdv.setUint32(42, offset, true);
      cd.set(nameBytes, 46);
      central.push(cd);

      offset += local.length + e.data.length;
    });

    var cdSize = central.reduce(function (a, c) { return a + c.length; }, 0);
    var eocd = new Uint8Array(22);
    var edv = new DataView(eocd.buffer);
    edv.setUint32(0, 0x06054b50, true);
    edv.setUint16(8, entries.length, true);
    edv.setUint16(10, entries.length, true);
    edv.setUint32(12, cdSize, true);
    edv.setUint32(16, offset, true);

    var all = chunks.concat(central, [eocd]);
    var total = all.reduce(function (a, c) { return a + c.length; }, 0);
    var out = new Uint8Array(total);
    var pos = 0;
    all.forEach(function (c) { out.set(c, pos); pos += c.length; });
    return out;
  }

  // ------------------------------------------------------------ 文字編碼
  function utf8Encode(str) {
    if (typeof TextEncoder !== 'undefined') { return new TextEncoder().encode(str); }
    var esc = unescape(encodeURIComponent(str));
    var arr = new Uint8Array(esc.length);
    for (var i = 0; i < esc.length; i++) { arr[i] = esc.charCodeAt(i); }
    return arr;
  }

  function utf8Decode(bytes) {
    if (typeof TextDecoder !== 'undefined') { return new TextDecoder('utf-8').decode(bytes); }
    var s = '';
    for (var i = 0; i < bytes.length; i++) { s += String.fromCharCode(bytes[i]); }
    return decodeURIComponent(escape(s));
  }

  // ------------------------------------------------------------- XML 產生
  function esc(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function paragraphXml(style, text) {
    // xml:space="preserve" 是必要的——填空用的全形空白在句首句尾都會被吃掉
    return '<w:p><w:pPr><w:pStyle w:val="' + esc(style) + '"/></w:pPr>'
      + '<w:r><w:t xml:space="preserve">' + esc(text) + '</w:t></w:r></w:p>';
  }

  /**
   * 用模板產生一份 docx。
   * @param {ArrayBuffer} templateBuffer  templates/練習卷範本.docx 的內容
   * @param {Array} paragraphs            [{ style: '卷_大題', text: '一、...' }]
   * @returns {Uint8Array}
   */
  function build(templateBuffer, paragraphs) {
    var entries = readZip(templateBuffer);
    var docEntry = null;
    for (var i = 0; i < entries.length; i++) {
      if (entries[i].name === DOC_PATH) { docEntry = entries[i]; break; }
    }
    if (!docEntry) { throw new Error('模板缺少 ' + DOC_PATH); }

    var xml = utf8Decode(docEntry.data);
    var bodyAt = xml.indexOf('<w:body>');
    var sectAt = xml.indexOf('<w:sectPr');
    if (bodyAt < 0 || sectAt < 0 || sectAt < bodyAt) {
      throw new Error('模板 document.xml 結構不符預期');
    }

    var body = paragraphs.map(function (p) {
      return paragraphXml(p.style, p.text);
    }).join('');

    var newXml = xml.slice(0, bodyAt + '<w:body>'.length) + body + xml.slice(sectAt);
    docEntry.data = utf8Encode(newXml);
    return writeZip(entries);
  }

  function download(bytes, filename) {
    var blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  return {
    build: build,
    download: download,
    _readZip: readZip,
    _writeZip: writeZip,
    _crc32: crc32
  };
})();

if (typeof module !== 'undefined' && module.exports) { module.exports = ZhuDocx; }
