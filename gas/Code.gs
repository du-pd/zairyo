const SHEET_ID = '1altLq8qp_TGzKcrUbyypcE2068HruaMY7_NkhVJppqM';

// 単価マスタの列（1行目は見出し）
var PRICE_HEADER = ['種別', '品名', '請求単価', '区分', '仕入日', '長さmm', '幅mm', '厚みmm', '本数', '購入量g', '購入金額', '仕入単価'];

function formatMaterials(materials) {
  if (!materials || materials.length === 0) return '';
  return materials.map(function(m) {
    var typeName = m.type === 'wood' ? '木材' : m.type === 'filament' ? 'フィラメント' : 'その他';
    var price = Number(m.price || 0).toLocaleString();
    return typeName + ': ' + (m.name || '') + '（' + (m.detail || '') + '）¥' + price;
  }).join(' / ');
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents || '{}');
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var now = new Date();

    if (payload.type === 'apply' || payload.type === 'approve') {
      var sheet = getOrCreateSheet_(ss, '記録');
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['日時', '種別', '学籍番号', '氏名', '材料内容', '合計金額（円）', '支払予定日', '備考', '承認教員']);
      }
      var materialsText = formatMaterials(payload.materials);
      if (payload.type === 'apply') {
        sheet.appendRow([now, '申請', payload.studentId||'', payload.studentName||'',
          materialsText, payload.totalAmount||0, '', payload.notes||'', '']);
      } else {
        sheet.appendRow([now, '承認', payload.studentId||'', payload.studentName||'',
          materialsText, payload.totalAmount||0, payload.payDate||'',
          payload.notes||'', payload.teacherName||'']);
      }
    } else if (payload.type === 'updatePrices') {
      var masterSheet = getOrCreateSheet_(ss, '単価マスタ');
      masterSheet.clearContents();
      var rows = [PRICE_HEADER];
      (payload.wood || []).forEach(function(w) {
        rows.push(['wood', w.name, w.pricePerM3, w.category || '', w.pcDate || '',
          v_(w.pcL), v_(w.pcW), v_(w.pcH), v_(w.pcQty), '', v_(w.pcCost), v_(w.pcUnit)]);
      });
      (payload.filament || []).forEach(function(f) {
        rows.push(['filament', f.name, f.pricePerG, '', f.pcDate || '',
          '', '', '', '', v_(f.pcWeight), v_(f.pcCost), v_(f.pcUnit)]);
      });
      (payload.other || []).forEach(function(o) {
        rows.push(['other', o.name, o.price, '', '', '', '', '', '', '', '', '']);
      });
      // 仕入日は日付に自動変換させず、文字のまま保存する
      masterSheet.getRange(1, 5, rows.length, 1).setNumberFormat('@');
      masterSheet.getRange(1, 1, rows.length, PRICE_HEADER.length).setValues(rows);
    }

    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
    if (action === 'getPrices') {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var sheet = ss.getSheetByName('単価マスタ');
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({status:'error',message:'単価マスタシートが見つかりません'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var tz = ss.getSpreadsheetTimeZone();
      var data = sheet.getDataRange().getValues();
      var wood = [];
      var filament = [];
      var other = [];
      for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var type = String(row[0] || '').trim();
        var name = String(row[1] || '').trim();
        var price = row[2];
        if (!type || !name || price === '' || price === null || price === undefined) continue;
        if (type === 'wood') {
          var w = { name: name, pricePerM3: Number(price) };
          if (row[3]) w.category = String(row[3]);
          if (row[10] !== '' && row[10] !== undefined) {
            w.pcDate = dateText_(row[4], tz);
            w.pcL = Number(row[5]); w.pcW = Number(row[6]); w.pcH = Number(row[7]);
            w.pcQty = Number(row[8]); w.pcCost = Number(row[10]); w.pcUnit = Number(row[11]);
          }
          wood.push(w);
        } else if (type === 'filament') {
          var f = { name: name, pricePerG: Number(price) };
          if (row[10] !== '' && row[10] !== undefined) {
            f.pcDate = dateText_(row[4], tz);
            f.pcWeight = Number(row[9]); f.pcCost = Number(row[10]); f.pcUnit = Number(row[11]);
          }
          filament.push(f);
        } else if (type === 'other') {
          other.push({ name: name, price: Number(price) });
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status:'ok', wood:wood, filament:filament, other:other}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:'unknown action'})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// 未入力（undefined / null）は空欄にする
function v_(x) {
  return (x === undefined || x === null) ? '' : x;
}

// 日付として入ってしまったセルも yyyy-MM-dd に戻す
function dateText_(x, tz) {
  if (x instanceof Date) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
  return String(x || '');
}
