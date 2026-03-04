// =============================================
// Sales Timer - Google Spreadsheet 連携スクリプト
// =============================================
// 【設置手順】
// 1. スプレッドシートを開く
// 2. 拡張機能 → Apps Script
// 3. このコードを全部貼り付けて保存
// 4. デプロイ → 新しいデプロイ
// 5. 種類: ウェブアプリ
// 6. アクセスできるユーザー: 「全員」
// 7. デプロイ → URLをコピー
// 8. Sales Timer の設定画面にURLを貼り付け
// =============================================

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName('調整');
    var dataSheet = ss.getSheetByName('入金管理リスト');

    if (!configSheet || !dataSheet) {
      return jsonResponse({
        error: 'シートが見つかりません',
        sheets: ss.getSheets().map(function(s) { return s.getName(); })
      });
    }

    // 調整シートから列参照を取得
    var dateColRef = String(configSheet.getRange('E6').getValue()).trim();
    var amountColRef = String(configSheet.getRange('E12').getValue()).trim();
    var statusColRef = String(configSheet.getRange('E15').getValue()).trim();

    // 列文字 → インデックス変換 (A=1, B=2, ..., AA=27)
    function colToIndex(col) {
      col = col.toUpperCase().replace(/[^A-Z]/g, '');
      var idx = 0;
      for (var i = 0; i < col.length; i++) {
        idx = idx * 26 + (col.charCodeAt(i) - 64);
      }
      return idx;
    }

    var dateIdx = colToIndex(dateColRef) - 1;
    var amountIdx = colToIndex(amountColRef) - 1;
    var statusIdx = colToIndex(statusColRef) - 1;

    // クエリパラメータから期間を取得
    var params = e ? (e.parameter || {}) : {};
    var fromDate = params.from ? new Date(params.from) : null;
    var toDate = params.to ? new Date(params.to) : null;

    // 日付の時間部分をリセット
    if (fromDate) { fromDate.setHours(0, 0, 0, 0); }
    if (toDate) { toDate.setHours(23, 59, 59, 999); }

    var data = dataSheet.getDataRange().getValues();
    var records = [];
    var total = 0;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = String(row[statusIdx] || '').trim();

      if (status === '成約' || status.indexOf('契約済み') >= 0) {
        var amount = Number(row[amountIdx]) || 0;
        var dateVal = row[dateIdx];
        var dateStr = '';
        var rowDate = null;

        if (dateVal instanceof Date) {
          rowDate = dateVal;
          dateStr = dateVal.toISOString();
        } else if (dateVal) {
          dateStr = String(dateVal);
          // 文字列の日付もパースを試みる
          var parsed = new Date(dateVal);
          if (!isNaN(parsed.getTime())) {
            rowDate = parsed;
          }
        }

        // 日付フィルタリング
        if (fromDate && rowDate && rowDate < fromDate) continue;
        if (toDate && rowDate && rowDate > toDate) continue;

        if (amount > 0) {
          records.push({
            date: dateStr,
            amount: amount,
            status: status
          });
          total += amount;
        }
      }
    }

    return jsonResponse({
      total: total,
      records: records,
      count: records.length,
      filter: {
        from: fromDate ? fromDate.toISOString() : null,
        to: toDate ? toDate.toISOString() : null
      },
      config: {
        dateCol: dateColRef,
        amountCol: amountColRef,
        statusCol: statusColRef
      },
      timestamp: new Date().toISOString()
    });

  } catch (err) {
    return jsonResponse({ error: err.toString(), stack: err.stack || '' });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
