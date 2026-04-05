/**
 * ============================================================
 *  チケット情報追記ツール
 * ============================================================
 *
 * 【概要】
 *   「結果」シートの受付番号をキーにして、全シートからチケット情報
 *   （リベシティ名・プロフィールURL）を検索し、実行時間と共に追記する。
 *
 * 【関数一覧】
 *   - appendTicketInformationWithTimestamp() : 結果シートにチケット情報と実行時間を追記する
 *
 * ============================================================
 */

/**
 * 「結果」シートに対し、全シートからチケット情報を検索して「実行時間」と共に追記する
 */
function appendTicketInformationWithTimestamp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('結果');

  if (!resultSheet) {
    SpreadsheetApp.getUi().alert('「結果」シートが見つかりません。');
    return;
  }

  const resultData = resultSheet.getDataRange().getValues();
  if (resultData.length <= 1) return;

  const resultHeader = resultData[0];
  const idColIdx = resultHeader.indexOf('受付番号');

  if (idColIdx === -1) {
    SpreadsheetApp.getUi().alert('「結果」シートに「受付番号」列が見つかりません。');
    return;
  }

  // 現在時刻を取得（実行時間の記録用）
  const timestamp = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss");

  // 2. 全シートから検索用データをキャッシュ
  const ticketMap = {};
  const allSheets = ss.getSheets();

  allSheets.forEach(sheet => {
    if (sheet.getName() === '結果' || sheet.getLastRow() < 1) return;

    const data = sheet.getDataRange().getValues();
    const header = data[0];

    const colIdx = {
      id: header.indexOf('受付番号'),
      ticket: header.indexOf('リベシティ名'),
      option: header.indexOf('リベシティプロフィールURL')
    };

    if (colIdx.id === -1) return;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[colIdx.id]).trim();
      if (!id) continue;

      if (!ticketMap[id]) {
        ticketMap[id] = {
          ticket: colIdx.ticket !== -1 ? row[colIdx.ticket] : '',
          option: colIdx.option !== -1 ? row[colIdx.option] : ''
        };
      }
    }
  });

  // 3. 書き込み用データの作成
  const outputRows = [];

  for (let i = 1; i < resultData.length; i++) {
    const id = String(resultData[i][idColIdx]).trim();
    const match = ticketMap[id];

    if (match) {
      // [チケット, オプション, 実行時間]
      outputRows.push([match.ticket, match.option, timestamp]);
    } else {
      outputRows.push(['', '', timestamp]);
    }
  }

  // 4. 「結果」シートに3列（チケット、オプション、実行時間）を追記
  const lastCol = resultHeader.length;
  const newHeaders = [['イベントチケット', 'イベントオプションチケット', '実行時間']];

  // ヘッダー追加
  resultSheet.getRange(1, lastCol + 1, 1, 3).setValues(newHeaders);

  // データ流し込み
  if (outputRows.length > 0) {
    resultSheet.getRange(2, lastCol + 1, outputRows.length, 3).setValues(outputRows);
  }

  resultSheet.autoResizeColumns(lastCol + 1, 3);
  SpreadsheetApp.getUi().alert('チケット情報と実行時間の追記が完了しました。');
}
