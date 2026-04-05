/**
 * ============================================================
 *  郵便番号・住所照合チェックツール
 * ============================================================
 *
 * 【概要】
 *   スプレッドシート上のチケット購入者データに対して、
 *   住所と郵便番号の整合性を自動チェックするツール。
 *   Gemini API を利用した住所分割と、KEN_ALL（郵便番号DB）による照合を行う。
 *
 * 【関数一覧】
 *
 *  ■ メイン関数（メニュー / 手動実行用）
 *   - startSearchByIds()        : 受付番号を指定して検索・処理する
 *   - startProcessByMemo()      : 「メモ」列が「変更」の行を自動抽出して処理する
 *   - startProcessAllSheets()   : TARGET_SHEETS の全件を一括処理する（中断・再開対応）
 *   - startFromSpecificIndex()  : 任意のインデックスから処理を再開する（リカバリ用）
 *   - resetProgress()           : 中断時の進捗情報をリセットする
 *
 *  ■ コア処理
 *   - processPostalData()       : 住所分割・郵便番号照合・判定を行うメイン処理ループ
 *   - saveProgress_()           : 処理結果を「結果」シートへ書き出し、進捗を保存する
 *
 *  ■ 郵便番号DB 検索
 *   - loadKenAllDbHighPrecision_()  : utf_ken_all シートから郵便番号DBを読み込む
 *   - findBestMatchNoSplit_()       : 正規化住所でDB内を最長一致検索する
 *
 *  ■ Gemini API 連携
 *   - callGeminiPreciseSplitBatch_() : Gemini API で住所を address1 / address2 に分割する
 *
 *  ■ 判定
 *   - checkTotalMatch_()        : 郵便番号一致・住所欠損・氏名空欄などの総合判定を行う
 *
 *  ■ 正規化・変換ヘルパー
 *   - normAddrHighPrecision_()  : 住所文字列を検索用に高精度正規化する
 *   - normAddrForCompare_()     : 住所文字列を比較用に正規化する
 *   - toHalfWidth_()            : 全角英数を半角に変換する
 *   - toKanjiNumber_()          : 数値を漢数字に変換する
 *   - formatZipCode_()          : 郵便番号を xxx-xxxx 形式にフォーマットする
 *   - unifyShortAddress2Strict_() : 短い address2 を address1 に統合する
 *
 * ============================================================
 */

/**
 * 設定
 */
const DB_SHEET_NAME = 'utf_ken_all';
const RESULT_SHEET_NAME = '結果';
const HEADER_ROWS = 1;
const MAX_EXECUTION_TIME = 300000;

// ★【追加】対象とする入力シート名をすべてここに記述してください
const TARGET_SHEETS = ['0303_1530まで分', '0305_1430まで分', 'シート3'];

/**
 * 1. 指定した受付番号を複数検索して実行する関数
 */
function startSearchByIds() {
  const ui = SpreadsheetApp.getUi();
  // 説明文も少し分かりやすく書き換えています
  const response = ui.prompt('受付番号検索', '検索したい番号を入力してください（スペース、カンマ、改行区切りOK）', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  // 正規表現で分割：カンマ(全半角)、スペース(全半角)、改行に対応
  const inputIds = response.getResponseText().split(/[,，\s\n]+/).map(s => s.trim()).filter(s => s !== "");

  if (inputIds.length === 0) {
    ui.alert('受付番号を入力してください。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = loadKenAllDbHighPrecision_();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  let totalFound = 0; // 見つかった合計件数のカウント用

  TARGET_SHEETS.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const allData = sh.getDataRange().getValues();
    const header = allData[0];
    const requestIdIdx = header.indexOf('受付番号');

    const targetRows = allData.slice(HEADER_ROWS).filter(row => {
      return inputIds.includes(row[requestIdIdx].toString());
    });

    if (targetRows.length > 0) {
      totalFound += targetRows.length;
      console.log(`${sheetName} から ${targetRows.length} 件見つかりました。`);
      processPostalData(0, targetRows, sheetName, db, apiKey);
    }
  });

  if (totalFound === 0) {
    ui.alert('指定された受付番号はどのシートにも見つかりませんでした。');
  } else {
    ui.alert(totalFound + ' 件の処理が完了しました。');
  }
}

/**
 * 1. 各シートの「メモ」列が「変更」となっている行を自動抽出して実行する関数
 */
function startProcessByMemo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = loadKenAllDbHighPrecision_();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    ui.alert('APIキーが設定されていません。');
    return;
  }

  let totalFound = 0;

  TARGET_SHEETS.forEach(sheetName => {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return;

    const allData = sh.getDataRange().getValues();
    const header = allData[0];

    // 各列のインデックスを取得
    const memoIdx = header.indexOf('メモ');
    const requestIdIdx = header.indexOf('受付番号');

    if (memoIdx === -1) {
      console.log(`シート「${sheetName}」に「メモ」列がないためスキップします。`);
      return;
    }

    // 「メモ」列が「変更」となっている行だけを抽出
    const targetRows = allData.slice(HEADER_ROWS).filter(row => {
      const memoValue = (row[memoIdx] || '').toString().trim();
      return memoValue === '変更';
    });

    if (targetRows.length > 0) {
      totalFound += targetRows.length;
      console.log(`${sheetName} から「変更」対象を ${targetRows.length} 件抽出しました。`);

      // 抽出したデータのみを処理に回す
      // 第1引数の startIndex は 0（抽出済みリストの先頭から）で固定
      processPostalData(0, targetRows, sheetName, db, apiKey);
    }
  });

  if (totalFound === 0) {
    ui.alert('「メモ」列が「変更」となっている行は見つかりませんでした。');
  } else {
    ui.alert(`全シート合計 ${totalFound} 件の処理が完了しました。`);
  }
}

/**
 * 2. 従来の全件処理を「複数シート対応」に拡張したメイン関数
 */
function startProcessAllSheets() {
  const props = PropertiesService.getScriptProperties();
  const db = loadKenAllDbHighPrecision_();
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

  // 現在どのシートを処理中か保持（中断・再開用）
  let currentSheetIdx = parseInt(props.getProperty('CURRENT_SHEET_IDX') || "0", 10);
  let lastIndex = parseInt(props.getProperty('LAST_INDEX') || "0", 10);

  for (let i = currentSheetIdx; i < TARGET_SHEETS.length; i++) {
    const sheetName = TARGET_SHEETS[i];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) continue;

    const dataRows = sh.getDataRange().getValues().slice(HEADER_ROWS);

    // 処理実行
    const isFinished = processPostalData(lastIndex, dataRows, sheetName, db, apiKey);

    if (!isFinished) {
      // 中断した場合、現在のシートインデックスを保存して終了
      props.setProperty('CURRENT_SHEET_IDX', i.toString());
      return;
    }
    // シート完了につきインデックスリセット
    lastIndex = 0;
    props.deleteProperty('LAST_INDEX');
  }
  props.deleteProperty('CURRENT_SHEET_IDX');
}

/**
 * コア処理部分（引数を拡張して再利用性を向上）
 * 戻り値: 全件完了したらtrue、中断したらfalse
 */
function processPostalData(startIndex, dataRows, currentSheetName, db, apiKey) {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSh = ss.getSheetByName(currentSheetName);
  const headerValues = inputSh.getDataRange().getValues()[0];

  const colMap = {
    requestId: headerValues.indexOf('受付番号'),
    zip:       headerValues.indexOf('郵便番号'),
    address:   headerValues.indexOf('住所'),
    lastName:  headerValues.indexOf('名前(姓)'),
    firstName: headerValues.indexOf('名前(名)')
  };

  const outRows = [];
  const BATCH_SIZE = 3;
  let processedCountInThisRun = 0;

  for (let i = startIndex; i < dataRows.length; i += BATCH_SIZE) {
    if (new Date().getTime() - startTime > MAX_EXECUTION_TIME) {
      saveProgress_(outRows, startIndex + processedCountInThisRun, dataRows.length, currentSheetName);
      return false; // 中断
    }

    const batch = dataRows.slice(i, i + BATCH_SIZE);
    const requestItems = batch.map(r => ({
      id: (r[colMap.requestId] || '').toString(),
      addr: (r[colMap.address] || '').toString()
    }));

    const splitResults = callGeminiPreciseSplitBatch_(requestItems, apiKey);

    for (let j = 0; j < batch.length; j++) {
      const row = batch[j];
      const targetId = (row[colMap.requestId] || '').toString();
      const inputZip = (row[colMap.zip] || '').toString();
      const fullAddress = (row[colMap.address] || '').toString();
      const lName = colMap.lastName !== -1 ? (row[colMap.lastName] || '') : '';
      const fName = colMap.firstName !== -1 ? (row[colMap.firstName] || '') : '';
      const formattedName = (lName + ' ' + fName).trim() + ' 様';

      const normAddrForSearch = normAddrHighPrecision_(fullAddress);
      const searchResult = findBestMatchNoSplit_(normAddrForSearch, db);
      let foundZip = searchResult.zipcode || 'no-data';
      let isPredicted = false;

      let splitInfo = (splitResults && Array.isArray(splitResults) && splitResults.find(res => res.id === targetId))
                      || { address1: fullAddress, address2: '', predictedZip: 'no-data' };

      if (foundZip === 'no-data' && splitInfo.predictedZip && splitInfo.predictedZip !== 'no-data') {
        foundZip = formatZipCode_(splitInfo.predictedZip);
        isPredicted = true;
      }

      splitInfo.address1 = toHalfWidth_(splitInfo.address1 || '');
      splitInfo.address2 = toHalfWidth_(splitInfo.address2 || '');
      const finalJudgement = checkTotalMatch_(fullAddress, splitInfo, foundZip, inputZip, lName, fName, isPredicted);

      outRows.push([targetId, foundZip, splitInfo.address1, splitInfo.address2, formattedName, inputZip, fullAddress, finalJudgement, currentSheetName]);
    }
    processedCountInThisRun += batch.length;
    Utilities.sleep(1000);
  }

  saveProgress_(outRows, startIndex + processedCountInThisRun, dataRows.length, currentSheetName);
  return true; // 完了
}

// saveProgress_ 内の引数に元シート名が含まれるよう修正
function saveProgress_(outRows, nextIndex, totalLength, currentSheetName) {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let resultSh = ss.getSheetByName(RESULT_SHEET_NAME) || ss.insertSheet(RESULT_SHEET_NAME);

  if (resultSh.getLastRow() === 0) {
    const headers = [["受付番号", "郵便番号(検索)", "住所1", "住所2", "氏名", "元の郵便番号", "元の住所", "判定", "元シート名"]];
    resultSh.getRange(1, 1, 1, 9).setValues(headers).setBackground('#f3f3f3').setFontWeight('bold');
    resultSh.setFrozenRows(1);
  }

  if (outRows.length > 0) {
    resultSh.getRange(resultSh.getLastRow() + 1, 1, outRows.length, 9).setValues(outRows);
  }

  if (nextIndex < totalLength) {
    props.setProperty('LAST_INDEX', nextIndex.toString());
    SpreadsheetApp.getUi().alert(`${currentSheetName} の処理を中断しました。再開するには再度実行してください。`);
  } else if (currentSheetName === TARGET_SHEETS[TARGET_SHEETS.length - 1]) {
    // 本当に最後のシートが終わった時
    props.deleteProperty('LAST_INDEX');
    props.deleteProperty('CURRENT_SHEET_IDX');
    SpreadsheetApp.getUi().alert('すべての対象シートの処理が完了しました！');
  }
}

// 以降、補助関数（normAddrForCompare_等）は元のまま維持してください。
/**
 * 進捗のリセット（最初からやり直したい時用）
 */
function resetProgress() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_INDEX');
  SpreadsheetApp.getUi().alert('進捗をリセットしました。');
}

// --- 以下の normAddrForCompare_ などの補助関数は元のコードのまま維持してください ---
/**
 * 比較用正規化ロジック (変更なし)
 */
function normAddrForCompare_(s) {
  if (!s) return '';
  let norm = s.toString();
  norm = toHalfWidth_(norm).replace(/[ぁ-ゖ]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0x60));
  const kanjiMap = { '一':1, '二':2, '三':3, '四':4, '五':5, '六':6, '七':7, '八':8, '九':9, '十':10 };
  norm = norm.replace(/[一二三四五六七八九十]+/g, function(match) {
    let res = 0, tmp = 0;
    for (let char of match) {
      if (char === '十') { res += (tmp === 0 ? 10 : tmp * 10); tmp = 0; }
      else { tmp = kanjiMap[char]; }
    }
    return res + tmp;
  });
  return norm.replace(/丁目/g, '-')
             .replace(/番地/g, '-')
             .replace(/号/g, '')
             .replace(/[\s\-ー－・、，,\.。町大字字]/g, "")
             .trim();
}

/**
 * 詳細判定 (変更なし)
 */
function checkTotalMatch_(originalAddr, splitInfo, foundZip, inputZip, lastName, firstName, isPredicted) {
  let errors = [];
  if (!lastName.toString().trim() && !firstName.toString().trim()) errors.push("氏名空欄");
  const z1 = (foundZip || '').toString().replace(/\D/g, '');
  const z2 = (inputZip || '').toString().replace(/\D/g, '');
  if (foundZip === 'no-data') errors.push("確認不可(no-data)");
  else if (!z2) errors.push("入力Zip空欄");
  else if (z1 !== z2) errors.push("Zip不一致");
  const combined = splitInfo.address1 + splitInfo.address2;
  if (normAddrForCompare_(originalAddr) !== normAddrForCompare_(combined)) {
    errors.push("住所欠損あり");
  }
  let status = errors.length === 0 ? "一致" : errors.join(" / ");
  if (isPredicted && errors.length === 0) status = "一致(AI推測)";
  return status;
}

/**
 * 高精度化・検索系補助関数 (変更なし)
 */
function normAddrHighPrecision_(s) {
  if (!s) return '';
  s = s.replace(/[ぁ-ゖ]/g, ch => String.fromCharCode(ch.charCodeAt(0) + 0x60));
  s = s.replace(/大字/g, '').replace(/字/g, '');
  s = s.replace(/\d+/g, numStr => toKanjiNumber_(parseInt(numStr, 10)));
  return s.replace(/[！-～]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
          .replace(/\u3000/g, ' ')
          .replace(/ヶ/g, 'ケ').replace(/ヵ/g, 'カ')
          .replace(/[‐–—―－ー]/g, '-')
          .replace(/町/g, '')
          .replace(/[\s\-・、，,\.。]/g, '')
          .replace(/[（(].*?[)）]/g, '')
          .trim();
}

function toKanjiNumber_(n) {
  if (isNaN(n) || n <= 0 || n >= 100) return String(n);
  const map = ['','一','二','三','四','五','六','七','八','九'];
  if (n <= 10) return ['', '一','二','三','四','五','六','七','八','九','十'][n];
  if (n < 20) return '十' + map[n - 10];
  const tens = Math.floor(n / 10);
  const ones = n % 10;
  return (tens === 1 ? '十' : map[tens] + '十') + (ones ? map[ones] : '');
}

function findBestMatchNoSplit_(normalizedAddress, dbRows) {
  const head = normalizedAddress.slice(0, 10);
  let best = null;
  for (let row of dbRows) {
    if (!row.key.startsWith(head.slice(0, Math.min(4, head.length)))) continue;
    if (normalizedAddress.indexOf(row.key) !== -1) {
      if (!best || row.key.length > best.len) best = { len: row.key.length, row: row };
    }
  }
  if (best) return { zipcode: best.row.zip };
  let best2 = null;
  for (let row of dbRows) {
    const key2 = normAddrHighPrecision_(row.city + row.town.replace(/町/g, ''));
    if (normalizedAddress.indexOf(key2) !== -1) {
      if (!best2 || key2.length > best2.len) best2 = { len: key2.length, row: row };
    }
  }
  return best2 ? { zipcode: best2.row.zip } : { zipcode: null };
}

function loadKenAllDbHighPrecision_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  const vals = dbSheet.getDataRange().getValues();
  return vals.slice(1).map(r => {
    const zipRaw = (r[2] || '').toString();
    const pref = (r[6] || '').toString().trim(), city = (r[7] || '').toString().trim();
    let town = (r[8] || '').toString().trim().replace(/（.*?）/g, '').replace(/以下に掲載がない場合/g, '');
    town = town.replace(/ヶ/g, 'ケ').replace(/ヵ/g, 'カ');
    const key = normAddrHighPrecision_(pref + city + town.replace(/町/g, ''));
    const zipFmt = zipRaw.length === 7 ? zipRaw.slice(0, 3) + '-' + zipRaw.slice(3) : zipRaw;
    return { zip: zipFmt, pref: pref, city: city, town: town, key: key };
  });
}

function callGeminiPreciseSplitBatch_(requestItems, apiKey) {
  const modelId = "gemini-2.5-flash-lite";
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelId}:generateContent?key=${apiKey}`;
const prompt = `以下の住所リストを解析し、JSON形式で返してください。
【分割の指針】
- address1: 都道府県 ＋ 市区町村 ＋ 町名 ＋ 番地（号）まで。
- address2: 建物名 ＋ 部屋番号。建物名がある場合は必ずこちらに分離してください。
- 数字、英字、ハイフンはすべて「半角」に変換してください。
- 語順は一切変えないでください。

【出力形式】
[{"id":"受付番号","predictedZip":"郵便番号","address1":"前半","address2":"後半"}]

リスト:
${requestItems.map(item => `ID:${item.id} 住所:${item.addr}`).join('\n')}`;
  const payload = { "contents": [{ "parts": [{ "text": prompt }] }], "generationConfig": { "response_mime_type": "application/json", "temperature": 0.1 } };
  const options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
  try {
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() === 200) return JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts[0].text);
  } catch (e) { console.error(e); }
  return [];
}

function unifyShortAddress2Strict_(splitInfo) {
  let a1 = (splitInfo.address1 || '').trim(), a2 = (splitInfo.address2 || '').replace(/^[ \-ー－]+/, '').trim();
  if (!a2) return { address1: a1, address2: "" };
  const shortRegex = /^[0-9０-９a-zA-Zａ-ｚＡ-Ｚ\-ー－号室]+$/;
  if (shortRegex.test(a2)) {
    const needsHyphen = /[0-9０-９a-zA-Zａ-ｚＡ-ｚ]$/.test(a1) && /^[0-9０-９a-zA-Zａ-ｚＡ-ｚ]/.test(a2);
    a1 = a1 + (needsHyphen ? "-" : " ") + a2; a2 = "";
  }
  return { address1: a1, address2: a2 };
}

function toHalfWidth_(str) {
  if (!str) return "";
  return str.replace(/[！-～]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/　/g, " ");
}

function formatZipCode_(zip) {
  const clean = zip.replace(/\D/g, '');
  return clean.length === 7 ? clean.slice(0, 3) + '-' + clean.slice(3) : zip;
}

/**
 * 指定したインデックスから処理を開始する（デバッグ・リカバリ用）
 */
function startFromSpecificIndex() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '開始インデックスの指定',
    '処理を開始するインデックス番号（0から始まる数値）を入力してください：',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const inputIndex = parseInt(response.getResponseText(), 10);

    if (isNaN(inputIndex) || inputIndex < 0) {
      ui.alert('有効な数値を入力してください。');
      return;
    }

    // 進捗プロパティを上書き保存
    PropertiesService.getScriptProperties().setProperty('LAST_INDEX', inputIndex.toString());

    ui.alert(`インデックス ${inputIndex} から開始します。`);
    processPostalData(inputIndex);
  }
}
