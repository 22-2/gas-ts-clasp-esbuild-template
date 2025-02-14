import { DATE_FORMAT, TIME_FORMAT, HOURLY_SHEET, DIARY_SHEET } from "./constants";
import { requestCharCount } from "./count-characters";

interface SheetConfig {
  NAME: string;
  COL_DATE: number;
  COL_DIFFERENCE: number;
  COL_COUNT: number;
  COL_TIME?: number;
}

/**
 * シートに記録を書き込む共通関数
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 書き込み対象のシート
 * @param {SheetConfig} sheetConfig - シートの設定オブジェクト
 * @param {boolean} isHourly - 時間別記録かどうか (true: 時間別, false: 日別)
 */
export function writeRecord(sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetConfig: SheetConfig, isHourly: boolean) {
  const count = requestCharCount();

  if (isNaN(count)) {
    throw new Error("文字数の取得に失敗しました。count-characters.ts内の設定を見直してください。");
  }

  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), DATE_FORMAT);

  let rowToInsert: number;

  if (isHourly) {
    // 時間別記録: 常に新しい行に追記
    rowToInsert = sheet.getLastRow() + 1;
    const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), TIME_FORMAT);
    sheet.getRange(rowToInsert, sheetConfig.COL_TIME!).setValue(formattedTime); // 時刻
  } else {
    // 日別記録: 今日の日付があればその行、なければ新しい行
    rowToInsert = findRowByDate(sheet, sheetConfig, now);
  }

  sheet.getRange(rowToInsert, sheetConfig.COL_DATE).setValue(formattedDate); // 日付
  sheet.getRange(rowToInsert, sheetConfig.COL_COUNT).setValue(count); // 文字数

  // 差分計算 (時間別と日別で関数を分ける)
  if (isHourly) {
    calculateHourlyDifference(sheet, sheetConfig, rowToInsert);
  } else {
    calculateDailyDifference(sheet, sheetConfig, rowToInsert);
  }
}

/**
 * 日付で記録行を検索する
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 検索対象のシート
 * @param {SheetConfig} sheetConfig - シート設定
 * @param {Date} date - 検索する日付
 * @returns {number} 見つかった行番号。見つからない場合は新しい行番号(最終行+1)を返す。
 */
function findRowByDate(sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetConfig: SheetConfig, date: Date): number {
  const lastRow = sheet.getLastRow();
  const dateValues = sheet.getRange(1, sheetConfig.COL_DATE, lastRow).getValues();

  for (let i = 0; i < dateValues.length; i++) {
    const rowDate = dateValues[i][0];
    //日付の比較
    if (rowDate instanceof Date && rowDate.toDateString() === date.toDateString()) {
      return i + 1;
    }
  }
  return lastRow + 1; // 見つからなかった場合は新しい行
}

/**
 * 差分を計算して書き込む (時間別記録用)
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 書き込み対象のシート
 * @param {SheetConfig} sheetConfig - シート設定
 * @param {number} currentRow - 現在の行番号
 */
function calculateHourlyDifference(sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetConfig: SheetConfig, currentRow: number) {
  if (currentRow <= 1) {
    sheet.getRange(currentRow, sheetConfig.COL_DIFFERENCE).setValue("N/A"); // 最初の行は "N/A"
    return;
  }

  const previousRow = currentRow - 1;
  const formula = `=IFERROR(${sheetConfig.NAME}!${getColName(sheetConfig.COL_COUNT)}${currentRow} - ${sheetConfig.NAME}!${getColName(sheetConfig.COL_COUNT)}${previousRow}, "N/A")`;
  sheet.getRange(currentRow, sheetConfig.COL_DIFFERENCE).setFormula(formula);
}

/**
 * 差分を計算して書き込む (日別記録用)
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 書き込み対象のシート
 * @param {SheetConfig} sheetConfig - シート設定
 * @param {number} currentRow - 現在の行番号
 */
function calculateDailyDifference(sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetConfig: SheetConfig, currentRow: number) {
  if (currentRow <= 1) {
    sheet.getRange(currentRow, sheetConfig.COL_DIFFERENCE).setValue("N/A");
    return;
  }

  // 今日の日付を取得
  const today = sheet.getRange(currentRow, sheetConfig.COL_DATE).getValue() as Date;
  const todayString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // INDIRECT関数とMATCH関数を組み合わせて、前日の日付の行を動的に検索
  const formula = `=IFERROR(${sheetConfig.NAME}!${getColName(sheetConfig.COL_COUNT)}${currentRow} - INDIRECT("${sheetConfig.NAME}!${getColName(sheetConfig.COL_COUNT)}"&MATCH(DATE(${todayString.split('-')[0]},${todayString.split('-')[1]},${todayString.split('-')[2]}-1),${sheetConfig.NAME}!${getColName(sheetConfig.COL_DATE)}1:${getColName(sheetConfig.COL_DATE)},0)), "N/A")`;
  sheet.getRange(currentRow, sheetConfig.COL_DIFFERENCE).setFormula(formula);
}

/**
 * 列番号を列名に変換する関数
 * @param {number} colNumber - 列番号(1始まり)
 * @returns {string} - 列名(A, B, C, ...)
 */
function getColName(colNumber: number): string {
  let colName = '';
  let temp: number;
  while (colNumber > 0) {
    temp = (colNumber - 1) % 26;
    colName = String.fromCharCode(65 + temp) + colName;
    colNumber = (colNumber - temp - 1) / 26;
  }
  return colName;
}
