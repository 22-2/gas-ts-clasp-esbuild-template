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

  // 差分計算
  calculateDifference(sheet, sheetConfig, rowToInsert, count);
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
 * 差分を計算して書き込む
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 書き込み対象のシート
 * @param {SheetConfig} sheetConfig - シート設定
 * @param {number} currentRow - 現在の行番号
 * @param {number} currentCount - 現在の文字数
 */
function calculateDifference(sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetConfig: SheetConfig, currentRow: number, currentCount: number) {
  if (currentRow <= 1) {
    sheet.getRange(currentRow, sheetConfig.COL_DIFFERENCE).setValue("N/A"); // 最初の行は "N/A"
    return;
  }

  let previousCount: number | "N/A" = "N/A";
  const values = sheet.getRange(1, sheetConfig.COL_DATE, currentRow - 1, 2).getValues();

  // 日別記録の場合、前日のデータを取得
  if (sheetConfig.COL_TIME === undefined) {
    for (let i = values.length - 1; i >= 0; i--) {
      const prebDate = values[i][0];
      const today = sheet.getRange(currentRow, sheetConfig.COL_DATE).getValue();
      if (!(prebDate instanceof Date && today instanceof Date && prebDate.toDateString() === today.toDateString())) {
        const count = values[i][1]; //日付の隣のセル（カウント）を取得
        if (typeof count === "number") {
          previousCount = count;
          break;
        }
      }
    }
  }
  //時間別記録の場合
  else {
    const count = values[values.length - 1][1];
    if (typeof count === "number") {
      previousCount = count;
    }
  }


  if (typeof previousCount === "number") {
    const formula = `=IFERROR(${currentCount} - ${previousCount}, "N/A")`;
    sheet.getRange(currentRow, sheetConfig.COL_DIFFERENCE).setFormula(formula);
  } else {
    sheet.getRange(currentRow, sheetConfig.COL_DIFFERENCE).setValue("N/A");
  }
}

