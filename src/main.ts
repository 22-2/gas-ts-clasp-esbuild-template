import { DIARY_SHEET, HOURLY_SHEET } from "./constants";
import { writeRecord } from "./helpers";

//--------------------------------------------------
// シートごと、日付ごとに記録する関数 (毎日トリガーで実行)
//--------------------------------------------------
function executeDaily() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(DIARY_SHEET.NAME);
  if (sheet) writeRecord(sheet, DIARY_SHEET, false);
  else {
    throw Error("Error: 記録対象のシートが見つかりません"); // 記録対象のシートが見つからなかった場合の処理}
  }
}
//--------------------------------------------------
// 1時間ごとに記録する関数 (毎時トリガーで実行)
//--------------------------------------------------
function executeHourly() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HOURLY_SHEET.NAME);
  if (sheet) writeRecord(sheet, HOURLY_SHEET, true);
  else {
    throw Error("Error: 記録対象のシートが見つかりません"); // 記録対象のシートが見つからなかった場合の処理}
  }
}

(global as any).executeHourly = executeHourly;
(global as any).executeDaily = executeDaily;
