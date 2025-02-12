import { DIARY_SHEET, HOURLY_SHEET } from "./constants";
import { dailyRecord, hourlyRecord } from "./record";

//--------------------------------------------------
// シートごと、日付ごとに記録する関数 (毎日トリガーで実行)
//--------------------------------------------------
function executeDaily() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(DIARY_SHEET.NAME);
  if (sheet) dailyRecord(sheet);
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
  if (sheet) hourlyRecord(sheet); // 記録処理
  else {
    throw Error("Error: 記録対象のシートが見つかりません"); // 記録対象のシートが見つからなかった場合の処理}
  }
}

(global as any).executeHourly = executeHourly;
(global as any).executeDaily = executeDaily;
