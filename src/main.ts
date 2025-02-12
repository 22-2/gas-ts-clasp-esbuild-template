import { requestCharCount } from "./count-characters"; // この関数は別途定義が必要です。
import { dailyRecord, hourlyRecord } from "./record";


//--------------------------------------------------
// シートごと、日付ごとに記録する関数 (毎日トリガーで実行)
//--------------------------------------------------
function executeDaily() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  dailyRecord(sheet);
}

//--------------------------------------------------
// 1時間ごとに記録する関数 (毎時トリガーで実行)
//--------------------------------------------------
function executeHourly() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  hourlyRecord(sheet); // 記録処理
}

(global as any).executeDaily = executeDaily;
(global as any).executeHourly = executeHourly;
