import { requestCharCount } from "./count-characters"; // この関数は別途定義が必要です。

// 定数定義 (マジックナンバーを排除 & 明確化)
const COLUMN_DATE = 1; // 日付列
const COLUMN_TIME = 2; // 時刻列
const COLUMN_DIFFERENCE = 3; // 前日比列
const COLUMN_COUNT = 4; // 文字数列
const DATE_FORMAT = "yyyy/MM/dd";
const TIME_FORMAT = "HH:mm:ss";
const SHEET_NAME = "シート名"; // ★★★ シート名を指定 ★★★

function executeCount() {
  try {
    const count = requestCharCount(); // 文字数を取得 (この関数の実装は別途必要)

    if (typeof count !== 'number' || isNaN(count)) {
      throw new Error("requestCharCount() は有効な数値を返しませんでした。");
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();

    if (!sheet) {
      throw new Error(`シート '${SHEET_NAME}' が見つかりません。`);
    }

    // 現在の日付と時刻を取得
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), DATE_FORMAT);
    const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), TIME_FORMAT);

    // 最終行+1を取得（新規行に書き込む）
    const newRow = sheet.getLastRow() + 1;

    // 新しい行に値を書き込む
    sheet.getRange(newRow, COLUMN_DATE).setValue(formattedDate);
    sheet.getRange(newRow, COLUMN_TIME).setValue(formattedTime);
    sheet.getRange(newRow, COLUMN_COUNT).setValue(count);

    // 前の記録との差分を計算 (2行目以降)
    if (newRow > 1) {
      //前回の文字数を取得。
      let previousCount = "N/A";
      if (newRow > 2) { //2行目の時は比較対象がないのでスキップ
        previousCount = sheet.getRange(newRow - 1, COLUMN_COUNT).getValue();
      }

      let formula = `=IFERROR(${count} - ${previousCount}, "N/A")`; //数式を修正

      if (previousCount === "N/A") {
        sheet.getRange(newRow, COLUMN_DIFFERENCE).setValue("N/A");
      } else {
        sheet.getRange(newRow, COLUMN_DIFFERENCE).setFormula(formula); //数式を使用
      }


    } else {
      // 最初の行の場合、前日比は "N/A"
      sheet.getRange(newRow, COLUMN_DIFFERENCE).setValue("N/A");
    }


  } catch (error: any) {
    // エラーハンドリング
    Logger.log("エラー: " + error.message);
    Browser.msgBox("エラー: " + error.message); // ユーザーにエラーを表示
  }
}

(global as any).executeCount = executeCount;
