import { requestCharCount } from "./count-characters";

// 定数定義 (マジックナンバーを排除)
const DATE_COLUMN = 1;
const TIME_COLUMN = 2;
const DIFFERENCE_COLUMN = 3;
const COUNT_COLUMN = 4;
const DATE_FORMAT = "yyyy/MM/dd";
const TIME_FORMAT = "HH:mm:ss"; // 時刻のフォーマット

function executeCount() {
  try {
    const count = requestCharCount(); // 文字数を取得 (この関数の実装は別途必要)

    if (typeof count !== 'number' || isNaN(count)) {
      throw new Error("requestCharCount() did not return a valid number.");
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet(); // または特定のシート名: spreadsheet.getSheetByName("シート名");

    // 今日の日付と時刻を取得
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), DATE_FORMAT);
    const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), TIME_FORMAT);

    // 最終行を取得 (A列のデータがある最後の行)
    const lastRow = sheet.getLastRow();

    // 今日の日付が既にA列に存在するかチェック
    let rowToInsert = -1;
    for (let i = 1; i <= lastRow; i++) {
      const dateValue = sheet.getRange(i, DATE_COLUMN).getValue(); // A列のi行目の値を取得

      //日付の比較方法を修正. getValue()で取得した日付はDateオブジェクト。
      if (dateValue instanceof Date && Utilities.formatDate(dateValue, Session.getScriptTimeZone(), DATE_FORMAT) === formattedDate) {
        rowToInsert = i;
        break; // 今日の日付が見つかったのでループを抜ける
      }
    }

    //今日のデータがない場合は、新しい行にデータを書き込む
    if (rowToInsert === -1) {
      rowToInsert = lastRow + 1;
      // A列に日付をセット
      sheet.getRange(rowToInsert, DATE_COLUMN).setValue(formattedDate);
    }

    // B列に時刻をセット
    sheet.getRange(rowToInsert, TIME_COLUMN).setValue(formattedTime);

    // D列に文字数をセット
    sheet.getRange(rowToInsert, COUNT_COLUMN).setValue(count);


    // C列に前日比を計算する数式をセット (2行目以降)
    if (rowToInsert > 1) {
      const formula = `=IFERROR(D${rowToInsert}-D${rowToInsert - 1}, "N/A")`;  // 数式を設定
      sheet.getRange(rowToInsert, DIFFERENCE_COLUMN).setFormula(formula);
    } else {
      // 最初の行の場合、数式は設定しない（または、必要に応じて別の処理）
      sheet.getRange(rowToInsert, DIFFERENCE_COLUMN).setValue("N/A");
    }

  } catch (error: any) {
    // エラーハンドリング (エラーが発生した場合の処理)
    Logger.log("Error: " + error.message);
    Browser.msgBox("Error: " + error.message); // ユーザーにエラーを表示
  }
}

(global as any).executeCount = executeCount;
