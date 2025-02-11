import { requestCharCount } from "./count-characters";

function executeCount() {
  try {
    const count = requestCharCount(); // 文字数を取得 (この関数の実装は別途必要)

    if (typeof count !== 'number' || isNaN(count)) {
      throw new Error("requestCharCount() did not return a valid number.");
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet(); // または特定のシート名: spreadsheet.getSheetByName("シート名");

    // 今日の日付を取得 (YYYY/MM/DD形式)
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy/MM/dd");

    // 最終行を取得 (A列のデータがある最後の行)
    const lastRow = sheet.getLastRow();

    // 今日の日付が既にA列に存在するかチェック
    let rowToInsert = -1;
    for (let i = 1; i <= lastRow; i++) {
      const dateValue = sheet.getRange(i, 1).getValue(); // A列のi行目の値を取得

      //日付の比較方法を修正. getValue()で取得した日付はDateオブジェクト。
      if (dateValue instanceof Date && Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "yyyy/MM/dd") === formattedDate) {
        rowToInsert = i;
        break; // 今日の日付が見つかったのでループを抜ける
      }
    }

    //今日のデータがない場合は、新しい行にデータを書き込む
    if (rowToInsert === -1) {
      rowToInsert = lastRow + 1;
      // A列に日付をセット
      sheet.getRange(rowToInsert, 1).setValue(formattedDate);
    }


    // C列に文字数をセット
    sheet.getRange(rowToInsert, 3).setValue(count);

    // B列に前日比を計算してセット (前日のデータがある場合のみ)
    if (rowToInsert > 1) {
      const previousCount = sheet.getRange(rowToInsert - 1, 3).getValue(); // 前日の文字数
      if (typeof previousCount === 'number') {  //数値であることのチェックを追加
        const difference = count - previousCount;
        sheet.getRange(rowToInsert, 2).setValue(difference);
      } else {
        sheet.getRange(rowToInsert, 2).setValue("N/A"); //前日のデータがない場合、"N/A"などをセット
      }
    } else {
      sheet.getRange(rowToInsert, 2).setValue("N/A"); // 一番最初のデータの場合
    }


    // (オプション) 書式設定など、必要に応じて追加の処理を行う
    // 例: sheet.getRange(lastRow + 1, 1, 1, 3).setFontWeight("bold"); // 新しい行を太字にする
    // 例：数値のフォーマットなど.

  } catch (error: any) {
    // エラーハンドリング (エラーが発生した場合の処理)
    Logger.log("Error: " + error.message);
    Browser.msgBox("Error: " + error.message); // ユーザーにエラーを表示
  }
}

(global as any).executeCount = executeCount;
