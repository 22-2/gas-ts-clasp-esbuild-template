import { DATE_FORMAT, TIME_FORMAT, COL_DATE, COL_TIME, COL_COUNT, COL_DIFFERENCE } from "./constants";
import { requestCharCount } from "./count-characters";

export function hourlyRecord(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const count = requestCharCount();

  if (typeof count !== "number" || isNaN(count)) {
    throw new Error("requestCharCount() は有効な数値を返しませんでした。");
  }

  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), DATE_FORMAT);
  const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), TIME_FORMAT);

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, COL_DATE).setValue(formattedDate); // 日付
  sheet.getRange(newRow, COL_TIME).setValue(formattedTime); // 時刻
  sheet.getRange(newRow, COL_COUNT).setValue(count); // 文字数

  // 差分の計算 (2行目以降)
  if (newRow > 1) {
    const previousCount = sheet.getRange(newRow - 1, COL_COUNT).getValue();
    const formula = `=IFERROR(${count} - ${previousCount}, "N/A")`;
    sheet.getRange(newRow, COL_DIFFERENCE).setFormula(formula);
  } else {
    sheet.getRange(newRow, COL_DIFFERENCE).setValue("N/A"); // 最初の行は "N/A"
  }
}

export function dailyRecord(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  const count = requestCharCount(); // 文字数を取得

  if (typeof count !== "number" || isNaN(count)) {
    throw new Error("requestCharCount() は有効な数値を返しませんでした。");
  }

  // 現在の日付と時刻を取得
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), DATE_FORMAT);

  // 最終行を取得
  const lastRow = sheet.getLastRow();

  // 今日の日付が既に記録されているか確認
  let rowToInsert = -1;
  for (let i = 1; i <= lastRow; i++) {
    const dateValue = sheet.getRange(i, COL_DATE).getValue();
    if (dateValue instanceof Date && Utilities.formatDate(dateValue, Session.getScriptTimeZone(), DATE_FORMAT) === formattedDate) {
      rowToInsert = i;
      break;
    }
  }

  //今日のデータがない場合は、新しい行にデータを書き込む
  if (rowToInsert === -1) {
    rowToInsert = lastRow + 1;
    sheet.getRange(rowToInsert, COL_DATE).setValue(formattedDate); // 日付
  }

  sheet.getRange(rowToInsert, COL_COUNT).setValue(count); //文字数

  // 差分の計算 (2行目以降、かつ前日が異なる場合のみ)
  if (rowToInsert > 1) {
    let previousCount: number | "N/A" = "N/A";
    for (let i = rowToInsert - 1; i >= 1; i--) {
      const prevDateValue = sheet.getRange(i, COL_DATE).getValue();
      if (prevDateValue instanceof Date && Utilities.formatDate(prevDateValue, Session.getScriptTimeZone(), DATE_FORMAT) !== formattedDate) {
        previousCount = sheet.getRange(i, COL_COUNT).getValue();
        break;
      }
    }

    if (typeof previousCount === "number") {
      const formula = `=IFERROR(${count} - ${previousCount}, "N/A")`;
      sheet.getRange(rowToInsert, COL_DIFFERENCE).setFormula(formula);
    } else {
      sheet.getRange(rowToInsert, COL_DIFFERENCE).setValue("N/A");
    }
  } else {
    sheet.getRange(rowToInsert, COL_DIFFERENCE).setValue("N/A"); // 最初の行は "N/A"
  }
}
