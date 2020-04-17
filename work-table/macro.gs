// 左側の年月
let leftYear;
let leftMonth;
// 右側の年月
let rightYear;
let rightMonth;

function reset() {
  // アクティブなシートを取得
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let regex = sheet.getName().match(/(\d+)年(\d+)月分/);

  // 年月を指定
  let year = parseInt(regex[1]);
  let month = parseInt(regex[2]);

  resetDate(sheet, year, month);
}

// 月日切換処理
function resetDate(sheet, year, month) {
  // 左側の年月
  leftYear = month == 1 ? year - 1 : year;
  leftMonth = month == 1 ? 12 : month - 1;
  // 右側の年月
  rightYear = year;
  rightMonth = month;

  // 表の月を記述
  sheet.getRange("A6").setValue(leftMonth);
  sheet.getRange("A17").setValue(rightMonth);
  sheet.getRange("S6").setValue(rightMonth);

  let cellsL = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "O"];
  let cellsR = ["S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AG"];
  common(sheet, cellsL, 21, leftYear, leftMonth, true);
  common(sheet, cellsR, 6, rightYear, rightMonth, false);
}

// 月日切換処理の共通部分
function common(sheet, cells, date, year, month, isLeft) {
  for(let i=6; i<=21; i++) {
    // 表の右側の21行目は休憩・実働・普通残業の合計の数式を設定
    if(i == 21 && !isLeft) {
      sheet.getRange(cells[11]+i).setValue(`=SUM(L6:L21) + SUM(AD6:AD20)`);
      sheet.getRange(cells[12]+i).setValue(`=SUM(M6:M21) + SUM(AE6:AE20)`);
      sheet.getRange(cells[13]+i).setValue(`=SUM(O6:O21) + SUM(AG6:AG20)`);
      break;
    }

    // 表の左側で月を切り替え
    if(i == 17 && isLeft) {
      year = rightYear;
      month = rightMonth;
      date = 1;
    }

    // 勤務開始・終了日時、休憩時間の定数
    let jobStartHour = "9";
    let jobStartMinute = "00";
    let jobEndHour = "18";
    let jobEndMinute = "00";
    let timeSplit = ":";
    let middle = "～";
    let restTime = "1.00";

    // 日本語の曜日を取得
    let dayStr = getDayStr(new Date(year + "/" + month + "/" + date).getDay());

    // 日付と曜日を設定
    sheet.getRange(cells[1]+i).setValue("／");
    sheet.getRange(cells[2]+i).setValue(date);
    sheet.getRange(cells[3]+i).setValue(dayStr);

    // 土日は背景色を付け、勤怠を設定しない
    if (dayStr == "土" || dayStr == "日") {
      sheet.getRange(cells[0]+i+":"+cells[3]+i).setBackground("cyan");
      jobStartHour = "";
      jobStartMinute = "";
      jobEndHour = "";
      jobEndMinute = "";
      timeSplit = "";
      middle = "";
      restTime = "";
    } else {
      // 平日は背景色なし、勤怠を設定（デフォルトは定時）
      sheet.getRange(cells[0]+i+":"+cells[3]+i).setBackground("white");
    }

    // 上記で設定した背景色と勤怠で既存データを上書き
    sheet.getRange(cells[4]+i).setValue(jobStartHour);
    sheet.getRange(cells[5]+i).setValue(timeSplit);
    sheet.getRange(cells[6]+i).setValue(jobStartMinute);
    sheet.getRange(cells[7]+i).setValue(middle);
    sheet.getRange(cells[8]+i).setValue(jobEndHour);
    sheet.getRange(cells[9]+i).setValue(timeSplit);
    sheet.getRange(cells[10]+i).setValue(jobEndMinute);
    sheet.getRange(cells[11]+i).setValue(restTime);

    // 実働、普通残業の数式を設定
    sheet.getRange(cells[12]+i).setValue(`=IF(${cells[11]}${i}=0.75,7.75,IF(${cells[4]}${i}="","",${cells[8]}${i}-${cells[4]}${i}-${cells[11]}${i}+IF(${cells[10]}${i}="30",0.5)))`);
    sheet.getRange(cells[13]+i).setValue(`=IF(${cells[5]}${i}="","",IF(${cells[12]}${i}<=8,"",${cells[12]}${i}-8+0.25))`);

    // 日付を更新
    date = date + 1;
  }
}

// 日本語の曜日を取得
function getDayStr(day) {
  let dayArray = ['日', '月', '火', '水', '木', '金', '土'];
  return dayArray[day];
}
