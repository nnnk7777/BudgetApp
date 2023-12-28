import { AnnualReport } from "./model/annualReport";

export function main() {
  const speadSheetName = "金銭メモテスト";
  const sheetName = "🐖 家計簿";

  // スプレッドシート、シートの初期化
  let spreadsheet;
  let sheet;

  // シート検索のための初期化
  const files = DriveApp.getFilesByName(speadSheetName);

  if (files.hasNext()) {
    // スプレッドシートが存在する場合、そのスプレッドシートを開き中身をクリア
    spreadsheet = SpreadsheetApp.open(files.next());
    sheet = spreadsheet.getSheets()[0];
    sheet.clear();
  } else {
    // スプレッドシートが存在しない場合、新しいスプレッドシートを作成
    spreadsheet = SpreadsheetApp.create(speadSheetName);
    // 新しいスプレッドシートにデフォルトで含まれるシートを取得
    sheet = spreadsheet.getSheets()[0];
    // シート名を変更
    sheet.setName(sheetName);
  }

  // columnOffsetを指定
  new AnnualReport(sheet);
}
