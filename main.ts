import { AnnualReport } from "./model/annualReport";

export function main() {
  const speadSheetName = "金銭メモ2024";
  const sheetName = "🐖 家計簿";

  // スプレッドシート、シートの初期化
  let spreadsheet;
  let sheet;
  // シート検索のための初期化
  const files = DriveApp.getFilesByName(speadSheetName);

  /**
   * シートの存在有無によって処理を分岐させる
   * - 存在する： 既存のシートに対してスタイルやバリデーションルールの再適用のみをする
   * - 存在しない： 新規シートを作成する
   */
  if (files.hasNext()) {
    console.log("同一名称のシートが存在する");

    // スプレッドシートが存在する場合、そのスプレッドシートを開き中身をクリア
    spreadsheet = SpreadsheetApp.open(files.next());
    sheet = spreadsheet.getSheets()[0];
    // sheet.clear();
  } else {
    console.log("同一名称のシートが存在しない");

    // スプレッドシートが存在しない場合、新しいスプレッドシートを作成
    spreadsheet = SpreadsheetApp.create(speadSheetName);
    // 新しいスプレッドシートにデフォルトで含まれるシートを取得
    sheet = spreadsheet.getSheets()[0];
    // シート名を変更
    sheet.setName(sheetName);
  }

  new AnnualReport(sheet);
}
