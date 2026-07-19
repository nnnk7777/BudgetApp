// GAS の実行画面から、既存シートの書式・入力規則・条件付き書式を再構築する入口。
function reapplySheetStyleManual() {
    Logger.log("シートスタイルの再適用を開始します");
    main();
    Logger.log("シートスタイルの再適用が完了しました");
    return "Sheet style reapplied";
}
