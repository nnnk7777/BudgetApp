import { columnToLetter } from "./columnToletter";

// 列番号と行番号を受け取り、そのセルの表記（例: "A1"）を返す関数
export function getCellNotation(column: number, row: number) {
    // 列のアルファベット表記を取得する（例: 1 -> A, 2 -> B）
    var columnLetter = columnToLetter(column);
    return columnLetter + row;
}
