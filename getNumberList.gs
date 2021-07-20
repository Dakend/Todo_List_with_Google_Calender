//****************************************
//関数名：getNumberList
//機能：タスクリストのNo.1から最後までのNoを配列に格納し戻す
//引数：なし
//戻り値：numberList(配列)
//****************************************
function getNumberList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTaskList = ss.getSheetByName("実行中タスクリスト");
  //実行中タスクリストのA3から最後の記載までの範囲のタスクデータを取得し、numberList(2次元配列)に格納
  const taskTableStartRow = 3;
  const taskTableStartColumn = 1;
  const taskTableRowOffset = sheetTaskList.getLastRow() - 2;
  const taskTableColumnOffset = 1;
  const numberList = sheetTaskList
    .getRange(
      taskTableStartRow,
      taskTableStartColumn,
      taskTableRowOffset,
      taskTableColumnOffset
    )
    .getValues();
  return numberList;
}
