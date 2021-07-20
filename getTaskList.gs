//****************************************
//関数名：getTaskList
//機能：実行中タスクリストからデータを1行づつオブジェクト型変数に格納し返す
//引数：なし
//戻り値：taskList(オブジェクト型)
//****************************************
function getTaskList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTaskList = ss.getSheetByName("実行中タスクリスト");

  //実行中タスクリストのA3から最後の記載までの範囲のタスクデータを取得し、taskTable(2次元配列)に格納
  const taskTableStartRow = 3;
  const taskTableStartColumn = 1;
  const taskTableRowOffset = sheetTaskList.getLastRow() - 2;
  const taskTableColumnOffset = 6;
  const taskTable = sheetTaskList
    .getRange(
      taskTableStartRow,
      taskTableStartColumn,
      taskTableRowOffset,
      taskTableColumnOffset
    )
    .getValues();

  //taskList(オブジェクト型)を作成して、実行中タスクリストのA2,B2,C2,D2,E2,F2の表記に紐づけて、タスクデータを追加
  const numberIndex = 0;
  const overViewIndex = 1;
  const detailsIndex = 2;
  const nameIndex = 3;
  const statusIndex = 4;
  const deliveryIndex = 5;
  const taskList = {};
  for (let i = 0; i < taskTable.length; i++) {
    taskList[taskTable[i][numberIndex]] = {
      "概要*": taskTable[i][overViewIndex],
      詳細: taskTable[i][detailsIndex],
      "担当者*": taskTable[i][nameIndex],
      "ステータス*": taskTable[i][statusIndex],
      "期日*": taskTable[i][deliveryIndex],
    };
  }
  return taskList;
}
