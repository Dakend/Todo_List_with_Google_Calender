//****************************************
//関数名：taskListUpdate
//機能：実行中タスクリストのCloseされているタスクを完了済タスクリストに移行させる
//引数：なし
//戻り値：なし
//****************************************
function taskListUpdate() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTaskList = ss.getSheetByName("実行中タスクリスト");
    const sheetClosedList = ss.getSheetByName("完了済タスクリスト");

    //実行中タスクリストのE3から最後の記載までの範囲のデータを取得し、keySearchListに格納
    const taskTableRangeStartRow = 3;
    const taskTableRangeStartColumn = 5;
    const taskTableRangeRowOffset = sheetTaskList.getLastRow() - 2;
    const taskTableRangeColumnOffset = 1;
    const keySearchList = sheetTaskList
      .getRange(
        taskTableRangeStartRow,
        taskTableRangeStartColumn,
        taskTableRangeRowOffset,
        taskTableRangeColumnOffset
      )
      .getValues();

    //完了済タスクリスト記載の最後のタスクの次の行を変数に代入
    const pushRow = sheetClosedList.getLastRow() + 1;
    const count = 0;

    //次にCloseタスクor空白の行を削除するため、一致したときにその行数を追加
    const deleteRowList = [];

    //完了済タスクリストに追加する行数をカウント
    const insertRowCountForCloseList = 0;

    //完了済タスクリストのA2~E2までの行・列を変数化
    const closeTableStartRow;
    const closeTableStartColum = 1;
    const closeTableStartRowOffeset = 1;
    const closeTableColumnOffset = 5;

    //実行中タスクリストにCloseのタスクが存在しているかどうか
    const existsCloseTask = false;

    //実行中タスクリストを1行づつ取得するために値を代入
    const taskTableStartRow;
    const taskTableStartColumn = 2;
    const taskTableRowOffset = 1;
    const taskTableColumnOffset = 5;

    const startTaskTableRowOffset = 3;
    for (var i = 0; i < keySearchList.length; i++) {
      if (keySearchList[i] == "Close") {
        closeTableStartRow = pushRow + count;
        taskTableStartRow = i + startTaskTableRowOffset;

        //実行中タスクリストにおけるCloseタスクを完了済タスクリストにセット
        sheetClosedList
          .getRange(
            closeTableStartRow,
            closeTableStartColum,
            closeTableStartRowOffeset,
            closeTableColumnOffset
          )
          .setValues(
            sheetTaskList
              .getRange(
                taskTableStartRow,
                taskTableStartColumn,
                taskTableRowOffset,
                taskTableColumnOffset
              )
              .getValues()
          );

        //Closeタスクの行をdeleteRowList(配列型)に追加
        deleteRowList.push(taskTableStartRow);

        //完了済タスクリストに追加する行数カウントをインクリメント
        insertRowCountForCloseList = insertRowCountForCloseList + 1;

        //次の行をサーチさせるため、インクリメント
        count = count + 1;

        //Closeタスクが存在していたので、trueを代入
        existsCloseTask = true;
      }
    }

    //実行中タスクリストのNo.の空白である行数を配列deleteRowListに追加
    const taskNumberList = getNumberList();
    //実行中タスクにCloseのタスクが存在しているかどうか
    let existsBlankTask = false;
    for (let i = 0; i < taskNumberList.length; i++) {
      if (taskNumberList[i] == "") {
        deleteRowList.push(i + startTaskTableRowOffset);
        existsBlankTask = true;
      }
    }

    if (existsCloseTask || existsBlankTask) {
      //実行中タスクリストにおけるCloseタスク行及び空白行を順次削除
      //deleteRowListの中身をソートしてから削除開始
      deleteRowList.sort(compareNumbers);
      for (let i = 0; i < deleteRowList.length; i++) {
        //削除すると指定行が-1されるのでiでデクリメント
        sheetTaskList.deleteRow(deleteRowList[i] - i);
      }

      //実行中タスクリスト最終行に空白行削除分(deleteRowListの要素数)だけ行を挿入
      const inserTaskListtRow = sheetTaskList.getRange("A:A").getLastRow();
      sheetTaskList.insertRows(inserTaskListtRow, deleteRowList.length);

      if (existsCloseTask) {
        //完了済タスクリスト最終行に追加分(=実行中タスクリスト削除分)だけ行を挿入
        const insertCloseListRow = sheetClosedList.getRange("A:A").getLastRow();
        sheetClosedList.insertRows(
          insertCloseListRow,
          insertRowCountForCloseList
        );

        //実行中タスクリストのNo.を採番(A3からAxまで)
        const startNumberingRow = 2;
        const startNumberingColumn = 1;
        startTaskTableRowOffset = 2;
        const taskTableLastRow = sheetTaskList.getLastRow() - 2;
        for (let i = 1; i <= taskTableLastRow; i++) {
          startNumberingRow = i + startTaskTableRowOffset;
          sheetTaskList
            .getRange(startNumberingRow, startNumberingColumn)
            .setValue(i);
        }
      }
    }
  } catch (e) {
    return;
  }
}

//****************************************
//関数名：compareNumbers
//機能：数値の昇順ソート用関数
//引数：a(number型), b(number型)
//戻り値：-
//****************************************
//数値の昇順ソート用関数
function compareNumbers(a, b) {
  return a - b;
}
