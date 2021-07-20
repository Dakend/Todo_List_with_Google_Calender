//****************************************
//関数名：getMemberList
//機能：名簿からデータを1行づつオブジェクト型変数に格納し返す
//引数：なし
//戻り値：memberList(オブジェクト型)
//****************************************
function getMemberList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMemberList = ss.getSheetByName("名簿");

  //名簿のA2から最後の記載までの範囲のタスクデータを取得し、memberTable(2次元配列)に格納
  const memberTableStartRow = 2;
  const memberTableStartColumn = 1;
  const memberTableRowOffset = sheetMemberList.getLastRow() - 1;
  const memberTableColumnOffset = 2;
  const memberTable = sheetMemberList
    .getRange(
      memberTableStartRow,
      memberTableStartColumn,
      memberTableRowOffset,
      memberTableColumnOffset
    )
    .getValues();

  //memberList(オブジェクト型)を作成して、名簿のA1,B1の表記に紐づけて、メンバーデータを追加
  const memberList = {};
  const nameIndex = 0;
  const addressIndex = 1;
  for (var i = 0; i < memberTable.length; i++) {
    memberList[memberTable[i][nameIndex]] = memberTable[i][addressIndex];
  }
  return memberList;
}
