//****************************************
//関数名：verifyItem
//機能："実行中タスクリスト"シートの指定タスクの未入力を検証
//引数：number(number型)、taskList(object型)
//戻り値：なし
//****************************************
function verifyItem(number, taskList) {
  if (
    taskList[number]["No*"] == "" ||
    taskList[number]["概要*"] == "" ||
    taskList[number]["担当者*"] == "" ||
    taskList[number]["ステータス*"] == "" ||
    taskList[number]["期日*"] == ""
  ) {
    Browser.msgBox("！！ 必須項目が入力されていません。！！");
    throw new Error();
  }
}
