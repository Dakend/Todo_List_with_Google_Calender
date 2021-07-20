//****************************************
//関数名：mainOperate
//機能：カレンダー操作に必要なクラスからインスタンスを生成、
//     登録・更新・削除に応じて、カレンダーを操作
//引数：operateType(String型)
//戻り値：なし
//****************************************
function mainOperate(operateType) {
  constlock = LockService.getDocumentLock();
  if (lock.tryLock(10000)) {
    try {
      constnumber = getInputNumber(operateType);
      verifyInputNumber(number);
      constmemberList = getMemberList();
      consttaskList = getTaskList();
      verifyItem(number, taskList);
      consttheCalendar = new OperateCalendar(number, memberList, taskList);
      switch (operateType) {
        case "登録/更新":
          theCalendar.modifyEvent();
          break;
        case "削除":
          theCalendar.deleteEvent();
          break;
      }
      Browser.msgBox("正常に処理されました。");
    } catch (e) {
      Browser.msgBox("!! 操作がキャンセルされました。!!");
      return;
    } finally {
      lock.releaseLock();
    }
  } else {
    Browser.msgBox(
      "現在Lockされています。\\n しばらくしてから再度実行させてください。"
    );
  }
}
