//****************************************
//関数名：autoModifyTask
//機能：タスクリストのNo.1から最後までの全てのタスクをカレンダー更新する
//引数：なし
//戻り値：なし
//****************************************
function autoModifyTask() {
  try {
    const numberList = getNumberList();
    const memberList = getMemberList();
    const taskList = getTaskList();
    const openTaskList = [];
    const closeTaskList = [];

    for (let i = 0; i < numberList.length; i++) {
      if (!(numberList[i] == "")) {
        if (taskList[numberList[i]]["ステータス*"] == "Open") {
          openTaskList.push(
            new OperateCalendar(numberList[i], memberList, taskList)
          );
        } else {
          closeTaskList.push(
            new OperateCalendar(numberList[i], memberList, taskList)
          );
        }
      }
    }

    for (let i = 0; i < openTaskList.length; i++) {
      openTaskList[i].modifyEvent();
    }

    for (let i = 0; i < closeTaskList.length; i++) {
      closeTaskList[i].deleteEvent();
    }
  } catch (e) {
    return;
  }
}
