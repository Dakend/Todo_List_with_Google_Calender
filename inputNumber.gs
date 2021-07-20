//****************************************
//関数名：getInputNumber
//機能：ユーザー入力画面の表示
//引数：operateType(string型)
//戻り値：number(number型)
//****************************************
function getInputNumber(operateType) {
  const confirmation = Browser.msgBox(
    "確認画面",
    "該当カレンダーに" + operateType + "しますか？",
    Browser.Buttons.OK_CANCEL
  );
  if (confirmation == "cancel") {
    throw new Error();
  }
  const number = Browser.inputBox(
    "ナンバー入力画面",
    operateType + "したいNo.を入力してください \\n ※半角数字のみ",
    Browser.Buttons.OK_CANCEL
  );
  if (number == "cancel") {
    throw new Error();
  }
  return number;
}

//****************************************
//関数名：verifyInputNumber
//機能：ユーザー入力の文字をタスクリストにあるものか、半角数字か検証
//引数：number(number型)
//戻り値：なし
//****************************************
function verifyInputNumber(number) {
  const numberListForVerification = getNumberList();
  const isNumber = isFinite(number);
  const numberMatchFlag = false;
  for (let i = 0; i < numberListForVerification.length; i++) {
    if (numberListForVerification[i] == number) {
      numberMatchFlag = true;
      break;
    }
  }
  if (!(numberMatchFlag && isNumber)) {
    Browser.msgBox(
      "!! 入力された値が、タスクリストに存在しない、もしくは半角数字ではありません。!!"
    );
    throw new Error();
  }
}
