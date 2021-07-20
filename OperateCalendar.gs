//****************************************
//クラス名：OperateCalendar
//機能：指定タスクの担当者の各カレンダー操作(登録、削除、更新)をもつクラスを定義
//引数：number(number型), memberList(配列), taskList(配列)
//戻り値：OperateCalendarインスタンス
//****************************************

class OperateCalendar {
  constructor(number, memberList, taskList) {
    this.number = number;
    this.memberList = memberList;
    this.taskList = taskList;

    //指定タスクの担当者のカレンダー操作の準備をする
    const calendarId = this.memberList[this.taskList[this.number]["担当者*"]];
    this.calendar = CalendarApp.getCalendarById(calendarId);

    //カレンダー検索時における検索範囲を設定. 今日から前後6ケ月で設定
    const startTime = new Date();
    const endTime = new Date();
    const monthScope = 6;
    startTime.setMonth(startTime.getMonth() - monthScope);
    endTime.setMonth(endTime.getMonth() + monthScope);

    //操作するイベントタイトルの終日イベント(確認用)と時間設定イベント(通知用) ※登録・更新・削除共通
    this.titleForAllDay =
      this.taskList[this.number]["概要*"] +
      "：" +
      this.taskList[this.number]["担当者*"];
    this.titleForPush = this.taskList[this.number]["概要*"];

    //上記設定範囲にある全てのイベントを変数eventList(配列)に格納
    this.eventList = this.calendar.getEvents(startTime, endTime);
  }

  //*************************
  //registerNewEventメソッド
  //機能：指定タスクの登録. 実行中タスクリストにおける概要、期日、詳細をカレンダーのタイトル、通知日、詳細としてセット
  //　　　期日は、外部関数dateResearchに記載日を渡して適正値を受け取り、値をセット. 終日登録(確認用)と//時間指定登録(通知用)を行う
  //引数：なし
  //戻り値：なし
  //*************************
  registerNewEvent() {
    //通知用イベントの時刻を開始18:00~終了18:00に設定、通知時刻を530分前(09:10)に設定
    //タスクリスト
    this.dateList = dateResearch(this.taskList[this.number]["期日*"]);
    this.registerHour = "18";
    this.remindMinuteBefore = 530;

    for (let i = 0; i < this.dateList.length; i++) {
      this.notificationStartEndTime = new Date(this.dateList[i]);
      this.notificationStartEndTime.setHours(this.registerHour);
      this.option = { description: this.taskList[this.number]["詳細"] };

      //終日登録(確認用)
      this.calendar.createAllDayEvent(
        this.titleForAllDay,
        this.dateList[i],
        this.option
      );

      //時間指定登録(通知用)
      this.calendar
        .createEvent(
          this.titleForPush,
          this.notificationStartEndTime,
          this.notificationStartEndTime
        )
        .addPopupReminder(this.remindMinuteBefore);
    }
  }

  //*************************
  //modifyEventメソッド
  //変数eventListに格納したイベントから指定したタスクのタイトルを検索し一致したものを削除、新規登録
  //検索にヒットしなかった場合は、新規登録としてタスク登録メソッドを実行
  //引数：なし
  //戻り値：なし
  //*************************
  modifyEvent() {
    //指定タスクの概要を一致したものを削除、registerNewEventを呼び出し、新規登録
    this.deleteEvent();
    this.registerNewEvent();
  }

  //*************************
  //deleteEventメソッド
  //変数eventListに格納したイベントから指定したタスクのタイトルを検索し一致したものを削除
  //引数：なし
  //戻り値：なし
  //*************************
  deleteEvent() {
    //変数eventsArrayTitle(配列)にeventListにあるイベントのタイトルのみを順次格納
    //指定タスクの概要を一致したものを削除
    this.eventsArrayTitle = [];
    for (let i = 0; i < this.eventList.length; i++) {
      this.eventsArrayTitle.push(this.eventList[i].getTitle());
      if (
        this.eventsArrayTitle[i] == this.titleForAllDay ||
        this.eventsArrayTitle[i] == this.titleForPush
      ) {
        this.eventList[i].deleteEvent();
      }
    }
  }
}
