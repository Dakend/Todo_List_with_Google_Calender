//****************************************
//関数名：dateResearch
//機能：受け取った日程を本日と期限日の期間に応じて適切に通知日(3日前/1日前/当日)を設定し返す
//引数：date(date型オブジェクト)
//戻り値：date型オブジェクト
//****************************************
function dateResearch(date) {
  //本日の値を取得、yyyy/mm/ddのみを抽出
  const today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  //営業日まで日を戻してくれるgetWorkDate関数で処理したdate型オブジェクトを新規変数に代入、date型変数の参照を切るために、newしている.
  //受け取った引数dateを営業日処理したものを変数に代入
  const normalizedDateList = [];
  const registerCount = 3;
  //3日前・1日前・当日通知用data型オブジェクトをnewし、Listに格納
  for (let i = 0; i < registerCount; i++) {
    normalizedDateList.push(new Date(getWorkDate(date).getTime()));
  }

  //本日todayを営業日処理したものを変数に代入
  const normalizedToday = new Date(getWorkDate(today).getTime());
  const dayGap = (normalizedDateList[0] - normalizedToday) / 86400000;
  const normalizedDayGap =
    dayGap - getNotWorkDayCount(normalizedDateList[0], normalizedToday);

  //適正化された2つの日付の差(normalizedDayGap)を検査し、営業日処理されたnormalizedDateを3日前/1日前/当日の処理をして、
  //結果をregisterListをの戻り値として返す
  const firstDayBack = 3;
  const secondDayBack = 1;
  const registerDateList = [];

  if (normalizedDayGap >= firstDayBack) {
    //Browser.msgBox("3日前,1日前,当日に設定");
    registerDateList.push(setDeliveryDate(normalizedDateList[2], firstDayBack));
    registerDateList.push(
      setDeliveryDate(normalizedDateList[1], secondDayBack)
    );
    registerDateList.push(normalizedDateList[0]);
    return registerDateList;
  } else if (normalizedDayGap >= secondDayBack) {
    //Browser.msgBox("1日前に設定");
    registerDateList.push(
      setDeliveryDate(normalizedDateList[1], secondDayBack)
    );
    registerDateList.push(normalizedDateList[0]);
    return registerDateList;
  } else if (normalizedDayGap == secondDayBack) {
    //Browser.msgBox("当日に設定");
    registerDateList.push(normalizedDateList[0]);
    return registerDateList;
  } else {
    //Browser.msgBox("期限超過、本日に設定");
    registerDateList.push(normalizedToday);
    return registerDateList;
  }
}

//****************************************
//関数名：setDeliveryDate
//機能：指定日から指定カウント分バックする(バックした日が土日祝日の場合営業日化する)
//引数：date(date型オブジェクト), backCount(number型)
//戻り値：date(date型オブジェクト)
//****************************************
function setDeliveryDate(date, backCount) {
  //Browser.msgBox("最初の営業日" + normalizedDate);
  for (let i = 0; i < backCount; i++) {
    getWorkDate(getBackDate(date));
  }
  return date;
}

//****************************************
//関数名：getNotWorkDayCount
//機能：受け取った2つ日程の差から土日祝日を除いた差を求めて返す
//引数：compareDateX(date型オブジェクト), compareDateY(date型オブジェクト)
//戻り値：count(number型)
//****************************************
function getNotWorkDayCount(compareDateX, compareDateY) {
  const comparedDayGap = (compareDateX - compareDateY) / 86400000;
  //差を求めた日数の中から土日・祝日を除いた日数を求める
  const dayGapSearchDate = new Date(getWorkDate(compareDateX).getTime());
  const saturdayNumber = 6;
  const sundayNumber = 0;
  const holidayNumberFlag = 1;
  const dayGapOffset = 1;
  const count = 0;
  for (let i = 0; i < comparedDayGap - dayGapOffset; i++) {
    getBackDate(dayGapSearchDate);
    if (
      dayGapSearchDate.getDay() == sundayNumber ||
      dayGapSearchDate.getDay() == saturdayNumber ||
      getHolidayNumberFlag(dayGapSearchDate) == holidayNumberFlag
    ) {
      count = count + 1;
    }
  }
  return count;
}

//****************************************
//関数名：getBackDate
//機能：受け取った日程の前日の値を返す
//引数：date(date型オブジェクト)
//戻り値：date(date型オブジェクト)
//****************************************
function getBackDate(date) {
  date.setDate(date.getDate() - 1);
  return date;
}

//****************************************
//関数名：getHolidayNumberFlag
//機能：受け取った日程を祝日判定し、祝日であれば1を返す
//引数：date(date型オブジェクト)
//戻り値：holidayNumberFlag(number型)
//****************************************
function getHolidayNumberFlag(date) {
  const calendarJapan = CalendarApp.getCalendarById(
    "ja.japanese#holiday@group.v.calendar.google.com"
  ); 
  //国指定のカレンダー情報取得。
  const holidayNumberFlag = calendarJapan.getEventsForDay(date).length;
  return holidayNumberFlag;
}

//****************************************
//関数名：getWorkDate
//機能：受け取った日程がもし営業日でなければ、営業日まで遡って日程を返す
//引数：date(date型オブジェクト)
//戻り値：date(date型オブジェクト)
//****************************************
function getWorkDate(date) {
  const saturdayNumber = 6;
  const sundayNumber = 0;
  const holidayNumberFlag = 1;
  if (
    date.getDay() == sundayNumber ||
    date.getDay() == saturdayNumber ||
    getHolidayNumberFlag(date) == holidayNumberFlag
  ) {
    let isNotWorkDay = true;
    const backDate;
    const backDateNumberFlag;
    while (isNotWorkDay) {
      backDate = new Date(getBackDate(date).getTime());
      backDateNumberFlag = backDate.getDay();
      if (
        !(
          backDateNumberFlag == sundayNumber ||
          backDateNumberFlag == saturdayNumber ||
          getHolidayNumberFlag(date) == holidayNumberFlag
        )
      ) {
        isNotWorkDay = false;
      }
    }
  }
  return date;
}
