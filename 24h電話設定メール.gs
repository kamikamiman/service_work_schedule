/* 
   担当が同じ場合はお知らせを無効にする。
   月を跨ぐ場合の処理の確認。
*/

function transferSettings() {
  
  // スプレットシートの情報を取得する。
  var ssGet = SpreadsheetApp.openById('1Wf2nEZEh4YfiKSfn2iNfBIs8hcxsFdYBBI8o6vwJYxY'); // 【サービス作業予定表】
  var period = 69; // 第〇〇期
  // 本日の日付を取得する。
  var date = new Date();
  var nowDay = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');
  // 本日の月を取得する。
  var nowMonth = Utilities.formatDate(date, 'Asia/Tokyo', 'M');
  // その月のシート情報を取得する。
  var schedule = ssGet.getSheetByName('${period}期${nowMonth}月'
                                      .replace('${period}', period)
                                      .replace('${nowMonth}', nowMonth));

  // 前月の月を取得する。
  var beforeMonth;
  if ( nowMonth !== 1 ) { 
     beforeMonth = nowMonth -1;
  } else {
     beforeMonth = 12;
  }
  
  // その月の前月のシート情報を取得する。
  var beforeSchedule = ssGet.getSheetByName('${period}期${beforeMonth}月'
                                      .replace('${period}', period)
                                      .replace('${beforeMonth}', beforeMonth));
  var lastColumn = beforeSchedule.getLastColumn();
  
//  Logger.log(beforeMonth);
//  Logger.log(lastColumn);
 
  // 日付データの取得範囲を指定する。
  const firstRow = 6;
  const _lastCol = schedule.getLastColumn();
  const lastCol = _lastCol - 1;
  console.log(lastCol);
  // 選択された日付データを取得する。
  var _getDays = schedule.getRange(firstRow, 2, 1, lastCol); 
  var getDays = _getDays.getValues();
  // 変数の初期設定。
  var number1 = 0;
  var dayNumber = 0;
  var dayOfWeek = ''; // 金曜日の番号取得が目的。
  
  // 本日の日付のセルの列番号を取得する。
  getDays.forEach(function(_getDay){
    _getDay.forEach(function(getDay){
      var day = (Utilities.formatDate(getDay, 'Asia/Tokyo', 'M/d')); // 日付の行の情報を取得する。
      if(day === nowDay) {                                           // 本日の日付と一致したら、
         dayNumber = number1;                                        // その日付のセルの列番号を取得し、
         dayOfWeek = new Date(getDay).getDay();                      // 曜日番号を取得する。
      }
      number1 += 1; 
    });  
  });
  
  
  /* 
  dayNumber: 説の列番号
  getLastColumn：　最終列
  */

  // 本日の夜勤当番管轄を取得する。
  const _names = schedule.getRange(5, dayNumber + 2, 1, 1);
  const  names = _names.getValues().flat();
  let  name;
  names.forEach(el => name = el);
  
  // 前日の夜勤当番管轄を取得する。
  let beforeNumber;
  let beforeName;
  let _beforeNames;
  let beforeNames;
  
  if ( dayNumber !== 1 ) {
    beforeNumber = dayNumber - 1;    
    _beforeNames = schedule.getRange(5, beforeNumber + 2, 1, 1);
  } else {
    _beforeNames = beforeSchedule.getRange(5, lastColumn + 2, 1, 1);
  }
  beforeNames = _beforeNames.getValues().flat();
  beforeNames.forEach(el => beforeName = el);
//  Logger.log(name);
//  Logger.log(beforeName);
  
      
  // 夜勤当番者が記入されているセルの色が黄色でなければ、メールを送信する。
  var _getCell = schedule.getRange(4, dayNumber + 2, 1, 1);
  var _getCellFri = schedule.getRange(5, dayNumber + 2, 1, 1);
  var getCellFri = _getCellFri.getValue();
  var getCellBackground = _getCell.getBackground();
  var getCellBackgroundFri = _getCellFri.getBackground();
  if(dayOfWeek !== 0) {         // 本日が土曜日以外なら実行する。
    if(dayOfWeek !== 6) {       // 本日が日曜日以外なら実行する。
      if(name !== beforeName) { // 当番が前日と管轄が違う場合に実行する。
        if(getCellFri !== '') {
          if(getCellBackground !== '#ffff00' || getCellBackgroundFri == '#ffff00'){
            // 夜勤当番者に電話が繋がるかの確認依頼メールを送信する。
            // メールの送信先
            var to = 'technical-support@isowa.co.jp';
//            var to = 'k.kamikura@isowa.co.jp';
            // メールのタイトル
            var subject = '24時間当番の電話転送設定の確認依頼';
            // メールの本文
            var body = '\
テクニカルサポートの皆様\n\n\
お仕事お疲れ様です。\n\
本日、24時間当番の電話転送設定の確認が取れていません。\n\
お手間ですが、転送設定の確認をお願いします。\n\n\n\
よろしくお願いします。';
          
            var options = { name: 'ISOWA_support',
                            bcc: 'k.kamikura@isowa.co.jp'};
          
            GmailApp.sendEmail(
              to,
              subject,
              body,
              options
            );
         };
       };
     };
   };
 };

  // ログ確認用
//  Logger.log(getCellFri);
//  Logger.log('---------------');  
//  Logger.log(dayNumber);
//  Logger.log('---------------');
  Logger.log(dayOfWeek);
//  Logger.log('---------------');
//  Logger.log(getCellBackground);
//  Logger.log('---------------');
//  Logger.log(getCellBackgroundFri);
//  Logger.log('---------------');  
  
}










