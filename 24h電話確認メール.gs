function mail24hours() {
  
  // スプレットシートの情報を取得する。
  var ssGet = SpreadsheetApp.openById('1Wf2nEZEh4YfiKSfn2iNfBIs8hcxsFdYBBI8o6vwJYxY'); // 【サービス作業予定表】
  var period = 69; // 第〇〇期
  var date = new Date();
  var nowDay = Utilities.formatDate(date, 'Asia/Tokyo', 'M/d');
  
  // 本日の月を取得する。
  var nowMonth = Utilities.formatDate(date, 'Asia/Tokyo', 'M');
  // その月のシート情報を取得する。
  var schedule = ssGet.getSheetByName('${period}期${nowMonth}月'
                                      .replace('${period}', period)
                                      .replace('${nowMonth}', nowMonth));
 
  // データの取得範囲を指定する。
  var firstRowDays = 6;    // セル選択開始行（日付データを取得する）
  var firstRowNameFri = 5; // セル選択開始行（金曜日の当番をデータ取得する目的）
  var firstRowName = 4;    // セル選択開始行（当番データを取得する）
  var lastColumn = 30;     // セル選択終了列 　要修正必要
  // 日付データを取得する。
  var _getDays = schedule.getRange(firstRowDays, 2, 1, lastColumn); 
  var getDays = _getDays.getValues();
     // 金曜日当番データ取得する目的
  var _getDaysFri = schedule.getRange(firstRowNameFri, 2, 1, lastColumn); 
  var getDaysFri = _getDaysFri.getValues(); 
  // 変数の初期設定。
  var number1 = 0;
  var dayNumber = 0;
  var dayOfWeek = ''; // 金曜日の番号取得が目的。
  
  // 本日の日付のセルの列番号を取得する。
  getDays.forEach(function(_getDay){
    _getDay.forEach(function(getDay){
      var day = (Utilities.formatDate(getDay, 'Asia/Tokyo', 'M/d')); // 日付の行の情報を取得する。
      if(day === nowDay) {                                           // 本日の日付と一致したら、
        dayNumber = number1;                                         // その日付のセルの列番号を取得する。
        dayOfWeek = new Date(getDay).getDay();
      }
      number1 += 1; 
    });
  });
    
  // 選択された範囲の夜勤当番者データを取得する。
  var _getNames = schedule.getRange(firstRowName, 2, 1, lastColumn);
  var getNames = _getNames.getValues();
  // 選択された範囲の夜勤当番者データ（金曜日 5行目）を取得する。
  var _getNamesFri = schedule.getRange(firstRowNameFri, 2, 1, lastColumn)
  var getNamesFri = _getNamesFri.getValues();
  // 変数の初期設定。
  var number2 = 0;
  var number4 = 0;
  var nameNumber = 0;
  var nightShiftDutys1 = ''; // 夜勤当番者2名
  var nightShiftDutys2 = ''; // 金曜当番者2名
  
  // 本日の日付のセルの列番号と同じ列にある人の名前を取得する。
  getNames.forEach(function(_getName){
    _getName.forEach(function(getName){
      if(dayNumber === number2) {                // 本日の日付のセル列番号と一致したら、
        nightShiftDutys1 = _getName[dayNumber];   // その列番号の夜勤当番者名(2名分)を取得する。
      }
      number2 += 1;
    });  
  });
  
  
  // 担当者名を個別で取得する。
  // nightShiftDutysの変数に'/'が含まれている場合に実行する。
  if ( nightShiftDutys1.indexOf('/') != -1) {
    var nightShiftDuty1 = nightShiftDutys1.split('/')[0];        // 1人目
    var nightShiftDuty2 = nightShiftDutys1.split('/')[1];        // 2人目
    var sheetAdress = ssGet.getSheetByName('メールアドレス（24h）');// スプレットシート情報を取得
    // メールアドレス一覧から送信対象者の情報を取得する。
    var names = sheetAdress.getRange(2, 2, 50, 2).getValues();
    // スプレットシートの行と列を反転させる。
    var _ = Underscore.load();
    var namesTrans = _.zip.apply(_, names);
    // 選択された送信対象者から送信先メールアドレスの情報を取得する。
    var selectedName1 = nightShiftDuty1;
    var selectedName2 = nightShiftDuty2;
    // 本日が金曜日だった場合は、selectedName1 ,selectedName2を上書きする。
    getNamesFri.forEach(function(_getNameFri){
      _getNameFri.forEach(function(getNameFri){
        if(dayOfWeek == 5) { // 本日の日付のセル列番号と一致し、金曜日の場合、
          nightShiftDutys2 = _getNameFri[dayNumber];   // その列番号の夜勤当番者名(2名分)を取得する。
          selectedName1 = nightShiftDutys2.split('/')[0]; // 1人目
          selectedName2 = nightShiftDutys2.split('/')[1]; // 2人目
        }
        number4 += 1;
      });  
    });
    var namesNumber1 = namesTrans[0].indexOf(selectedName1);
    var namesNumber2 = namesTrans[0].indexOf(selectedName2);
    var selctedAdress1 = namesTrans[1][namesNumber1];
    var selctedAdress2 = namesTrans[1][namesNumber2];
    
    // 夜勤当番者が記入されているセルの色が赤色でなければ、メールを送信する。
    // 夜勤当番者のセル情報を取得する。
    var getCell = schedule.getRange(firstRowName, dayNumber+2, 1, 1);
    var getCellBackground = getCell.getBackground();
    if(getCellBackground !== '#ff0000'){

      // 夜勤当番者に電話が繋がるかの確認依頼メールを送信する。
      // メールの送信先
      var to = selctedAdress1,selctedAdress2;
      // メールのタイトル
      var subject = '24時間当番の電話確認をお願いします。';
      // メールの本文
      var body = '\
${nightShiftDuty1}さん,${nightShiftDuty2}さん\n\n\
お仕事お疲れ様です。\n\
本日、24時間当番お願いします。\n\
まだ電話確認が取れていません。\n\
確認をお願いします。'
.replace('${nightShiftDuty1}', nightShiftDuty1)
.replace('${nightShiftDuty2}', nightShiftDuty2);
      
      var options = { name: 'ISOWA_support',
                     bcc: 'k.kamikura@isowa.co.jp'
                    };
    }
  }
  // ログ確認用
  Logger.log(dayOfWeek);          // 本日の曜日番号
  Logger.log('---------------');
  Logger.log(nightShiftDutys2);   // 金曜日の当番（２人分）
  Logger.log('---------------');
  Logger.log(namesTrans[0]);      // メールアドレス一覧
  Logger.log('---------------');  
  Logger.log(selectedName1);      // 当番の名前（1人目）
  Logger.log('---------------');
  Logger.log(namesNumber1);       // 当番の番号（1人目）
  Logger.log('---------------');
  Logger.log(selctedAdress1);     // 当番のアドレス（１人目）
  Logger.log('---------------');
  Logger.log(selectedName2);      // 当番の名前（2人目）
  Logger.log('---------------');
  Logger.log(namesNumber2);       // 当番の番号（2人目）
  Logger.log('---------------');
  Logger.log(selctedAdress2);     // 当番のアドレス（2人目）
  Logger.log('---------------');
  Logger.log(nightShiftDutys1);    // 当日の当番（2人分）
  Logger.log('---------------');
  Logger.log(dayNumber);          // 日付の番号
  Logger.log('---------------');
  Logger.log(getCellBackground);  // 当日のセル背景色 
  Logger.log('---------------');
}