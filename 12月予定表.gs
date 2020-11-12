/* 

コピーする際の変更箇所
① 12行目 実行関数名
② 21行目 読込先スプレットシート
③ 36行目 書込先スプレットシート
④ 39行目 メンバー数


*/

function workSchedule_Month12() {
  
  // スプレットシートの情報を取得する。
  const ssGet = SpreadsheetApp.openById('1Wf2nEZEh4YfiKSfn2iNfBIs8hcxsFdYBBI8o6vwJYxY'); // 【サービス作業予定表】
  const ssSet = SpreadsheetApp.openById('1xuz9S0jg2xxHHYDwbi0LLc2R_Cg5dO-6OgyazJumpbk'); // 【eサービス作業予定】
  
  // 設定項目
  const period = 69; // 第〇〇期
  
  // 読込先スプレットシートを取得する。
  const schedule = ssGet.getSheetByName('${period}期12月'.replace('${period}', period));

  // eサービス、協力会社のシートを取得する。
  const ssMembers  = ssGet.getSheetByName('eサービスメンバー + 協力会社');

  const row1 = ssMembers.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();   // eサービスメンバー 最終行
  const row2 = ssMembers.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();   // 協力会社 最終行
  const isowabito  = ssMembers.getRange( 1, 1, row1, 1).getValues().flat();                        // eサービスメンバー
  const tasukebito = ssMembers.getRange( 1, 2, row2, 1).getValues().flat();                        // 協力会社
  const members = schedule.getRange( 1, 1, 180, 1 ).getValues().flat();
  const lastCol = schedule.getRange(6, 2).getNextDataCell(SpreadsheetApp.Direction.NEXT).getColumn();
  const enColL = lastCol - 15;  // 予定表の終了列(下旬)
  let row;
  
  // 書込先のスプレットシートを取得する。  
  const workSchedule = ssSet.getSheetByName('${period}期12月'.replace('${period}', period));
  
  const memberNum = 20; // 人数が増減する場合はこの値を減算
  const blankRow  = 2;  // 上旬と下旬間の空白行数
  const dateNumU  = 1;  // 日付(上旬)の行数
  const dateNumL  = 1;  // 日付(下旬)の行数
  const numRowU   = dateNumU + memberNum + blankRow; // 予定(上旬)の使用行数
  let wRowU = dateNumU + 1;                          // 日付の行(1) + 最初のメンバー記入行(1) = 2
  let wRowL = wRowU + numRowU;                       // 日付(下旬)の最初のメンバー記入行
  const dateRowL  = dateNumU + memberNum + blankRow + dateNumL; // 日付(下旬)行
  const scheLastRow = dateNumU + (memberNum * 2) + blankRow + dateNumL; // 予定表の最終行


  // ログ確認
  Logger.log(row1);
  Logger.log(row2);
  Logger.log(isowabito);
  Logger.log(tasukebito);
  Logger.log(members);
  
  
  const eMembers = [ isowabito, tasukebito ]; // eサービスメンバー + 協力会社
  
  // 日付(上旬・下旬)を 取得 ・ 書込
  const dateU = schedule.getRange(6, 2, 1, 15).getValues();        // 日付(上旬)取得
  const dateL = schedule.getRange(6, 17, 1, enColL).getValues();   // 日付(下旬)取得
  workSchedule.getRange(1, 2, 1, 15).setValues(dateU);             // 日付(上旬)書込
  workSchedule.getRange(dateRowL, 2, 1, enColL).setValues(dateL);  // 日付(下旬)書込
  
  // セルに メンバー名 ・ 予定 を書込
  eMembers.forEach( eMember => {
    eMember.forEach( el => {
      members.forEach( member => {
        if ( el === member ) {
          InputValue(el);  
          wRowU++;
          wRowL = wRowU + numRowU;
        }
      })     
    })
  });


/******************************************** 
* 指定した列と特定の文字列が一致する行番号を取得    *
* 行番号の情報を取得し指定したシートに書き込む関数  *
********************************************/
function InputValue(member){
    
  // 指定した列と特定の文字列が一致する行番号を取得
  const key = member;             // 文字列を指定(メンバー情報) 
  const col = "A";                // 指定した文字列を検索する列を指定
  const readSh  = schedule;       // 読込先のスプレットシートを指定
  row = GetRow(key, col, readSh); //【関数】get_row()を実行、rowに該当する行番号を渡す
  
  
  // 行番号の情報を取得
  const val1  = readSh.getRange(row, 1, 1, 16).getValues();                   // 予定表(上旬)
  const val2  = readSh.getRange(row,17, 1, enColL).getValues();               // 予定表(下旬)
  const val3  = readSh.getRange(row, 1, 1, 1).getValues();                    // メンバー名(下旬)
  const getNotesU = readSh.getRange( row, 1, 1, 16).getNotes();               // メモ(上旬)
  const getNotesL = readSh.getRange( row, 17, 1, enColL).getNotes();          // メモ(下旬)
  const getFontColorU = readSh.getRange( row, 1, 1, 16).getFontColors();      // フォント色(上旬)
  const getFontColorL = readSh.getRange( row, 17, 1, enColL).getFontColors(); // フォント色(下旬)
  
  // 行番号の情報を指定したシートに書き込む
  const whiteSh = workSchedule;  // 書込先のスプレットシート
  whiteSh.getRange(wRowU, 1, 1, 16).setValues(val1);                 // 予定表(上旬)
  whiteSh.getRange(wRowL, 2, 1, enColL).setValues(val2);             // 予定表(下旬)
  whiteSh.getRange(wRowL, 1, 1, 1).setValues(val3);                  // メンバー名(下旬)
  whiteSh.getRange(wRowU, 1, 1, 16).setNotes(getNotesU);             // メモ(上旬)
  whiteSh.getRange(wRowL, 2, 1, enColL).setNotes(getNotesL);         // メモ(下旬)
  whiteSh.getRange(wRowU, 1, 1, 16).setFontColors(getFontColorU)     // フォント色(上旬)
  whiteSh.getRange(wRowL, 2, 1, enColL).setFontColors(getFontColorL) // フォント色(下旬)
//  ColorCoding();  // 行の色を塗り潰す 　初回のみ実行する(実行時間が大きい為)
//  HolidayColor(); // 休日を色分けする　 初回のみ実行する(実行時間が大きい為)
 }

  //【関数】GetRow()  [ 指定したsh内のkeyと一致する行番号を取得する関数 ]
　function GetRow(key, col, sh){
   const array = GetArray(sh, col);   //【関数】get_array()を実行、arrayに配列を渡す
   const row = array.indexOf(key) + 1; // 配列の中から該当するメンバーと一致する行番号をrowに渡す
   return row;
 }

  //【関数】GetArray()  [ 指定したシート(sh)の情報を取得して配列に格納する ]
　function GetArray(sh, col) {
   const lastRow = sh.getLastRow();   // シートの最終行目を取得する
   const range = sh.getRange(col + "1:" + col + lastRow)  // シートの選択範囲を指定
   const values = range.getValues();   // シートの選択範囲の情報を取得
   const array = [];  // 配列arrayを定義
   
   // シートの選択数、順番に取得した情報を配列arrayに格納
   for(let i = 0; i < values.length; i++){
     array.push(values[i][0]);
   }
   return array;
 }


/******************************************** 
*           各行を色分けする関数               *
********************************************/
function ColorCoding() {
  
  if ( wRowU % 2 === 0) {
    workSchedule.getRange( wRowU, 1, 1, 16 ).setBackgroundColor("#d1f4ec");
  } else {
    workSchedule.getRange( wRowU, 1, 1, 16 ).setBackgroundColor("#ffffb7");
  }
  
  if ( wRowL % 2 === 0) {
    workSchedule.getRange( wRowL, 1, 1, enColL ).setBackgroundColor("#d1f4ec");
  } else {
    workSchedule.getRange( wRowL, 1, 1, enColL ).setBackgroundColor("#ffffb7");
  }
  
}


/******************************************** 
*           休日を色分けする関数               *
********************************************/
function HolidayColor () {
  
  const days = ['B','C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P','Q'] // 列の配列
  const firstRowU =  1;                      // 開始行(上旬)
  const firstRowL = dateRowL;                // 開始行(下旬)
  const lastRowU  = dateNumU + memberNum;    // 終了行(上旬)
  const lastRowL  = scheLastRow;             // 終了行(下旬)
    
    days.forEach(function(day){
      
      // 上旬のセルの色変更範囲を指定する。
      const getRangeU = workSchedule.getRange('${day}${firstRowU}:${day}${lastRowU}'
                                              .replace('${day}', day)
                                              .replace('${firstRowU}', firstRowU)
                                              .replace('${day}', day)
                                              .replace('${lastRowU}', lastRowU)
                                             )
      
      // 上旬の文字フォント色変更範囲を指定する。
      const getFontU = workSchedule.getRange('${day}${firstRowU}'
                                             .replace('${day}', day)
                                             .replace('${firstRowU}', firstRowU)
                                            )
      
      // 上旬の日付を検索する。
      const dayU = new Date(getFontU.getValues()).getDay();
      
      
      // 下旬のセルの色変更範囲を指定する。
      const getRangeL = workSchedule.getRange('${day}${firstRowL}:${day}${lastRowL}'
                                              .replace('${day}', day)
                                              .replace('${firstRowL}', firstRowL)
                                              .replace('${day}', day)
                                              .replace('${lastRowL}', lastRowL)
                                             )
      
      // 下旬の文字フォント色変更範囲を指定する。
      const getFontL = workSchedule.getRange('${day}${firstRowL}'
                                             .replace('${day}', day)
                                             .replace('${firstRowL}', firstRowL)
                                            )
      
      // 下旬の予定を検索する。
      const dayL = new Date(getFontL.getValues()).getDay();
      
      
      // 土日だった場合の処理。
      const allDay = [ dayU, dayL ];
      allDay.forEach( day => {
                     if ( day === 0 ) {
        if ( day === dayU ) {
          getRangeU.setBackground('#ffd2ff'); // 日曜日のセルの色
          getFontU.setFontColor('red');       // 日曜日のフォント色 
        } else if ( day === dayL ) {
          getRangeL.setBackground('#ffd2ff'); // 日曜日のセルの色
          getFontL.setFontColor('red');       // 日曜日のフォント色
        }
      } else if ( day === 6 ) {
        if ( day === dayU ) {
          getRangeU.setBackground('#a1caff'); // 土曜日のセルの色
          getFontU.setFontColor('blue');      // 土曜日のフォント色
        } else if ( day === dayL) {
          getRangeL.setBackground('#a1caff'); // 土曜日のセルの色
          getFontL.setFontColor('blue');      // 土曜日のフォント色                  
        }
      }
    });
  });

  }

}