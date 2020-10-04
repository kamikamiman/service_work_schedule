/******************************
月の日数           endColLの値     
 1月 31日          >>> 17
 2月 28日 or 29日  >>> 14 or 15
 3月 31日          >>> 17
 4月 30日          >>> 16
 5月 31日          >>> 17
 6月 30日          >>> 16
 7月 31日          >>> 17
 8月 31日          >>> 17
 9月 30日          >>> 16
10月 31日          >>> 17
11月 30日          >>> 16
12月 31日          >>> 17
******************************/

function workSchedule_Month11() {
  
  // スプレットシートの情報を取得する。
  const ssGet = SpreadsheetApp.openById('1Wf2nEZEh4YfiKSfn2iNfBIs8hcxsFdYBBI8o6vwJYxY'); // 【サービス作業予定表】
  const ssSet = SpreadsheetApp.openById('1xuz9S0jg2xxHHYDwbi0LLc2R_Cg5dO-6OgyazJumpbk'); // 【eサービス作業予定】
  
  // 設定項目
  const period = 69; // 第〇〇期
  
// 読込先スプレットシートを取得する。
const schedule = ssGet.getSheetByName('${period}期11月'.replace('${period}', period));
// 書込先のスプレットシートを取得する。  
const workSchedule = ssSet.getSheetByName('${period}期11月'.replace('${period}', period));
// eサービス、協力会社のシートを取得する。
const ssMembers  = ssGet.getSheetByName('eサービスメンバー + 協力会社');

const row1 = ssMembers.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();   // eサービスメンバー 最終行
const row2 = ssMembers.getRange(1, 2).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();   // 協力会社 最終行
const isowabito  = ssMembers.getRange( 1, 1, row1, 1).getValues().flat();                        // eサービスメンバー
const tasukebito = ssMembers.getRange( 1, 2, row2, 1).getValues().flat();                        // 協力会社
const members = schedule.getRange( 1, 1, 180, 1 ).getValues().flat();
const enColL = 16;  // 終了列(下旬)  * 月により変動
let row;
const num = 24;  // 人数が増減する場合はこの値を減算
let wRowU = 2;
let wRowL = wRowU + num;  // 人数が増減により可変


// ログ確認
Logger.log(row1);
Logger.log(row2);
Logger.log(isowabito);
Logger.log(tasukebito);
Logger.log(members);


const eMembers = [ isowabito, tasukebito ]; // eサービスメンバー + 協力会社

eMembers.forEach( eMember => {
   eMember.forEach( el => {
      members.forEach( member => {
         if ( el === member ) {
             InputValue(el);  
             wRowU++;
             wRowL = wRowU + num; // 人数が増減により可変
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
//  const stColU =  1;  // 開始列(上旬)  * 基本固定値
//  const stColL = 17;  // 開始列(下旬)  * 基本固定値
//  const enColU = 16;  // 終了列(上旬)  * 基本固定値
  const val1 = readSh.getRange(row, 1, 1, 16).getValues();                    // 予定表(上旬)
  const val2 = readSh.getRange(row,17, 1, enColL).getValues();                // 予定表(下旬)
  const val3 = readSh.getRange(row, 1, 1, 1).getValues();                     // メンバー名(下旬)
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
//  ColorCoding();  // 行の色を塗り潰す 　初回のみ実行する
//  HolidayColor(); // 休日を色分けする　 初回のみ実行する
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
//  const cell = workSchedule.getRange("A2").getValues();
//  Logger.log(cell);
  
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
  const firstRowU =  1;     // 開始行(上旬)  * 基本固定値
  const firstRowL = 25;     // 開始行(下旬)  * 人数増減により変動 * 1 ここを変更します
  const lastRowU  = 23;     // 終了行(上旬)  * 人数増減により変動 * 1 ここを変更します
  const lastRowL  = 47;     // 終了行(下旬)  * 人数増減により変動 * 2 ここを変更します
    
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
      
      
      // 上旬の日付を検索する。
      var dayU = new Date(workSchedule.getRange('${day}${firstRowU}'
                                                .replace('${day}', day)
                                                .replace('${firstRowU}', firstRowU)
                                               ).getValues()).getDay();
      
      
      // 下旬の予定を検索する。
      var dayL = new Date(workSchedule.getRange('${day}${firstRowL}'
                                                .replace('${day}', day)
                                                .replace('${firstRowL}', firstRowL)
                                               ).getValues()).getDay();
      
      
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







