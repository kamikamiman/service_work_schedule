/*******************************************************************************************************************
*                        指定した列の情報と特定の文字列が一致する行番号を返す関数                                           *
*******************************************************************************************************************/

function input_value(member){

  const key = member;                                      // 文字列を指定(メンバー情報) 
  const col = "A";                                         // 指定した文字列を検索する列を指定
  const readSh = SpreadsheetApp.openById("idを入れる").getSheetByName("シート名を入れる");       // 読込先のスプレットシート
  row = get_row(key, col, readSh);                         //【関数】get_row()を実行、rowに該当する行番号を渡す

}


  //【関数】get_row()  [ 指定したsh内のkeyと一致する行番号を取得する関数 ]
　function get_row(key, col, sh){
   
   const array = get_array(sh, col);                       // get_array()を実行、arrayに配列を渡す
   const row = array.indexOf(key) + 1;                     // 配列の中から該当するメンバーと一致する行番号をrowに渡す
   return row;
   
 }


  // 【関数】get_array()  [ 指定したシート(sh)の情報を取得して配列に格納する ]
　function get_array(sh, col) {
   
   const last_row = sh.getLastRow();                       // シートの最終行目を取得する
   const range = sh.getRange(col + "1:" + col + last_row)  // シートの選択範囲を指定
   const values = range.getValues();                       // シートの選択範囲の情報を取得
   const array = [];                                       // 配列arrayを定義
   for(let i = 0; i < values.length; i++){                 // シートの選択数分を順番に取得し、配列arrayに格納
     array.push(values[i][0]);
   }
   return array;
   
 }


/*******************************************************************************************************************
*                        月末日を取得する関数                                                                         *
*******************************************************************************************************************/

function last_date() {
  const today = new Date();      // 今日の日付を取得する関数を実行
  const _lastDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  const  lastDate = _lastDate.getDate();
}









