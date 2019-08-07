var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('参照データ');

var sheet2 = SpreadsheetApp.getActiveSpreadsheet(); 
var my_sheet = sheet2.getActiveSheet();
Logger.log("見ているシートは" + my_sheet.getName());
// ---------------------------------------------------------------------------------------------------


function testSum() {

// 範囲を取得する(getValues複数形)　、値を比べる　一番大きい値を取得する --------------------------------
  var basic_skill = sheet.getRange('A82:A89');
  var basic_skill_array = basic_skill.getValues();  // [[55.0], [65.0], [75.0], [0.0], ....]
  
  // 配列から最大値を取り出せる
  var max_basic_value = Math.max.apply(null,basic_skill_array);
  Logger.log("最大値は" + max_basic_value);



 // 合計算出方法---------------------------------
  // A90:A96までのセルの値を取得し、合計する
   sum = 0;
   var selectRow = 96;
   
   for(var i = 90; i <= selectRow; i++) {
   var cell_value = sheet.getRange(i, 1).getValue(); // i行目, ７列目セルから指定行までの値を取得できる
   sum += cell_value;
  }
  
  Logger.log("A90~A" + selectRow +"の合計は" + sum);
  var total_value = max_basic_value + sum
  

  var a97 = sheet.getRange("A97");  // 合計値を格納するセルを取得
  a97.setValue(total_value);        // 値をセット
  Logger.log("全ての合計は" + total_value);
}

// ------------------------------------------------------------------------------------

function sumStartValue() {
  const FIRST_FIXED_PRICE =40;
  const FIRST_CAREER_RETOUCH = my_sheet.getRange('D15').getValue(); // 0.75

  var age = my_sheet.getRange('C3').getValue();
  
  // 経験年数
  var career = my_sheet.getRange('C4').getValue();
  
  var total_sum_range = my_sheet.getRange("L47");
  
  if (career < 1) {
    total_sum_range.setValue(FIRST_FIXED_PRICE);
    Logger.log("経験年数１年未満なので固定値");
    return;

  } else {  
    // 開発工程のチェックTRUE,FALSEを取得
    var enhance_range =  my_sheet.getRange('C5');
    var enhance = enhance_range.getValue();        // 改修

    var implement_range =  my_sheet.getRange('C6'); // 実装
    var implement = implement_range.getValue();
    
    var detailed_design_range =  my_sheet.getRange('C7'); // 詳細設計
    var detailed_design = detailed_design_range.getValue();
    
    var basic_design_range =  my_sheet.getRange('C8'); // 基本設計
    var basic_design = basic_design_range.getValue();
    
    var require_define_range =  my_sheet.getRange('C9'); // 要件定義
    var require_define = require_define_range.getValue();
    
    var agile_range =  my_sheet.getRange('C10'); // アジャイル
    var agile = agile_range.getValue();
    
    // ------------------------------
    var leader_range = my_sheet.getRange('C11'); // リーダー経験
    var leader = leader_range.getValue();
    
    var game_or_service_range = my_sheet.getRange('C12'); // ゲームorサービス経験
    var game_or_service = game_or_service_range.getValue();
    
    const LEADER_ADD_PRICE = 5;
    const GAME_OR_SERVICE_ADD_PRICE = 10;
    
    
    // 開発工程ごとの経験年数
    var enhance_career = my_sheet.getRange('F5').getValue();
    var implement_career = my_sheet.getRange('F6').getValue();
    var detailed_design_career = my_sheet.getRange('F7').getValue();
    var basic_design_career = my_sheet.getRange('F8').getValue();
    var require_define_career = my_sheet.getRange('F9').getValue();
    var agile_career = my_sheet.getRange('F10').getValue();
    Logger.log("経験年数は" + enhance_career + "年");
    
    
    // 基本単価
    // 経験1年以内は補正する
    var enhance_unit_price = my_sheet.getRange('D5').getValue();
     switch (true) {
       case enhance == false:
        enhance_unit_price = 0
        break
       case enhance && enhance_career <= 1:
        enhance_unit_price *= FIRST_CAREER_RETOUCH
        break
    }   
    Logger.log("改修の基本単価は" + enhance_unit_price);
    
    var implement_unit_price = my_sheet.getRange('D6').getValue();
      switch (true) {
        case implement == false:
         implement_unit_price = 0
         break
        case implement && implement_career <= 1:
         implement_unit_price *= FIRST_CAREER_RETOUCH
         break
      }    
    var detailed_design_unit_price = my_sheet.getRange('D7').getValue();
      switch (true) {
        case detailed_design == false:
         detailed_design_unit_price = 0
         break
        case detailed_design && detailed_design_career <= 1:
         detailed_design_unit_price *= FIRST_CAREER_RETOUCH
         break
      }    
    var basic_design_unit_price = my_sheet.getRange('D8').getValue();
      switch (true) {
        case basic_design == false:
         basic_design_unit_price = 0
         break
        case basic_design && basic_design_career <= 1:
         basic_design_unit_price *= FIRST_CAREER_RETOUCH
         break
      }    
    var require_define_unit_price = my_sheet.getRange('D9').getValue();
      switch (true) {
        case require_define == false:
         require_define_unit_price = 0
         break
        case require_define && require_define_career <= 1:
         require_define_unit_price *= FIRST_CAREER_RETOUCH
         break
      }    
    var agile_unit_price = my_sheet.getRange('D10').getValue();
      switch (true) {
        case agile == false:
         agile_unit_price = 0
         break
        case agile && agile_career <= 1:
         agile_unit_price *= FIRST_CAREER_RETOUCH
         break
      }
    
    var unit_price_array =[enhance_unit_price, implement_unit_price, detailed_design_unit_price, basic_design_unit_price, require_define_unit_price, agile_unit_price];
    var max_unit_price = Math.max.apply(null,unit_price_array);
    Logger.log("単価の最大は" + max_unit_price);

    /* ----------------------------------------------------
    加算・減算項目
    ------------------------------------------------------*/
    // リーダー、ゲーム
    var leader_add_price = 0;
    var game_or_service_add_price = 0;
    if (leader) {leader_add_price = LEADER_ADD_PRICE;}
    if (game_or_service) {game_or_service_add_price = GAME_OR_SERVICE_ADD_PRICE;}  
    Logger.log("リーダー加算は" + leader_add_price);
    Logger.log("ゲームサービス加算は" + game_or_service_add_price);
    
    // 経験年数の減算
    var career_reduce= 0;
    switch (true) {
      case career >= 3:
        career_reduce = 0
        break
      case career >= 2:
        career_reduce = -3
        break
      case career >= 1:
        career_reduce = -8
        break      
    }
    Logger.log("経験年数減算は" + career_reduce　 + "万円");
    
    // 年齢減算
    var age_reduce= 0;
    switch (true) {
      case age >= 50:
        age_reduce = -10
        break
      case age < 24:
        age_reduce = -5
        break    
    }   
    Logger.log("年齢減算は" + age_reduce　 + "万円");
    
    
    // 初期値要素の合計；
    start_element_array = [max_unit_price, leader_add_price, game_or_service_add_price,　career_reduce, age_reduce];
    var total_first_value = start_element_array.reduce(function(prev,nx) {
      return prev + nx;
    });
    
    total_sum_range.setValue(total_first_value);
    Logger.log("初期値合計は" + total_first_value　 + "万円");
    
    // 形態による掛け値から支給額を算出
    var l48 = sheet.getRange("L48");  // 合計値を格納するセルを取得
    var employ_rate = sheet.getRange("K45").getValue();
    Logger.log("掛け率は" + employ_rate　);
    var first_payment_amount = Math.round(total_first_value　* employ_rate);  // 四捨五入
    Logger.log("掛け率は" + first_payment_amount　);
    l48.setValue(first_payment_amount);  // 表示
    
    // -----------------------------------------------------------------------------------
    /* １スロットの１〜5年目の計算式：
    for繰り返しを使って　i=1 -> 1年目のセルに値を入れる　-> １〜５までiを繰り返す...
    
    */
 }
}
