// ここは適宜変更してください
function setting(team){
  /*
  SETTING 設定用関数
    それぞれのチームのシートと、そこで見るべきセルの位置を指定する
  
  Parameters
    team : string
      チーム名
  
  Returns
    {team, sheetname_26, sheetname_25, achievecell_26, achievecell_25, goalcell_26, goalcell_25} : list
      以下が格納されたリスト
      sheetname_26 : string
        26卒シート名
      sheetname_25 : string
        25卒シート名
      achievecell_26 : string
        達成数の書かれたセル(26卒)
      achievecell_25 : string
        達成数の書かれたセル(25卒)
      goalcell_26 : string
        目標数の書かれたセル(26卒)
      goalcell_25 : string
        目標数の書かれたセル(25卒)
  */
  var team = team; // チーム名
  var sheetname_26 = team+" 26卒"; // 25卒のシートをbynameで指定
  var sheetname_25 = team+" 25卒"; // 26卒のシートをbynameで指定
  var achievecell_26 = "D5"; // 25卒の実績が格納されるセルを指定
  var achievecell_25 = "D5"; // 26卒の実績が格納されるセルを指定
  var goalcell_26 = "E5"; // 25卒の目標が格納されるセルを指定
  var goalcell_25 = "E5"; // 26卒の目標が格納されるセルを指定
  return {team, sheetname_26, sheetname_25, achievecell_26, achievecell_25, goalcell_26, goalcell_25}
}

function makeTeamlist(){
  /*
  MAKETEAMLIST 設定用関数
    チームの数と名前を指定する
  
  Parameters 
    なし

  Returns
    teamlist : list
      チーム名(string)が格納されたリスト
  */
  var teamlist = ["A", "B", "C"]; // チームの増減で適宜変更
  return teamlist 
}

function vacation(){
  /*
  VACATION 設定用関数
    休暇日数を指定する。月のはじめが休暇だった時は正の整数、他は0を記入（例：1/4まで休暇→4と記入）
  
  Parameters
    なし

  Returns
    vacation : integer
      休暇日数
  */
  var vacation = 4; 
  return vacation
}
// ここまで

function checkAndCreateSheet(sheet_name) {
  /*
  CHECKANDCREATESHEET 計算・グラフ出力用シートの存在を担保する
    シートがない場合は作成し、シートがある場合はスキップする

  Parameters
    sheet_name : string
      計算・グラフ出力用シートの名前
  
  Returns
    なし
  */
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = sheet_name;

  var existingSheet = spreadsheet.getSheetByName(sheetName);
  if (!existingSheet) {
    spreadsheet.insertSheet(sheetName);
    Logger.log(sheetName + 'The calcurate cheet created , welcome!');
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    sheet.appendRow(['date','26卒','25卒','目標','日別目標']);
  }
}

function difineCalcurateSheet(team){
  /*
  DIFINECALCURATESHEET 計算・グラフ出力用シートの名前を定義する

  Parameters
    team : string
      チーム名
  
  Returns
    calsheet : object
      計算・グラフ出力用シートオブジェクト
  */
  var calsheet_name = '目標達成グラフ_'+team+'チーム';
  checkAndCreateSheet(calsheet_name);
  var calsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(calsheet_name);
  return calsheet
}

function getData(team, sheetname_26, sheetname_25, achievecell_26, achievecell_25, goalcell_26, goalcell_25) {
  /*
  GETDATA データの取得と計算
    データを取ってきて、達成率などの計算をしてセルに出力する

  Parameters
    team : string
      チーム名
    sheetname_26 : string
      26卒シート名
    sheetname_25 : string
      25卒シート名
    achievecell_26 : string
      達成数の書かれたセル(26卒)
    achievecell_25 : string
      達成数の書かれたセル(25卒)
    goalcell_26 : string
      目標数の書かれたセル(26卒)
    goalcell_25 : string
      目標数の書かれたセル(25卒)
  
  Returns
    なし
  */
  var sheet_26 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname_26);
  var sheet_25 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname_25);
  var calsheet = difineCalcurateSheet(team);

  var data_26 = sheet_26.getRange(achievecell_26).getValue();
  var data_25 = sheet_25.getRange(achievecell_25).getValue();

  // 日付を出力
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'M月d日');

  // 達成率計算
  var achievement_num_26 = sheet_26.getRange(goalcell_26).getValue();
  var achievement_num_25 = sheet_25.getRange(goalcell_25).getValue();
  var achievement_rate_26 = data_26 / achievement_num_26 * 100;
  var achievement_rate_25 = data_25 / achievement_num_25 * 100;

  // 日程消化率計算
  var today = new Date();
  var vacation_int = vacation();
  var daysPassed = today.getDate() - vacation_int ;
  var month = today.getMonth() + 1; //today.getMonth()の出力値がmonth_listのため+1
  if(month == 1 || month == 3 || month == 5 || month == 7 || month == 8 || month == 10 || month == 12){
    var day_num = 31
  }else if(month == 2){
    var day_num = 28 //note : 閏年だけ気を付けて。まあ達成率は1日ずれてもそこまで問題はないと判断し閏年の処理は未実装
  }else{
    var day_num = 30
  };
  var achievement_rate_day = daysPassed / day_num * 100;

  calsheet.appendRow([formattedDate, achievement_rate_26, achievement_rate_25, 100, achievement_rate_day]);
}

function createChart(team) {
  /*
  CREATECHART グラフ出力
    既存のグラフを削除したのち、新しいグラフを作成する

  Parameters
    team : string
      チーム名
  
  Returns
    なし
  */
  var calsheet = difineCalcurateSheet(team);

  // 既存のグラフがあれば削除
  var existingCharts = calsheet.getCharts();
  for (var i = 0; i < existingCharts.length; i++) {
    calsheet.removeChart(existingCharts[i]);
  }

  //月取得
  var currentDate = new Date();
  var formattedMonth = Utilities.formatDate(currentDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'M月');

  // チャートを作成
  var chart = calsheet.newChart()
    .asLineChart()
    .setOption('title', '👑'+formattedMonth+'目標達成グラフ '+team+'チーム👑')
    .setOption('titleTextStyle', { fontSize: 30, color: '#e60033', bold: true, alignment: 'center'})  // タイトルのスタイルを設定
    .setOption('legend', {position: 'bottom'}) // 凡例の設定
    .addRange(calsheet.getRange('A:E')) // 日時とデータの列を指定
    .setOption('series', { 0: { labelInLegend: '26卒' }, 1: { labelInLegend: '25卒' }, 2: { labelInLegend: '目標' }, 3: { labelInLegend: '日別目標' } })
    .setOption('series', { 0: { pointSize: 7 }, 1: { pointSize: 7 }, 2: { pointSize: 0 }, 3: { pointSize: 0} }) // プロットのサイズを設定
    .setPosition(5, 5, 0, 0) // グラフの位置を指定
    .build();

  calsheet.insertChart(chart);
}

function main() {
  /*
  MAIN メイン関数
    実際に動いているのはこれ

  Parameters
    なし
  
  Returns
    なし
  */  
  var teamlist = makeTeamlist();
    for (var i = 0; i < teamlist.length; i++) {
      var {team, sheetname_26, sheetname_25, achievecell_26, achievecell_25, goalcell_26, goalcell_25} = setting(teamlist[i]);
      getData(team, sheetname_26, sheetname_25, achievecell_26, achievecell_25, goalcell_26, goalcell_25);
      createChart(team);
  }
}