// ここは適宜変更してください
function setting(team){
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
  var teamlist = ["A", "B", "C"]; // チームの増減で適宜変更
  return teamlist 
}

function checkAndCreateSheet(sheet_name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = sheet_name;

  // シートが存在するか確認
  var existingSheet = spreadsheet.getSheetByName(sheetName);

  if (!existingSheet) {
    // シートが存在しない場合は新しいシートを作成
    spreadsheet.insertSheet(sheetName);
    Logger.log(sheetName + 'The calcurate cheet created , welcome!');
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    sheet.appendRow(['date','26卒','25卒','目標','日別目標']);
  }
}

function difineCalcurateSheet(team){
  // 計算用シートを定義
  var calsheet_name = '目標達成グラフ_'+team+'チーム';
  checkAndCreateSheet(calsheet_name);
  var calsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(calsheet_name);
  return calsheet
}

function getData(team, sheetname_26, sheetname_25, achievecell_26, achievecell_25, goalcell_26, goalcell_25) {
  // 必要なシートを取得
  var sheet_26 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname_26);
  var sheet_25 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetname_25);
  var calsheet = difineCalcurateSheet(team);

  // 現在の値を取得
  var data_26 = sheet_26.getRange(achievecell_26).getValue();
  var data_25 = sheet_25.getRange(achievecell_25).getValue();

  // 現在の日時を取得
  var currentDate = new Date();
  var formattedDate = Utilities.formatDate(currentDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'M月d日');

  // 達成率算出
  var achievement_num_26 = sheet_26.getRange(goalcell_26).getValue();
  var achievement_num_25 = sheet_25.getRange(goalcell_25).getValue();
  var achievement_rate_26 = data_26 / achievement_num_26 * 100;
  var achievement_rate_25 = data_25 / achievement_num_25 * 100;

  // 今日は月の何日目か調べて日別目標値を算出
  var today = new Date();
  var daysPassed = today.getDate();
  var month = today.getMonth() + 1; //today.getMonth()の出力値がmonth_listのため+1
  if(month == 1 || month == 3 || month == 5 || month == 7 || month == 8 || month == 10 || month == 12){
    var day_num = 31
  }else if(month == 2){
    var day_num = 28 //閏年だけ気を付けて。まあ達成率は1日ずれてもそこまで問題はないと判断し閏年の処理は未実装
  }else{
    var day_num = 30
  };
  var achievement_rate_day = daysPassed / day_num * 100;

  // グラフを作成するためのデータをシートに追加
  calsheet.appendRow([formattedDate, achievement_rate_26, achievement_rate_25, 100, achievement_rate_day]);
}

function createChart(team) {

  // シートを取得
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
    .setOption('title', '👑'+formattedMonth+'目標達成グラフ_'+team+'チーム👑')
    .setOption('titleTextStyle', { fontSize: 30, color: '#e60033', bold: true, alignment: 'center'})  // タイトルのスタイルを設定
    .setOption('legend', {position: 'bottom'}) // 凡例の設定
    .addRange(calsheet.getRange('A:E')) // 日時とデータの列を指定
    .setOption('series', { 0: { labelInLegend: '26卒' }, 1: { labelInLegend: '25卒' }, 2: { labelInLegend: '目標' }, 3: { labelInLegend: '日別目標' } })
    .setOption('series', { 0: { pointSize: 7 }, 1: { pointSize: 7 }, 2: { pointSize: 0 }, 3: { pointSize: 0} }) // プロットのサイズを設定
    .setPosition(5, 5, 0, 0) // グラフの位置を指定
    .build();

  // シートにグラフを挿入
  calsheet.insertChart(chart);
}

function main() {
  var teamlist = makeTeamlist();
    for (var i = 0; i < teamlist.length; i++) {
      var {team, sheetname_26, sheetname_25, achievecell_26, achievecell_25, goalcell_26, goalcell_25} = setting(teamlist[i]);
      getData(team, sheetname_26, sheetname_25, achievecell_26, achievecell_25, goalcell_26, goalcell_25);
      createChart(team);
  }
}