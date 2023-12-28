function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // カスタムメニューを作成
    ui.createMenu('自動グラフ作成')
      .addItem('データ取得とグラフ作成', 'getDataAndCreateChart')
      .addToUi();
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
      sheet.appendRow(['date','26卒','25卒','目標']);
    }
  }
  
  function difineCalcurateSheet(){
    // 計算用シートを定義
    var calsheet_name = '目標達成グラフ';
    checkAndCreateSheet(calsheet_name);
    var calsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(calsheet_name);
    return calsheet
  }
  
  function getDataAndCreateChart() {
    // 必要なシートを取得
    var sheet_26 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bチーム_26卒LG招待');
    var sheet_25 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bチーム_25卒初回面談予約');
    var calsheet = difineCalcurateSheet();
  
    // 現在の値を取得
    var data_26 = sheet_26.getRange('H3').getValue();
    var data_25 = sheet_25.getRange('G3').getValue();
  
    // 現在の日時を取得
    var currentDate = new Date();
    var formattedDate = Utilities.formatDate(currentDate, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'M月d日');
  
    // 達成率計算
    var achievement_num_26 = sheet_26.getRange('G3').getValue();
    var achievement_num_25 = sheet_25.getRange('F3').getValue();
    var achievement_rate_26 = data_26 / achievement_num_26 * 100;
    var achievement_rate_25 = data_25 / achievement_num_25 * 100;
  
    // グラフを作成するためのデータをシートに追加
    calsheet.appendRow([formattedDate, achievement_rate_26, achievement_rate_25, 100]);
    
    // グラフを作成
    createChart();
  }
  
  function createChart() {
  
    // シートを取得
    var calsheet = difineCalcurateSheet();
  
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
      .setOption('title', '👑'+formattedMonth+'目標達成グラフ👑')
      .setOption('titleTextStyle', { fontSize: 30, color: '#e60033', bold: true, alignment: 'center'})  // タイトルのスタイルを設定
      .setOption('legend', {position: 'bottom'}) // 凡例の設定
      .addRange(calsheet.getRange('A1:D29')) // 日時とデータの列を指定
      .setOption('series', { 0: { labelInLegend: '26卒' }, 1: { labelInLegend: '25卒' }, 2: { labelInLegend: '目標' } })
      .setOption('series', { 0: { pointSize: 7 }, 1: { pointSize: 7 }, 2: { pointSize: 0 } }) // プロットのサイズを設定
      .setPosition(5, 5, 0, 0) // グラフの位置を指定
      .build();
  
    // シートにグラフを挿入
    calsheet.insertChart(chart);
  }