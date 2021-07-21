// テストなど
function _reCalc() {
  const sheet = SpreadsheetApp.getActiveSheet();
  var s = sheet.getRange(11, 3).getValue();
  sheet.getRange(11, 3).setValue("dummy");
  sheet.getRange(11, 3).setValue(s);
}

function _test() {
  let ekimuSheet = SpreadsheetApp.getActive().getSheetByName('役務算出');
  let ekimuData = ekimuSheet.getDataRange().getValues();
  // monthは1月=0...12月=11
  var from = new Date(2021, 6, 1);
  var to = new Date(2021, 6, 5);
  var ret = calcWorkRate2('岩田俊朗', from, to, CALC_TYPE_INPUTRATE, ekimuData, 2, 8);
  Logger.log(ret);
}

function _test2() {
  var seiban = "TD16A001";
  var seibans = "TD16A001:1\nTD16A002:2";
  var hour = 3;
  seibans = getActSeiban(seiban, seibans, hour);
  Logger.log(seibans);
}

// 業務ロジック
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('ITDC');
  menu.addItem('役務算出集計', 'execCalc');
  menu.addToUi();
}

// 研究/開発/教育製番及び汎用製番
var RANDD_GENERAL_SEIBAN = ['TD16A002', 'TD16A006', 'TD16G002', 'TD16G005', 'TD17A002', 'TD17A005', 'TD17G002', 'TD17G005', 'TD18A002', 'TD18A005', 'TD18G002', 'TD18G005', 'TD19A002', 'TD19A005', 'TD19G002', 'TD19G005', 'TD20A002', 'TD20A005', 'TD20G002', 'TD20G005', 'TD21A002', 'TD21A005', 'TD21G002', 'TD21G005'];

var WORKDAYS_OF_A_WEEK_DEFAULT = 5;
var HOURS_OF_A_DAY_DEFAULT = 8;

var CALC_TYPE_ACTRATE = 1;
var CALC_TYPE_INPUTRATE = 2;
var CALC_TYPE_ACTSEIBAN = 3;
var CALC_TYPE_SEIBAN_TOTAL = 4;

let SHEET_TYPE_PROJECT = 1;
let SHEET_TYPE_MEMBER = 2;
let SHEET_TYPE_UNDEF = -1;

// 役務算出シート
var COL_SEIBAN = 2;
var COL_DATE = 3;
var COL_NAME = 5;
// var COL_HOUR = 7;  // ～37期下期
var COL_HOUR = 9; // 38期上期
var ROW_DATA_START = 2;

// プロジェクトシート、メンバーシート
let IDX_KEYWORD = 2;
let COL_PROJECT_SEIBAN = 3;
let COL_MEMBER_NAME = 3;
let COL_DATE_START_PROJECT = 5;
let COL_DATE_START_MEMBER = 4;

function execCalc() {
  // ITDC->役務算出集計コマンドが実行されたシートを判定
  let sheet_type = getSheetType(SpreadsheetApp.getActiveSheet().getSheetName())
  if (sheet_type == SHEET_TYPE_UNDEF) {
    Browser.msgBox("プロジェクトシートかメンバーシートを選択してから実行してください。");
    return;
  }
  // アクティブシートとアクティブセルの取得
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let activeCell = activeSheet.getCurrentCell();
  let activeCellValue = activeCell.getValue().toString().trim();
  let activeCellRow = activeCell.getRow();
  let activeCellColumn = activeCell.getColumn();
  let targetRow = 0;
  if (sheet_type == SHEET_TYPE_PROJECT) {
    if (isSeibanString(activeCellValue) && activeCellColumn == COL_PROJECT_SEIBAN) {
      let ret = Browser.msgBox("製番の実績を集計します\\nはい（Yes）：" + activeCellRow + "行目の " + activeCellValue + " だけを集計する\\nいいえ（No）：全ての製番を集計する", Browser.Buttons.YES_NO_CANCEL);
      if (ret == "yes") {
        targetRow = activeCellRow;
      } else if (ret == "cancel") {
        return;
      }
    } else {
      let ret = Browser.msgBox("全ての製番の実績を集計します", Browser.Buttons.OK_CANCEL);
      if (ret == "cancel") {
        return;
      }
    }
  } else if (sheet_type == SHEET_TYPE_MEMBER) {
    if (isMemberNameCell(activeSheet, activeCellRow, activeCellColumn)) {
      let ret = Browser.msgBox("メンバーの実績を集計します\\nはい（Yes）：" + activeCellValue + " だけを集計する\\nいいえ（No）：全てのメンバーを集計する", Browser.Buttons.YES_NO_CANCEL);
      if (ret == "yes") {
        targetRow = activeCellRow;
      } else if (ret == "cancel") {
        return;
      }
    } else {
      let ret = Browser.msgBox("全てのメンバーの実績を集計します", Browser.Buttons.OK_CANCEL);
      if (ret == "cancel") {
        return;
      }
    }
  } else {
    Browser.msgBox("集計対象シートではありません。");
    return;
  }
  // 役務算出シートのデータが入力されている範囲の全データを取得
  let ekimuSheet = SpreadsheetApp.getActive().getSheetByName('役務算出');
  let ekimuData = ekimuSheet.getDataRange().getValues();
  let rowKeyword = getKeywordRowBySheetType(sheet_type, activeSheet);
  for (let cntRow = 1; cntRow < activeSheet.getLastRow(); cntRow++) {
    if (sheet_type == SHEET_TYPE_PROJECT) {
      let seiban = activeSheet.getRange(cntRow, COL_PROJECT_SEIBAN).getValue();
      if (isSeibanString(seiban)) {
        if (targetRow > 0 && targetRow != cntRow) {
          continue;
        }
        for (let cntCol = COL_DATE_START_PROJECT; activeSheet.getLastColumn(); cntCol++) {
          let rowDate = rowKeyword;
          let dateFrom = activeSheet.getRange(rowDate, cntCol).getValue();
          let dateTo = activeSheet.getRange(rowDate, cntCol + 1).getValue();
          if (dateTo.toString().trim().length <= 0) {
            break;
          }
          let value = calcWorkRate2(seiban, dateFrom, dateTo, CALC_TYPE_SEIBAN_TOTAL, ekimuData);
          value = isNumber(value) ? value : 0;
          activeSheet.getRange(cntRow, cntCol).setValue(value);
        }
        SpreadsheetApp.getActiveSpreadsheet().toast(cntRow + "行目 " + seiban + " 集計済", "進捗表示", 1.5);
        if (targetRow > 0) {
          break;
        }
      }
    } else if (sheet_type == SHEET_TYPE_MEMBER) {
      let memberName = activeSheet.getRange(cntRow, COL_MEMBER_NAME).getValue();
      if (isMemberNameCell(activeSheet, cntRow, COL_MEMBER_NAME)) {
        if (targetRow > 0 && targetRow != cntRow) {
          continue;
        }
        for (let cntCol = COL_DATE_START_MEMBER; activeSheet.getLastColumn(); cntCol++) {
          let rowDate = rowKeyword + 2;
          let dateFrom = activeSheet.getRange(rowDate, cntCol).getValue();
          let dateTo = activeSheet.getRange(rowDate, cntCol + 1).getValue();
          if (dateTo.toString().trim().length <= 0) {
            break;
          }
          let rowWorkdays = rowKeyword;
          let workdays = activeSheet.getRange(rowWorkdays, cntCol).getValue();
          // 稼動実績
          let v1 = calcWorkRate2(memberName, dateFrom, dateTo, CALC_TYPE_ACTRATE, ekimuData, workdays);
          v1 = isNumber(v1) ? v1 : 0;
          activeSheet.getRange(cntRow + 3, cntCol).setValue(v1);
          // 製番毎実績
          let v2 = calcWorkRate2(memberName, dateFrom, dateTo, CALC_TYPE_ACTSEIBAN, ekimuData);
          activeSheet.getRange(cntRow + 4, cntCol).setValue(v2);
          // 入力率
          let v3 = calcWorkRate2(memberName, dateFrom, dateTo, CALC_TYPE_INPUTRATE, ekimuData, workdays);
          v3 = isNumber(v3) ? v3 : 0;
          activeSheet.getRange(cntRow + 5, cntCol).setValue(v3);
        }
        SpreadsheetApp.getActiveSpreadsheet().toast(memberName + " 集計済", "進捗表示", 1.5);
        if (targetRow > 0) {
          break;
        }
      }
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("集計完了", "進捗表示", 1.5);
}

function getSheetType(sheetName) {
  let sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  let data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    let keyword = data[i][IDX_KEYWORD];
    if (keyword == '案件名') {
      return SHEET_TYPE_PROJECT;
    } else if (keyword == '実稼働日数') {
      return SHEET_TYPE_MEMBER;
    }
  }
  return SHEET_TYPE_UNDEF;
}

function getKeywordRowBySheetType(sheetType, activeSheet) {
  for (let i = 1; i <= activeSheet.getLastRow(); i++) {
    let keyword = activeSheet.getRange(i, COL_PROJECT_SEIBAN).getValue();
    if (sheetType == SHEET_TYPE_PROJECT) {
      if (keyword == '案件名') {
        return i;
      }
    } else if (sheetType == SHEET_TYPE_MEMBER) {
      if (keyword == '実稼働日数') {
        return i;
      }
    }
  }
  return -1;
}

// 引数として渡された文字列が製番かどうかを判定する
function isSeibanString(s) {
  let reg = /[A-Z]{2}[0-9]{2}[A-Z][0-9]{3}/gi;
  return reg.test(s)
}

// 引数として渡されたシートのセルがメンバー名のセルかどうかを判定する
function isMemberNameCell(sht, r, c) {
  let v1 = sht.getRange(r + 1, c).getValue();
  let v2 = sht.getRange(r + 2, c).getValue();
  let v3 = sht.getRange(r + 3, c).getValue();
  if (v1 == "保守案件" && v2 == "稼動予定" && v3 == "稼動実績") {
    return true;
  } else {
    return false;
  }
}

// 数値かどうかを判定する
function isNumber(v) {
  return ((typeof v === 'number') && (isFinite(v)));
}

// from <= 日付 < to とする 
function calcWorkRate2(name, from, to, calctype, data, workdays, workhours) {
  
  // オプション引数の初期値の設定
  if (workdays == null) {
    workdays = WORKDAYS_OF_A_WEEK_DEFAULT;
  }
  if (workhours == null) {
    workhours = HOURS_OF_A_DAY_DEFAULT;
  }

  // 実稼働日数が0の場合は稼働率の計算ができない（ゼロ割となり無限大に発散する）ので0を返す
  if (workdays == 0) {
    return 0;
  }

  // 役務算出シートのデータが入力されている範囲の全データを取得
//  var sheet = SpreadsheetApp.getActive().getSheetByName('役務算出');
//  var data = sheet.getDataRange().getValues();
  
  // nameのスタッフのfrom toの期間の作業時間の合計を取得
  var sum_hour = 0;
  var seibans = "";
  // dataには対象シートの上の行から順番に配列として格納されているので、順ループの場合はシートの上の行から順番に取得する処理になる
  for (var i = ROW_DATA_START; i < data.length; i++) {
    // from日付より前の日付なら期間外なのでループの最初に戻って次の行へ
    if (data[i][COL_DATE] < from) {
      continue;
    }
    // to日付を含む先の日付まで到達したらこれ以降は全て期間外なのでループを抜け余分な処理をしない（WDRからコピペしたままの役務算出シートは日付で昇順であるため）
    if (to <= data[i][COL_DATE]) {
      break;
    }
    // 製番実績集計を得たいのではなく
    if (calctype != CALC_TYPE_SEIBAN_TOTAL) {
      // 対象者でなければループの最初に戻って次の行へ
      if (data[i][COL_NAME] != name) {
        continue;
      }
    }
    
    // calctypeによって作業時間の合計を得る
    if (calctype == CALC_TYPE_ACTRATE) {
      // 稼働率を得たい場合は製番有りの作業だけを加算
      sum_hour += getActHour(data[i][COL_SEIBAN], data[i][COL_HOUR]);
    } else if (calctype == CALC_TYPE_INPUTRATE) {
      // 入力率を得たい場合は全ての作業を加算
      sum_hour += data[i][COL_HOUR];
    } else if (calctype == CALC_TYPE_ACTSEIBAN) {
      // 実績製番リストを得たい場合は重複無しの製番リストを編集
      seibans = getActSeiban(data[i][COL_SEIBAN], seibans, data[i][COL_HOUR]);
    } else if (calctype == CALC_TYPE_SEIBAN_TOTAL) {
      // 製番毎の実績時間を得たい場合は対象製番の時間を加算（この計算タイプの場合、nameには製番が入っている）
      sum_hour += getTargetSeibanHour(name, data[i][COL_SEIBAN], data[i][COL_HOUR]);
    } else {
      // undefined (Do nothing)
    }
  }
  
  // 長期休暇や祝日等で週の稼働日が5日に満たない場合の稼動率や入力率を100%にするための係数
  var coefficient = workdays / WORKDAYS_OF_A_WEEK_DEFAULT;

  if (calctype == CALC_TYPE_ACTSEIBAN) {
    // 末尾の改行を削除
    seibans = seibans.replace(/\n+$/g,'');
    // 実績製番リストを返す
    return seibans;
  } else if (calctype == CALC_TYPE_SEIBAN_TOTAL) {
    // 対象製番の合計時間を返す
    return sum_hour;
  } else {
    // 百分率を計算して返す（1週間単位固定とする）
    var hoursofaweek = workhours * WORKDAYS_OF_A_WEEK_DEFAULT;
    var rete = sum_hour / hoursofaweek / coefficient;
    return rete;
  }
}

function getActHour(seiban, hour) {
  
  // 製番無し作業は稼動に入れない
  if (seiban != "") {
    // 研究/開発/教育製番及び汎用製番は稼動には入れない
    if (RANDD_GENERAL_SEIBAN.indexOf(seiban) == -1) {
      return hour;
    }
  }
  return 0;
}

function getActSeiban(seiban, seibans, hour) {
  
  if (seiban == "") {
    seiban = "製番なし";
  }
  // 重複無しの製番リストを改行区切りで編集
  if (seiban != "") {
    if (seibans.indexOf(seiban) == -1) {
      seibans += seiban + ":" + hour + "\n";
    } else {
      var ary_seiban = seibans.split("\n");
      for (var i = 0; i < ary_seiban.length; i++) {
        var s = ary_seiban[i];
        if (s.indexOf(seiban) != -1) {
          var IDX_SEIBAN = 0;
          var IDX_HOUR = 1;
          var ary_tmp = s.split(":");
          ary_tmp[IDX_HOUR] = parseFloat(ary_tmp[IDX_HOUR]) + hour;
          ary_seiban[i] = ary_tmp[IDX_SEIBAN] + ":" + ary_tmp[IDX_HOUR];
          break;
        }
      }
      seibans = ary_seiban.join("\n");
    }
  }
  return seibans;
}

function getTargetSeibanHour(target_seiban, seiban, hour) {
  
  if (seiban != "") {
    if (target_seiban == seiban) {
      // 対象製番と一致した製番の時間を返す
      return hour;
    }
  }
  return 0;
}