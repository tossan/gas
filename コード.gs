function _test() {
  // monthは1月=0...12月=11
  var from = new Date(2018, 3, 2);
  var to = new Date(2018, 3, 9);
  var ret = calcWorkRate('吉澤瑞恵', from, to, 2, 5, 8);
  Logger.log(ret);
}

function _test2() {
  var seiban = "TD16A001";
  var seibans = "TD16A001:1\nTD16A002:2";
  var hour = 3;
  seibans = getActSeiban(seiban, seibans, hour);
  Logger.log(seibans);
}

var COL_SEIBAN = 2;
var COL_DATE = 3;
var COL_NAME = 5;
var COL_HOUR = 7;
var ROW_DATA_START = 2;

var WORKDAYS_OF_A_WEEK_DEFAULT = 5;
var HOURS_OF_A_DAY_DEFAULT = 8;

var CALC_TYPE_ACTRATE = 1;
var CALC_TYPE_INPUTRATE = 2;
var CALC_TYPE_ACTSEIBAN = 3;
var CALC_TYPE_SEIBAN_TOTAL = 4;

// 研究/開発/教育製番及び汎用製番
var RANDD_GENERAL_SEIBAN = ['TD16A002', 'TD16A006', 'TD16G002', 'TD16G005', 'TD17A002', 'TD17A005', 'TD17G002', 'TD17G005', 'TD18A002', 'TD18A005', 'TD18G002', 'TD18G005', 'TD19A002', 'TD19A005', 'TD19G002', 'TD19G005', 'TD20A002', 'TD20A005', 'TD20G002', 'TD20G005'];
  
// from <= 日付 < to とする 
function calcWorkRate(name, from, to, calctype, workdays, workhours) {
  
  // オプション引数の初期値の設定
  if (workdays == null) {
    workdays = WORKDAYS_OF_A_WEEK_DEFAULT;
  }
  if (workhours == null) {
    workhours = HOURS_OF_A_DAY_DEFAULT;
  }
  
  // 役務算出シートのデータが入力されている範囲の全データを取得
  var sheet = SpreadsheetApp.getActive().getSheetByName('役務算出');
  var data = sheet.getDataRange().getValues();
  
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