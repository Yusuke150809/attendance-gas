function doGet(e) {
  var selectedEmpId = e.parameter.empId;
  var student       = e.parameter.student;
  var page          = e.parameter.page; 
  var from          = e.parameter.from || ""; 

  // 生徒ページ
  if (page === 'students') {
    return HtmlService.createTemplateFromFile('view_students')
      .evaluate().setTitle('生徒ページ');
  }

  // 従業員ページ
  if (page === 'employees') {
    return HtmlService.createTemplateFromFile('view_employees')
      .evaluate().setTitle('従業員ページ');
  }

  if (page === 'admin') {
    var tmpl = HtmlService.createTemplateFromFile('view_admin_home');
    tmpl.from = from; 
    return tmpl.evaluate().setTitle('塾長ページ');
  }

  // 給与計算ページ
  if (page === 'admin_salary') {
    return HtmlService.createTemplateFromFile('view_admin_salary')
      .evaluate().setTitle('給与計算ページ');
  }

  // 授業分析ページ
  if (page === 'admin_analysis') {
    return HtmlService.createTemplateFromFile('view_admin_analysis')
      .evaluate().setTitle('授業分析ページ');
  }

  // 勤務状況ページ
  if (page === 'admin_attendance') {
    return HtmlService.createTemplateFromFile('view_admin_attendance')
      .evaluate().setTitle('勤務状況ページ');
  }

  // QRコードページ 
  if (page === 'qr') {
    return HtmlService.createTemplateFromFile('view_qr')
      .evaluate().setTitle('QR打刻ページ');
  }

  // フィードバック(従業員)
  if (page === 'feedback_emp') {
    return HtmlService.createTemplateFromFile('view_feedback_emp')
      .evaluate().setTitle('フィードバック（従業員）');
  }

  // 生徒詳細（Feedbackページ）
  if (student != undefined) {
    PropertiesService.getUserProperties().setProperty('selectedStudent', student.toString());
    if (selectedEmpId != undefined) {
      PropertiesService.getUserProperties().setProperty('selectedEmpId', selectedEmpId.toString());
    }
    return HtmlService.createTemplateFromFile("view_feedback")
      .evaluate().setTitle("Feedback: " + student.toString());
  }

  // 年別集計ページ
  if (page === 'yearly') {
    return HtmlService.createTemplateFromFile('yearly')
      .evaluate().setTitle('年別集計ページ');
  }

  // 従業員IDが指定されていない場合はホーム画面
  if (selectedEmpId == undefined) {
    return HtmlService.createTemplateFromFile("view_home")
      .evaluate().setTitle("Home");
  }

  // 従業員詳細ページ
  PropertiesService.getUserProperties().setProperty('selectedEmpId', selectedEmpId.toString());
  return HtmlService.createTemplateFromFile("view_detail")
    .evaluate().setTitle("Detail: " + selectedEmpId.toString());
}


/**
 * このアプリのURLを返す
 */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getSelectedEmpId() {
  return PropertiesService.getUserProperties().getProperty('selectedEmpId') || "";
}

function getSelectedStudent() {
  return PropertiesService.getUserProperties().getProperty('selectedStudent') || "";
}
function setSelectedEmpId(empId) {
  PropertiesService.getUserProperties().setProperty('selectedEmpId', empId);
}

function getEmployees() {  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // 1行目に「従業員番号」「名前」を持つシートを探す
  let sh = null, colId = 0, colName = 1;
  for (const s of sheets) {
    const lastCol = s.getLastColumn();
    if (lastCol < 2) continue;
    const headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
    const iId   = headers.indexOf('従業員番号');
    const iName = headers.indexOf('名前');
    if (iId !== -1 && iName !== -1) {
      sh = s; colId = iId; colName = iName;
      break;
    }
  }

  if (!sh) return []; // 見つからない場合は空

  const last = sh.getLastRow();
  if (last < 2) return []; // データなし

  // 2行目以降（ヘッダー除外）を取得
  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

  const list = [];
  for (let i = 0; i < values.length; i++) {
    const id   = String(values[i][colId]   || '').trim();
    const name = String(values[i][colName] || '').trim();
    if (!id) continue; // 空行スキップ
    list.push({ id: id, name: name });
  }
  return list;
}


/**
 * 従業員情報の取得
 * ※ デバッグするときにはselectedEmpIdを存在するIDで書き換えてください
 */
function getEmployeeName() {                                
  const selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!selectedEmpId) return "";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let sh = null, colId = 0, colName = 1;
  for (const s of sheets) {
    const lastCol = s.getLastColumn();
    if (lastCol < 2) continue;
    const headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
    const iId   = headers.indexOf('従業員番号');
    const iName = headers.indexOf('名前');
    if (iId !== -1 && iName !== -1) {
      sh = s; colId = iId; colName = iName;
      break;
    }
  }
  if (!sh) return "";

  const last = sh.getLastRow();
  if (last < 2) return "";

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const target = String(selectedEmpId).trim();

  for (let i = 0; i < values.length; i++) {
    const id = String(values[i][colId] || '').trim();
    if (id === target) {
      return String(values[i][colName] || '').trim();
    }
  }
  return "";
}



// 日時用と労働時間用に分ける
function formatDateTime(value) {
  var tz = "Asia/Tokyo";
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, "yyyy-MM-dd HH:mm");
  }
  return "";
}

function formatWorkingTime(value) {
  var tz = "Asia/Tokyo";
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, "HH:mm");
  }
  return "";
}


/**
 * 勤怠情報の取得
 * 今月における今日までの勤怠情報が取得される
 */
function getTimeClocks() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]; // 打刻履歴
  var last_row = sh.getLastRow();
  if (last_row < 2) return [];

 
  var range = sh.getRange(2, 1, last_row-1, 7);
  var rows = range.getNumRows();
  var empTimeClocks = [];

  for (var i = 1; i <= rows; i++) {
    var empId    = range.getCell(i, 1).getValue(); // A列: 従業員ID
    var type     = range.getCell(i, 2).getValue(); // B列: 種別
    var datetime = range.getCell(i, 3).getValue(); // C列: 日時
    var subject  = range.getCell(i, 4).getValue(); // D列: 科目
    var wt       = range.getCell(i, 5).getValue(); // E列: 労働時間
    var student  = range.getCell(i, 6).getValue(); // F列: 生徒名
    var fb       = range.getCell(i, 7).getValue(); // G列: フィードバック 

    if (empId === "") break;

    if (String(empId) == String(selectedEmpId)) {
      empTimeClocks.push({
        'date': formatDateTime(datetime),
        'type': type,
        'subject': subject,
        'workingtime': formatWorkingTime(wt),
        'student': student || "",
        'feedback': fb || "" ,
        'row': i + 1 
      });
    }
  }

  // 日付で昇順ソート
  empTimeClocks.sort(function(a, b) {
    return new Date(a.date) - new Date(b.date);
  });

  return empTimeClocks;
}




/**
 * 勤怠情報登録
 */
function saveWorkRecord(form) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');

  var targetDate = form.target_date;
  var targetTime = form.target_time;
  var subject    = form.subject || "";
  var student    = form.student || "";
  var feedback   = form.feedback || "";

  var targetType = '';
  switch (form.target_type) {
    case 'clock_in':    targetType = '出勤'; break;
    case 'break_begin': targetType = '休憩開始'; break;
    case 'break_end':   targetType = '休憩終了'; break;
    case 'clock_out':   targetType = '退勤'; break;
  }

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3];
  var r = sh.getLastRow() + 1;

  sh.getRange(r, 1).setValue(selectedEmpId);
  sh.getRange(r, 2).setValue(targetType);

  var dateObj = new Date(targetDate + 'T' + targetTime + ':00+09:00');
  sh.getRange(r, 3).setValue(dateObj).setNumberFormat("yyyy-MM-dd HH:mm");

  sh.getRange(r, 4).setValue(subject);

  if (targetType === '退勤') { 
    recordTotalWorkingHours(sh, r); 
  }

  sh.getRange(r, 6).setValue(student);

  if (targetType === '退勤' && feedback) {
    sh.getRange(r, 7).setValue(feedback); 
  }

  return targetType + "を記録しました";
}


// 総労働時間を計算
function recordTotalWorkingHours(sh, rowOut) {
  const [empIdOut, typeOut, outStr] = sh.getRange(rowOut, 1, 1, 3).getValues()[0];
  if (typeOut !== '退勤') return;

  const outAt = new Date(outStr);

  // 対応する出勤を探す
  let r = rowOut - 1, inAt;
  for (; r >= 2; r--) {
    const [e, t, s] = sh.getRange(r, 1, 1, 3).getValues()[0];
    if (e == empIdOut && t === '出勤') {
      inAt = new Date(s); 
      break; 
    }
    if (e === "") break;
  }

  if (!inAt) return sh.getRange(rowOut, 5).setValue('');

  // 出勤～退勤の間の休憩を集計
  const between = sh.getRange(r, 1, rowOut - r + 1, 3).getValues();
  let breakMs = 0, last = null;

  for (let i = 1; i < between.length - 1; i++) {
    const [e, t, s] = between[i];
    if (e != empIdOut) continue;
    if (t === '休憩開始') last = new Date(s);
    if (t === '休憩終了' && last) {
      breakMs += (new Date(s) - last);
      last = null;
    }
  }

  const workingTime = Math.max(0, (outAt - inAt) - breakMs);
  const m = Math.floor(workingTime / 60000);
  const hh = ('0' + Math.floor(m / 60)).slice(-2);
  const mm = ('0' + (m % 60)).slice(-2);

  sh.getRange(rowOut, 5).setValue(hh + ':' + mm);
}


// 直近の勤怠データ削除
function deleteLastWork() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]; // 打刻履歴
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) { 
    return "削除できる勤怠データがありません。";
  }
  sheet.deleteRow(lastRow);
  return "直近の勤怠データを削除しました。";
}


// password
const PASSWORD = "yusuke";  // 勤怠用
const ADMIN_PASSWORD = "yusuke"; // 塾長用

function deleteLastWorkWithPassword(password) {
  if (password !== PASSWORD) {
    throw new Error("パスワードが違います。");
  }
  return deleteLastWork();
}

function checkAdminPassword(pw) {
  if (pw !== ADMIN_PASSWORD) {
    throw new Error("パスワードが違います。");
  }
  return "OK";
}


// メモ関連
function getEmpMemo() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  var checkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var last_row = checkSheet.getLastRow();
  var timeClocksRange = checkSheet.getRange(2, 1, last_row, 2);

  var checkResult = "";
  var i = 1;
  while (true) {
    var empId = timeClocksRange.getCell(i, 1).getValue();
    var result = timeClocksRange.getCell(i, 2).getValue();
    if (empId === "") break;
    if (empId == selectedEmpId){
      checkResult = result;
      break;
    }
    i++;
  }
  return checkResult;
}




function saveMemo(form) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  var memo = form.memo;

  var targetRowNumber = getTargetEmpRowNumber(selectedEmpId);
  var sheet = SpreadsheetApp.getActiveSheet();

  if (targetRowNumber == null) {
    targetRowNumber = sheet.getLastRow() + 1;
    sheet.getRange(targetRowNumber, 1).setValue(selectedEmpId);
  }
  sheet.getRange(targetRowNumber, 2).setValue(memo);
}

function getTargetEmpRowNumber(empId) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var last_row = sheet.getLastRow();
  var data_range = sheet.getRange(1, 1, last_row, 1);
  var sheetRows = data_range.getValues();

  for (var i = 0; i <= sheetRows.length - 1; i++) {
    var row = sheetRows[i];
    if (row[0] == empId) {
      return i + 1;
    }
  }
  return null;
}


// 生徒一覧取得
function getStudents() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3];
  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2, 6, last - 1, 1).getValues();
  var seen = {};
  var list = [];

  for (var i = 0; i < vals.length; i++) {
    var name = String(vals[i][0] || '').trim();
    if (name && !seen[name]) {
      seen[name] = true;
      list.push(name);
    }
  }

  list.sort(function(a,b){ return a.localeCompare(b, 'ja'); });
  return list;
}

function getFeedback() { return ""; }


// 科目ごとの労働時間
function getSubjectHours(empId){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3];
  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2,1,last-1,6).getValues();
  var result = {};

  for (var i=0; i<vals.length; i++){
    var id   = vals[i][0];
    var type = vals[i][1];
    var wt   = vals[i][4];

    if (String(id) !== String(empId)) continue;
    if (!wt || typeof wt !== "string") continue;

    var subject = vals[i][3] || "未設定";
    var parts = wt.split(":");
    var h = parseInt(parts[0],10);
    var m = parseInt(parts[1],10);

    if (!result[subject]) result[subject] = {h:0, m:0};
    result[subject].h += h;
    result[subject].m += m;
  }

  return Object.keys(result).map(function(subj){
    var totalH = result[subj].h;
    var totalM = result[subj].m;
    totalH += Math.floor(totalM / 60);
    totalM  = totalM % 60;
    return {subject: subj, hoursStr: totalH+"時間"+totalM+"分", hours: totalH + totalM/60};
  });
}


// 給与集計（全体）
function getSalaryData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[3];
  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2, 1, last - 1, 6).getValues();

  // 時給マップ
  var wageSh = ss.getSheetByName("給与設定");
  var wageMap = {};
  if (wageSh) {
    var wVals = wageSh.getRange(2,1,wageSh.getLastRow()-1,3).getValues();
    wVals.forEach(function(r){
      wageMap[r[0]+"_"+r[1]] = r[2];
    });
  }

  var empMap = {};
  vals.forEach(function(row){
    var empId   = String(row[0] || "");
    var subject = String(row[3] || "その他");
    var wt      = row[4];
    var empName = getEmployeeNameById(empId);
    if (!empName) return;

    var minutes = 0;
    if (wt instanceof Date) {
      minutes = wt.getHours() * 60 + wt.getMinutes();
    } else if (typeof wt === "string" && wt.match(/^\d{1,2}:\d{2}$/)) {
      var parts = wt.split(":");
      minutes = parseInt(parts[0],10) * 60 + parseInt(parts[1],10);
    }

    if (!empMap[empName]) empMap[empName] = {};
    if (!empMap[empName][subject]) empMap[empName][subject] = 0;
    empMap[empName][subject] += minutes;
  });

  var result = [];
  for (var emp in empMap) {
    var subjects = [];
    for (var subj in empMap[emp]) {
      var mins = empMap[emp][subj];
      var hours = (mins / 60).toFixed(2);
      var key = emp+"_"+subj;
      var wage = wageMap[key] || 0;
      subjects.push({
        subject: subj,
        hoursStr: Math.floor(mins/60) + "時間" + (mins%60) + "分",
        hours: parseFloat(hours),
        wage: wage
      });
    }
    result.push({ employee: emp, subjects: subjects });
  }
  return result;
}
// 年別給与集計
function getYearlySalaryData(year) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[3]; // 打刻履歴
  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2, 1, last - 1, 6).getValues();
  var empMap = {};

  vals.forEach(function(row){
    var empId  = String(row[0] || "");
    var type   = row[1];
    var dt     = row[2];
    var subject= String(row[3] || "その他");
    var wt     = row[4];

    if (!(dt instanceof Date)) return;
    if (dt.getFullYear() !== year) return;

    var empName = getEmployeeNameById(empId);
    if (!empName) return;

    var minutes = 0;
    if (typeof wt === "string" && wt.match(/^\d{2}:\d{2}$/)) {
      var parts = wt.split(":");
      minutes = parseInt(parts[0]) * 60 + parseInt(parts[1]);
    }

    if (!empMap[empName]) empMap[empName] = {};
    if (!empMap[empName][subject]) empMap[empName][subject] = 0;
    empMap[empName][subject] += minutes;
  });

  var result = [];
  for (var emp in empMap) {
    var subjects = [];
    for (var subj in empMap[emp]) {
      var mins = empMap[emp][subj];
      subjects.push({
        subject: subj,
        hoursStr: Math.floor(mins/60) + "時間" + (mins%60) + "分",
        hours: (mins / 60)
      });
    }
    result.push({ employee: emp, subjects: subjects });
  }
  return result;
}


// 利用可能な年一覧
function getAvailableYears() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[3];
  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2, 3, last - 1, 1).getValues(); // C列=日時
  var years = {};
  vals.forEach(function(r){
    var d = r[0];
    if (d instanceof Date) {
      years[d.getFullYear()] = true;
    }
  });
  return Object.keys(years).sort().reverse(); // 新しい順
}


// 従業員IDから名前を取得
function getEmployeeNameById(empId) {
  const employees = getEmployees();
  for (var i=0;i<employees.length;i++){
    if (employees[i].id === empId) return employees[i].name;
  }
  return "";
}


// フィードバック保存
function saveFeedback(row, feedback) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3];
  var last = sh.getLastRow();
  if (last < 2) return "保存対象なし";
  if (!row || row < 2 || row > last) {
    return "対象の行番号が不正です";
  }
  sh.getRange(row, 7).setValue(feedback); // G列
  return "OK";
}


// 月別給与集計
function getMonthlySalaryData(year, month) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3];
  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2, 1, last - 1, 6).getValues();
  var empMap = {};

  vals.forEach(function(row){
    var empId   = String(row[0] || "");
    var type    = row[1];
    var dt      = new Date(row[2]);
    var subject = String(row[3] || "その他");
    var wt      = row[4];
    var empName = getEmployeeNameById(empId);
    if (!empName) return;

    if (dt.getFullYear() !== year || (dt.getMonth()+1) !== month) return;

    var minutes = 0;
    if (wt instanceof Date) {
      minutes = wt.getHours() * 60 + wt.getMinutes();
    } else if (typeof wt === "string" && wt.match(/^\d{2}:\d{2}$/)) {
      var parts = wt.split(":");
      minutes = parseInt(parts[0]) * 60 + parseInt(parts[1]);
    }

    if (!empMap[empName]) empMap[empName] = {};
    if (!empMap[empName][subject]) empMap[empName][subject] = 0;
    empMap[empName][subject] += minutes;
  });

  var result = [];
  for (var emp in empMap) {
    var subjects = [];
    for (var subj in empMap[emp]) {
      var mins = empMap[emp][subj];
      var hours = (mins / 60).toFixed(2);
      subjects.push({
        subject: subj,
        hoursStr: Math.floor(mins/60) + "時間" + (mins%60) + "分",
        hours: parseFloat(hours)
      });
    }
    result.push({ employee: emp, subjects: subjects });
  }
  return result;
}


// 給与設定保存
function saveWage(empName, subject, wage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("給与設定");
  if (!sh) sh = ss.insertSheet("給与設定");

  var last = sh.getLastRow();
  var range = sh.getRange(2, 1, last-1, 3).getValues();

  for (var i=0; i<range.length; i++) {
    if (range[i][0] === empName && range[i][1] === subject) {
      sh.getRange(i+2, 3).setValue(wage);
      return;
    }
  }
  sh.appendRow([empName, subject, wage]);
}


// 打刻履歴から利用可能な月を取得
function getAvailableMonths() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheets()[3];
  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2, 3, last-1, 1).getValues(); // C列=日時
  var months = {};
  vals.forEach(function(r){
    var d = r[0];
    if (d instanceof Date) {
      var y = d.getFullYear();
      var m = d.getMonth() + 1;
      var key = y + "-" + ("0" + m).slice(-2);
      months[key] = true;
    }
  });
  return Object.keys(months).sort().reverse();
}
// 科目一覧を返す
function getSubjects() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]; // 打刻履歴シート
  var last = sh.getLastRow();
  if (last < 2) return [];
  
  var vals = sh.getRange(2, 4, last - 1, 1).getValues(); // D列=科目
  var seen = {};
  var list = [];
  vals.forEach(function(r){
    var subj = String(r[0] || "").trim();
    if (subj && !seen[subj]) {
      seen[subj] = true;
      list.push(subj);
    }
  });
  list.sort();
  return list;
}
function getLessonSessions() {
  var empId   = getSelectedEmpId();
  var student = getSelectedStudent();
  if (!empId || !student) return [];

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]; // 打刻履歴
  var last = sh.getLastRow();
  if (last < 2) return [];

  var vals = sh.getRange(2, 1, last - 1, 7).getValues(); // A〜G列
  var rows = vals.filter(function(r){
    return String(r[0]) === String(empId) && String(r[5]) === String(student);
  }).sort(function(a,b){
    return new Date(a[2]) - new Date(b[2]);
  });

  // フォーム回答マップを取得
  var answeredMap = getAnsweredSessions();

  var sessions = [];
  var currentStart = null;
  var currentSubject = "";
  for (var i = 0; i < rows.length; i++) {
    var type = rows[i][1];
    var dt   = new Date(rows[i][2]);
    var subj = rows[i][3] || "—";
    var fb   = rows[i][6] || "";
    var stu  = rows[i][5] || "";

    if (type === '出勤') {
      currentStart = dt;
      currentSubject = subj;
    }
    if (type === '退勤' && currentStart) {
      var startStr = Utilities.formatDate(currentStart, "Asia/Tokyo", "yyyy-MM-dd HH:mm");
      var endStr   = Utilities.formatDate(dt, "Asia/Tokyo", "yyyy-MM-dd HH:mm");

      var key = startStr + "_" + stu;
      Logger.log("セッションキー: " + key); // 🔍 デバッグ用ログ

      var answered = answeredMap[key] ? "回答済み" : "未回答";

      sessions.push({
        start: startStr,
        end: endStr,
        empName: getEmployeeNameById(empId),
        subject: currentSubject,
        feedback: fb,
        student: stu,
        answered: answered,  
        row: i+2
      });
      currentStart = null;
      currentSubject = "";
    }
  }

  return sessions;
}


function getAnsweredSessions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSh = ss.getSheetByName("フォームの回答 1");
  if (!formSh) return {};
  
  var last = formSh.getLastRow();
  if (last < 2) return {};

  var vals = formSh.getRange(2, 2, last - 1, 5).getValues(); // B～F列
  var answeredMap = {};

  vals.forEach(function(r){
    var start = r[0];   // 授業開始時間
    var stu   = String(r[4] || "").trim();

    if (start && stu) {
      var dt = new Date(start); // 文字列でもDateでもここで統一
      if (!isNaN(dt)) {
        var key = Utilities.formatDate(dt, "Asia/Tokyo", "yyyy-MM-dd HH:mm") + "_" + stu;
        Logger.log("回答キー: " + key); //  デバッグ用
        answeredMap[key] = true;
      }
    }
  });
  return answeredMap;
}

function saveFeedbackRow(row, inputId){
  var val = document.getElementById(inputId).value;
  google.script.run
    .withSuccessHandler(function(res){
      if (res === "OK") {
        alert("保存しました！");
      } else {
        alert("エラー: " + res);
      }
    })
    .saveFeedback(row, val);
}