/**
 * mini-CEX.tw.entrust 平台 — Google Sheets 資料後端
 *
 * 部署方式：
 * 1. 開啟你的 Google 試算表 → 擴充功能 → Apps Script
 * 2. 把本檔內容貼上，存檔
 * 3. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *    - 執行身分：我
 *    - 具有存取權的使用者：測試階段可選「知道連結的任何人」；接真實資料時務必改為
 *      「機構內的任何使用者」或「僅限我」，並改用需登入的驗證流程
 * 4. 複製產生的網址，填入前端 index.html 的 window.__SHEETS_API__
 *
 * 安全提醒：
 * 本腳本的權限判斷僅供原型使用。目前 role 由前端傳入，可被偽造。
 * 接真實資料前，必須改為由 Session.getActiveUser().getEmail() 取得登入者，
 * 再回頭查 Users 表決定角色，絕不可信任前端傳來的身分。
 */

/**
 * 【必填】試算表 ID
 *
 * 若這份腳本是從試算表的「擴充功能 → Apps Script」建立的（綁定型），
 * 這裡留空字串即可，腳本會自動使用所屬的試算表。
 *
 * 若這份腳本是從 script.google.com 直接建立的（獨立型），
 * 必須填入試算表 ID，否則會出現
 * 「Cannot read properties of null (reading 'getSheetByName')」。
 *
 * ID 是試算表網址中間那一長串，例如：
 * https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz/edit
 *                                        ^^^^^^^^^^^^^^^^^^^^^^^^^^ 這一段
 */
var SPREADSHEET_ID = '1JxVytbX9WzKKFHO5OlYsjD2f8DFNU5hgmXGLHSgpo20';

/** 取得試算表。獨立型腳本用 ID 開啟，綁定型腳本用所屬試算表 */
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      '找不到試算表。這是獨立型 Apps Script 專案，請在 Code.gs 最上方的 SPREADSHEET_ID 填入試算表 ID。'
    );
  }
  return ss;
}

var SHEETS = [
  'Users', 'Depts', 'Evaluations', 'Rotations', 'Officers',
  'DeptWeights', 'Scores', 'GlobalScores', 'RubricFiles',
  'ContentPages', 'AuditLogs', 'ContentTasks'
];

/** 讀取單一工作表，回傳物件陣列（第一列為欄名） */
function readSheet(name) {
  var sh = getSpreadsheet().getSheetByName(name);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var header = values[0].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row.every(function (c) { return c === '' || c === null; })) continue;
    var obj = {};
    for (var j = 0; j < header.length; j++) {
      if (!header[j]) continue;
      obj[header[j]] = normalize(row[j]);
    }
    out.push(obj);
  }
  return out;
}

/** 日期物件轉字串，布林字串轉真布林 */
function normalize(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      .replace(' 00:00', '');
  }
  if (v === 'TRUE' || v === true) return true;
  if (v === 'FALSE' || v === false) return false;
  return v;
}

/** 依角色過濾評量資料。原型版本，正式版須改為後端驗證身分後判斷 */
function scopeEvaluations(rows, role, staffId, scopeDeptCode) {
  if (role === 'student') {
    return rows.filter(function (r) { return r.trainee_staff_id === staffId; });
  }
  if (role === 'teacher') {
    return rows.filter(function (r) { return r.teacher_staff_id === staffId; });
  }
  if (role === 'deptAdmin' || role === 'deptPromoter') {
    if (!scopeDeptCode) return [];
    return rows.filter(function (r) {
      return r.dept_code === scopeDeptCode ||
             String(r.dept_code).indexOf(scopeDeptCode + '-') === 0;
    });
  }
  return rows; // hospital / sysAdmin
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    var payload = {};

    if (p.sheet) {
      payload[p.sheet] = readSheet(p.sheet);
    } else {
      SHEETS.forEach(function (n) { payload[n] = readSheet(n); });
    }

    if (payload.Evaluations && p.role) {
      payload.Evaluations = scopeEvaluations(
        payload.Evaluations, p.role, p.staffId || '', p.scopeDept || ''
      );
    }

    return json({ ok: true, generatedAt: new Date().toISOString(), data: payload });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** 寫入用：目前支援新增稽核紀錄與教材待辦 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var name = body.sheet;
    if (SHEETS.indexOf(name) === -1) throw new Error('未知的工作表：' + name);
    if (name !== 'AuditLogs' && name !== 'ContentTasks') {
      throw new Error('此工作表不開放寫入');
    }
    var sh = getSpreadsheet().getSheetByName(name);
    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var row = header.map(function (h) { return body.row[h] !== undefined ? body.row[h] : ''; });
    sh.appendRow(row);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 選單：一鍵建立所有工作表與欄位標題（首次設定用）
 * 僅綁定型腳本會出現此選單。獨立型腳本請直接在編輯器執行 setupSheets。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('平台資料')
    .addItem('建立所有工作表', 'setupSheets')
    .addToUi();
}

var HEADERS = {
  Users: ['user_id', 'staff_id', 'name_zh', 'name_en', 'email', 'title_zh', 'title_en', 'rank', 'dept_code', 'scope_dept_code', 'roles', 'grade', 'active'],
  Depts: ['dept_code', 'dept_zh', 'dept_en', 'parent_code', 'is_sub', 'active'],
  Evaluations: ['eval_id', 'visit_datetime', 'apply_datetime', 'eval_datetime', 'trainee_staff_id', 'teacher_staff_id', 'dept_code', 'complexity', 'entrust_level', 'item_interview', 'item_physical', 'item_procedure', 'item_counseling', 'item_judgment', 'item_organization', 'item_professional', 'praise', 'suggestion', 'status', 'in_rotation_flag'],
  Rotations: ['rotation_id', 'trainee_staff_id', 'dept_code', 'start_date', 'end_date', 'weeks'],
  Officers: ['officer_id', 'staff_id', 'dept_code', 'since_date', 'until_date', 'active'],
  DeptWeights: ['weight_id', 'dept_code', 'track', 'academic_year', 'item_zh', 'item_en', 'weight_pct', 'source_zh', 'source_en'],
  Scores: ['score_id', 'trainee_staff_id', 'dept_code', 'academic_year', 'score', 'sub_weight_pct', 'status'],
  GlobalScores: ['trainee_staff_id', 'academic_year', 'edu_office_score', 'adjustment', 'dept_ratio', 'edu_ratio'],
  RubricFiles: ['rubric_id', 'academic_year', 'track', 'file_name', 'file_url', 'updated_date'],
  ContentPages: ['page_id', 'kind', 'name_zh', 'name_en', 'enabled', 'body_url'],
  AuditLogs: ['log_id', 'timestamp', 'actor_staff_id', 'action_zh', 'action_en', 'target', 'device', 'ip'],
  ContentTasks: ['task_id', 'created_at', 'insight_text', 'mention_count', 'dept_code', 'assignee_staff_id', 'status']
};

function setupSheets() {
  var ss = getSpreadsheet();
  var created = [];
  SHEETS.forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      var h = HEADERS[name];
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
      sh.setFrozenRows(1);
      created.push(name);
    }
  });
  var msg = '工作表已建立完成（' + created.length + ' 張）。請依 CSV 範本貼入資料。';
  Logger.log(msg);
  // 獨立型腳本沒有試算表介面，getUi() 會失敗，因此包在 try 中
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    // 從編輯器直接執行時，改看「執行紀錄」即可
  }
  return msg;
}

/** 診斷用：從編輯器執行這個函式，可確認是否連得上試算表 */
function testConnection() {
  var ss = getSpreadsheet();
  var names = ss.getSheets().map(function (s) { return s.getName(); });
  var msg = '連線成功：' + ss.getName() + '\n工作表：' + names.join(', ');
  Logger.log(msg);
  return msg;
}
