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
  'ContentPages', 'AuditLogs', 'AuditChanges', 'ContentTasks', 'Assessments'
];

/** 稽核表為唯讀對外，只能透過 logEvent 寫入，不開放前端任意 append */
var APPEND_ALLOWED = ['ContentTasks', 'Assessments'];

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

/**
 * 讀取端點。
 *
 * 身分一律由後端從 Google 登入狀態取得，絕不接受前端傳來的角色參數。
 * 這是整個權限模型的基礎：前端程式碼是公開的，任何人都能改，
 * 因此前端說自己是誰完全不能採信。
 */
function doGet(e) {
  var p = (e && e.parameter) || {};

  // 沒帶 api 參數時，直接把前端網頁吐出來（同源，可取得登入身分）
  if (!p.api) return serveApp();

  try {
    var me = getIdentity();
    if (!me) {
      return json({ ok: false, error: 'NOT_AUTHENTICATED', message: '請先以院內帳號登入。' });
    }
    if (!me.active) {
      return json({ ok: false, error: 'ACCOUNT_DISABLED', message: '此帳號已停用。' });
    }

    var payload = {};
    SHEETS.forEach(function (n) {
      if (n === 'AuditLogs' || n === 'AuditChanges') return;   // 稽核另外走 canSeeAudit
      payload[n] = readSheet(n);
    });

    // 依身分裁切。裁切在回傳之前完成，前端拿不到範圍外的資料。
    payload.Evaluations = scopeEvaluations(payload.Evaluations, me.roles, me.staffId, me.scopeDept);
    payload.Scores = scopeScores(payload.Scores, me.roles, me.staffId);
    payload.Users = scopeUsers(payload.Users, me.roles);
    payload.DeptStats = deptStatsFor(me.roles, me.scopeDept);

    if (canSeeAudit(me.roles)) {
      payload.AuditLogs = readSheet('AuditLogs');
      payload.AuditChanges = readSheet('AuditChanges');
    }

    logEvent(actorOf(me, p), 'read', 'load_data', 'session', '', 'success', '', '');

    return json({
      ok: true,
      generatedAt: new Date().toISOString(),
      me: { staffId: me.staffId, name: me.name, roles: me.roles, dept: me.dept, scopeDept: me.scopeDept },
      data: payload
    });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** 把打包好的前端網頁送出。前端檔案存為 Apps Script 專案中的 index.html */
function serveApp() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('mini-CEX.tw.entrust')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * 取得目前登入者。
 * 必須部署為「執行身分：存取應用程式的使用者」才能取得 email。
 * 回傳 null 表示未登入或不在 Users 名冊中。
 */
function getIdentity() {
  var email = Session.getActiveUser().getEmail();
  if (!email) return null;

  var rows = readSheet('Users');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].email || '').toLowerCase() === email.toLowerCase()) {
      return {
        email: email,
        staffId: String(rows[i].staff_id || ''),
        name: String(rows[i].name || ''),
        roles: String(rows[i].roles || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean),
        dept: String(rows[i].dept_code || ''),
        scopeDept: String(rows[i].scope_dept || rows[i].dept_code || ''),
        active: String(rows[i].active || 'TRUE').toUpperCase() !== 'FALSE'
      };
    }
  }
  return null;
}

function has(roles, r) { return roles.indexOf(r) !== -1; }
function isHospital(roles) { return has(roles, 'hospital') || has(roles, 'sysAdmin'); }
function isDeptScoped(roles) { return has(roles, 'deptAdmin') || has(roles, 'deptPromoter'); }
function canSeeAudit(roles) { return has(roles, 'sysAdmin'); }

function actorOf(me, p) {
  return {
    sessionId: (p && p.session) || '',
    staffId: me.staffId, name: me.name, role: me.roles.join(','),
    device: (p && p.device) || '', userAgent: (p && p.ua) || '',
    ip: '(未取得)', source: 'web'
  };
}

/** 評量：學員只見自己被評的，教師只見自己評的，科部限負責科含次專科 */
function scopeEvaluations(rows, roles, staffId, scopeDept) {
  if (isHospital(roles)) return rows;
  return rows.filter(function (r) {
    if (isDeptScoped(roles) && String(r.dept_code || '').indexOf(scopeDept) === 0) return true;
    if (has(roles, 'teacher') && String(r.teacher_staff_id) === staffId) return true;
    if (has(roles, 'student') && String(r.student_staff_id) === staffId) return true;
    return false;
  });
}

/** 成績：學員只見自己的 */
function scopeScores(rows, roles, staffId) {
  if (isHospital(roles) || isDeptScoped(roles)) return rows;
  if (has(roles, 'student')) {
    return rows.filter(function (r) { return String(r.student_staff_id) === staffId; });
  }
  return [];
}

/** 名冊：只有系統管理者看得到完整名冊含聯絡方式，其餘只給顯示用的最小欄位 */
function scopeUsers(rows, roles) {
  if (has(roles, 'sysAdmin')) return rows;
  return rows.map(function (r) {
    return { staff_id: r.staff_id, name: r.name, title: r.title, dept_code: r.dept_code };
  });
}

/** 跨科比較：科部角色只拿得到自己負責的科，全院角色才拿得到全部 */
function deptStatsFor(roles, scopeDept) {
  var evals = readSheet('Evaluations');
  var depts = readSheet('Depts');
  var out = [];
  depts.forEach(function (d) {
    var code = String(d.dept_code || '');
    if (isDeptScoped(roles) && !isHospital(roles) && code.indexOf(scopeDept) !== 0) return;
    var mine = evals.filter(function (e) { return String(e.dept_code || '').indexOf(code) === 0; });
    if (mine.length === 0) return;
    var done = mine.filter(function (e) { return String(e.status) === '已評量'; });
    var sum = 0, n = 0;
    done.forEach(function (e) {
      var v = Number(e.entrust_level);
      if (v) { sum += v; n++; }
    });
    out.push({
      dept_code: code,
      name: d.name,
      parent_code: d.parent_code || '',
      total: mine.length,
      done: done.length,
      completion: mine.length ? Math.round(done.length / mine.length * 100) : 0,
      avg: n ? Math.round(sum / n * 10) / 10 : 0
    });
  });
  return out;
}

/**
 * 寫入端點。body.op 決定行為：
 *   append  — 新增一列到白名單工作表
 *   update  — 更新既有列，並自動比對欄位差異寫入稽核
 *   log     — 僅寫入一筆事件（登入、登出、檢視、匯出等不改資料的行為）
 * 所有寫入一律留下稽核紀錄，無法略過。
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var op = body.op || 'log';

    // 身分由後端取得，忽略前端送來的 actor
    var me = getIdentity();
    if (!me || !me.active) return json({ ok: false, error: 'NOT_AUTHENTICATED' });
    var actor = actorOf(me, body);

    // 寫入權限檢查。前端送什麼都不影響這裡的判斷。
    if ((op === 'append' || op === 'update') && !canWrite(me.roles, body.sheet)) {
      logEvent(actor, 'write', body.action || ('deny_' + op), body.sheet || '', body.keyValue || '',
               'fail', '權限不足', '');
      return json({ ok: false, error: 'FORBIDDEN', message: '你的角色沒有修改此資料的權限。' });
    }

    if (op === 'log') {
      var id = logEvent(actor, body.category || 'read', body.action || 'unknown',
                        body.targetType || '', body.targetId || '',
                        body.result || 'success', body.failReason || '', '');
      return json({ ok: true, logId: id });
    }

    if (op === 'append') {
      var name = body.sheet;
      if (APPEND_ALLOWED.indexOf(name) === -1) throw new Error('此工作表不開放新增：' + name);
      var sh = getSpreadsheet().getSheetByName(name);
      if (!sh) throw new Error('找不到工作表：' + name);
      // 身分欄位一律以後端取得的為準，忽略前端送來的值
      if (name === 'Assessments') {
        body.row.student_staff_id = me.staffId;
        body.row.student_name = me.name;
        body.row.form_id = nextId('F', 'Assessments', 'form_id');
        body.row.applied_at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        body.row.source = 'web';
      }

      var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      sh.appendRow(header.map(function (h) { return body.row[h] !== undefined ? body.row[h] : ''; }));

      var csId = nextId('CS', 'AuditChanges', 'change_set_id');
      header.forEach(function (h) {
        if (!h || body.row[h] === undefined || body.row[h] === '') return;
        writeChange(csId, name, body.row[header[0]] || '', h, '', body.row[h]);
      });
      logEvent(actor, 'write', 'create_' + name, name, body.row[header[0]] || '', 'success', '', csId);
      return json({ ok: true, changeSetId: csId });
    }

    if (op === 'update') {
      var name2 = body.sheet;
      if (SHEETS.indexOf(name2) === -1) throw new Error('未知的工作表：' + name2);
      if (name2 === 'AuditLogs' || name2 === 'AuditChanges') throw new Error('稽核表不可修改');
      var sh2 = getSpreadsheet().getSheetByName(name2);
      if (!sh2) throw new Error('找不到工作表：' + name2);

      var values = sh2.getDataRange().getValues();
      var header2 = values[0].map(function (h) { return String(h).trim(); });
      var keyCol = header2.indexOf(body.keyField);
      if (keyCol < 0) throw new Error('找不到鍵值欄位：' + body.keyField);

      var rowIndex = -1;
      for (var i = 1; i < values.length; i++) {
        if (String(values[i][keyCol]) === String(body.keyValue)) { rowIndex = i; break; }
      }
      if (rowIndex < 0) throw new Error('找不到資料列：' + body.keyValue);

      var own = canWriteRow(me, name2, body.keyValue, body.changes);
      if (!own.ok) {
        logEvent(actor, 'write', body.action || ('deny_update_' + name2), name2, body.keyValue,
                 'fail', own.why, '');
        return json({ ok: false, error: 'FORBIDDEN', message: own.why });
      }

      var csId2 = nextId('CS', 'AuditChanges', 'change_set_id');
      var changed = 0;
      Object.keys(body.changes || {}).forEach(function (field) {
        var col = header2.indexOf(field);
        if (col < 0) return;
        var before = normalize(values[rowIndex][col]);
        var after = body.changes[field];
        if (String(before) === String(after)) return;
        sh2.getRange(rowIndex + 1, col + 1).setValue(after);
        writeChange(csId2, name2, body.keyValue, field, before, after);
        changed++;
      });

      logEvent(actor, 'write', body.action || ('update_' + name2), name2, body.keyValue,
               'success', '', changed > 0 ? csId2 : '');
      return json({ ok: true, changedFields: changed, changeSetId: changed > 0 ? csId2 : null });
    }

    throw new Error('未知的 op：' + op);
  } catch (err) {
    try {
      var b = JSON.parse(e.postData.contents || '{}');
      logEvent(b.actor || {}, 'write', b.action || 'unknown', b.sheet || '', b.keyValue || '',
               'fail', String(err), '');
    } catch (ignored) {}
    return json({ ok: false, error: String(err) });
  }
}

/** 誰能改哪張表。這是後端唯一的真相來源。 */
function canWrite(roles, sheet) {
  if (sheet === 'AuditLogs' || sheet === 'AuditChanges') return false;   // 任何人都不可改稽核
  if (has(roles, 'sysAdmin')) return true;
  if (sheet === 'Evaluations') return has(roles, 'teacher');
  if (sheet === 'ContentTasks') return isDeptScoped(roles) || has(roles, 'hospital');
  if (sheet === 'Assessments') return has(roles, 'student') || has(roles, 'teacher');
  return false;
}

/**
 * 逐列的所有權檢查。canWrite 只管「能不能碰這張表」，
 * 這裡管「能不能碰這一列」。少了這一層，任何學員都能改別人的申請。
 */
function canWriteRow(me, sheet, keyValue, changes) {
  if (has(me.roles, 'sysAdmin')) return { ok: true };
  if (sheet !== 'Assessments') return { ok: true };

  var sh = getSpreadsheet().getSheetByName('Assessments');
  if (!sh) return { ok: false, why: '找不到工作表' };
  var rows = readSheet('Assessments');
  var row = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].form_id) === String(keyValue)) { row = rows[i]; break; }
  }
  if (!row) return { ok: false, why: '找不到該筆申請' };

  var fields = Object.keys(changes || {});

  // 學員：只能改自己的申請，且只能改醫療資訊三欄
  if (String(row.student_staff_id) === me.staffId) {
    var allowed = ['chart_no', 'location', 'condition', 'visit_at', 'status'];
    for (var j = 0; j < fields.length; j++) {
      if (allowed.indexOf(fields[j]) === -1) {
        return { ok: false, why: '學員不可修改評量內容：' + fields[j] };
      }
    }
    // 已評量後 7 日內才可修改
    if (String(row.status) === '已評量') {
      var days = (new Date() - new Date(row.assessed_at)) / 86400000;
      if (days > 7) return { ok: false, why: '已評量超過 7 日，不可再修改' };
    }
    return { ok: true };
  }

  // 教師：只能評自己被指定的，且不可改學員填的醫療資訊
  if (String(row.teacher_staff_id) === me.staffId) {
    var blocked = ['chart_no', 'location', 'condition', 'student_staff_id', 'teacher_staff_id'];
    for (var k = 0; k < fields.length; k++) {
      if (blocked.indexOf(fields[k]) !== -1) {
        return { ok: false, why: '教師不可修改學員填寫的醫療資訊：' + fields[k] };
      }
    }
    if (String(row.status) === '未填寫') {
      return { ok: false, why: '學員醫療資訊未填寫完整，無法進行評量' };
    }
    return { ok: true };
  }

  return { ok: false, why: '這筆申請與你無關' };
}

/** 寫入一筆事件到 AuditLogs。回傳 log_id */
function logEvent(actor, category, action, targetType, targetId, result, failReason, changeSetId) {
  var sh = getSpreadsheet().getSheetByName('AuditLogs');
  if (!sh) return '';
  var id = nextId('L', 'AuditLogs', 'log_id');
  sh.appendRow([
    id,
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    actor.sessionId || '',
    actor.staffId || '',
    actor.name || '',
    actor.role || '',
    category, action, targetType, targetId,
    result, failReason, changeSetId,
    actor.device || '', actor.userAgent || '',
    actor.ip || '(未取得)',
    actor.source || 'web'
  ]);
  return id;
}

/** 寫入一筆欄位層級的異動明細 */
function writeChange(changeSetId, targetType, targetId, field, before, after) {
  var sh = getSpreadsheet().getSheetByName('AuditChanges');
  if (!sh) return;
  sh.appendRow([
    nextId('C', 'AuditChanges', 'change_id'),
    changeSetId, targetType, targetId, field,
    before === null || before === undefined ? '' : before,
    after === null || after === undefined ? '' : after
  ]);
}

/** 產生遞增流水號，例如 L0016、CS-0005 */
function nextId(prefix, sheetName, column) {
  var sh = getSpreadsheet().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return prefix + (prefix === 'CS' ? '-0001' : '0001');
  var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var col = header.indexOf(column);
  if (col < 0) return prefix + (prefix === 'CS' ? '-0001' : '0001');
  var vals = sh.getRange(2, col + 1, sh.getLastRow() - 1, 1).getValues();
  var max = 0;
  vals.forEach(function (r) {
    var m = String(r[0]).match(/(\d+)$/);
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  var n = ('0000' + (max + 1)).slice(-4);
  return prefix === 'CS' ? 'CS-' + n : prefix + n;
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
  AuditLogs: ['log_id', 'timestamp', 'session_id', 'actor_staff_id', 'actor_name', 'actor_role', 'category', 'action', 'target_type', 'target_id', 'result', 'fail_reason', 'change_set_id', 'device', 'user_agent', 'ip', 'source'],
  Assessments: [
    'form_id', 'status', 'student_staff_id', 'student_name', 'student_rank',
    'teacher_staff_id', 'teacher_name', 'teacher_rank',
    'dept_code', 'sub_dept', 'applied_at', 'visit_at', 'assessed_at',
    'chart_no', 'location', 'condition', 'item',
    'complexity', 'entrust',
    'r_interview', 'r_exam', 'r_procedure', 'r_counseling', 'r_judgment', 'r_efficiency', 'r_humanistic',
    'phrases', 'narrative', 'source'
  ],
  AuditChanges: ['change_id', 'change_set_id', 'target_type', 'target_id', 'field', 'before_value', 'after_value'],
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
