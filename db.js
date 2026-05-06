'use strict';
const { DatabaseSync } = require('node:sqlite');
const path = require('path');
const crypto = require('crypto');

const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'data', 'workiq.db');
const db = new DatabaseSync(DB_PATH);

// ── Schema ────────────────────────────────────────────────────────────────────
db.exec(`
  PRAGMA journal_mode=WAL;
  PRAGMA foreign_keys=ON;

  CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS calendars (
    id         TEXT PRIMARY KEY,
    name       TEXT NOT NULL,
    timezone   TEXT NOT NULL DEFAULT 'UTC',
    weekends   TEXT NOT NULL DEFAULT 'Sat-Sun',
    color      TEXT NOT NULL DEFAULT '#0078d4',
    holidays   TEXT NOT NULL DEFAULT '[]',
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS departments (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT NOT NULL UNIQUE,
    teams      TEXT NOT NULL DEFAULT '[]',
    cal_id     TEXT REFERENCES calendars(id),
    created_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS employees (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    entra_id     TEXT UNIQUE,
    name         TEXT NOT NULL,
    email        TEXT UNIQUE,
    initials     TEXT NOT NULL DEFAULT 'XX',
    dept         TEXT,
    team         TEXT,
    cal_id       TEXT REFERENCES calendars(id),
    role         TEXT NOT NULL DEFAULT 'employee',
    title        TEXT,
    bal_type     TEXT NOT NULL DEFAULT 'fixed',
    password_hash TEXT,
    av_bg        TEXT DEFAULT '#e6f2fb',
    av_color     TEXT DEFAULT '#004e8c',
    active       INTEGER DEFAULT 1,
    created_at   TEXT DEFAULT (datetime('now')),
    updated_at   TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS leave_types (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    name         TEXT NOT NULL UNIQUE,
    days_per_year INTEGER NOT NULL DEFAULT 14,
    approval_levels INTEGER NOT NULL DEFAULT 2,
    carry_over   TEXT NOT NULL DEFAULT 'no',
    carry_over_max INTEGER DEFAULT 0,
    doc_required INTEGER DEFAULT 0,
    applies_to   TEXT DEFAULT 'all',
    color        TEXT DEFAULT '#0078d4',
    active       INTEGER DEFAULT 1,
    created_at   TEXT DEFAULT (datetime('now')),
    updated_at   TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS leave_balances (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id  INTEGER NOT NULL REFERENCES employees(id),
    leave_type_id INTEGER NOT NULL REFERENCES leave_types(id),
    year         INTEGER NOT NULL,
    entitled     REAL NOT NULL DEFAULT 0,
    accrued      REAL NOT NULL DEFAULT 0,
    used         REAL NOT NULL DEFAULT 0,
    adjustment   REAL NOT NULL DEFAULT 0,
    UNIQUE(employee_id, leave_type_id, year)
  );

  CREATE TABLE IF NOT EXISTS leave_requests (
    id           TEXT PRIMARY KEY,
    employee_id  INTEGER NOT NULL REFERENCES employees(id),
    leave_type_id INTEGER NOT NULL REFERENCES leave_types(id),
    from_date    TEXT NOT NULL,
    to_date      TEXT NOT NULL,
    days         REAL NOT NULL,
    reason       TEXT,
    doc_url      TEXT,
    status       TEXT NOT NULL DEFAULT 'pending_manager',
    created_at   TEXT DEFAULT (datetime('now')),
    updated_at   TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS leave_approvals (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    request_id   TEXT NOT NULL REFERENCES leave_requests(id),
    level        INTEGER NOT NULL,
    role         TEXT NOT NULL,
    approver_id  INTEGER REFERENCES employees(id),
    approver_name TEXT,
    status       TEXT NOT NULL DEFAULT 'waiting',
    comment      TEXT,
    decided_at   TEXT,
    UNIQUE(request_id, level)
  );

  CREATE TABLE IF NOT EXISTS activity_log (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id  INTEGER REFERENCES employees(id),
    entra_id     TEXT,
    event_type   TEXT NOT NULL,
    event_time   TEXT NOT NULL,
    ip_address   TEXT,
    location     TEXT,
    raw          TEXT,
    synced_at    TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS sessions (
    token        TEXT PRIMARY KEY,
    employee_id  INTEGER NOT NULL REFERENCES employees(id),
    created_at   TEXT DEFAULT (datetime('now')),
    expires_at   TEXT NOT NULL
  );

  CREATE TABLE IF NOT EXISTS audit_log (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    actor_id     INTEGER REFERENCES employees(id),
    actor_name   TEXT,
    action       TEXT NOT NULL,
    target       TEXT,
    detail       TEXT,
    ts           TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS entra_config (
    id           INTEGER PRIMARY KEY DEFAULT 1,
    tenant_id    TEXT,
    client_id    TEXT,
    client_secret TEXT,
    redirect_uri TEXT,
    sync_interval_min INTEGER DEFAULT 5,
    idle_threshold_min INTEGER DEFAULT 15,
    auto_add_users INTEGER DEFAULT 1,
    default_cal_id TEXT,
    last_sync_at TEXT,
    last_sync_status TEXT,
    last_sync_msg TEXT,
    updated_at   TEXT DEFAULT (datetime('now'))
  );
`);

// ── Seed defaults if empty ────────────────────────────────────────────────────
function seedIfEmpty() {
  const hasCalendars = db.prepare('SELECT COUNT(*) as c FROM calendars').get().c;
  if (hasCalendars > 0) return;

  // Calendars
  const insertCal = db.prepare(`INSERT INTO calendars (id,name,timezone,weekends,color,holidays) VALUES (?,?,?,?,?,?)`);
  insertCal.run('cal-ae','UAE Standard','Asia/Dubai (UTC+4)','Fri-Sat','#0078d4',JSON.stringify(['Jan 1','Mar 30','Jun 25','Dec 2']));
  insertCal.run('cal-uk','UK Office','Europe/London (UTC+0)','Sat-Sun','#0d5c30',JSON.stringify(['Jan 1','Apr 18','May 5','Dec 25']));
  insertCal.run('cal-intl','Remote / International','UTC','Sat-Sun','#7a4a00',JSON.stringify(['Jan 1','Dec 25']));

  // Departments
  const insertDept = db.prepare(`INSERT INTO departments (name,teams,cal_id) VALUES (?,?,?)`);
  insertDept.run('Engineering', JSON.stringify(['Backend','Frontend','DevOps']), 'cal-ae');
  insertDept.run('Design',      JSON.stringify(['UX','Brand']),                  'cal-ae');
  insertDept.run('Product',     JSON.stringify(['Core','Growth']),               'cal-uk');
  insertDept.run('Marketing',   JSON.stringify(['Growth','Content']),            'cal-uk');
  insertDept.run('Finance',     JSON.stringify(['Accounts','FP&A']),             'cal-intl');
  insertDept.run('HR',          JSON.stringify(['Ops','Talent']),                'cal-ae');

  // Leave types
  const insertLT = db.prepare(`INSERT INTO leave_types (name,days_per_year,approval_levels,carry_over,carry_over_max,doc_required,color) VALUES (?,?,?,?,?,?,?)`);
  insertLT.run('Annual Leave',          24, 2, 'yes', 5, 0, '#0078d4');
  insertLT.run('Sick Leave',            14, 2, 'no',  0, 1, '#22c55e');
  insertLT.run('Emergency Leave',        5, 2, 'no',  0, 0, '#ef4444');
  insertLT.run('Maternity/Paternity',   90, 3, 'no',  0, 1, '#a855f7');
  insertLT.run('Unpaid Leave',          30, 3, 'no',  0, 0, '#9ca3af');
  insertLT.run('Study / Exam Leave',     7, 2, 'no',  0, 1, '#f59e0b');

  // Demo employees
  function hash(p) { return crypto.createHash('sha256').update(p).digest('hex'); }
  const insertEmp = db.prepare(`INSERT INTO employees (entra_id,name,email,initials,dept,team,cal_id,role,title,bal_type,password_hash,av_bg,av_color) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`);
  let _empSeq = 1;
  const empId = (name, email, ini, dept, team, cal, role, title, balType, pass, bg, fg) =>
    insertEmp.run(`demo-${role}-${_empSeq++}`, name, email, ini, dept, team, cal, role, title, balType, hash(pass), bg, fg).lastInsertRowid;

  const eId = empId('Alex Johnson','alex@company.com','AJ','Engineering','Backend','cal-ae','employee','Software Engineer','accrual','employee123','#e6f2fb','#004e8c');
  const mId = empId('Maria Chen',  'maria@company.com','MC','Engineering','Backend','cal-ae','manager', 'Team Lead',         'fixed',   'manager123', '#e6f9f0','#0d5c30');
  empId('James Okafor','james@company.com','JO','Design',    'UX',     'cal-ae',  'employee','UX Designer',        'accrual','emp456','#fef3e6','#7a4a00');
  empId('Priya Singh', 'priya@company.com','PS','Product',   'Core',   'cal-uk',  'employee','Product Manager',    'fixed',  'emp789','#f0e9fb','#4a1a8a');
  empId('Luis Rivera', 'luis@company.com', 'LR','Marketing', 'Growth', 'cal-uk',  'employee','Marketing Lead',     'fixed',  'emp000','#e6f2fb','#004e8c');
  const hId = empId('Sarah Ahmed',  'sarah@company.com','SA','HR',        'Ops',    'cal-ae',  'hr',     'HR Administrator',    'fixed',  'hr123',  '#fef3e6','#7a4a00');
  empId('Tom Wei',     'tom@company.com',  'TW','Finance',   'Accounts','cal-intl','employee','Finance Analyst',    'accrual','emp111','#fef3e6','#7a4a00');
  const dId = empId('David Park',   'david@company.com','DP','Executive', '',       'cal-ae',  'director','C-Level Director',   'fixed',  'director123','#f0e9fb','#4a1a8a');

  // Leave balances for current year
  const year = new Date().getFullYear();
  const ltRows = db.prepare('SELECT id FROM leave_types').all();
  const empRows = db.prepare('SELECT id, bal_type FROM employees').all();
  const insertBal = db.prepare(`INSERT OR IGNORE INTO leave_balances (employee_id,leave_type_id,year,entitled,accrued,used) VALUES (?,?,?,?,?,?)`);
  const monthsElapsed = new Date().getMonth(); // 0-based
  empRows.forEach(emp => {
    ltRows.forEach(lt => {
      const ltData = db.prepare('SELECT days_per_year FROM leave_types WHERE id=?').get(lt.id);
      const entitled = ltData.days_per_year;
      const accrued  = emp.bal_type === 'accrual' ? (monthsElapsed * 2) : entitled;
      insertBal.run(emp.id, lt.id, year, entitled, accrued, 0);
    });
  });

  // Sample leave requests
  const insertReq = db.prepare(`INSERT INTO leave_requests (id,employee_id,leave_type_id,from_date,to_date,days,reason,status) VALUES (?,?,?,?,?,?,?,?)`);
  const insertApproval = db.prepare(`INSERT INTO leave_approvals (request_id,level,role,approver_name,status) VALUES (?,?,?,?,?)`);

  insertReq.run('LR-001', eId, 1, '2025-05-10','2025-05-14', 5, 'Family vacation',   'pending_manager');
  insertApproval.run('LR-001',1,'Manager','Maria Chen','pending');
  insertApproval.run('LR-001',2,'HR Admin','Sarah Ahmed','waiting');
  insertApproval.run('LR-001',3,'Director','David Park','waiting');

  insertReq.run('LR-002', 3,   2, '2025-05-07','2025-05-08', 2, 'Medical appointment','pending_hr');
  insertApproval.run('LR-002',1,'Manager','Maria Chen','done');
  insertApproval.run('LR-002',2,'HR Admin','Sarah Ahmed','pending');

  insertReq.run('LR-003', 4,   1, '2025-05-20','2025-05-23', 4, 'Travel',            'approved');
  insertApproval.run('LR-003',1,'Manager','Maria Chen','done');
  insertApproval.run('LR-003',2,'HR Admin','Sarah Ahmed','done');
  insertApproval.run('LR-003',3,'Director','David Park','done');

  insertReq.run('LR-004', 5,   3, '2025-05-05','2025-05-06', 2, 'Family emergency',  'pending_director');
  insertApproval.run('LR-004',1,'Manager','Maria Chen','done');
  insertApproval.run('LR-004',2,'HR Admin','Sarah Ahmed','done');
  insertApproval.run('LR-004',3,'Director','David Park','pending');

  // Settings
  const setSetting = db.prepare(`INSERT OR REPLACE INTO settings (key,value) VALUES (?,?)`);
  setSetting.run('org_name','Acme Corporation');
  setSetting.run('default_timezone','Asia/Dubai (UTC+4)');
  setSetting.run('working_hours_day','8');
  setSetting.run('leave_year_reset','jan1');
  setSetting.run('min_notice_days','3');
  setSetting.run('max_consecutive_days','5');
  setSetting.run('allow_backdated','7');
  setSetting.run('notify_employee','1');
  setSetting.run('notify_manager','1');
  setSetting.run('monthly_balance_summary','1');
  setSetting.run('hr_daily_digest','0');
  setSetting.run('approval_reminder_days','2');
  setSetting.run('gdpr_mode','1');
  setSetting.run('data_retention_years','5');
  setSetting.run('leave_balance_policy','both');
  setSetting.run('fixed_annual_days','24');
  setSetting.run('fixed_carry_over_cap','5');
  setSetting.run('accrual_days_per_month','2');
  setSetting.run('accrual_max_cap','30');
  setSetting.run('accrual_credit_date','last_day');
  setSetting.run('approval_l1_role','line_manager');
  setSetting.run('approval_l1_sla','1');
  setSetting.run('approval_l1_escalate','hr');
  setSetting.run('approval_l2_role','hr_admin');
  setSetting.run('approval_l2_sla','1');
  setSetting.run('approval_l2_bypass','yes');
  setSetting.run('approval_l3_threshold_days','5');
  setSetting.run('approval_l3_role','director');
  setSetting.run('approval_l3_sla','2');

  // Entra config placeholder
  db.prepare(`INSERT OR IGNORE INTO entra_config (id) VALUES (1)`).run();

  // Sample activity log
  const insertAct = db.prepare(`INSERT INTO activity_log (employee_id,entra_id,event_type,event_time,ip_address) VALUES (?,?,?,?,?)`);
  const today = new Date().toISOString().slice(0,10);
  [[eId,'SignIn',`${today}T08:42:00Z','192.168.1.10`],[mId,'SignIn',`${today}T08:15:00Z`,'192.168.1.11'],[hId,'SignIn',`${today}T08:10:00Z`,'192.168.1.12'],[dId,'SignIn',`${today}T09:00:00Z`,'192.168.1.13']].forEach(([eid,type,time,ip])=>{
    try{ insertAct.run(eid,`demo-${eid}`,type,time,ip||'unknown'); } catch{}
  });

  // Audit log seed
  const insertAudit = db.prepare(`INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)`);
  insertAudit.run('System','Database initialised — seed data loaded','system');
}

seedIfEmpty();

// ── Query helpers ─────────────────────────────────────────────────────────────
const Q = {
  // Settings
  getSetting:  (key) => { const r = db.prepare('SELECT value FROM settings WHERE key=?').get(key); return r ? r.value : null; },
  getSettings: ()    => { const rows = db.prepare('SELECT key,value FROM settings').all(); const o={}; rows.forEach(r=>o[r.key]=r.value); return o; },
  setSetting:  (key,value,actorName='system') => {
    db.prepare('INSERT OR REPLACE INTO settings (key,value,updated_at) VALUES (?,?,datetime(\'now\'))').run(key, String(value));
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName,'Setting updated',key,String(value));
  },
  setSettings: (obj, actorName='system') => {
    const stmt = db.prepare('INSERT OR REPLACE INTO settings (key,value,updated_at) VALUES (?,?,datetime(\'now\'))');
    Object.entries(obj).forEach(([k,v]) => stmt.run(k, String(v)));
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName,'Bulk settings saved',Object.keys(obj).join(', '),JSON.stringify(obj));
  },

  // Calendars
  getCalendars: () => db.prepare('SELECT * FROM calendars ORDER BY name').all().map(parseCalendar),
  getCalendar:  (id) => { const r = db.prepare('SELECT * FROM calendars WHERE id=?').get(id); return r ? parseCalendar(r) : null; },
  saveCalendar: (data, actorName) => {
    if (data.id) {
      db.prepare('UPDATE calendars SET name=?,timezone=?,weekends=?,color=?,holidays=?,updated_at=datetime(\'now\') WHERE id=?')
        .run(data.name,data.timezone,data.weekends,data.color,JSON.stringify(data.holidays||[]),data.id);
      db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,`Calendar updated`,data.name);
      return data.id;
    } else {
      const id = 'cal-'+Date.now();
      db.prepare('INSERT INTO calendars (id,name,timezone,weekends,color,holidays) VALUES (?,?,?,?,?,?)')
        .run(id,data.name,data.timezone||'UTC',data.weekends||'Sat-Sun',data.color||'#0078d4',JSON.stringify(data.holidays||[]));
      db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,`Calendar created`,data.name);
      return id;
    }
  },
  deleteCalendar: (id, actorName) => {
    const cal = db.prepare('SELECT name FROM calendars WHERE id=?').get(id);
    db.prepare('DELETE FROM calendars WHERE id=?').run(id);
    db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Calendar deleted',cal?.name||id);
  },

  // Employees
  getEmployees: () => db.prepare('SELECT * FROM employees WHERE active=1 ORDER BY name').all(),
  getEmployee:  (id) => db.prepare('SELECT * FROM employees WHERE id=?').get(id),
  getEmployeeByEmail: (email) => db.prepare('SELECT * FROM employees WHERE email=?').get(email),
  saveEmployee: (data, actorName) => {
    if (data.id) {
      db.prepare('UPDATE employees SET name=?,email=?,dept=?,team=?,cal_id=?,role=?,title=?,bal_type=?,av_bg=?,av_color=?,updated_at=datetime(\'now\') WHERE id=?')
        .run(data.name,data.email,data.dept,data.team,data.cal_id,data.role||'employee',data.title,data.bal_type||'fixed',data.av_bg||'#e6f2fb',data.av_color||'#004e8c',data.id);
      db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Employee updated',data.name);
      return data.id;
    } else {
      const ini = (data.name||'XX').split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
      const r = db.prepare('INSERT INTO employees (name,email,initials,dept,team,cal_id,role,title,bal_type,av_bg,av_color) VALUES (?,?,?,?,?,?,?,?,?,?,?)')
        .run(data.name,data.email,ini,data.dept,data.team,data.cal_id,data.role||'employee',data.title,data.bal_type||'fixed',data.av_bg||'#e6f2fb',data.av_color||'#004e8c');
      db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Employee added',data.name);
      return r.lastInsertRowid;
    }
  },
  assignCalendar: (employeeId, calId, actorName) => {
    const emp = db.prepare('SELECT name FROM employees WHERE id=?').get(employeeId);
    const cal = db.prepare('SELECT name FROM calendars WHERE id=?').get(calId);
    db.prepare('UPDATE employees SET cal_id=?,updated_at=datetime(\'now\') WHERE id=?').run(calId,employeeId);
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName,`Calendar assigned`,emp?.name||employeeId,`→ ${cal?.name||calId}`);
  },
  assignCalendarByDept: (dept, calId, overrideIndividual, actorName) => {
    const count = db.prepare('SELECT COUNT(*) as c FROM employees WHERE dept=? AND active=1').get(dept).c;
    db.prepare('UPDATE employees SET cal_id=?,updated_at=datetime(\'now\') WHERE dept=? AND active=1').run(calId,dept);
    const cal = db.prepare('SELECT name FROM calendars WHERE id=?').get(calId);
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName,`Bulk calendar assignment`,dept,`${count} employees → ${cal?.name||calId}`);
    return count;
  },

  // Departments
  getDepartments: () => db.prepare('SELECT * FROM departments ORDER BY name').all().map(d=>({...d,teams:JSON.parse(d.teams||'[]')})),
  saveDepartment: (data, actorName) => {
    if (data.id) {
      db.prepare('UPDATE departments SET name=?,teams=?,cal_id=? WHERE id=?').run(data.name,JSON.stringify(data.teams||[]),data.cal_id,data.id);
    } else {
      db.prepare('INSERT INTO departments (name,teams,cal_id) VALUES (?,?,?)').run(data.name,JSON.stringify(data.teams||[]),data.cal_id);
    }
    db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Department saved',data.name);
  },

  // Leave types
  getLeaveTypes: () => db.prepare('SELECT * FROM leave_types WHERE active=1 ORDER BY name').all(),
  saveLeaveType: (data, actorName) => {
    if (data.id) {
      db.prepare('UPDATE leave_types SET name=?,days_per_year=?,approval_levels=?,carry_over=?,carry_over_max=?,doc_required=?,applies_to=?,color=?,updated_at=datetime(\'now\') WHERE id=?')
        .run(data.name,data.days_per_year,data.approval_levels,data.carry_over,data.carry_over_max||0,data.doc_required?1:0,data.applies_to||'all',data.color||'#0078d4',data.id);
      db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Leave type updated',data.name);
    } else {
      db.prepare('INSERT INTO leave_types (name,days_per_year,approval_levels,carry_over,carry_over_max,doc_required,applies_to,color) VALUES (?,?,?,?,?,?,?,?)')
        .run(data.name,data.days_per_year,data.approval_levels,data.carry_over,data.carry_over_max||0,data.doc_required?1:0,data.applies_to||'all',data.color||'#0078d4');
      db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Leave type created',data.name);
    }
  },
  deleteLeaveType: (id, actorName) => {
    const lt = db.prepare('SELECT name FROM leave_types WHERE id=?').get(id);
    db.prepare('UPDATE leave_types SET active=0 WHERE id=?').run(id);
    db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Leave type deactivated',lt?.name||id);
  },

  // Leave balances
  getBalances: (employeeId, year) => {
    year = year || new Date().getFullYear();
    return db.prepare(`SELECT lb.*, lt.name as type_name, lt.color, lt.days_per_year FROM leave_balances lb JOIN leave_types lt ON lb.leave_type_id=lt.id WHERE lb.employee_id=? AND lb.year=?`).all(employeeId, year);
  },
  getAllBalances: (year) => {
    year = year || new Date().getFullYear();
    return db.prepare(`SELECT lb.*, e.name as emp_name, e.initials, e.dept, e.av_bg, e.av_color, e.bal_type, lt.name as type_name FROM leave_balances lb JOIN employees e ON lb.employee_id=e.id JOIN leave_types lt ON lb.leave_type_id=lt.id WHERE lb.year=? AND lt.name='Annual Leave' ORDER BY e.name`).all(year);
  },
  adjustBalance: (employeeId, leaveTypeId, year, adjustment, actorName, reason) => {
    year = year || new Date().getFullYear();
    db.prepare('INSERT OR IGNORE INTO leave_balances (employee_id,leave_type_id,year,entitled,accrued,used,adjustment) VALUES (?,?,?,0,0,0,0)').run(employeeId,leaveTypeId,year);
    db.prepare('UPDATE leave_balances SET adjustment=adjustment+? WHERE employee_id=? AND leave_type_id=? AND year=?').run(adjustment,employeeId,leaveTypeId,year);
    const emp = db.prepare('SELECT name FROM employees WHERE id=?').get(employeeId);
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName,`Balance adjusted: ${adjustment>0?'+':''}${adjustment}d`,emp?.name||employeeId,reason||'Manual adjustment');
  },
  runMonthlyAccrual: (actorName) => {
    const year = new Date().getFullYear();
    const accrualRate = parseFloat(Q.getSetting('accrual_days_per_month') || '2');
    const maxCap = parseFloat(Q.getSetting('accrual_max_cap') || '30');
    const annualLT = db.prepare("SELECT id FROM leave_types WHERE name='Annual Leave'").get();
    if (!annualLT) return 0;
    const empAccrual = db.prepare("SELECT id FROM employees WHERE bal_type='accrual' AND active=1").all();
    empAccrual.forEach(emp => {
      db.prepare('INSERT OR IGNORE INTO leave_balances (employee_id,leave_type_id,year,entitled,accrued,used,adjustment) VALUES (?,?,?,0,0,0,0)').run(emp.id,annualLT.id,year);
      db.prepare('UPDATE leave_balances SET accrued=MIN(accrued+?,?) WHERE employee_id=? AND leave_type_id=? AND year=?').run(accrualRate,maxCap,emp.id,annualLT.id,year);
    });
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName||'System',`Monthly accrual run`,`${empAccrual.length} employees`,`+${accrualRate} days each`);
    return empAccrual.length;
  },

  // Leave requests
  getLeaveRequests: (filters={}) => {
    let sql = `SELECT lr.*, e.name as emp_name, e.dept, e.initials, e.av_bg, e.av_color, lt.name as type_name FROM leave_requests lr JOIN employees e ON lr.employee_id=e.id JOIN leave_types lt ON lr.leave_type_id=lt.id WHERE 1=1`;
    const params = [];
    if (filters.employeeId) { sql += ' AND lr.employee_id=?'; params.push(filters.employeeId); }
    if (filters.status)     { sql += ' AND lr.status=?';      params.push(filters.status); }
    if (filters.dept)       { sql += ' AND e.dept=?';         params.push(filters.dept); }
    sql += ' ORDER BY lr.created_at DESC';
    const rows = db.prepare(sql).all(...params);
    return rows.map(r => ({
      ...r,
      approvals: db.prepare('SELECT * FROM leave_approvals WHERE request_id=? ORDER BY level').all(r.id)
    }));
  },
  submitLeaveRequest: (empId, data, actorName) => {
    const id = 'LR-' + String(Date.now()).slice(-6);
    const lt = db.prepare('SELECT * FROM leave_types WHERE id=?').get(data.leave_type_id);
    db.prepare('INSERT INTO leave_requests (id,employee_id,leave_type_id,from_date,to_date,days,reason,status) VALUES (?,?,?,?,?,?,?,?)')
      .run(id, empId, data.leave_type_id, data.from_date, data.to_date, data.days, data.reason||'', 'pending_manager');
    // Build approval chain
    const levels = data.days > (parseFloat(Q.getSetting('approval_l3_threshold_days')||'5')) || (lt && lt.approval_levels >= 3) ? 3 : 2;
    const managerName = db.prepare("SELECT name FROM employees WHERE dept=(SELECT dept FROM employees WHERE id=?) AND role='manager' LIMIT 1").get(empId)?.name || 'Manager';
    const hrName      = db.prepare("SELECT name FROM employees WHERE role='hr' LIMIT 1").get()?.name || 'HR Admin';
    const dirName     = db.prepare("SELECT name FROM employees WHERE role='director' LIMIT 1").get()?.name || 'Director';
    db.prepare('INSERT INTO leave_approvals (request_id,level,role,approver_name,status) VALUES (?,?,?,?,?)').run(id,1,'Manager',managerName,'pending');
    db.prepare('INSERT INTO leave_approvals (request_id,level,role,approver_name,status) VALUES (?,?,?,?,?)').run(id,2,'HR Admin',hrName,'waiting');
    if (levels === 3) db.prepare('INSERT INTO leave_approvals (request_id,level,role,approver_name,status) VALUES (?,?,?,?,?)').run(id,3,'Director',dirName,'waiting');
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName,'Leave request submitted',id,`${data.days} days from ${data.from_date}`);
    return id;
  },
  approveLeaveRequest: (requestId, approverEmpId, comment, actorName) => {
    const req = db.prepare('SELECT * FROM leave_requests WHERE id=?').get(requestId);
    if (!req) throw new Error('Request not found');
    const pending = db.prepare("SELECT * FROM leave_approvals WHERE request_id=? AND status='pending' ORDER BY level LIMIT 1").get(requestId);
    if (!pending) throw new Error('No pending approval step');
    db.prepare("UPDATE leave_approvals SET status='done',approver_id=?,comment=?,decided_at=datetime('now') WHERE id=?").run(approverEmpId,comment||null,pending.id);
    // Check next step
    const nextWaiting = db.prepare("SELECT * FROM leave_approvals WHERE request_id=? AND status='waiting' ORDER BY level LIMIT 1").get(requestId);
    if (nextWaiting) {
      db.prepare("UPDATE leave_approvals SET status='pending' WHERE id=?").run(nextWaiting.id);
      db.prepare("UPDATE leave_requests SET status=?,updated_at=datetime('now') WHERE id=?").run(`pending_${nextWaiting.role.toLowerCase().replace(/\s+/g,'_')}`,requestId);
    } else {
      db.prepare("UPDATE leave_requests SET status='approved',updated_at=datetime('now') WHERE id=?").run(requestId);
      // Deduct balance
      const year = new Date(req.from_date).getFullYear();
      db.prepare('UPDATE leave_balances SET used=used+? WHERE employee_id=? AND leave_type_id=? AND year=?').run(req.days,req.employee_id,req.leave_type_id,year);
    }
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName,'Leave approved',requestId,comment||'');
  },
  rejectLeaveRequest: (requestId, approverEmpId, reason, actorName) => {
    db.prepare("UPDATE leave_approvals SET status='rejected',approver_id=?,comment=?,decided_at=datetime('now') WHERE request_id=? AND status='pending'").run(approverEmpId,reason||null,requestId);
    db.prepare("UPDATE leave_approvals SET status='rejected' WHERE request_id=? AND status='waiting'").run(requestId);
    db.prepare("UPDATE leave_requests SET status='rejected',updated_at=datetime('now') WHERE id=?").run(requestId);
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(actorName,'Leave rejected',requestId,reason||'');
  },

  // Activity log
  getActivityLog: (filters={}) => {
    let sql = `SELECT al.*, e.name as emp_name, e.initials, e.dept, e.av_bg, e.av_color FROM activity_log al LEFT JOIN employees e ON al.employee_id=e.id WHERE 1=1`;
    const params = [];
    if (filters.employeeId) { sql += ' AND al.employee_id=?'; params.push(filters.employeeId); }
    if (filters.date)       { sql += ' AND date(al.event_time)=?'; params.push(filters.date); }
    sql += ' ORDER BY al.event_time DESC LIMIT 200';
    return db.prepare(sql).all(...params);
  },
  logActivity: (empId, entraId, type, time, ip, raw) => {
    db.prepare('INSERT INTO activity_log (employee_id,entra_id,event_type,event_time,ip_address,raw) VALUES (?,?,?,?,?,?)').run(empId||null,entraId||null,type,time,ip||null,raw?JSON.stringify(raw):null);
  },
  getTodaySummary: () => {
    const today = new Date().toISOString().slice(0,10);
    return db.prepare(`SELECT e.id, e.name, e.initials, e.dept, e.av_bg, e.av_color,
      MIN(CASE WHEN al.event_type='SignIn' THEN al.event_time END) as first_login,
      MAX(CASE WHEN al.event_type='SignOut' THEN al.event_time END) as last_logout
      FROM employees e LEFT JOIN activity_log al ON e.id=al.employee_id AND date(al.event_time)=?
      WHERE e.active=1 GROUP BY e.id ORDER BY e.name`).all(today);
  },

  // Sessions
  createSession: (empId) => {
    const token = crypto.randomBytes(32).toString('hex');
    const expiresAt = new Date(Date.now() + 8*60*60*1000).toISOString();
    db.prepare('DELETE FROM sessions WHERE employee_id=?').run(empId);
    db.prepare('INSERT INTO sessions (token,employee_id,expires_at) VALUES (?,?,?)').run(token,empId,expiresAt);
    return token;
  },
  getSession: (token) => {
    const s = db.prepare("SELECT * FROM sessions WHERE token=? AND expires_at > datetime('now')").get(token);
    if (!s) return null;
    return db.prepare('SELECT * FROM employees WHERE id=?').get(s.employee_id);
  },
  deleteSession: (token) => db.prepare('DELETE FROM sessions WHERE token=?').run(token),

  // Auth
  verifyPassword: (emp, password) => {
    const hash = crypto.createHash('sha256').update(password).digest('hex');
    return emp.password_hash === hash;
  },
  setPassword: (empId, password, actorName) => {
    const hash = crypto.createHash('sha256').update(password).digest('hex');
    db.prepare('UPDATE employees SET password_hash=? WHERE id=?').run(hash,empId);
    db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Password changed',String(empId));
  },

  // Entra config
  getEntraConfig: () => db.prepare('SELECT * FROM entra_config WHERE id=1').get() || {},
  saveEntraConfig: (data, actorName) => {
    db.prepare('INSERT OR IGNORE INTO entra_config (id) VALUES (1)').run();
    db.prepare('UPDATE entra_config SET tenant_id=?,client_id=?,client_secret=?,redirect_uri=?,sync_interval_min=?,idle_threshold_min=?,auto_add_users=?,default_cal_id=?,updated_at=datetime(\'now\') WHERE id=1')
      .run(data.tenant_id,data.client_id,data.client_secret,data.redirect_uri,data.sync_interval_min||5,data.idle_threshold_min||15,data.auto_add_users?1:0,data.default_cal_id||null);
    db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(actorName,'Entra ID config updated','entra_config');
  },

  // Audit log
  getAuditLog: (limit=100) => db.prepare('SELECT * FROM audit_log ORDER BY ts DESC LIMIT ?').all(limit),
};

function parseCalendar(r) {
  return { ...r, holidays: typeof r.holidays === 'string' ? JSON.parse(r.holidays) : (r.holidays||[]) };
}

module.exports = { db, Q };
