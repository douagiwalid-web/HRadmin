'use strict';
const http  = require('http');
const https = require('https');
const fs    = require('fs');
const path  = require('path');
const url   = require('url');
const crypto = require('crypto');
const { Q, db } = require('./db');
const entra = require('./entra');

const PORT = process.env.PORT || 3000;

// ── SSE clients ───────────────────────────────────────────────────────────────
const sseClients = new Set();
function broadcast(data) {
  const payload = `data: ${JSON.stringify(data)}\n\n`;
  sseClients.forEach(res => { try { res.write(payload); } catch {} });
}

entra.startSyncScheduler(broadcast);

// ── Helpers ───────────────────────────────────────────────────────────────────
function json(res, status, data) {
  const body = JSON.stringify(data);
  res.writeHead(status, { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) });
  res.end(body);
}
function readBody(req) {
  return new Promise(resolve => {
    let b = '';
    req.on('data', d => b += d);
    req.on('end', () => { try { resolve(JSON.parse(b)); } catch { resolve({}); } });
  });
}
function getCookieToken(req) {
  const c = (req.headers.cookie || '').split(';').map(s => s.trim()).find(s => s.startsWith('wiq_session='));
  return c ? c.split('=')[1] : null;
}
function requireAuth(req, res) {
  const token = getCookieToken(req);
  if (!token) { json(res, 401, { error: 'Not authenticated' }); return null; }
  const emp = Q.getSession(token);
  if (!emp) { json(res, 401, { error: 'Session expired' }); return null; }
  return emp;
}
function requireRole(emp, res, ...roles) {
  if (!roles.includes(emp.role)) { json(res, 403, { error: 'Forbidden' }); return false; }
  return true;
}

// ── PKCE verifier store (state → {verifier, redirectUri}, TTL 10 min) ─────────
const pkceStore = new Map();
setInterval(() => {
  const now = Date.now();
  pkceStore.forEach((v, k) => { if (now - v.ts > 10 * 60 * 1000) pkceStore.delete(k); });
}, 60 * 1000);

function generatePKCE() {
  const verifier  = crypto.randomBytes(32).toString('base64url');
  const challenge = crypto.createHash('sha256').update(verifier).digest('base64url');
  return { verifier, challenge };
}

// ── Microsoft SSO: exchange auth code for user info ───────────────────────────
async function exchangeMsftCode(code, redirectUri, cfg, codeVerifier) {
  // Confidential client (has client_secret) + PKCE:
  // Azure requires BOTH client_secret AND code_verifier together.
  // Public client (no secret): send code_verifier only.
  const params = {
    grant_type:    'authorization_code',
    client_id:     cfg.client_id,
    code,
    redirect_uri:  redirectUri,
    scope:         'openid profile email User.Read',
  };
  if (codeVerifier)      params.code_verifier = codeVerifier;
  if (cfg.client_secret) params.client_secret  = cfg.client_secret;
  const body = new URLSearchParams(params).toString();
  return new Promise((resolve, reject) => {
    const opts = {
      hostname: 'login.microsoftonline.com',
      path: `/${cfg.tenant_id}/oauth2/v2.0/token`,
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    };
    let resp = '';
    const r = https.request(opts, res => { res.on('data', d => resp += d); res.on('end', () => { try { resolve(JSON.parse(resp)); } catch { reject(new Error('Token parse error')); } }); });
    r.on('error', reject);
    r.write(body); r.end();
  });
}

async function getMsftUserInfo(accessToken) {
  return new Promise((resolve, reject) => {
    const opts = {
      hostname: 'graph.microsoft.com',
      path: '/v1.0/me?$select=id,displayName,mail,userPrincipalName,department,jobTitle',
      headers: { Authorization: `Bearer ${accessToken}` },
    };
    let resp = '';
    const r = https.request(opts, res => { res.on('data', d => resp += d); res.on('end', () => { try { resolve(JSON.parse(resp)); } catch { reject(new Error('User info parse error')); } }); });
    r.on('error', reject);
    r.end();
  });
}

const MIME = { '.html':'text/html','.css':'text/css','.js':'application/javascript','.json':'application/json','.ico':'image/x-icon','.png':'image/png' };

// ── Router ────────────────────────────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  const parsed   = url.parse(req.url, true);
  const pathname = parsed.pathname;
  const method   = req.method;

  // ── Password login (local accounts) ────────────────────────────────────────
  if (pathname === '/api/login' && method === 'POST') {
    const body = await readBody(req);
    const emp = Q.getEmployeeByEmail(body.username) ||
                db.prepare("SELECT * FROM employees WHERE name=? OR LOWER(role)=LOWER(?) LIMIT 1").get(body.username, body.username);
    if (!emp || !Q.verifyPassword(emp, body.password)) return json(res, 401, { error: 'Invalid credentials' });
    if (!emp.active) return json(res, 403, { error: 'Account disabled' });
    const token = Q.createSession(emp.id);
    res.writeHead(200, { 'Set-Cookie': `wiq_session=${token}; Path=/; HttpOnly; SameSite=Lax`, 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ ok: true, user: safeUser(emp) }));
  }

  // ── Microsoft SSO: get auth URL ─────────────────────────────────────────────
  if (pathname === '/api/sso/url' && method === 'GET') {
    const cfg = Q.getEntraConfig();
    if (!cfg.tenant_id || !cfg.client_id) return json(res, 400, { error: 'Entra ID not configured — enter Tenant ID and Client ID in Entra config first' });
    const redirectUri = cfg.redirect_uri || `http://${req.headers.host}/auth/callback`;
    const state = crypto.randomBytes(16).toString('hex');
    const { verifier, challenge } = generatePKCE();
    // Store verifier keyed by state — callback retrieves it in a moment
    pkceStore.set(state, { verifier, redirectUri, ts: Date.now() });
    const authUrl = 'https://login.microsoftonline.com/' + cfg.tenant_id + '/oauth2/v2.0/authorize?' + new URLSearchParams({
      client_id:             cfg.client_id,
      response_type:         'code',
      redirect_uri:          redirectUri,
      scope:                 'openid profile email User.Read',
      state,
      prompt:                'select_account',
      code_challenge:        challenge,
      code_challenge_method: 'S256',
    });
    return json(res, 200, { url: authUrl, state });
  }

  // ── Microsoft SSO: callback — exchange code for session ─────────────────────
  if (pathname === '/auth/callback') {
    const code  = parsed.query.code;
    const state = parsed.query.state;
    const error = parsed.query.error;
    if (error) {
      res.writeHead(302, { Location: `/?sso_error=${encodeURIComponent(parsed.query.error_description||error)}` });
      return res.end();
    }
    const cfg = Q.getEntraConfig();
    // Retrieve PKCE verifier stored when auth URL was generated
    const pkce = pkceStore.get(state);
    pkceStore.delete(state); // one-time use
    const redirectUri = pkce?.redirectUri || cfg.redirect_uri || `http://${req.headers.host}/auth/callback`;
    const codeVerifier = pkce?.verifier || null;
    try {
      const tokens = await exchangeMsftCode(code, redirectUri, cfg, codeVerifier);
      if (tokens.error) throw new Error(tokens.error_description || tokens.error);
      const msUser = await getMsftUserInfo(tokens.access_token);
      const email  = msUser.mail || msUser.userPrincipalName;
      if (!email) throw new Error('No email returned from Microsoft');

      // Find or create employee
      let emp = Q.getEmployeeByEmail(email);
      if (!emp) {
        const ini = (msUser.displayName||'XX').split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
        const r = db.prepare('INSERT INTO employees (entra_id,name,email,initials,dept,title,cal_id,role,bal_type,av_bg,av_color) VALUES (?,?,?,?,?,?,?,?,?,?,?)')
          .run(msUser.id, msUser.displayName||email, email, ini, msUser.department||'', msUser.jobTitle||'',
               cfg.default_cal_id||'cal-ae', 'employee', 'fixed', '#e6f2fb', '#004e8c');
        const year = new Date().getFullYear();
        const ltRows = db.prepare('SELECT id, days_per_year FROM leave_types WHERE active=1').all();
        ltRows.forEach(lt => {
          db.prepare('INSERT OR IGNORE INTO leave_balances (employee_id,leave_type_id,year,entitled,accrued,used,adjustment) VALUES (?,?,?,?,?,0,0)')
            .run(r.lastInsertRowid, lt.id, year, lt.days_per_year, lt.days_per_year);
        });
        emp = Q.getEmployee(r.lastInsertRowid);
        db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run('System', 'SSO user created', email);
      } else {
        db.prepare("UPDATE employees SET entra_id=?,name=?,dept=?,title=?,updated_at=datetime('now') WHERE id=?")
          .run(msUser.id, msUser.displayName||emp.name, msUser.department||emp.dept, msUser.jobTitle||emp.title, emp.id);
        emp = Q.getEmployee(emp.id);
      }
      if (!emp.active) throw new Error('Account disabled');
      const sessionToken = Q.createSession(emp.id);
      res.writeHead(302, { 'Set-Cookie': `wiq_session=${sessionToken}; Path=/; HttpOnly; SameSite=Lax`, Location: '/' });
      return res.end();
    } catch (e) {
      res.writeHead(302, { Location: `/?sso_error=${encodeURIComponent(e.message)}` });
      return res.end();
    }
  }

  if (pathname === '/api/logout' && method === 'POST') {
    const token = getCookieToken(req);
    if (token) Q.deleteSession(token);
    res.writeHead(200, { 'Set-Cookie': 'wiq_session=; Path=/; Max-Age=0', 'Content-Type': 'application/json' });
    return res.end(JSON.stringify({ ok: true }));
  }

  if (pathname === '/api/me') {
    const emp = requireAuth(req, res);
    if (!emp) return;
    return json(res, 200, safeUser(emp));
  }

  // ── Settings ────────────────────────────────────────────────────────────────
  if (pathname === '/api/settings') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'GET') return json(res, 200, Q.getSettings());
    if (method === 'POST') {
      if (!requireRole(emp, res, 'hr', 'director')) return;
      const body = await readBody(req);
      Q.setSettings(body, emp.name);
      return json(res, 200, { ok: true });
    }
  }

  // ── Email settings (stored as regular settings keys) ─────────────────────────
  if (pathname === '/api/email-settings') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    if (method === 'GET') {
      const all = Q.getSettings();
      return json(res, 200, {
        smtp_host:     all.smtp_host     || '',
        smtp_port:     all.smtp_port     || '587',
        smtp_secure:   all.smtp_secure   || '0',
        smtp_user:     all.smtp_user     || '',
        smtp_pass:     all.smtp_pass     ? '••••••••' : '',
        smtp_from:     all.smtp_from     || '',
        smtp_from_name: all.smtp_from_name || 'WorkIQ HR',
        email_enabled: all.email_enabled || '0',
      });
    }
    if (method === 'POST') {
      const body = await readBody(req);
      const toSave = { ...body };
      // Don't overwrite password if masked
      if (toSave.smtp_pass === '••••••••') delete toSave.smtp_pass;
      Q.setSettings(toSave, emp.name);
      return json(res, 200, { ok: true });
    }
  }

  // ── Test email ────────────────────────────────────────────────────────────────
  if (pathname === '/api/email-settings/test' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    const s = Q.getSettings();
    if (!s.smtp_host) return json(res, 400, { error: 'SMTP not configured' });
    // We don't have nodemailer (no npm), so return config validation result
    const missing = [];
    if (!s.smtp_host) missing.push('SMTP host');
    if (!s.smtp_user) missing.push('SMTP username');
    if (!s.smtp_pass) missing.push('SMTP password');
    if (!s.smtp_from) missing.push('From address');
    if (missing.length) return json(res, 400, { error: `Missing: ${missing.join(', ')}` });
    db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(emp.name, 'Email test triggered', s.smtp_host, `to ${emp.email}`);
    return json(res, 200, { ok: true, message: `Config validated. To send real emails, install nodemailer: npm install nodemailer` });
  }

  // ── Employee calendar info (for holiday-aware day count) ──────────────────────
  if (/^\/api\/my-calendar-info$/.test(pathname)) {
    const emp = requireAuth(req, res); if (!emp) return;
    const empRecord = db.prepare('SELECT cal_id FROM employees WHERE id=?').get(emp.id);
    const cal = empRecord?.cal_id ? Q.getCalendar(empRecord.cal_id) : null;
    const leaveRows = db.prepare(
      "SELECT from_date, to_date FROM leave_requests WHERE employee_id=? AND status='approved'"
    ).all(emp.id);
    return json(res, 200, { calendar: cal, approvedLeave: leaveRows });
  }

  // ── Calendars ────────────────────────────────────────────────────────────────
  if (pathname === '/api/calendars') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'GET') return json(res, 200, Q.getCalendars());
    if (method === 'POST') {
      if (!requireRole(emp, res, 'hr')) return;
      const body = await readBody(req);
      const id = Q.saveCalendar(body, emp.name);
      return json(res, 200, { ok: true, id });
    }
  }
  if (/^\/api\/calendars\/([^/]+)$/.test(pathname)) {
    const calId = pathname.split('/')[3];
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'PUT') {
      if (!requireRole(emp, res, 'hr')) return;
      const body = await readBody(req);
      body.id = calId;
      Q.saveCalendar(body, emp.name);
      return json(res, 200, { ok: true });
    }
    if (method === 'DELETE') {
      if (!requireRole(emp, res, 'hr')) return;
      Q.deleteCalendar(calId, emp.name);
      return json(res, 200, { ok: true });
    }
  }

  // ── Employees ────────────────────────────────────────────────────────────────
  if (pathname === '/api/employees') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'GET') {
      const emps = Q.getEmployees();
      return json(res, 200, emps.map(safeUser));
    }
    if (method === 'POST') {
      if (!requireRole(emp, res, 'hr')) return;
      const body = await readBody(req);
      const id = Q.saveEmployee(body, emp.name);
      return json(res, 200, { ok: true, id });
    }
  }
  if (/^\/api\/employees\/(\d+)$/.test(pathname)) {
    const empId = parseInt(pathname.split('/')[3]);
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'GET') return json(res, 200, safeUser(Q.getEmployee(empId)));
    if (method === 'PUT') {
      if (!requireRole(emp, res, 'hr')) return;
      const body = await readBody(req); body.id = empId;
      Q.saveEmployee(body, emp.name);
      return json(res, 200, { ok: true });
    }
    if (method === 'DELETE') {
      if (!requireRole(emp, res, 'hr')) return;
      Q.deleteEmployee(empId, emp.name);
      return json(res, 200, { ok: true });
    }
  }
  if (/^\/api\/employees\/(\d+)\/calendar$/.test(pathname)) {
    const empId = parseInt(pathname.split('/')[3]);
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    const body = await readBody(req);
    Q.assignCalendar(empId, body.cal_id, emp.name);
    return json(res, 200, { ok: true });
  }
  if (pathname === '/api/employees/bulk-calendar' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    const body = await readBody(req);
    let count = 0;
    if (body.team)      count = Q.assignCalendarByTeam(body.team, body.cal_id, emp.name);
    else if (body.dept) count = Q.assignCalendarByDept(body.dept, body.cal_id, body.override, emp.name);
    return json(res, 200, { ok: true, count });
  }

  // ── Departments ───────────────────────────────────────────────────────────────
  if (pathname === '/api/departments') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'GET') return json(res, 200, Q.getDepartments());
    if (method === 'POST') {
      if (!requireRole(emp, res, 'hr')) return;
      const body = await readBody(req);
      Q.saveDepartment(body, emp.name);
      return json(res, 200, { ok: true });
    }
  }
  if (/^\/api\/departments\/(\d+)$/.test(pathname)) {
    const deptId = parseInt(pathname.split('/')[3]);
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'DELETE') {
      if (!requireRole(emp, res, 'hr')) return;
      db.prepare('DELETE FROM departments WHERE id=?').run(deptId);
      db.prepare('INSERT INTO audit_log (actor_name,action,target) VALUES (?,?,?)').run(emp.name, 'Department deleted', String(deptId));
      return json(res, 200, { ok: true });
    }
  }

  // ── Leave types ───────────────────────────────────────────────────────────────
  if (pathname === '/api/leave-types') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'GET') return json(res, 200, Q.getLeaveTypes());
    if (method === 'POST') {
      if (!requireRole(emp, res, 'hr')) return;
      const body = await readBody(req);
      Q.saveLeaveType(body, emp.name);
      return json(res, 200, { ok: true });
    }
  }
  if (/^\/api\/leave-types\/(\d+)$/.test(pathname)) {
    const ltId = parseInt(pathname.split('/')[3]);
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    if (method === 'PUT') {
      const body = await readBody(req); body.id = ltId;
      Q.saveLeaveType(body, emp.name);
      return json(res, 200, { ok: true });
    }
    if (method === 'DELETE') {
      Q.deleteLeaveType(ltId, emp.name);
      return json(res, 200, { ok: true });
    }
  }

  // ── Leave balances ────────────────────────────────────────────────────────────
  if (pathname === '/api/leave-balances') {
    const emp = requireAuth(req, res); if (!emp) return;
    const year = parseInt(parsed.query.year) || new Date().getFullYear();
    if (['hr','director'].includes(emp.role)) return json(res, 200, Q.getAllBalances(year));
    return json(res, 200, Q.getBalances(emp.id, year));
  }
  if (pathname === '/api/leave-balances/adjust' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    const body = await readBody(req);
    Q.adjustBalance(body.employee_id, body.leave_type_id, body.year, body.adjustment, emp.name, body.reason);
    return json(res, 200, { ok: true });
  }
  if (pathname === '/api/leave-balances/accrue' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    const count = Q.runMonthlyAccrual(emp.name);
    return json(res, 200, { ok: true, count });
  }

  // ── Leave requests ────────────────────────────────────────────────────────────
  if (pathname === '/api/leave-requests') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (method === 'GET') {
      const filters = {};
      // Employee: own leaves only
      if (emp.role === 'employee') filters.employeeId = emp.id;
      // Manager: their dept + optional status filter
      if (emp.role === 'manager') {
        const mgr = db.prepare('SELECT dept FROM employees WHERE id=?').get(emp.id);
        if (mgr?.dept) filters.managerDept = mgr.dept;
      }
      if (parsed.query.status) filters.status = parsed.query.status;
      if (parsed.query.dept)   filters.dept   = parsed.query.dept;
      if (parsed.query.team)   filters.team   = parsed.query.team;
      return json(res, 200, Q.getLeaveRequests(filters));
    }
    if (method === 'POST') {
      const body = await readBody(req);
      const id = Q.submitLeaveRequest(emp.id, body, emp.name);
      return json(res, 200, { ok: true, id });
    }
  }
  if (/^\/api\/leave-requests\/([^/]+)\/(approve|reject)$/.test(pathname)) {
    const parts = pathname.split('/');
    const reqId = parts[3], action = parts[4];
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'manager','hr','director')) return;
    const body = await readBody(req);
    try {
      if (action === 'approve') Q.approveLeaveRequest(reqId, emp.id, body.comment, emp.name);
      else Q.rejectLeaveRequest(reqId, emp.id, body.reason, emp.name);
      return json(res, 200, { ok: true });
    } catch (e) { return json(res, 400, { error: e.message }); }
  }

  // ── Activity endpoints ────────────────────────────────────────────────────────
  if (pathname === '/api/activity') {
    const emp = requireAuth(req, res); if (!emp) return;
    const filters = {};
    if (emp.role === 'employee') filters.employeeId = emp.id;
    if (parsed.query.date)  filters.date  = parsed.query.date;
    if (parsed.query.dept)  filters.dept  = parsed.query.dept;
    if (parsed.query.team)  filters.team  = parsed.query.team;
    return json(res, 200, Q.getActivityLog(filters));
  }

  if (pathname === '/api/activity/24h') {
    const emp = requireAuth(req, res); if (!emp) return;
    const filters = {};
    if (emp.role === 'employee') filters.employeeId = emp.id;
    if (parsed.query.date)       filters.date = parsed.query.date;
    if (parsed.query.dept)       filters.dept = parsed.query.dept;
    if (parsed.query.team)       filters.team = parsed.query.team;
    if (parsed.query.employee_id && ['manager','hr','director'].includes(emp.role))
      filters.employeeId = parseInt(parsed.query.employee_id);
    return json(res, 200, Q.getActivity24h(filters));
  }

  if (pathname === '/api/activity/live') {
    const emp = requireAuth(req, res); if (!emp) return;
    return json(res, 200, Q.getLiveStats());
  }

  if (pathname === '/api/activity/rollup') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'manager','hr','director')) return;
    const groupBy = parsed.query.by || 'dept';
    const date    = parsed.query.date || new Date().toISOString().slice(0,10);
    return json(res, 200, Q.getActivityRollup(groupBy, date));
  }

  if (pathname === '/api/activity/today-summary') {
    const emp = requireAuth(req, res); if (!emp) return;
    return json(res, 200, Q.getTodaySummary());
  }


  // ── Entra ID ──────────────────────────────────────────────────────────────────
  if (pathname === '/api/entra/config') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    if (method === 'GET') {
      const cfg = Q.getEntraConfig();
      // Mask secret
      if (cfg.client_secret) cfg.client_secret = '••••••••••••••••';
      return json(res, 200, cfg);
    }
    if (method === 'POST') {
      const body = await readBody(req);
      // Don't overwrite secret if masked
      if (body.client_secret === '••••••••••••••••') {
        const current = Q.getEntraConfig();
        body.client_secret = current.client_secret;
      }
      Q.saveEntraConfig(body, emp.name);
      entra.stopSyncScheduler();
      entra.startSyncScheduler(broadcast);
      return json(res, 200, { ok: true });
    }
  }
  if (pathname === '/api/entra/test' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    const body = await readBody(req);
    const cfg = { ...Q.getEntraConfig(), ...body };
    try {
      const result = await entra.testConnection(cfg);
      return json(res, 200, result);
    } catch (e) {
      return json(res, 400, { error: e.message });
    }
  }
  if (pathname === '/api/entra/sync' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    const cfg = Q.getEntraConfig();
    if (!cfg.tenant_id) return json(res, 400, { error: 'Entra not configured' });
    entra.runSync(cfg, broadcast).catch(e => console.error('[Entra]', e));
    return json(res, 200, { ok: true, message: 'Sync started — watch the SSE stream for progress' });
  }

  // ── SSE stream for sync logs ──────────────────────────────────────────────────
  if (pathname === '/api/sync-stream') {
    const emp = requireAuth(req, res); if (!emp) return;
    res.writeHead(200, { 'Content-Type':'text/event-stream','Cache-Control':'no-cache','Connection':'keep-alive' });
    res.write('data: {"type":"connected"}\n\n');
    sseClients.add(res);
    req.on('close', () => sseClients.delete(res));
    return;
  }

  // ── Audit log ─────────────────────────────────────────────────────────────────
  if (pathname === '/api/audit-log') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr', 'director')) return;
    const limit = parseInt(parsed.query.limit) || 100;
    return json(res, 200, Q.getAuditLog(limit));
  }

  // ── Change password ────────────────────────────────────────────────────────────
  if (pathname === '/api/change-password' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    const body = await readBody(req);
    if (!Q.verifyPassword(emp, body.current_password)) return json(res, 400, { error: 'Current password incorrect' });
    Q.setPassword(emp.id, body.new_password, emp.name);
    return json(res, 200, { ok: true });
  }

  // ── Static files ──────────────────────────────────────────────────────────────
  let filePath = path.join(__dirname, 'public', pathname === '/' ? 'index.html' : pathname);
  if (!fs.existsSync(filePath)) filePath = path.join(__dirname, 'public', 'index.html');
  try {
    const content = fs.readFileSync(filePath);
    const ext  = path.extname(filePath);
    res.writeHead(200, { 'Content-Type': MIME[ext] || 'text/plain' });
    res.end(content);
  } catch {
    res.writeHead(404); res.end('Not found');
  }
});

server.listen(PORT, () => console.log(`\nWorkIQ running → http://localhost:${PORT}\n`));

function safeUser(emp) {
  if (!emp) return null;
  const { password_hash, ...safe } = emp;
  return safe;
}
