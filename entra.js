'use strict';
const https = require('https');
const { Q, db } = require('./db');

// ── Microsoft Graph helpers ───────────────────────────────────────────────────
function graphRequest(accessToken, path) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'graph.microsoft.com',
      path: `/v1.0/${path}`,
      method: 'GET',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    };
    let body = '';
    const req = https.request(options, res => {
      res.on('data', d => body += d);
      res.on('end', () => {
        try { resolve(JSON.parse(body)); }
        catch (e) { reject(new Error('Invalid JSON from Graph API')); }
      });
    });
    req.on('error', reject);
    req.setTimeout(10000, () => { req.destroy(); reject(new Error('Graph API timeout')); });
    req.end();
  });
}

function getAccessToken(tenantId, clientId, clientSecret) {
  return new Promise((resolve, reject) => {
    const body = new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'https://graph.microsoft.com/.default',
    }).toString();
    const options = {
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body),
      },
    };
    let resp = '';
    const req = https.request(options, res => {
      res.on('data', d => resp += d);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(resp);
          if (parsed.access_token) resolve(parsed.access_token);
          else reject(new Error(parsed.error_description || 'Token request failed'));
        } catch (e) { reject(new Error('Token parse error')); }
      });
    });
    req.on('error', reject);
    req.setTimeout(10000, () => { req.destroy(); reject(new Error('Token request timeout')); });
    req.write(body);
    req.end();
  });
}

// ── Full sync ─────────────────────────────────────────────────────────────────
async function runSync(cfg, broadcast) {
  const log = (msg, type='info') => {
    console.log(`[Entra Sync] ${msg}`);
    if (broadcast) broadcast({ type: 'sync_log', msg, level: type, ts: new Date().toISOString() });
  };

  log('Starting Entra ID sync...');
  let token;
  try {
    token = await getAccessToken(cfg.tenant_id, cfg.client_id, cfg.client_secret);
    log('✓ Access token obtained');
  } catch (err) {
    log(`✗ Auth failed: ${err.message}`, 'error');
    throw err;
  }

  // ── 1. Sync user directory ──────────────────────────────────────────────────
  let usersAdded = 0, usersUpdated = 0;
  try {
    // Follow @odata.nextLink pagination to get ALL users, not just first 100
    let nextUrl = 'users?$select=id,displayName,mail,department,jobTitle&$top=999';
    const allUsers = [];
    while (nextUrl) {
      const resp = await graphRequest(token, nextUrl);
      (resp.value || []).forEach(u => allUsers.push(u));
      // nextLink contains the full URL; extract just the path after /v1.0/
      if (resp['@odata.nextLink']) {
        nextUrl = resp['@odata.nextLink'].replace(/^https:\/\/graph\.microsoft\.com\/v1\.0\//, '');
      } else {
        nextUrl = null;
      }
    }
    log(`Found ${allUsers.length} users in directory`);

    const year = new Date().getFullYear();
    const monthsElapsed = new Date().getMonth(); // 0-based
    const ltRows = db.prepare('SELECT id, days_per_year FROM leave_types WHERE active=1').all();
    const insertBal = db.prepare(
      'INSERT OR IGNORE INTO leave_balances (employee_id,leave_type_id,year,entitled,accrued,used,adjustment) VALUES (?,?,?,?,?,0,0)'
    );

    if (cfg.auto_add_users) {
      allUsers.forEach(u => {
        if (!u.mail) return;
        const existing = Q.getEmployeeByEmail(u.mail);
        const ini = (u.displayName||'XX').split(' ').map(w=>w[0]).join('').slice(0,2).toUpperCase();
        let empId;
        if (!existing) {
          const r = db.prepare('INSERT OR IGNORE INTO employees (entra_id,name,email,initials,dept,title,cal_id,role,bal_type,av_bg,av_color) VALUES (?,?,?,?,?,?,?,?,?,?,?)')
            .run(u.id, u.displayName||u.mail, u.mail, ini, u.department||'', u.jobTitle||'', cfg.default_cal_id||'cal-ae', 'employee', 'fixed', '#e6f2fb', '#004e8c');
          empId = r.lastInsertRowid;
          usersAdded++;
        } else {
          db.prepare("UPDATE employees SET entra_id=?,name=?,dept=?,title=?,updated_at=datetime('now') WHERE email=?")
            .run(u.id, u.displayName||existing.name, u.department||existing.dept, u.jobTitle||existing.title, u.mail);
          empId = existing.id;
          usersUpdated++;
        }
        // Ensure leave_balance rows exist for every leave type this year
        if (empId) {
          ltRows.forEach(lt => {
            const accrued = lt.days_per_year; // fixed by default for synced users
            insertBal.run(empId, lt.id, year, lt.days_per_year, accrued);
          });
        }
      });
      log(`Users: ${usersAdded} added, ${usersUpdated} updated`);
    }
  } catch (err) {
    log(`⚠ User sync skipped: ${err.message}`, 'warn');
  }

  // ── 2. Sync sign-in logs ────────────────────────────────────────────────────
  let logsImported = 0;
  try {
    const since = new Date(Date.now() - 24*60*60*1000).toISOString();
    // Only encode the datetime value itself, not the OData operators
    const sinceVal = since.replace(/:/g, '%3A').replace(/\+/g, '%2B');
    const insertLog = db.prepare(
      'INSERT OR IGNORE INTO activity_log (employee_id,entra_id,event_type,event_time,ip_address,raw) VALUES (?,?,?,?,?,?)'
    );
    const dupCheck = db.prepare(
      'SELECT id FROM activity_log WHERE entra_id=? AND event_time=? AND event_type=? LIMIT 1'
    );

    // Paginate sign-in logs
    let logUrl = `auditLogs/signIns?$filter=createdDateTime ge ${sinceVal}&$top=100&$select=id,createdDateTime,userPrincipalName,userId,status,ipAddress`;
    let totalLogs = 0;
    while (logUrl) {
      const logsResp = await graphRequest(token, logUrl);
      const logs = logsResp.value || [];
      totalLogs += logs.length;
      logs.forEach(entry => {
        const emp = Q.getEmployeeByEmail(entry.userPrincipalName);
        const eventType = (entry.status?.errorCode === 0) ? 'SignIn' : 'SignInFailed';
        try {
          const exists = dupCheck.get(entry.userId, entry.createdDateTime, eventType);
          if (!exists) {
            insertLog.run(emp?.id||null, entry.userId, eventType, entry.createdDateTime, entry.ipAddress||null, JSON.stringify(entry));
            logsImported++;
          }
        } catch {}
      });
      if (logsResp['@odata.nextLink']) {
        logUrl = logsResp['@odata.nextLink'].replace(/^https:\/\/graph\.microsoft\.com\/v1\.0\//, '');
      } else {
        logUrl = null;
      }
    }
    log(`✓ ${logsImported} sign-in events imported (${totalLogs} fetched)`);
  } catch (err) {
    log(`⚠ Sign-in log sync skipped: ${err.message}`, 'warn');
  }

  // ── 3. Update sync status ───────────────────────────────────────────────────
  const summary = `${usersAdded} added, ${usersUpdated} updated, ${logsImported} events`;
  db.prepare("UPDATE entra_config SET last_sync_at=datetime('now'),last_sync_status='success',last_sync_msg=? WHERE id=1").run(summary);
  db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run('System', 'Entra ID sync completed', 'entra_config', summary);
  log(`✓ Sync complete: ${summary}`);
  return { usersAdded, usersUpdated, logsImported };
}

// ── Test connection only ──────────────────────────────────────────────────────
async function testConnection(cfg) {
  const token = await getAccessToken(cfg.tenant_id, cfg.client_id, cfg.client_secret);
  const org = await graphRequest(token, 'organization?$select=displayName,verifiedDomains');
  const orgName = org.value?.[0]?.displayName || 'Unknown';
  const userCount = await graphRequest(token, 'users/$count');
  return { ok: true, orgName, userCount: typeof userCount === 'number' ? userCount : '?' };
}

// ── Periodic sync scheduler ───────────────────────────────────────────────────
let syncTimer = null;
function startSyncScheduler(broadcast) {
  if (syncTimer) clearInterval(syncTimer);
  const cfg = Q.getEntraConfig();
  if (!cfg.tenant_id || !cfg.client_id || !cfg.client_secret) return;
  const intervalMs = (cfg.sync_interval_min || 5) * 60 * 1000;
  console.log(`[Entra] Sync scheduler started — every ${cfg.sync_interval_min||5} min`);
  syncTimer = setInterval(async () => {
    try {
      const freshCfg = Q.getEntraConfig();
      if (freshCfg.tenant_id) await runSync(freshCfg, broadcast);
    } catch (err) {
      console.error('[Entra] Scheduled sync error:', err.message);
      db.prepare("UPDATE entra_config SET last_sync_status='error',last_sync_msg=? WHERE id=1").run(err.message);
    }
  }, intervalMs);
}

function stopSyncScheduler() {
  if (syncTimer) { clearInterval(syncTimer); syncTimer = null; }
}

module.exports = { runSync, testConnection, startSyncScheduler, stopSyncScheduler };
