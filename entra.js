'use strict';
const https = require('https');
const { Q, db } = require('./db');

// ── Generic HTTPS POST ────────────────────────────────────────────────────────
function httpsPost(hostname, path, headers, body) {
  return new Promise((resolve, reject) => {
    const opts = { hostname, path, method: 'POST', headers: { ...headers, 'Content-Length': Buffer.byteLength(body) } };
    let resp = '';
    const r = https.request(opts, res => { res.on('data', d => resp += d); res.on('end', () => { try { resolve(JSON.parse(resp)); } catch { reject(new Error('JSON parse error')); } }); });
    r.on('error', reject);
    r.setTimeout(15000, () => { r.destroy(); reject(new Error('Request timeout')); });
    r.write(body); r.end();
  });
}

// ── Generic HTTPS GET — handles both full URLs and /v1.0/ relative paths ──────
function httpsGet(urlOrPath, accessToken) {
  return new Promise((resolve, reject) => {
    let hostname = 'graph.microsoft.com';
    let path = urlOrPath;
    if (urlOrPath.startsWith('https://')) {
      try { const u = new URL(urlOrPath); hostname = u.hostname; path = u.pathname + u.search; } catch {}
    } else if (!urlOrPath.startsWith('/')) {
      path = '/v1.0/' + urlOrPath;
    }
    const opts = {
      hostname, path, method: 'GET',
      headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json', ConsistencyLevel: 'eventual' },
    };
    let body = '';
    const r = https.request(opts, res => {
      res.on('data', d => body += d);
      res.on('end', () => {
        try {
          const parsed = JSON.parse(body);
          if (parsed.error) reject(new Error(`Graph ${parsed.error.code}: ${parsed.error.message}`));
          else resolve(parsed);
        } catch { reject(new Error(`Graph returned non-JSON (status likely non-200)`)); }
      });
    });
    r.on('error', reject);
    r.setTimeout(15000, () => { r.destroy(); reject(new Error('Graph API timeout')); });
    r.end();
  });
}

// ── POST to Graph API ─────────────────────────────────────────────────────────
function graphPost(accessToken, relativePath, bodyObj) {
  return new Promise((resolve, reject) => {
    const bodyStr = JSON.stringify(bodyObj);
    const opts = {
      hostname: 'graph.microsoft.com',
      path: '/v1.0/' + relativePath,
      method: 'POST',
      headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(bodyStr) },
    };
    let resp = '';
    const r = https.request(opts, res => { res.on('data', d => resp += d); res.on('end', () => { try { resolve(JSON.parse(resp)); } catch { reject(new Error('POST JSON parse error')); } }); });
    r.on('error', reject);
    r.setTimeout(15000, () => { r.destroy(); reject(new Error('POST timeout')); });
    r.write(bodyStr); r.end();
  });
}

// ── OAuth2 client credentials token ──────────────────────────────────────────
function getAccessToken(tenantId, clientId, clientSecret) {
  const body = new URLSearchParams({
    grant_type: 'client_credentials', client_id: clientId,
    client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default',
  }).toString();
  return httpsPost('login.microsoftonline.com', `/${tenantId}/oauth2/v2.0/token`,
    { 'Content-Type': 'application/x-www-form-urlencoded' }, body
  ).then(r => {
    if (r.error) throw new Error(`Auth: ${r.error_description || r.error}`);
    return r.access_token;
  });
}

// ── Paginate all Graph pages ──────────────────────────────────────────────────
async function graphGetAll(token, startUrl) {
  const items = [];
  let url = startUrl;
  let pages = 0;
  while (url && pages < 50) {
    const resp = await httpsGet(url, token);
    (resp.value || []).forEach(i => items.push(i));
    url = resp['@odata.nextLink'] || null;
    pages++;
  }
  return items;
}

// ── Unified status model ──────────────────────────────────────────────────────
function unifyStatus({ hasSignInToday, msSinceLastSignIn, teamsAvailability, teamsActivity, isOnLeave, isHoliday, isWeekend, idleThresholdMs }) {
  if (isOnLeave)  return 'on_leave';
  if (isHoliday)  return 'holiday';
  if (isWeekend)  return 'weekend';

  // Teams presence — highest fidelity signal
  if (teamsAvailability) {
    const av = teamsAvailability.toLowerCase();
    if (['donotdisturb','presenting'].includes(av))                    return 'do_not_disturb';
    if (['available','busy','inacall','inaconferencecall','inameeting',
         'urgentinterruptionsonly'].includes(av))                      return 'active';
    if (['away','berightback'].includes(av))                           return 'idle';
    if (['offline','presenceunknown','offwork'].includes(av))          return 'offline';
  }

  // Fallback: sign-in log timing
  if (!hasSignInToday)                                                 return 'offline';
  if (msSinceLastSignIn !== null && msSinceLastSignIn < idleThresholdMs)         return 'active';
  if (msSinceLastSignIn !== null && msSinceLastSignIn < idleThresholdMs * 6)     return 'idle';
  return 'offline';
}

// ── Main sync ─────────────────────────────────────────────────────────────────
async function runSync(cfg, broadcast) {
  const log = (msg, level = 'info') => {
    console.log(`[Entra ${level}] ${msg}`);
    if (broadcast) broadcast({ type: 'sync_log', msg, level, ts: new Date().toISOString() });
  };

  log('Starting Entra ID sync…');
  let token;
  try {
    token = await getAccessToken(cfg.tenant_id, cfg.client_id, cfg.client_secret);
    log('✓ Access token obtained');
  } catch (err) {
    log(`✗ Authentication failed: ${err.message}`, 'error');
    db.prepare("UPDATE entra_config SET last_sync_status='error',last_sync_msg=? WHERE id=1").run(err.message);
    throw err;
  }

  let usersAdded = 0, usersUpdated = 0, logsImported = 0, presenceSynced = 0;

  // ── 1. User directory sync ────────────────────────────────────────────────────
  try {
    const users = await graphGetAll(token,
      'users?$select=id,displayName,mail,userPrincipalName,department,jobTitle&$top=999');
    log(`Found ${users.length} users in directory`);

    const year = new Date().getFullYear();
    const ltRows = db.prepare('SELECT id, days_per_year FROM leave_types WHERE active=1').all();
    const insertBal = db.prepare('INSERT OR IGNORE INTO leave_balances (employee_id,leave_type_id,year,entitled,accrued,used,adjustment) VALUES (?,?,?,?,?,0,0)');

    if (cfg.auto_add_users) {
      users.forEach(u => {
        const email = u.mail || u.userPrincipalName;
        if (!email) return;
        const existing = Q.getEmployeeByEmail(email);
        const ini = (u.displayName || 'XX').split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
        let empId;
        if (!existing) {
          const r = db.prepare('INSERT OR IGNORE INTO employees (entra_id,name,email,initials,dept,title,cal_id,role,bal_type,av_bg,av_color) VALUES (?,?,?,?,?,?,?,?,?,?,?)')
            .run(u.id, u.displayName || email, email, ini, u.department || '', u.jobTitle || '',
              cfg.default_cal_id || 'cal-ae', 'employee', 'fixed', '#e6f2fb', '#004e8c');
          empId = r.lastInsertRowid;
          usersAdded++;
        } else {
          db.prepare("UPDATE employees SET entra_id=?,name=?,dept=?,title=?,updated_at=datetime('now') WHERE email=?")
            .run(u.id, u.displayName || existing.name, u.department || existing.dept,
              u.jobTitle || existing.title, email);
          empId = existing.id;
          usersUpdated++;
        }
        if (empId) ltRows.forEach(lt => insertBal.run(empId, lt.id, year, lt.days_per_year, lt.days_per_year));
      });
      log(`✓ Users: ${usersAdded} added, ${usersUpdated} updated`);
    }
  } catch (err) {
    log(`⚠ User sync skipped: ${err.message}`, 'warn');
  }

  // ── 2. Sign-in logs sync ──────────────────────────────────────────────────────
  try {
    const since = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
    // Node's https.request rejects unescaped spaces and colons in the path.
    // Encode the entire filter value: spaces → %20, colons → %3A
    // Do NOT use URL class — it encodes $ → %24 breaking OData operators.
    const filterVal = `createdDateTime ge ${since}`.replace(/ /g, '%20').replace(/:/g, '%3A');
    const startPath = `/v1.0/auditLogs/signIns?$filter=${filterVal}&$top=100&$select=id,createdDateTime,userPrincipalName,userId,status,ipAddress,appDisplayName,location`;

    const insertLog = db.prepare('INSERT OR IGNORE INTO activity_log (employee_id,entra_id,event_type,event_time,ip_address,raw) VALUES (?,?,?,?,?,?)');
    const dupCheck  = db.prepare('SELECT id FROM activity_log WHERE entra_id=? AND event_time=? AND event_type=? LIMIT 1');

    let url = startPath, totalFetched = 0, pages = 0;
    while (url && pages < 20) {
      const resp = await httpsGet(url, token);
      if (resp.error) throw new Error(`${resp.error.code}: ${resp.error.message}`);
      const entries = resp.value || [];
      totalFetched += entries.length;

      entries.forEach(entry => {
        const email   = entry.userPrincipalName;
        const emp     = email ? Q.getEmployeeByEmail(email) : null;
        const isSuccess = entry.status?.errorCode === 0;
        const eventType = isSuccess ? 'SignIn' : 'SignInFailed';
        const entraId   = entry.userId || entry.id;
        try {
          if (!dupCheck.get(entraId, entry.createdDateTime, eventType)) {
            insertLog.run(emp?.id || null, entraId, eventType, entry.createdDateTime,
              entry.ipAddress || null,
              JSON.stringify({ app: entry.appDisplayName, city: entry.location?.city }));
            if (isSuccess) logsImported++;
          }
        } catch {}
      });

      // nextLink is absolute URL — pass as-is to httpsGet which handles both
      url = resp['@odata.nextLink'] || null;
      pages++;
    }
    log(`✓ Sign-in logs: ${logsImported} successful sign-ins imported (${totalFetched} fetched)`);
  } catch (err) {
    log(`⚠ Sign-in log sync failed: ${err.message}`, 'warn');
    if (err.message.includes('403') || err.message.includes('Authorization_RequestDenied') || err.message.includes('Forbidden')) {
      log('  → AuditLog.Read.All denied — your Azure AD license may not include P1/P2', 'warn');
      log('  → Check: portal.azure.com → Licenses → All products', 'warn');
    } else if (err.message.includes('unescaped') || err.message.includes('escape') || err.message.includes('invalid')) {
      log(`  → URL issue: ${err.message}`, 'warn');
    }
  }

  // ── 3. Teams presence sync ────────────────────────────────────────────────────
  try {
    // Ensure presence table exists
    db.exec(`CREATE TABLE IF NOT EXISTS employee_presence (
      employee_id INTEGER PRIMARY KEY REFERENCES employees(id),
      entra_id TEXT, availability TEXT, activity TEXT,
      status_message TEXT, synced_at TEXT
    )`);

    const empWithEntra = db.prepare("SELECT id, entra_id FROM employees WHERE active=1 AND entra_id IS NOT NULL AND entra_id != ''").all();
    if (empWithEntra.length === 0) {
      log('⚠ No employees with Entra IDs yet — sync users first', 'warn');
    } else {
      const upsert = db.prepare('INSERT OR REPLACE INTO employee_presence (employee_id,entra_id,availability,activity,status_message,synced_at) VALUES (?,?,?,?,?,?)');
      const now = new Date().toISOString();
      const BATCH = 650;

      for (let i = 0; i < empWithEntra.length; i += BATCH) {
        const batch = empWithEntra.slice(i, i + BATCH);
        const ids = batch.map(e => e.entra_id);
        const resp = await graphPost(token, 'communications/getPresencesByUserId', { ids });
        if (resp.error) throw new Error(`${resp.error.code}: ${resp.error.message}`);
        (resp.value || []).forEach(p => {
          const emp = batch.find(e => e.entra_id === p.id);
          if (!emp) return;
          upsert.run(emp.id, p.id, p.availability, p.activity, p.statusMessage?.message?.content || null, now);
          presenceSynced++;
        });
      }
      log(`✓ Teams presence: ${presenceSynced} employees synced`);
    }
  } catch (err) {
    log(`⚠ Teams presence sync failed: ${err.message}`, 'warn');
    if (err.message.includes('Authorization') || err.message.includes('Forbidden') || err.message.includes('403')) {
      log('  → Add Presence.Read.All permission in your Azure App Registration → API permissions', 'warn');
    }
  }

  // ── 4. Finalise ───────────────────────────────────────────────────────────────
  const summary = `${usersAdded} added, ${usersUpdated} updated, ${logsImported} sign-ins, ${presenceSynced} presence`;
  db.prepare("UPDATE entra_config SET last_sync_at=datetime('now'),last_sync_status='success',last_sync_msg=? WHERE id=1").run(summary);
  db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run('System', 'Entra sync completed', 'entra_config', summary);
  log(`✓ Sync complete: ${summary}`);
  return { usersAdded, usersUpdated, logsImported, presenceSynced };
}

// ── Test connection with permission check ─────────────────────────────────────
async function testConnection(cfg) {
  const token = await getAccessToken(cfg.tenant_id, cfg.client_id, cfg.client_secret);
  const org = await httpsGet('organization?$select=displayName', token);
  const orgName = org.value?.[0]?.displayName || 'Unknown';

  const perms = { users: false, auditLogs: false, presence: false };
  try { await httpsGet('users?$top=1&$select=id', token); perms.users = true; } catch {}
  try { await httpsGet('auditLogs/signIns?$top=1&$select=id', token); perms.auditLogs = true; } catch {}
  try { await graphPost(token, 'communications/getPresencesByUserId', { ids: [] }); perms.presence = true; } catch (e) {
    // empty ids returns 200 with empty value — any auth error means no permission
    if (!e.message.includes('BadRequest') && !e.message.includes('400')) perms.presence = false;
    else perms.presence = true;
  }

  return { ok: true, orgName, perms };
}

// ── Scheduler — two separate timers ──────────────────────────────────────────
// Presence: every 1 minute (Teams status changes in seconds)
// Full sync: every N minutes (configurable, default 5) for users + sign-in logs
let syncTimer    = null;
let presenceTimer = null;

async function syncPresenceOnly(cfg, broadcast) {
  const log = (msg, level='info') => {
    if (broadcast) broadcast({ type:'sync_log', msg, level, ts: new Date().toISOString() });
  };
  let token;
  try {
    token = await getAccessToken(cfg.tenant_id, cfg.client_id, cfg.client_secret);
  } catch { return; } // silent — full sync will surface auth errors

  let presenceSynced = 0;
  try {
    db.exec(`CREATE TABLE IF NOT EXISTS employee_presence (
      employee_id INTEGER PRIMARY KEY REFERENCES employees(id),
      entra_id TEXT, availability TEXT, activity TEXT,
      status_message TEXT, synced_at TEXT
    )`);
    const empWithEntra = db.prepare("SELECT id, entra_id FROM employees WHERE active=1 AND entra_id IS NOT NULL AND entra_id != ''").all();
    if (!empWithEntra.length) return;

    const upsert = db.prepare('INSERT OR REPLACE INTO employee_presence (employee_id,entra_id,availability,activity,status_message,synced_at) VALUES (?,?,?,?,?,?)');
    const now = new Date().toISOString();
    const BATCH = 650;

    for (let i = 0; i < empWithEntra.length; i += BATCH) {
      const batch = empWithEntra.slice(i, i + BATCH);
      const resp  = await graphPost(token, 'communications/getPresencesByUserId', { ids: batch.map(e => e.entra_id) });
      if (resp.error) return;
      (resp.value || []).forEach(p => {
        const emp = batch.find(e => e.entra_id === p.id);
        if (emp) { upsert.run(emp.id, p.id, p.availability, p.activity, p.statusMessage?.message?.content || null, now); presenceSynced++; }
      });
    }
    if (presenceSynced > 0) log(`↻ Presence refreshed: ${presenceSynced} employees`);
  } catch {}
}

function startSyncScheduler(broadcast) {
  if (syncTimer)     clearInterval(syncTimer);
  if (presenceTimer) clearInterval(presenceTimer);

  const cfg = Q.getEntraConfig();
  if (!cfg.tenant_id || !cfg.client_id || !cfg.client_secret) return;

  const fullIntervalMs     = (cfg.sync_interval_min || 5) * 60 * 1000;
  const presenceIntervalMs = 60 * 1000; // always 1 minute

  console.log(`[Entra] Full sync every ${cfg.sync_interval_min || 5} min | Presence every 1 min`);

  // Full sync: users + sign-in logs + presence
  syncTimer = setInterval(async () => {
    try {
      const freshCfg = Q.getEntraConfig();
      if (freshCfg.tenant_id) await runSync(freshCfg, broadcast);
    } catch (err) {
      console.error('[Entra] Full sync error:', err.message);
      db.prepare("UPDATE entra_config SET last_sync_status='error',last_sync_msg=? WHERE id=1").run(err.message);
    }
  }, fullIntervalMs);

  // Presence-only fast refresh
  presenceTimer = setInterval(async () => {
    try {
      const freshCfg = Q.getEntraConfig();
      if (freshCfg.tenant_id) await syncPresenceOnly(freshCfg, broadcast);
    } catch {}
  }, presenceIntervalMs);
}

function stopSyncScheduler() {
  if (syncTimer)     { clearInterval(syncTimer);     syncTimer     = null; }
  if (presenceTimer) { clearInterval(presenceTimer); presenceTimer = null; }
}

module.exports = { runSync, testConnection, startSyncScheduler, stopSyncScheduler, unifyStatus };
