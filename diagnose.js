/**
 * WorkIQ Diagnostic Tool
 * Run: node diagnose.js
 * Shows exactly what data is being retrieved from Entra ID and Teams
 */

'use strict';
const https  = require('https');
const path   = require('path');

// ── Load config from DB ───────────────────────────────────────────────────────
let cfg, db, Q;
try {
  const mod = require(path.join(__dirname, 'db'));
  db = mod.db; Q = mod.Q;
  cfg = Q.getEntraConfig();
} catch(e) {
  console.error('ERROR: Cannot load db.js —', e.message);
  process.exit(1);
}

const SEP  = '─'.repeat(60);
const OK   = '\x1b[32m✓\x1b[0m';
const FAIL = '\x1b[31m✗\x1b[0m';
const WARN = '\x1b[33m⚠\x1b[0m';
const INFO = '\x1b[36mℹ\x1b[0m';
const BOLD = s => `\x1b[1m${s}\x1b[0m`;

function printSection(title) {
  console.log('\n' + SEP);
  console.log(BOLD(' ' + title));
  console.log(SEP);
}

// ── HTTPS helpers ─────────────────────────────────────────────────────────────
function httpsPost(hostname, path, headers, body) {
  return new Promise((resolve, reject) => {
    const opts = { hostname, path, method:'POST', headers:{...headers,'Content-Length':Buffer.byteLength(body)} };
    let r='';
    const req = https.request(opts, res=>{res.on('data',d=>r+=d);res.on('end',()=>{try{resolve({status:res.statusCode,body:JSON.parse(r)});}catch{resolve({status:res.statusCode,body:r});}});});
    req.on('error',reject);
    req.setTimeout(15000,()=>{req.destroy();reject(new Error('Timeout'));});
    req.write(body); req.end();
  });
}

function httpsGet(urlOrPath, token) {
  return new Promise((resolve, reject) => {
    let hostname='graph.microsoft.com', p=urlOrPath;
    if (urlOrPath.startsWith('https://')) { try{const u=new URL(urlOrPath);hostname=u.hostname;p=u.pathname+u.search;}catch{} }
    else if (!urlOrPath.startsWith('/')) p='/v1.0/'+urlOrPath;
    const opts = { hostname, path:p, method:'GET', headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json',ConsistencyLevel:'eventual'} };
    let r='';
    const req = https.request(opts, res=>{res.on('data',d=>r+=d);res.on('end',()=>{try{resolve({status:res.statusCode,body:JSON.parse(r)});}catch{resolve({status:res.statusCode,body:r});}});});
    req.on('error',reject);
    req.setTimeout(15000,()=>{req.destroy();reject(new Error('Timeout'));});
    req.end();
  });
}

function httpsPostGraph(token, relPath, bodyObj) {
  const bodyStr = JSON.stringify(bodyObj);
  return new Promise((resolve, reject) => {
    const opts = { hostname:'graph.microsoft.com', path:'/v1.0/'+relPath, method:'POST',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json','Content-Length':Buffer.byteLength(bodyStr)} };
    let r='';
    const req = https.request(opts, res=>{res.on('data',d=>r+=d);res.on('end',()=>{try{resolve({status:res.statusCode,body:JSON.parse(r)});}catch{resolve({status:res.statusCode,body:r});}});});
    req.on('error',reject);
    req.setTimeout(15000,()=>{req.destroy();reject(new Error('Timeout'));});
    req.write(bodyStr); req.end();
  });
}

// ── Main diagnostic ───────────────────────────────────────────────────────────
async function main() {
  console.log('\n' + '═'.repeat(60));
  console.log(BOLD('  WorkIQ Diagnostic Tool — ' + new Date().toISOString()));
  console.log('═'.repeat(60));

  // ── 1. DB state ─────────────────────────────────────────────────────────────
  printSection('1. DATABASE STATE');
  const empCount  = db.prepare('SELECT COUNT(*) as c FROM employees WHERE active=1').get().c;
  const balCount  = db.prepare('SELECT COUNT(*) as c FROM leave_balances').get().c;
  const actCount  = db.prepare('SELECT COUNT(*) as c FROM activity_log').get().c;
  const actToday  = db.prepare("SELECT COUNT(*) as c FROM activity_log WHERE date(event_time)=date('now')").get().c;
  const today     = new Date().toISOString().slice(0,10);

  let presCount = 0;
  try { presCount = db.prepare('SELECT COUNT(*) as c FROM employee_presence').get().c; } catch {}

  const entraLinked = db.prepare("SELECT COUNT(*) as c FROM employees WHERE active=1 AND entra_id IS NOT NULL AND entra_id!=''").get().c;

  console.log(`${empCount>0?OK:FAIL} Employees in DB:        ${empCount}`);
  console.log(`${entraLinked>0?OK:WARN} Employees with Entra ID: ${entraLinked} / ${empCount}`);
  console.log(`${balCount>0?OK:WARN} Leave balance rows:      ${balCount}`);
  console.log(`${actCount>0?OK:WARN} Activity log total:      ${actCount} events`);
  console.log(`${actToday>0?OK:WARN} Activity log today:      ${actToday} events (${today})`);
  console.log(`${presCount>0?OK:WARN} Presence rows:           ${presCount}`);

  if (actCount === 0) {
    console.log(`\n${INFO} Activity log is EMPTY — this is why the dashboard shows everyone offline.`);
    console.log(`   Either sign-in log sync is failing, or Entra credentials aren't saved yet.`);
  }
  if (entraLinked === 0) {
    console.log(`\n${WARN} No employees have Entra IDs — user directory sync has not run yet.`);
  }

  // Show last 5 activity events
  if (actCount > 0) {
    console.log(`\n  Last 5 activity events:`);
    const recent = db.prepare("SELECT event_type,event_time,entra_id,ip_address FROM activity_log ORDER BY event_time DESC LIMIT 5").all();
    recent.forEach(r=>console.log(`    ${r.event_type.padEnd(12)} ${r.event_time?.slice(0,19)} ${r.ip_address||''}`));
  }

  // ── 2. Entra config ──────────────────────────────────────────────────────────
  printSection('2. ENTRA ID CONFIGURATION');
  console.log(`  Tenant ID:    ${cfg.tenant_id   ? OK+' '+cfg.tenant_id : FAIL+' NOT SET'}`);
  console.log(`  Client ID:    ${cfg.client_id   ? OK+' '+cfg.client_id : FAIL+' NOT SET'}`);
  console.log(`  Client secret:${cfg.client_secret? OK+' '+'•'.repeat(8)  : FAIL+' NOT SET'}`);
  console.log(`  Redirect URI: ${cfg.redirect_uri || '(auto from request host)'}`);
  console.log(`  Sync interval: ${cfg.sync_interval_min||5} min (full) + 1 min (presence)`);
  console.log(`  Idle threshold: ${cfg.idle_threshold_min||15} min`);
  console.log(`  Last sync: ${cfg.last_sync_at||'never'}`);
  console.log(`  Last status: ${cfg.last_sync_status||'—'}`);
  console.log(`  Last message: ${cfg.last_sync_msg||'—'}`);

  if (!cfg.tenant_id || !cfg.client_id || !cfg.client_secret) {
    console.log(`\n${FAIL} CREDENTIALS NOT CONFIGURED — go to HR Admin → Entra ID config and save your credentials.`);
    console.log('   Skipping all API tests.\n');
    return;
  }

  // ── 3. Authentication ────────────────────────────────────────────────────────
  printSection('3. AUTHENTICATION TEST');
  let token;
  try {
    const body = new URLSearchParams({
      grant_type:'client_credentials', client_id:cfg.client_id,
      client_secret:cfg.client_secret, scope:'https://graph.microsoft.com/.default'
    }).toString();
    const r = await httpsPost('login.microsoftonline.com', `/${cfg.tenant_id}/oauth2/v2.0/token`,
      {'Content-Type':'application/x-www-form-urlencoded'}, body);
    if (r.body.access_token) {
      token = r.body.access_token;
      const exp = r.body.expires_in;
      console.log(`${OK} Access token obtained (expires in ${exp}s)`);
      console.log(`   Token preview: ${token.slice(0,40)}...`);
    } else {
      console.log(`${FAIL} Auth failed: ${r.body.error} — ${r.body.error_description}`);
      return;
    }
  } catch(e) {
    console.log(`${FAIL} Auth exception: ${e.message}`);
    return;
  }

  // ── 4. Permission tests ──────────────────────────────────────────────────────
  printSection('4. GRAPH API PERMISSION TESTS');

  // 4a. Organization info
  try {
    const r = await httpsGet('organization?$select=displayName', token);
    const name = r.body.value?.[0]?.displayName;
    console.log(`${OK} Organization.Read: ${name}`);
  } catch(e) { console.log(`${FAIL} Organization.Read: ${e.message}`); }

  // 4b. Users
  let sampleEntraIds = [];
  try {
    const r = await httpsGet('users?$select=id,displayName,mail,userPrincipalName,department&$top=5', token);
    if (r.status === 200 && r.body.value) {
      const users = r.body.value;
      sampleEntraIds = users.map(u=>u.id).filter(Boolean);
      console.log(`${OK} User.Read.All: ${users.length} sample users`);
      users.forEach(u=>console.log(`     ${(u.displayName||'?').padEnd(25)} ${u.mail||u.userPrincipalName||'no email'} | dept: ${u.department||'—'}`));
      // Count total
      try {
        const countR = await httpsGet('users/$count', token);
        if (typeof countR.body === 'number') console.log(`   Total users in directory: ${countR.body}`);
      } catch {}
    } else {
      console.log(`${FAIL} User.Read.All: HTTP ${r.status} — ${JSON.stringify(r.body.error||r.body).slice(0,120)}`);
    }
  } catch(e) { console.log(`${FAIL} User.Read.All: ${e.message}`); }

  // 4c. Sign-in logs
  console.log('');
  try {
    // Use URL object to avoid encoding issues
    const fv2h = `createdDateTime ge ${new Date(Date.now()-2*60*60*1000).toISOString()}`.replace(/ /g,'%20').replace(/:/g,'%3A');
    const r = await httpsGet(`/v1.0/auditLogs/signIns?$filter=${fv2h}&$top=5&$select=id,createdDateTime,userPrincipalName,status,ipAddress`, token);
    if (r.status === 200 && r.body.value) {
      const logs = r.body.value;
      console.log(`${OK} AuditLog.Read.All (sign-in logs): ${logs.length} events in last 2 hours`);
      if (logs.length === 0) {
        const fv24 = `createdDateTime ge ${new Date(Date.now()-24*60*60*1000).toISOString()}`.replace(/ /g,'%20').replace(/:/g,'%3A');
        const r24 = await httpsGet(`/v1.0/auditLogs/signIns?$filter=${fv24}&$top=5&$select=id,createdDateTime,userPrincipalName,status`, token);
        if (r24.status===200) console.log(`   Last 24h: ${r24.body.value?.length||0} events found`);
        else console.log(`${WARN} No events in last 2h — try checking Azure Portal sign-in logs directly`);
      } else {
        logs.forEach(l=>console.log(`     ${(l.createdDateTime||'').slice(0,19)} ${(l.userPrincipalName||'?').slice(0,35).padEnd(35)} ${l.status?.errorCode===0?'Success':'Failed'} ${l.ipAddress||''}`));
      }
    } else if (r.status === 403) {
      console.log(`${FAIL} AuditLog.Read.All: PERMISSION DENIED (HTTP 403)`);
      console.log(`   → This permission requires Azure AD P1 or P2 license`);
      console.log(`   → Without it, activity tracking uses Teams presence only`);
    } else {
      console.log(`${FAIL} AuditLog.Read.All: HTTP ${r.status} — ${JSON.stringify(r.body).slice(0,200)}`);
    }
  } catch(e) { console.log(`${FAIL} AuditLog.Read.All: ${e.message}`); }

  // 4d. Teams presence
  console.log('');
  if (sampleEntraIds.length > 0) {
    try {
      const r = await httpsPostGraph(token, 'communications/getPresencesByUserId', { ids: sampleEntraIds.slice(0,3) });
      if (r.status === 200 && r.body.value) {
        const pres = r.body.value;
        console.log(`${OK} Presence.Read.All (Teams): ${pres.length} results`);
        pres.forEach(p=>console.log(`     ${p.id?.slice(0,8)}... availability: ${p.availability||'?'} | activity: ${p.activity||'?'}`));
      } else if (r.status === 403) {
        console.log(`${FAIL} Presence.Read.All: PERMISSION DENIED (HTTP 403)`);
        console.log(`   → Go to Azure Portal → App Registration → API permissions`);
        console.log(`   → Add: Microsoft Graph → Application permissions → Presence.Read.All`);
        console.log(`   → Then click "Grant admin consent"`);
      } else {
        console.log(`${FAIL} Presence.Read.All: HTTP ${r.status} — ${JSON.stringify(r.body).slice(0,200)}`);
      }
    } catch(e) { console.log(`${FAIL} Presence.Read.All: ${e.message}`); }
  } else {
    console.log(`${WARN} Presence test skipped — no Entra IDs available from user test`);
  }

  // ── 5. Run a live sync and show results ──────────────────────────────────────
  printSection('5. LIVE SYNC TEST (running now)');
  console.log('Running a full sync... this may take 10-30 seconds.\n');
  try {
    const { runSync } = require(path.join(__dirname, 'entra'));
    const result = await runSync(cfg, (event) => {
      if (event.type === 'sync_log') {
        const icon = event.level==='error'?FAIL:event.level==='warn'?WARN:OK;
        console.log(`  ${icon} [${event.ts.slice(11,19)}] ${event.msg}`);
      }
    });
    console.log(`\n  Result: users +${result.usersAdded} added, ${result.usersUpdated} updated | ${result.logsImported} sign-ins | ${result.presenceSynced} presence`);
  } catch(e) {
    console.log(`${FAIL} Sync failed: ${e.message}`);
  }

  // ── 6. Post-sync DB state ────────────────────────────────────────────────────
  printSection('6. DATABASE STATE AFTER SYNC');
  const actAfter    = db.prepare('SELECT COUNT(*) as c FROM activity_log').get().c;
  const actTodayAfter = db.prepare("SELECT COUNT(*) as c FROM activity_log WHERE date(event_time)=date('now')").get().c;
  let presAfter = 0;
  try { presAfter = db.prepare('SELECT COUNT(*) as c FROM employee_presence').get().c; } catch {}
  const entraAfter = db.prepare("SELECT COUNT(*) as c FROM employees WHERE active=1 AND entra_id IS NOT NULL AND entra_id!=''").get().c;

  console.log(`  Activity log total: ${actCount} → ${actAfter} (+${actAfter-actCount})`);
  console.log(`  Activity log today: ${actToday} → ${actTodayAfter} (+${actTodayAfter-actToday})`);
  console.log(`  Presence rows:      ${presCount} → ${presAfter} (+${presAfter-presCount})`);
  console.log(`  Employees with Entra ID: ${entraLinked} → ${entraAfter}`);

  if (actTodayAfter > 0) {
    console.log(`\n${OK} Activity data now in DB — dashboard should show real data.`);
    console.log('  Latest events:');
    const latest = db.prepare("SELECT al.event_type,al.event_time,e.name,al.ip_address FROM activity_log al LEFT JOIN employees e ON al.employee_id=e.id ORDER BY al.event_time DESC LIMIT 5").all();
    latest.forEach(r=>console.log(`    ${r.event_type.padEnd(12)} ${r.event_time?.slice(0,19)} ${(r.name||'unknown').padEnd(25)} ${r.ip_address||''}`));
  } else if (actAfter === 0) {
    console.log(`\n${FAIL} Still no activity data after sync.`);
    console.log('  Most likely reason: AuditLog.Read.All permission denied (needs P1/P2 license).');
    console.log('  Action: Enable Teams presence (Presence.Read.All) as the primary tracking method.');
  }

  if (presAfter > 0) {
    console.log(`\n${OK} Teams presence data available:`);
    try {
      const pres = db.prepare('SELECT e.name, ep.availability, ep.activity, ep.synced_at FROM employee_presence ep JOIN employees e ON ep.employee_id=e.id LIMIT 10').all();
      pres.forEach(p=>console.log(`    ${(p.name||'?').padEnd(25)} ${(p.availability||'?').padEnd(20)} ${p.activity||''}`));
    } catch {}
  }

  // Manager relationship check
  console.log('');
  try {
    const withManager = db.prepare('SELECT COUNT(*) as c FROM employees WHERE manager_id IS NOT NULL AND active=1').get().c;
    const totalEmp    = db.prepare("SELECT COUNT(*) as c FROM employees WHERE active=1 AND role='employee'").get().c;
    console.log(`Manager links: ${withManager} / ${totalEmp} employees have a manager assigned`);
    if (withManager === 0) {
      console.log(`  ${WARN} No manager assignments yet. Options:`);
      console.log(`    a) Set managers in Azure AD (user profile → Manager field) — synced automatically`);
      console.log(`    b) Assign manually in HR Admin → Employee Directory → Edit employee → Direct manager`);
      console.log(`    c) If column missing: restart the server (pm2 restart workiq) to apply migration`);
    } else {
      const sample = db.prepare(`SELECT e.name, m.name as mgr_name FROM employees e
        JOIN employees m ON e.manager_id=m.id WHERE e.active=1 LIMIT 5`).all();
      console.log('  Sample assignments:');
      sample.forEach(s => console.log(`    ${s.name.padEnd(25)} → ${s.mgr_name}`));
    }
  } catch(e) {
    if (e.message.includes('no such column: manager_id')) {
      console.log(`${WARN} manager_id column not yet in DB.`);
      console.log(`  Fix: pm2 restart workiq  (the migration runs automatically on startup)`);
      console.log(`  Then run: node diagnose.js  again`);
    } else {
      console.log(`${WARN} Manager check error: ${e.message}`);
    }
  }

  // ── 7. Summary + action items ────────────────────────────────────────────────
  printSection('7. SUMMARY & ACTION ITEMS');
  const actOk   = actTodayAfter > 0;
  const presOk  = presAfter > 0;
  const empOk   = entraAfter > 0;

  if (actOk && presOk) {
    console.log(`${OK} Everything working — dashboard has real Entra + Teams data.`);
  } else {
    console.log('Action items to fix:');
    if (!empOk)  console.log(`  ${FAIL} Users not synced — check User.Read.All permission`);
    if (!actOk)  console.log(`  ${FAIL} No sign-in logs — likely missing AuditLog.Read.All (requires Azure AD P1/P2)`);
    if (!presOk) console.log(`  ${FAIL} No Teams presence — add Presence.Read.All permission in Azure Portal`);

    console.log('\nQuick fix steps:');
    console.log('  1. Azure Portal → Entra ID → App registrations → your app');
    console.log('  2. API permissions → Add a permission → Microsoft Graph → Application permissions');
    console.log('  3. Add: Presence.Read.All');
    if (!actOk) {
      console.log('  4. For sign-in logs: confirm your Azure AD license is P1 or P2');
      console.log('     (Free/Basic license = AuditLog.Read.All denied)');
    }
    console.log('  5. Click "Grant admin consent for [your org]"');
    console.log('  6. Run: node diagnose.js  (to verify)');
  }

  console.log('\n' + '═'.repeat(60) + '\n');
}

main().catch(e => {
  console.error('\n' + FAIL + ' Fatal error:', e.message);
  console.error(e.stack);
  process.exit(1);
});