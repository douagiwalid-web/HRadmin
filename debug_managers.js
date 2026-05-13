'use strict';
/**
 * WorkIQ Manager Debug Tool
 * Run: node debug_managers.js
 * Shows exactly what Graph API returns for manager expand
 */
const https = require('https');
const path  = require('path');
const { Q, db } = require(path.join(__dirname, 'db'));

function httpsPost(hostname, path, headers, body) {
  return new Promise((resolve, reject) => {
    const opts = { hostname, path, method:'POST', headers:{...headers,'Content-Length':Buffer.byteLength(body)} };
    let r='';
    const req = https.request(opts, res=>{res.on('data',d=>r+=d);res.on('end',()=>{try{resolve(JSON.parse(r));}catch{reject(new Error('JSON parse'));}});});
    req.on('error',reject);
    req.setTimeout(20000,()=>{req.destroy();reject(new Error('Timeout'));});
    req.write(body); req.end();
  });
}

function httpsGet(urlOrPath, token) {
  return new Promise((resolve, reject) => {
    let hostname='graph.microsoft.com', p=urlOrPath;
    if (urlOrPath.startsWith('https://')) { try{const u=new URL(urlOrPath);hostname=u.hostname;p=u.pathname+u.search;}catch{} }
    else if (!urlOrPath.startsWith('/')) p='/v1.0/'+urlOrPath;
    const opts = { hostname, path:p, method:'GET',
      headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json',ConsistencyLevel:'eventual'} };
    let r='';
    const req = https.request(opts, res=>{res.on('data',d=>r+=d);res.on('end',()=>{try{resolve({status:res.statusCode,body:JSON.parse(r)});}catch{resolve({status:res.statusCode,body:r});}});});
    req.on('error',reject);
    req.setTimeout(20000,()=>{req.destroy();reject(new Error('Timeout'));});
    req.end();
  });
}

const SEP  = '─'.repeat(60);
const OK   = '\x1b[32m✓\x1b[0m';
const FAIL = '\x1b[31m✗\x1b[0m';
const WARN = '\x1b[33m⚠\x1b[0m';
const INFO = '\x1b[36mℹ\x1b[0m';

async function main() {
  console.log('\n' + '═'.repeat(60));
  console.log('\x1b[1m  WorkIQ Manager Debug Tool — ' + new Date().toISOString() + '\x1b[0m');
  console.log('═'.repeat(60));

  const cfg = Q.getEntraConfig();
  if (!cfg.tenant_id || !cfg.client_id || !cfg.client_secret) {
    console.error(FAIL, 'Entra credentials not configured. Run node diagnose.js first.');
    process.exit(1);
  }

  // Get token
  console.log('\nGetting access token...');
  const body = new URLSearchParams({
    grant_type:'client_credentials', client_id:cfg.client_id,
    client_secret:cfg.client_secret, scope:'https://graph.microsoft.com/.default'
  }).toString();
  const tokenResp = await httpsPost('login.microsoftonline.com',
    `/${cfg.tenant_id}/oauth2/v2.0/token`,
    {'Content-Type':'application/x-www-form-urlencoded'}, body);
  if (tokenResp.error) { console.error(FAIL, 'Auth failed:', tokenResp.error_description); process.exit(1); }
  const token = tokenResp.access_token;
  console.log(OK, 'Token obtained');

  // ── TEST 1: Raw $expand=manager on a few users ────────────────────────────
  console.log('\n' + SEP);
  console.log('\x1b[1m TEST 1: $expand=manager on first 5 users\x1b[0m');
  console.log(SEP);
  console.log('URL: users?$select=id,displayName,mail,userPrincipalName&$expand=manager&$top=5\n');

  const r1 = await httpsGet('users?$select=id,displayName,mail,userPrincipalName&$expand=manager&$top=5', token);
  if (r1.status !== 200) {
    console.log(FAIL, 'HTTP', r1.status, JSON.stringify(r1.body).slice(0,200));
  } else {
    const users = r1.body.value || [];
    users.forEach(u => {
      const m = u.manager;
      console.log(`User: ${(u.displayName||'?').padEnd(30)} mail: ${u.mail||u.userPrincipalName||'(none)'}`);
      if (m) {
        console.log(`  ✓ manager object keys: ${Object.keys(m).join(', ')}`);
        console.log(`  ✓ manager.id:                ${m.id||'(none)'}`);
        console.log(`  ✓ manager.displayName:       ${m.displayName||'(none)'}`);
        console.log(`  ✓ manager.mail:              ${m.mail||'(null)'}`);
        console.log(`  ✓ manager.userPrincipalName: ${m.userPrincipalName||'(null)'}`);
        // Check if this manager exists in our DB
        const byId    = m.id    ? db.prepare('SELECT id,name FROM employees WHERE entra_id=?').get(m.id)    : null;
        const byEmail = (m.mail||m.userPrincipalName) ? db.prepare('SELECT id,name FROM employees WHERE email=?').get(m.mail||m.userPrincipalName) : null;
        console.log(`  DB lookup by entra_id: ${byId    ? OK+' '+byId.name    : FAIL+' not found'}`);
        console.log(`  DB lookup by email:    ${byEmail ? OK+' '+byEmail.name : FAIL+' not found (mail was null)'}`);
      } else {
        console.log(`  ${INFO} no manager set`);
      }
      console.log('');
    });
  }

  // ── TEST 2: Find a known user with manager (Abdelhak Amami from screenshot) ─
  console.log(SEP);
  console.log('\x1b[1m TEST 2: Look up Abdelhak Amami specifically\x1b[0m');
  console.log(SEP);
  const r2 = await httpsGet('users?$filter=startswith(mail,\'abdelhak.amami\')&$expand=manager&$select=id,displayName,mail,userPrincipalName&$top=1', token);
  if (r2.status === 200 && r2.body.value?.length > 0) {
    const u = r2.body.value[0];
    const m = u.manager;
    console.log(`User: ${u.displayName} (${u.mail||u.userPrincipalName})`);
    if (m) {
      console.log(OK, 'Manager object received:');
      console.log(JSON.stringify(m, null, 2));
      const byId = db.prepare('SELECT id,name,email FROM employees WHERE entra_id=?').get(m.id);
      console.log(`\nDB lookup by manager.id (${m.id?.slice(0,8)}...): ${byId ? OK+' Found: '+byId.name : FAIL+' NOT IN DB'}`);
      if (!byId) {
        console.log(`  ${WARN} Manager "${m.displayName||'?'}" has entra_id "${m.id}" but is not in employees table`);
        console.log(`  → Either they are not in the sync filter, or their email doesn't match`);
        // Try to find by name
        const byName = m.displayName ? db.prepare("SELECT id,name,email,entra_id FROM employees WHERE name LIKE ? LIMIT 3").all('%'+(m.displayName||'').split(' ')[0]+'%') : [];
        if (byName.length > 0) {
          console.log(`  Possible match by name:`);
          byName.forEach(e => console.log(`    ${e.name} | email: ${e.email} | entra_id: ${e.entra_id||'(none)'}`));
        }
      }
    } else {
      console.log(FAIL, 'No manager object — $expand=manager returned nothing for this user');
      console.log('This means either:');
      console.log('  a) The manager field requires $expand but is not being returned');
      console.log('  b) Try fetching the manager directly:');
      const r2b = await httpsGet(`users/${u.id}/manager?$select=id,displayName,mail,userPrincipalName`, token);
      console.log(`\nDirect manager endpoint (GET /users/${u.id}/manager):`);
      console.log(`  HTTP ${r2b.status}:`, JSON.stringify(r2b.body).slice(0,300));
    }
  } else {
    console.log(WARN, 'Abdelhak Amami not found via filter, HTTP:', r2.status);
  }

  // ── TEST 3: Check employees.entra_id values vs what Graph returns ─────────
  console.log('\n' + SEP);
  console.log('\x1b[1m TEST 3: entra_id column integrity\x1b[0m');
  console.log(SEP);
  const total    = db.prepare("SELECT COUNT(*) as c FROM employees WHERE active=1").get().c;
  const withId   = db.prepare("SELECT COUNT(*) as c FROM employees WHERE entra_id IS NOT NULL AND entra_id NOT LIKE 'demo-%' AND active=1").get().c;
  const withNull = db.prepare("SELECT COUNT(*) as c FROM employees WHERE entra_id IS NULL AND active=1").get().c;
  const withDemo = db.prepare("SELECT COUNT(*) as c FROM employees WHERE entra_id LIKE 'demo-%' AND active=1").get().c;
  console.log(`Total employees:      ${total}`);
  console.log(`With real entra_id:   ${withId} ${withId > 0 ? OK : FAIL}`);
  console.log(`With demo entra_id:   ${withDemo}`);
  console.log(`With NULL entra_id:   ${withNull}`);
  console.log('');

  // Show sample of entra_ids
  const sample = db.prepare("SELECT name, entra_id, email FROM employees WHERE active=1 LIMIT 5").all();
  console.log('Sample entra_ids:');
  sample.forEach(e => console.log(`  ${(e.name||'?').padEnd(28)} entra_id: ${(e.entra_id||'NULL').slice(0,36)}`));

  // ── TEST 4: Count users that HAVE manager set in Entra ─────────────────────
  console.log('\n' + SEP);
  console.log('\x1b[1m TEST 4: Count users with manager set in Entra (sample of 100)\x1b[0m');
  console.log(SEP);
  const r4 = await httpsGet('users?$select=id,displayName,mail&$expand=manager($select=id,displayName)&$top=100', token);
  if (r4.status === 200) {
    const users4 = r4.body.value || [];
    const withMgr = users4.filter(u => u.manager && u.manager.id);
    const noMgr   = users4.filter(u => !u.manager);
    console.log(`Sample: ${users4.length} users | with manager: ${withMgr.length} | without: ${noMgr.length}`);
    if (withMgr.length > 0) {
      console.log('\nFirst 5 with manager:');
      withMgr.slice(0,5).forEach(u => {
        const m = u.manager;
        console.log(`  ${(u.displayName||'?').padEnd(28)} → ${m.displayName||'?'} (id: ${m.id?.slice(0,8)}...)`);
        // Check if manager is in our DB
        const inDB = m.id ? db.prepare('SELECT name FROM employees WHERE entra_id=?').get(m.id) : null;
        console.log(`     In WorkIQ DB: ${inDB ? OK+' '+inDB.name : FAIL+' NOT FOUND — entra_id mismatch'}`);
      });
    }
  } else {
    console.log(FAIL, 'HTTP', r4.status, JSON.stringify(r4.body).slice(0,150));
  }

  // ── TEST 5: Simulate the actual sync second pass ──────────────────────────
  console.log('\n' + SEP);
  console.log('\x1b[1m TEST 5: Simulate manager sync pass\x1b[0m');
  console.log(SEP);
  if (r4.status === 200) {
    const users4 = r4.body.value || [];
    const withMgr = users4.filter(u => u.manager && u.manager.id);
    let linked=0, notFound=0, noObj=0;
    withMgr.forEach(u => {
      const email = u.mail || u.userPrincipalName;
      const mgrId = u.manager?.id;
      if (!mgrId) { noObj++; return; }
      const empRecord = email ? Q.getEmployeeByEmail(email) : db.prepare('SELECT * FROM employees WHERE entra_id=?').get(u.id);
      const mgrRecord = db.prepare('SELECT * FROM employees WHERE entra_id=?').get(mgrId);
      if (empRecord && mgrRecord && mgrRecord.id !== empRecord?.id) linked++;
      else if (!mgrRecord) notFound++;
    });
    console.log(`Simulation result (${withMgr.length} users with manager in Entra):`);
    console.log(`  Would link:       ${linked}`);
    console.log(`  Manager not in DB: ${notFound}`);
    if (notFound > 0) {
      console.log(`\n  ${WARN} ${notFound} managers exist in Entra but not in WorkIQ DB`);
      console.log('  This is the core issue — managers have different entra_ids in DB vs Graph');
      // Show which ones are missing
      let shown = 0;
      withMgr.forEach(u => {
        if (shown >= 5) return;
        const mgrId = u.manager?.id;
        const mgrRecord = mgrId ? db.prepare('SELECT * FROM employees WHERE entra_id=?').get(mgrId) : null;
        if (!mgrRecord) {
          console.log(`  Missing: ${(u.manager?.displayName||'?').padEnd(28)} entra_id: ${mgrId?.slice(0,36)}`);
          // Check if they exist by name
          const byName = db.prepare("SELECT name,entra_id,email FROM employees WHERE name LIKE ? LIMIT 1")
            .get('%'+(u.manager?.displayName||'').split(' ').slice(0,1).join('')+'%');
          if (byName) console.log(`    Possible match in DB: ${byName.name} | stored entra_id: ${byName.entra_id||'(none)'}`);
          shown++;
        }
      });
    }
  }

  console.log('\n' + '═'.repeat(60) + '\n');
}

main().catch(e => {
  console.error('\n\x1b[31m✗\x1b[0m Fatal:', e.message);
  process.exit(1);
});