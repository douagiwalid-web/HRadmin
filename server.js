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

// ── Leave email notifications ─────────────────────────────────────────────────
async function sendLeaveEmail(triggerKey, leaveReq, extraInfo = {}) {
  try {
    const s   = Q.getSettings();
    if (s.email_enabled !== '1' || s[triggerKey] !== '1') return; // disabled
    const cfg    = Q.getEntraConfig();
    const sender = s.email_sender;
    if (!cfg.tenant_id || !cfg.client_id || !cfg.client_secret || !sender) return;
    const fromName = s.email_from_name || 'WorkIQ HR';

    const { empEmail, empName, managerEmail, managerName, reqId, typeName, fromDate, toDate, days, status, comment } = extraInfo;

    const baseStyle = `font-family:Arial,sans-serif;max-width:520px;margin:0 auto`;
    const pill = (text, color) => `<span style="display:inline-block;padding:2px 10px;border-radius:12px;background:${color}22;color:${color};font-size:12px;font-weight:600">${text}</span>`;

    const leaveCard = `
      <div style="background:#f8f9fa;border-radius:8px;padding:16px;margin:14px 0;border-left:4px solid #3b82f6">
        <div style="font-size:13px;color:#555">Request ID: <strong>${reqId}</strong></div>
        <div style="font-size:13px;color:#555">Type: <strong>${typeName}</strong></div>
        <div style="font-size:13px;color:#555">Period: <strong>${fromDate} → ${toDate}</strong> (${days} day${days>1?'s':''})</div>
      </div>`;

    let subject, body, recipients = [];

    if (triggerKey === 'email_on_submit') {
      // → Employee gets confirmation; manager gets pending notification
      subject = `Leave request ${reqId} submitted — awaiting approval`;
      const empBody = `<div style="${baseStyle}">
        <h2 style="color:#1e293b">Leave request submitted ✓</h2>
        <p>Hi ${empName},</p>
        <p>Your leave request has been submitted and is pending manager approval.</p>
        ${leaveCard}
        <p style="color:#64748b;font-size:12px">You will be notified at each approval step.<br>— ${fromName}</p>
      </div>`;
      const mgrBody = `<div style="${baseStyle}">
        <h2 style="color:#1e293b">Action required: Leave approval</h2>
        <p>Hi ${managerName||'Manager'},</p>
        <p><strong>${empName}</strong> has submitted a leave request that requires your approval.</p>
        ${leaveCard}
        <p>Please log in to <strong>WorkIQ</strong> to approve or reject this request.</p>
        <p style="color:#64748b;font-size:12px">— ${fromName}</p>
      </div>`;
      if (empEmail)     await sendGraphMail(cfg, sender, empEmail,     subject, empBody).catch(()=>{});
      if (managerEmail) await sendGraphMail(cfg, sender, managerEmail, `Action required: ${empName} leave approval`, mgrBody).catch(()=>{});
      return;
    }

    if (triggerKey === 'email_on_approve') {
      const isFinal = status === 'approved';
      subject = isFinal ? `Leave request ${reqId} fully approved ✓` : `Leave request ${reqId} — next approval pending`;
      body = `<div style="${baseStyle}">
        <h2 style="color:#22c55e">Leave request ${isFinal ? 'approved ✓' : 'progressing...'}</h2>
        <p>Hi ${empName},</p>
        <p>Your leave request has been <strong>${isFinal ? 'fully approved' : 'approved at one level and is awaiting the next approver'}</strong>.</p>
        ${leaveCard}
        ${comment ? `<p>Comment: <em>${comment}</em></p>` : ''}
        <p style="color:#64748b;font-size:12px">— ${fromName}</p>
      </div>`;
      recipients = [empEmail];
    }

    if (triggerKey === 'email_on_reject') {
      subject = `Leave request ${reqId} rejected`;
      body = `<div style="${baseStyle}">
        <h2 style="color:#ef4444">Leave request rejected</h2>
        <p>Hi ${empName},</p>
        <p>Your leave request has been <strong>rejected</strong>.</p>
        ${leaveCard}
        ${comment ? `<p>Reason: <em>${comment}</em></p>` : ''}
        <p>Please contact HR if you have questions.</p>
        <p style="color:#64748b;font-size:12px">— ${fromName}</p>
      </div>`;
      recipients = [empEmail];
    }

    for (const to of recipients.filter(Boolean)) {
      await sendGraphMail(cfg, sender, to, subject, body).catch(() => {});
    }
  } catch (e) {
    console.error('[WorkIQ] Email notification error:', e.message);
  }
}

// Helper: build email context from a leave request ID
function getLeaveEmailContext(reqId) {
  const req = db.prepare(`
    SELECT lr.*, e.name as emp_name, e.email as emp_email,
           e.manager_id, lt.name as type_name
    FROM leave_requests lr
    JOIN employees e ON lr.employee_id=e.id
    JOIN leave_types lt ON lr.leave_type_id=lt.id
    WHERE lr.id=?`).get(reqId);
  if (!req) return null;
  let managerEmail = null, managerName = null;
  if (req.manager_id) {
    const mgr = db.prepare('SELECT name, email FROM employees WHERE id=?').get(req.manager_id);
    managerEmail = mgr?.email;
    managerName  = mgr?.name;
  } else {
    // Fallback: find manager by dept
    const mgr = db.prepare("SELECT name, email FROM employees WHERE dept=(SELECT dept FROM employees WHERE id=?) AND role='manager' AND active=1 LIMIT 1").get(req.employee_id);
    managerEmail = mgr?.email;
    managerName  = mgr?.name;
  }
  return {
    reqId: req.id, empEmail: req.emp_email, empName: req.emp_name,
    managerEmail, managerName,
    typeName: req.type_name, fromDate: req.from_date, toDate: req.to_date,
    days: req.days, status: req.status,
  };
}
async function getClientCredToken(cfg) {
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: cfg.client_id,
    client_secret: cfg.client_secret,
    scope: 'https://graph.microsoft.com/.default',
  }).toString();
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
    r.setTimeout(15000, () => { r.destroy(); reject(new Error('Token request timeout')); });
    r.write(body); r.end();
  });
}

async function sendGraphMail(cfg, senderUpn, toEmail, subject, htmlBody) {
  if (!cfg?.tenant_id || !cfg?.client_id || !cfg?.client_secret)
    throw new Error('Entra ID credentials not configured');
  if (!senderUpn) throw new Error('Sender UPN not configured (set in Email settings)');

  const tokenResp = await getClientCredToken(cfg);
  if (!tokenResp.access_token)
    throw new Error(tokenResp.error_description || tokenResp.error || 'Failed to get access token');

  const s    = Q.getSettings();
  const fromName = s.email_from_name || 'WorkIQ HR';
  const payload  = JSON.stringify({
    message: {
      subject,
      body: { contentType: 'HTML', content: htmlBody },
      toRecipients: [{ emailAddress: { address: toEmail } }],
      from: { emailAddress: { address: senderUpn, name: fromName } },
    },
    saveToSentItems: false,
  });

  return new Promise((resolve, reject) => {
    const opts = {
      hostname: 'graph.microsoft.com',
      path: `/v1.0/users/${encodeURIComponent(senderUpn)}/sendMail`,
      method: 'POST',
      headers: {
        Authorization: `Bearer ${tokenResp.access_token}`,
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(payload),
      },
    };
    let resp = '';
    const r = https.request(opts, res => {
      res.on('data', d => resp += d);
      res.on('end', () => {
        // 202 Accepted = success, no body
        if (res.statusCode === 202) return resolve({ ok: true });
        try {
          const data = JSON.parse(resp);
          reject(new Error(data?.error?.message || `Graph sendMail HTTP ${res.statusCode}`));
        } catch {
          reject(new Error(`Graph sendMail HTTP ${res.statusCode}`));
        }
      });
    });
    r.on('error', reject);
    r.setTimeout(20000, () => { r.destroy(); reject(new Error('sendMail timeout')); });
    r.write(payload);
    r.end();
  });
}

const MIME = { '.html':'text/html','.css':'text/css','.js':'application/javascript','.json':'application/json','.ico':'image/x-icon','.png':'image/png' };

// Gender name sets — built once at startup, used by /api/hr-insights
const GENDER_FEMALE = new Set(['ghaida','houwayda','kana','mariko','mayssa','ons','molka','nanya','rahma','raoudha','salwa','sarra','semah','shalaka','sondos','taeko','yassmine','yuuna','amira','amina','amani','amel','asma','beya','chahira','dhouha','emna','fatima','fatma','ghada','hajer','hana','houda','ikram','imen','ines','jihen','khaoula','latifa','leila','lina','mabrouka','maissa','mariem','meriem','mouna','nadia','najla','narjes','nesrine','nour','doniez','dorsaf','eya','nouha','olfa','rania','rim','rima','safa','sahar','salma','sana','sara','sarah','sirine','sonia','wafa','wiem','yasmine','yousra','zahra','zaineb','zina','zohra','hind','nora','noura','maha','dina','randa','rana','jihan','heba','dalia','maya','mira','luna','layla','laila','nada','mona','hala','rita','roula','samira','siham','soumaya','souad','rajaa','kaoutar','imane','ilham','hanane','hasna','ghizlane','fadwa','chaima','assia','asmaa','aisha','aicha','aida','zara','zineb','widad','oumaima','nabila','mounya','loubna','lilia','khadija','jihane','jamila','hiba','fatna','farida','fatiha','dounia','dalila','chama','bouchra','asmae','priya','maria','emma','julia','diana','nina','elena','anna','lisa','sophie','claire','alice','julie','marie','camille','lea','chloe','lucie','manon','irina','olga','natasha','svetlana','jennifer','jessica','ashley','amanda','megan','lauren','kelly','rachel','samantha','nicole','aiko','akane','akemi','aki','akiko','amane','ami','asahi','asuka','ayaka','ayako','ayame','ayane','ayano','ayumi','azusa','chie','chieko','chika','chikako','chisato','chizu','eiko','emi','emiko','eri','erika','eriko','fujiko','fumi','fumie','fumiko','hanako','hayami','hikari','hikaru','hina','hinako','hiroko','hisako','hitomi','honami','ichika','itsuki','izumi','junko','kaede','kanae','kanako','kanon','kaori','kaoru','kasumi','kazue','kazuko','keiko','kiko','kimiko','kirara','kiriko','kiyomi','koharu','kotomi','kotone','kumiko','kumi','kuniko','kurumi','kyoko','machiko','madoka','mai','maiko','makiko','mami','mamiko','manami','mao','masako','matsuri','mayuko','megumi','miharu','mihoko','mika','mikako','mikoto','miku','minako','minori','mirei','misaki','misato','mitsuki','miwa','miyako','miyuki','mizuki','moe','moeko','momoka','momoko','nana','nanami','naoko','naomi','natsuki','natsuko','natsumi','nobuko','nodoka','noriko','reika','reiko','reina','rena','rina','rio','rion','risa','sachiko','saki','sakura','satsuki','sayaka','sayuri','setsuko','shiho','shika','shizuka','sumire','suzume','tamako','tamami','teruko','tomoe','tomoko','tomomi','tsukiko','tsukimi','tsukasa','umi','wakana','waka','yoko','yoriko','yoshiko','yukako','yukari','yuki','yukiko','yuko','yumi','yumiko','yuna','yuri','yurika','yurina','yuriko','yuuko']);
const GENDER_MALE   = new Set(['adem','alex','amrikar','bilel','david','duchaufour','elyes','feres','fernando','firas','gaith','ghazi','hachem','hamdy','hassen','haykel','hedi','hesham','hichem','hiroyuki','houssem','imed','islem','iyed','jalel','james','jaouher','jawaad','kal','kang','karam','kerem','kido','kishore','laroussi','layth','louay','luis','mabrouki','marouen','marwen','matthew','med','medamine','moetez','mohamedali','mones','morita','mouhamed','nadhir','nadim','nicholas','nidhal','noorallah','radhouan','rajshekhar','rayen','razi','rchid','saber','salim','sami','seifallah','seifeddine','skander','slim','sofiene','sousuke','syed','tarek','tetsumin','tetsuya','tom','toshiharu','wael','wahib','walid','wassim','woodson','xavier','yanseiji','yassine','yessin','youssef','zied','abdelhak','abdelaziz','abdelkader','abdelmalek','abderrahim','abderrahmane','abdessalam','adil','adnane','adnan','ahmed','ali','amine','anter','amenallah','amir','anis','aymen','aziz','achraf','achref','belhassen','bilal','brahim','chakib','driss','daly','dhia','farid','hamid','hamza','hani','hassan','hatim','houssam','iliyas','ilyes','iskander','ismail','issa','jawad','kamal','kamel','karim','khaled','khalid','latif','mahdi','marouane','mehdi','mohamed','mohammed','mouad','mourad','moussa','mustapha','nabil','nassim','nizar','omar','oussama','rachid','rafik','rami','redouane','riadh','saad','samir','slim','soufiane','tarak','tariq','tawfiq','yacine','yasser','zakaria','anas','fares','ghassen','hatem','helmi','jaber','jamel','jawher','kaies','khalil','larbi','maher','malek','mansour','mohsen','mondher','mongi','mounir','nader','oualid','ramzi','ridha','salah','samy','sayed','selim','souheil','taher','thameur','wajdi','wissem','yousri','zouhaier','michael','john','robert','william','daniel','mark','paul','peter','thomas','andrew','george','kevin','brian','steven','edward','richard','charles','joseph','christopher','anthony','donald','kenneth','joshua','timothy','ryan','eric','jason','jeffrey','frank','gary','stephen','patrick','raymond','scott','jack','dennis','walter','henry','arthur','joe','juan','carlos','miguel','alejandro','akihiko','akihiro','akio','akira','akito','atsushi','daichi','daiki','daisuke','fumihiro','fumio','genki','goro','hideo','hiroki','hiroshi','hiroto','hiroya','hisashi','hisato','isamu','isao','jiro','junichi','junpei','junsuke','katsuhiko','katsuhiro','katsuya','kazuki','kazuma','kazuya','keigo','keiji','keiki','keisuke','kenji','kenta','kentaro','kenzo','koji','kotaro','koya','kuniaki','kunio','makoto','masahiro','masaki','masanobu','masaru','masashi','masato','masayuki','minoru','naoki','naoto','naoya','noboru','nobuhiro','nobuo','nobuyuki','noriaki','norio','noriyuki','osamu','raito','reiji','reito','ren','riku','rinto','ryo','ryohei','ryoji','ryosuke','ryota','ryoto','ryuichi','ryuji','ryunosuke','satoru','satoshi','seiji','seiya','shingo','shinji','shinsuke','shintaro','shiro','shota','shuhei','shun','shunsuke','soichiro','sota','subaru','suguru','taichi','taiga','taiki','taishi','taito','takahiro','takashi','takato','takayuki','takehiro','takeshi','takuya','takuma','takuro','tomohiro','tomoki','tomoya','toshiaki','toshihiko','toshihiro','toshiki','toshio','toshiyuki','yasuhiro','yasuo','yasushi','yasuyuki','yoshihiko','yoshihiro','yoshiki','yoshio','yoshitaka','yoshito','yudai','yuichi','yuji','yukihiro','yukio','yukito','yusuke','yuto','yuya','yuuki','yuuma','hayato','haruki','rento','haruto','sosuke','hinata','kohei','issei']);

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

  // ── Email settings (Microsoft Graph / Azure service principal) ───────────────
  if (pathname === '/api/email-settings') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    if (method === 'GET') {
      const all = Q.getSettings();
      const cfg  = Q.getEntraConfig();
      return json(res, 200, {
        // Graph-based email settings
        email_enabled:    all.email_enabled    || '0',
        email_sender:     all.email_sender     || '',   // UPN of the mailbox to send from
        email_from_name:  all.email_from_name  || 'WorkIQ HR',
        // Show which Entra creds will be used (read-only reference)
        entra_tenant:     cfg.tenant_id  ? cfg.tenant_id.slice(0,8)+'...' : '',
        entra_client:     cfg.client_id  ? cfg.client_id.slice(0,8)+'...' : '',
        entra_configured: !!(cfg.tenant_id && cfg.client_id && cfg.client_secret),
        // Trigger toggles
        email_on_submit:   all.email_on_submit   || '1',
        email_on_approve:  all.email_on_approve  || '1',
        email_on_reject:   all.email_on_reject   || '1',
        email_on_pending:  all.email_on_pending  || '1',
        email_reminder:    all.email_reminder    || '1',
        email_monthly_bal: all.email_monthly_bal || '1',
        email_on_sync:     all.email_on_sync     || '0',
      });
    }
    if (method === 'POST') {
      const body = await readBody(req);
      Q.setSettings(body, emp.name);
      return json(res, 200, { ok: true });
    }
  }

  // ── Test email via Microsoft Graph ────────────────────────────────────────────
  if (pathname === '/api/email-settings/test' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr')) return;
    const s   = Q.getSettings();
    const cfg = Q.getEntraConfig();
    if (!cfg.tenant_id || !cfg.client_id || !cfg.client_secret)
      return json(res, 400, { error: 'Entra ID credentials not configured — go to Entra ID config first' });
    const sender = s.email_sender || emp.email;
    if (!sender)
      return json(res, 400, { error: 'No sender email address. Set "Send from" mailbox in Email settings.' });
    try {
      const result = await sendGraphMail(cfg, sender, emp.email, `WorkIQ Email Test — ${new Date().toLocaleString()}`,
        `<p>This is a test email from <strong>WorkIQ HR Platform</strong>.</p>
         <p>Sent via Microsoft Graph API using your Azure service principal.</p>
         <p style="color:#888;font-size:12px">Sent at ${new Date().toISOString()}</p>`);
      db.prepare('INSERT INTO audit_log (actor_name,action,target,detail) VALUES (?,?,?,?)').run(emp.name, 'Email test sent', sender, `to ${emp.email}`);
      return json(res, 200, { ok: true, message: `Test email sent to ${emp.email} via Microsoft Graph` });
    } catch(e) {
      return json(res, 400, { error: `Graph sendMail failed: ${e.message}` });
    }
  }

  // ── Send email endpoint (internal use) ────────────────────────────────────────
  if (pathname === '/api/send-email' && method === 'POST') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr', 'director')) return;
    const body = await readBody(req);
    const cfg  = Q.getEntraConfig();
    const s    = Q.getSettings();
    if (s.email_enabled !== '1') return json(res, 400, { error: 'Email notifications are disabled' });
    const sender = s.email_sender || '';
    if (!sender) return json(res, 400, { error: 'Sender email not configured' });
    try {
      await sendGraphMail(cfg, sender, body.to, body.subject, body.html || body.text);
      return json(res, 200, { ok: true });
    } catch(e) { return json(res, 400, { error: e.message }); }
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
      if (emp.role === 'employee') {
        filters.employeeId = emp.id;
      }
      // Manager: only direct reports (with dept fallback)
      if (emp.role === 'manager') {
        const mgrRecord = db.prepare('SELECT dept FROM employees WHERE id=?').get(emp.id);
        filters.managerId   = emp.id;
        filters.managerDept = mgrRecord?.dept || null;
      }
      // HR and director: all (no filter added)
      if (parsed.query.status) filters.status = parsed.query.status;
      if (parsed.query.dept)   filters.dept   = parsed.query.dept;
      if (parsed.query.team)   filters.team   = parsed.query.team;
      return json(res, 200, Q.getLeaveRequests(filters));
    }
    if (method === 'POST') {
      const body = await readBody(req);
      const id = Q.submitLeaveRequest(emp.id, body, emp.name);
      // Send email notifications (non-blocking)
      const ctx = getLeaveEmailContext(id);
      if (ctx) sendLeaveEmail('email_on_submit', null, ctx);
      return json(res, 200, { ok: true, id });
    }
  }
  if (/^\/api\/leave-requests\/([^/]+)\/(approve|reject)$/.test(pathname)) {
    const parts = pathname.split('/');
    const reqId = parts[3], action = parts[4];
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'manager','hr','director')) return;

    // Manager scope guard: manager can only approve their own direct reports' requests
    if (emp.role === 'manager') {
      const leaveReq = db.prepare('SELECT employee_id FROM leave_requests WHERE id=?').get(reqId);
      if (leaveReq) {
        const directReports = db.prepare('SELECT id FROM employees WHERE manager_id=? AND active=1').all(emp.id).map(r=>r.id);
        // Fallback: dept match if no direct reports assigned
        const inScope = directReports.length > 0
          ? directReports.includes(leaveReq.employee_id)
          : (() => {
              const mgrDept = db.prepare('SELECT dept FROM employees WHERE id=?').get(emp.id)?.dept;
              const empDept = db.prepare('SELECT dept FROM employees WHERE id=?').get(leaveReq.employee_id)?.dept;
              return mgrDept && mgrDept === empDept;
            })();
        if (!inScope) return json(res, 403, { error: 'This employee is not in your team' });
      }
    }

    const body = await readBody(req);
    try {
      if (action === 'approve') Q.approveLeaveRequest(reqId, emp.id, body.comment, emp.name);
      else Q.rejectLeaveRequest(reqId, emp.id, body.reason, emp.name);
      // Send email after status is updated in DB (non-blocking)
      const ctx = getLeaveEmailContext(reqId);
      if (ctx) {
        ctx.comment = body.comment || body.reason || '';
        const triggerKey = action === 'approve' ? 'email_on_approve' : 'email_on_reject';
        sendLeaveEmail(triggerKey, null, ctx);
      }
      return json(res, 200, { ok: true });
    } catch (e) { return json(res, 400, { error: e.message }); }
  }

  // ── HR Insights ───────────────────────────────────────────────────────────────
  if (pathname === '/api/hr-insights') {
    const emp = requireAuth(req, res); if (!emp) return;
    if (!requireRole(emp, res, 'hr', 'director')) return;
    const today = new Date().toISOString().slice(0, 10);
    const year  = new Date().getFullYear();

    // Base filter — exclude employees with no company assigned
    const BASE = "active=1 AND company IS NOT NULL AND company != ''";

    // Headcount (company-assigned only)
    const total     = db.prepare(`SELECT COUNT(*) as c FROM employees WHERE ${BASE}`).get().c;
    const excluded  = db.prepare("SELECT COUNT(*) as c FROM employees WHERE active=1 AND (company IS NULL OR company='')").get().c;
    const byRole    = db.prepare(`SELECT role, COUNT(*) as c FROM employees WHERE ${BASE} GROUP BY role ORDER BY c DESC`).all();
    const byDept    = db.prepare(`SELECT dept, COUNT(*) as c FROM employees WHERE ${BASE} AND dept IS NOT NULL AND dept!='' GROUP BY dept ORDER BY c DESC`).all();
    const byTeam    = db.prepare(`SELECT team, COUNT(*) as c FROM employees WHERE ${BASE} AND team IS NOT NULL AND team!='' GROUP BY team ORDER BY c DESC`).all();

    // Team split: Interns vs non-Interns per team (for stacked bar chart)
    const byTeamSplit = db.prepare(`
      SELECT
        team,
        SUM(CASE WHEN LOWER(title) LIKE '%intern%' THEN 1 ELSE 0 END) as interns,
        SUM(CASE WHEN LOWER(title) NOT LIKE '%intern%' OR title IS NULL THEN 1 ELSE 0 END) as employees
      FROM employees
      WHERE ${BASE} AND team IS NOT NULL AND team != ''
      GROUP BY team
      ORDER BY (interns + employees) DESC
    `).all();
    const byCompany = db.prepare(`SELECT company, COUNT(*) as c FROM employees WHERE ${BASE} GROUP BY company ORDER BY c DESC`).all();
    const internCount = db.prepare(`SELECT COUNT(*) as c FROM employees WHERE ${BASE} AND LOWER(title) LIKE '%intern%'`).get().c;

    // Manager span of control (company-assigned only)
    const spans    = db.prepare(`SELECT manager_id, COUNT(*) as c FROM employees WHERE ${BASE} AND manager_id IS NOT NULL GROUP BY manager_id`).all();
    const avgSpan  = spans.length ? (spans.reduce((s,r)=>s+r.c,0)/spans.length).toFixed(1) : 0;
    const maxSpan  = spans.length ? Math.max(...spans.map(r=>r.c)) : 0;
    const unmanaged= db.prepare(`SELECT COUNT(*) as c FROM employees WHERE ${BASE} AND manager_id IS NULL AND role='employee'`).get().c;

    // Coverage gaps (company-assigned only)
    const noCal    = db.prepare(`SELECT COUNT(*) as c FROM employees WHERE ${BASE} AND (cal_id IS NULL OR cal_id='')`).get().c;
    const noCompany= excluded; // already computed above

    // Leave analytics — only for employees with a company
    const leaveByType = db.prepare(`
      SELECT lt.name, COUNT(*) as c, SUM(lr.days) as total_days
      FROM leave_requests lr
      JOIN leave_types lt ON lr.leave_type_id=lt.id
      JOIN employees e ON lr.employee_id=e.id
      WHERE lr.status='approved' AND e.company IS NOT NULL AND e.company!=''
      GROUP BY lt.id ORDER BY total_days DESC`).all();

    const leaveByDept = db.prepare(`
      SELECT e.dept, COUNT(*) as c, SUM(lr.days) as total_days
      FROM leave_requests lr
      JOIN employees e ON lr.employee_id=e.id
      WHERE lr.status='approved' AND e.dept IS NOT NULL
        AND e.company IS NOT NULL AND e.company!=''
      GROUP BY e.dept ORDER BY total_days DESC`).all();

    const onLeaveToday = db.prepare(`
      SELECT COUNT(DISTINCT lr.employee_id) as c
      FROM leave_requests lr
      JOIN employees e ON lr.employee_id=e.id
      WHERE lr.status='approved' AND lr.from_date<=? AND lr.to_date>=?
        AND e.company IS NOT NULL AND e.company!=''`).get(today, today).c;

    const pendingLeave = db.prepare(`
      SELECT COUNT(*) as c FROM leave_requests lr
      JOIN employees e ON lr.employee_id=e.id
      WHERE lr.status NOT IN ('approved','rejected')
        AND e.company IS NOT NULL AND e.company!=''`).get().c;

    const rejectedRate = db.prepare(`
      SELECT COUNT(*) as c FROM leave_requests lr
      JOIN employees e ON lr.employee_id=e.id
      WHERE lr.status='rejected' AND e.company IS NOT NULL AND e.company!=''`).get().c;

    const totalReqs = db.prepare(`
      SELECT COUNT(*) as c FROM leave_requests lr
      JOIN employees e ON lr.employee_id=e.id
      WHERE e.company IS NOT NULL AND e.company!=''`).get().c;

    // Avg leave days used per employee this year (company-assigned only)
    const avgUsed = db.prepare(`
      SELECT AVG(lb.used) as avg FROM leave_balances lb
      JOIN employees e ON lb.employee_id=e.id
      WHERE lb.year=? AND lb.used>0 AND e.company IS NOT NULL AND e.company!=''`).get(year)?.avg?.toFixed(1) || 0;

    // Monthly trend (company-assigned only)
    const monthlyTrend = db.prepare(`
      SELECT strftime('%Y-%m', created_at) as month, COUNT(*) as joined
      FROM employees WHERE ${BASE} GROUP BY month ORDER BY month`).all();

    // Monthly leave trend (company-assigned only)
    const leaveTrend = db.prepare(`
      SELECT strftime('%Y-%m', lr.created_at) as month, COUNT(*) as c
      FROM leave_requests lr
      JOIN employees e ON lr.employee_id=e.id
      WHERE e.company IS NOT NULL AND e.company!=''
      GROUP BY month ORDER BY month DESC LIMIT 12`).all().reverse();

    // Presence — only for company-assigned employees
    const presenceCounts = (() => {
      try {
        return db.prepare(`
          SELECT ep.availability, COUNT(*) as c
          FROM employee_presence ep
          JOIN employees e ON ep.employee_id=e.id
          WHERE e.company IS NOT NULL AND e.company!=''
          GROUP BY ep.availability ORDER BY c DESC`).all();
      } catch { return []; }
    })();
    const activeNow  = presenceCounts.filter(p=>['Available','Busy','InACall','InAConferenceCall','InAMeeting'].includes(p.availability)).reduce((s,r)=>s+r.c,0);
    const awayNow    = presenceCounts.filter(p=>['Away','BeRightBack'].includes(p.availability)).reduce((s,r)=>s+r.c,0);
    const offlineNow = presenceCounts.filter(p=>['Offline','PresenceUnknown','OffWork'].includes(p.availability)).reduce((s,r)=>s+r.c,0);
    const dndNow     = presenceCounts.filter(p=>['DoNotDisturb','Presenting'].includes(p.availability)).reduce((s,r)=>s+r.c,0);

    // Gender (company-assigned only)
    const genderStats = { male: 0, female: 0, unknown: 0 };
    db.prepare(`SELECT name FROM employees WHERE ${BASE}`).all().forEach(e => {
      const first = (e.name||'').trim().split(/\s+/)[0].toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
      if (GENDER_FEMALE.has(first))      genderStats.female++;
      else if (GENDER_MALE.has(first))   genderStats.male++;
      else                               genderStats.unknown++;
    });

    // Balance utilisation (company-assigned only)
    const balUtil = db.prepare(`
      SELECT AVG(CASE WHEN lb.entitled>0 THEN CAST(lb.used AS REAL)/lb.entitled*100 ELSE 0 END) as pct
      FROM leave_balances lb JOIN employees e ON lb.employee_id=e.id
      WHERE lb.year=? AND e.company IS NOT NULL AND e.company!=''`).get(year)?.pct?.toFixed(1) || 0;

    const topLeaveDesc = leaveByDept[0] || null;

    return json(res, 200, {
      headcount: { total, excluded, internCount, byRole, byDept, byTeam, byTeamSplit, byCompany, monthlyTrend },
      management: { managers: byRole.find(r=>r.role==='manager')?.c||0, avgSpan, maxSpan, unmanaged },
      coverage: { noCal, noCompany },
      leave: { byType: leaveByType, byDept: leaveByDept, onLeaveToday, pendingLeave, totalReqs, rejectedRate, avgUsed, leaveTrend, balUtil, topLeaveDesc },
      presence: { activeNow, awayNow, offlineNow, dndNow, total },
      gender: genderStats,
    });
  }


  // ── IP info — office recognition + external reputation with cache ────────────
  if (/^\/api\/ip-info\//.test(pathname)) {
    const ipAddr = decodeURIComponent(pathname.split('/')[3] || '');
    if (!ipAddr) return json(res, 400, { error: 'No IP' });

    if (!global._ipCache) global._ipCache = new Map();

    // Build FQDN→IP resolution cache (resolved once, cached)
    if (!global._fqdnCache) global._fqdnCache = new Map();

    let officeIps = [];
    try { officeIps = JSON.parse(Q.getSetting('office_ips') || '[]'); } catch {}

    // Resolve all office FQDNs to IPs (lazy, cached per entry)
    const isIp = s => /^\d{1,3}(\.\d{1,3}){3}$/.test(s);

    // Check each office entry — match by IP directly OR by resolving FQDN
    for (const office of officeIps) {
      const entries = (office.ips || []).map(s => s.trim()).filter(Boolean);
      // Direct IP match
      if (entries.includes(ipAddr)) {
        return json(res, 200, { type:'office', label:office.label, ip:ipAddr });
      }
      // FQDN resolution match
      for (const entry of entries) {
        if (!isIp(entry)) {
          // Resolve FQDN to IP(s), use cache
          let resolved = global._fqdnCache.get(entry);
          if (!resolved) {
            try {
              resolved = await new Promise((resolve, reject) => {
                require('dns').resolve4(entry, (err, addrs) => err ? reject(err) : resolve(addrs));
              });
              global._fqdnCache.set(entry, resolved);
              // Refresh FQDN cache every 5 minutes
              setTimeout(() => global._fqdnCache.delete(entry), 5 * 60 * 1000);
            } catch {
              global._fqdnCache.set(entry, []);
            }
          }
          if ((resolved || []).includes(ipAddr)) {
            return json(res, 200, { type:'office', label:`${office.label} (${entry})`, ip:ipAddr });
          }
        }
      }
    }

    if (global._ipCache.has(ipAddr)) return json(res, 200, global._ipCache.get(ipAddr));

    const isPrivate = /^(10\.|172\.(1[6-9]|2\d|3[01])\.|192\.168\.|127\.|::1)/.test(ipAddr);
    if (isPrivate) {
      const r = { type:'private', label:'Private / Internal', ip:ipAddr };
      global._ipCache.set(ipAddr, r);
      return json(res, 200, r);
    }

    try {
      const result = await new Promise((resolve, reject) => {
        const r = https.request({
          hostname:'ip-api.com',
          path:`/json/${ipAddr}?fields=status,country,countryCode,city,isp,org,proxy,hosting,query`,
          method:'GET', headers:{'User-Agent':'WorkIQ/1.0'}
        }, res => {
          let b=''; res.on('data',d=>b+=d);
          res.on('end',()=>{ try{resolve(JSON.parse(b));}catch{reject(new Error('parse'));} });
        });
        r.on('error',reject);
        r.setTimeout(5000,()=>{ r.destroy(); reject(new Error('timeout')); });
        r.end();
      });
      let risk='low', riskColor='#22c55e';
      if (result.proxy || result.hosting) { risk='medium'; riskColor='#f59e0b'; }
      if (result.proxy && result.hosting) { risk='high';   riskColor='#ef4444'; }
      const info = {
        type:'external', ip:ipAddr,
        country:result.country||'—', countryCode:result.countryCode||'',
        city:result.city||'—', isp:result.isp||'—',
        proxy:result.proxy||false, hosting:result.hosting||false,
        risk, riskColor,
        label:`${result.city||''}, ${result.country||'Unknown'}`.replace(/^, /,'')
      };
      global._ipCache.set(ipAddr, info);
      if (global._ipCache.size > 2000) global._ipCache.delete(global._ipCache.keys().next().value);
      return json(res, 200, info);
    } catch(e) {
      return json(res, 200, { type:'unknown', ip:ipAddr, label:'Lookup failed', risk:'unknown' });
    }
  }

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