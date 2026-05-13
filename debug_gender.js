'use strict';
const path = require('path');
const { db } = require(path.join(__dirname, 'db'));
const fs = require('fs');

// Load gender sets from server.js
const src = fs.readFileSync(path.join(__dirname, 'server.js'), 'utf8');
const femaleMatch = src.match(/const GENDER_FEMALE = new Set\(([\s\S]*?)\);/);
const maleMatch   = src.match(/const GENDER_MALE\s*= new Set\(([\s\S]*?)\);/);
const GENDER_FEMALE = new Set(eval(femaleMatch[1]));
const GENDER_MALE   = new Set(eval(maleMatch[1]));

const SEP  = '─'.repeat(60);
const OK   = '\x1b[32m✓\x1b[0m';
const WARN = '\x1b[33m⚠\x1b[0m';
const BOLD = s => `\x1b[1m${s}\x1b[0m`;

console.log('\n' + '═'.repeat(60));
console.log(BOLD('  Gender Unknown Employees — ' + new Date().toISOString().slice(0,10)));
console.log('═'.repeat(60));
console.log(`  Female set: ${GENDER_FEMALE.size} names | Male set: ${GENDER_MALE.size} names\n`);

const employees = db.prepare("SELECT id, name, dept, company, role FROM employees WHERE active=1 ORDER BY name").all();

const results = { male:[], female:[], unknown:[] };

employees.forEach(e => {
  const raw   = (e.name||'').trim().split(/\s+/)[0];
  const first = raw.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  if      (GENDER_FEMALE.has(first)) results.female.push({ ...e, first });
  else if (GENDER_MALE.has(first))   results.male.push({ ...e, first });
  else                               results.unknown.push({ ...e, first });
});

console.log(SEP);
console.log(BOLD(` UNKNOWN gender: ${results.unknown.length} employees`));
console.log(SEP);

if (results.unknown.length === 0) {
  console.log(`${OK} All employees have a recognised gender name!\n`);
} else {
  // Table header
  console.log(`  ${'#'.padEnd(4)} ${'First name'.padEnd(15)} ${'Full name'.padEnd(30)} ${'Dept'.padEnd(25)} ${'Company'.padEnd(15)} ${'Role'}`);
  console.log('  ' + '─'.repeat(105));
  results.unknown.forEach((e, i) => {
    const num     = String(i+1).padEnd(4);
    const first   = e.first.padEnd(15);
    const name    = (e.name||'').padEnd(30);
    const dept    = (e.dept||'—').padEnd(25);
    const company = (e.company||'—').padEnd(15);
    const role    = e.role||'—';
    console.log(`  ${num} ${first} ${name} ${dept} ${company} ${role}`);
  });
  console.log('');

  // Summary by first letter
  const byLetter = {};
  results.unknown.forEach(e => {
    const l = e.first[0]?.toUpperCase() || '?';
    byLetter[l] = (byLetter[l]||[]);
    byLetter[l].push(e.first);
  });
  console.log(SEP);
  console.log(BOLD(' Unknown first names grouped by letter (add these to server.js):'));
  console.log(SEP);
  Object.keys(byLetter).sort().forEach(letter => {
    const unique = [...new Set(byLetter[letter])].sort();
    console.log(`  ${letter}: ${unique.join(', ')}`);
  });
}

console.log('\n' + SEP);
console.log(BOLD(' SUMMARY'));
console.log(SEP);
const total = employees.length;
console.log(`  Total active employees: ${total}`);
console.log(`  ${OK} Male:    ${results.male.length} (${Math.round(results.male.length/total*100)}%)`);
console.log(`  ${OK} Female:  ${results.female.length} (${Math.round(results.female.length/total*100)}%)`);
console.log(`  ${WARN} Unknown: ${results.unknown.length} (${Math.round(results.unknown.length/total*100)}%)`);
console.log('\n' + '═'.repeat(60) + '\n');