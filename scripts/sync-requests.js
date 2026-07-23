// Sync renewal requests from the local CSU Daily Report pipeline output
// (D:\CSU Daily Report\dashboard_data.json) into Firestore `requests`.
// Signs in with a normal app account whose credentials live in scripts/.env
// (SYNC_EMAIL / SYNC_PASSWORD) so writes pass the auth-only security rules.
// Run: node scripts/sync-requests.js   (or use Sync Requests.bat)
const fs = require('fs');
const path = require('path');

const API_KEY = 'AIzaSyAeKliE-8JWcYMKUwIkcqUGwwMEEWX6Oj8';
const PROJECT_ID = 'renewal-tracker-app';
const DATA_FILE = process.env.DATA_FILE || 'D:/CSU Daily Report/dashboard_data.json';
const CONCURRENCY = 15;

// minimal .env loader (no dependencies)
const envPath = path.join(__dirname, '.env');
if (fs.existsSync(envPath)) {
  for (const line of fs.readFileSync(envPath, 'utf8').split('\n')) {
    const m = line.match(/^\s*([A-Z_]+)\s*=\s*(.+?)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2];
  }
}

const EMAIL = process.env.SYNC_EMAIL;
const PASSWORD = process.env.SYNC_PASSWORD;
if (!EMAIL || !PASSWORD) {
  console.error('Missing SYNC_EMAIL / SYNC_PASSWORD. Copy scripts/.env.example to scripts/.env and fill it in.');
  process.exit(1);
}

function docId(r, index) {
  // Index-based to guarantee uniqueness: blank/continuation rows in the sheet
  // can share identical date+plate+sale_order+mobile, which previously made
  // content-hash IDs collide and silently overwrite each other (data loss).
  return 'row-' + String(index).padStart(5, '0');
}

function toFsFields(r) {
  const f = {};
  for (const [k, v] of Object.entries(r)) {
    if (typeof v === 'boolean') f[k] = { booleanValue: v };
    else if (typeof v === 'number') f[k] = { doubleValue: v };
    else f[k] = { stringValue: String(v ?? '') };
  }
  f.synced_at = { timestampValue: new Date().toISOString() };
  return f;
}

async function main() {
  const raw = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
  const records = raw.records || [];
  console.log(`loaded ${records.length} records from ${DATA_FILE} (generated_at: ${raw.generated_at})`);

  console.log('signing in...');
  const authRes = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${API_KEY}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email: EMAIL, password: PASSWORD, returnSecureToken: true })
  });
  const auth = await authRes.json();
  if (!authRes.ok) {
    console.error('sign-in failed:', auth.error && auth.error.message);
    process.exit(1);
  }
  console.log('signed in as', auth.email);

  // Don't overwrite rows the team edited in the app but hasn't pushed to
  // SharePoint yet (needs_sharepoint_sync=true) — the sheet doesn't have
  // those edits, so syncing over them would silently discard the edits.
  const pendingRes = await fetch(`https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents:runQuery`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${auth.idToken}` },
    body: JSON.stringify({
      structuredQuery: {
        from: [{ collectionId: 'requests' }],
        where: { fieldFilter: { field: { fieldPath: 'needs_sharepoint_sync' }, op: 'EQUAL', value: { booleanValue: true } } }
      }
    })
  });
  const pendingRows = await pendingRes.json();
  if (!pendingRes.ok) { console.error('pending-edit query failed:', JSON.stringify(pendingRows).slice(0, 200)); process.exit(1); }
  const skipIds = new Set(pendingRows.filter(r => r.document).map(r => r.document.name.split('/').pop()));
  if (skipIds.size) console.log(`skipping ${skipIds.size} row(s) with unpushed app edits (run "Push to SharePoint.bat" to push them)`);

  async function upsertOne(r, index) {
    if (skipIds.has(docId(r, index))) return;
    const url = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/requests/${docId(r, index)}`;
    const res = await fetch(url, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${auth.idToken}` },
      body: JSON.stringify({ fields: toFsFields(r) })
    });
    if (!res.ok) throw new Error(`${res.status}: ${(await res.text()).slice(0, 200)}`);
  }

  let written = 0;
  let failed = 0;
  let firstError = null;
  for (let i = 0; i < records.length; i += CONCURRENCY) {
    const slice = records.slice(i, i + CONCURRENCY);
    const results = await Promise.allSettled(slice.map((r, j) => upsertOne(r, i + j)));
    for (const r of results) {
      if (r.status === 'fulfilled') written++;
      else { failed++; if (!firstError) firstError = r.reason.message; }
    }
    process.stdout.write(`\rsynced ${written}/${records.length}${failed ? ' (' + failed + ' failed)' : ''}`);
  }
  console.log('\ndone.' + (failed ? ` ${failed} record(s) failed. First error: ${firstError}` : ''));
}

main().catch(err => { console.error(err.message); process.exit(1); });
