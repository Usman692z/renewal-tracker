// Pushes edits made in the Requests tab (Sale Order, Renewal/Wasl status,
// Remarks) back into the live "New renewal sheet" in SharePoint.
// Matches rows by Plate + Mobile (the stable fields we never overwrite).
// Run: node scripts/push-to-sharepoint.js         (dry run, writes nothing)
//      node scripts/push-to-sharepoint.js --apply (actually writes to SharePoint)
const fs = require('fs');
const path = require('path');

const APPLY = process.argv.includes('--apply');

// ---- Firebase (this app's) credentials, from scripts/.env ----
const FB_API_KEY = 'AIzaSyAeKliE-8JWcYMKUwIkcqUGwwMEEWX6Oj8';
function loadEnv(file) {
  const env = {};
  if (!fs.existsSync(file)) return env;
  for (const line of fs.readFileSync(file, 'utf8').split('\n')) {
    const m = line.match(/^\s*([A-Z_]+)\s*=\s*(.+?)\s*$/);
    if (m) env[m[1]] = m[2];
  }
  return env;
}
const fbEnv = loadEnv(path.join(__dirname, '.env'));
const FB_EMAIL = fbEnv.SYNC_EMAIL;
const FB_PASSWORD = fbEnv.SYNC_PASSWORD;

// ---- Microsoft Graph (SharePoint) credentials, from the CSU pipeline's .env ----
const graphEnv = loadEnv('D:/CSU Daily Report/.env');
const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_PATH, FILE_NAME } = graphEnv;
const HOSTNAME = 'trackingco-my.sharepoint.com';
const SHEET_NAME = 'New renewal sheet';

// Column layout of the sheet (0-based), confirmed from its header row.
const COL = { date: 0, encoder: 1, mobile: 2, plate: 3, saleOrder: 10, remarks: 11, renewalStatus: 12, waslStatus: 13 };

function colLetter(i) { return String.fromCharCode(65 + i); }

async function fbSignIn() {
  const res = await fetch(`https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=${FB_API_KEY}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email: FB_EMAIL, password: FB_PASSWORD, returnSecureToken: true })
  });
  const j = await res.json();
  if (!res.ok) throw new Error('Firebase sign-in failed: ' + (j.error && j.error.message));
  return j.idToken;
}

async function fbFetchPending(idToken) {
  const res = await fetch('https://firestore.googleapis.com/v1/projects/renewal-tracker-app/databases/(default)/documents:runQuery', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + idToken },
    body: JSON.stringify({
      structuredQuery: {
        from: [{ collectionId: 'requests' }],
        where: { fieldFilter: { field: { fieldPath: 'needs_sharepoint_sync' }, op: 'EQUAL', value: { booleanValue: true } } }
      }
    })
  });
  const rows = await res.json();
  if (!res.ok) throw new Error('Firestore query failed: ' + JSON.stringify(rows).slice(0, 300));
  return rows.filter(r => r.document).map(r => {
    const id = r.document.name.split('/').pop();
    const f = r.document.fields;
    const val = (field) => (f[field] ? (f[field].stringValue ?? f[field].booleanValue ?? f[field].doubleValue) : '');
    return { id, plate: val('plate'), mobile: val('mobile'), sale_order: val('sale_order'), remarks: val('remarks'), renewal_status: val('renewal_status'), wasl_status: val('wasl_status') };
  });
}

async function fbClearFlag(idToken, docId) {
  const url = `https://firestore.googleapis.com/v1/projects/renewal-tracker-app/databases/(default)/documents/requests/${docId}?updateMask.fieldPaths=needs_sharepoint_sync`;
  const res = await fetch(url, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + idToken },
    body: JSON.stringify({ fields: { needs_sharepoint_sync: { booleanValue: false } } })
  });
  if (!res.ok) throw new Error('Failed to clear flag for ' + docId + ': ' + (await res.text()).slice(0, 200));
}

async function graphToken() {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' })
  });
  const j = await res.json();
  if (!res.ok) throw new Error('Graph auth failed: ' + JSON.stringify(j));
  return j.access_token;
}

async function getSiteAndFileIds(token) {
  const siteRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${HOSTNAME}:${SITE_PATH}`, { headers: { Authorization: 'Bearer ' + token } });
  const site = await siteRes.json();
  if (!siteRes.ok) throw new Error('Site lookup failed: ' + JSON.stringify(site));
  const safeName = FILE_NAME.replace(/'/g, '');
  const searchRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${site.id}/drive/root/search(q='${safeName}')`, { headers: { Authorization: 'Bearer ' + token } });
  const items = (await searchRes.json()).value;
  if (!items || !items.length) throw new Error('File not found');
  return { siteId: site.id, itemId: items[0].id };
}

async function getUsedRange(token, siteId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/workbook/worksheets('${SHEET_NAME}')/usedRange`;
  const res = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
  const j = await res.json();
  if (!res.ok) throw new Error('usedRange fetch failed: ' + JSON.stringify(j));
  return j.values;
}

function findRow(values, plate, mobile) {
  const matches = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[COL.plate] ?? '').trim() === String(plate).trim() &&
        String(row[COL.mobile] ?? '').trim() === String(mobile).trim()) {
      matches.push(i);
    }
  }
  return matches;
}

async function main() {
  if (!FB_EMAIL || !FB_PASSWORD) { console.error('Missing scripts/.env (SYNC_EMAIL/SYNC_PASSWORD).'); process.exit(1); }
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) { console.error('Missing D:/CSU Daily Report/.env Graph credentials.'); process.exit(1); }

  console.log(APPLY ? '=== LIVE MODE: will write to SharePoint ===' : '=== DRY RUN: no writes will be made (pass --apply to write) ===');

  const idToken = await fbSignIn();
  const pending = await fbFetchPending(idToken);
  console.log(`${pending.length} request(s) marked needs_sharepoint_sync`);
  if (!pending.length) return;

  const graphTok = await graphToken();
  const { siteId, itemId } = await getSiteAndFileIds(graphTok);
  const values = await getUsedRange(graphTok, siteId, itemId);
  console.log(`fetched live sheet, ${values.length} rows (incl. header)`);

  for (const p of pending) {
    const matches = findRow(values, p.plate, p.mobile);
    if (matches.length !== 1) {
      console.log(`SKIP doc ${p.id}: found ${matches.length} matching row(s) for plate="${p.plate}" mobile="${p.mobile}" (need exactly 1)`);
      continue;
    }
    const rowIndex = matches[0]; // 0-based index into `values`, row 1 in Excel == index 0
    const excelRow = rowIndex + 1;
    const rangeAddr = `${colLetter(COL.saleOrder)}${excelRow}:${colLetter(COL.waslStatus)}${excelRow}`;
    const newValues = [p.sale_order, p.remarks, p.renewal_status, p.wasl_status];
    console.log(`${APPLY ? 'WRITE' : 'WOULD WRITE'} doc ${p.id} -> sheet row ${excelRow} (${rangeAddr}):`, JSON.stringify(newValues));

    if (APPLY) {
      const patchUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/workbook/worksheets('${SHEET_NAME}')/range(address='${rangeAddr}')`;
      const patchRes = await fetch(patchUrl, {
        method: 'PATCH',
        headers: { Authorization: 'Bearer ' + graphTok, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: [newValues] })
      });
      if (!patchRes.ok) { console.log('  FAILED:', (await patchRes.text()).slice(0, 300)); continue; }
      await fbClearFlag(idToken, p.id);
      console.log('  done, flag cleared.');
    }
  }
  console.log(APPLY ? 'All done.' : 'Dry run complete. Re-run with --apply to actually write these changes.');
}

main().catch(err => { console.error(err.message); process.exit(1); });
