// ═══════════════════════════════════════════════════
//  SERVICE WORKER — Daily Material Used background sync
//
//  This lets queued Daily Material Used submissions sync
//  to Airtable + Excel even if the tab/app was fully closed
//  after going offline. Only Android/Chrome (and desktop
//  Chrome/Edge) support the Background Sync event that
//  wakes this up automatically; iOS Safari and Chrome-on-iOS
//  are both WebKit under the hood and do not support it, so
//  for those devices the page itself retries as soon as it's
//  reopened/foregrounded (see the visibilitychange/pageshow
//  handling in index.html) — this file changes nothing for
//  them, it's a bonus path for the devices that support it.
// ═══════════════════════════════════════════════════

const AT_TOKEN = 'patbyyxDeG1vyXDt1.bf541626ff51d7ff78db16476566f413b3fd3159d7670838cb27aa2527f754c7';
const AT_BASE  = 'appLWbPBLepjoO6F4';
const AT_API   = 'https://api.airtable.com/v0';

const PA_SSR_WEBHOOK = 'https://default32a9c90e5bf140a39c28672bbed3bb.f7.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/6c7551321b01416c8d7a0b6127f58fb0/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=7ZhHQogCqtp7a6gp-90tXAiP4whoyVyKyZMVLBXYvL4';
const PA_SJR_WEBHOOK = 'https://default32a9c90e5bf140a39c28672bbed3bb.f7.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/7e65653d80fa42bea5045d2cb94c4080/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=30TUImpQ82nwDumYglJsEOl6Zn3wRlkb3kmYtyxlS4A';

const TBL = {
  materialsUsed:            'tblgfw9MvYDAtZEHz',
  materialsUsedLines:       'tblxktgCxv6few93n',
  subcontractorStockReport: 'tblAkiFxWxDYV2Alm',
  subcontractorJobReport:   'tbl5pCmx7a4tpNCNk',
  representatives:          'tblutA51et1wwkEVA',
  materials:                'tbldLNJHOmXy6MdRa', // Materials Quantities table
};

const FLD = {
  mu_company:   'fld7hG8n3HpRMA4r9',
  mu_rep:       'fldxfXXZuYzZt0FTz',
  mu_email:     'fld89I88Qttc5cRLF',
  mu_mobile:    'fldMzcCEAzo92W1Wx',
  mu_date:      'fldtSi37Fokls6K4O',
  mu_lineCount: 'fldzpeGI8LCVazURo',
};

// ── Airtable helpers (same behavior as the page's copies) ──
async function atFetchAll(tableId, body) {
  let records = [], offset = null;
  do {
    const payload = { ...body };
    if (offset) payload.offset = offset;
    const res = await fetch(`${AT_API}/${AT_BASE}/${tableId}/listRecords`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${AT_TOKEN}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) throw new Error(`AT listRecords ${res.status}: ${await res.text()}`);
    const data = await res.json();
    records = records.concat(data.records || []);
    offset = data.offset || null;
  } while (offset);
  return records;
}
const atListRecords = atFetchAll;

async function atCreate(tableId, fields, retries = 3) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const res = await fetch(`${AT_API}/${AT_BASE}/${tableId}`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${AT_TOKEN}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields }),
      });
      if (res.ok) return res.json();
      const errText = await res.text();
      if (res.status >= 400 && res.status < 500) throw new Error(`AT create ${res.status}: ${errText}`);
      if (attempt === retries) throw new Error(`AT create ${res.status}: ${errText}`);
      await new Promise(r => setTimeout(r, 2000 * attempt));
    } catch (err) {
      if (attempt === retries || (err.message && err.message.startsWith('AT create 4'))) throw err;
      await new Promise(r => setTimeout(r, 2000 * attempt));
    }
  }
}

async function atUpdate(tableId, recordId, fields, retries = 3) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const res = await fetch(`${AT_API}/${AT_BASE}/${tableId}/${recordId}`, {
        method: 'PATCH',
        headers: { Authorization: `Bearer ${AT_TOKEN}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ fields }),
      });
      if (res.ok) return res.json();
      const errText = await res.text();
      if (res.status >= 400 && res.status < 500) throw new Error(`AT update ${res.status}: ${errText}`);
      if (attempt === retries) throw new Error(`AT update ${res.status}: ${errText}`);
      await new Promise(r => setTimeout(r, 2000 * attempt));
    } catch (err) {
      if (attempt === retries || (err.message && err.message.startsWith('AT update 4'))) throw err;
      await new Promise(r => setTimeout(r, 2000 * attempt));
    }
  }
}

async function atCreateBatch(tableId, fieldsArray, retries = 3) {
  const records = fieldsArray.map(fields => ({ fields }));
  for (let attempt = 1; attempt <= retries; attempt++) {
    const res = await fetch(`${AT_API}/${AT_BASE}/${tableId}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${AT_TOKEN}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ records }),
    });
    if (res.ok) { const data = await res.json(); return data.records; }
    const errText = await res.text();
    if (res.status >= 400 && res.status < 500) throw new Error(`AT createBatch ${res.status}: ${errText}`);
    if (attempt === retries) throw new Error(`AT createBatch ${res.status}: ${errText}`);
    await new Promise(r => setTimeout(r, 2000 * attempt));
  }
}

// ── Excel/Power Automate pushes (mirrors the page's sendSSRBatchToPA / sendSJRBatchToPA) ──
async function sendSSRBatchToPA(repRecId) {
  if (!PA_SSR_WEBHOOK) return false;
  try {
    const resp = await fetch(`${AT_API}/${AT_BASE}/${TBL.subcontractorStockReport}/listRecords`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${AT_TOKEN}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filterByFormula: `FIND("${repRecId}", ARRAYJOIN({Rep Record ID (from Representative)}))`,
        fields: [
          'fldzxd5Odt0mpzDW8', 'fldc33IYWc8kgLdmN', 'fld0Cyg9IBXd8XXEH', 'fldrbeZkBMPwpaOxg',
          'fldPpj5hGp4lbQugL', 'fldAZUYVyQ5gLWKWk', 'fld5YH3SMorI6dUFe', 'fldkKoXhDLLQwo8np',
          'fld2r0bsneVGQIfnn', 'fldmqsBzGd05tJlk8', 'fldkeIz1IRPi8tc9j', 'fldpr4nM1K7MkLdEI',
          'fldtl8EaNWqzVVe8O',
        ],
        returnFieldsByFieldId: true,
      }),
    });
    const data = await resp.json();
    const rows = data.records || [];
    if (!rows.length) return true;

    const u = v => Array.isArray(v) ? v[0] || '' : v || '';
    const payload = rows.map(row => {
      const f = row.fields;
      return {
        reportKey: u(f['fldzxd5Odt0mpzDW8']), fullName: u(f['fldc33IYWc8kgLdmN']),
        companyName: u(f['fld0Cyg9IBXd8XXEH']), inventoryCode: u(f['fldrbeZkBMPwpaOxg']),
        opCode: u(f['fldPpj5hGp4lbQugL']), itCode: u(f['fldAZUYVyQ5gLWKWk']),
        description: u(f['fld5YH3SMorI6dUFe']), totalStockReceived: Number(f['fldkKoXhDLLQwo8np']) || 0,
        totalqtyRequested: Number(f['fld2r0bsneVGQIfnn']) || 0, totalStockUsed: Number(f['fldmqsBzGd05tJlk8']) || 0,
        totalStockReturned: Number(f['fldkeIz1IRPi8tc9j']) || 0, currentStockOnHand: Number(f['fldpr4nM1K7MkLdEI']) || 0,
        currentStockRequested: Number(f['fldtl8EaNWqzVVe8O']) || 0,
      };
    });

    const pushResp = await fetch(PA_SSR_WEBHOOK, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ records: payload }),
    });
    return pushResp.ok || pushResp.status === 202;
  } catch (err) {
    console.warn('[sw] sendSSRBatchToPA error:', err.message);
    return false;
  }
}

async function sendSJRBatchToPA(payload) {
  if (!PA_SJR_WEBHOOK) return false;
  if (!payload || !payload.length) return true;
  try {
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const sjrRes = await fetch(PA_SJR_WEBHOOK, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ records: payload }),
        });
        if (sjrRes.ok || sjrRes.status === 202) return true;
        if (sjrRes.status === 502 || sjrRes.status === 503) { await new Promise(r => setTimeout(r, 5000)); continue; }
        return false;
      } catch (fetchErr) {
        if (attempt === 3) return false;
        await new Promise(r => setTimeout(r, 5000));
      }
    }
    return false;
  } catch (err) {
    console.warn('[sw] sendSJRBatchToPA error:', err.message);
    return false;
  }
}

// ── Daily Material Used replay (mirrors the page's submitDailyMaterialUsed) ──
async function submitDailyMaterialUsed(payload) {
  const { repRecId, usedItems, fullJobCode, dmDate, dmKey } = payload;

  let mainUsedId;
  let existingMulLineIds = [];
  const dmExisting = await atFetchAll(TBL.materialsUsed, {
    filterByFormula: `{Submission Key}="${dmKey}"`,
    fields: ['Submission Key', 'Materials Used Lines'],
  });
  if (dmExisting.length > 0) {
    mainUsedId = dmExisting[0].id;
    existingMulLineIds = (dmExisting[0].fields['Materials Used Lines'] || []).map(r => r && r.id ? r.id : r);
  } else {
    const newUsedRecord = await atCreate(TBL.materialsUsed, {
      [FLD.mu_company]: [repRecId], [FLD.mu_rep]: [repRecId], [FLD.mu_email]: [repRecId],
      [FLD.mu_mobile]: [repRecId], [FLD.mu_date]: dmDate, 'Job Code': fullJobCode,
      [FLD.mu_lineCount]: usedItems.length, 'fldPlhNRGhoYuh40E': dmKey,
    });
    mainUsedId = newUsedRecord.id;
  }

  let createdUsedLineIds = [];
  let remainingItems = usedItems;
  if (existingMulLineIds.length > 0) {
    const idFormula = 'OR(' + existingMulLineIds.map(id => `RECORD_ID()="${id}"`).join(',') + ')';
    const existingMulRows = await atFetchAll(TBL.materialsUsedLines, { filterByFormula: idFormula, fields: ['Material', 'Qty Used'] }).catch(() => []);
    const alreadyCreatedMatIds = new Set(existingMulRows.map(r => (r.fields['Material'] || [])[0]).filter(Boolean));
    createdUsedLineIds = existingMulRows.map(r => ({ lineId: r.id, recId: (r.fields['Material'] || [])[0], qty: r.fields['Qty Used'] || 0 }));
    remainingItems = usedItems.filter(item => !alreadyCreatedMatIds.has(item.recId));
  }
  const mulBatchFields = remainingItems.map(item => ({
    'fldbEHOSuv028hMjD': [mainUsedId], 'fld0u8aUHdRsSIJLm': [item.recId],
    'fldGV26CqV55FYGv6': item.qty, 'flduIFkorq5wVuNNz': dmDate,
  }));
  for (let i = 0; i < mulBatchFields.length; i += 10) {
    const created = await atCreateBatch(TBL.materialsUsedLines, mulBatchFields.slice(i, i + 10));
    created.forEach((rec, idx) => createdUsedLineIds.push({ lineId: rec.id, recId: remainingItems[i + idx].recId, qty: remainingItems[i + idx].qty }));
  }

  const allSsrRows = await atListRecords(TBL.subcontractorStockReport, {
    fields: ['fldTqTbr6ysQAimLX', 'fldXm5gW03QOzBE1P', 'fldozoAjUrm312vj9'],
    returnFieldsByFieldId: true,
  });
  const ssrUpdates = {};
  const ssrCreates = [];
  createdUsedLineIds.forEach(item => {
    const ssrMatch = allSsrRows.find(row => {
      const repLookup = row.fields['fldTqTbr6ysQAimLX'] || [];
      const matLinked = (row.fields['fldXm5gW03QOzBE1P'] || []).map(r => r && r.id ? r.id : r);
      return repLookup.includes(repRecId) && matLinked.includes(item.recId);
    });
    if (ssrMatch) {
      if (!ssrUpdates[ssrMatch.id]) {
        ssrUpdates[ssrMatch.id] = { existingIds: (ssrMatch.fields['fldozoAjUrm312vj9'] || []).map(r => r && r.id ? r.id : r), newLineIds: [] };
      }
      if (!ssrUpdates[ssrMatch.id].existingIds.includes(item.lineId) && !ssrUpdates[ssrMatch.id].newLineIds.includes(item.lineId)) {
        ssrUpdates[ssrMatch.id].newLineIds.push(item.lineId);
      }
    } else {
      const existing = ssrCreates.find(c => c.matRecId === item.recId);
      if (existing) existing.lineIds.push(item.lineId);
      else ssrCreates.push({ matRecId: item.recId, lineIds: [item.lineId] });
    }
  });

  const ssrUpdateEntries = Object.entries(ssrUpdates);
  for (let i = 0; i < ssrUpdateEntries.length; i += 4) {
    await Promise.all(ssrUpdateEntries.slice(i, i + 4).map(([ssrId, { existingIds, newLineIds }]) =>
      atUpdate(TBL.subcontractorStockReport, ssrId, { 'fldozoAjUrm312vj9': [...existingIds, ...newLineIds] }).catch(() => {})
    ));
  }
  for (let i = 0; i < ssrCreates.length; i += 4) {
    await Promise.all(ssrCreates.slice(i, i + 4).map(({ matRecId, lineIds }) =>
      atCreate(TBL.subcontractorStockReport, { 'fldV1yEmErUzd4oFw': [repRecId], 'fldXm5gW03QOzBE1P': [matRecId], 'fldozoAjUrm312vj9': lineIds }).catch(() => {})
    ));
  }

  const existingSjrRows = await atListRecords(TBL.subcontractorJobReport, { fields: ['fldkCo5VFj09XALYv'], returnFieldsByFieldId: true }).catch(() => []);
  const mulIdsWithSjr = new Set();
  existingSjrRows.forEach(row => (row.fields['fldkCo5VFj09XALYv'] || []).forEach(r => mulIdsWithSjr.add(r && r.id ? r.id : r)));
  const sjrPendingItems = createdUsedLineIds.filter(item => !mulIdsWithSjr.has(item.lineId));
  for (let i = 0; i < sjrPendingItems.length; i += 4) {
    await Promise.all(sjrPendingItems.slice(i, i + 4).map(async item => {
      try {
        await atCreate(TBL.subcontractorJobReport, { 'fldOsnJEHcWjfe7fV': [repRecId], 'fldkCo5VFj09XALYv': [item.lineId] }, 1);
      } catch (err) { console.warn('[sw] SJR create failed:', err.message); }
    }));
  }

  // Build the SJR Excel payload from two small targeted lookups (rep + the
  // handful of materials involved) instead of the old full-table SJR fetch.
  // No shared CACHE here (this runs in the service worker's own context), so
  // this is as cheap as it gets without one — a couple of scoped API calls
  // instead of paginating through the entire SJR table like before.
  const repRows = await atFetchAll(TBL.representatives, {
    filterByFormula: `RECORD_ID()="${repRecId}"`,
    fields: ['Full Name', 'Company Name (from Company)', 'Email', 'Mobile'],
  }).catch(() => []);
  const repRow = repRows[0];
  const cn = repRow?.fields?.['Company Name (from Company)'];
  const sjrCompanyName = (Array.isArray(cn) ? cn[0] : cn) || '';
  const sjrFullName = repRow?.fields?.['Full Name'] || '';
  const sjrEmail    = repRow?.fields?.['Email'] || '';
  const sjrMobile   = repRow?.fields?.['Mobile'] || '';

  const uniqueMatIds = [...new Set(createdUsedLineIds.map(item => item.recId))];
  let matRows = [];
  if (uniqueMatIds.length) {
    const matFormula = 'OR(' + uniqueMatIds.map(id => `RECORD_ID()="${id}"`).join(',') + ')';
    matRows = await atFetchAll(TBL.materials, {
      filterByFormula: matFormula,
      fields: ['Inventory Code', 'OP Code', 'IT Code', 'Description'],
    }).catch(() => []);
  }
  const matById = {};
  matRows.forEach(r => { matById[r.id] = r.fields; });

  const sjrExcelPayload = createdUsedLineIds.map(item => {
    const mf = matById[item.recId] || {};
    const inventoryCode = mf['Inventory Code'] || '';
    return {
      reportKey:     `${sjrCompanyName} - ${sjrFullName} | ${inventoryCode} | ${fullJobCode}`,
      dateInstalled: dmDate,
      jobCode:       fullJobCode,
      companyName:   sjrCompanyName,
      fullName:      sjrFullName,
      email:         sjrEmail,
      mobile:        sjrMobile,
      inventoryCode,
      opCode:        mf['OP Code'] || '',
      itCode:        mf['IT Code'] || '',
      description:   mf['Description'] || '',
      qtyUsed:       item.qty || 0,
    };
  });

  // No 20s artificial wait here — by the time a background sync event fires,
  // enough real time has already passed for rollups to have settled.
  const ssrOk = await sendSSRBatchToPA(repRecId);
  const sjrOk = await sendSJRBatchToPA(sjrExcelPayload);
  if (!ssrOk || !sjrOk) {
    // Leave a lightweight retry-only record behind for next sync event
    throw new Error('Excel sync failed during background replay');
  }
}

async function replayWebhookRetry(payload) {
  const { repRecId, sjrExcelPayload } = payload;
  const ssrOk = await sendSSRBatchToPA(repRecId);
  const sjrOk = await sendSJRBatchToPA(sjrExcelPayload);
  if (!ssrOk || !sjrOk) throw new Error('Excel sync retry failed');
}

// ── IndexedDB queue access (same DB/store the page writes to) ──
const QUEUE_DB_NAME = 'mmp_offline_queue';
const QUEUE_STORE = 'submissions';

function openQueueDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(QUEUE_DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(QUEUE_STORE)) {
        const store = db.createObjectStore(QUEUE_STORE, { keyPath: 'id', autoIncrement: true });
        store.createIndex('status', 'status');
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function queueGetAll() {
  const db = await openQueueDB();
  return new Promise((resolve, reject) => {
    const req = db.transaction(QUEUE_STORE, 'readonly').objectStore(QUEUE_STORE).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function queueDelete(id) {
  const db = await openQueueDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(QUEUE_STORE, 'readwrite');
    tx.objectStore(QUEUE_STORE).delete(id);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function queueUpdate(id, changes) {
  const db = await openQueueDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(QUEUE_STORE, 'readwrite');
    const store = tx.objectStore(QUEUE_STORE);
    const getReq = store.get(id);
    getReq.onsuccess = () => {
      const rec = getReq.result;
      if (!rec) return resolve(null);
      Object.assign(rec, changes);
      const putReq = store.put(rec);
      putReq.onsuccess = () => resolve(rec);
      putReq.onerror = () => reject(putReq.error);
    };
    getReq.onerror = () => reject(getReq.error);
  });
}

const REPLAYERS = {
  dailyMaterialUsed: submitDailyMaterialUsed,
  dailyMaterialUsedWebhookRetry: replayWebhookRetry,
};

async function processQueueInWorker() {
  const items = (await queueGetAll()).filter(r => r.status === 'pending');
  for (const item of items) {
    const replay = REPLAYERS[item.formType];
    if (!replay) continue; // formType not handled here (e.g. contractorReturn) — leave for the page to retry
    try {
      await replay(item.payload);
      await queueDelete(item.id);
      console.log(`[sw] synced #${item.id} (${item.formType})`);
    } catch (err) {
      await queueUpdate(item.id, { attempts: (item.attempts || 0) + 1, lastError: err.message });
      console.warn(`[sw] retry failed #${item.id} (${item.formType}):`, err.message);
      throw err; // tells the Background Sync API to try again later
    }
  }
}

// ── Service worker lifecycle ──
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', event => event.waitUntil(self.clients.claim()));

self.addEventListener('sync', event => {
  if (event.tag === 'sync-offline-queue') {
    event.waitUntil(processQueueInWorker());
  }
});

// Allows the page to ask the (already-running) worker to sync immediately,
// e.g. right after queuing something while the tab is still open.
self.addEventListener('message', event => {
  if (event.data === 'sync-now') {
    event.waitUntil(processQueueInWorker().catch(() => {}));
  }
});
