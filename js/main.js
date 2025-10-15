// Entry point for the app
import { initDB, dbSave, dbLoadAll, STORES, db, dbClear, dbDelete } from './db.js';
import { categories, addCategory, addLibraryItem, getLibraryItems, clearLibrary, removeLibraryItemById } from './categories.js';
import { addNoteToCategory, getNotes } from './notes.js';
import { renderCategory, processAllDataAndDisplay, createPdf, rehydrateItem, normalizeFileDates, parseDateByLocale, formatByLocale, parseFile } from './ui.js';

// libraryFiles are maintained in the categories module; use addLibraryItem / getLibraryItems

function toast(message, duration = 3000) {
    const toastContainer = document.getElementById('toast-container');
    const toastEl = document.createElement('div');
    toastEl.className = 'toast fade-in';
    toastEl.textContent = message;
    toastContainer.appendChild(toastEl);
    setTimeout(() => toastEl.remove(), duration);
}

// Small HTML escaper used by UI snippets in this module (keeps this file self-contained)
function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (m) { return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]; });
}

// Helpers to enable/disable alerts debug mode and ensure console.debug is visible
function enableAlertsDebug() {
    try {
        window.__ALERTS_DEBUG = true;
        if (console && console.debug && console.log && console.debug !== console.log) console.debug = console.log.bind(console);
        console.log('Alerts debug enabled');
    } catch (e) { console.warn('enableAlertsDebug failed', e); }
}

function disableAlertsDebug() {
    try {
        window.__ALERTS_DEBUG = false;
        console.log('Alerts debug disabled');
    } catch (e) { console.warn('disableAlertsDebug failed', e); }
}
// Expose helpers on window so they can be invoked from the console (module scope isn't global)
try { if (typeof window !== 'undefined') { window.enableAlertsDebug = enableAlertsDebug; window.disableAlertsDebug = disableAlertsDebug; if (typeof window.__ALERTS_DEBUG === 'undefined') window.__ALERTS_DEBUG = false; } } catch (e) { /* ignore */ }
// Migration helper: re-normalize stored parsed file rows to YYYY-MM-DD using local date components
try {
    if (typeof window !== 'undefined') {
        window.migrateNormalizeStoredDates = async function migrateNormalizeStoredDates() {
            try {
                const items = (typeof getLibraryItems === 'function') ? getLibraryItems() : [];
                // Note: PNG export removed. Use SVG export (exportElementAsSvg) for reliable image exports.
                // PNG export removed — offer SVG export instead
                const canonicalToAliases = {};
                Object.keys(aliasMap).forEach(k => {
                    const can = aliasMap[k]; canonicalToAliases[can] = canonicalToAliases[can] || new Set(); canonicalToAliases[can].add(k);
                });
                const candidatesFor = (canonical) => { const s = new Set(); s.add(canonical); (canonicalToAliases[canonical] || []).forEach(a=>s.add(a)); return Array.from(s); };
                const lookup = (row, canonical) => {
                    if (!row) return undefined;
                    for (const key of candidatesFor(canonical)) { if (key in row) return row[key]; }
                    // fallback: try direct canonical if alias map empty
                    if (canonical in row) return row[canonical];
                    return undefined;
                };
                for (const f of files || []) {
                    (f.dataRows || []).forEach((r, idx) => {
                        try {
                            // diagnostic: if a row contains 'Shift Type' as an own property, log its value
                            try {
                                if (r && Object.prototype.hasOwnProperty.call(r, 'Shift Type')) {
                                    console.debug('__alertsIgnoredDuties: row has Shift Type key', f.name, idx, 'value', r['Shift Type']);
                                }
                                // also detect any string field that contains the word 'rest' to surface hidden matches
                                if (r && typeof r === 'object') {
                                    for (const k of Object.keys(r)) {
                                        try {
                                            const v = r[k];
                                            if (v != null && typeof v === 'string' && /\brest\b/i.test(v.trim())) {
                                                console.debug('__alertsIgnoredDuties: row field contains rest', f.name, idx, 'field', k, 'value', v);
                                                break;
                                            }
                                        } catch (e) { /* ignore */ }
                                    }
                                }
                            } catch (e) { /* ignore diag errors */ }
                            // try to find staff id but don't require it — we still want to detect Rest rows even if staff missing
                            let staffVal = null;
                            for (const sc of staffCols) { const v = lookup(r, sc); if (v != null) { staffVal = v; break; } }
                            // find duty date value across dutyCols (may be missing)
                            let dutyRaw = null;
                            for (const dc of dutyCols) { const v = lookup(r, dc); if (v != null) { dutyRaw = v; break; } }
                            // don't abort the entire helper for a single row missing a duty date, keep processing so
                            // we can still detect 'Rest' rows even when the date cell is absent or empty.
                            if (!dutyRaw) {
                                try { if (window.__ALERTS_DEBUG) console.debug('__alertsIgnoredDuties: row missing duty date (will continue to check shift type)', f.name, idx, Object.keys(r || {}).slice(0,10)); } catch (e) {}
                                // continue processing — dutyRaw remains null
                            }
                            // locate shift-type value across several likely header names
                            let dutyTypeVal = lookup(r, 'Shift Type');
                            if (dutyTypeVal == null) {
                                // try common alternate headers
                                const tryAlts = ['ShiftType','Type','Roster Type','Assignment Info','Assignment'];
                                for (const alt of tryAlts) {
                                    const v = lookup(r, alt);
                                    if (v != null) { dutyTypeVal = v; break; }
                                }
                            }
                            if (dutyTypeVal == null) { out.push({ file: f.name, rowIndex: idx, rowObj: r, reason: 'Shift Type missing', dutyRaw }); return; }
                            const low = String(dutyTypeVal).trim().toLowerCase();
                            // Per the rule: ignore (audit) this row as Rest if Shift Type contains 'rest'
                            // or Assignment Info contains 'do'. Any other Shift Type should be counted.
                            let assignmentInfoVal = null;
                            try {
                                const tryAlts = ['Assignment Info','Assignment','AssignmentInfo','Assign Info'];
                                for (const alt of tryAlts) { const v = lookup(r, alt); if (v != null) { assignmentInfoVal = v; break; } }
                            } catch (e) { /* ignore */ }
                            const assignLow = assignmentInfoVal != null ? String(assignmentInfoVal).trim().toLowerCase() : '';
                            if (/(?:\b|^)rest(?:\b|$)/i.test(low) || /(?:\b|^)do(?:\b|$)/i.test(assignLow)) {
                                out.push({ file: f.name, rowIndex: idx, rowObj: r, reason: 'Rest shift', dutyRaw, shift: dutyTypeVal, staff: staffVal || null });
                                return;
                            }
                            // Not a Rest according to strict rules — proceed and require a parsable duty date
                            const dtObj = parseToDate(dutyRaw, f._locale || 'uk');
                            if (!dtObj) { out.push({ file: f.name, rowIndex: idx, rowObj: r, reason: 'Duty Date parse failed', dutyRaw, shift: dutyTypeVal }); return; }
                        } catch (e) { /* ignore row errors */ }
                    });
                }
            } catch (e) { console.debug('computeIgnoredDutiesFromLibrary failed', e); }
            return out;
        };
        window.__alertsIgnoredDuties.print = function () { try { console.log('Ignored duties:', window.__alertsIgnoredDuties()); } catch (e) { console.warn('print failed', e); } };
        // Summary helper: returns counts and groupings for quick inspection
        window.__alertsIgnoredDuties.summary = function () {
            try {
                const arr = window.__alertsIgnoredDuties();
                const total = Array.isArray(arr) ? arr.length : 0;
                const byFile = {};
                const byStaff = {};
                (arr || []).forEach(it => {
                    const f = it.file || 'UNKNOWN'; byFile[f] = (byFile[f] || 0) + 1;
                    const s = (it.staff != null && String(it.staff).trim() !== '') ? String(it.staff) : 'UNKNOWN'; byStaff[s] = (byStaff[s] || 0) + 1;
                });
                return { total, byFile, byStaff };
            } catch (e) { console.warn('summary failed', e); return null; }
        };
        // Export helper: triggers CSV download of ignored duties (usable from browser console)
        window.__alertsIgnoredDuties.exportCsv = function (filename = 'ignored_duties.csv') {
            try {
                const arr = window.__alertsIgnoredDuties() || [];
                const cols = ['file','rowIndex','staff','shift','dutyRaw','matchedField','reason'];
                const rows = [cols.join(',')];
                for (const it of arr) {
                    const vals = cols.map(c => {
                        const v = it[c];
                        if (v == null) return '';
                        // escape double quotes
                        const s = String(v).replace(/"/g, '""');
                        // wrap in quotes if contains comma/newline
                        return /[",\n\r]/.test(s) ? `"${s}"` : s;
                    });
                    rows.push(vals.join(','));
                }
                const csv = rows.join('\r\n');
                // create blob and trigger download
                try {
                    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
                    const link = document.createElement('a');
                    const url = URL.createObjectURL(blob);
                    link.setAttribute('href', url);
                    link.setAttribute('download', filename);
                    link.style.visibility = 'hidden';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                    URL.revokeObjectURL(url);
                    console.log('Export started:', filename, 'rows:', arr.length);
                } catch (e) {
                    // fallback: print CSV to console
                    console.log('CSV export fallback - printing CSV to console. Rows:', arr.length);
                    console.log(csv);
                }
            } catch (e) { console.warn('exportCsv failed', e); }
        };
    }
} catch (e) { /* ignore */ }

// Column links: allow mapping multiple header names to a canonical header name
// Stored in localStorage under 'TERRA_COLUMN_LINKS' as an array of {id, canonical, aliases: []}
function getColumnLinks() {
    try {
        const raw = localStorage.getItem('TERRA_COLUMN_LINKS');
        if (!raw) return [];
        return JSON.parse(raw) || [];
    } catch (e) { console.warn('failed to load column links', e); return []; }
}

function saveColumnLinks(list) {
    try {
        localStorage.setItem('TERRA_COLUMN_LINKS', JSON.stringify(list || []));
    } catch (e) { console.warn('failed to save column links', e); }
}

function buildColumnAliasMap() {
    const links = getColumnLinks();
    const aliasToCanonical = {};
    links.forEach(l => {
        const can = (l.canonical || '').trim();
        if (!can) return;
        // map canonical to itself
        aliasToCanonical[can] = can;
        (l.aliases || []).forEach(a => {
            const aa = (a || '').trim(); if (!aa) return;
            aliasToCanonical[aa] = can;
        });
    });
    return aliasToCanonical;
}

function canonicalizeHeader(h) {
    if (!h) return h;
    const map = buildColumnAliasMap();
    return map[h] || h;
}

// Open (or create) the Column Links modal. Reusable so multiple UI places can call it.
function openColumnLinksModal() {
    let linksModal = document.getElementById('column-links-modal');
    if (!linksModal) {
        linksModal = document.createElement('div');
        linksModal.id = 'column-links-modal';
        linksModal.className = 'modal-backdrop';
        linksModal.innerHTML = `
            <div class="modal">
                <div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Manage Column Links</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div>
                <div class="modal-body p-2 max-h-[60vh] overflow-auto">
                    <div id="column-links-list" class="space-y-2"></div>
                    <div class="mt-3 flex gap-2"><input id="new-canonical" placeholder="Canonical name (e.g. Assignment No)" class="p-1 border rounded flex-1"/><input id="new-aliases" placeholder="Aliases (comma separated)" class="p-1 border rounded flex-1"/><button id="add-column-link" class="px-3 py-1 rounded bg-blue-600 text-white">Add</button></div>
                </div>
                <div class="modal-footer mt-3 text-right"><button id="close-links" class="btn-secondary px-3 py-1 rounded">Close</button></div>
            </div>`;
        document.body.appendChild(linksModal);
        const listEl = linksModal.querySelector('#column-links-list');
        function renderLinks() {
            const current = getColumnLinks();
            listEl.innerHTML = '';
            current.forEach((l, i) => {
                const row = document.createElement('div');
                row.className = 'p-2 border rounded flex items-center gap-2';
                const txt = document.createElement('div'); txt.className = 'flex-1 text-sm'; txt.innerHTML = `<strong>${escapeHtml(l.canonical)}</strong><div class="text-subtle text-xs">${escapeHtml((l.aliases||[]).join(', '))}</div>`;
                const del = document.createElement('button'); del.className = 'text-red-600'; del.textContent = 'Remove';
                row.appendChild(txt); row.appendChild(del); listEl.appendChild(row);
                del.addEventListener('click', () => {
                    const arr = getColumnLinks(); arr.splice(i,1); saveColumnLinks(arr); renderLinks();
                });
            });
        }
        linksModal.querySelector('.modal-close-btn').addEventListener('click', () => linksModal.classList.remove('active'));
        linksModal.querySelector('#close-links').addEventListener('click', () => linksModal.classList.remove('active'));
        linksModal.addEventListener('click', (e) => { if (e.target === linksModal) linksModal.classList.remove('active'); });
        linksModal.querySelector('#add-column-link').addEventListener('click', () => {
            const can = linksModal.querySelector('#new-canonical').value.trim();
            const aliasTxt = linksModal.querySelector('#new-aliases').value.trim();
            if (!can) return alert('Canonical name required');
            const aliases = aliasTxt ? aliasTxt.split(',').map(s=>s.trim()).filter(Boolean) : [];
            const arr = getColumnLinks(); arr.push({ id: `cl_${Date.now()}`, canonical: can, aliases }); saveColumnLinks(arr); renderLinks();
            linksModal.querySelector('#new-canonical').value = ''; linksModal.querySelector('#new-aliases').value = '';
        });
        renderLinks();
    }
    linksModal.classList.add('active');
}

async function init() {
    // apply saved theme early so styles load correctly
    const savedTheme = localStorage.getItem('TERRA_THEME') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);
    await initDB();
    // Cleanup: remove any null/undefined or malformed records from stores
    try {
        await Promise.all(Object.values(STORES).map(s => db && db.cleanNullishRecords ? db.cleanNullishRecords(s) : null));
    } catch (e) {
        // Older environments may not expose the function on db; fallback to importing helper
        try {
            // import the cleanup helper from db.js dynamically (works in module context)
            const dbModule = await import('./db.js');
            await Promise.all(Object.values(STORES).map(s => dbModule.cleanNullishRecords(s)));
        } catch (e2) { console.warn('cleanup: failed to clean DB records', e2); }
    }
    // One-time migration: if the date-display preference was saved when labels were swapped,
    // flip it back so stored values match labels. Use a flag to avoid repeating.
    try {
        if (!localStorage.getItem('TERRA_DATE_DISPLAY_MIGRATED')) {
            const v = localStorage.getItem('TERRA_DATE_DISPLAY');
            if (v === 'uk' || v === 'us') {
                // during the temporary label swap, the meanings were inverted; flip them now
                const flipped = v === 'uk' ? 'us' : 'uk';
                localStorage.setItem('TERRA_DATE_DISPLAY', flipped);
            }
            localStorage.setItem('TERRA_DATE_DISPLAY_MIGRATED', '1');
        }
    } catch (e) { /* ignore */ }
    // Load library files from DB
    const allData = await Promise.all(Object.values(STORES).map(store => {
        return new Promise((resolve, reject) => {
            const tx = db.transaction(store, 'readonly');
            const storeObj = tx.objectStore(store);
            const req = storeObj.getAll();
            req.onsuccess = () => resolve(req.result);
            req.onerror = reject;
        });
    }));
    // populate shared libraryFiles via addLibraryItem
    clearLibrary();
    // flatten DB results and drop any null/undefined entries that might have been stored by mistake
    const rawItems = allData.flat().filter(it => it != null);
    rawItems.forEach(it => { if (it) addLibraryItem(it); });
    const libraryFiles = getLibraryItems();
    // Ensure parsed files have a default locale (uk) so initial display is consistent
    for (const item of libraryFiles) {
        if (item && item.isParsedFile) {
            if (!item._locale) {
                try {
                    item._locale = 'uk';
                    await dbSave(STORES.FILES, { ...item, _locale: item._locale });
                    console.debug('set default _locale=uk for file', item.name, item.id);
                } catch (e) {
                    console.error('failed to persist default locale for', item.id, e);
                }
            }
        }
    }

    // Heuristic detection: inspect numeric slashed dates in each file to detect probable locale
    function detectLocaleForFile(rows) {
        // Bias towards UK (DD/MM) as the safe default. Require stronger confidence to pick US.
        let ukVotes = 0, usVotes = 0, total = 0;
        const slashRe = /^(\s*)(\d{1,2})\/(\d{1,2})\/(\d{2,4})(\s*)$/;
        for (const r of rows || []) {
            for (const k of Object.keys(r || {})) {
                const v = r[k];
                if (typeof v !== 'string') continue;
                const m = v.match(slashRe);
                if (m) {
                    const a = parseInt(m[2], 10), b = parseInt(m[3], 10);
                    if (a > 12 && b <= 12) { ukVotes++; total++; }
                    else if (b > 12 && a <= 12) { usVotes++; total++; }
                    else if (a <= 12 && b <= 12) { /* ambiguous */ total++; }
                }
            }
        }
        if (total === 0) return { locale: null, confidence: 0 };
        // For ambiguous entries (both <=12) we don't vote, instead rely on explicit votes
        const explicitVotes = ukVotes + usVotes;
        const confidence = explicitVotes > 0 ? Math.max(ukVotes, usVotes) / explicitVotes : 0;
        // If UK has more votes, accept it with moderate confidence
        if (ukVotes > usVotes && confidence >= 0.5) return { locale: 'uk', confidence };
        // Only pick US if it has strong majority (reduce false positives)
        if (usVotes > ukVotes && confidence >= 0.75) return { locale: 'us', confidence };
        // Otherwise stay undecided and let default 'uk' apply elsewhere
        return { locale: null, confidence };
    }

    // Run detection and update _locale when confident
    for (const item of libraryFiles) {
        if (item && item.isParsedFile && Array.isArray(item.dataRows)) {
            try {
                const det = detectLocaleForFile(item.dataRows);
                if (det.locale && det.locale !== item._locale) {
                    const old = item._locale;
                    item._locale = det.locale;
                    try {
                        await dbSave(STORES.FILES, { ...item, _locale: item._locale });
                        console.debug('auto-detected locale', det.locale, 'for', item.name, 'confidence', det.confidence, 'old', old);
                    } catch (e) {
                        console.error('failed to persist auto-detected locale for', item.id, e);
                    }
                }
            } catch (e) { console.error('locale detection error for', item.id, e); }
        }
    }
    // Migration: ensure all persisted parsed files have date-like cells stored as ISO (YYYY-MM-DD).
    try {
        const filesStoreItems = libraryFiles.filter(i => i && i.isParsedFile);
        for (const item of filesStoreItems) {
            let changed = false;
            const fileLocale = item._locale || 'uk';
            const newRows = (item.dataRows || []).map(row => {
                const out = { ...row };
                Object.keys(out).forEach(k => {
                    const v = out[k];
                    if (typeof v === 'string') {
                        // If already ISO, skip
                        if (/^\s*\d{4}-\d{2}-\d{2}\s*$/.test(v)) return;
                        const parsed = parseDateByLocale(v, fileLocale);
                        if (parsed) {
                            const m = parsed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
                            if (m) {
                                const d = m[1].padStart(2,'0'); const mo = m[2].padStart(2,'0'); const yy = m[3].length === 2 ? (parseInt(m[3],10) > 50 ? '19'+m[3] : '20'+m[3]) : m[3];
                                const iso = `${yy}-${mo}-${d}`;
                                if (iso !== v) { out[k] = iso; changed = true; }
                            }
                        }
                    }
                    if (v instanceof Date) {
                        const iso = toLocalIso(v);
                        if (iso !== v) { out[k] = iso; changed = true; }
                    }
                });
                return out;
            });
            if (changed) {
                try { await dbSave(STORES.FILES, { ...item, dataRows: newRows }); item.dataRows = newRows; console.debug('migration: updated file to ISO dates', item.name); } catch (e) { console.error('migration: failed to update file', item.id, e); }
            }
        }
    } catch (e) { console.error('migration: failed', e); }
    // Normalize stored parsed files by applying locale-aware parsing to every string cell
    // Store dates as unambiguous ISO (YYYY-MM-DD). The UI will still display as UK via formatByLocale.
    for (const item of libraryFiles) {
        if (item && item.isParsedFile && Array.isArray(item.dataRows)) {
            const fileLocale = item._locale || 'uk';
            const normalized = item.dataRows.map(row => {
                const out = { ...row };
                Object.keys(out).forEach(k => {
                    const v = out[k];
                    if (typeof v === 'string') {
                        const parsed = parseDateByLocale(v, fileLocale);
                        if (parsed) {
                            // parseDateByLocale now returns ISO (YYYY-MM-DD) when it can parse a date
                            out[k] = parsed;
                        }
                    }
                    if (v instanceof Date) out[k] = toLocalIso(v);
                });
                return out;
            });
            const before = JSON.stringify(item.dataRows);
            const after = JSON.stringify(normalized);
            if (before !== after) {
                try {
                    await dbSave(STORES.FILES, { ...item, dataRows: normalized });
                    console.debug('normalized file', item.name, 'locale', item._locale, 'sample before', JSON.stringify(item.dataRows[0]), 'after', JSON.stringify(normalized[0]));
                    item.dataRows = normalized;
                } catch (e) { console.error('failed to persist normalized file', e); }
            }
        }
    }
    // rehydrate library items into categories so customization can edit them
    for (const item of libraryFiles) {
        try { rehydrateItem(item); } catch (e) { console.error('rehydrate error', e); }
    }
    console.debug('main:init loaded library items:', libraryFiles.length, libraryFiles.map(i => (i && typeof i === 'object' && 'name' in i) ? i.name : i));
    toast(`Loaded ${libraryFiles.length} item(s) from library.`);
    // Load saved alert rules from DB (if any)
    try {
        const savedRules = await dbLoadAll(STORES.RULES).catch(()=>[]);
        // attach to a global holder so Alerts UI can pick them up when rendering
        window.__ALERT_RULES = Array.isArray(savedRules) ? savedRules : [];
    } catch (e) { console.warn('Failed to load saved alert rules', e); window.__ALERT_RULES = []; }
    // Initial UI rendering
    renderCategoryUI();
    setupSettings();
}

function renderCategoryUI() {
    const mainContent = document.getElementById('main-content');
    // Tab header + content area
    mainContent.innerHTML = `
        <div class="tabs mb-6 flex items-center gap-3">
            <button id="tab-builder" class="px-4 py-2 rounded-md font-semibold bg-white shadow-sm">Report Builder</button>
            <button id="tab-alerts" class="px-4 py-2 rounded-md text-sm text-muted hover:bg-white">Alerts</button>
        </div>
        <div id="tab-content"></div>
    `;

    const tabContent = document.getElementById('tab-content');
    const tabBuilder = document.getElementById('tab-builder');
    const tabAlerts = document.getElementById('tab-alerts');

    function setActiveTab(tab) {
        if (tab === 'builder') {
            tabBuilder.classList.add('bg-white');
            tabBuilder.classList.remove('text-muted');
            tabAlerts.classList.remove('bg-white');
            tabAlerts.classList.add('text-muted');
            // Defer heavy rendering to idle time for faster click response
            const doRender = () => { try { renderReportBuilder(); } catch (e) { console.error('renderReportBuilder failed', e); } };
            if (typeof requestIdleCallback === 'function') requestIdleCallback(doRender, { timeout: 300 }); else setTimeout(doRender, 20);
        } else {
            tabAlerts.classList.add('bg-white');
            tabAlerts.classList.remove('text-muted');
            tabBuilder.classList.remove('bg-white');
            tabBuilder.classList.add('text-muted');
            const doRender = () => { try { renderAlerts(); } catch (e) { console.error('renderAlerts failed', e); } };
            if (typeof requestIdleCallback === 'function') requestIdleCallback(doRender, { timeout: 300 }); else setTimeout(doRender, 20);
        }
    }

    tabBuilder.addEventListener('click', () => setActiveTab('builder'));
    tabAlerts.addEventListener('click', () => setActiveTab('alerts'));

    // default to Report Builder — render synchronously on initial load so the page is ready
    tabBuilder.classList.add('bg-white'); tabBuilder.classList.remove('text-muted'); tabAlerts.classList.remove('bg-white'); tabAlerts.classList.add('text-muted');
    try { renderReportBuilder(); } catch (e) { console.error('initial renderReportBuilder failed', e); }
}

// Build an in-memory index mapping numeric staff keys -> array of { file, rowIndex, row }
function buildStaffIndex() {
    try {
        const files = (typeof getParsedFiles === 'function') ? getParsedFiles() : (typeof window !== 'undefined' && window.__getParsedFiles ? window.__getParsedFiles() : []);
        const idx = {};
        (files || []).forEach(f => {
            (f.dataRows || []).forEach((r, i) => {
                try {
                    const k = String(getStaffKey(r) || '').replace(/\D/g, '');
                    if (!k) return;
                    idx[k] = idx[k] || [];
                    idx[k].push({ file: f, rowIndex: i, row: r });
                } catch (e) { /* ignore per-row */ }
            });
        });
        try { if (typeof window !== 'undefined') { window.__staffIndex = idx; window.rebuildStaffIndex = buildStaffIndex; } } catch (e) {}
        return idx;
    } catch (e) { console.error('buildStaffIndex failed', e); return {}; }
}

// Render the existing report-builder UI into the tab content (keeps previous behaviour)
function renderReportBuilder() {
    const tabContent = document.getElementById('tab-content');
    if (!tabContent) return;

    // Defensive: clear any lingering active modal backdrops that could block UI clicks
    try {
        const stuck = Array.from(document.querySelectorAll('.modal-backdrop.active'));
        stuck.forEach(s => { try { s.classList.remove('active'); } catch (e) {} });
    } catch (e) { /* ignore */ }
    tabContent.innerHTML = `<div id="step1"><div id="categories-container" class="space-y-6"></div><div class="text-center mt-6 pt-6 border-t border-base space-x-4"><button id="add-category-btn" class="btn-secondary px-5 py-2.5 rounded-lg font-semibold">+ Add Category</button><button id="proceed-btn" class="bg-blue-600 text-white px-8 py-3 rounded-lg font-semibold text-lg hover:bg-blue-700 transition-all shadow-md disabled:bg-slate-300 disabled:cursor-not-allowed" disabled>Customize Report</button></div><div id="error-message" class="hidden mt-4 text-center text-red-600 bg-red-100 p-3 rounded-lg"></div></div>`;
    const addBtn = document.getElementById('add-category-btn');
    addBtn.addEventListener('click', () => {
        const categoryId = `cat_${Date.now()}`;
        const category = { id: categoryId, name: `Category ${categories.length + 1}`, files: [], data: [], headers: [], selectedColumns: [], filterRules: [] };
        addCategory(category);
        renderCategory(category);
    });
    // Render existing categories
    const container = document.getElementById('categories-container');
    categories.forEach(cat => renderCategory(cat));
    const proceedBtn = document.getElementById('proceed-btn');
    if (proceedBtn) proceedBtn.addEventListener('click', renderStep2);
}

// Render the Alerts tab content styled like the report builder
function renderAlerts() {
    const tabContent = document.getElementById('tab-content');
    if (!tabContent) return;

    // State for alerts UI
    let selectedFileIds = new Set();
    // Persist selected file ids so the user's selection survives reloads
    function saveAlertsSelectedFiles() {
        try { localStorage.setItem('TERRA_ALERTS_SELECTED_FILES', JSON.stringify(Array.from(selectedFileIds))); } catch(e){ console.warn('Failed to save selected files', e); }
    }
    function loadAlertsSelectedFiles() {
        try {
            const raw = localStorage.getItem('TERRA_ALERTS_SELECTED_FILES');
            if (!raw) return new Set();
            const arr = JSON.parse(raw || '[]');
            return new Set(Array.isArray(arr) ? arr : []);
        } catch (e) { console.warn('Failed to load selected files', e); return new Set(); }
    }
    // initialize the selection from storage
    try { selectedFileIds = loadAlertsSelectedFiles(); } catch(e){ selectedFileIds = new Set(); }
    // initialize rules from DB-loaded global (set during init)
    let rules = (window.__ALERT_RULES && Array.isArray(window.__ALERT_RULES)) ? window.__ALERT_RULES.map(r => ({ ...r })) : [];

    // If no rules are present, show a lightweight hint and render a lightweight rules editor.
    // Avoid scanning files/headers synchronously here — schedule the heavy scan slightly later.
    if (!rules || rules.length === 0) {
        try {
            const resultsEl = tabContent.querySelector('#alerts-results');
            if (resultsEl) {
                resultsEl.innerHTML = '';
                const hint = document.createElement('div');
                hint.className = 'p-4 rounded border bg-white text-sm';
                hint.innerHTML = `<div class="text-lg font-semibold mb-2">No rules defined</div><div class="text-sm">Add one or more rules to see matched results. The file browser, RTW summary and NMC pins are available while you configure rules.</div>`;
                resultsEl.appendChild(hint);
            }
        } catch (e) { /* ignore render hint errors */ }
        // Render rules editor light-weight (no files/headers) then schedule a deferred real render
        try { renderRulesEditor([], []); } catch (e) { /* ignore */ }
        try {
            setTimeout(() => {
                try { renderRulesEditor(getParsedFiles(), unionHeaders(getParsedFiles())); } catch (e) { /* ignore */ }
            }, 250);
        } catch (e) { /* ignore */ }
    }

    async function saveAllRules() {
        try {
            await dbClear(STORES.RULES);
            for (const r of rules) {
                // ensure id not lost — let db assign if missing
                const toSave = { ...r };
                // remove transient properties if present
                delete toSave.__temp;
                const savedId = await dbSave(STORES.RULES, toSave);
                // update local rule id if newly assigned
                r.id = savedId;
            }
            // refresh global snapshot
            window.__ALERT_RULES = rules.map(r => ({ ...r }));
        } catch (e) { console.warn('Failed to persist alert rules', e); }
    }

    function getParsedFiles() {
        return (getLibraryItems() || []).filter(i => i && i.isParsedFile && Array.isArray(i.dataRows));
    }

    function unionHeaders(files) {
        const s = new Set();
        const aliasMap = buildColumnAliasMap();
        files.forEach(f => (f.headers || []).forEach(h => {
            const can = aliasMap[h] || h;
            s.add(can);
        }));
        const debug = !!window.__ALERTS_DEBUG;
        if (debug) {
            try {
                console.debug('unionHeaders: input files', files.map(f => ({ id: f.id, name: f.name, headers: f.headers }))); 
                console.debug('unionHeaders: aliasMap', aliasMap);
                console.debug('unionHeaders: result', Array.from(s));
            } catch (e) { console.debug('unionHeaders debug failed', e); }
        }
        return Array.from(s);
    }

    // Given a canonical header name, find the first matching value in a row by
    // checking canonical and alias header keys.
    function getValueByCanonical(row, canonical) {
        if (!row || !canonical) return undefined;
        const aliasMap = buildColumnAliasMap();
        // build canonical -> aliases map
        const canonicalToAliases = {};
        Object.keys(aliasMap).forEach(k => {
            const can = aliasMap[k];
            canonicalToAliases[can] = canonicalToAliases[can] || new Set();
            canonicalToAliases[can].add(k);
        });
        const candidates = new Set();
        candidates.add(canonical);
        (canonicalToAliases[canonical] || []).forEach(a => candidates.add(a));
        const debug = !!window.__ALERTS_DEBUG;
        if (debug) console.debug('getValueByCanonical: looking for', canonical, 'candidates', Array.from(candidates), 'in row keys', Object.keys(row));
        for (const key of candidates) {
            if (key in row) {
                if (debug) console.debug('getValueByCanonical: found value for', key, '=>', row[key]);
                return row[key];
            }
        }
        // Fallback: try tolerant matching against row keys (case-insensitive, punctuation-insensitive, singular/plural)
        try {
            const normalize = (s) => (String(s || '')).toLowerCase().replace(/[^a-z0-9]/g, '').replace(/s$/,'');
            const want = normalize(canonical);
            for (const rk of Object.keys(row)) {
                try {
                    const cand = normalize(rk);
                    if (!cand) continue;
                    if (cand === want) {
                        if (debug) console.debug('getValueByCanonical: tolerant match', rk, 'for', canonical);
                        return row[rk];
                    }
                    // allow partial matches for common pairs like 'forename' ~ 'forenames' or 'dutydate' ~ 'duty'
                    if (cand.includes(want) || want.includes(cand)) {
                        if (debug) console.debug('getValueByCanonical: partial tolerant match', rk, 'for', canonical);
                        return row[rk];
                    }
                } catch (e) { /* ignore key-specific errors */ }
            }
        } catch (e) { if (debug) console.debug('getValueByCanonical: tolerant fallback failed', e); }
        if (debug) console.debug('getValueByCanonical: no value found for', canonical);
        return undefined;
    }

    // Robust extraction of a staff identifier from a row using common header names
    function getStaffKey(row) {
        if (!row) return '';
        const candidates = ['Staff Number','Staff No','Employee Number','Employee No','Assignment No','Assignment No.','Staff','StaffID','ID','EmployeeID'];
        const normalize = (val) => {
            if (val == null) return '';
            let s = String(val || '').trim();
            // If the assignment includes a hyphen suffix (e.g. 27029932-2), keep only the part before the hyphen
            if (s.includes('-')) s = s.split('-')[0];
            // Keep digits only (staff numbers are numeric). This removes spaces and other punctuation.
            s = s.replace(/\D+/g, '');
            return s;
        };
        for (const c of candidates) {
            try {
                const v = getValueByCanonical(row, c);
                const n = normalize(v);
                if (n) return n;
            } catch (e) { /* ignore per-field errors */ }
        }
        // fallback: look for any key that contains 'staff' or 'assign' and return its value
        try {
            for (const k of Object.keys(row)) {
                const nk = String(k || '').toLowerCase();
                if (nk.includes('staff') || nk.includes('assign') || nk.includes('employee')) {
                    const v = row[k];
                    const n = normalize(v);
                    if (n) return n;
                }
            }
        } catch (e) { /* ignore */ }
        return '';
    }

    // Helper to parse a date-like value into a Date object, trying various formats and locales.
    function parseToDate(input, locale) {
        if (!input && input !== 0) return null;
        if (input instanceof Date && !isNaN(input)) return input;
        const s = String(input || '').trim();
        // helper to convert canonical ISO YYYY-MM-DD to Date
        const fromCanonical = (canonical) => {
            if (!canonical) return null;
            const m = canonical.match(/^(\d{4})-(\d{2})-(\d{2})$/);
            if (!m) return null;
            const y = Number(m[1]); const mon = Number(m[2]); const d = Number(m[3]);
            const dt = new Date(y, mon - 1, d);
            return isNaN(dt) ? null : dt;
        };
        // If the input is already ISO YYYY-MM-DD, prefer that (unambiguous)
        const isoMatch = s.match(/^\s*(\d{4})-(\d{2})-(\d{2})\s*$/);
        if (isoMatch) {
            const y = Number(isoMatch[1]), mo = Number(isoMatch[2]), d = Number(isoMatch[3]);
            const dtIso = new Date(y, mo - 1, d);
            if (!isNaN(dtIso)) return dtIso;
        }
        try {
            // First try file-local parsing
            const parsedStr = parseDateByLocale(s, locale);
            const dt1 = fromCanonical(parsedStr);
            if (dt1) return dt1;
            // If input is a numeric slashed date, try the alternate locale as a fallback
            if (/^\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*$/.test(s)) {
                const other = (locale === 'uk') ? 'us' : 'uk';
                    const parsed2 = parseDateByLocale(s, other);
                    const dt2 = fromCanonical(parsed2);
                if (dt2) return dt2;
            }
        } catch (e) { /* ignore */ }
        // final fallback: try Date.parse on the original string
        try {
            const t = Date.parse(s);
            if (!isNaN(t)) return new Date(t);
        } catch (e) { /* ignore */ }
        return null;
    }

    // CSV safe quoting helper used by modal export buttons
    function csvSafe(v) {
        if (v === null || v === undefined) return '""';
        const s = String(v);
        return '"' + s.replace(/"/g, '""') + '"';
    }

    // Helper to find a human name for a staff id by searching parsed files
    function findStaffName(staffId) {
        try {
            const files = getParsedFiles();
            // First pass: prefer rows that include both first and last in the same row
            const firstCols = ['Forename','Forenames','First Name','Given Name','GivenName','Forename(s)','Name'];
            const lastCols = ['Surname','Last Name','Lastname','Family Name','FamilyName'];
            const nameCols = ['Name','Full Name','FullName','Display Name','DisplayName'];
            for (const f of files) {
                for (const r of (f.dataRows || [])) {
                    const key = String(getStaffKey(r) || '').trim();
                    if (!key) continue;
                    if (key === String(staffId)) {
                        let first = null; let last = null;
                        for (const fc of firstCols) {
                            const fv = getValueByCanonical(r, fc);
                            if (fv != null && String(fv).toString().trim()) { first = String(fv).trim(); break; }
                        }
                        for (const lc of lastCols) {
                            const lv = getValueByCanonical(r, lc);
                            if (lv != null && String(lv).toString().trim()) { last = String(lv).trim(); break; }
                        }
                        if (first && last) return `${first} ${last}`;
                        // if row contains a single-name field like Full Name, prefer it
                        for (const nc of nameCols) {
                            const nv = getValueByCanonical(r, nc);
                            if (nv != null && String(nv).toString().trim()) return String(nv).trim();
                        }
                        if (first) return first;
                        if (last) return last;
                    }
                }
            }
            // Second pass: if no single row had both, collect first and last across rows and combine
            let anyFirst = null, anyLast = null;
            for (const f of files) {
                for (const r of (f.dataRows || [])) {
                    const key = String(getStaffKey(r) || '').trim();
                    if (!key) continue; if (key !== String(staffId)) continue;
                    if (!anyFirst) {
                        for (const fc of firstCols) {
                            const fv = getValueByCanonical(r, fc);
                            if (fv != null && String(fv).toString().trim()) { anyFirst = String(fv).trim(); break; }
                        }
                    }
                    if (!anyLast) {
                        for (const lc of lastCols) {
                            const lv = getValueByCanonical(r, lc);
                            if (lv != null && String(lv).toString().trim()) { anyLast = String(lv).trim(); break; }
                        }
                    }
                    if (anyFirst && anyLast) return `${anyFirst} ${anyLast}`;
                }
            }
            if (anyFirst && anyLast) return `${anyFirst} ${anyLast}`;
            if (anyFirst) return anyFirst; if (anyLast) return anyLast;
        } catch (e) { /* ignore */ }
        return null;
    }

    // Format a Date using local timezone components into YYYY-MM-DD to avoid UTC shift
    function toLocalIso(dt) {
        if (!(dt instanceof Date) || isNaN(dt)) return null;
        const y = dt.getFullYear();
        const m = String(dt.getMonth() + 1).padStart(2, '0');
        const d = String(dt.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
    }

    function evaluateRuleOnValue(val, rule, locale) {
        if (val == null) return false;
        const debug = !!window.__ALERTS_DEBUG;
        const ruleValRaw = String(rule.value || '');
        const ruleVal = ruleValRaw.toLowerCase();
        const numericRe = /^\s*-?\d+(?:\.\d+)?\s*$/;
        const isValNum = numericRe.test(String(val));
        const isRuleNum = numericRe.test(String(rule.value));

        // Build multiple textual representations for the cell value so filters match ISO, canonical and display forms.
        const buildRepresentations = (v) => {
            const reps = new Set();
            if (v == null) return Array.from(reps);
            try { reps.add(String(v)); } catch (e) { /* ignore */ }
            try { reps.add(String(formatByLocale(v, locale))); } catch (e) { /* ignore */ }
            // If a string, try canonical parse (DD/MM/YYYY) using parseDateByLocale
            if (typeof v === 'string') {
                try {
                    const canon = parseDateByLocale(v, locale);
                    if (canon) reps.add(canon);
                    // if canonical matched DD/MM/YYYY convert to ISO
                    const m = canon && canon.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
                    if (m) {
                        let d = m[1].padStart(2,'0');
                        let mo = m[2].padStart(2,'0');
                        let yy = m[3];
                        if (yy.length === 2) yy = (parseInt(yy,10) > 50 ? '19'+yy : '20'+yy);
                        reps.add(`${yy}-${mo}-${d}`);
                    }
                    // if input was ISO, add canonical form
                    const isoMatch = v.match(/^\s*(\d{4})-(\d{2})-(\d{2})\s*$/);
                    if (isoMatch) {
                        const y = isoMatch[1], mo = isoMatch[2], d = isoMatch[3]; reps.add(`${d}/${mo}/${y}`);
                    }
                } catch (e) { /* ignore */ }
            } else if (v instanceof Date) {
                try { reps.add(toLocalIso(v)); } catch(e){}
                try { reps.add(formatByLocale(toLocalIso(v), locale)); } catch(e){}
            }
            return Array.from(reps).map(x => (x == null ? '' : String(x)));
        };

        const valReps = buildRepresentations(val).map(s => s.toLowerCase());
        const ruleValLower = ruleVal.toLowerCase();

        // Date-aware parse helper to compare date equality when possible
        const parsePossibleDate = (input) => {
            if (!input && input !== 0) return null;
            if (input instanceof Date && !isNaN(input)) return input;
            const s = String(input || '').trim();
            const isoMatch = s.match(/^\s*(\d{4})-(\d{2})-(\d{2})\s*$/);
            if (isoMatch) {
                const y = Number(isoMatch[1]), mo = Number(isoMatch[2]), d = Number(isoMatch[3]);
                const dt = new Date(y, mo - 1, d);
                if (!isNaN(dt)) return dt;
            }
            try {
                const canon = parseDateByLocale(s, locale);
                if (canon) {
                    const m = canon.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
                    if (m) {
                        const d = Number(m[1]); const mo = Number(m[2]); let yy = m[3];
                        if (yy.length === 2) yy = (parseInt(yy,10) > 50 ? '19' + yy : '20' + yy);
                        const dt = new Date(Number(yy), mo - 1, d);
                        if (!isNaN(dt)) return dt;
                    }
                }
            } catch (e) { /* ignore */ }
            const t = Date.parse(s);
            if (!isNaN(t)) return new Date(t);
            return null;
        };

        if (rule.operator === 'contains') {
            const res = valReps.some(r => r.includes(ruleValLower));
            if (debug) console.debug('rule contains', rule, val, 'representations', valReps, '=>', res);
            return res;
        }
        if (rule.operator === 'not_contains') {
            const res = !valReps.some(r => r.includes(ruleValLower));
            if (debug) console.debug('rule not_contains', rule, val, 'representations', valReps, '=>', res);
            return res;
        }
        if (rule.operator === 'equals') {
            // First attempt string equality against any representation
            const strMatch = valReps.some(r => r === ruleValLower);
            if (strMatch) { if (debug) console.debug('rule equals string-match', rule, val, '=>', true); return true; }
            // If both sides can be parsed as dates, compare date-only equality
            try {
                const leftDate = parsePossibleDate(val);
                const rightDate = parsePossibleDate(rule.value);
                if (leftDate && rightDate) {
                    const d1 = `${leftDate.getFullYear()}-${String(leftDate.getMonth()+1).padStart(2,'0')}-${String(leftDate.getDate()).padStart(2,'0')}`;
                    const d2 = `${rightDate.getFullYear()}-${String(rightDate.getMonth()+1).padStart(2,'0')}-${String(rightDate.getDate()).padStart(2,'0')}`;
                    const res = d1 === d2;
                    if (debug) console.debug('rule equals date-compare', { rule, val, leftDate: d1, rightDate: d2, res });
                    return res;
                }
            } catch (e) { if (debug) console.debug('rule equals date-compare failed', e); }
            if (debug) console.debug('rule equals', rule, val, '=>', false);
            return false;
        }
        // Blank checks (treat common placeholder tokens as blank)
        const _normalizeForBlank = (v) => {
            if (v === null || v === undefined) return '';
            const s = String(v).trim();
            if (!s) return '';
            // strip trailing dots and normalize
            const token = s.replace(/\.+$/,'').trim().toLowerCase();
            const BLANK_TOKENS = new Set(['-','n/a','na','none','']);
            return BLANK_TOKENS.has(token) ? '' : s;
        };
        if (rule.operator === 'is_blank') {
            const normalized = _normalizeForBlank(val);
            const isBlank = normalized === '';
            if (debug) console.debug('rule is_blank', rule, val, 'normalized->', normalized, '=>', isBlank);
            return isBlank;
        }
        if (rule.operator === 'is_not_blank') {
            const normalized = _normalizeForBlank(val);
            const isNotBlank = normalized !== '';
            if (debug) console.debug('rule is_not_blank', rule, val, 'normalized->', normalized, '=>', isNotBlank);
            return isNotBlank;
        }
        // Date: within last N days (rule.value is number of days)
        if (rule.operator === 'within_days') {
            const days = Number(String(rule.value || '').trim());
            if (debug) console.debug('within_days check', { val, ruleValue: rule.value, days, locale });
            if (!Number.isFinite(days) || days < 0) return false;
            const dateVal = parseToDate(val, locale);
            if (!dateVal) { if (debug) console.debug('within_days: failed to parse date from value', val); return false; }
            // consider "within last N days" relative to now (inclusive)
            // Use local YYYY-MM-DD strings for date-only comparisons to avoid timezone shifts.
            const toLocalIso = (dt) => {
                if (!(dt instanceof Date) || isNaN(dt)) return null;
                const y = dt.getFullYear(); const m = String(dt.getMonth() + 1).padStart(2, '0'); const d = String(dt.getDate()).padStart(2, '0');
                return `${y}-${m}-${d}`;
            };
            const now = new Date();
            const nowIso = toLocalIso(now);
            const start = new Date(now);
            start.setDate(start.getDate() - (Math.floor(days) - 1));
            const startIso = toLocalIso(start);
            const cmpIso = toLocalIso(dateVal);
            const match = cmpIso && startIso && nowIso ? (cmpIso >= startIso && cmpIso <= nowIso) : false;
            if (debug) console.debug('within_days compare', { cmpIso, startIso, nowIso, match });
            return match;
        }
        // Date: within next N days (future-looking)
        if (rule.operator === 'within_next_days') {
            const days = Number(String(rule.value || '').trim());
            if (debug) console.debug('within_next_days check', { val, ruleValue: rule.value, days, locale });
            if (!Number.isFinite(days) || days < 0) return false;
            const dateVal = parseToDate(val, locale);
            if (!dateVal) { if (debug) console.debug('within_next_days: failed to parse date from value', val); return false; }
            // Use local YYYY-MM-DD strings for date-only comparisons
            const toLocalIso = (dt) => {
                if (!(dt instanceof Date) || isNaN(dt)) return null;
                const y = dt.getFullYear(); const m = String(dt.getMonth() + 1).padStart(2, '0'); const d = String(dt.getDate()).padStart(2, '0');
                return `${y}-${m}-${d}`;
            };
            const now = new Date();
            const nowIso = toLocalIso(now);
            const end = new Date(now);
            end.setDate(end.getDate() + Math.floor(days));
            const endIso = toLocalIso(end);
            const cmpIso = toLocalIso(dateVal);
            const match = cmpIso && nowIso && endIso ? (cmpIso >= nowIso && cmpIso <= endIso) : false;
            if (debug) console.debug('within_next_days compare', { cmpIso, nowIso, endIso, match });
            return match;
        }
        if (['numeric_gt','numeric_gte','numeric_lt','numeric_lte'].includes(rule.operator) && isValNum && isRuleNum) {
            const a = Number(String(val).trim());
            const b = Number(String(rule.value).trim());
            if (rule.operator === 'numeric_gt') return a > b;
            if (rule.operator === 'numeric_gte') return a >= b;
            if (rule.operator === 'numeric_lt') return a < b;
            if (rule.operator === 'numeric_lte') return a <= b;
        }
        return false;
    }

    function evaluateRowAgainstRules(row, file, rules) {
        if (!rules || rules.length === 0) return true;
        const debug = !!window.__ALERTS_DEBUG;
        return rules.every(rule => {
            const val = getValueByCanonical(row, rule.column);
            if (debug) console.debug('evaluating rule on row', { rule, fileId: file && file.id, row });
            if (val === undefined) { if (debug) console.debug('no value found for canonical column', rule.column); return false; }
            const res = evaluateRuleOnValue(val, rule, file._locale || 'uk');
            if (debug) console.debug('rule result', { rule, value: val, res });
            return res;
        });
    }

    function buildResults() {
        const debug = !!window.__ALERTS_DEBUG;
        const files = getParsedFiles().filter(f => selectedFileIds.has(f.id));
        if (debug) console.debug('buildResults: selectedFileIds', Array.from(selectedFileIds));
        if (debug) console.debug('buildResults: files considered', files.map(f => ({ id: f.id, name: f.name, rows: (f.dataRows||[]).length })));
        const headers = unionHeaders(files);
        if (debug) console.debug('buildResults: headers', headers);
        const rows = [];
        const diagnostics = []; // collect per-row parsing/match info for debug
        let processed = 0;
        files.forEach(f => {
            (f.dataRows || []).forEach(r => {
                processed++;
                try {
                    // For diagnostics: inspect each rule's parsing for this row
                    try {
                        rules.forEach(rule => {
                            if (!rule || !rule.column) return;
                            const rawVal = getValueByCanonical(r, rule.column);
                            // only interested in date window operators for now
                            if (rawVal == null) return;
                            const isWindowOp = rule.operator === 'within_days' || rule.operator === 'within_next_days';
                            if (!isWindowOp) return;
                            // parse date using the same helper used by evaluation
                            const parsedDate = parseToDate(rawVal, f._locale || 'uk');
                            const toLocalIso = (dt) => { if (!(dt instanceof Date) || isNaN(dt)) return null; const y = dt.getFullYear(); const m = String(dt.getMonth()+1).padStart(2,'0'); const d = String(dt.getDate()).padStart(2,'0'); return `${y}-${m}-${d}`; };
                            const cmpIso = parsedDate ? toLocalIso(parsedDate) : null;
                            const now = new Date(); const nowIso = toLocalIso(now);
                            let inWindow = false;
                            const days = Number(String(rule.value || '').trim());
                            if (rule.operator === 'within_days') {
                                const start = new Date(now); start.setDate(start.getDate() - (Math.floor(days) - 1)); const startIso = toLocalIso(start);
                                inWindow = cmpIso ? (cmpIso >= startIso && cmpIso <= nowIso) : false;
                            } else if (rule.operator === 'within_next_days') {
                                const end = new Date(now); end.setDate(end.getDate() + Math.floor(days)); const endIso = toLocalIso(end);
                                inWindow = cmpIso ? (cmpIso >= nowIso && cmpIso <= endIso) : false;
                            }
                            diagnostics.push({ fileId: f.id, fileName: f.name, rowIndex: processed - 1, column: rule.column, raw: rawVal, parsedIso: cmpIso, inWindow, rule: { operator: rule.operator, value: rule.value } });
                        });
                    } catch (e) { if (debug) console.debug('buildResults: diagnostics collection failed', e); }

                    if (evaluateRowAgainstRules(r, f, rules)) {
                        const computedStaffKey = getStaffKey(r);
                        // store canonical staff key on the raw row so callers/indexers can reuse it without recomputing
                        try { if (computedStaffKey) r.__staffKey = String(computedStaffKey); } catch (e) {}
                        // ...existing code...
                        const wrapper = { __fileName: f.name, __fileId: f.id, __locale: f._locale || 'uk', data: r, __staffKey: computedStaffKey };
                        rows.push(wrapper);
                    }
                } catch (e) { if (debug) console.debug('buildResults: evaluateRowAgainstRules error', e, { file: f.name, row: r }); }
            });
        });
        if (debug) console.debug('buildResults: processed rows', processed, 'matched rows', rows.length, 'sample', rows.slice(0,3));
        // Compute next upcoming Duty Date (>= today) and last duty (<= today) per person using Staff Number or Assignment No
    const nextDutyMap = {};
    const lastDutyMap = {};
    const nextDutyDetailMap = {};
    const lastDutyDetailMap = {};
    // meta maps to store priority for same-date conflicts (higher wins)
    const nextDutyMeta = {};
    const lastDutyMeta = {};
    const ignoredDuties = [];
        try {
            const toLocalIso = (dt) => { if (!(dt instanceof Date) || isNaN(dt)) return null; const y = dt.getFullYear(); const m = String(dt.getMonth()+1).padStart(2,'0'); const d = String(dt.getDate()).padStart(2,'0'); return `${y}-${m}-${d}`; };
            const today = new Date(); today.setHours(0,0,0,0);
            // scan all selected files to collect upcoming duty dates and last duties
            files.forEach(f => {
        (f.dataRows || []).forEach(r => {
                    try {
            const staffVal = getStaffKey(r);
                        if (!staffVal) return;
                const dutyRaw = getValueByCanonical(r, 'Duty Date') || getValueByCanonical(r, 'Next Duty');
                                if (!dutyRaw) return;
                                // Require Shift Type and accept only Day or Night (ignore Rest/Combined)
                                const dutyTypeVal = getValueByCanonical(r, 'Shift Type');
                                if (dutyTypeVal == null) {
                                    ignoredDuties.push({ file: f.name, rowObj: r, reason: 'Shift Type missing', dutyRaw });
                                    if (window.__ALERTS_DEBUG) console.debug('dutyMap: skipping row because Shift Type missing', { file: f.name, row: r });
                                    return;
                                }
                                const low = String(dutyTypeVal).trim().toLowerCase();
                                const tokens = low.split(/[^a-z]+/).filter(Boolean);
                                    // Accept 'combined' (prio=3), 'day' or 'night' (prio=2); skip 'rest'
                                    let prio = -1;
                                        if (/(?:\b|^)(?:ld|mgt|do)(?:\b|$)/i.test(low) || tokens.includes('rest')) prio = 0;
                                        else if (tokens.includes('combined')) prio = 3;
                                        else if (tokens.includes('day') || tokens.includes('night') || tokens.includes('n')) prio = 2;
                                        else if (tokens.includes('rest')) prio = 0;
                                    // If token not recognised, skip and record
                                    if (prio < 0) {
                                        ignoredDuties.push({ file: f.name, rowObj: r, reason: 'Shift Type not counted (not Day/Night)', dutyRaw, shift: dutyTypeVal });
                                        if (window.__ALERTS_DEBUG) console.debug('dutyMap: skipping row because Shift Type not Day/Night', { file: f.name, row: r, shiftType: dutyTypeVal, tokens });
                                        return;
                                    }
                                    // If this is a Rest (prio === 0), record it in diagnostics and skip adding it as a duty
                                    if (prio === 0) {
                                        ignoredDuties.push({ file: f.name, rowObj: r, reason: 'Rest shift', dutyRaw, shift: dutyTypeVal, staff: String(staffVal || '').trim() || null });
                                        if (window.__ALERTS_DEBUG) console.debug('dutyMap: skipping Rest shift (audited)', { file: f.name, row: r, shiftType: dutyTypeVal });
                                        return;
                                    }
                                const dt = parseToDate(dutyRaw, f._locale || 'uk');
                        if (!dt) return;
                        const d0 = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
                        const iso = toLocalIso(dt) || toLocalIso(d0);
                        if (!iso) return;
                        const key = String(staffVal).trim();
                        if (!key) return;
                                // nextDuty: earliest iso > today (strict future); prefer higher priority on same date
                                if (d0.getTime() > today.getTime()) {
                                    const meta = nextDutyMeta[key];
                                    if (!meta) {
                                        nextDutyMeta[key] = { iso, prio };
                                        nextDutyMap[key] = iso;
                                        nextDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw };
                                    } else {
                                        if (iso < meta.iso) {
                                            nextDutyMeta[key] = { iso, prio };
                                            nextDutyMap[key] = iso;
                                            nextDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw };
                                        } else if (iso === meta.iso && prio > meta.prio) {
                                            nextDutyMeta[key] = { iso, prio };
                                            nextDutyMap[key] = iso;
                                            nextDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw };
                                        }
                                    }
                                }
                                // lastDuty: latest iso <= today; prefer higher priority on same date
                                if (d0.getTime() <= today.getTime()) {
                                    const meta = lastDutyMeta[key];
                                    if (!meta) {
                                        lastDutyMeta[key] = { iso, prio };
                                        lastDutyMap[key] = iso;
                                        lastDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw };
                                    } else {
                                        if (iso > meta.iso) {
                                            lastDutyMeta[key] = { iso, prio };
                                            lastDutyMap[key] = iso;
                                            lastDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw };
                                        } else if (iso === meta.iso && prio > meta.prio) {
                                            lastDutyMeta[key] = { iso, prio };
                                            lastDutyMap[key] = iso;
                                            lastDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw };
                                        }
                                    }
                                }
                    } catch (e) { if (window.__ALERTS_DEBUG) console.debug('dutyMap: row scan failed', e); }
                });
            });
        } catch (e) { if (window.__ALERTS_DEBUG) console.debug('dutyMap: failed', e); }

                diagnostics.ignoredDuties = ignoredDuties;
                // compute per-staff RTW summary for quick lookup in the UI
                let perStaffArr = [];
                try { perStaffArr = computeRTWStats(getParsedFiles(), null) || []; } catch (e) { try { perStaffArr = (typeof window !== 'undefined' && typeof window.computeRTWLibraryStats === 'function') ? window.computeRTWLibraryStats() : []; } catch (e2) { perStaffArr = []; } }
                const perStaffMap = {};
                perStaffArr.forEach(p => { try { if (p && p.staff) perStaffMap[String(p.staff)] = p; } catch (e) {} });
                // Annotate each result row with RTW flags for fast lookup during render and details display
                try {
                    rows.forEach(wrap => {
                        try {
                            const rk = String(wrap.__staffKey || (wrap.data && wrap.data.__staffKey) || '').trim();
                            if (!rk) return;
                            let staffObj = perStaffMap[rk] || null;
                            if (!staffObj) {
                                // try numeric-only match across perStaffMap keys
                                const norm = rk.replace(/\D/g, '');
                                if (norm) {
                                    for (const k of Object.keys(perStaffMap || {})) {
                                        try {
                                            if (String(k).replace(/\D/g, '') === norm) { staffObj = perStaffMap[k]; break; }
                                        } catch (e) { /* ignore */ }
                                    }
                                }
                            }
                            if (staffObj) {
                                wrap.__rtwFlags = {
                                    rtwDone: !!staffObj.rtwDone,
                                    hadShiftAfter: !!staffObj.hadShiftAfter,
                                    continuingSickness: !!staffObj.continuingSickness
                                };
                                try { if (wrap.data) wrap.data.__rtwFlags = Object.assign({}, wrap.__rtwFlags); } catch (e) {}
                            }
                        } catch (e) { /* ignore per-row annotate errors */ }
                    });
                } catch (e) { /* ignore annotate failures */ }

                return { headers, rows, diagnostics, nextDutyMap, lastDutyMap, nextDutyDetailMap, lastDutyDetailMap, perStaffMap };
    }

    // Scan selected files for NMC pin expiry rows and return summary
    function scanNmcExpiry(files) {
        const debug = !!window.__ALERTS_DEBUG;
        const results = { total: 0, within30: [], within7: [] };
        try {
            const now = new Date(); now.setHours(0,0,0,0);
            const inDays = (dt) => { const d = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()); const diff = Math.round((d - now) / (1000*60*60*24)); return diff; };
            files.forEach(f => {
                (f.dataRows || []).forEach(r => {
                    try {
                        // permissive detection: require at least Surname and Valid To or Date Gained/Valid To
                        const surname = getValueByCanonical(r, 'Surname') || getValueByCanonical(r, 'Last Name') || getValueByCanonical(r, 'Family Name');
                        const forenames = getValueByCanonical(r, 'Forenames') || getValueByCanonical(r, 'Forename') || getValueByCanonical(r, 'First Name');
                        const staffNo = getStaffKey(r) || getValueByCanonical(r, 'Staff No') || getValueByCanonical(r, 'Employee Number');
                        const validToRaw = getValueByCanonical(r, 'Valid To') || getValueByCanonical(r, 'Expiry Date') || getValueByCanonical(r, 'ValidTo');
                        if (!surname || !validToRaw) return;
                        const dt = parseToDate(validToRaw, f._locale || 'uk');
                        if (!dt) return;
                        const days = inDays(dt);
                        const entry = { file: f.name, surname: String(surname).trim(), forenames: String(forenames||'').trim(), staffNumber: String(staffNo||'').trim(), validTo: toLocalIso(dt), daysUntil: days };
                        results.total += 1;
                        if (days <= 30 && days >= 0) results.within30.push(entry);
                        if (days <= 7 && days >= 0) results.within7.push(entry);
                    } catch (e) { if (debug) console.debug('scanNmcExpiry: row scan failed', e); }
                });
            });
            // sort lists by ascending daysUntil
            results.within30.sort((a,b) => a.daysUntil - b.daysUntil);
            results.within7.sort((a,b) => a.daysUntil - b.daysUntil);
        } catch (e) { if (debug) console.debug('scanNmcExpiry failed', e); }
        return results;
    }

    function renderRulesEditor(files, headers) {
        const rulesContainer = tabContent.querySelector('#alerts-rules');
        if (!rulesContainer) return;
        rulesContainer.innerHTML = '';
        rules.forEach((rule, idx) => {
            const el = document.createElement('div');
            el.className = 'p-2 border rounded mb-2 flex gap-2 items-center';
            const colSel = document.createElement('select'); colSel.className = 'p-1 rounded border text-sm';
            headers.forEach(h => { const o = document.createElement('option'); o.value = h; o.textContent = h; if (h === rule.column) o.selected = true; colSel.appendChild(o); });
            const opSel = document.createElement('select'); opSel.className = 'p-1 rounded border text-sm';
            // Friendly operator labels for non-technical users. Keep the option value as the internal operator code.
            const operatorDefs = [
                { code: 'contains', label: 'Includes' },
                { code: 'not_contains', label: "Doesn't include" },
                { code: 'equals', label: 'Is exactly' },
                { code: 'is_blank', label: 'Is blank' },
                { code: 'is_not_blank', label: 'Is not blank' },
                { code: 'within_days', label: 'Within last N days' },
                { code: 'within_next_days', label: 'Within next N days' },
                { code: 'numeric_gt', label: 'Greater than' },
                { code: 'numeric_gte', label: 'Greater than or equal to' },
                { code: 'numeric_lt', label: 'Less than' },
                { code: 'numeric_lte', label: 'Less than or equal to' },
            ];
            operatorDefs.forEach(op => { const o = document.createElement('option'); o.value = op.code; o.textContent = op.label; if (op.code === rule.operator) o.selected = true; opSel.appendChild(o); });
            const valInp = document.createElement('input'); valInp.type = 'text'; valInp.value = rule.value || ''; valInp.className = 'p-1 rounded border flex-1 text-sm';
            // placeholder hint for special operators
            valInp.placeholder = (rule.operator === 'within_days' || rule.operator === 'within_next_days') ? 'days (e.g. 7)' : '';
            const del = document.createElement('button'); del.className = 'px-2 py-1 text-red-600'; del.textContent = 'Remove';
            el.appendChild(colSel); el.appendChild(opSel); el.appendChild(valInp); el.appendChild(del);
            rulesContainer.appendChild(el);
            colSel.addEventListener('change', async (e) => { rule.column = e.target.value; updateAlerts(); await saveAllRules(); });
            opSel.addEventListener('change', async (e) => {
                rule.operator = e.target.value;
                // update placeholder / enabled state for operators that expect a value
                if (rule.operator === 'within_days' || rule.operator === 'within_next_days') {
                    valInp.placeholder = 'days (e.g. 7)';
                    valInp.disabled = false;
                    valInp.style.opacity = '';
                } else if (rule.operator === 'is_blank' || rule.operator === 'is_not_blank') {
                    // these operators don't use the value input
                    valInp.placeholder = '';
                    valInp.value = '';
                    valInp.disabled = true;
                    valInp.style.opacity = '0.6';
                } else {
                    valInp.placeholder = '';
                    valInp.disabled = false;
                    valInp.style.opacity = '';
                }
                updateAlerts();
                await saveAllRules();
            });
            valInp.addEventListener('input', async (e) => { rule.value = e.target.value; updateAlerts(); await saveAllRules(); });
            del.addEventListener('click', async () => {
                const removed = rules.splice(idx,1)[0];
                try {
                    if (removed && removed.id) await dbDelete(STORES.RULES, removed.id);
                } catch(e){ console.warn('Failed to delete saved rule', e); }
                try { await saveAllRules(); } catch(e){ console.warn('Failed to persist rules after delete', e); }
                renderAlerts();
            });
        });
    }

    function updateAlerts() {
        const debug = !!window.__ALERTS_DEBUG;
        const files = getParsedFiles();
        const selectedFiles = files.filter(f => selectedFileIds.has(f.id));
        const headers = unionHeaders(selectedFiles);
        if (debug) {
            console.debug('updateAlerts: selectedFiles', selectedFiles.map(f => ({ id: f.id, name: f.name }))); 
            console.debug('updateAlerts: rules', rules);
            console.debug('updateAlerts: headers', headers);
        }
        // update rules editor
        renderRulesEditor(selectedFiles, headers);
        // helper: visible columns selection (persisted in localStorage)
        function getVisibleColumns(allHeaders) {
            try {
                const raw = localStorage.getItem('TERRA_ALERTS_VISIBLE_COLS');
                if (!raw) return allHeaders;
                const arr = JSON.parse(raw);
                if (!Array.isArray(arr) || arr.length === 0) return allHeaders;
                // validate entries exist in allHeaders; if not, fall back to allHeaders
                const filtered = arr.filter(a => allHeaders.includes(a));
                return filtered.length ? filtered : allHeaders;
            } catch (e) { return allHeaders; }
        }

        function showColumnsModal(allHeaders) {
            const id = 'alerts-columns-modal';
            let modal = document.getElementById(id);
            if (!modal) {
                modal = document.createElement('div'); modal.id = id; modal.className = 'modal-backdrop';
                modal.innerHTML = `<div class="modal"><div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Choose columns</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div><div class="modal-body p-2"><div id="alerts-columns-list" class="space-y-2"></div></div><div class="modal-footer mt-3 text-right"><button id="alerts-columns-save" class="px-3 py-1 rounded bg-blue-600 text-white">Save</button><button id="alerts-columns-cancel" class="ml-2">Cancel</button></div></div>`;
                document.body.appendChild(modal);
                modal.querySelector('.modal-close-btn').addEventListener('click', () => modal.classList.remove('active'));
                modal.querySelector('#alerts-columns-cancel').addEventListener('click', () => modal.classList.remove('active'));
                modal.querySelector('#alerts-columns-save').addEventListener('click', () => {
                    const inputs = Array.from(modal.querySelectorAll('#alerts-columns-list input[type=checkbox]'));
                    const picks = inputs.filter(i=>i.checked).map(i=>i.dataset.h);
                    localStorage.setItem('TERRA_ALERTS_VISIBLE_COLS', JSON.stringify(picks));
                    modal.classList.remove('active');
                    try { updateAlerts(); } catch (e) { /* ignore */ }
                });
            }
            const list = modal.querySelector('#alerts-columns-list');
            list.innerHTML = '';
            const curr = getVisibleColumns(allHeaders);
            allHeaders.forEach(h => {
                const row = document.createElement('div'); row.className = 'flex items-center gap-2';
                const cb = document.createElement('input'); cb.type = 'checkbox'; cb.dataset.h = h; cb.checked = curr.includes(h);
                const lbl = document.createElement('label'); lbl.textContent = h; lbl.className = 'text-sm';
                row.appendChild(cb); row.appendChild(lbl); list.appendChild(row);
            });
            // no width inputs in this modal; column widths are controlled by drag handles in the table
            modal.classList.add('active');
        }
        // build results (robust): prefer buildResults when available.
        // If no rules are defined, do NOT call buildResults (it treats empty rules as match-all)
        // and instead synthesize a minimal result but keep per-staff RTW stats so summaries work.
        let res = null;
        try {
            if (!rules || rules.length === 0) {
                // synthesize minimal result when there are no rules: do not compute matched rows
                try {
                    // Use a cached per-staff map if available to avoid blocking work during initial render
                    const cached = (typeof window !== 'undefined' && window.__RTW_PER_STAFF_CACHE) ? window.__RTW_PER_STAFF_CACHE : null;
                    if (cached) {
                        res = { headers: [], rows: [], diagnostics: { ignoredDuties: [] }, nextDutyMap: {}, lastDutyMap: {}, nextDutyDetailMap: {}, lastDutyDetailMap: {}, perStaffMap: cached };
                    } else {
                        // create an empty result quickly and schedule a background compute to populate perStaffMap
                        res = { headers: [], rows: [], diagnostics: { ignoredDuties: [] }, nextDutyMap: {}, lastDutyMap: {}, nextDutyDetailMap: {}, lastDutyDetailMap: {}, perStaffMap: {} };
                        const computeAndCache = () => {
                            try {
                                const perArr = (typeof computeRTWLibraryStats === 'function') ? computeRTWLibraryStats({ includeSources: true }) : [];
                                const perMap = {};
                                (perArr || []).forEach(p => { try { if (p && p.staff) perMap[String(p.staff)] = p; } catch (e) {} });
                                try { window.__RTW_PER_STAFF_CACHE = perMap; } catch (e) {}
                                // if the alerts tab is still visible, refresh the UI to show updated summaries
                                try { const isAlertsVisible = document.querySelector('#tab-alerts') && document.getElementById('tab-alerts').classList.contains('bg-white'); if (isAlertsVisible) updateAlerts(); } catch (e) {}
                            } catch (e) { if (window.__ALERTS_DEBUG) console.debug('background computeRTWLibraryStats failed', e); }
                        };
                        if (typeof requestIdleCallback === 'function') requestIdleCallback(computeAndCache, { timeout: 2000 }); else setTimeout(computeAndCache, 200);
                    }
                    if (window.__ALERTS_DEBUG) console.debug('updateAlerts: synthesized minimal res because no rules defined (fast path)');
                } catch (e) {
                    res = { headers: [], rows: [], diagnostics: { ignoredDuties: [] }, nextDutyMap: {}, lastDutyMap: {}, nextDutyDetailMap: {}, lastDutyDetailMap: {}, perStaffMap: {} };
                }
            } else {
                if (typeof buildResults === 'function') res = buildResults();
            }
        } catch (e) {
            if (window.__ALERTS_DEBUG) console.debug('buildResults call failed, will fallback to library scan', e);
            res = null;
        }
        // If buildResults didn't produce a usable result, synthesize a minimal result using computeRTWLibraryStats
        if (!res) {
            try {
                // Use cached per-staff map when available to avoid blocking
                const cached = (typeof window !== 'undefined' && window.__RTW_PER_STAFF_CACHE) ? window.__RTW_PER_STAFF_CACHE : null;
                if (cached) {
                    res = { headers: [], rows: [], diagnostics: { ignoredDuties: [] }, nextDutyMap: {}, lastDutyMap: {}, nextDutyDetailMap: {}, lastDutyDetailMap: {}, perStaffMap: cached };
                } else {
                    // create a minimal empty result and schedule a deferred compute to populate the cache
                    res = { headers: [], rows: [], diagnostics: { ignoredDuties: [] }, nextDutyMap: {}, lastDutyMap: {}, nextDutyDetailMap: {}, lastDutyDetailMap: {}, perStaffMap: {} };
                    try {
                        setTimeout(() => {
                            try {
                                const perArr = (typeof computeRTWLibraryStats === 'function') ? computeRTWLibraryStats({ includeSources: true }) : [];
                                const perMap = {};
                                (perArr || []).forEach(p => { try { if (p && p.staff) perMap[String(p.staff)] = p; } catch (e) {} });
                                try { window.__RTW_PER_STAFF_CACHE = perMap; } catch (e) {}
                                try { updateAlerts(); } catch (e) {}
                            } catch (e) { if (window.__ALERTS_DEBUG) console.debug('deferred computeRTWLibraryStats failed', e); }
                        }, 800);
                    } catch (e) { /* ignore schedule failure */ }
                }
            } catch (e) {
                // final fallback: empty result
                res = { headers: [], rows: [], diagnostics: { ignoredDuties: [] }, nextDutyMap: {}, lastDutyMap: {}, nextDutyDetailMap: {}, lastDutyDetailMap: {}, perStaffMap: {} };
            }
        } else {
            // Ensure perStaffMap exists; attach fallback map when missing or empty
            try {
                const hasPer = res.perStaffMap && Object.keys(res.perStaffMap).length;
                if (!hasPer) {
                    const perArr = (typeof computeRTWLibraryStats === 'function') ? computeRTWLibraryStats({ includeSources: true }) : [];
                    const perMap = {};
                    (perArr || []).forEach(p => { try { if (p && p.staff) perMap[String(p.staff)] = p; } catch (e) {} });
                    res.perStaffMap = perMap;
                    if (window.__ALERTS_DEBUG) console.debug('attached fallback perStaffMap from computeRTWLibraryStats', Object.keys(perMap).length);
                }
            } catch (e) { /* ignore per-staff fallback errors */ }
        }
// New computeRTWStats implementation - replace the existing function body with this
function computeRTWStats(files, rows) {
    const sicknessCols = ['Sickness End','Sickness End Date','Sick End Date','SicknessEnd','Sick End','Sickness_End','End','End Date'];
    const dutyCols = ['Duty Date','DutyDate','Duty','Shift Date','ShiftDate','Next Duty','NextDuty','Next Duty Date','NextDutyDate'];
    const rtwCols = ['Return to Work','RTW','Return to work','Return','ReturnToWork','Return To Work Interview Completed','Return To Work Interview','RTW Interview Completed','RTW Interview','RTW Date','Return to Work Date'];
    const staffKeyFn = (r) => getStaffKey(r) || getValueByCanonical(r, 'Assignment No') || getValueByCanonical(r, 'Staff');

    const staffMap = {};
    const feedRows = [];

    // Build feedRows with rowIndex so we can link flags -> ranges by source
    if (Array.isArray(files) && files.length) {
        files.forEach(f => {
            (f.dataRows || []).forEach((r, idx) => {
                feedRows.push({ data: r, __locale: f._locale || 'uk', __fileId: f.id, __fileName: f.name, __rowIndex: idx });
            });
        });
    } else if (Array.isArray(rows) && rows.length) {
        // rows are wrappers (as returned by buildResults) — preserve any __rowIndex if present
        rows.forEach(r => feedRows.push({ ...r, __rowIndex: (r.__rowIndex != null ? r.__rowIndex : null) }));
    }

    // collect sickness ranges, duties and RTW flag objects (flags include parsedDate & source)
    const sicknessStartCols = ['Sickness Start','Sickness Start Date','Sick Start','Start','Start Date','SicknessStart'];

    feedRows.forEach(r => {
        const staff = String(staffKeyFn(r.data) || '').trim();
        if (!staff) return;
        staffMap[staff] = staffMap[staff] || { staff, sicknessRanges: [], dutyDates: [], rtwFlags: [] };
        const entry = staffMap[staff];

        // collect sickness start/end (prefer explicit start+end columns). Keep source reference.
        let startRaw = null, endRaw = null;
        for (const sc of sicknessCols) { const raw = getValueByCanonical(r.data, sc); if (raw != null) { endRaw = raw; break; } }
        for (const ss of sicknessStartCols) { const raw = getValueByCanonical(r.data, ss); if (raw != null) { startRaw = raw; break; } }
        try {
            const startDt = startRaw != null ? parseToDate(startRaw, r.__locale || 'uk') : null;
            const endDt = endRaw != null ? parseToDate(endRaw, r.__locale || 'uk') : null;
            const reasonCols = ['Reason','Sickness Reason','Absence Reason','Notes','Description'];
            let reasonVal = null;
            for (const rc of reasonCols) {
                const rv = getValueByCanonical(r.data, rc);
                if (rv != null && String(rv || '').trim() !== '') { reasonVal = String(rv).trim(); break; }
            }
            const source = { fileId: r.__fileId, fileName: r.__fileName, rowIndex: r.__rowIndex };

            if (endDt && !startDt) entry.sicknessRanges.push({ start: endDt, end: endDt, reason: reasonVal, sources: [source] });
            else if (startDt && !endDt) entry.sicknessRanges.push({ start: startDt, end: startDt, reason: reasonVal, sources: [source] });
            else if (startDt && endDt) entry.sicknessRanges.push({ start: startDt, end: endDt, reason: reasonVal, sources: [source] });
        } catch (e) {
            for (const sc of sicknessCols) {
                const raw = getValueByCanonical(r.data, sc);
                if (raw != null) {
                    const dt = parseToDate(raw, r.__locale || 'uk');
                    if (dt) entry.sicknessRanges.push({ start: dt, end: dt, reason: null, sources: [{ fileId: r.__fileId, fileName: r.__fileName, rowIndex: r.__rowIndex }] });
                    break;
                }
            }
        }

        // collect duty dates (skip Rest rows)
        for (const dc of dutyCols) {
            const raw = getValueByCanonical(r.data, dc);
            if (raw != null) {
                let dutyTypeVal = getValueByCanonical(r.data, 'Shift Type');
                if (dutyTypeVal == null) {
                    const tryAlts = ['ShiftType','Type','Roster Type','Assignment Info','Assignment'];
                    for (const alt of tryAlts) { const v = getValueByCanonical(r.data, alt); if (v != null) { dutyTypeVal = v; break; } }
                }
                let prio = -1;
                if (dutyTypeVal != null) {
                    const low = String(dutyTypeVal).trim().toLowerCase();
                    const tokens = low.split(/[^a-z]+/).filter(Boolean);
                    if (tokens.includes('combined')) prio = 3;
                    else if (tokens.includes('day') || tokens.includes('night')) prio = 2;
                    else if (tokens.includes('rest')) prio = 0;
                }
                if (prio < 0) {
                    try {
                        for (const key of Object.keys(r.data || {})) {
                            const s = (r.data[key] == null ? '' : (typeof r.data[key] === 'string' ? r.data[key] : JSON.stringify(r.data[key])));
                            if (!s) continue;
                            if (/\brest\b/i.test(s)) { prio = 0; break; }
                        }
                    } catch (e) { /* ignore */ }
                }
                if (prio < 0) break;
                if (prio === 0) break;
                const dt = parseToDate(raw, r.__locale || 'uk');
                if (dt) {
                    const iso = toLocalIso(dt);
                    entry.dutyDates.push({ date: dt, iso, prio, file: r.__fileName });
                }
                break;
            }
        }

        // collect RTW flags: store as objects with parsedDate (if available) + source
        for (const rc of rtwCols) {
            const raw = getValueByCanonical(r.data, rc);
            if (raw != null) {
                // attempt to find an associated date for the RTW flag: RTW Date column, or this row's sickness end, or duty date
                let dateRaw = getValueByCanonical(r.data, 'RTW Date') || getValueByCanonical(r.data, 'Return to Work Date') || endRaw;
                if (!dateRaw) {
                    for (const dc of dutyCols) {
                        const rd = getValueByCanonical(r.data, dc);
                        if (rd != null) { dateRaw = rd; break; }
                    }
                }
                const parsedDate = dateRaw ? parseToDate(dateRaw, r.__locale || 'uk') : null;
                entry.rtwFlags.push({
                    value: String(raw || '').trim(),
                    parsedDate: parsedDate,
                    iso: parsedDate ? toLocalIso(parsedDate) : null,
                    source: { fileId: r.__fileId, fileName: r.__fileName, rowIndex: r.__rowIndex }
                });
                break;
            }
        }
    });

    // Evaluate per-staff
    const results = [];
    const today = new Date(); today.setHours(0,0,0,0);

    Object.values(staffMap).forEach(s => {
        const requireReasonMatch = (typeof localStorage !== 'undefined' && localStorage.getItem('TERRA_MERGE_SICKNESS_BY_REASON') === 'true');
        const rawRanges = (s.sicknessRanges || []).filter(rr => rr && rr.start && rr.end).map(rr => ({
            start: new Date(rr.start.getFullYear(), rr.start.getMonth(), rr.start.getDate()),
            end: new Date(rr.end.getFullYear(), rr.end.getMonth(), rr.end.getDate()),
            reason: rr.reason != null ? String(rr.reason).trim() : null,
            sources: rr.sources || []
        }));
        rawRanges.sort((a,b) => a.start.getTime() - b.start.getTime());

        // merge adjacent/overlapping ranges while preserving sources
        const merged = [];
        for (const r0 of rawRanges) {
            if (!merged.length) { merged.push({ start: new Date(r0.start), end: new Date(r0.end), count:1, reason: r0.reason || null, sources: (r0.sources || []).slice() }); continue; }
            const last = merged[merged.length - 1];
            const nextDayAfterLast = new Date(last.end.getTime()); nextDayAfterLast.setDate(nextDayAfterLast.getDate() + 1);
            const lastReason = last.reason != null ? String(last.reason).trim() : null;
            const thisReason = r0.reason != null ? String(r0.reason).trim() : null;
            if (r0.start.getTime() <= nextDayAfterLast.getTime() && (!requireReasonMatch || lastReason === thisReason)) {
                if (r0.end.getTime() > last.end.getTime()) last.end = new Date(r0.end);
                last.count = (last.count || 1) + 1;
                last.sources = (last.sources || []).concat(r0.sources || []);
            } else {
                merged.push({ start: new Date(r0.start), end: new Date(r0.end), count: 1, reason: r0.reason || null, sources: (r0.sources || []).slice() });
            }
        }

        // latest merged range considered the 'current' sickness episode
        const latestRange = merged.length ? merged[merged.length - 1] : null;
        const sickness = latestRange ? new Date(latestRange.end.getTime()) : null;

        // process dutyDates (keep highest priority per date)
        let processedDuties = [];
        if (s.dutyDates.length) {
            const byIso = {};
            s.dutyDates.forEach(dd => {
                const iso = dd.iso || (dd instanceof Date ? toLocalIso(dd) : null);
                if (!iso) return;
                if (!byIso[iso] || (dd.prio != null && dd.prio > byIso[iso].prio)) byIso[iso] = dd;
            });
            processedDuties = Object.values(byIso).map(x => x.date).sort((a,b) => a.getTime() - b.getTime());
        }

        // nextDuty and lastDuty as before
        const nextDuty = processedDuties.find(d => {
            const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate());
            return d0.getTime() >= today.getTime();
        }) || null;
        const lastOnOrBefore = processedDuties.filter(d => {
            const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate());
            return d0.getTime() <= today.getTime();
        });
        const lastDuty = lastOnOrBefore.length ? lastOnOrBefore[lastOnOrBefore.length - 1] : null;

        // helper to identify yes-like RTW flag values
        const yesMatches = (f) => {
            if (!f || !f.value) return false;
            const v = String(f.value || '').toLowerCase();
            return v === 'y' || v === 'yes' || v === 'true' || v === '1';
        };

        // determine if an RTW flag belongs to the given merged range
        const flagBelongsToRange = (flag, range) => {
            if (!flag) return false;
            // Prefer date association when present. Allow a short post-return window.
            if (flag.parsedDate instanceof Date && !isNaN(flag.parsedDate)) {
                const fd = new Date(flag.parsedDate.getFullYear(), flag.parsedDate.getMonth(), flag.parsedDate.getDate());
                const POST_WINDOW_DAYS = 7; // configurable window after sickness end to accept RTW recorded shortly after return
                const windowEnd = new Date(range.end.getTime() + (POST_WINDOW_DAYS * 24*60*60*1000));
                return fd.getTime() >= range.start.getTime() && fd.getTime() <= windowEnd.getTime();
            }
            // If no usable date, accept the flag if it comes from a source row that contributed to this range
            try {
                if (flag.source && Array.isArray(range.sources) && range.sources.length) {
                    return range.sources.some(src => (src && src.fileId === flag.source.fileId && (src.rowIndex == null || flag.source.rowIndex == null || src.rowIndex === flag.source.rowIndex)));
                }
            } catch (e) { /* ignore */ }
            // Conservative fallback: do not attribute the flag to the range
            return false;
        };

        // compute rtwDone only considering flags that belong to the latestRange
        let rtwDoneForLatest = false;
        if (latestRange) {
            for (const flag of (s.rtwFlags || [])) {
                try {
                    if (!yesMatches(flag)) continue;
                    if (flagBelongsToRange(flag, latestRange)) { rtwDoneForLatest = true; break; }
                } catch (e) { /* ignore per-flag errors */ }
            }
        } else {
            // fallback: no sickness range detected — use broad heuristic (any yes-like RTW flag)
            rtwDoneForLatest = (s.rtwFlags || []).some(f => yesMatches(f));
        }

        // remaining semantics preserved
        const sicknessEnded = !!(sickness && sickness.getTime() <= today.getTime());
        const hadShiftAfter = sickness && lastDuty && lastDuty.getTime() >= sickness.getTime();
        const hadShiftAfterSickness = !!hadShiftAfter;
        const rtwInterviewDone = !!rtwDoneForLatest;
        const rawMultiple = Array.isArray(rawRanges) && rawRanges.length > 1;
        const continuingSickness = !!(latestRange && (latestRange.end.getTime() >= today.getTime() || (latestRange.count && latestRange.count > 1))) || rawMultiple;
        const mergedRangesOut = merged.map(mr => ({ start: toLocalIso(mr.start), end: toLocalIso(mr.end), parts: mr.count || 1, reason: (mr.reason != null ? String(mr.reason) : null) }));

        results.push({
            staff: s.staff,
            sickness,
            sicknessEnded,
            hadShiftAfterSickness,
            hadShiftAfter: !!hadShiftAfter,
            rtwInterviewDone,
            rtwDone: !!rtwDoneForLatest,
            sampleDuty: processedDuties.length ? processedDuties[0] : null,
            nextDuty,
            lastDuty,
            continuingSickness,
            sicknessRanges: mergedRangesOut
        });
    });

    return results;
}

    // Compute RTW stats across the whole library (all parsed files), not only matched alert rows
    const stats = computeRTWStats(getParsedFiles(), /*rows*/ null);
    const hadShift = stats.filter(s => s.hadShiftAfter);
    const noRtw = hadShift.filter(s => !s.rtwDone);
    const completed = hadShift.length - noRtw.length;
    const completedPct = hadShift.length ? Math.round((completed / hadShift.length) * 100) : 0;
    const incompletePct = hadShift.length ? Math.round((noRtw.length / hadShift.length) * 100) : 0;
        const resultsEl = tabContent.querySelector('#alerts-results');
        if (!resultsEl) return;
    // Render table (add Next Duty column) and respect user-visible columns
    // include 'Source File' in the available headers so user can hide it
    const visibleHeaders = getVisibleColumns(['Source File', ...res.headers]);
    const cols = [...visibleHeaders, 'Last Duty', 'Next Duty'];
    // Helper to compute RTW flags for a row (lazy)
    function computeRowRTWFlags(row) {
        const staffKey = String(row.__staffKey || (row.data && row.data.__staffKey) || '').trim();
        let staffObj = res && res.perStaffMap && staffKey ? res.perStaffMap[staffKey] : null;
        if (!staffObj && staffKey && res && res.perStaffMap) {
            // try numeric-only match
            const norm = staffKey.replace(/\D/g, '');
            if (norm) {
                for (const k of Object.keys(res.perStaffMap)) {
                    if (String(k).replace(/\D/g, '') === norm) { staffObj = res.perStaffMap[k]; break; }
                }
            }
        }
        return {
            rtwDone: !!(staffObj && staffObj.rtwDone),
            hadShiftAfter: !!(staffObj && staffObj.hadShiftAfter),
            continuingSickness: !!(staffObj && staffObj.continuingSickness)
        };
    }
        // render summary above the table
        resultsEl.innerHTML = '';
    // Determine status color by thresholds: >=85 green, 65-84 amber, <65 red
    let statusColor = '#dc2626'; // red
    let statusLabel = 'Low';
    if (completedPct >= 85) { statusColor = '#16a34a'; statusLabel = 'Good'; }
    else if (completedPct >= 65) { statusColor = '#f59e0b'; statusLabel = 'Warning'; }
    const pill = `<span style="display:inline-block;width:10px;height:10px;border-radius:999px;background:${statusColor};margin-right:8px;vertical-align:middle"></span>`;
    // compute ignored duties counts by reason for a small badge
    const ignoredList = (res && res.diagnostics && res.diagnostics.ignoredDuties) || [];
    const ignoredCounts = ignoredList.reduce((acc, it) => { const k = (it && it.reason) ? String(it.reason) : 'Unknown'; acc[k] = (acc[k] || 0) + 1; return acc; }, {});
    const ignoredTotal = ignoredList.length;
    // expose to console for debugging
    try {
        if (typeof window !== 'undefined') {
            // If a library-scanning helper already exists (defined earlier), don't overwrite it.
            // Keep a UI-specific view of the ignored list available as a separate helper.
            try {
                if (typeof window.__alertsIgnoredDuties !== 'function') {
                    // no existing helper — expose the current UI list as the default
                    window.__alertsIgnoredDuties = () => ignoredList;
                }
            } catch (e) { /* ignore */ }
            // always expose the UI-provided ignored list separately
            window.__alertsIgnoredDutiesFromUI = () => ignoredList;
            // helpful printers
            try { window.__alertsIgnoredDuties.print = () => console.log('Ignored duties (library scan or UI):', window.__alertsIgnoredDuties()); } catch (e) { /* ignore */ }
            try { window.__alertsIgnoredDutiesFromUI.print = () => console.log('Ignored duties (from UI):', ignoredList); } catch (e) { /* ignore */ }
        }
    } catch (e) { /* ignore */ }

    const summary = document.createElement('div'); summary.className = 'p-3 mb-3 rounded border bg-yellow-50 text-sm';
    const ignoredBadge = ignoredTotal ? `<span title="Ignored duties" style="display:inline-block;background:#ef4444;color:white;padding:2px 6px;border-radius:999px;margin-left:8px;font-size:0.8em">${ignoredTotal}</span>` : '';
    const ignoredBreakdown = ignoredTotal ? ` (${Object.entries(ignoredCounts).map(([k,v])=>`${k}: ${v}`).join(', ')})` : '';
    const urgentCount = noRtw.length;

    summary.innerHTML = `
    <div role="region" aria-label="RTW summary" style="display:flex;gap:10px;align-items:center;flex-wrap:nowrap;min-width:420px;padding:6px;border-radius:10px;border:1px solid rgba(0,0,0,0.06);background:var(--color-surface);box-sizing:border-box">
        <!-- Left card: Title + two metric tiles -->
        <div style="display:flex;align-items:center;gap:12px;padding:8px 10px;border-radius:8px;background:transparent;min-width:300px;box-sizing:border-box;height:64px">
        <div style="display:flex;flex-direction:column;justify-content:center;align-items:flex-start;min-width:160px">
            <div style="display:flex;align-items:center;gap:8px;">
            <div aria-hidden="true" style="width:10px;height:10px;border-radius:999px;background:${statusColor || '#f59e0b'}"></div>
            <div style="font-size:13px;font-weight:600;line-height:1;color:var(--color-text)">${escapeHtml('RTW Summary')}</div>
            <div style="margin-left:6px;font-size:11px;color:var(--color-text-muted)">${escapeHtml(String(statusLabel))}</div>
            </div>
            <div style="display:flex;gap:8px;margin-top:6px">
            <div style="display:flex;flex-direction:column;justify-content:center;align-items:flex-start;padding:4px 8px;border-radius:6px;min-width:90px">
                <div style="font-weight:700;font-size:14px;color:var(--color-text)">${hadShift.length}</div>
                <div style="font-size:11px;color:var(--color-text-muted)">had a shift</div>
            </div>
            <div style="display:flex;flex-direction:column;justify-content:center;align-items:flex-start;padding:4px 8px;border-radius:6px;min-width:120px">
                <div style="display:flex;align-items:center;gap:8px">
                <div style="font-weight:700;font-size:14px;color:var(--color-text)">${completed}</div>
                <div style="font-size:12px;color:var(--color-text-muted)">RTW</div>
                <div style="margin-left:6px;font-size:11px;padding:2px 6px;border-radius:999px;background:rgba(0,0,0,0.06);color:var(--color-text);font-weight:600">${completedPct}%</div>
                </div>
                <!-- subtle inline progress bar -->
                <div style="margin-top:6px;width:100%;height:6px;background:rgba(0,0,0,0.06);border-radius:6px;overflow:hidden">
                <div style="width:${Math.max(0, Math.min(100, Number(completedPct) || 0))}%;height:100%;background:${(completedPct >= 85) ? '#16a34a' : (completedPct >= 65) ? '#f59e0b' : '#ef4444'};border-radius:6px"></div>
                </div>
            </div>
            </div>
        </div>
        </div>

        <!-- Right card: Missing RTW -->
        <div style="display:flex;flex-direction:column;justify-content:center;align-items:center;padding:8px 12px;border-radius:8px;background:#fff7f7;min-width:120px;box-sizing:border-box;height:64px">
        <div style="font-size:11px;color:var(--color-text-muted);margin-bottom:4px">Missing RTW</div>
        <div style="display:flex;align-items:center;gap:8px">
            <div style="width:34px;height:34px;border-radius:999px;display:flex;align-items:center;justify-content:center;background:#fff;border:1px solid rgba(0,0,0,0.06);font-weight:700;color:#b91c1c">${urgentCount}</div>
            <div style="font-size:12px;color:var(--color-text-muted)">out of <strong>${hadShift.length}</strong></div>
        </div>
        </div>

        <!-- Actions -->
        <div style="margin-left:auto;display:flex;gap:8px;align-items:center">
        <button id="alerts-show-no-rtw" class="px-2 py-1 border rounded text-sm" style="font-size:12px;padding:6px 8px">Show list</button>
        </div>
    </div>
    `;

    // Replace previous updateNmcPinsCard IIFE with this one (place immediately after summary.innerHTML)
    (function updateNmcPinsCard() {
    try {
        const nmcEl = (typeof resultsEl !== 'undefined' && resultsEl && resultsEl.querySelector)
        ? resultsEl.querySelector('.p-3.mb-3.rounded.border.bg-yellow-50') || resultsEl.querySelector('.nmc-summary')
        : document.querySelector('.p-3.mb-3.rounded.border.bg-yellow-50') || document.querySelector('.nmc-summary');
        if (!nmcEl) return;

        // Attempt to extract a numeric count from existing content
        const rawText = (nmcEl.textContent || '').trim();
        const found = rawText.match(/(\d+)\s*(pins?|expir)/i);
        const nmcCount = found ? Number(found[1]) : (typeof window.__nmcCount !== 'undefined' ? window.__nmcCount : '—');

        // Build markup using shared classes so it matches RTW summary exactly
        nmcEl.innerHTML = `
        <div class="alert-card" style="background:#fffaf0;">
            <div class="card-left">
            <div style="display:flex;align-items:center;gap:8px;">
                <div aria-hidden="true" style="width:10px;height:10px;border-radius:999px;background:#f59e0b"></div>
                <div class="card-title">NMC Pins</div>
                <div class="card-sub">Alert</div>
            </div>
            <div style="margin-top:6px;display:flex;gap:8px;align-items:center">
                <div class="alert-pill" style="color:#b85705">${escapeHtml(String(nmcCount))}</div>
                <div style="font-size:12px;color:var(--color-text-muted)">pins expiring soon</div>
            </div>
            </div>
            <div class="alert-actions">
            <button id="nmc-expiry-show" class="px-2 py-1 border rounded text-sm" style="font-size:12px;padding:6px 8px">Show list</button>
            </div>
        </div>
        `.trim();

        // Wire the new Show list button to the existing handler if present
        const btn = nmcEl.querySelector('#nmc-expiry-show');
        if (btn) {
        btn.addEventListener('click', (ev) => {
            try {
            ev.preventDefault();
            // Prefer original existing controls if present
            const existing = document.querySelector('#nmc-expiry-export') || document.querySelector('#nmc-expiry-show-existing') || document.querySelector('#nmc-expiry-show');
            if (existing) {
                existing.click();
                return;
            }
            // Fallback: dispatch custom event for any listeners
            const evt = new CustomEvent('nmc:show', { detail: {} });
            nmcEl.dispatchEvent(evt);
            } catch (e) { /* ignore */ }
        });
        }
    } catch (e) {
        /* safe no-op */
    }
    })();

    (function ensureAlertCardsWrapper() {
  try {
    // Helper: find elements by visible text when classes vary
    const findByText = (txt) => Array.from(document.querySelectorAll('body *')).find(el => el && el.textContent && el.textContent.trim().includes(txt));

    // Locate the RTW summary node more robustly — search for visible string 'RTW Summary'
    const rtwNode = findByText('RTW Summary') || document.querySelector('.p-3.mb-3.rounded.border') || document.querySelector('.rtw-summary');

    // Locate NMC element (common selectors and content)
    const nmcNode = document.querySelector('.p-3.mb-3.rounded.border.bg-yellow-50') ||
                    document.querySelector('.nmc-summary') ||
                    findByText('NMC Pins') ||
                    findByText('pins expiring');

    if (!rtwNode && !nmcNode) return; // nothing to do

    // If both nodes exist and are already inside a shared wrapper with our class, ensure classes present and exit
    const existingWrapper = (rtwNode && rtwNode.closest && rtwNode.closest('.alert-cards-wrapper')) || (nmcNode && nmcNode.closest && nmcNode.closest('.alert-cards-wrapper'));
    if (existingWrapper) {
      try { if (rtwNode) rtwNode.classList.add('alert-card'); } catch(e) {}
      try { if (nmcNode) nmcNode.classList.add('alert-card'); } catch(e) {}
      return;
    }

    // Create wrapper and place it before the earliest of the two nodes
    const anchor = rtwNode || nmcNode;
    const wrapper = document.createElement('div');
    wrapper.className = 'alert-cards-wrapper';

    // Insert wrapper into DOM
    anchor.parentNode.insertBefore(wrapper, anchor);

    // Move rtwNode and nmcNode into wrapper in that order (if present)
    if (rtwNode && rtwNode.parentNode !== wrapper) wrapper.appendChild(rtwNode);
    if (nmcNode && nmcNode.parentNode !== wrapper) wrapper.appendChild(nmcNode);

    // Apply alert-card class so inner layout/pills use shared CSS
    if (rtwNode) {
      rtwNode.classList.add('alert-card');
      // ensure inner count pill uses alert-pill if present
      const pill = rtwNode.querySelector('.alert-pill');
      if (pill) pill.classList.add('alert-pill');
    }
    if (nmcNode) {
      nmcNode.classList.add('alert-card');
      // try to find numeric count and wrap/ensure pill class
      let foundPill = nmcNode.querySelector('.alert-pill');
      if (!foundPill) {
        const countMatch = (nmcNode.textContent || '').match(/(\d{1,4})/);
        if (countMatch) {
          // Create pill element and insert it near the start of the node's left area
          const pillEl = document.createElement('div');
          pillEl.className = 'alert-pill';
          pillEl.textContent = countMatch[1];
          // find sensible insertion point
          const left = nmcNode.querySelector('div') || nmcNode.firstElementChild || nmcNode;
          left.insertBefore(pillEl, left.firstChild);
        }
      } else {
        foundPill.classList.add('alert-pill');
      }
    }
  } catch (e) {
    // safe no-op - if this fails, it won't break the rest of the page
    console.warn('ensureAlertCardsWrapper failed', e);
  }
})();

    // keep `summary` element available; it will be placed into the unified alert bar below
    // wire ignored button immediately so users don't need to click 'Show list' first
        const ignoredBtnImmediate = resultsEl.querySelector('#alerts-show-ignored');
        if (ignoredBtnImmediate && !ignoredBtnImmediate.dataset.bound) {
            ignoredBtnImmediate.dataset.bound = '1';
            ignoredBtnImmediate.addEventListener('click', (e) => {
            e.preventDefault();
            const id = 'alerts-ignored-modal';
            let im = document.getElementById(id);
            if (!im) {
                im = document.createElement('div'); im.id = id; im.className = 'modal-backdrop';
                im.innerHTML = `<div class="modal"><div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Ignored duties</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div><div class="modal-body p-2"><div id="alerts-ignored-list" class="text-sm"></div></div><div class="modal-footer mt-3 text-right"><button id="alerts-ignored-export" class="px-3 py-1 mr-2 rounded bg-slate-100 text-sm">Export CSV</button><button id="alerts-ignored-close" class="px-3 py-1 py-1 rounded">Close</button></div></div>`;
                document.body.appendChild(im);
                im.querySelector('.modal-close-btn').addEventListener('click', ()=> im.classList.remove('active'));
                im.querySelector('#alerts-ignored-close').addEventListener('click', ()=> im.classList.remove('active'));
            }
                            const il = im.querySelector('#alerts-ignored-list'); il.innerHTML = '';
                            try {
                                let ig = (res && res.diagnostics && res.diagnostics.ignoredDuties) || [];
                                if ((!ig || ig.length === 0) && typeof window !== 'undefined' && typeof window.__alertsIgnoredDuties === 'function') {
                                    try { ig = window.__alertsIgnoredDuties() || []; } catch (e) { /* ignore */ }
                                }
                                if (!ig || ig.length === 0) {
                                    il.textContent = 'No ignored duties found for the current selection or library.';
                                } else {
                                    // render a table of ignored duties
                                    const tbl = document.createElement('table'); tbl.className = 'w-full text-sm border';
                                    tbl.innerHTML = '<thead><tr><th class="px-2 py-1 border-r">File</th><th class="px-2 py-1 border-r">Staff</th><th class="px-2 py-1 border-r">Reason</th><th class="px-2 py-1">Duty</th></tr></thead>';
                                    const tb = document.createElement('tbody');
                                    ig.forEach(it => {
                                        const tr = document.createElement('tr'); tr.innerHTML = `<td class="px-2 py-1 border-r">${escapeHtml(it.file||'')}</td><td class="px-2 py-1 border-r">${escapeHtml(String(it.staff||''))}</td><td class="px-2 py-1 border-r">${escapeHtml(String(it.reason||''))}</td><td class="px-2 py-1">${escapeHtml(String(it.dutyRaw||''))}</td>`;
                                        tb.appendChild(tr);
                                    });
                                    tbl.appendChild(tb);
                                    il.appendChild(tbl);
                                    // store rows on modal for export
                                    im.__exportRows = ig.map(it => ({ file: it.file||'', staff: String(it.staff||''), reason: String(it.reason||''), duty: String(it.dutyRaw||'') }));
                                    const exportBtn = im.querySelector('#alerts-ignored-export');
                                    if (exportBtn) exportBtn.addEventListener('click', () => {
                                        try {
                                            const rows = im.__exportRows || [];
                                            const csv = ['File,Staff,Reason,Duty', ...rows.map(r => `${csvSafe(r.file)},${csvSafe(r.staff)},${csvSafe(r.reason)},${csvSafe(r.duty)}`)].join('\n');
                                            const b = new Blob([csv], { type: 'text/csv' }); const url = URL.createObjectURL(b); const a = document.createElement('a'); a.href = url; a.download = 'ignored_duties.csv'; document.body.appendChild(a); try { a.click(); } catch (e) { window.open(url); } setTimeout(()=>{ try{ a.remove(); URL.revokeObjectURL(url); } catch(e){} },3000);
                                        } catch (e) { alert('Export failed'); }
                                    });
                                }
                            } catch (e) { il.textContent = 'Failed to render ignored duties'; }
            im.classList.add('active');
        });
        }

        // wire show list button
        const btn = resultsEl.querySelector('#alerts-show-no-rtw');
        if (btn && !btn.dataset.bound) {
            btn.dataset.bound = '1';
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                const modalId = 'alerts-no-rtw-modal';
                let m = document.getElementById(modalId);
                if (!m) {
                    m = document.createElement('div'); m.id = modalId; m.className = 'modal-backdrop';
                    m.innerHTML = `<div class="modal"><div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Staff with shift after sickness and no RTW</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div><div class="modal-body p-2"><div id="alerts-no-rtw-list" class="text-sm"></div></div><div class="modal-footer mt-3 text-right"><button id="alerts-no-rtw-export" class="px-3 py-1 mr-2 rounded bg-slate-100 text-sm">Export CSV</button><button id="alerts-no-rtw-close" class="px-3 py-1 rounded">Close</button></div></div>`;
                    document.body.appendChild(m);
                    m.querySelector('.modal-close-btn').addEventListener('click', ()=> m.classList.remove('active'));
                    m.querySelector('#alerts-no-rtw-close').addEventListener('click', ()=> m.classList.remove('active'));
                }
                const list = m.querySelector('#alerts-no-rtw-list'); list.innerHTML = '';
                // attempt to enrich the list with source filenames for duties if available
                let srcMap = {};
                try {
                    if (typeof window.computeRTWLibraryStats === 'function') {
                        const lib = window.computeRTWLibraryStats({ includeSources: true }) || [];
                        lib.forEach(x => { if (x && x.staff) srcMap[String(x.staff)] = x.sources || {}; });
                    }
                } catch (e) { /* ignore */ }
                // ignored duties button is wired globally above
                function findStaffName(staffId) {
                    try {
                        const files = getParsedFiles();
                        for (const f of files) {
                            for (const r of (f.dataRows || [])) {
                                const key = String(getStaffKey(r) || '').trim();
                                if (!key) continue;
                                if (key === String(staffId)) {
                                    // prefer first/given + surname/family
                                    const firstCols = ['Forename','First Name','Given Name','GivenName','Forename(s)','Name'];
                                    const lastCols = ['Surname','Last Name','Lastname','Family Name','FamilyName'];
                                    let first = null; let last = null;
                                    for (const fc of firstCols) {
                                        const fv = getValueByCanonical(r, fc);
                                        if (fv != null && String(fv).toString().trim()) { first = String(fv).trim(); break; }
                                    }
                                    for (const lc of lastCols) {
                                        const lv = getValueByCanonical(r, lc);
                                        if (lv != null && String(lv).toString().trim()) { last = String(lv).trim(); break; }
                                    }
                                    if (first && last) return `${first} ${last}`;
                                    if (first) return first;
                                    if (last) return last;
                                    // fallback: any name-like column
                                    const nameCols = ['Name','Full Name','FullName','Display Name','DisplayName'];
                                    for (const nc of nameCols) {
                                        const nv = getValueByCanonical(r, nc);
                                        if (nv != null && String(nv).toString().trim()) return String(nv).trim();
                                    }
                                }
                            }
                        }
                    } catch (e) { /* ignore */ }
                    return null;
                }

                // helper: find all rows for a staff id across parsed files
                // Improved findRowsForStaff: exact -> numeric-only -> relaxed field-scan fallback
                function findRowsForStaff(staffId) {
                    const out = [];
                    try {
                        const files = (typeof getParsedFiles === 'function') ? getParsedFiles() : (window.__getParsedFiles ? window.__getParsedFiles() : []);
                        const wanted = String(staffId || '').trim();
                        if (!wanted) return out;
                        const wantedDigits = wanted.replace(/\D/g, '');
                        const sampleKeys = [];

                        // First pass: prefer exact or numeric-only matches
                        files.forEach(f => {
                            (f.dataRows || []).forEach((r, idx) => {
                                try {
                                    const rawKey = String((r && r.__staffKey) || getStaffKey(r) || '').trim();
                                    const rawDigits = rawKey.replace(/\D/g, '');
                                    if (sampleKeys.length < 12) sampleKeys.push(rawKey || String(getStaffKey(r) || ''));
                                    if (rawKey && (rawKey === wanted || (rawDigits && wantedDigits && rawDigits === wantedDigits))) {
                                        out.push({ file: f, row: r, rowIndex: idx });
                                        return; // continue to next row
                                    }
                                    // also accept when the extractor on-the-fly matches the wanted value
                                    const alt = String(getStaffKey(r) || '').trim();
                                    const altDigits = alt.replace(/\D/g, '');
                                    if (alt && (alt === wanted || (altDigits && wantedDigits && altDigits === wantedDigits))) {
                                        out.push({ file: f, row: r, rowIndex: idx });
                                        return;
                                    }
                                } catch (e) { /* ignore per-row errors */ }
                            });
                        });

                        if (out.length) return out;

                        // Second pass: relaxed scan across all cell values (substring or digits match)
                        files.forEach(f => {
                            (f.dataRows || []).forEach((r, idx) => {
                                try {
                                    // build a string containing relevant cell text
                                    const rowText = Object.keys(r || {}).map(k => {
                                        try {
                                            const v = r[k];
                                            if (v == null) return '';
                                            if (typeof v === 'string') return v;
                                            if (typeof v === 'number') return String(v);
                                            try { return JSON.stringify(v); } catch (e) { return String(v); }
                                        } catch (e) { return ''; }
                                    }).join(' | ').toLowerCase();

                                    // direct substring match
                                    if (wanted && String(rowText).includes(String(wanted).toLowerCase())) {
                                        out.push({ file: f, row: r, rowIndex: idx });
                                        return;
                                    }

                                    // digits-only containment: e.g. cell "ID:27029932-2" should match "27029932"
                                    if (wantedDigits) {
                                        const rowDigits = String(rowText).replace(/\D/g, '');
                                        if (rowDigits && rowDigits.indexOf(wantedDigits) !== -1) {
                                            out.push({ file: f, row: r, rowIndex: idx });
                                            return;
                                        }
                                    }
                                } catch (e) { /* ignore per-row errors */ }
                            });
                        });

                        // If still empty, log diagnostics once to help debugging
                        if (!out.length) {
                            try { console.debug('findRowsForStaff: no match', { staffId, wantedDigits, filesScanned: (files||[]).length, sampleKeys }); } catch (e) {}
                        }
                    } catch (e) { /* ignore */ }
                    return out;
                }

                // Render entries compactly for non-technical users
                function renderStaffDetailsEntries(body, entries) {
                    body.innerHTML = '';
                    if (!entries || !entries.length) {
                        body.textContent = 'No rows found for this staff in the library.';
                        return;
                    }
                    entries.forEach(en => {
                        const fileName = en.file && (en.file.name || en.file.id) ? (en.file.name || en.file.id) : 'unknown source';
                        const container = document.createElement('div'); container.className = 'mb-3';
                        const hdr = document.createElement('div'); hdr.className = 'flex items-center justify-between mb-1';
                        const title = document.createElement('div'); title.className = 'font-semibold'; title.textContent = `${fileName} — row ${en.rowIndex}`;
                        hdr.appendChild(title);
                        container.appendChild(hdr);

                        const grid = document.createElement('div'); grid.className = 'detail-grid';

                        const getVal = (key) => {
                            try { const v = (typeof getValueByCanonical === 'function') ? getValueByCanonical(en.row, key) : (en.row && en.row[key] !== undefined ? en.row[key] : undefined); return v == null ? '' : String(v); } catch (e) { return ''; }
                        };

                        const addRow = (label, val) => {
                            if (val === null || val === undefined || String(val).toString().trim() === '') return;
                            const row = document.createElement('div'); row.className = 'detail-row';
                            const lab = document.createElement('div'); lab.className = 'detail-label text-subtle'; lab.textContent = label;
                            const valEl = document.createElement('div'); valEl.className = 'detail-value'; valEl.innerHTML = `<strong>${escapeHtml(String(val))}</strong>`;
                            row.appendChild(lab); row.appendChild(valEl); grid.appendChild(row);
                        };

                        // Name
                        let forename = getVal('Forename') || getVal('First Name') || getVal('Given Name') || '';
                        let surname = getVal('Surname') || getVal('Last Name') || '';

                        // Normalize to trimmed strings
                        forename = forename ? String(forename).trim() : '';
                        surname = surname ? String(surname).trim() : '';

                        let fullName = '';
                        if (forename && surname) {
                            // Avoid duplicating identical values (e.g. when both fields contain the same value)
                            if (forename === surname) fullName = forename;
                            else fullName = `${forename} ${surname}`;
                        } else if (forename) {
                            fullName = forename;
                        } else if (surname) {
                            fullName = surname;
                        } else {
                            fullName = getVal('Name') || getVal('Full Name') || '';
                        }

                        addRow('Name', (fullName && String(fullName).trim()) ? fullName : 'Unknown');

                        // Staff number / assignment
                        const staffNo = getVal('Staff Number') || getVal('Assignment No') || getVal('Assignment No.') || getVal('AssignmentNo') || '';
                        addRow('Staff number', staffNo);

                        // Duties
                        const lastDuty = getVal('Last Duty') || getVal('Duty Date') || '';
                        const nextDuty = getVal('Next Duty') || getVal('Next Duty Date') || getVal('NextDuty') || '';
                        try { addRow('Last duty', lastDuty ? formatByLocale(lastDuty, en.row && en.row.__sourceLocale ? en.row.__sourceLocale : undefined) : ''); } catch (e) { addRow('Last duty', lastDuty); }
                        try { addRow('Next duty', nextDuty ? formatByLocale(nextDuty, en.row && en.row.__sourceLocale ? en.row.__sourceLocale : undefined) : ''); } catch (e) { addRow('Next duty', nextDuty); }

                        // Sickness / RTW
                        const sickStart = getVal('Sickness Start') || getVal('Sickness') || '';
                        const sickEnd = getVal('Sickness End') || getVal('SicknessEnd') || '';
                        const rtw = getVal('Return to Work') || getVal('RTW') || '';
                        const reason = getVal('Reason') || getVal('Sickness Reason') || '';
                        try { if (sickStart) addRow('Sickness start', formatByLocale(sickStart, en.row && en.row.__sourceLocale ? en.row.__sourceLocale : undefined)); } catch(e){ if (sickStart) addRow('Sickness start', sickStart); }
                        try { if (sickEnd) addRow('Sickness end', formatByLocale(sickEnd, en.row && en.row.__sourceLocale ? en.row.__sourceLocale : undefined)); } catch(e){ if (sickEnd) addRow('Sickness end', sickEnd); }
                        addRow('RTW interview', rtw || 'Not recorded');
                        if (reason) addRow('Reason', reason);

                        // Minimal raw toggle (per entry)
                        const rawToggle = document.createElement('a'); rawToggle.href = '#'; rawToggle.className = 'text-xs text-subtle'; rawToggle.style.display = 'inline-block'; rawToggle.style.marginTop = '6px'; rawToggle.textContent = 'Show raw details';
                        const pre = document.createElement('pre'); pre.className = 'small-raw'; pre.style.display = 'none'; pre.textContent = JSON.stringify(en.row, null, 2);
                        rawToggle.addEventListener('click', (ev) => { try { ev.preventDefault(); if (pre.style.display === 'none') { pre.style.display = 'block'; rawToggle.textContent = 'Hide raw details'; } else { pre.style.display = 'none'; rawToggle.textContent = 'Show raw details'; } } catch(e){} });

                        container.appendChild(grid);
                        container.appendChild(rawToggle);
                        container.appendChild(pre);
                        body.appendChild(container);
                    });
                }

                function showStaffDetails(staffId) {
                    console.debug('showStaffDetails called', { staffId });
                    const entries = findRowsForStaff(staffId);
                    console.debug('showStaffDetails found entries count', { count: (entries && entries.length) || 0 });
                    const modalId = 'staff-detail-modal';
                    let md = document.getElementById(modalId);
                    if (!md) {
                        console.debug('showStaffDetails creating modal element');
                        md = document.createElement('div'); md.id = modalId; md.className = 'modal-backdrop';
                        md.innerHTML = `<div class="modal" style="max-width:900px"><div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Staff details</h2><div class="modal-header-actions"><button id="staff-detail-toggle-raw" class="px-2 py-1 text-sm border rounded">Show raw data</button><button id="staff-detail-copy" class="px-2 py-1 text-sm border rounded">Copy raw</button></div><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div><div class="modal-body p-2"><div id="staff-detail-body" class="text-sm"></div></div><div class="modal-footer mt-3 text-right"><button id="staff-detail-close" class="px-3 py-1 rounded">Close</button></div></div>`;
                        document.body.appendChild(md);
                        md.querySelector('.modal-close-btn').addEventListener('click', () => md.classList.remove('active'));
                        md.querySelector('#staff-detail-close').addEventListener('click', () => md.classList.remove('active'));
                        // wire header controls
                        try {
                            const toggle = md.querySelector('#staff-detail-toggle-raw');
                            const copyBtn = md.querySelector('#staff-detail-copy');
                            if (toggle) toggle.addEventListener('click', (ev) => {
                                try {
                                    ev.preventDefault();
                                    if (md.classList.contains('show-raw')) { md.classList.remove('show-raw'); toggle.textContent = 'Show raw data'; }
                                    else { md.classList.add('show-raw'); toggle.textContent = 'Hide raw data'; }
                                } catch (e) { /* ignore */ }
                            });
                            if (copyBtn) copyBtn.addEventListener('click', async (ev) => {
                                try {
                                    ev.preventDefault();
                                    const bodyEl = md.querySelector('#staff-detail-body');
                                    const pres = Array.from(bodyEl.querySelectorAll('pre')).map(p => p.textContent).join('\n\n----\n\n');
                                    if (navigator.clipboard && navigator.clipboard.writeText) await navigator.clipboard.writeText(pres);
                                    else { const ta = document.createElement('textarea'); ta.value = pres; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); ta.remove(); }
                                } catch (e) { console.warn('copy raw failed', e); }
                            });
                        } catch (e) { /* ignore */ }
                    }
                    const body = md.querySelector('#staff-detail-body'); body.innerHTML = '';
                    // Use compact renderer to display only relevant info per entry
                    try { renderStaffDetailsEntries(body, entries); } catch (e) { body.textContent = 'No rows found for this staff in the library.'; }
                    md.classList.add('active');
                    console.debug('showStaffDetails activated modal', { modalId, active: md.classList.contains('active') });
                }
                    try { if (typeof window !== 'undefined') window.showStaffDetails = showStaffDetails; } catch (e) { /* ignore */ }
                noRtw.forEach(s => {
                    const row = document.createElement('div'); row.className = 'p-1 border-b flex items-center justify-between';
                    const nextStr = s.nextDuty ? toLocalIso(s.nextDuty) : '';
                    const name = findStaffName(s.staff) || String(s.staff);
                    const left = document.createElement('div');
                    const link = document.createElement('a'); link.href = '#'; link.className = 'text-sm font-medium staff-link'; link.dataset.staff = String(s.staff); link.textContent = name;
                    link.addEventListener('click', (ev) => { ev.preventDefault(); try { scheduleShowStaffDetails(link.dataset.staff); } catch (e) { console.error('scheduleShowStaffDetails failed', e); } });
                    left.appendChild(link);
                    const right = document.createElement('div'); right.className = 'text-sm text-subtle'; right.textContent = `next duty: ${nextStr || 'N/A'}`;
                    row.appendChild(left); row.appendChild(right);
                    list.appendChild(row);
                });
                m.classList.add('active');
            });
        }
    // Add export buttons: CSV for ignored duties and PNG snapshot of the alerts table
    const exportBar = document.createElement('div'); exportBar.className = 'flex items-center gap-2 mb-3';
    const exportCsvBtn = document.createElement('button'); exportCsvBtn.className = 'px-3 py-1 rounded bg-slate-100 text-sm'; exportCsvBtn.textContent = 'Export Ignored Duties (CSV)';
    const exportSvgBtn = document.createElement('button'); exportSvgBtn.className = 'px-3 py-1 rounded bg-slate-100 text-sm'; exportSvgBtn.textContent = 'Export Alerts (SVG)';
    exportBar.appendChild(exportCsvBtn); exportBar.appendChild(exportSvgBtn);
    resultsEl.appendChild(exportBar);

    // Show NMC PIN expiry summary for selected files
    try {
        const selectedFiles = getParsedFiles().filter(f => selectedFileIds.has(f.id));
        const nmc = scanNmcExpiry(selectedFiles);
    const nmcBox = document.createElement('div'); nmcBox.className = 'p-3 mb-3 rounded border bg-yellow-50';
    const urgentCount = (nmc.within7 || []).length; const soonCount = (nmc.within30 || []).length;
    // include a Show list button for NMC entries
    nmcBox.innerHTML = `<div class="flex items-center justify-between"><div><strong>NMC PINs</strong> — ${soonCount} expiring within 30 days</div><div class="text-sm text-muted">${urgentCount} expiring within 7 days</div></div><div class="mt-2"><button id="nmc-show-list" class="px-2 py-1 ml-2 border rounded text-sm">Show list</button></div>`;
        // If urgent entries exist, render a small table of them
        if (urgentCount > 0) {
            const tbl = document.createElement('table'); tbl.className = 'w-full text-sm mt-2 border';
            tbl.innerHTML = '<thead><tr><th class="px-2 py-1 border-r">Name</th><th class="px-2 py-1 border-r">Staff No</th><th class="px-2 py-1 border-r">Valid To</th><th class="px-2 py-1">Days</th></tr></thead>';
            const tb = document.createElement('tbody');
            nmc.within7.slice(0,50).forEach(e => {
                const tr = document.createElement('tr');
                tr.innerHTML = `<td class="px-2 py-1 border-r">${escapeHtml(e.surname + (e.forenames ? ', ' + e.forenames : ''))}</td><td class="px-2 py-1 border-r">${escapeHtml(e.staffNumber)}</td><td class="px-2 py-1 border-r">${escapeHtml(e.validTo)}</td><td class="px-2 py-1">${escapeHtml(String(e.daysUntil))}</td>`;
                tb.appendChild(tr);
            });
            tbl.appendChild(tb);
            nmcBox.appendChild(tbl);
        }
        // create a unified alert bar containing RTW summary and NMC Pins
            try {
                const alertBar = document.createElement('div');
                alertBar.className = 'alerts-bar flex flex-col md:flex-row items-stretch gap-3 mb-3';
                // ensure children stretch equally: give both boxes flex-1
                try { if (typeof summary !== 'undefined' && summary) { summary.classList.add('flex-1'); alertBar.appendChild(summary); } } catch (e) {}
                try { if (nmcBox) { nmcBox.classList.add('flex-1'); nmcBox.innerHTML = nmcBox.innerHTML.replace('NMC PINs', 'NMC Pins'); alertBar.appendChild(nmcBox); } } catch (e) { /* ignore */ }
                try { resultsEl.insertBefore(alertBar, exportBar); } catch (e) { resultsEl.appendChild(alertBar); }

                // Re-wire handlers for the newly created alert bar content
                const rtwBtn = resultsEl.querySelector('#alerts-show-no-rtw');
                if (rtwBtn) {
                    rtwBtn.addEventListener('click', (e) => {
                        e.preventDefault();
                        const modalId = 'alerts-no-rtw-modal';
                        let m = document.getElementById(modalId);
                        if (!m) {
                            m = document.createElement('div'); m.id = modalId; m.className = 'modal-backdrop';
                            m.innerHTML = `<div class="modal"><div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Staff with shift after sickness and no RTW</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div><div class="modal-body p-2"><div id="alerts-no-rtw-list" class="space-y-2 text-sm"></div></div><div class="modal-footer mt-3 text-right"><button id="alerts-no-rtw-export" class="px-3 py-1 mr-2 rounded bg-slate-100 text-sm">Export CSV</button><button id="alerts-no-rtw-close" class="px-3 py-1 rounded">Close</button></div></div>`;
                            document.body.appendChild(m);
                            m.querySelector('.modal-close-btn').addEventListener('click', ()=> m.classList.remove('active'));
                            m.querySelector('#alerts-no-rtw-close').addEventListener('click', ()=> m.classList.remove('active'));
                        }
                        const list = m.querySelector('#alerts-no-rtw-list');
                        list.innerHTML = ''; // Clear previous content
                        if (!noRtw || !noRtw.length) { list.textContent = 'No staff found.'; }
                        else {
                            const tbl = document.createElement('table'); tbl.className = 'w-full text-sm border';
                            tbl.innerHTML = '<thead><tr><th class="px-2 py-1 border-r">Name</th><th class="px-2 py-1 border-r">Staff No</th><th class="px-2 py-1 border-r">Next Duty</th><th class="px-2 py-1">RTW Flags</th></tr></thead>';
                            const tb = document.createElement('tbody');
                            const rowsForExport = [];
                            noRtw.forEach(s => {
                                const tr = document.createElement('tr');
                                const name = (typeof findStaffName === 'function') ? (findStaffName(s.staff) || String(s.staff)) : String(s.staff);
                                const nextStr = s.nextDuty ? toLocalIso(s.nextDuty) : 'N/A';
                                tr.innerHTML = `<td class="px-2 py-1 border-r"><a href="#" class="staff-link" data-staff="${escapeHtml(String(s.staff))}">${escapeHtml(name)}</a></td><td class="px-2 py-1 border-r">${escapeHtml(String(s.staff))}</td><td class="px-2 py-1 border-r">${escapeHtml(nextStr)}</td><td class="px-2 py-1">${escapeHtml(String((s.rtwFlags||[]).join('; ')))}</td>`;
                                tb.appendChild(tr);
                                rowsForExport.push({ name: name, staff: String(s.staff), nextDuty: nextStr, rtwFlags: (s.rtwFlags||[]).join('; ') });
                            });
                            tbl.appendChild(tb); list.appendChild(tbl);
                            // delegate staff-link clicks to the modal container to avoid per-link listeners
                            try {
                                list.addEventListener('click', (ev) => {
                                    try {
                                        const a = ev.target && ev.target.closest ? ev.target.closest('.staff-link') : null;
                                        if (!a) return;
                                        ev.preventDefault();
                                        try { scheduleShowStaffDetails(a.dataset.staff); } catch (e) { console.error('scheduleShowStaffDetails failed', e); }
                                    } catch (e) { /* ignore per-click errors */ }
                                });
                            } catch (e) { /* ignore */ }
                            const exportBtn = m.querySelector('#alerts-no-rtw-export');
                            if (exportBtn) exportBtn.addEventListener('click', () => {
                                try {
                                    const csv = ['Name,Staff,NextDuty,RTWFlags', ...rowsForExport.map(r => `${csvSafe(r.name)},${csvSafe(r.staff)},${csvSafe(r.nextDuty)},${csvSafe(r.rtwFlags)}`)].join('\n');
                                    const b = new Blob([csv], { type: 'text/csv' }); const url = URL.createObjectURL(b); const a = document.createElement('a'); a.href = url; a.download = 'no_rtw_list.csv'; document.body.appendChild(a); try { a.click(); } catch (e) { window.open(url); } setTimeout(()=>{ try{ a.remove(); URL.revokeObjectURL(url); } catch(e){} },3000);
                                } catch (e) { alert('Export failed'); }
                            });
                        }
                        m.classList.add('active');
                    });
                }
                const nmcBtn = resultsEl.querySelector('#nmc-show-list');
                if (nmcBtn) {
                    nmcBtn.addEventListener('click', (e) => {
                        e.preventDefault();
                        const id = 'nmc-expiry-modal';
                        let im = document.getElementById(id);
                        if (!im) {
                            im = document.createElement('div'); im.id = id; im.className = 'modal-backdrop';
                            im.innerHTML = `<div class="modal"><div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">NMC Pins expiring</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div><div class="modal-body p-2"><div id="nmc-expiry-list" class="text-sm"></div></div><div class="modal-footer mt-3 text-right"><button id="nmc-expiry-export" class="px-3 py-1 mr-2 rounded bg-slate-100 text-sm">Export CSV</button><button id="nmc-expiry-close" class="px-3 py-1 py-1 rounded">Close</button></div></div>`;
                            document.body.appendChild(im);
                            im.querySelector('.modal-close-btn').addEventListener('click', ()=> im.classList.remove('active'));
                            im.querySelector('#nmc-expiry-close').addEventListener('click', ()=> im.classList.remove('active'));
                        }
                        const il = im.querySelector('#nmc-expiry-list'); il.innerHTML = '';
                        const entries = (nmc.within30 || []).slice(0,200);
                        if (!entries || entries.length === 0) il.textContent = 'No expiring NMC Pins found.';
                        else {
                            const tbl = document.createElement('table'); tbl.className = 'w-full text-sm border';
                            tbl.innerHTML = '<thead><tr><th class="px-2 py-1 border-r">Name</th><th class="px-2 py-1 border-r">Staff No</th><th class="px-2 py-1 border-r">Valid To</th><th class="px-2 py-1">Days</th></tr></thead>';
                            const tb = document.createElement('tbody');
                            const rowsForExport = [];
                            entries.forEach(e => {
                                const tr = document.createElement('tr'); tr.innerHTML = `<td class="px-2 py-1 border-r"><a href="#" class="staff-link" data-staff="${escapeHtml(String(e.staffNumber))}">${escapeHtml(e.surname + (e.forenames ? ', '+e.forenames : ''))}</a></td><td class="px-2 py-1 border-r">${escapeHtml(e.staffNumber)}</td><td class="px-2 py-1 border-r">${escapeHtml(e.validTo)}</td><td class="px-2 py-1">${escapeHtml(String(e.daysUntil))}</td>`;
                                tb.appendChild(tr);
                                rowsForExport.push({ name: e.surname + (e.forenames ? ', '+e.forenames : ''), staff: e.staffNumber, validTo: e.validTo, days: e.daysUntil });
                            });
                            tbl.appendChild(tb); il.appendChild(tbl);
                            try {
                                il.addEventListener('click', (ev) => {
                                    try {
                                        const a = ev.target && ev.target.closest ? ev.target.closest('.staff-link') : null;
                                        if (!a) return;
                                        ev.preventDefault();
                                        try { scheduleShowStaffDetails(a.dataset.staff); } catch (err) { console.error('scheduleShowStaffDetails failed', err); }
                                    } catch (e) { /* ignore */ }
                                });
                            } catch (e) { /* ignore */ }
                            const exportBtn = im.querySelector('#nmc-expiry-export');
                            if (exportBtn) exportBtn.addEventListener('click', () => {
                                try {
                                    const csv = ['Name,Staff,ValidTo,Days', ...rowsForExport.map(r => `${csvSafe(r.name)},${csvSafe(r.staff)},${csvSafe(r.validTo)},${csvSafe(String(r.days))}`)].join('\n');
                                    const b = new Blob([csv], { type: 'text/csv' }); const url = URL.createObjectURL(b); const a = document.createElement('a'); a.href = url; a.download = 'nmc_expiry.csv'; document.body.appendChild(a); try { a.click(); } catch (e) { window.open(url); } setTimeout(()=>{ try{ a.remove(); URL.revokeObjectURL(url); } catch(e){} },3000);
                                } catch (e) { alert('Export failed'); }
                            });
                        }
                        im.classList.add('active');
                    });
                }
            } catch (e) { if (window.__ALERTS_DEBUG) console.debug('renderAlerts: failed to insert alertBar', e); }
    } catch (e) { if (window.__ALERTS_DEBUG) console.debug('renderAlerts: NMC summary render failed', e); }

    exportCsvBtn.addEventListener('click', () => {
        try {
            if (window.__alertsIgnoredDuties && typeof window.__alertsIgnoredDuties.exportCsv === 'function') {
                window.__alertsIgnoredDuties.exportCsv();
            } else {
                alert('CSV export function not available.');
            }
        } catch (e) { console.warn('exportCsv failed', e); alert('Export failed'); }
    });

    // Helper: capture an element (table) as PNG via SVG foreignObject
    async function exportElementAsPng(el, filename = 'alerts_table.png', forcedWidth = null, forcedHeight = null) {
        try {
            let width = forcedWidth, height = forcedHeight;
            // Try multiple ways to measure size; fall back to scroll sizes when rect is zero
            if (width == null || height == null) {
                try {
                    const rect = (el.getBoundingClientRect && el.getBoundingClientRect()) || { width: 0, height: 0 };
                    width = Math.ceil(rect.width || 0);
                    height = Math.ceil(rect.height || 0);
                } catch (e) { width = 0; height = 0; }
                // fallback to scroll sizes for detached/cloned nodes or when bounding rect is 0
                if ((!width || !height) && el.scrollWidth && el.scrollHeight) {
                    width = Math.ceil(el.scrollWidth);
                    height = Math.ceil(el.scrollHeight);
                }
            }
            // If still zero, set conservative defaults to avoid canvas errors
            if (!width) width = 1000; if (!height) height = 600;
            const clone = el.cloneNode(true);
            // Inline wrapper for foreignObject; include basic font stack and background
            const svg = `<?xml version="1.0" encoding="utf-8"?>\n` +
                `<svg xmlns='http://www.w3.org/2000/svg' width='${width}' height='${height}'>` +
                `<foreignObject width='100%' height='100%'><div xmlns='http://www.w3.org/1999/xhtml' style='font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial; font-size:12px; color:#111; background:#fff;'>${clone.outerHTML}</div></foreignObject></svg>`;
            const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const img = new Image();
            // Ensure anonymous crossOrigin just in case (blob URLs typically don't need it)
            try { img.crossOrigin = 'anonymous'; } catch (e) {}
            let handled = false;
            img.onload = () => {
                try {
                    const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height;
                    const ctx = canvas.getContext('2d');
                    ctx.fillStyle = '#fff'; ctx.fillRect(0,0,width,height);
                    ctx.drawImage(img, 0, 0);
                    URL.revokeObjectURL(url);
                    canvas.toBlob(b => {
                        try {
                            const blobUrl = URL.createObjectURL(b);
                            const link = document.createElement('a');
                            link.href = blobUrl;
                            link.download = filename;
                            link.style.display = 'none';
                            document.body.appendChild(link);
                            try { link.click(); }
                            catch (err) {
                                console.warn('Programmatic download blocked, opening blob URL instead', err);
                                window.open(blobUrl);
                            }
                            handled = true;
                            setTimeout(() => { try { link.remove(); URL.revokeObjectURL(blobUrl); } catch (e) {} }, 3000);
                        } catch (e) { console.warn('export PNG blob handling failed', e); alert('PNG export failed'); }
                    }, 'image/png');
                } catch (e) {
                    console.warn('export PNG draw failed', e);
                    // fallback: open the generated SVG so user can save manually
                    try { handled = true; window.open(url); } catch (e2) { alert('PNG export failed'); }
                }
            };
            img.onerror = (e) => {
                console.warn('export PNG image load failed', e);
                // fallback: open the SVG blob URL so user can save the SVG
                try { window.open(url); } catch (e2) { alert('PNG export failed'); }
            };
            img.src = url;
            // safety net: if not handled after a short delay, open the SVG for manual save
            setTimeout(() => { try { if (!handled) window.open(url); } catch (e) {} }, 2500);
        } catch (e) { console.warn('exportElementAsPng failed', e); alert('Export failed'); }
    }

        // PNG export removed; SVG export remains available via the Export Alerts (SVG) button.

    // Export the element as an SVG file (downloads the constructed SVG directly)
    function exportElementAsSvg(el, filename = 'alerts_table.svg', forcedWidth = null, forcedHeight = null) {
        try {
            let width = forcedWidth, height = forcedHeight;
            try { const rect = (el.getBoundingClientRect && el.getBoundingClientRect()) || {}; width = width || Math.ceil(rect.width || el.scrollWidth || 1000); height = height || Math.ceil(rect.height || el.scrollHeight || 600); } catch (e) { width = width || 1000; height = height || 600; }
            const clone = el.cloneNode(true);
            const svg = `<?xml version="1.0" encoding="utf-8"?>\n` +
                `<svg xmlns='http://www.w3.org/2000/svg' width='${width}' height='${height}'>` +
                `<foreignObject width='100%' height='100%'><div xmlns='http://www.w3.org/1999/xhtml' style='font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial; font-size:12px; color:#111; background:#fff;'>${clone.outerHTML}</div></foreignObject></svg>`;
            const blob = new Blob([svg], { type: 'image/svg+xml;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a'); a.href = url; a.download = filename; a.style.display = 'none'; document.body.appendChild(a);
            try { a.click(); } catch (e) { window.open(url); }
            setTimeout(() => { try { a.remove(); URL.revokeObjectURL(url); } catch (e) {} }, 3000);
        } catch (e) { console.warn('exportElementAsSvg failed', e); alert('SVG export failed'); }
    }

    exportSvgBtn.addEventListener('click', () => {
        try {
            // Prefer the unified alert bar (contains RTW summary and NMC box)
            const alertBar = resultsEl.querySelector('.alerts-bar');
            const tableEl = resultsEl.querySelector('table');

            // If we have an alerts bar + table, export them together
            if (alertBar && tableEl) {
                const sRect = alertBar.getBoundingClientRect();
                const tRect = tableEl.getBoundingClientRect();
                const width = Math.ceil(Math.max(sRect.width, tRect.width));
                const height = Math.ceil(sRect.height + tRect.height);
                const wrapper = document.createElement('div');
                wrapper.style.background = '#fff';
                wrapper.appendChild(alertBar.cloneNode(true));
                wrapper.appendChild(tableEl.cloneNode(true));
                exportElementAsSvg(wrapper, 'alerts_with_summary.svg', width, height);
                return;
            }

            // Fallback: try to locate the summary and the NMC box individually
            const summaryEl = resultsEl.querySelector('.p-3.mb-3.rounded.border') || resultsEl.querySelector('.p-3.mb-3.rounded.border.bg-yellow-50');
            // If we don't have an alerts bar but we have summary + table, export them
            if (summaryEl && tableEl) {
                const sRect = summaryEl.getBoundingClientRect();
                const tRect = tableEl.getBoundingClientRect();
                const width = Math.ceil(Math.max(sRect.width, tRect.width));
                const height = Math.ceil(sRect.height + tRect.height);
                const wrapper = document.createElement('div');
                wrapper.style.background = '#fff';
                wrapper.appendChild(summaryEl.cloneNode(true));
                wrapper.appendChild(tableEl.cloneNode(true));
                exportElementAsSvg(wrapper, 'alerts_with_summary.svg', width, height);
                return;
            }

            // If only the table exists, export table (unchanged)
            if (tableEl) {
                exportElementAsSvg(tableEl, 'alerts_table.svg');
                return;
            }

            // If only the summary exists, export summary (unchanged)
            if (summaryEl) {
                exportElementAsSvg(summaryEl, 'alerts_summary.svg');
                return;
            }

            alert('Nothing to export');
        } catch (e) {
            console.warn('SVG export failed', e);
            alert('SVG export failed');
        }
    });

    const table = document.createElement('table'); table.className = 'w-full text-sm';
    // Use auto layout so columns size to content; enable horizontal scroll on the container
    table.style.tableLayout = 'auto';
    table.style.width = 'auto';
    try { resultsEl.style.overflowX = 'auto'; resultsEl.style.maxWidth = '100%'; } catch (e) {}
        // apply any saved widths
        let colWidths = {};
        try { const raw = localStorage.getItem('TERRA_ALERTS_COL_WIDTHS'); if (raw) colWidths = JSON.parse(raw) || {}; } catch (e) { colWidths = {}; }
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        cols.forEach(c => {
            const th = document.createElement('th'); th.className = 'px-2 py-1 text-left border-b';
            th.style.boxSizing = 'border-box';
            th.style.whiteSpace = 'nowrap';
            if (colWidths[c]) th.style.width = colWidths[c];
            th.textContent = c;
            // show full header on hover
            th.title = String(c);
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        const tbody = document.createElement('tbody');
    // Helper: display mode - global vs per-file
    const displayMode = localStorage.getItem('TERRA_ALERTS_DISPLAY_BY') || 'global';
    // Helper: format a value according to the user's selected display preference
    const dateDisplayPref = localStorage.getItem('TERRA_DATE_DISPLAY_V2') || localStorage.getItem('TERRA_DATE_DISPLAY') || 'uk';
        if (!!window.__ALERTS_DEBUG) console.debug('renderAlerts: dateDisplayPref', dateDisplayPref, 'TERRA_DATE_DISPLAY', localStorage.getItem('TERRA_DATE_DISPLAY'), 'TERRA_DATE_DISPLAY_V2', localStorage.getItem('TERRA_DATE_DISPLAY_V2'));
        const formatForDisplay = (val, rowLocale) => {
            if (val == null) return '';
            if (dateDisplayPref === 'iso') { if (!!window.__ALERTS_DEBUG) console.debug('formatForDisplay: iso branch', val); return String(val); }
            // For 'uk' and 'us' prefer the formatByLocale helper but pass the requested display locale
            if (dateDisplayPref === 'uk') { if (!!window.__ALERTS_DEBUG) console.debug('formatForDisplay: uk branch', val); return formatByLocale(val, 'uk'); }
            if (dateDisplayPref === 'us') { if (!!window.__ALERTS_DEBUG) console.debug('formatForDisplay: us branch', val); return formatByLocale(val, 'us'); }
            if (!!window.__ALERTS_DEBUG) console.debug('formatForDisplay: fallback branch', val);
            return String(val);
        };
        res.rows.slice(0,200).forEach((r, ridx) => {
            const tr = document.createElement('tr');
            // apply zebra banding for readability
            if (ridx % 2 === 0) {
                tr.style.background = 'transparent';
            } else {
                tr.style.background = 'rgba(0,0,0,0.02)';
            }
            // build base cells from visibleHeaders (includes 'Source File' if selected)
            let rowHtml = visibleHeaders.map(h => {
                // render Source File specially
                if (h === 'Source File') {
                    const sf = escapeHtml(r.__fileName);
                    const w = colWidths[h] ? ` style="width:${escapeHtml(colWidths[h])};box-sizing:border-box"` : ' style="box-sizing:border-box"';
                    return `<td class="px-2 py-1 border-b" title="${sf}"${w}>${sf}</td>`;
                }
                const rawVal = getValueByCanonical(r.data, h);
                const val = formatForDisplay(rawVal, r.__locale);
                // show truncated inline value with tooltip; allows table to remain compact
                const displayText = escapeHtml(val);
                const truncated = `<span class="truncate-inline" title="${displayText}">${displayText}</span>`;
                const w = colWidths[h] ? ` style="width:${escapeHtml(colWidths[h])};box-sizing:border-box;white-space:nowrap"` : ' style="box-sizing:border-box;white-space:nowrap"';
                return `<td class="px-2 py-1 border-b" title="${escapeHtml(String(val))}"${w}>${truncated}</td>`;
            }).join('');
            // determine staff key for this row (prefer precomputed __staffKey added during buildResults)
            let staffVal = (r.__staffKey && String(r.__staffKey).trim()) || getStaffKey(r.data) || getValueByCanonical(r.data, 'Assignment No');
            let lastIso = null; let nextIso = null;
            if (staffVal) {
                try { staffVal = String(staffVal).trim(); } catch (e) { staffVal = String(staffVal || ''); }
                if (res.lastDutyMap) lastIso = res.lastDutyMap[staffVal] || null;
                if (res.nextDutyMap) nextIso = res.nextDutyMap[staffVal] || null;
            }
            const lastDisplayBase = lastIso ? String(formatForDisplay(lastIso, r.__locale)) : '';
            const nextDisplayBase = nextIso ? String(formatForDisplay(nextIso, r.__locale)) : '';

            // Add a compact RTW badge column (green if RTW done, red if missing and had shift after sickness)
            const staffKey = (staffVal ? String(staffVal) : '').trim();
            let badgeHtml = '';
            try {
                // Resolve staffObj from perStaffMap. Try exact match first, then fall back to
                // numeric-only matching across keys (handles '27029932-2', spacing, leading zeros, etc.).
                let staffObj = null;
                if (res.perStaffMap && staffKey) staffObj = res.perStaffMap[staffKey] || null;
                if (!staffObj && res.perStaffMap && staffKey) {
                    const norm = String(staffKey).replace(/\D/g, '');
                    if (norm) {
                        for (const k of Object.keys(res.perStaffMap || {})) {
                            try {
                                if (String(k).replace(/\D/g, '') === norm) { staffObj = res.perStaffMap[k]; break; }
                            } catch (e) { /* ignore per-key errors */ }
                        }
                    }
                }
                if (staffObj) {
                    if (staffObj.rtwDone) badgeHtml = `<span class="badge green" title="RTW recorded">RTW</span>`;
                    else if (staffObj.hadShiftAfter) badgeHtml = `<span class="badge red" title="Had shift after sickness, no RTW">No RTW</span>`;
                    else badgeHtml = `<span class="badge gray" title="No recent duty">OK</span>`;
                    // highlight continuing sickness if detected (merged ranges or ongoing end date)
                }
            } catch (e) { badgeHtml = ''; }
            // Enrich the Next Duty display with the shift detail when available (uses nextDutyDetailMap produced by buildResults)
            let nextDisplayWithShift = nextDisplayBase ? escapeHtml(nextDisplayBase) : '';
            try {
                if (nextDisplayBase && res.nextDutyDetailMap) {
                    const detail = res.nextDutyDetailMap[staffKey] || null;
                    if (detail && detail.shift) {
                        const shiftText = String(detail.shift || '').trim();
                        if (shiftText) nextDisplayWithShift = `${escapeHtml(nextDisplayBase)} — ${escapeHtml(shiftText)}`;
                    }
                }
            } catch (e) { /* ignore enrichment errors and fall back to base display */ }

            const wLast = colWidths['Last Duty'] ? ` style="width:${escapeHtml(colWidths['Last Duty'])};box-sizing:border-box;white-space:nowrap"` : ' style="box-sizing:border-box;white-space:nowrap"';
            rowHtml += `<td class="px-2 py-1 border-b" title="${escapeHtml(String(lastDisplayBase))}"${wLast}>${escapeHtml(String(lastDisplayBase))}</td>`;
            const wNext = colWidths['Next Duty'] ? ` style="width:${escapeHtml(colWidths['Next Duty'])};box-sizing:border-box;white-space:nowrap"` : ' style="box-sizing:border-box;white-space:nowrap"';
            rowHtml += `<td class="px-2 py-1 border-b" title="${escapeHtml(String(nextDisplayWithShift))}"${wNext}>${nextDisplayWithShift}</td>`;

            rowHtml += `<td class="px-2 py-1 border-b" title="RTW status">${badgeHtml}</td>`;

            // Per-row actions: Details / Export row
            const actions = `<td class="px-2 py-1 border-b text-right"><button class="row-action" data-row="${ridx}" data-staff="${escapeHtml(staffKey)}" data-file="${escapeHtml(r.__fileName||'')}">Details</button></td>`;
            rowHtml += actions;
            tr.innerHTML = rowHtml;
            tbody.appendChild(tr);
            // per-row listeners removed in favor of a delegated handler attached to the table
        });
        table.appendChild(thead); table.appendChild(tbody);
        resultsEl.appendChild(table);

        // Delegated handler for per-row actions (Details). This is more reliable than attaching
        // individual listeners to each row element and survives reflows.
        try {
            // resilient fallback to show staff details even when the scoped showStaffDetails is not exposed
            try {
                if (typeof window !== 'undefined' && !window.showStaffDetailsSafe) {
                    window.showStaffDetailsSafe = function (staffId) {
                        try {
                            console.debug('showStaffDetailsSafe called', { staffId });
                            const entries = [];
                            const wanted = String(staffId || '').trim();
                            const wantedDigits = wanted.replace(/\D/g, '');

                            // Fast path: use prebuilt index when available (maps numeric staff -> hits)
                            try {
                                const idx = (typeof window !== 'undefined' && window.__staffIndex) ? window.__staffIndex : null;
                                if (idx && wantedDigits) {
                                    const hits = idx[wantedDigits] || [];
                                    if (hits && hits.length) {
                                        hits.forEach(h => { try { entries.push({ file: h.file, row: h.row, rowIndex: h.rowIndex }); } catch (e) {} });
                                    }
                                }
                            } catch (e) { /* ignore index errors */ }

                            // If index miss, fall back to scanning files (previous logic)
                            if (!entries.length) {
                                const files = (typeof getParsedFiles === 'function') ? getParsedFiles() : (window.__getParsedFiles ? window.__getParsedFiles() : []);
                                // Primary pass: prefer explicit staff extraction (precomputed or via getStaffKey)
                                (files || []).forEach(f => {
                                    (f.dataRows || []).forEach((r, idx) => {
                                        try {
                                            const rawKey = String((r && r.__staffKey) || getStaffKey(r) || '').trim();
                                            const rawDigits = rawKey.replace(/\D/g, '');
                                            if (rawKey && (rawKey === wanted || (rawDigits && wantedDigits && rawDigits === wantedDigits))) {
                                                entries.push({ file: f, row: r, rowIndex: idx });
                                                return;
                                            }
                                            const alt = String(getStaffKey(r) || '').trim();
                                            const altDigits = alt.replace(/\D/g, '');
                                            if (alt && (alt === wanted || (altDigits && wantedDigits && altDigits === wantedDigits))) {
                                                entries.push({ file: f, row: r, rowIndex: idx });
                                                return;
                                            }
                                        } catch (e) { /* ignore per-row errors */ }
                                    });
                                });
                            }

                            // If none found, do a relaxed scan across all fields (some sources put staff id in odd columns)
                            if (!entries.length) {
                                const sampleKeys = [];
                                (files || []).forEach(f => {
                                    (f.dataRows || []).forEach((r, idx) => {
                                        try {
                                            if (sampleKeys.length < 30) {
                                                const sample = String(getStaffKey(r) || (r && (r['Staff Number'] || r['Assignment No'] || r['Staff'] || '')) || '');
                                                sampleKeys.push(sample);
                                            }
                                            // Scan every cell value for the wanted digits as a substring match
                                            const found = Object.keys(r || {}).some(k => {
                                                try {
                                                    const v = r[k];
                                                    if (v == null) return false;
                                                    const s = typeof v === 'string' ? v : (typeof v === 'number' ? String(v) : JSON.stringify(v));
                                                    if (!s) return false;
                                                    if (s.includes(wanted)) return true;
                                                    const digits = String(s).replace(/\D/g, '');
                                                    if (digits && wantedDigits && digits === wantedDigits) return true;
                                                    // also accept when digits contain the wantedDigits (e.g. cell 'ID:27029932-2')
                                                    if (digits && wantedDigits && digits.indexOf(wantedDigits) !== -1) return true;
                                                    return false;
                                                } catch (e) { return false; }
                                            });
                                            if (found) entries.push({ file: f, row: r, rowIndex: idx });
                                        } catch (e) { /* ignore */ }
                                    });
                                });
                                if (!entries.length) {
                                    try { console.debug('showStaffDetailsSafe: no direct matches; diagnostic sampleKeys', { wanted, wantedDigits, sampleKeys: sampleKeys.slice(0,20) }); } catch (e) {}
                                }
                            }

                            // create modal similar to showStaffDetails implementation
                            const modalId = 'staff-detail-modal';
                            let md = document.getElementById(modalId);
                            if (!md) {
                                md = document.createElement('div'); md.id = modalId; md.className = 'modal-backdrop';
                                md.innerHTML = `<div class="modal" style="max-width:900px"><div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Staff details</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div><div class="modal-body p-2"><div id="staff-detail-body" class="text-sm"></div></div><div class="modal-footer mt-3 text-right"><button id="staff-detail-close" class="px-3 py-1 rounded">Close</button></div></div>`;
                                document.body.appendChild(md);
                                try { md.querySelector('.modal-close-btn').addEventListener('click', () => md.classList.remove('active')); } catch (e) {}
                                try { md.querySelector('#staff-detail-close').addEventListener('click', () => md.classList.remove('active')); } catch (e) {}
                            }
                            const body = md.querySelector('#staff-detail-body'); body.innerHTML = '';
                            try {
                                const countEl = document.createElement('div'); countEl.className = 'text-sm text-subtle mb-2'; countEl.textContent = `Matches: ${entries.length}`;
                                body.appendChild(countEl);
                                renderStaffDetailsEntries(body, entries);
                            } catch (e) { body.textContent = 'No rows found for this staff in the library.'; }
                            md.classList.add('active');
                            console.debug('showStaffDetailsSafe returning', { staffId, matches: entries.length });
                            return true;
                        } catch (ex) { console.error('showStaffDetailsSafe failed', ex); return false; }
                    };
                }
            } catch (e) { /* ignore */ }

            // helper: open a lightweight modal immediately and defer heavy scanning to idle time
            function scheduleShowStaffDetails(staff) {
                try {
                    const modalId = 'staff-detail-modal';
                    let md = document.getElementById(modalId);
                    if (!md) {
                        md = document.createElement('div'); md.id = modalId; md.className = 'modal-backdrop';
                        md.innerHTML = `<div class="modal" style="max-width:900px"><div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Staff details</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div><div class="modal-body p-2"><div id="staff-detail-body" class="text-sm">Loading…</div></div><div class="modal-footer mt-3 text-right"><button id="staff-detail-close" class="px-3 py-1 rounded">Close</button></div></div>`;
                        document.body.appendChild(md);
                        try { md.querySelector('.modal-close-btn').addEventListener('click', () => md.classList.remove('active')); } catch (e) {}
                        try { md.querySelector('#staff-detail-close').addEventListener('click', () => md.classList.remove('active')); } catch (e) {}
                    } else {
                        const body = md.querySelector('#staff-detail-body'); if (body) body.textContent = 'Loading…';
                    }
                    md.classList.add('active');
                    // defer the heavy search to idle time to keep click handler responsive
                    const invoke = () => { try { if (window && typeof window.showStaffDetailsSafe === 'function') window.showStaffDetailsSafe(staff); else if (typeof showStaffDetails === 'function') showStaffDetails(staff); } catch (e) { console.error('deferred showStaffDetails failed', e); } };
                    if (typeof requestIdleCallback === 'function') requestIdleCallback(invoke, { timeout: 500 });
                    else setTimeout(invoke, 20);
                } catch (e) { console.error('scheduleShowStaffDetails failed', e); }
            }

            // handler extracted so it can be used both by delegated listener and direct per-button bindings
            function handleRowAction(btn) {
                try {
                    // Prefer explicit data-staff attribute if available to avoid index mismatch
                    const staffFromAttr = btn.dataset && btn.dataset.staff ? String(btn.dataset.staff).trim() : null;
                    let staff = staffFromAttr;
                    if (!staff) {
                        const idx = Number(btn.dataset.row);
                        if (Number.isNaN(idx)) return;
                        const row = (res && Array.isArray(res.rows)) ? res.rows[idx] : null;
                        if (!row) return;
                        // Prefer a precomputed key then centralized helper for robust extraction
                        staff = String((row.__staffKey && String(row.__staffKey).trim()) || getStaffKey(row.data) || '').trim();
                    }
                    console.debug('alerts table action clicked', { staff });
                    if (staff) {
                        try {
                            console.debug('scheduling showStaffDetails', { staff });
                            scheduleShowStaffDetails(staff);
                        } catch (e) { console.error('failed to schedule staff details', e); }
                    }
                } catch (e) { console.error('alerts table action failed', e); }
            }

            table.addEventListener('click', (ev) => {
                try {
                    const btn = ev.target && ev.target.closest ? ev.target.closest('.row-action') : null;
                    if (!btn) return;
                    ev.preventDefault();
                    handleRowAction(btn);
                } catch (e) { /* ignore click handling errors */ }
            });

            // Per-button bindings removed: delegation handles clicks for dynamic rows and avoids double-invocation.
            // (If someone expects per-button bindings, they can be reintroduced, but delegation is preferred.)
        } catch (e) { /* ignore delegation setup errors */ }
        // Make columns resizable by dragging the right edge of headers
        (function makeColumnsResizable(tbl, colsArr) {
            if (!tbl) return;
            const thead = tbl.querySelector('thead');
            if (!thead) return;
            const ths = Array.from(thead.querySelectorAll('th'));
            // remove any existing resizers to avoid duplicates
            ths.forEach(th => { const prev = th.querySelector('.col-resizer'); if (prev) prev.remove(); });
            let dragging = null;
            let startX = 0; let startW = 0; let colIndex = -1;
            const onMouseMove = (e) => {
                if (!dragging) return;
                const clientX = (e.touches && e.touches[0]) ? e.touches[0].clientX : e.clientX;
                const dx = clientX - startX;
                const newW = Math.max(30, startW + dx);
                try { dragging.th.style.width = newW + 'px'; dragging.th.style.boxSizing = 'border-box'; } catch (e) {}
                // apply width to all body cells in that column (same index)
                try {
                    tbl.querySelectorAll('tbody tr').forEach(tr => {
                        const td = tr.children[colIndex]; if (td) { td.style.width = newW + 'px'; td.style.boxSizing = 'border-box'; }
                    });
                } catch (e) {}
            };
            const onMouseUp = (e) => {
                if (!dragging) return;
                // persist width
                try {
                    const headerName = colsArr[colIndex] || dragging.th.textContent || `col_${colIndex}`;
                    const raw = localStorage.getItem('TERRA_ALERTS_COL_WIDTHS');
                    const map = raw ? JSON.parse(raw || '{}') : {};
                    const finalW = dragging.th.style.width || `${dragging.th.getBoundingClientRect().width}px`;
                    map[headerName] = finalW;
                    localStorage.setItem('TERRA_ALERTS_COL_WIDTHS', JSON.stringify(map));
                } catch (e) { /* ignore */ }
                dragging = null; colIndex = -1; startX = 0; startW = 0;
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
            };
            ths.forEach((th, idx) => {
                // ensure th is positioned so grip sits inside
                if (!th.style.position) th.style.position = 'relative';
                const grip = document.createElement('div');
                grip.className = 'col-resizer';
                grip.style.position = 'absolute';
                grip.style.top = '0';
                grip.style.right = '0';
                grip.style.width = '12px';
                grip.style.cursor = 'col-resize';
                grip.style.userSelect = 'none';
                grip.style.height = '100%';
                grip.style.zIndex = '20';
                grip.style.background = 'transparent';
                // subtle hover feedback
                grip.addEventListener('mouseenter', () => { grip.style.background = 'rgba(0,0,0,0.03)'; });
                grip.addEventListener('mouseleave', () => { if (!dragging) grip.style.background = 'transparent'; });
                th.appendChild(grip);
                grip.addEventListener('pointerdown', (ev) => {
                    try { ev.preventDefault(); ev.stopPropagation(); } catch (e) {}
                    dragging = { th };
                    colIndex = idx;
                    startX = ev.clientX;
                    startW = th.getBoundingClientRect().width;
                    // capture the pointer so move/up events are delivered to the grip
                    try { if (grip.setPointerCapture) grip.setPointerCapture(ev.pointerId); } catch (e) {}
                    document.addEventListener('pointermove', onMouseMove);
                    document.addEventListener('pointerup', onMouseUp);
                });
            });
        })(table, cols);
        // If debug mode, render diagnostics below the table to show parsedIso/inWindow for each evaluated date cell
        if (debug) {
            const diagEl = document.createElement('div');
            diagEl.className = 'mt-3 p-2 border rounded text-xs bg-surface-alt overflow-auto';
            diagEl.style.maxHeight = '200px';
            diagEl.innerHTML = `<div class="font-semibold mb-2">Diagnostics (parsedIso / inWindow)</div>`;
            const dl = document.createElement('div');
            if (res.diagnostics && res.diagnostics.length) {
                res.diagnostics.slice(0,200).forEach(d => {
                    const row = document.createElement('div');
                    row.style.padding = '2px 0';
                    row.textContent = `${d.fileName} [row:${d.rowIndex}] ${d.column} raw="${String(d.raw)}" parsedIso=${d.parsedIso} inWindow=${d.inWindow} op=${d.rule.operator} val=${d.rule.value}`;
                    dl.appendChild(row);
                });
            } else {
                dl.textContent = 'No diagnostics available';
            }
            diagEl.appendChild(dl);
            resultsEl.appendChild(diagEl);
        }
        const countEl = tabContent.querySelector('#alerts-count'); if (countEl) countEl.textContent = `${res.rows.length} matched rows`;
        // wire display-mode selector
        const dispSel = tabContent.querySelector('#alerts-display-mode');
        if (dispSel) {
            dispSel.value = localStorage.getItem('TERRA_ALERTS_DISPLAY_BY') || 'global';
            dispSel.addEventListener('change', (e) => {
                localStorage.setItem('TERRA_ALERTS_DISPLAY_BY', e.target.value);
                try { updateAlerts(); } catch (err) { /* ignore */ }
            });
        }
        // wire Choose Columns button
        const chooseBtn = tabContent.querySelector('#choose-columns-btn');
        if (chooseBtn) {
            chooseBtn.addEventListener('click', (e) => {
                e.preventDefault(); showColumnsModal(res.headers);
            });
        }
    }

    // Listen for date display preference changes so Alerts can refresh immediately
    if (!document.__terraDateDisplayListenerRegistered) {
        document.addEventListener('terra:dateDisplayChanged', () => {
            try { updateAlerts(); } catch (e) { /* ignore */ }
        });
        document.__terraDateDisplayListenerRegistered = true;
    }

    function buildUI() {
        const files = getParsedFiles();
        const headers = unionHeaders(files);
        tabContent.innerHTML = `
            <div class="grid lg:grid-cols-3 gap-6">
                <div class="col-span-1 p-4 border rounded-lg bg-surface-alt">
                    <h3 class="font-semibold mb-2">Files</h3>
                    <div id="alerts-files" class="space-y-2 text-sm max-h-[60vh] overflow-y-auto"></div>
                </div>
                <div class="col-span-2">
                    <div class="p-4 border rounded-lg bg-surface-alt mb-4">
                        <div class="flex items-center justify-between"><h3 class="font-semibold">Rules</h3><div id="alerts-count" class="text-sm text-subtle">0 matched rows</div></div>
                        <div class="mt-2 flex items-center gap-2"><button id="choose-columns-btn" class="px-2 py-1 rounded border text-sm">Choose columns</button><button id="alerts-add-rule" class="px-3 py-1 rounded bg-blue-600 text-white">+ Add Rule</button>
                        </div>
                        <div id="alerts-rules" class="mt-3"></div>
                    </div>
                    <div id="alerts-results" class="p-2 border rounded-lg bg-white max-h-[60vh] overflow-auto"></div>
                </div>
            </div>
        `;

        const filesEl = tabContent.querySelector('#alerts-files');
        // add a drag-and-drop area at the top of files list
        const dropZone = document.createElement('div');
        dropZone.className = 'p-2 mb-2 border-dashed border-2 rounded text-sm text-muted text-center';
        dropZone.textContent = 'Drag CSV or XLSX files here to add them to the library for alerting';
        filesEl.parentNode.insertBefore(dropZone, filesEl);
        // hidden file input for fallback browse
        const ddInput = document.createElement('input'); ddInput.type = 'file'; ddInput.accept = '.csv, .xlsx'; ddInput.multiple = true; ddInput.className = 'hidden';
        filesEl.parentNode.insertBefore(ddInput, filesEl);
        dropZone.style.cursor = 'pointer';
        dropZone.addEventListener('click', () => ddInput.click());
        dropZone.addEventListener('dragenter', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
        dropZone.addEventListener('dragover', (e) => { e.preventDefault(); try { if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy'; } catch (_){} dropZone.classList.add('drag-over'); });
        dropZone.addEventListener('dragleave', (e) => { dropZone.classList.remove('drag-over'); });
        dropZone.addEventListener('drop', async (e) => {
            e.preventDefault(); dropZone.classList.remove('drag-over');
            const files = e.dataTransfer ? e.dataTransfer.files : null;
            console.debug('dropZone: files dropped', files);
            if (!files || files.length === 0) { toast('No files detected in drop'); return; }
            await handleDroppedFiles(files);
        });
        ddInput.addEventListener('change', async (e) => { const files = e.target.files; if (files && files.length) await handleDroppedFiles(files); ddInput.value = ''; });

        async function handleDroppedFiles(files) {
            const list = Array.from(files);
            toast(`Adding ${list.length} file(s) to library...`);
            for (const f of list) {
                try {
                    const lower = (f.name || '').toLowerCase();
                    if (!lower.endsWith('.csv') && !lower.endsWith('.xlsx')) { console.debug('skipping unsupported file', f.name); continue; }
                    console.debug('handleDroppedFiles: parsing', f.name);
                    const parsed = await parseFile(f);
                    const item = { name: f.name, isParsedFile: true, dataRows: parsed, headers: parsed.length ? Object.keys(parsed[0]) : [], _locale: 'uk' };
                    const id = await dbSave(STORES.FILES, item);
                    const saved = { ...item, id };
                    addLibraryItem(saved);
                    try { console.debug('handleDroppedFiles: added to library', saved.name, 'id', saved.id, 'library now', (typeof getLibraryItems === 'function' ? getLibraryItems().map(x=>({name:x.name,id:x.id,rows:(x.dataRows||[]).length})) : 'getLibraryItems unavailable')); } catch (e) { console.debug('handleDroppedFiles: debug getLibraryItems failed', e); }
                    rehydrateItem(saved);
                    // select the newly added file for alerts
                    selectedFileIds.add(id);
                    saveAlertsSelectedFiles();
                    toast(`Added ${f.name}`);
                } catch (err) {
                    console.error('Failed to handle dropped file', f.name, err);
                    toast(`Failed to add ${f.name}`);
                }
            }
            // rebuild staff index after files are added
            try { if (typeof buildStaffIndex === 'function') { buildStaffIndex(); } } catch (e) { console.warn('rebuild staff index after drop failed', e); }
            // rebuild UI to show new files and update alerts
            buildUI();
            updateAlerts();
            toast('Files added to library');
        }
        files.forEach(f => {
            const div = document.createElement('div');
            div.className = 'flex items-center gap-2';
            const cb = document.createElement('input'); cb.type = 'checkbox'; cb.dataset.id = f.id; cb.checked = selectedFileIds.has(f.id);
            const lbl = document.createElement('label'); lbl.textContent = `${f.name} (${(f.dataRows||[]).length} rows)`; lbl.className = 'text-sm flex-1';
            // locale badge (click to toggle UK/US) to help debug ambiguous slashed dates
            const localeBtn = document.createElement('button'); localeBtn.className = 'locale-badge text-xs px-2 py-1 border rounded';
            const currentLocale = (f._locale || 'uk').toString().toLowerCase(); localeBtn.textContent = currentLocale.toUpperCase();
            localeBtn.title = 'Click to toggle locale preference for this file';
            // delete button to remove file from library and DB
            const delBtn = document.createElement('button');
            delBtn.className = 'px-2 py-1 rounded border text-sm text-red-600';
            delBtn.title = `Delete ${f.name}`;
            delBtn.textContent = 'Delete';
            // mark actions so we can use a single delegated handler instead of per-element listeners
            localeBtn.dataset.action = 'toggle-locale';
            delBtn.dataset.action = 'delete-file';
            div.appendChild(cb); div.appendChild(lbl); div.appendChild(localeBtn); div.appendChild(delBtn); filesEl.appendChild(div);
            // per-item listeners removed in favor of a delegated handler bound once below
        });

        // Delegated handlers for files list (one-time bindings rather than per-file)
        try {
            filesEl.addEventListener('change', (e) => {
                try {
                    const cb = e.target && e.target.closest ? e.target.closest('input[type=checkbox]') : null;
                    if (!cb) return;
                    const id = cb.dataset && cb.dataset.id ? cb.dataset.id : null;
                    if (!id) return;
                    if (cb.checked) selectedFileIds.add(id); else selectedFileIds.delete(id);
                    saveAlertsSelectedFiles(); updateAlerts();
                } catch (err) { /* ignore per-change errors */ }
            });
            filesEl.addEventListener('click', async (e) => {
                try {
                    const btn = e.target && e.target.closest ? e.target.closest('button') : null;
                    if (!btn) return;
                    const action = btn.dataset && btn.dataset.action ? btn.dataset.action : null;
                    // find the nearest checkbox label to extract file id
                    const wrapper = btn.closest && btn.closest('div') ? btn.closest('div') : null;
                    let fileId = null;
                    if (wrapper) {
                        const cb = wrapper.querySelector && wrapper.querySelector('input[type=checkbox]') ? wrapper.querySelector('input[type=checkbox]') : null;
                        if (cb && cb.dataset) fileId = cb.dataset.id;
                    }
                    if (action === 'toggle-locale' && fileId) {
                        e.preventDefault();
                        const lib = getLibraryItems();
                        const item = lib.find(x => x && x.id === fileId);
                        if (!item) return alert('File not found in library');
                        const newLocale = ((item._locale || 'uk').toString().toLowerCase() === 'uk') ? 'us' : 'uk';
                        try {
                            item._locale = newLocale;
                            try { if (Array.isArray(item.dataRows)) item.dataRows = normalizeFileDates(item.dataRows, newLocale); } catch (errNorm) { console.warn('normalize on locale toggle failed', errNorm); }
                            await dbSave(STORES.FILES, { ...item, _locale: newLocale, dataRows: item.dataRows });
                            // update UI
                            try { window.__ALERTS_FORCE_FILE_DISPLAY = true; } catch (e) {}
                            try { buildUI(); } catch (e) {}
                            try { updateAlerts(); } catch (e) {}
                            try { document.dispatchEvent(new CustomEvent('terra:dateDisplayChanged', { detail: { value: newLocale } })); } catch (e) {}
                            setTimeout(() => { try { window.__ALERTS_FORCE_FILE_DISPLAY = false; buildUI(); updateAlerts(); document.dispatchEvent(new CustomEvent('terra:dateDisplayChanged', { detail: { value: localStorage.getItem('TERRA_DATE_DISPLAY_V2') || localStorage.getItem('TERRA_DATE_DISPLAY') || 'uk' } })); } catch (e) {} }, 3000);
                            toast(`File locale for ${item.name} set to ${newLocale}`);
                        } catch (err) { console.error('Failed to persist file locale', err); alert('Failed to update file locale'); }
                    } else if (action === 'delete-file' && fileId) {
                        e.preventDefault();
                        const ok = confirm(`Delete "${fileId}" from the library? This cannot be undone.`);
                        if (!ok) return;
                        try {
                            await dbDelete(STORES.FILES, fileId);
                            try { removeLibraryItemById(fileId); } catch (inner) { console.warn('removeLibraryItemById failed', inner); }
                            try { selectedFileIds.delete(fileId); saveAlertsSelectedFiles(); } catch (inner) {}
                            try { buildUI(); } catch (inner) {}
                            try { updateAlerts(); } catch (inner) {}
                            try { if (typeof buildStaffIndex === 'function') buildStaffIndex(); } catch (e) { console.warn('rebuild staff index after delete failed', e); }
                            toast(`Deleted file`);
                        } catch (err) { console.error('Failed to delete file', err); alert('Failed to delete file from library'); }
                    }
                } catch (e) { /* ignore delegated click errors */ }
            });
        } catch (e) { /* ignore delegation setup errors */ }

        const addBtn = tabContent.querySelector('#alerts-add-rule');
        addBtn.addEventListener('click', async () => {
            // create a new rule defaulting to first header
            const filesNow = getParsedFiles().filter(fi => selectedFileIds.has(fi.id));
            const hdrs = unionHeaders(filesNow);
            const rule = { id: `r_${Date.now()}`, column: hdrs[0] || '', operator: 'contains', value: '' };
            rules.push(rule);
            await saveAllRules();
            updateAlerts();
        });

        // Note: bulk locale and manage-column-links controls have been removed for a cleaner UI.

        // default: select all parsed files initially if there was no saved selection
        if (files.length && selectedFileIds.size === 0) files.forEach(f => selectedFileIds.add(f.id));
        // persist any initial selection
        saveAlertsSelectedFiles();
        updateAlerts();
    }

        buildUI();
    // Defer heavy work (alerts build and staff index) to idle so initial UI is snappy
    try {
        const doDeferred = () => {
            try { updateAlerts(); } catch (e) { console.warn('deferred updateAlerts failed', e); }
            try { if (typeof buildStaffIndex === 'function') buildStaffIndex(); } catch (e) { console.warn('deferred buildStaffIndex failed', e); }
        };
        if (typeof requestIdleCallback === 'function') requestIdleCallback(doDeferred, { timeout: 300 }); else setTimeout(doDeferred, 50);
    } catch (e) { console.warn('scheduling deferred alerts work failed', e); }
}

// Expose debug helpers for the browser console
    try {
    if (typeof window !== 'undefined') {
        window.buildResultsDebug = function () {
            try {
                if (typeof buildResults === 'function') return buildResults();
            } catch (e) { if (window.__ALERTS_DEBUG) console.debug('buildResults call failed in buildResultsDebug', e); }
            // synthesize a result using library scan and diagnostics
            try {
                const perArr = (typeof computeRTWLibraryStats === 'function') ? computeRTWLibraryStats({ includeSources: true }) : [];
                const perMap = {};
                (perArr || []).forEach(p => { try { if (p && p.staff) perMap[String(p.staff)] = p; } catch (e) {} });
                const diag = (typeof runAlertsDiagnostics === 'function') ? runAlertsDiagnostics() : { nextDutyMap: {}, lastDutyMap: {}, nextDutyDetailMap: {}, lastDutyDetailMap: {}, ignoredDuties: [] };
                return { headers: [], rows: [], diagnostics: diag, nextDutyMap: diag.nextDutyMap || {}, lastDutyMap: diag.lastDutyMap || {}, nextDutyDetailMap: diag.nextDutyDetailMap || {}, lastDutyDetailMap: diag.lastDutyDetailMap || {}, perStaffMap: perMap };
            } catch (e) { console.error('buildResultsDebug synth failed', e); return null; }
        };
        window.updateAlertsDebug = function () { try { if (typeof updateAlerts === 'function') return updateAlerts(); console.warn('updateAlerts not defined in this scope'); } catch (e) { console.error('updateAlertsDebug failed', e); } };
        window.getAlertsState = function () {
            try {
                if (typeof buildResults === 'function') return buildResults();
            } catch (e) { if (window.__ALERTS_DEBUG) console.debug('getAlertsState: buildResults failed', e); }
            try { return window.buildResultsDebug(); } catch (e) { console.error('getAlertsState failed to synthesize', e); return null; }
        };
    }
} catch (e) { /* ignore window attach errors */ }

        // Robust diagnostic that does not rely on buildResults existing as a callable symbol
        try {
            if (typeof window !== 'undefined') {
                window.runAlertsDiagnostics = function () {
                    try {
                        const res = { nextDutyMap: {}, lastDutyMap: {}, nextDutyDetailMap: {}, lastDutyDetailMap: {}, ignoredDuties: [] };
                        const rawSel = localStorage.getItem('TERRA_ALERTS_SELECTED_FILES');
                        const sel = rawSel ? JSON.parse(rawSel) : [];
                        let files = [];
                        if (typeof getParsedFiles === 'function') files = getParsedFiles().filter(f => !sel.length || sel.includes(f.id));
                        else if (typeof getLibraryItems === 'function') files = (getLibraryItems() || []).filter(f => f && f.isParsedFile && (!sel.length || sel.includes(f.id)));
                        const toLocalIso = (dt) => { if (!(dt instanceof Date) || isNaN(dt)) return null; const y = dt.getFullYear(); const m = String(dt.getMonth()+1).padStart(2,'0'); const d = String(dt.getDate()).padStart(2,'0'); return `${y}-${m}-${d}`; };
                        const today = new Date(); today.setHours(0,0,0,0);
                        const nextDutyMeta = {}, lastDutyMeta = {};
                        files.forEach(f => {
                            (f.dataRows || []).forEach(r => {
                                try {
                                    const staffVal = (typeof getStaffKey === 'function') ? getStaffKey(r) : (String(r['Staff Number']||r['Assignment No']||r['Staff']||'').trim());
                                    if (!staffVal) return;
                                    const dutyRaw = (typeof getValueByCanonical === 'function') ? (getValueByCanonical(r, 'Duty Date') || getValueByCanonical(r, 'Next Duty')) : (r['Duty Date']||r['Next Duty']||r['Duty']||null);
                                    if (!dutyRaw) return;
                                    let dutyTypeVal = (typeof getValueByCanonical === 'function') ? getValueByCanonical(r, 'Shift Type') : (r['Shift Type']||r['ShiftType']||r['Type']||null);
                                    if (dutyTypeVal == null) {
                                        res.ignoredDuties.push({ file: f.name, rowObj: r, reason: 'Shift Type missing', dutyRaw });
                                        return;
                                    }
                                    const low = String(dutyTypeVal || '').trim().toLowerCase();
                                    const tokens = low.split(/[^a-z]+/).filter(Boolean);
                                    let prio = -1;
                                    if (/(?:\b|^)(?:ld|mgt|do)(?:\b|$)/i.test(low) || tokens.includes('rest')) prio = 0;
                                    else if (tokens.includes('combined')) prio = 3;
                                    else if (tokens.includes('day') || tokens.includes('night') || tokens.includes('n') || tokens.includes('am') || tokens.includes('pm') || tokens.includes('early') || tokens.includes('late') || tokens.includes('eve') || tokens.includes('evening') || tokens.includes('morning') || tokens.includes('oncall') ) prio = 2;
                                    else if (tokens.includes('rest')) prio = 0;
                                    if (prio < 0) { res.ignoredDuties.push({ file: f.name, rowObj: r, reason: 'Shift Type not counted', dutyRaw, shift: dutyTypeVal }); return; }
                                    // parseToDate if available
                                    let dt = null;
                                    try { dt = (typeof parseToDate === 'function') ? parseToDate(dutyRaw, f._locale || 'uk') : (new Date(String(dutyRaw))); } catch (e) { dt = null; }
                                    if (!dt || isNaN(dt)) { res.ignoredDuties.push({ file: f.name, rowObj: r, reason: 'Duty date parse failed', dutyRaw }); return; }
                                    const d0 = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
                                    const iso = toLocalIso(dt) || toLocalIso(d0);
                                    if (!iso) return;
                                    const key = String(staffVal).trim();
                                    if (!key) return;
                                    if (d0.getTime() > today.getTime()) {
                                        const meta = nextDutyMeta[key];
                                        if (!meta) { nextDutyMeta[key] = { iso, prio }; res.nextDutyMap[key] = iso; res.nextDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw }; }
                                        else { if (iso < meta.iso) { nextDutyMeta[key] = { iso, prio }; res.nextDutyMap[key] = iso; res.nextDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw }; } else if (iso === meta.iso && prio > meta.prio) { nextDutyMeta[key] = { iso, prio }; res.nextDutyMap[key] = iso; res.nextDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw }; } }
                                    }
                                    if (d0.getTime() <= today.getTime()) {
                                        const meta = lastDutyMeta[key];
                                        if (!meta) { lastDutyMeta[key] = { iso, prio }; res.lastDutyMap[key] = iso; res.lastDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw }; }
                                        else { if (iso > meta.iso) { lastDutyMeta[key] = { iso, prio }; res.lastDutyMap[key] = iso; res.lastDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw }; } else if (iso === meta.iso && prio > meta.prio) { lastDutyMeta[key] = { iso, prio }; res.lastDutyMap[key] = iso; res.lastDutyDetailMap[key] = { iso, prio, shift: dutyTypeVal, dutyRaw }; } }
                                    }
                                } catch (e) { /* per-row ignore */ }
                            });
                        });
                        return res;
                    } catch (e) { console.error('runAlertsDiagnostics failed', e); return null; }
                };
            }
        } catch (e) { /* ignore */ }

// Diagnostic helper: run from browser console to print per-file/row date parsing and last-N-days match info
try {
    if (typeof window !== 'undefined') {
        window.__alertDiagnostics = function(days = 7) {
            try {
                const raw = localStorage.getItem('TERRA_ALERTS_SELECTED_FILES');
                const sel = raw ? JSON.parse(raw) : [];
                const files = (getLibraryItems() || []).filter(f => f && f.isParsedFile && sel.includes(f.id));
                const toLocalIso = (dt) => { if (!(dt instanceof Date) || isNaN(dt)) return null; const y = dt.getFullYear(); const m = String(dt.getMonth()+1).padStart(2,'0'); const d = String(dt.getDate()).padStart(2,'0'); return `${y}-${m}-${d}`; };
                const now = new Date(); const nowIso = toLocalIso(now);
                const start = new Date(now); start.setDate(start.getDate() - (Math.floor(days) - 1)); const startIso = toLocalIso(start);
                console.group(`Alerts diagnostics - last ${days} day(s) window ${startIso} -> ${nowIso}`);
                files.forEach(f => {
                    console.group(`File: ${f.name} (locale: ${f._locale || 'uk'}) rows:${(f.dataRows||[]).length}`);
                    (f.dataRows || []).forEach((r, i) => {
                            Object.keys(r).forEach(k => {
                                const rawVal = r[k];
                                // diagParse mirrors parseToDate behaviour but is local to this helper
                                const diagParse = (input, locale) => {
                                    if (!input && input !== 0) return null;
                                    if (input instanceof Date && !isNaN(input)) return input;
                                    const s = String(input || '').trim();
                                    const isoMatch = s.match(/^\s*(\d{4})-(\d{2})-(\d{2})\s*$/);
                                    if (isoMatch) {
                                        const y = Number(isoMatch[1]), mo = Number(isoMatch[2]), d = Number(isoMatch[3]);
                                        const dtIso = new Date(y, mo - 1, d);
                                        if (!isNaN(dtIso)) return dtIso;
                                    }
                                    try {
                                        const parsedStr = parseDateByLocale(s, locale);
                                        if (parsedStr) {
                                            const m = parsedStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
                                            if (m) {
                                                const d = Number(m[1]); const mon = Number(m[2]); let yy = m[3];
                                                if (yy.length === 2) yy = (parseInt(yy,10) > 50 ? '19'+yy : '20'+yy);
                                                const dt = new Date(Number(yy), mon - 1, d);
                                                if (!isNaN(dt)) return dt;
                                            }
                                        }
                                        if (/^\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*$/.test(s)) {
                                            const other = (locale === 'uk') ? 'us' : 'uk';
                                            const parsed2 = parseDateByLocale(s, other);
                                            if (parsed2) {
                                                const m2 = parsed2.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
                                                if (m2) {
                                                    const d2 = Number(m2[1]); const mon2 = Number(m2[2]); let yy2 = m2[3];
                                                    if (yy2.length === 2) yy2 = (parseInt(yy2,10) > 50 ? '19'+yy2 : '20'+yy2);
                                                    const dt2 = new Date(Number(yy2), mon2 - 1, d2);
                                                    if (!isNaN(dt2)) return dt2;
                                                }
                                            }
                                        }
                                    } catch (e) { /* ignore */ }
                                    const t = Date.parse(s);
                                    if (!isNaN(t)) return new Date(t);
                                    return null;
                                };
                                const parsed = diagParse(rawVal, f._locale || 'uk');
                                const parsedIso = parsed ? toLocalIso(parsed) : null;
                                const inWindow = parsedIso ? (parsedIso >= startIso && parsedIso <= nowIso) : false;
                                console.log({rowIndex: i, column: k, raw: rawVal, parsedIso, fileLocale: f._locale || 'uk', inWindow});
                            });
                        });
                    console.groupEnd();
                });
                console.groupEnd();
            } catch (e) { console.error('alertDiagnostics failed', e); }
        };
    }
} catch (e) { /* ignore */ }

        // Future-looking diagnostics helper for within_next_days
        try {
            if (typeof window !== 'undefined') {
                window.__alertFutureDiagnostics = function(days = 7) {
                    try {
                        const raw = localStorage.getItem('TERRA_ALERTS_SELECTED_FILES');
                        const sel = raw ? JSON.parse(raw) : [];
                        const files = (getLibraryItems() || []).filter(f => f && f.isParsedFile && sel.includes(f.id));
                        const toLocalIso = (dt) => { if (!(dt instanceof Date) || isNaN(dt)) return null; const y = dt.getFullYear(); const m = String(dt.getMonth()+1).padStart(2,'0'); const d = String(dt.getDate()).padStart(2,'0'); return `${y}-${m}-${d}`; };
                        const now = new Date(); const nowIso = toLocalIso(now);
                        const end = new Date(now); end.setDate(end.getDate() + (Math.floor(days) || 0)); const endIso = toLocalIso(end);
                        console.group(`Alerts future diagnostics - next ${days} day(s) window ${nowIso} -> ${endIso}`);
                        files.forEach(f => {
                            console.group(`File: ${f.name} (locale: ${f._locale || 'uk'}) rows:${(f.dataRows||[]).length}`);
                            (f.dataRows || []).forEach((r,i) => {
                                Object.keys(r).forEach(k => {
                                    const rawVal = r[k];
                                    const parsed = (function diagParse(input, locale) {
                                        if (!input && input !== 0) return null;
                                        if (input instanceof Date && !isNaN(input)) return input;
                                        const s = String(input || '').trim();
                                        const isoMatch = s.match(/^\s*(\d{4})-(\d{2})-(\d{2})\s*$/);
                                        if (isoMatch) { const y = Number(isoMatch[1]), mo = Number(isoMatch[2]), d = Number(isoMatch[3]); const dtIso = new Date(y, mo-1, d); if (!isNaN(dtIso)) return dtIso; }
                                        try { const parsedStr = parseDateByLocale(s, locale); if (parsedStr) { const m = parsedStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/); if (m) { const d = Number(m[1]); const mon = Number(m[2]); let yy = m[3]; if (yy.length===2) yy = (parseInt(yy,10) > 50 ? '19'+yy : '20'+yy); const dt = new Date(Number(yy), mon-1, d); if (!isNaN(dt)) return dt; } } } catch(e){}
                                        const t = Date.parse(s); if (!isNaN(t)) return new Date(t); return null;
                                    })(rawVal, f._locale || 'uk');
                                    const parsedIso = parsed ? toLocalIso(parsed) : null;
                                    const inWindow = parsedIso ? (parsedIso >= nowIso && parsedIso <= endIso) : false;
                                    if (inWindow) console.log({rowIndex: i, column: k, raw: rawVal, parsedIso, fileLocale: f._locale || 'uk', inWindow});
                                });
                            });
                            console.groupEnd();
                        });
                        console.groupEnd();
                    } catch (e) { console.error('alertFutureDiagnostics failed', e); }
                };
                // Debug helpers exposed for console inspection
                try {
                    if (typeof window !== 'undefined') {
                        window.__getParsedFiles = function () {
                            try {
                                if (typeof getLibraryItems === 'function') return (getLibraryItems() || []).filter(i => i && i.isParsedFile && Array.isArray(i.dataRows));
                                return [];
                            } catch (e) { return []; }
                        };
                        window.__dumpParsedFile = function (fileNameOrId, rowIndex) {
                            try {
                                const files = window.__getParsedFiles();
                                const f = files.find(x => x.name === fileNameOrId || x.id === fileNameOrId || (x.name||'').includes(fileNameOrId));
                                if (!f) { console.warn('file not found', fileNameOrId); return null; }
                                if (typeof rowIndex === 'number') { console.log('row', rowIndex, f.dataRows[rowIndex]); return f.dataRows[rowIndex]; }
                                console.log('file', f.name, 'rows', f.dataRows.length); return f.dataRows;
                            } catch (e) { console.error('dumpParsedFile failed', e); return null; }
                        };
                    }
                } catch (e) { /* ignore */ }
            }
        } catch (e) { /* ignore */ }

        // Global helper: compute RTW stats across the whole library and expose for console use.
        try {
            if (typeof window !== 'undefined') {
                window.computeRTWLibraryStats = function (opts) {
                    try {
                        opts = opts || {};
                        const includeSources = !!opts.includeSources;
                        // get parsed files from in-memory library; prefer getParsedFiles if available, otherwise use getLibraryItems
                        let files = [];
                        try {
                            if (typeof getParsedFiles === 'function') files = getParsedFiles();
                            else if (typeof getLibraryItems === 'function') files = (getLibraryItems() || []).filter(i => i && i.isParsedFile);
                        } catch (e) { files = []; }
                        // reuse computeRTWStats if available, otherwise compute locally to avoid ReferenceError
                        let perStaff = [];
                        // local helper: resolve a canonical header value from a row using alias mapping
                        // Use the global, tolerant getValueByCanonical when available so we don't duplicate
                        // matching logic (case/punctuation-insensitive, alias map, plural-aware).
                        const aliasMapForHelper = (typeof buildColumnAliasMap === 'function') ? buildColumnAliasMap() : {};
                        function localGetValueByCanonical(row, canonical) {
                            try {
                                if (typeof getValueByCanonical === 'function') return getValueByCanonical(row, canonical);
                            } catch (e) { /* ignore and fallback */ }
                            // Fallback: simple direct lookup using alias map if global helper isn't present
                            if (!row || !canonical) return undefined;
                            const canonicalToAliases = {};
                            Object.keys(aliasMapForHelper).forEach(k => {
                                const can = aliasMapForHelper[k];
                                canonicalToAliases[can] = canonicalToAliases[can] || new Set();
                                canonicalToAliases[can].add(k);
                            });
                            const candidates = new Set();
                            candidates.add(canonical);
                            (canonicalToAliases[canonical] || []).forEach(a => candidates.add(a));
                            for (const key of candidates) {
                                if (key in row) return row[key];
                            }
                            return undefined;
                        }
                        if (typeof computeRTWStats === 'function') {
                            perStaff = computeRTWStats(files, null);
                        } else {
                            // local computation mirroring computeRTWStats
                            const sicknessCols = ['Sickness End','Sickness End Date','Sick End Date','SicknessEnd','Sick End','Sickness_End','End','End Date'];
                            const dutyCols = ['Duty Date','DutyDate','Duty','Shift Date','ShiftDate','Next Duty','NextDuty','Next Duty Date','NextDutyDate'];
                            const rtwCols = ['Return to Work','RTW','Return to work','Return','ReturnToWork','Return To Work Interview Completed','Return To Work Interview','RTW Interview Completed','RTW Interview'];
                            const staffKeyFn = (r) => {
                                try { return String((typeof getValueByCanonical === 'function' ? getValueByCanonical(r, 'Staff Number') : localGetValueByCanonical(r, 'Staff Number')) || (typeof getStaffKey === 'function' ? getStaffKey(r) : localGetValueByCanonical(r, 'Assignment No')) || '').trim(); } catch (e) { return String(localGetValueByCanonical(r, 'Staff Number') || localGetValueByCanonical(r, 'Assignment No') || localGetValueByCanonical(r, 'Staff') || '').trim(); }
                            };
                            // local parseToDate fallback using parseDateByLocale (imported from ui.js) where available
                            function localParseToDate(input, locale) {
                                if (!input && input !== 0) return null;
                                if (input instanceof Date && !isNaN(input)) return input;
                                const s = String(input || '').trim();
                                // If already ISO YYYY-MM-DD, prefer that
                                const isoMatch = s.match(/^\s*(\d{4})-(\d{2})-(\d{2})\s*$/);
                                if (isoMatch) {
                                    const y = Number(isoMatch[1]), mo = Number(isoMatch[2]), d = Number(isoMatch[3]);
                                    const dtIso = new Date(y, mo - 1, d);
                                    if (!isNaN(dtIso)) return dtIso;
                                }
                                try {
                                    if (typeof parseDateByLocale === 'function') {
                                        const parsedStr = parseDateByLocale(s, locale);
                                        if (parsedStr) {
                                            // parsedStr may be ISO or DD/MM/YYYY depending on helper
                                            const m = parsedStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
                                            if (m) {
                                                const d = Number(m[1]); const mon = Number(m[2]); let yy = m[3];
                                                if (yy.length === 2) yy = (parseInt(yy,10) > 50 ? '19'+yy : '20'+yy);
                                                const dt = new Date(Number(yy), mon - 1, d);
                                                if (!isNaN(dt)) return dt;
                                            }
                                            // If parseDateByLocale returned ISO, parse it
                                            const iso2 = parsedStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
                                            if (iso2) {
                                                const y2 = Number(iso2[1]), mo2 = Number(iso2[2]), d2 = Number(iso2[3]);
                                                const dt2 = new Date(y2, mo2 - 1, d2);
                                                if (!isNaN(dt2)) return dt2;
                                            }
                                        }
                                    }
                                    // If slashed ambiguous, try alternate locale
                                    if (/^\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*$/.test(s)) {
                                        const other = (locale === 'uk') ? 'us' : 'uk';
                                        if (typeof parseDateByLocale === 'function') {
                                            const parsed2 = parseDateByLocale(s, other);
                                            if (parsed2) {
                                                const m2 = parsed2.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
                                                if (m2) {
                                                    const d2 = Number(m2[1]); const mon2 = Number(m2[2]); let yy2 = m2[3];
                                                    if (yy2.length === 2) yy2 = (parseInt(yy2,10) > 50 ? '19'+yy2 : '20'+yy2);
                                                    const dt2 = new Date(Number(yy2), mon2 - 1, d2);
                                                    if (!isNaN(dt2)) return dt2;
                                                }
                                            }
                                        }
                                    }
                                } catch (e) { /* ignore */ }
                                const t = Date.parse(s);
                                if (!isNaN(t)) return new Date(t);
                                return null;
                            }
                            const staffMap = {};
                            files.forEach(f => {
                                (f.dataRows || []).forEach(r => {
                                    const staff = String(staffKeyFn(r) || '').trim();
                                    if (!staff) return;
                                    staffMap[staff] = staffMap[staff] || { staff, sicknessDates: [], dutyDates: [], rtwFlags: [] };
                                    const entry = staffMap[staff];
                                    for (const sc of sicknessCols) {
                                        const raw = localGetValueByCanonical(r, sc);
                                        if (raw != null) { const dt = localParseToDate(raw, f._locale || 'uk'); if (dt) entry.sicknessDates.push(dt); break; }
                                    }
                                        // Determine whether this row represents a Rest shift. We must skip Rest rows when
                                        // collecting dutyDates so RTW and Next Duty calculations don't count them.
                                        let isRestShift = false;
                                        try {
                                            // Try canonical Shift Type first, then common alternates.
                                            let dutyTypeVal = localGetValueByCanonical(r, 'Shift Type');
                                            if (dutyTypeVal == null) {
                                                const tryAlts = ['ShiftType','Type','Roster Type','Assignment Info','Assignment'];
                                                for (const alt of tryAlts) {
                                                    const v = localGetValueByCanonical(r, alt);
                                                    if (v != null) { dutyTypeVal = v; break; }
                                                }
                                            }
                                            if (dutyTypeVal != null) {
                                                const low = String(dutyTypeVal || '').trim().toLowerCase();
                                                const tokens = low.split(/[^a-z]+/).filter(Boolean);
                                                if (tokens.includes('rest') || tokens.includes('ld') || tokens.includes('mgt') || tokens.includes('do')) isRestShift = true;
                                            }
                                            // Permissive fallback: scan all field values for the whole word 'rest'
                                            if (!isRestShift) {
                                                const fieldToString = (val) => {
                                                    if (val == null) return '';
                                                    if (typeof val === 'string') return val.trim();
                                                    try { if (typeof val === 'object') return JSON.stringify(val); } catch (e) {}
                                                    try { return String(val); } catch (e) { return ''; }
                                                };
                                                for (const key of Object.keys(r || {})) {
                                                    try {
                                                        const s = fieldToString(r[key]);
                                                        if (!s) continue;
                                                        if (/(?:\b|^)(?:rest|ld|mgt|do)(?:\b|$)/i.test(s)) { isRestShift = true; break; }
                                                    } catch (e) { /* ignore per-field errors */ }
                                                }
                                            }
                                        } catch (e) { /* ignore detection errors */ }

                                        // Only add dutyDates when this is NOT a Rest shift
                                        if (!isRestShift) {
                                            for (const dc of dutyCols) {
                                                const raw = localGetValueByCanonical(r, dc);
                                                if (raw != null) { const dt = localParseToDate(raw, f._locale || 'uk'); if (dt) entry.dutyDates.push(dt); break; }
                                            }
                                        } else {
                                            // For diagnostics it's useful to know we skipped a Rest row
                                            try { if (window.__ALERTS_DEBUG) console.debug('computeRTWLibraryStats: skipping Rest row for staff', entry.staff, 'file', f.name, 'row', idx); } catch (e) {}
                                        }
                                    for (const rc of rtwCols) {
                                        const raw = localGetValueByCanonical(r, rc);
                                        if (raw != null) { entry.rtwFlags.push(String(raw || '').trim()); break; }
                                    }
                                });
                            });
                            // Evaluate per-staff
                            const today = new Date(); today.setHours(0,0,0,0);
                            Object.values(staffMap).forEach(s => {
                                const sickness = s.sicknessDates.length ? new Date(Math.max(...s.sicknessDates.map(d=>d.getTime()))) : null;
                                const rtwDone = s.rtwFlags.some(f => { const v = String(f || '').toLowerCase(); return v === 'y' || v === 'yes' || v === 'true' || v === '1'; });
                                const sortedDuties = s.dutyDates.length ? s.dutyDates.slice().sort((a,b) => a.getTime() - b.getTime()) : [];
                                const nextDuty = sortedDuties.find(d => { const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate()); return d0.getTime() > today.getTime(); }) || null;
                                const lastOnOrBefore = sortedDuties.filter(d => { const d0 = new Date(d.getFullYear(), d.getMonth(), d.getDate()); return d0.getTime() <= today.getTime(); });
                                const lastDuty = lastOnOrBefore.length ? lastOnOrBefore[lastOnOrBefore.length - 1] : null;
                                const hadShiftAfter = sickness && lastDuty && lastDuty.getTime() >= sickness.getTime();
                                const sampleDuty = sortedDuties.length ? sortedDuties[0] : null;
                                const sicknessEnded = !!(sickness && sickness.getTime() <= today.getTime());
                                const hadShiftAfterSickness = !!hadShiftAfter;
                                const rtwInterviewDone = !!rtwDone;
                                // Heuristic: if there are multiple raw sickness dates recorded for this staff
                                // treat as continuing sickness so the UI can surface it even when ranges
                                // didn't merge (for example due to differing Reason text).
                                const continuingSickness = (Array.isArray(s.sicknessDates) && s.sicknessDates.length > 1) || false;
                                perStaff.push({ staff: s.staff, sickness, sicknessEnded, hadShiftAfterSickness, hadShiftAfter: !!hadShiftAfter, rtwInterviewDone, rtwDone: !!rtwDone, sampleDuty, nextDuty, lastDuty, continuingSickness });
                            });
                        }
                        if (!includeSources) return perStaff;
                        // If sources requested, build a map of which files contributed sickness/duty/rtw for each staff
                        const staffSources = {};
                        files.forEach(f => {
                (f.dataRows || []).forEach(r => {
                    // Prefer centralized getStaffKey when present, otherwise fall back to localGetValueByCanonical
                    let staff = '';
                    try {
                        if (typeof getStaffKey === 'function') staff = String(getStaffKey(r) || '').trim();
                        else {
                            try {
                                if (typeof getStaffKey === 'function') staff = String(getStaffKey(r) || '').trim();
                                else staff = String(localGetValueByCanonical(r, 'Staff Number') || localGetValueByCanonical(r, 'Assignment No') || localGetValueByCanonical(r, 'Staff') || '').trim();
                            } catch (e) { staff = String(localGetValueByCanonical(r, 'Staff Number') || localGetValueByCanonical(r, 'Assignment No') || localGetValueByCanonical(r, 'Staff') || '').trim(); }
                        }
                    } catch (e) { /* ignore staff extraction errors */ }
                    if (!staff) return;
                    staffSources[staff] = staffSources[staff] || { sicknessFiles: new Set(), dutyFiles: new Set(), rtwFiles: new Set() };
                                // sickness
                                const sicknessCols = ['Sickness End','Sickness End Date','Sick End Date','SicknessEnd','Sick End','Sickness_End'];
                                for (const sc of sicknessCols) {
                                    const raw = localGetValueByCanonical(r, sc);
                                    if (raw != null) { staffSources[staff].sicknessFiles.add(f.name || f.id || 'unknown'); break; }
                                }
                                // duty
                                const dutyCols = ['Duty Date','DutyDate','Duty','Shift Date','ShiftDate'];
                                for (const dc of dutyCols) {
                                    const raw = localGetValueByCanonical(r, dc);
                                    if (raw != null) { staffSources[staff].dutyFiles.add(f.name || f.id || 'unknown'); break; }
                                }
                                // rtw
                                const rtwCols = ['Return to Work','RTW','Return to work','Return','ReturnToWork'];
                                for (const rc of rtwCols) {
                                    const raw = localGetValueByCanonical(r, rc);
                                    if (raw != null) { staffSources[staff].rtwFiles.add(f.name || f.id || 'unknown'); break; }
                                }
                            });
                        });
                        // attach sources lists to the perStaff objects
                        return perStaff.map(s => ({ ...s, sources: { sicknessFiles: Array.from((staffSources[s.staff] && staffSources[s.staff].sicknessFiles) || []), dutyFiles: Array.from((staffSources[s.staff] && staffSources[s.staff].dutyFiles) || []), rtwFiles: Array.from((staffSources[s.staff] && staffSources[s.staff].rtwFiles) || []) } }));
                    } catch (e) {
                        console.error('computeRTWLibraryStats failed', e);
                        return [];
                    }
                };
            }
        } catch (e) { /* ignore */ }

function setupSettings() {
    const settingsBtn = document.getElementById('settings-btn');
    settingsBtn.addEventListener('click', () => {
        // create a simple settings modal
        let modal = document.getElementById('settings-modal');
        if (!modal) {
            modal = document.createElement('div');
            modal.id = 'settings-modal';
            modal.className = 'modal-backdrop';
            modal.innerHTML = `
                <div class="modal">
                    <h2 class="text-xl font-bold">Settings</h2>
                    <div class="mt-4">
                        <label for="theme-swatches">Theme</label>
                        <div id="theme-swatches" class="theme-swatch-grid mt-2" role="list" aria-label="Theme choices">
                          <button role="listitem" class="theme-swatch" data-theme="light" title="Light" aria-pressed="false"><span class="swatch-label">Light</span></button>
                          <button role="listitem" class="theme-swatch" data-theme="dark" title="Dark" aria-pressed="false"><span class="swatch-label">Dark</span></button>
                          <button role="listitem" class="theme-swatch" data-theme="midnight" title="Midnight" aria-pressed="false"><span class="swatch-label">Midnight</span></button>
                          <button role="listitem" class="theme-swatch" data-theme="teal" title="Teal" aria-pressed="false"><span class="swatch-label">Teal</span></button>
                          <button role="listitem" class="theme-swatch" data-theme="coral" title="Coral" aria-pressed="false"><span class="swatch-label">Coral</span></button>
                          <button role="listitem" class="theme-swatch" data-theme="forest" title="Forest" aria-pressed="false"><span class="swatch-label">Forest</span></button>
                          <button role="listitem" class="theme-swatch" data-theme="sunset" title="Sunset" aria-pressed="false"><span class="swatch-label">Sunset</span></button>
                          <button role="listitem" class="theme-swatch" data-theme="lavender" title="Lavender" aria-pressed="false"><span class="swatch-label">Lavender</span></button>
                        </div>
                        <div class="mt-3">
                          <button id="save-theme-btn" class="px-3 py-1 rounded bg-blue-600 text-white">Save theme</button>
                          <button id="reset-theme-btn" class="px-3 py-1 rounded btn-secondary ml-2">Reset</button>
                        </div>
                    </div>
                    <div class="mt-4">
                        <label for="merge-by-reason">Merge adjacent sickness rows only when Reason matches</label>
                        <div class="mt-2 text-sm"><input type="checkbox" id="merge-by-reason" /> <span class="text-subtle">Require identical Reason text to merge adjacent sickness ranges</span></div>
                    </div>
                    <div class="mt-4">
                        <label for="date-display-select">Date display</label>
                        <select id="date-display-select" class="block mt-2 p-2">
                            <option value="uk">UK (DD/MM/YYYY)</option>
                            <option value="us">US (MM/DD/YYYY)</option>
                            <option value="iso">ISO (YYYY-MM-DD)</option>
                        </select>
                        <div id="date-display-preview" class="text-sm text-subtle mt-2">Preview: </div>
                    </div>
                    <div class="mt-4">
                        <label for="ambig-date-pref">Ambiguous slashed dates</label>
                        <select id="ambig-date-pref" class="block mt-2 p-2">
                            <option value="auto">Auto (prefer UK)</option>
                            <option value="uk">Always UK (DD/MM)</option>
                            <option value="us">Always US (MM/DD)</option>
                        </select>
                    </div>
                    <div class="mt-4">
                        <button id="clear-draft-btn" class="text-red-600">Clear all saved data</button>
                        <button id="close-settings" class="ml-2">Close</button>
                    </div>
                </div>`;
            document.body.appendChild(modal);
            modal.querySelector('#close-settings').addEventListener('click', () => modal.classList.remove('active'));
            modal.querySelector('#clear-draft-btn').addEventListener('click', async () => {
                if (!confirm('Delete ALL saved data?')) return;
                await Promise.all(Object.values(STORES).map(s => dbClear(s)));
                // clear runtime state: categories and library files
                try {
                    categories.length = 0;
                } catch (e) { console.warn('Failed to clear categories array', e); }
                try {
                    // clear the shared library maintained in categories module
                    clearLibrary();
                } catch (e) { console.warn('Failed to clear libraryFiles', e); }
                // re-render the initial UI
                try { renderCategoryUI(); } catch (e) { console.error('Failed to re-render UI after clearing data', e); }
                toast('Cleared saved data.');
                modal.classList.remove('active');
            });
            // wire theme swatches
            const swatchesContainer = modal.querySelector('#theme-swatches');
            const swatches = Array.from(modal.querySelectorAll('.theme-swatch'));
            let pendingTheme = localStorage.getItem('TERRA_THEME') || 'light';

            function applyPreviewTheme(name) {
                if (!name) return;
                document.documentElement.setAttribute('data-theme', name);
            }

            function updateSwatchStates(activeName) {
                swatches.forEach(s => {
                    const name = s.dataset.theme;
                    const pressed = name === activeName;
                    s.setAttribute('aria-pressed', pressed ? 'true' : 'false');
                    s.classList.toggle('selected', pressed);
                });
            }

            // initialize
            updateSwatchStates(pendingTheme);

            // initialize merge-by-reason checkbox
            const mergeByReasonCheckbox = modal.querySelector('#merge-by-reason');
            try {
                const saved = localStorage.getItem('TERRA_MERGE_SICKNESS_BY_REASON');
                mergeByReasonCheckbox.checked = saved === 'true';
            } catch (e) { mergeByReasonCheckbox.checked = false; }

            swatches.forEach(s => {
                const name = s.dataset.theme;
                // live preview on mouseenter/focus (temporary)
                s.addEventListener('mouseenter', () => applyPreviewTheme(name));
                s.addEventListener('focus', () => applyPreviewTheme(name));
                // revert preview when pointer leaves or element loses focus
                s.addEventListener('mouseleave', () => applyPreviewTheme(pendingTheme));
                s.addEventListener('blur', () => applyPreviewTheme(pendingTheme));
                // commit selection on click (updates pendingTheme and previews it)
                s.addEventListener('click', (e) => {
                    e.preventDefault();
                    pendingTheme = name;
                    updateSwatchStates(pendingTheme);
                    applyPreviewTheme(pendingTheme);
                });
                // keyboard: Space/Enter should select and preview; arrow keys navigate
                s.addEventListener('keydown', (e) => {
                    if (e.key === ' ' || e.key === 'Enter') { e.preventDefault(); pendingTheme = name; updateSwatchStates(pendingTheme); applyPreviewTheme(pendingTheme); }
                    if (e.key === 'ArrowRight' || e.key === 'ArrowDown') { e.preventDefault(); const idx = swatches.indexOf(s); const next = swatches[(idx+1)%swatches.length]; next.focus(); }
                    if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') { e.preventDefault(); const idx = swatches.indexOf(s); const prev = swatches[(idx-1+swatches.length)%swatches.length]; prev.focus(); }
                });
            });

            // Unsaved indicator element next to save/reset
            const btnHolder = modal.querySelector('.mt-3');
            const unsavedEl = document.createElement('span');
            unsavedEl.id = 'settings-unsaved';
            unsavedEl.className = 'ml-3 text-sm text-subtle';
            unsavedEl.style.display = 'none';
            unsavedEl.textContent = 'Unsaved changes';
            if (btnHolder) btnHolder.appendChild(unsavedEl);

            const saveBtn = modal.querySelector('#save-theme-btn');
            const resetBtn = modal.querySelector('#reset-theme-btn');

            // capture saved values at modal open for change detection
            const savedThemeAtOpen = localStorage.getItem('TERRA_THEME') || 'light';
            const savedMergeAtOpen = localStorage.getItem('TERRA_MERGE_SICKNESS_BY_REASON') === 'true';
            const savedDateDisplayAtOpen = localStorage.getItem('TERRA_DATE_DISPLAY_V2') || localStorage.getItem('TERRA_DATE_DISPLAY') || 'uk';
            const savedAmbigAtOpen = localStorage.getItem('TERRA_AMBIG_DATE_PREF') || 'auto';

            function updateUnsavedIndicator() {
                try {
                    const mergeNow = modal.querySelector('#merge-by-reason') ? modal.querySelector('#merge-by-reason').checked : false;
                    const dateNow = modal.querySelector('#date-display-select') ? modal.querySelector('#date-display-select').value : savedDateDisplayAtOpen;
                    const ambigNow = modal.querySelector('#ambig-date-pref') ? modal.querySelector('#ambig-date-pref').value : savedAmbigAtOpen;
                    const themeNow = pendingTheme || savedThemeAtOpen;
                    const changed = (String(themeNow) !== String(savedThemeAtOpen)) || (mergeNow !== savedMergeAtOpen) || (dateNow !== savedDateDisplayAtOpen) || (ambigNow !== savedAmbigAtOpen);
                    unsavedEl.style.display = changed ? 'inline-block' : 'none';
                } catch (e) { /* ignore */ }
            }

            // wire up controls to update unsaved indicator
            try {
                modal.querySelectorAll('.theme-swatch').forEach(s => s.addEventListener('click', updateUnsavedIndicator));
                const mergeCheckbox = modal.querySelector('#merge-by-reason'); if (mergeCheckbox) mergeCheckbox.addEventListener('change', updateUnsavedIndicator);
                const dateSel = modal.querySelector('#date-display-select'); if (dateSel) dateSel.addEventListener('change', updateUnsavedIndicator);
                const ambigSel = modal.querySelector('#ambig-date-pref'); if (ambigSel) ambigSel.addEventListener('change', updateUnsavedIndicator);
            } catch (e) { /* ignore wiring errors */ }

            // initialise unsaved indicator
            updateUnsavedIndicator();

            saveBtn.addEventListener('click', () => {
                const t = pendingTheme || 'light';
                localStorage.setItem('TERRA_THEME', t);
                document.documentElement.setAttribute('data-theme', t);
                updateSwatchStates(t);
                try {
                    const mergeByReasonCheckboxNow = modal.querySelector('#merge-by-reason');
                    if (mergeByReasonCheckboxNow) localStorage.setItem('TERRA_MERGE_SICKNESS_BY_REASON', mergeByReasonCheckboxNow.checked ? 'true' : 'false');
                } catch (e) { /* ignore */ }
                try {
                    const dateVal = modal.querySelector('#date-display-select'); if (dateVal) localStorage.setItem('TERRA_DATE_DISPLAY_V2', dateVal.value || 'uk');
                } catch (e) { /* ignore */ }
                try {
                    const ambigVal = modal.querySelector('#ambig-date-pref'); if (ambigVal) localStorage.setItem('TERRA_AMBIG_DATE_PREF', ambigVal.value || 'auto');
                } catch (e) { /* ignore */ }
                // show a general saved toast
                toast('Settings saved');
                // reset unsaved indicator baseline
                unsavedEl.style.display = 'none';
            });

            resetBtn.addEventListener('click', () => {
                localStorage.removeItem('TERRA_THEME');
                const sys = 'light';
                pendingTheme = sys;
                document.documentElement.setAttribute('data-theme', sys);
                updateSwatchStates(sys);
                toast('Theme reset to default');
                try { localStorage.removeItem('TERRA_MERGE_SICKNESS_BY_REASON'); const cb = modal.querySelector('#merge-by-reason'); if (cb) cb.checked = false; } catch (e) { /* ignore */ }
                // reset other persisted settings
                try { localStorage.removeItem('TERRA_DATE_DISPLAY_V2'); } catch (e) {}
                try { localStorage.removeItem('TERRA_AMBIG_DATE_PREF'); } catch (e) {}
                // clear unsaved indicator since reset counts as a committed action
                unsavedEl.style.display = 'none';
            });
        }
        // set current value, default to saved or light
        const saved = localStorage.getItem('TERRA_THEME') || 'light';
        // If the swatch UI is present, update its state; otherwise ensure theme is applied
        const swatchesNow = modal.querySelectorAll('.theme-swatch');
        if (swatchesNow && swatchesNow.length) {
            swatchesNow.forEach(s => {
                const isActive = s.dataset.theme === saved;
                s.classList.toggle('selected', isActive);
                s.setAttribute('aria-pressed', isActive ? 'true' : 'false');
            });
            document.documentElement.setAttribute('data-theme', saved);
        } else {
            document.documentElement.setAttribute('data-theme', saved);
        }
            // Prefer the v2 key to avoid confusion with earlier swapped-label era
            let savedDateDisplay = localStorage.getItem('TERRA_DATE_DISPLAY_V2');
            if (!savedDateDisplay) {
                // fallback and attempt best-effort conversion from legacy key
                const legacy = localStorage.getItem('TERRA_DATE_DISPLAY');
                if (legacy === 'uk' || legacy === 'us' || legacy === 'iso') savedDateDisplay = legacy; else savedDateDisplay = 'uk';
                // normalize into v2
                localStorage.setItem('TERRA_DATE_DISPLAY_V2', savedDateDisplay);
            }
            const dateSel = modal.querySelector('#date-display-select'); if (dateSel) dateSel.value = savedDateDisplay;
            const preview = modal.querySelector('#date-display-preview');
            const updatePreview = (val) => {
                if (!preview) return;
                const sample = { sampleDate: '2025-10-12' };
                const text = val === 'iso' ? sample.sampleDate : (val === 'uk' ? formatByLocale(sample.sampleDate, 'uk') : formatByLocale(sample.sampleDate, 'us'));
                preview.textContent = `Preview: ${text}`;
            };
            if (preview) updatePreview(savedDateDisplay);
            if (dateSel) dateSel.addEventListener('change', (e) => { const v = e.target.value || 'uk'; localStorage.setItem('TERRA_DATE_DISPLAY_V2', v); updatePreview(v); toast(`Date display set to ${v}`); try { document.dispatchEvent(new CustomEvent('terra:dateDisplayChanged', { detail: { value: v } })); } catch (err) { /* ignore */ } });
            const savedAmbig = localStorage.getItem('TERRA_DATE_AMBIGUOUS') || 'auto';
            const ambigSel = modal.querySelector('#ambig-date-pref'); if (ambigSel) ambigSel.value = savedAmbig;
            if (ambigSel) ambigSel.addEventListener('change', (e) => { const v = e.target.value || 'auto'; localStorage.setItem('TERRA_DATE_AMBIGUOUS', v); toast(`Ambiguous date preference set to ${v}`); });
        modal.classList.add('active');
    });

    // Add a quick link to edit column links
    const manageLinksBtn = document.getElementById('manage-column-links');
    if (manageLinksBtn) {
        manageLinksBtn.addEventListener('click', () => {
            let linksModal = document.getElementById('column-links-modal');
            if (!linksModal) {
                linksModal = document.createElement('div');
                linksModal.id = 'column-links-modal';
                linksModal.className = 'modal-backdrop';
                linksModal.innerHTML = `
                    <div class="modal">
                        <div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold">Manage Column Links</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div>
                        <div class="modal-body p-2 max-h-[60vh] overflow-auto">
                            <div id="column-links-list" class="space-y-2"></div>
                            <div class="mt-3 flex gap-2"><input id="new-canonical" placeholder="Canonical name (e.g. Assignment No)" class="p-1 border rounded flex-1"/><input id="new-aliases" placeholder="Aliases (comma separated)" class="p-1 border rounded flex-1"/><button id="add-column-link" class="px-3 py-1 rounded bg-blue-600 text-white">Add</button></div>
                        </div>
                        <div class="modal-footer mt-3 text-right"><button id="close-links" class="btn-secondary px-3 py-1 rounded">Close</button></div>
                    </div>`;
                document.body.appendChild(linksModal);
                const listEl = linksModal.querySelector('#column-links-list');
                function renderLinks() {
                    const current = getColumnLinks();
                    listEl.innerHTML = '';
                    current.forEach((l, i) => {
                        const row = document.createElement('div');
                        row.className = 'p-2 border rounded flex items-center gap-2';
                        const txt = document.createElement('div'); txt.className = 'flex-1 text-sm'; txt.innerHTML = `<strong>${escapeHtml(l.canonical)}</strong><div class="text-subtle text-xs">${escapeHtml((l.aliases||[]).join(', '))}</div>`;
                        const del = document.createElement('button'); del.className = 'text-red-600'; del.textContent = 'Remove';
                        row.appendChild(txt); row.appendChild(del); listEl.appendChild(row);
                        del.addEventListener('click', () => {
                            const arr = getColumnLinks(); arr.splice(i,1); saveColumnLinks(arr); renderLinks();
                        });
                    });
                }
                linksModal.querySelector('.modal-close-btn').addEventListener('click', () => linksModal.classList.remove('active'));
                linksModal.querySelector('#close-links').addEventListener('click', () => linksModal.classList.remove('active'));
                linksModal.addEventListener('click', (e) => { if (e.target === linksModal) linksModal.classList.remove('active'); });
                linksModal.querySelector('#add-column-link').addEventListener('click', () => {
                    const can = linksModal.querySelector('#new-canonical').value.trim();
                    const aliasTxt = linksModal.querySelector('#new-aliases').value.trim();
                    if (!can) return alert('Canonical name required');
                    const aliases = aliasTxt ? aliasTxt.split(',').map(s=>s.trim()).filter(Boolean) : [];
                    const arr = getColumnLinks(); arr.push({ id: `cl_${Date.now()}`, canonical: can, aliases }); saveColumnLinks(arr); renderLinks();
                    linksModal.querySelector('#new-canonical').value = ''; linksModal.querySelector('#new-aliases').value = '';
                });
                renderLinks();
            }
            linksModal.classList.add('active');
        });
    }

    // Wire up the Docs / How-to modal (navbar link) — loads README.md when available
    const docsLink = document.getElementById('docs-link');
    if (docsLink) {
        docsLink.addEventListener('click', async (ev) => {
            ev.preventDefault();
            let howto = document.getElementById('howto-modal');
            if (!howto) {
                howto = document.createElement('div');
                howto.id = 'howto-modal';
                howto.className = 'modal-backdrop';
                howto.innerHTML = `
                    <div class="modal" role="dialog" aria-modal="true" aria-labelledby="howto-title">
                        <div class="flex justify-between items-center mb-2">
                            <h2 id="howto-title" class="text-lg font-semibold">Read Me</h2>
                            <button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading" aria-label="Close">&times;</button>
                        </div>
                        <div class="modal-body max-h-[70vh] overflow-y-auto p-2" tabindex="0" style="line-height:1.6">
                          <div class="howto-tab-buttons mb-3" role="tablist" aria-label="Documentation sections">
                            <button id="howto-tab-builder" role="tab" aria-selected="true" class="tab-btn mr-2">Report Builder</button>
                            <button id="howto-tab-alerts" role="tab" aria-selected="false" class="tab-btn">Alerts</button>
                          </div>
                          <div class="howto-tab-content">
                            <div id="howto-panel-builder" class="tab-panel" role="tabpanel" aria-labelledby="howto-tab-builder">Loading content…</div>
                            <div id="howto-panel-alerts" class="tab-panel" role="tabpanel" aria-labelledby="howto-tab-alerts" hidden></div>
                          </div>
                        </div>
                        <div class="modal-footer mt-3 text-right">
                            <a id="howto-github" class="px-3 py-1 rounded btn-tertiary" target="_blank" rel="noopener">View on GitHub</a>
                            <button id="howto-close" class="btn-secondary px-3 py-1 rounded ml-2">Close</button>
                        </div>
                    </div>`;
                document.body.appendChild(howto);

                // event wiring: close buttons and backdrop
                howto.querySelector('.modal-close-btn').addEventListener('click', () => howto.classList.remove('active'));
                howto.querySelector('#howto-close').addEventListener('click', () => howto.classList.remove('active'));
                howto.addEventListener('click', (e) => { if (e.target === howto) howto.classList.remove('active'); });

                // focus trap helpers
                let lastFocused = null;
                function trapFocus(modalEl) {
                    const focusable = modalEl.querySelectorAll('a[href], button:not([disabled]), textarea, input, select, [tabindex]:not([tabindex="-1"])');
                    const first = focusable[0];
                    const last = focusable[focusable.length - 1];
                    function handleKey(e) {
                        if (e.key === 'Tab') {
                            if (e.shiftKey) {
                                if (document.activeElement === first) { e.preventDefault(); last.focus(); }
                            } else {
                                if (document.activeElement === last) { e.preventDefault(); first.focus(); }
                            }
                        }
                    }
                    modalEl._trapHandler = handleKey;
                    document.addEventListener('keydown', handleKey);
                    // remember and focus first
                    lastFocused = document.activeElement;
                    if (first) first.focus();
                }
                function releaseFocus(modalEl) {
                    if (modalEl && modalEl._trapHandler) {
                        document.removeEventListener('keydown', modalEl._trapHandler);
                        modalEl._trapHandler = null;
                    }
                    try { if (lastFocused && typeof lastFocused.focus === 'function') lastFocused.focus(); } catch (e) { /* ignore */ }
                }

                // small, conservative markdown -> HTML converter for README.md (keeps things simple)
                function escapeHtml(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
                function simpleMarkdownToHtml(md) {
                    if (!md) return '';
                    // Normalize CRLF
                    let out = escapeHtml(md);
                    // code fences
                    out = out.replace(/```([\s\S]*?)```/g, (m, code) => `<pre><code>${escapeHtml(code)}</code></pre>`);
                    // headings
                    out = out.replace(/^######\s?(.*)$/gm, '<h6>$1</h6>');
                    out = out.replace(/^#####\s?(.*)$/gm, '<h5>$1</h5>');
                    out = out.replace(/^####\s?(.*)$/gm, '<h4>$1</h4>');
                    out = out.replace(/^###\s?(.*)$/gm, '<h3>$1</h3>');
                    out = out.replace(/^##\s?(.*)$/gm, '<h2>$1</h2>');
                    out = out.replace(/^#\s?(.*)$/gm, '<h1>$1</h1>');
                    // bold and italics
                    out = out.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
                    out = out.replace(/\*(.*?)\*/g, '<em>$1</em>');
                    // links [text](url)
                    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>');
                    // unordered lists
                    out = out.replace(/(^|\n)\s*[-\*]\s+(.*)/g, '$1<li>$2</li>');
                    out = out.replace(/(<li>[\s\S]*<\/li>)/g, (m) => '<ul>' + m + '</ul>');
                    // paragraphs (simple)
                    out = out.replace(/\n{2,}/g, '</p><p>');
                    out = '<p>' + out + '</p>';
                    // fix wrapping around lists
                    out = out.replace(/<p>\s*<ul>/g, '<ul>');
                    out = out.replace(/<\/ul>\s*<\/p>/g, '</ul>');
                    return out;
                }

                // attempt to fetch README.md, fall back to embedded quick-start and alerts docs
                (async () => {
                    const panelBuilder = howto.querySelector('#howto-panel-builder');
                    const panelAlerts = howto.querySelector('#howto-panel-alerts');
                    const GH_URL = 'https://github.com/leapcoded/ITU-Site-Breakdown#readme';
                    howto.querySelector('#howto-github').setAttribute('href', GH_URL);
                    try {
                        const res = await fetch('./README.md');
                        if (!res.ok) throw new Error('not ok');
                        const md = await res.text();
                        if (md && md.trim()) {
                            // split markdown into builder and alerts sections if possible
                            const alertsRe = new RegExp('^##\\s*Alerts\\b.*$', 'im');
                            const m = alertsRe.exec(md);
                            if (m && typeof m.index === 'number') {
                                const start = m.index;
                                // find next top-level '## ' after start
                                const after = md.slice(start + 1);
                                const nextRe = /^##\s+/m;
                                const nextMatch = nextRe.exec(after);
                                const alertsMd = nextMatch ? md.slice(start, start + 1 + nextMatch.index) : md.slice(start);
                                const builderMd = md.slice(0, start);
                                panelBuilder.innerHTML = simpleMarkdownToHtml(builderMd || '# Report Builder\n\nNo builder docs found.');
                                panelAlerts.innerHTML = simpleMarkdownToHtml(alertsMd);
                            } else {
                                // no Alerts heading found: show full README in builder and provide an alerts summary
                                panelBuilder.innerHTML = simpleMarkdownToHtml(md);
                                panelAlerts.innerHTML = `<h2>Alerts</h2><p>The README does not contain a separate Alerts section. Use the Alerts tab to create rules (contains, equals, within_days, numeric comparisons) to surface rows matching your criteria.</p>`;
                            }
                        } else throw new Error('empty');
                    } catch (err) {
                        // simple, friendly fallback content for non-technical users
                        panelBuilder.innerHTML = `
                            <h3 class="text-md font-semibold">Quick start</h3>
                            <ol class="pl-4 mt-2 text-sm text-subtle">
                                <li class="mb-2">Click + Add Category to make a section.</li>
                                <li class="mb-2">Add files by clicking the area or dragging them in (tables and pictures).</li>
                                <li class="mb-2">Use + Note to add short notes to a category.</li>
                                <li class="mb-2">Click Customize Report to pick columns and order.</li>
                                <li class="mb-2">Click Generate PDF Report to download your report.</li>
                            </ol>
                            <p class="mt-4 text-sm text-subtle">For more help, open the full README on GitHub: <a href="${GH_URL}" target="_blank" rel="noopener">Open README on GitHub</a></p>
                        `;
                        panelAlerts.innerHTML = `
                            <h3 class="text-md font-semibold">Alerts</h3>
                            <p class="text-sm text-subtle">Alerts help you find rows that match conditions you care about, such as "Role contains nurse" or "Expiry Date is within 30 days". Open the Alerts tab to create and save rules.</p>
                            <p class="text-sm text-subtle">For step-by-step examples, open the full README on GitHub: <a href="${GH_URL}" target="_blank" rel="noopener">Open README on GitHub</a></p>
                        `;
                    }
                })();

                // tab switching logic
                const tabBuilderBtn = howto.querySelector('#howto-tab-builder');
                const tabAlertsBtn = howto.querySelector('#howto-tab-alerts');
                const panelB = howto.querySelector('#howto-panel-builder');
                const panelA = howto.querySelector('#howto-panel-alerts');
                function activateTab(which) {
                    const isBuilder = which === 'builder';
                    tabBuilderBtn.setAttribute('aria-selected', isBuilder ? 'true' : 'false');
                    tabAlertsBtn.setAttribute('aria-selected', isBuilder ? 'false' : 'true');
                    if (isBuilder) { panelB.hidden = false; panelA.hidden = true; panelB.focus(); }
                    else { panelB.hidden = true; panelA.hidden = false; panelA.focus(); }
                }
                tabBuilderBtn.addEventListener('click', () => activateTab('builder'));
                tabAlertsBtn.addEventListener('click', () => activateTab('alerts'));
                // keyboard navigation for tabs
                [tabBuilderBtn, tabAlertsBtn].forEach((btn, idx, arr) => {
                    btn.addEventListener('keydown', (e) => {
                        if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
                            e.preventDefault();
                            const next = arr[(idx + (e.key === 'ArrowRight' ? 1 : arr.length - 1)) % arr.length];
                            next.focus();
                        }
                        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); btn.click(); }
                    });
                });

                // when modal is shown, trap focus; release on close
                const observer = new MutationObserver((records) => {
                    records.forEach(r => {
                        if (r.attributeName === 'class') {
                            const isActive = howto.classList.contains('active');
                            if (isActive) trapFocus(howto);
                            else releaseFocus(howto);
                        }
                    });
                });
                observer.observe(howto, { attributes: true });
            }
            howto.classList.add('active');
        });
    }
}

function renderStep2() {
    const mainContent = document.getElementById('main-content');
    mainContent.innerHTML = `<div id="step2"><div class="flex justify-between items-center"><h2 class="text-2xl">Customize Your Report</h2><button id="reset-btn">Start Over</button></div><div id="customization-container"></div><div class="mt-4 text-center"><button id="generate-pdf-btn" class="btn-primary"><span>Generate PDF Report</span></button></div></div>`;
    document.getElementById('reset-btn').addEventListener('click', () => { if (confirm('Start over?')) { categories.length = 0; renderCategoryUI(); } });
    document.getElementById('generate-pdf-btn').addEventListener('click', () => createPdf());
    // Render customization UI including editable files/notes/images
    processAllDataAndDisplay();
}

document.addEventListener('DOMContentLoaded', init);

// Global convenience: allow closing modals by clicking the backdrop or pressing Escape
document.addEventListener('click', (e) => {
    try {
        const t = e.target;
        if (t && t.classList && t.classList.contains('modal-backdrop')) {
            t.classList.remove('active');
        }
    } catch (err) { /* ignore */ }
});

document.addEventListener('keydown', (e) => {
    try {
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal-backdrop.active').forEach(m => m.classList.remove('active'));
        }
    } catch (err) { /* ignore */ }
});
