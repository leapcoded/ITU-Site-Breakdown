// Entry point for the app
import { initDB, dbSave, STORES, db, dbClear, dbDelete } from './db.js';
import { categories, addCategory, addLibraryItem, getLibraryItems, clearLibrary, removeLibraryItemById } from './categories.js';
import { addNoteToCategory, getNotes } from './notes.js';
import { renderCategory, processAllDataAndDisplay, createPdf, rehydrateItem, normalizeUKDates, parseDateByLocale } from './ui.js';

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

async function init() {
    // apply saved theme early so styles load correctly
    const savedTheme = localStorage.getItem('TERRA_THEME') || 'light';
    document.documentElement.setAttribute('data-theme', savedTheme);
    await initDB();
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
    const rawItems = allData.flat();
    rawItems.forEach(it => addLibraryItem(it));
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
        const confidence = (ukVotes + usVotes) > 0 ? Math.max(ukVotes, usVotes) / (ukVotes + usVotes) : 0;
        if (ukVotes > usVotes && confidence >= 0.6) return { locale: 'uk', confidence };
        if (usVotes > ukVotes && confidence >= 0.6) return { locale: 'us', confidence };
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
    // Normalize stored parsed files by applying locale-aware parsing to every string cell
    for (const item of libraryFiles) {
        if (item && item.isParsedFile && Array.isArray(item.dataRows)) {
            const fileLocale = item._locale || 'uk';
            const normalized = item.dataRows.map(row => {
                const out = { ...row };
                Object.keys(out).forEach(k => {
                    const v = out[k];
                    if (typeof v === 'string') {
                        const parsed = parseDateByLocale(v, fileLocale);
                        if (parsed) out[k] = parsed;
                    }
                    if (v instanceof Date) out[k] = v.toISOString().split('T')[0];
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
    console.debug('main:init loaded library items:', libraryFiles.length, libraryFiles.map(i=>i.name));
    toast(`Loaded ${libraryFiles.length} item(s) from library.`);
    // Initial UI rendering
    renderCategoryUI();
    setupSettings();
}

function renderCategoryUI() {
    const mainContent = document.getElementById('main-content');
    mainContent.innerHTML = `<div id="step1"><div id="categories-container" class="space-y-6"></div><div class="text-center mt-6 pt-6 border-t border-base space-x-4"><button id="add-category-btn" class="btn-secondary px-5 py-2.5 rounded-lg font-semibold">+ Add Category</button><button id="proceed-btn" class="bg-blue-600 text-white px-8 py-3 rounded-lg font-semibold text-lg hover:bg-blue-700 transition-all shadow-md disabled:bg-slate-300 disabled:cursor-not-allowed" disabled>Customize Report</button></div><div id="error-message" class="hidden mt-4 text-center text-red-600 bg-red-100 p-3 rounded-lg"></div></div>`;
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
                        <label for="theme-select">Theme</label>
                        <select id="theme-select" class="block mt-2 p-2">
                            <option value="light">Light</option>
                            <option value="dark">Dark</option>
                            <option value="midnight">Midnight</option>
                            <option value="teal">Teal</option>
                            <option value="coral">Coral</option>
                            <option value="forest">Forest</option>
                            <option value="sunset">Sunset</option>
                            <option value="lavender">Lavender</option>
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
            // wire theme select change
            const themeSelect = modal.querySelector('#theme-select');
            themeSelect.addEventListener('change', (e) => {
                const t = e.target.value || 'light';
                document.documentElement.setAttribute('data-theme', t);
                localStorage.setItem('TERRA_THEME', t);
                toast(`Theme set to ${t}`);
            });
        }
        // set current value, default to saved or light
        const saved = localStorage.getItem('TERRA_THEME') || 'light';
        const sel = modal.querySelector('#theme-select'); if (sel) sel.value = saved;
        modal.classList.add('active');
    });

    // Wire up the Docs / How-to modal (navbar link)
    const docsLink = document.getElementById('docs-link');
    if (docsLink) {
        docsLink.addEventListener('click', (ev) => {
            ev.preventDefault();
            let howto = document.getElementById('howto-modal');
            if (!howto) {
                howto = document.createElement('div');
                howto.id = 'howto-modal';
                howto.className = 'modal-backdrop';
                howto.innerHTML = `
                    <div class="modal" role="dialog" aria-modal="true">
                        <div class="flex justify-between items-center mb-2">
                            <h2 class="text-lg font-semibold">Read Me — Quick Start</h2>
                            <button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button>
                        </div>
                        <div class="modal-body max-h-[70vh] overflow-y-auto p-2" style="line-height:1.6">
                            <h3 class="text-md font-semibold">Getting started</h3>
                            <ol class="pl-4 mt-2 text-sm text-subtle">
                                <li class="mb-2"><strong>1.</strong> Add a category: Click "+ Add Category" to create a group for related files.</li>
                                <li class="mb-2"><strong>2.</strong> Add files: Click the category area or drag files in. Use CSV or XLSX for tables, and images for photos.</li>
                                <li class="mb-2"><strong>3.</strong> Add notes: Click the <em>+ Note</em> button inside a category to write helpful comments about the files.</li>
                                <li class="mb-2"><strong>4.</strong> Customize your report: Click <em>Customize Report</em> to choose which columns to include and the order they appear in.</li>
                                <li class="mb-2"><strong>5.</strong> Title & logo: Use the project controls to add a report title, subtitle, and upload a small logo if you want it on the PDF.</li>
                                <li class="mb-2"><strong>6.</strong> Create the PDF: Click <em>Generate PDF Report</em> to build and download the report (this may take a moment for many images).</li>
                            </ol>

                            <h3 class="text-md font-semibold mt-4">Helpful tips</h3>
                            <ul class="pl-4 mt-2 text-sm text-subtle">
                                <li class="mb-2">If a date looks wrong, you can change the file's date format in the customization screen.</li>
                                <li class="mb-2">Images are optimized automatically to keep the download size small.</li>
                                <li class="mb-2">Use the Library to reuse items across categories (notes, files, images).</li>
                                <li class="mb-2">If things look unexpected, open <em>Settings</em> and choose <em>Clear saved data</em> to start fresh.</li>
                            </ul>

                            <p class="mt-4 text-sm text-subtle">Need help? Contact the person who shared this app with you or your project administrator for assistance.</p>
                        </div>
                        <div class="modal-footer mt-3 text-right"><button id="howto-close" class="btn-secondary px-3 py-1 rounded">Close</button></div>
                    </div>`;
                document.body.appendChild(howto);
                howto.querySelector('.modal-close-btn').addEventListener('click', () => howto.classList.remove('active'));
                const closeBtn = howto.querySelector('#howto-close'); if (closeBtn) closeBtn.addEventListener('click', () => howto.classList.remove('active'));
                howto.addEventListener('click', (e) => { if (e.target === howto) howto.classList.remove('active'); });
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
