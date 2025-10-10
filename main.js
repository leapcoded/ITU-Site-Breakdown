"use strict";

document.addEventListener('DOMContentLoaded', () => {
    
    // --- STATE ---
    // --- IMAGE PREVIEW ---
    // Image preview is now handled per category upload
    let categories = [];
    let libraryFiles = [];
    
    // --- DOM ELEMENTS ---
    const mainContent = document.getElementById('main-content');
    const settingsBtn = document.getElementById('settings-btn');
    const toastContainer = document.getElementById('toast-container');

    // --- TEMPLATES (HTML strings for dynamic content) ---
    const step1HTML = `
        <div id="step1">
            <div id="categories-container" class="space-y-6"></div>
            <div class="text-center mt-6 pt-6 border-t border-base space-x-4">
                <button id="add-category-btn" class="btn-secondary px-5 py-2.5 rounded-lg font-semibold">+ Add Category</button>
                <button id="proceed-btn" class="bg-blue-600 text-white px-8 py-3 rounded-lg font-semibold text-lg hover:bg-blue-700 transition-all shadow-md disabled:bg-slate-300 disabled:cursor-not-allowed" disabled>Customize Report</button>
            </div>
            <div id="error-message" class="hidden mt-4 text-center text-red-600 bg-red-100 p-3 rounded-lg"></div>
        </div>`;

    const step2HTML = `
        <div id="step2" class="fade-in">
             <div class="flex flex-wrap justify-between items-center mb-6 gap-4">
                <div>
                   <h2 class="text-2xl font-bold text-heading">Customize Your Report</h2>
                   <p id="file-summary" class="text-muted text-sm mt-1"></p>
                </div>
                 <button id="reset-btn" class="btn-secondary px-4 py-2 rounded-lg font-medium">Start Over</button>
            </div>
            <div id="customization-container" class="space-y-12"></div>
            <div class="mt-12 text-center border-t border-base pt-8 flex items-center justify-center">
                <button id="generate-pdf-btn" class="bg-blue-600 text-white px-8 py-3 rounded-lg font-semibold text-lg hover:bg-blue-700 transition-all shadow-md disabled:bg-slate-300 disabled:cursor-not-allowed">Generate PDF Report</button>
                <p id="pdf-button-status" class="text-sm text-subtle ml-4"></p>
            </div>
        </div>`;
    
    // --- MODALS ---
    function createModal(id, title, content, footer) {
        const modal = document.createElement('div');
        modal.id = id;
        modal.className = 'modal-backdrop';
        modal.setAttribute('aria-hidden', 'true');
        modal.innerHTML = `
            <div class="modal" role="dialog" aria-modal="true" aria-labelledby="${id}-title">
                <div class="flex justify-between items-center mb-4">
                    <h2 id="${id}-title" class="text-xl font-bold text-heading">${title}</h2>
                    <button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button>
                </div>
                <div class="modal-content">${content}</div>
                <div class="modal-footer mt-6 text-right space-x-2">${footer}</div>
            </div>`;
        document.body.appendChild(modal);
        
        const close = () => {
            modal.classList.remove('active');
            modal.setAttribute('aria-hidden', 'true');
        }

        modal.querySelector('.modal-close-btn').addEventListener('click', close);
        modal.addEventListener('click', (e) => { if(e.target === modal) close(); });
        return { element: modal, open: () => { modal.classList.add('active'); modal.setAttribute('aria-hidden', 'false');}, close };
    }

    const settingsModal = createModal('settings-modal', 'Settings', `
        <div>
            <label for="theme-select" class="block text-sm font-medium text-muted">Theme</label>
            <select id="theme-select" class="w-full mt-1 p-2 rounded-md input-base">
                <option value="light">Light</option>
                <option value="dark">Dark</option>
                <option value="midnight">Midnight</option>
            </select>
        </div>
        <div class="mt-4">
            <h3 class="text-lg font-semibold text-heading">Data Management</h3>
            <button id="clear-draft-btn" class="mt-2 text-sm text-red-600 hover:underline">Clear all saved data (notes, images, files)</button>
        </div>
    `, `<button class="modal-close-action btn-secondary px-4 py-2 rounded-md">Close</button>`);
    settingsModal.element.querySelector('.modal-close-action').addEventListener('click', settingsModal.close);

    const libraryModal = createModal('library-modal', 'File Library', `<div id="library-list" class="max-h-[60vh] overflow-y-auto"></div>`, `
        <button id="library-cancel" class="btn-secondary px-4 py-2 rounded-md">Cancel</button>
        <button id="library-attach" class="bg-blue-600 text-white px-4 py-2 rounded-md">Attach Selected</button>
    `);
    libraryModal.element.querySelector('#library-cancel').addEventListener('click', libraryModal.close);
    
    let captionCallback = null;
    const captionModal = createModal('caption-modal', 'Add Caption/Note', `<textarea id="caption-input" class="w-full p-2 border rounded-md input-base" rows="4"></textarea>`, `
        <button id="caption-cancel" class="btn-secondary px-4 py-2 rounded-md">Cancel</button>
        <button id="caption-save" class="bg-blue-600 text-white px-4 py-2 rounded-md">Save</button>
    `);
    captionModal.element.querySelector('#caption-cancel').addEventListener('click', captionModal.close);
    captionModal.element.querySelector('#caption-save').addEventListener('click', () => {
        if (captionCallback) {
            captionCallback(document.getElementById('caption-input').value);
        }
        captionModal.close();
    });

    function openCaptionModal(initialText = '', callback) {
        document.getElementById('caption-input').value = initialText;
        captionCallback = callback;
        captionModal.open();
    }
    
    // --- DATABASE ---
    const DB_NAME = 'TERRA_DB';
    const DB_VERSION = 1;
    const STORES = { NOTES: 'notes', IMAGES: 'images', FILES: 'files' };
    let db;
    
    async function initDB() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);
            request.onerror = (e) => reject("DB Error: " + e.target.errorCode);
            request.onsuccess = (e) => { db = e.target.result; resolve(db); };
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                Object.values(STORES).forEach(storeName => {
                    if (!db.objectStoreNames.contains(storeName)) {
                        db.createObjectStore(storeName, { keyPath: 'id', autoIncrement: true });
                    }
                });
            };
        });
    }

    async function dbAction(storeName, mode, action, ...args) {
        return new Promise((resolve, reject) => {
            const tx = db.transaction(storeName, mode);
            const store = tx.objectStore(storeName);
            const request = store[action](...args);
            request.onsuccess = (e) => resolve(e.target.result);
            request.onerror = (e) => reject(e.target.error);
        });
    }
    const dbSave = (store, data) => dbAction(store, 'readwrite', data.id ? 'put' : 'add', data);
    const dbLoadAll = (store) => dbAction(store, 'readonly', 'getAll');
    const dbDelete = (store, id) => dbAction(store, 'readwrite', 'delete', id);

    // --- MAIN APP ---
    
    function toast(message, duration = 3000) {
        const toastEl = document.createElement('div');
        toastEl.className = 'toast fade-in';
        toastEl.textContent = message;
        toastContainer.appendChild(toastEl);
        setTimeout(() => toastEl.remove(), duration);
    }

    function renderStep1() {
        mainContent.innerHTML = step1HTML;
        const addBtn = document.getElementById('add-category-btn');
        const procBtn = document.getElementById('proceed-btn');
        addBtn.addEventListener('click', addCategory);
        procBtn.addEventListener('click', renderStep2);
        document.getElementById('categories-container').innerHTML = '';
        if (categories.length === 0) addCategory();
        else categories.forEach(renderCategory);
    }
    
    function renderStep2() {
        mainContent.innerHTML = step2HTML;
        const resetButton = document.getElementById('reset-btn');
        const generateButton = document.getElementById('generate-pdf-btn');
        resetButton.addEventListener('click', () => {
             if(confirm('Are you sure you want to start over? This will clear your current session.')) {
                 categories = [];
                 renderStep1();
             }
        });
        generateButton.addEventListener('click', createPdf);
        processAllDataAndDisplay();
    }

    function addCategory() {
        const categoryId = `cat_${Date.now()}`;
        const category = { id: categoryId, name: `Category ${categories.length + 1}`, files: [], data: [], headers: [], selectedColumns: [], filterRules: [] };
        categories.push(category);
        if (document.getElementById('categories-container')) {
            renderCategory(category);
        }
    }

    // Format date-like values to UK display (DD/MM/YYYY)
    function formatToUK(val) {
        if (val == null) return '';
        if (val instanceof Date) {
            const d = String(val.getDate()).padStart(2,'0');
            const m = String(val.getMonth()+1).padStart(2,'0');
            const y = val.getFullYear();
            return `${d}/${m}/${y}`;
        }
        if (typeof val === 'string') {
            const iso = val.match(/^\s*(\d{4})-(\d{2})-(\d{2})(?:T.*)?\s*$/);
            if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
            const sl = val.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
            if (sl) {
                const day = sl[1].padStart(2,'0');
                const mon = sl[2].padStart(2,'0');
                const year = sl[3].length === 2 ? (parseInt(sl[3],10) > 50 ? '19'+sl[3] : '20'+sl[3]) : sl[3];
                return `${day}/${mon}/${year}`;
            }
            const dash = val.match(/^\s*(\d{1,2})-(\d{1,2})-(\d{2,4})\s*$/);
            if (dash) {
                const day = dash[1].padStart(2,'0');
                const mon = dash[2].padStart(2,'0');
                const year = dash[3].length === 2 ? (parseInt(dash[3],10) > 50 ? '19'+dash[3] : '20'+dash[3]) : dash[3];
                return `${day}/${mon}/${year}`;
            }
            // Try parseable text
            const parsed = Date.parse(val);
            if (!Number.isNaN(parsed)) {
                const dt = new Date(parsed);
                return `${String(dt.getDate()).padStart(2,'0')}/${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
            }
        }
        return String(val);
    }

    function renderCategory(category) {
        const container = document.getElementById('categories-container');
        if (!container) return;

        const categoryEl = document.createElement('div');
        categoryEl.id = category.id;
        categoryEl.className = 'p-4 border rounded-xl bg-surface-alt category-box';
        const fileAcceptTypes = ".csv, text/csv, .xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, image/*";
    categoryEl.innerHTML = `
            <div class="flex justify-between items-center mb-3">
                <h3 class="text-lg font-bold text-heading p-1 -m-1" contenteditable="true">${category.name}</h3>
                <button data-action="remove-category" class="text-slate-400 hover:text-red-500 p-1 text-2xl font-bold leading-none">&times;</button>
            </div>
            <div class="drop-zone border-2 border-dashed rounded-xl p-6 text-center cursor-pointer">
                <p class="font-semibold pointer-events-none text-muted">Drop files, or click to browse</p>
            </div>
            <input type="file" class="hidden" multiple accept="${fileAcceptTypes}">
            <div class="mt-3 flex flex-col gap-2">
                <ul class="file-list text-sm text-muted"></ul>
                <div class="flex-shrink-0">
                    <button class="add-note-btn text-sm btn-tertiary px-2 py-1 rounded-md">+ Note</button>
                    <button class="attach-from-library-btn text-sm btn-tertiary px-2 py-1 rounded-md">Library</button>
                </div>
                <div class="image-preview-container flex flex-col items-center mt-2"></div>
                <div class="image-delete-container flex flex-col items-center mt-2"></div>
            </div>`;
        container.appendChild(categoryEl);
        attachCategoryListeners(categoryEl);
    }
    
    function attachCategoryListeners(categoryEl) {
        const id = categoryEl.id;
        categoryEl.querySelector('.drop-zone').addEventListener('click', () => categoryEl.querySelector('input[type=file]').click());
        categoryEl.querySelector('input[type=file]').addEventListener('change', (e) => handleFiles(e.target.files, id));
        categoryEl.querySelector('.drop-zone').addEventListener('dragover', (e) => { e.preventDefault(); e.currentTarget.classList.add('drag-over'); });
        categoryEl.querySelector('.drop-zone').addEventListener('dragleave', (e) => e.currentTarget.classList.remove('drag-over'));
        categoryEl.querySelector('.drop-zone').addEventListener('drop', (e) => { e.preventDefault(); e.currentTarget.classList.remove('drag-over'); handleFiles(e.dataTransfer.files, id); });
        categoryEl.querySelector('.add-note-btn').addEventListener('click', () => handleAddNote(id));
        categoryEl.querySelector('.attach-from-library-btn').addEventListener('click', () => openLibraryPicker(id));
        categoryEl.querySelector('[data-action="remove-category"]').addEventListener('click', () => {
            if(categories.length > 1) {
                categories = categories.filter(c => c.id !== id);
                categoryEl.remove();
            } else {
                toast("You must have at least one category.");
            }
        });
        categoryEl.querySelector('h3').addEventListener('blur', (e) => {
            const cat = categories.find(c => c.id === id);
            if(cat) cat.name = e.target.textContent;
        });
    }
    
    async function handleFiles(files, categoryId) {
        const category = categories.find(c => c.id === categoryId);
        if (!category) return;
        toast(`Processing ${files.length} file(s)...`);
        let lastImageId = null;
        for (const file of files) {
            try {
                if (file.type.startsWith('image/')) {
                    const dataUrl = await new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onload = e => resolve(e.target.result);
                        reader.onerror = reject;
                        reader.readAsDataURL(file);
                    });
                    const img = { name: file.name, isImage: true, dataUrl, categoryName: category.name };
                    const id = await dbSave(STORES.IMAGES, img);
                    const itemWithId = { ...img, id };
                    libraryFiles.push(itemWithId);
                    rehydrateItem(itemWithId);
                    lastImageId = id;
                    // Show image preview in category
                    const catEl = document.getElementById(category.id);
                    if (catEl) {
                        const previewContainer = catEl.querySelector('.image-preview-container');
                        if (previewContainer) {
                                previewContainer.innerHTML = `<img src="${dataUrl}" alt="${file.name}" class="preview-image" />`;
                        }
                        const deleteContainer = catEl.querySelector('.image-delete-container');
                        if (deleteContainer) {
                            deleteContainer.innerHTML = `<button class="delete-image-btn bg-red-500 text-white px-3 py-1 rounded mt-2" data-image-id="${id}">Delete Image</button>`;
                            deleteContainer.querySelector('.delete-image-btn').addEventListener('click', async function() {
                                await deleteImageFromCategory(category.id, id);
                            });
                        }
                    }
                } else {
                    const data = await parseFile(file);
                    const parsed = { name: file.name, isParsedFile: true, dataRows: data, headers: data.length ? Object.keys(data[0]) : [], categoryName: category.name };
                    const id = await dbSave(STORES.FILES, parsed);
                    const itemWithId = { ...parsed, id };
                    libraryFiles.push(itemWithId);
                    rehydrateItem(itemWithId);
                }
            } catch (e) {
                toast(`Error processing ${file.name}: ${e.message || e}`);
                console.error(e);
            }
        }
        document.getElementById('proceed-btn').disabled = false;
    }

    async function deleteImageFromCategory(categoryId, imageId) {
        // Remove from category.files
        const category = categories.find(c => c.id === categoryId);
        if (!category) return;
        category.files = category.files.filter(f => !(f.isImage && f.id === imageId));
        // Remove from libraryFiles
        const idx = libraryFiles.findIndex(f => f.isImage && f.id === imageId);
        if (idx !== -1) libraryFiles.splice(idx, 1);
        // Remove preview
        const catEl = document.getElementById(categoryId);
        if (catEl) {
            const previewContainer = catEl.querySelector('.image-preview-container');
            if (previewContainer) previewContainer.innerHTML = '';
            const deleteContainer = catEl.querySelector('.image-delete-container');
            if (deleteContainer) deleteContainer.innerHTML = '';
        }
        updateFileList(categoryId);
        toast('Image deleted.');
    }

    async function handleAddNote(categoryId) {
        openCaptionModal('', async (content) => {
            if (!content) return;
            const category = categories.find(c => c.id === categoryId);
            const note = { isNote: true, name: `Note - ${new Date().toLocaleTimeString()}`, content, categoryName: category.name };
            const id = await dbSave(STORES.NOTES, note);
            const itemWithId = { ...note, id };
            libraryFiles.push(itemWithId);
            rehydrateItem(itemWithId);
            document.getElementById('proceed-btn').disabled = false;
        });
    }

    function rehydrateItem(item) {
        let cat = categories.find(c => c.name === item.categoryName);
        if (!cat) {
            addCategory();
            cat = categories[categories.length - 1];
            cat.name = item.categoryName;
            const catEl = document.getElementById(cat.id);
            if(catEl) catEl.querySelector('h3').textContent = cat.name;
        }

        if (cat.files.some(f => f.id === item.id)) return;

        const fileEntry = { id: item.id, name: item.name };
        if (item.isNote) {
            fileEntry.isNote = true;
            fileEntry.content = item.content;
            cat.data.push({ Note: item.content });
            if (!cat.headers.includes('Note')) cat.headers.push('Note');
        } else if (item.isImage) {
            fileEntry.isImage = true;
            fileEntry.dataUrl = item.dataUrl;
            cat.data.push({ Image: item.name });
            if (!cat.headers.includes('Image')) cat.headers.push('Image');
        } else if (item.isParsedFile) {
            fileEntry.isParsedFile = true;
            cat.data.push(...item.dataRows);
            item.headers.forEach(h => {
                if (!cat.headers.includes(h)) cat.headers.push(h);
            });
        }
        cat.files.push(fileEntry);
        if(document.getElementById(cat.id)) {
            updateFileList(cat.id);
        }
    }
    
    function updateFileList(categoryId) {
        const cat = categories.find(c => c.id === categoryId);
        const listEl = document.querySelector(`#${cat.id} .file-list`);
        if (listEl) {
            listEl.innerHTML = cat.files.map(f => `<li class="truncate" title="${f.name}">- ${f.name}</li>`).join('');
        }
    }

    function openLibraryPicker(targetCategoryId) {
        const listEl = document.getElementById('library-list');
        listEl.innerHTML = libraryFiles.map((file, i) => `
            <div class="flex items-center p-2 border-b border-base">
                <input type="checkbox" class="mr-3" data-index="${i}">
                <span class="text-sm">${file.name} (${file.isNote ? 'Note' : file.isImage ? 'Image' : 'File'})</span>
            </div>
        `).join('') || '<p class="text-muted text-sm p-4 text-center">Library is empty.</p>';
        
        const attachBtn = document.getElementById('library-attach');
        
        const handler = () => {
            const selectedIndexes = [...listEl.querySelectorAll('input:checked')].map(cb => parseInt(cb.dataset.index));
            const cat = categories.find(c => c.id === targetCategoryId);
            selectedIndexes.forEach(i => {
                rehydrateItem({ ...libraryFiles[i], categoryName: cat.name });
            });
            libraryModal.close();
            document.getElementById('proceed-btn').disabled = false;
        };

        // Replace button to ensure only one listener is active
        const newAttachBtn = attachBtn.cloneNode(true);
        attachBtn.parentNode.replaceChild(newAttachBtn, attachBtn);
        newAttachBtn.addEventListener('click', handler);

        libraryModal.open();
    }
    
    function parseFile(file) {
        return new Promise((resolve, reject) => {
            if (file.name.endsWith('.csv')) {
                Papa.parse(file, { header: true, skipEmptyLines: true, complete: res => resolve(res.data), error: reject });
            } else if (file.name.endsWith('.xlsx')) {
                const reader = new FileReader();
                reader.onload = e => {
                    const workbook = XLSX.read(e.target.result, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    resolve(XLSX.utils.sheet_to_json(sheet));
                };
                reader.onerror = reject;
                reader.readAsArrayBuffer(file);
            } else {
                reject(new Error('Unsupported file type'));
            }
        });
    }

    function processAllDataAndDisplay() {
        const container = document.getElementById('customization-container');
        if (!container) return;
        container.innerHTML = '';
        categories.forEach(renderCategoryCustomization);
    }

    function renderCategoryCustomization(category) {
        const section = document.createElement('div');
        section.id = `custom-${category.id}`;
        section.className = "category-customization-block border-t pt-8";
        section.innerHTML = `
            <h3 class="text-2xl font-bold text-heading mb-6">${category.name}</h3>
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-8">
                <div>
                    <h4 class="font-semibold text-lg mb-3 pb-2 border-b">Select & Reorder Columns</h4>
                    <div class="columns-container space-y-2 max-h-[400px] overflow-y-auto pr-2"></div>
                </div>
                <div>
                    <h4 class="font-semibold text-lg mb-3 pb-2 border-b">Filter Data</h4>
                    <div class="filter-container p-4 filter-box rounded-lg">
                        <div class="filter-rules-container space-y-3"></div>
                        <button class="add-filter-btn mt-3 text-sm font-semibold text-blue-600 hover:text-blue-800">+ Add Filter Rule</button>
                    </div>
                </div>
            </div>
            <div class="mt-8">
                <div class="flex justify-between items-center mb-3">
                    <h4 class="font-semibold text-lg">Data Preview</h4>
                    <p class="preview-row-count text-sm text-subtle"></p>
                </div>
                <div class="preview-container overflow-auto max-h-[400px] rounded-lg border-base">
                    <table class="w-full text-sm text-left text-base">
                        <thead class="preview-head text-xs uppercase sticky top-0 preview-table-header"></thead>
                        <tbody class="preview-body"></tbody>
                    </table>
                </div>
            </div>
        `;
        document.getElementById('customization-container').appendChild(section);

        populateColumnsSelector(category.id);
        section.querySelector('.add-filter-btn').addEventListener('click', () => addFilterRule(category.id));
        updatePreview(category.id);
    }

    function populateColumnsSelector(categoryId) {
        const cat = categories.find(c => c.id === categoryId);
        const container = document.querySelector(`#custom-${categoryId} .columns-container`);
        cat.selectedColumns = [...cat.headers]; // Default to all selected
        container.innerHTML = cat.headers.map(h => `
            <div class="column-item flex items-center bg-surface p-2 rounded-md border-base border" draggable="true" data-column-name="${h}">
                <svg class="w-5 h-5 text-slate-400 mr-2 drag-handle" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5" /></svg>
                <input type="checkbox" id="col-${categoryId}-${h}" value="${h}" checked class="h-4 w-4 rounded">
                <label for="col-${categoryId}-${h}" class="ml-3 block text-sm">${h}</label>
            </div>
        `).join('');
        
        // Add event listeners for checkboxes and drag/drop
        container.querySelectorAll('input[type=checkbox]').forEach(cb => {
            cb.addEventListener('change', () => {
                cat.selectedColumns = [...container.querySelectorAll('input:checked')].map(c => c.value);
                updatePreview(categoryId);
            });
        });
        // Drag-drop logic would be added here
    }

    function addFilterRule(categoryId) {
        const cat = categories.find(c => c.id === categoryId);
        const container = document.querySelector(`#custom-${categoryId} .filter-rules-container`);
        const ruleId = `rule-${Date.now()}`;
        const rule = { id: ruleId, column: cat.headers[0], operator: 'contains', value: '' };
        cat.filterRules.push(rule);
        
        const ruleEl = document.createElement('div');
        ruleEl.id = ruleId;
        ruleEl.className = 'grid grid-cols-[1fr,auto,1fr,auto] gap-2 items-center';
        ruleEl.innerHTML = `
            <select data-type="column" class="w-full rounded-md input-base text-sm p-2">${cat.headers.map(h => `<option value="${h}">${h}</option>`).join('')}</select>
            <select data-type="operator" class="w-full rounded-md input-base text-sm p-2">
                <option value="contains">contains</option>
                <option value="not_contains">not contains</option>
                <option value="equals">equals</option>
            </select>
            <input type="text" data-type="value" class="w-full rounded-md input-base text-sm p-2">
            <button data-action="remove" class="p-1 text-red-500">&times;</button>
        `;
        container.appendChild(ruleEl);
        
        ruleEl.addEventListener('change', (e) => {
            rule[e.target.dataset.type] = e.target.value;
            updatePreview(categoryId);
        });
        ruleEl.addEventListener('keyup', (e) => {
            if (e.target.dataset.type === 'value') {
                rule.value = e.target.value;
                updatePreview(categoryId);
            }
        });
        ruleEl.querySelector('[data-action=remove]').addEventListener('click', () => {
            cat.filterRules = cat.filterRules.filter(r => r.id !== ruleId);
            ruleEl.remove();
            updatePreview(categoryId);
        });
    }

    function updatePreview(categoryId) {
        const cat = categories.find(c => c.id === categoryId);
        const section = document.querySelector(`#custom-${categoryId}`);
        if(!section) return;

        const filteredData = cat.data.filter(row => {
            return cat.filterRules.every(rule => {
                if (!rule.value) return true;
                const rowVal = String(row[rule.column] || '').toLowerCase();
                const filterVal = rule.value.toLowerCase();
                if (rule.operator === 'contains') return rowVal.includes(filterVal);
                if (rule.operator === 'not_contains') return !rowVal.includes(filterVal);
                if (rule.operator === 'equals') return rowVal === filterVal;
                return true;
            });
        });

        const head = section.querySelector('.preview-head');
        const body = section.querySelector('.preview-body');
        head.innerHTML = `<tr>${cat.selectedColumns.map(h => `<th class="px-6 py-3">${h}</th>`).join('')}</tr>`;
        body.innerHTML = filteredData.slice(0, 50).map(row => `
            <tr class="border-b border-base hover:bg-surface-hover">
                ${cat.selectedColumns.map(h => `<td class="px-6 py-4">${formatToUK(row[h] || '')}</td>`).join('')}
            </tr>
        `).join('');

        section.querySelector('.preview-row-count').textContent = `${filteredData.length} rows`;
    }

    function createPdf() {
        toast('Generating PDF...');
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: 'landscape' });

        let y = 20;
        doc.setFontSize(18);
        doc.text("Report", 14, y);
        y += 10;

        categories.forEach(cat => {
            if (y > 20) y += 10;
            doc.setFontSize(14);
            doc.text(cat.name, 14, y);
            y += 7;

            const filteredData = cat.data.filter(row => cat.filterRules.every(rule => !rule.value || String(row[rule.column] || '').toLowerCase().includes(rule.value.toLowerCase())));
            
            doc.autoTable({
                startY: y,
                head: [cat.selectedColumns],
                body: filteredData.map(row => cat.selectedColumns.map(h => formatToUK(row[h] || ''))),
                theme: 'striped',
                headStyles: { fillColor: [41, 128, 185] }
            });
            y = doc.autoTable.previous.finalY;
        });
        
        doc.save('report.pdf');
    }

    async function init() {
        renderStep1();
        try {
            await initDB();
            const allData = await Promise.all(Object.values(STORES).map(dbLoadAll));
            libraryFiles = allData.flat();
            toast(`Loaded ${libraryFiles.length} item(s) from library.`);
        } catch (e) {
            console.error("Failed to init DB or load data", e);
            toast("Could not load saved data.", 5000);
        }
        
        settingsBtn.addEventListener('click', () => {
             settingsBtn.classList.add("spin");
             setTimeout(() => {
                settingsBtn.classList.remove("spin");
                settingsModal.open();
            }, 220);
        });
        
        const storedTheme = localStorage.getItem('TERRA_THEME') || 'light';
        document.documentElement.setAttribute('data-theme', storedTheme);
        const themeSelectEl = document.getElementById('theme-select');
        if (themeSelectEl) {
            try { themeSelectEl.value = storedTheme; } catch(e){}
            themeSelectEl.addEventListener('change', (e) => {
                const theme = e.target.value;
                document.documentElement.setAttribute('data-theme', theme);
                localStorage.setItem('TERRA_THEME', theme);
                toast(`Theme set to ${theme}`);
            });
        }

        document.getElementById('clear-draft-btn').addEventListener('click', async () => {
            if(confirm('Are you sure you want to delete ALL saved data? This cannot be undone.')) {
                try {
                    await Promise.all(Object.values(STORES).map(store => dbAction(store, 'readwrite', 'clear')));
                    libraryFiles = [];
                    categories = [];
                    renderStep1();
                    toast('All data cleared.');
                    settingsModal.close();
                } catch(e) {
                    toast('Failed to clear data.');
                }
            }
        });
    }
    
    init();
});

