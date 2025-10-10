// UI rendering module
import { categories, addCategory, getCategory, addLibraryItem, getLibraryItems, removeLibraryItemById } from './categories.js';
import { dbSave, dbLoadAll, dbDelete, dbClear, STORES } from './db.js';

// Simple HTML escaper available to all functions in this module
function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (m) { return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]; });
}

// Simple modal factory to centralize modal creation and reuse
let draggingCategoryId = null;
function createSimpleModal(id, title) {
    let modal = document.getElementById(id);
    if (!modal) {
        modal = document.createElement('div');
        modal.id = id;
        modal.className = 'modal-backdrop';
        modal.innerHTML = `
            <div class="modal" role="dialog">
                <div class="flex justify-between items-center mb-2"><h2 class="text-lg font-semibold" id="${id}-title">${title}</h2><button class="modal-close-btn p-1 text-2xl font-bold leading-none text-muted hover:text-heading">&times;</button></div>
                <div class="modal-body" id="${id}-body"></div>
                <div class="modal-footer mt-3 text-right" id="${id}-footer"></div>
            </div>`;
        document.body.appendChild(modal);
        modal.querySelector('.modal-close-btn').addEventListener('click', () => modal.classList.remove('active'));
        modal.addEventListener('click', (e) => { if (e.target === modal) modal.classList.remove('active'); });
    }
    return {
        el: modal,
        setBody: (htmlOrNode) => {
            const body = modal.querySelector(`#${id}-body`);
            if (typeof htmlOrNode === 'string') {
                body.innerHTML = htmlOrNode;
            } else {
                body.innerHTML = '';
                if (Array.isArray(htmlOrNode)) {
                    htmlOrNode.forEach(node => {
                        if (node instanceof Node) body.appendChild(node);
                        else if (typeof node === 'string') body.insertAdjacentHTML('beforeend', node);
                    });
                } else if (htmlOrNode instanceof Node) {
                    body.appendChild(htmlOrNode);
                } else if (htmlOrNode == null) {
                    // nothing to do
                } else {
                    body.textContent = String(htmlOrNode);
                }
            }
        },
        setFooter: (htmlOrNode) => {
            const f = modal.querySelector(`#${id}-footer`);
            if (typeof htmlOrNode === 'string') {
                f.innerHTML = htmlOrNode;
            } else {
                f.innerHTML = '';
                if (Array.isArray(htmlOrNode)) {
                    htmlOrNode.forEach(node => {
                        if (node instanceof Node) f.appendChild(node);
                        else if (typeof node === 'string') f.insertAdjacentHTML('beforeend', node);
                    });
                } else if (htmlOrNode instanceof Node) {
                    f.appendChild(htmlOrNode);
                } else if (htmlOrNode == null) {
                    // nothing
                } else {
                    f.textContent = String(htmlOrNode);
                }
            }
        },
        open: () => modal.classList.add('active'),
        close: () => modal.classList.remove('active'),
        node: modal
    };
}

// Generic toast helper using #toast-container
function showToast(message, duration = 3000) {
    const el = document.getElementById('toast-container');
    if (!el) return;
    const d = document.createElement('div'); d.className = 'toast'; d.textContent = message; el.appendChild(d);
    setTimeout(() => d.remove(), duration);
}

// Resize/compress a DataURL image to a max dimension (preserve aspect ratio).
// If the source is PNG we keep PNG output (to preserve transparency), otherwise output JPEG with quality.
export function resizeImageDataUrl(dataUrl, maxDim = 800, quality = 0.8) {
    return new Promise((resolve, reject) => {
        if (!dataUrl || typeof dataUrl !== 'string') return resolve(dataUrl);
        const mimeMatch = dataUrl.match(/^data:(image\/[a-zA-Z+]+);base64,/);
        const srcMime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
        const img = new Image();
        img.onload = () => {
            let { width, height } = img;
            let targetW = width, targetH = height;
            if (width > maxDim || height > maxDim) {
                const ratio = width / height;
                if (ratio > 1) { targetW = maxDim; targetH = Math.round(maxDim / ratio); }
                else { targetH = maxDim; targetW = Math.round(maxDim * ratio); }
            }
            const canvas = document.createElement('canvas');
            canvas.width = targetW;
            canvas.height = targetH;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, targetW, targetH);
            try {
                if (srcMime === 'image/png') {
                    // Preserve PNG if source was PNG (keeps transparency)
                    resolve(canvas.toDataURL('image/png'));
                } else {
                    resolve(canvas.toDataURL('image/jpeg', quality));
                }
            } catch (err) {
                // Fallback: return original
                resolve(dataUrl);
            }
        };
        img.onerror = (e) => reject(e || new Error('Failed to load image for resizing'));
        img.src = dataUrl;
    });
}

export function renderCategory(category) {
    const container = document.getElementById('categories-container');
    if (!container) return;

    const categoryEl = document.createElement('div');
    categoryEl.id = category.id;
    categoryEl.className = 'p-4 border rounded-xl bg-surface-alt category-box';
    const fileAcceptTypes = ".csv, text/csv, .xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, image/*";
    categoryEl.innerHTML = `
        <div class="flex justify-between items-center mb-3">
            <div class="flex items-center gap-2">
                <span class="drag-handle mr-2 cursor-move" title="Drag to reorder">☰</span>
                <h3 class="text-lg font-bold text-heading p-1 -m-1" contenteditable="true">${category.name}</h3>
            </div>
            <div class="flex items-center gap-2">
                <button data-action="move-up" title="Move up" class="text-slate-500 hover:text-slate-800 p-1">⬆️</button>
                <button data-action="move-down" title="Move down" class="text-slate-500 hover:text-slate-800 p-1">⬇️</button>
                <button data-action="remove-category" class="text-slate-400 hover:text-red-500 p-1 text-2xl font-bold leading-none">&times;</button>
            </div>
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

// Re-render all categories in the container according to the `categories` array order
export function renderAllCategories() {
    const container = document.getElementById('categories-container');
    if (!container) return;
    container.innerHTML = '';
    categories.forEach(cat => renderCategory(cat));
    // refresh file lists for each category so attachments display correctly after re-render
    categories.forEach(cat => { try { updateFileList(cat.id); } catch (_) {} });
    // disable move buttons for top/bottom categories
    const catEls = Array.from(document.querySelectorAll('#categories-container .category-box'));
    catEls.forEach((el, idx) => {
        const up = el.querySelector('[data-action="move-up"]');
        const down = el.querySelector('[data-action="move-down"]');
        if (up) up.disabled = idx === 0;
        if (down) down.disabled = idx === catEls.length - 1;
    });
}

function attachCategoryListeners(categoryEl) {
    const id = categoryEl.id;
    categoryEl.querySelector('.drop-zone').addEventListener('click', () => categoryEl.querySelector('input[type=file]').click());
    categoryEl.querySelector('input[type=file]').addEventListener('change', (e) => handleFiles(e.target.files, id));
    categoryEl.querySelector('.drop-zone').addEventListener('dragover', (e) => { e.preventDefault(); e.currentTarget.classList.add('drag-over'); });
    categoryEl.querySelector('.drop-zone').addEventListener('dragleave', (e) => e.currentTarget.classList.remove('drag-over'));
    categoryEl.querySelector('.drop-zone').addEventListener('drop', (e) => { e.preventDefault(); e.currentTarget.classList.remove('drag-over'); handleFiles(e.dataTransfer.files, id); });
    categoryEl.querySelector('.add-note-btn').addEventListener('click', () => openCaptionModal(id));
    categoryEl.querySelector('.attach-from-library-btn').addEventListener('click', () => openLibraryPicker(id));
    categoryEl.querySelector('[data-action="remove-category"]').addEventListener('click', () => {
        const idx = categories.findIndex(c => c.id === id);
        if (idx !== -1) categories.splice(idx, 1);
        categoryEl.remove();
    });
    categoryEl.querySelector('[data-action="move-up"]').addEventListener('click', () => {
        const idx = categories.findIndex(c => c.id === id);
        if (idx > 0) {
            const tmp = categories[idx - 1];
            categories[idx - 1] = categories[idx];
            categories[idx] = tmp;
            renderAllCategories();
        }
    });
    categoryEl.querySelector('[data-action="move-down"]').addEventListener('click', () => {
        const idx = categories.findIndex(c => c.id === id);
        if (idx !== -1 && idx < categories.length - 1) {
            const tmp = categories[idx + 1];
            categories[idx + 1] = categories[idx];
            categories[idx] = tmp;
            renderAllCategories();
        }
    });
    categoryEl.querySelector('h3').addEventListener('blur', (e) => {
        const cat = categories.find(c => c.id === id);
        if (cat) cat.name = e.target.textContent;
    });

    // Drag & drop reordering using drag-handle only
    const handle = categoryEl.querySelector('.drag-handle');
    if (handle) {
        handle.draggable = true;
        handle.addEventListener('dragstart', (e) => {
            try { e.dataTransfer.setData('text/plain', id); } catch (_) {}
            draggingCategoryId = id;
            categoryEl.classList.add('dragging');
        });
        handle.addEventListener('dragend', () => {
            categoryEl.classList.remove('dragging');
            draggingCategoryId = null;
            document.querySelectorAll('.category-box.drop-target-before, .category-box.drop-target-after').forEach(el => el.classList.remove('drop-target-before','drop-target-after'));
        });
    }
    // allow categories to show drop targets when dragging
    categoryEl.addEventListener('dragover', (e) => { e.preventDefault(); });
    categoryEl.addEventListener('dragenter', (e) => {
        const draggingId = draggingCategoryId || (() => { try { return e.dataTransfer.getData('text/plain'); } catch (_) { return null; } })();
        if (!draggingId || draggingId === id) return;
        const rect = categoryEl.getBoundingClientRect();
        const midY = rect.top + rect.height / 2;
        if (e.clientY < midY) { categoryEl.classList.add('drop-target-before'); categoryEl.classList.remove('drop-target-after'); }
        else { categoryEl.classList.add('drop-target-after'); categoryEl.classList.remove('drop-target-before'); }
    });
    categoryEl.addEventListener('dragleave', () => { categoryEl.classList.remove('drop-target-before','drop-target-after'); });
    categoryEl.addEventListener('drop', (e) => {
        e.preventDefault();
        const fromId = draggingCategoryId || (() => { try { return e.dataTransfer.getData('text/plain'); } catch (_) { return null; } })();
        const toId = id;
        if (!fromId || fromId === toId) return;
        const insertBefore = categoryEl.classList.contains('drop-target-before');
        const fromIdx = categories.findIndex(c => c.id === fromId);
        const toIdx = categories.findIndex(c => c.id === toId);
        if (fromIdx === -1 || toIdx === -1) return;
        const item = categories.splice(fromIdx, 1)[0];
        let targetIdx = toIdx;
        if (fromIdx < toIdx) targetIdx = toIdx - 1;
        if (!insertBefore) targetIdx = targetIdx + 1;
        if (targetIdx < 0) targetIdx = 0;
        if (targetIdx > categories.length) targetIdx = categories.length;
        categories.splice(targetIdx, 0, item);
        renderAllCategories();
    });
}

async function handleFiles(files, categoryId) {
    const category = categories.find(c => c.id === categoryId);
    if (!category) return;
    const toast = (msg) => { const el = document.getElementById('toast-container'); const d = document.createElement('div'); d.className='toast'; d.textContent=msg; el.appendChild(d); setTimeout(()=>d.remove(),3000); };
    toast(`Processing ${files.length} file(s)...`);
    for (const file of files) {
        try {
            if (file.type.startsWith('image/')) {
                let dataUrl = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = e => resolve(e.target.result);
                    reader.onerror = reject;
                    reader.readAsDataURL(file);
                });
                try { dataUrl = await resizeImageDataUrl(dataUrl, 800, 0.8); } catch (e) { console.warn('Image resize failed, using original', e); }
                const img = { name: file.name, isImage: true, dataUrl, categoryName: category.name, categoryId: category.id };
                const id = await dbSave(STORES.IMAGES, img);
                const itemWithId = { ...img, id };
                addLibraryItem(itemWithId);
                rehydrateItem(itemWithId);
                // show preview
                const catEl = document.getElementById(category.id);
                if (catEl) {
                    const previewContainer = catEl.querySelector('.image-preview-container');
                    previewContainer.innerHTML = `<img src="${dataUrl}" alt="${file.name}" class="preview-image" />`;
                    const deleteContainer = catEl.querySelector('.image-delete-container');
                    deleteContainer.innerHTML = `<button class="delete-image-btn bg-red-500 text-white px-3 py-1 rounded mt-2" data-image-id="${id}">Delete Image</button>`;
                    deleteContainer.querySelector('.delete-image-btn').addEventListener('click', async function() {
                        await dbDelete(STORES.IMAGES, id);
                        rehydrateRemove(category.id, id);
                    });
                }
            } else {
                const data = await parseFile(file);
                const parsed = { name: file.name, isParsedFile: true, dataRows: data, headers: data.length ? Object.keys(data[0]) : [], categoryName: category.name, categoryId: category.id };
                const id = await dbSave(STORES.FILES, parsed);
                const itemWithId = { ...parsed, id };
                addLibraryItem(itemWithId);
                rehydrateItem(itemWithId);
            }
        } catch (e) {
            console.error(e);
            const el = document.getElementById('toast-container'); const d = document.createElement('div'); d.className='toast'; d.textContent=`Error processing ${file.name}`; el.appendChild(d); setTimeout(()=>d.remove(),3000);
        }
    }
    document.getElementById('proceed-btn').disabled = false;
}

function rehydrateItem(item) {
    let cat = categories.find(c => c.name === item.categoryName);
    if (!cat) {
        const categoryId = `cat_${Date.now()}`;
        const newCat = { id: categoryId, name: item.categoryName, files: [], data: [], headers: [], selectedColumns: [], filterRules: [] };
        addCategory(newCat);
        cat = newCat;
        renderCategory(cat);
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
        // only include rows that are not hidden
        const visible = (item.dataRows || []).filter(r => !r.__hidden).map(r => ({ ...r, __sourceLocale: item._locale || 'uk', __sourceFileId: item.id }));
        cat.data.push(...visible);
        item.headers.forEach(h => { if (!cat.headers.includes(h)) cat.headers.push(h); });
    }
    cat.files.push(fileEntry);
    if(document.getElementById(cat.id)) updateFileList(cat.id);
}

function rehydrateRemove(categoryId, imageId) {
    const cat = categories.find(c => c.id === categoryId);
    if (!cat) return;
    cat.files = cat.files.filter(f => !(f.isImage && f.id === imageId));
    const catEl = document.getElementById(categoryId);
    if (catEl) {
        const previewContainer = catEl.querySelector('.image-preview-container'); if (previewContainer) previewContainer.innerHTML = '';
        const deleteContainer = catEl.querySelector('.image-delete-container'); if (deleteContainer) deleteContainer.innerHTML = '';
    }
    updateFileList(categoryId);
}

function updateFileList(categoryId) {
    const cat = categories.find(c => c.id === categoryId);
    const listEl = document.querySelector(`#${cat.id} .file-list`);
    if (listEl) {
        listEl.innerHTML = cat.files.map(f => `<li class="truncate" title="${f.name}">- ${f.name}</li>`).join('');
    }
}

// Render the files list and attach handlers for the customization panel for a given categoryId
export function renderFilesListForCategory(categoryId) {
    const category = categories.find(c => c.id === categoryId);
    const filesListEl = document.querySelector(`#custom-${categoryId} .files-list`);
    if (!category || !filesListEl) return;
    filesListEl.innerHTML = (category.files || []).map(f => {
        const full = getLibraryItems().find(i => i.id === f.id) || f;
        const type = full.isNote ? 'Note' : full.isImage ? 'Image' : full.isParsedFile ? 'File' : 'File';
        const overridden = full._locale_overridden ? ' •' : '';
        const localeBadge = full.isParsedFile ? `<button class="locale-btn text-xs px-2 py-1 border rounded mr-3" data-id="${f.id}" title="Detected: ${(full._locale||'uk').toUpperCase()}${overridden} (click to change)">${(full._locale||'uk').toUpperCase()}${overridden}</button>` : '';
        const notePreview = full.isNote ? `<div class="note-preview text-sm text-muted mt-1" style="white-space:pre-wrap;"><button class="view-note-btn text-xs mr-2" data-id="${f.id}">View</button>${escapeHtml(full.content || '')}</div>` : '';
        return `<div class="p-2 border rounded"><div class="flex items-center justify-between"><div class="text-sm truncate">${localeBadge}${full.name} <span class="text-xs text-muted">(${type})</span></div><div class="flex items-center gap-2"><button class="edit-file-btn text-sm px-2 py-1 bg-blue-600 text-white rounded" data-id="${f.id}">Edit</button><button class="remove-file-btn text-sm px-2 py-1 bg-red-500 text-white rounded" data-id="${f.id}">Remove</button></div></div>${notePreview}</div>`;
    }).join('') || '<div class="text-muted text-sm">No files attached.</div>';

    // attach handlers
    filesListEl.querySelectorAll('.edit-file-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            const id = parseInt(e.currentTarget.dataset.id);
            const full = getLibraryItems().find(x => x.id === id);
            if (!full) return;
            if (full.isNote) {
                openCaptionModal(category.id, full.content || '', full.id);
                return;
            } else if (full.isImage) {
                let modal = document.getElementById('image-rename-modal');
                if (!modal) {
                    modal = document.createElement('div');
                    modal.id = 'image-rename-modal';
                    modal.className = 'modal-backdrop';
                    modal.innerHTML = `
                            <div class="modal">
                                <h2 class="text-lg font-semibold mb-2">Rename image</h2>
                                <input id="image-rename-input" class="w-full p-2 border rounded" />
                                <div class="mt-3 text-right"><button id="image-rename-cancel" class="btn-secondary px-3 py-1 rounded">Cancel</button><button id="image-rename-save" class="bg-blue-600 text-white px-3 py-1 rounded">Save</button></div>
                            </div>`;
                    document.body.appendChild(modal);
                    modal.querySelector('#image-rename-cancel').addEventListener('click', () => modal.classList.remove('active'));
                }
                modal.querySelector('#image-rename-input').value = full.name || '';
                modal.classList.add('active');
                modal.querySelector('#image-rename-save').onclick = async () => {
                    const newName = modal.querySelector('#image-rename-input').value;
                    if (!newName) { modal.classList.remove('active'); return; }
                    try {
                        full.name = newName;
                        await dbSave(STORES.IMAGES, { id: full.id, isImage: true, name: newName, dataUrl: full.dataUrl, categoryName: full.categoryName });
                        showToast('Image renamed');
                    } catch (err) { console.error('Failed to rename image', err); showToast('Failed to rename image'); }
                    modal.classList.remove('active');
                    renderFilesListForCategory(category.id);
                    updatePreview(category.id);
                };
                return;
            } else if (full.isParsedFile) {
                openParsedEditor(full, category);
                return;
            }
            renderFilesListForCategory(category.id);
            updatePreview(category.id);
        });
    });

    filesListEl.querySelectorAll('.remove-file-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            const id = parseInt(e.currentTarget.dataset.id);
            category.files = (category.files || []).filter(x => x.id !== id);
            removeLibraryItemById(id);
            await dbDelete(STORES.NOTES, id).catch(()=>{});
            await dbDelete(STORES.IMAGES, id).catch(()=>{});
            await dbDelete(STORES.FILES, id).catch(()=>{});
            rebuildCategoryData(category);
            renderFilesListForCategory(category.id);
            updatePreview(category.id);
        });
    });

    // locale badge
    filesListEl.querySelectorAll('.locale-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = parseInt(e.currentTarget.dataset.id, 10);
            const current = (getLibraryItems().find(i => i.id === id) || {})._locale || 'uk';
            const sel = document.createElement('select');
            sel.className = 'text-xs p-1 rounded border';
            ['uk','us','auto'].forEach(opt => {
                const o = document.createElement('option'); o.value = opt; o.textContent = opt.toUpperCase(); if (opt === current) o.selected = true; sel.appendChild(o);
            });
            e.currentTarget.replaceWith(sel);
            sel.focus();
            const finish = async (newVal) => {
                try {
                    const lib = getLibraryItems().find(i => i.id === id);
                    if (!lib) return;
                    lib._locale = newVal;
                    await dbSave(STORES.FILES, { ...lib, _locale: lib._locale });
                    rebuildCategoryData(category);
                    updatePreview(category.id);
                    renderFilesListForCategory(category.id);
                } catch (err) {
                    console.error('Failed to update locale', err);
                    renderFilesListForCategory(category.id);
                }
            };
            sel.addEventListener('change', (ev) => finish(ev.target.value));
            sel.addEventListener('blur', () => finish(sel.value));
        });
    });

    filesListEl.querySelectorAll('.view-note-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const id = parseInt(e.currentTarget.dataset.id, 10);
            const lib = getLibraryItems().find(i => i.id === id);
            const modal = createSimpleModal('view-note-modal', lib ? lib.name : 'Note');
            const pre = document.createElement('div'); pre.style.whiteSpace = 'pre-wrap'; pre.className = 'p-2'; pre.textContent = lib ? lib.content : '';
            modal.setBody(pre);
            modal.setFooter(`<button class="btn-secondary px-3 py-1 rounded" onclick="document.getElementById('view-note-modal').classList.remove('active')">Close</button>`);
            modal.open();
        });
    });
}

function openLibraryPicker(targetCategoryId) {
    // Ensure modal exists (or create it)
    let modalEl = document.getElementById('library-modal');
    if (!modalEl) {
        const modal = document.createElement('div');
        modal.id = 'library-modal';
        modal.className = 'modal-backdrop';
        modal.innerHTML = `
            <div class="modal" role="dialog">
                <h2 class="text-xl font-bold mb-2">File Library</h2>
                <div id="library-list" class="max-h-[60vh] overflow-y-auto mb-4"></div>
                <div class="modal-footer text-right"><button id="library-cancel" class="btn-secondary px-4 py-2 rounded-md">Cancel</button><button id="library-attach" class="bg-blue-600 text-white px-4 py-2 rounded-md">Attach Selected</button></div>
            </div>`;
        document.body.appendChild(modal);
        modal.querySelector('#library-cancel').addEventListener('click', () => modal.classList.remove('active'));
        modalEl = modal;
    }

    const listEl = document.getElementById('library-list');
    const items = getLibraryItems();
    console.debug('openLibraryPicker targetCategoryId=', targetCategoryId, 'libraryCount=', (items && items.length) || 0);
    // render either items or an empty message
    if (items && items.length) {
        listEl.innerHTML = items.map((file, i) => `
            <div class="flex items-center p-2 border-b border-base">
                <input type="checkbox" class="mr-3" data-index="${i}">
                <span class="text-sm">${file.name} (${file.isNote ? 'Note' : file.isImage ? 'Image' : 'File'})</span>
            </div>
        `).join('');
    } else {
        listEl.innerHTML = '<p class="text-muted text-sm p-4 text-center">Library is empty.</p>';
    }

    // Attach handler only if attach button exists (it won't when library is empty)
    const attachBtn = document.getElementById('library-attach');
    if (attachBtn && items && items.length) {
        // ensure we don't double-bind handlers: replace with a fresh node
        const handler = () => {
            const selectedIndexes = [...listEl.querySelectorAll('input:checked')].map(cb => parseInt(cb.dataset.index));
            const cat = categories.find(c => c.id === targetCategoryId);
            selectedIndexes.forEach(i => { rehydrateItem({ ...items[i], categoryName: cat.name }); });
            const m = document.getElementById('library-modal'); if (m) m.classList.remove('active');
            const proceed = document.getElementById('proceed-btn'); if (proceed) proceed.disabled = false;
        };
        const newAttachBtn = attachBtn.cloneNode(true);
        attachBtn.parentNode.replaceChild(newAttachBtn, attachBtn);
        newAttachBtn.addEventListener('click', handler);
    }

    // show modal
    modalEl.classList.add('active');
}

function openCaptionModal(categoryId, initialText = '', existingNoteId = null) {
    const modal = createSimpleModal('caption-modal', existingNoteId ? 'Edit note' : 'Add a note');
    const textarea = document.createElement('textarea'); textarea.id = 'caption-textarea'; textarea.className = 'w-full p-2 border rounded-md'; textarea.rows = 6; textarea.value = initialText || '';
    modal.setBody(textarea);
    const saveBtn = document.createElement('button'); saveBtn.className = 'bg-blue-600 text-white px-3 py-1 rounded'; saveBtn.textContent = 'Save';
    const cancelBtn = document.createElement('button'); cancelBtn.className = 'btn-secondary px-3 py-1 rounded mr-2'; cancelBtn.textContent = 'Cancel';
    cancelBtn.addEventListener('click', () => modal.close());
    saveBtn.addEventListener('click', async () => {
        const content = textarea.value;
        if (!content) { modal.close(); return; }
        const cat = categories.find(c => c.id === categoryId);
        const note = { isNote: true, name: existingNoteId ? undefined : `Note - ${new Date().toLocaleTimeString()}`, content, categoryName: cat ? cat.name : 'Uncategorized' };
        try {
            if (existingNoteId) {
                // update
                await dbSave(STORES.NOTES, { id: existingNoteId, isNote: true, name: note.name || undefined, content, categoryName: note.categoryName });
                const lib = getLibraryItems().find(i => i.id === existingNoteId);
                if (lib) { lib.content = content; }
            } else {
                const id = await dbSave(STORES.NOTES, note);
                const itemWithId = { ...note, id };
                addLibraryItem(itemWithId);
                rehydrateItem(itemWithId);
            }
            const proceedBtn = document.getElementById('proceed-btn'); if (proceedBtn) proceedBtn.disabled = false;
        } catch (err) { console.error('Failed to save note', err); }
    modal.close();
    renderFilesListForCategory(categoryId);
        updatePreview(categoryId);
    });
    modal.setFooter([cancelBtn, saveBtn]);
    modal.open();
}

function parseFile(file) {
    return new Promise((resolve, reject) => {
        if (file.name.endsWith('.csv')) {
            window.Papa.parse(file, { header: true, skipEmptyLines: true, complete: res => {
                const data = normalizeUKDates(res.data);
                resolve(data);
            }, error: reject });
        } else if (file.name.endsWith('.xlsx')) {
            const reader = new FileReader();
            reader.onload = e => {
                const workbook = window.XLSX.read(e.target.result, { type: 'array', cellDates: true });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                // request formatted output then normalize dates
                const raw = window.XLSX.utils.sheet_to_json(sheet, { raw: false, dateNF: 'yyyy-mm-dd' });
                const data = normalizeUKDates(raw);
                resolve(data);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        } else {
            reject(new Error('Unsupported file type'));
        }
    });
}

// Convert UK-style date strings (DD/MM/YYYY, DD-MM-YYYY, D MMM YYYY) and Date objects to ISO YYYY-MM-DD strings
function normalizeUKDates(rows) {
    if (!Array.isArray(rows)) return rows;
    const monthNames = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
    const ukRegex = /^\s*(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\s*$/;
    return rows.map(row => {
        const out = { ...row };
        Object.keys(out).forEach(k => {
            const v = out[k];
            if (v == null) return;
            if (v instanceof Date) {
                const d = v.getDate().toString().padStart(2,'0');
                const m = (v.getMonth()+1).toString().padStart(2,'0');
                const y = v.getFullYear();
                out[k] = `${d}/${m}/${y}`;
                return;
            }
            if (typeof v === 'string') {
                const m = v.match(ukRegex);
                if (m) {
                    let day = parseInt(m[1], 10);
                    let month = parseInt(m[2], 10);
                    let year = parseInt(m[3], 10);
                    if (year < 100) year += year >= 50 ? 1900 : 2000;
                    const dateObj = new Date(year, month - 1, day);
                    if (!isNaN(dateObj.getTime()) && dateObj.getFullYear() === year && dateObj.getMonth() === month - 1 && dateObj.getDate() === day) {
                        const dd = String(day).padStart(2,'0');
                        const mm = String(month).padStart(2,'0');
                        const yyyy = String(year).padStart(4,'0');
                        out[k] = `${dd}/${mm}/${yyyy}`;
                        return;
                    }
                }
                // Try formats like '1 Jan 2025' or '01 January 2025'
                const m2 = v.match(/^\s*(\d{1,2})\s+([A-Za-z]{3,})\.?,?\s+(\d{4})\s*$/);
                if (m2) {
                    const day = parseInt(m2[1], 10);
                    const monthName = m2[2].toLowerCase().slice(0,3);
                    const year = parseInt(m2[3], 10);
                    const mi = monthNames.indexOf(monthName);
                    if (mi !== -1) {
                        const dateObj = new Date(year, mi, day);
                        if (!isNaN(dateObj.getTime())) {
                            const dd = String(day).padStart(2,'0');
                            const mm = String(mi+1).padStart(2,'0');
                            const yyyy = String(year).padStart(4,'0');
                            out[k] = `${dd}/${mm}/${yyyy}`;
                        }
                    }
                }
            }
        });
        return out;
    });
}
// export helper so stored files can be normalized on load
export { normalizeUKDates };

// Export small helpers used by other modules
export { handleFiles, rehydrateItem };

// --- Customization / Report generation ---
export function processAllDataAndDisplay() {
    const container = document.getElementById('customization-container');
    if (!container) return;
    container.innerHTML = '';
    // project-level controls: report title, subtitle and logo
    const logoControls = document.createElement('div');
    logoControls.id = 'logo-controls';
    logoControls.className = 'mb-4 flex flex-col gap-3';
    const savedTitle = localStorage.getItem('APP_REPORT_TITLE') || 'Report';
    const savedSubtitle = localStorage.getItem('APP_REPORT_SUBTITLE') || '';
    logoControls.innerHTML = `
        <div class="flex items-center gap-4">
            <div id="logo-preview" class="flex items-center"></div>
            <input id="logo-input" type="file" accept="image/*" class="hidden" />
            <div class="flex flex-col gap-2">
                <div class="flex gap-2 items-center">
                    <button id="logo-upload" class="btn-tertiary px-3 py-1 rounded">Upload logo</button>
                    <button id="logo-remove" class="btn-secondary px-3 py-1 rounded">Remove logo</button>
                </div>
                <div class="text-sm text-muted">Logo will appear on the generated PDF</div>
            </div>
        </div>
        <div class="flex gap-4 items-center mt-2">
            <div class="flex-1">
                <label class="text-sm text-subtle">Report title</label>
                <input id="report-title-input" class="w-full p-2 rounded border" value="${escapeHtml(savedTitle)}" />
            </div>
            <div class="flex-1">
                <label class="text-sm text-subtle">Report subtitle</label>
                <input id="report-subtitle-input" class="w-full p-2 rounded border" value="${escapeHtml(savedSubtitle)}" />
            </div>
        </div>
    `;
    container.appendChild(logoControls);
    // hookup
    const logoPreviewEl = document.getElementById('logo-preview');
    const logoInput = document.getElementById('logo-input');
    const logoUploadBtn = document.getElementById('logo-upload');
    const logoRemoveBtn = document.getElementById('logo-remove');
    function updateLogoPreview() {
        const data = localStorage.getItem('APP_LOGO_DATAURL');
        if (data) {
            logoPreviewEl.innerHTML = `<img src="${data}" alt="logo" style="height:48px; object-fit:contain;"/>`;
            logoRemoveBtn.style.display = '';
        } else {
            logoPreviewEl.innerHTML = '<div class="text-sm text-muted">No logo set</div>';
            logoRemoveBtn.style.display = 'none';
        }
    }
    logoUploadBtn.addEventListener('click', () => logoInput.click());
    logoInput.addEventListener('change', async (e) => {
        const f = e.target.files && e.target.files[0];
        if (!f) return;
        try {
            const dataUrl = await new Promise((res, rej) => {
                const r = new FileReader(); r.onload = (ev) => res(ev.target.result); r.onerror = rej; r.readAsDataURL(f);
            });
            let resized = dataUrl;
            try { resized = await resizeImageDataUrl(dataUrl, 200, 0.8); } catch (err) { console.warn('Logo resize failed', err); }
            localStorage.setItem('APP_LOGO_DATAURL', resized);
            updateLogoPreview();
            showToast('Logo uploaded');
        } catch (err) { console.error('Logo upload failed', err); showToast('Failed to upload logo'); }
    });
    logoRemoveBtn.addEventListener('click', () => { localStorage.removeItem('APP_LOGO_DATAURL'); updateLogoPreview(); showToast('Logo removed'); });
    updateLogoPreview();
    // wiring for title/subtitle inputs
    const titleInput = document.getElementById('report-title-input');
    const subtitleInput = document.getElementById('report-subtitle-input');
    if (titleInput) {
        titleInput.addEventListener('blur', (e) => {
            localStorage.setItem('APP_REPORT_TITLE', e.target.value || 'Report');
            showToast('Report title saved');
        });
    }
    if (subtitleInput) {
        subtitleInput.addEventListener('blur', (e) => {
            localStorage.setItem('APP_REPORT_SUBTITLE', e.target.value || '');
            showToast('Report subtitle saved');
        });
    }
    categories.forEach(renderCategoryCustomization);
}

export function renderCategoryCustomization(category) {
    const section = document.createElement('div');
    section.id = `custom-${category.id}`;
    section.className = 'category-customization-block border-t pt-8';
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

    // Files list (notes, images, parsed files) with edit/remove controls
    const filesBlock = document.createElement('div');
    filesBlock.className = 'mt-6';
    filesBlock.innerHTML = `<h4 class="font-semibold mb-2">Files & Notes</h4><div class="files-list space-y-2"></div>`;
    section.appendChild(filesBlock);
    const filesListEl = filesBlock.querySelector('.files-list');
    function renderFilesList() {
        filesListEl.innerHTML = (category.files || []).map(f => {
            const full = getLibraryItems().find(i => i.id === f.id) || f;
            const type = full.isNote ? 'Note' : full.isImage ? 'Image' : full.isParsedFile ? 'File' : 'File';
            // show locale badge for parsed files so users can quickly see/change detection
            const overridden = full._locale_overridden ? ' •' : '';
            const localeBadge = full.isParsedFile ? `<button class="locale-btn text-xs px-2 py-1 border rounded mr-3" data-id="${f.id}" title="Detected: ${(full._locale||'uk').toUpperCase()}${overridden} (click to change)">${(full._locale||'uk').toUpperCase()}${overridden}</button>` : '';
            const notePreview = full.isNote ? `<div class="note-preview text-sm text-muted mt-1" style="white-space:pre-wrap;"><button class="view-note-btn text-xs mr-2" data-id="${f.id}">View</button>${escapeHtml(full.content || '')}</div>` : '';
            return `<div class="p-2 border rounded"><div class="flex items-center justify-between"><div class="text-sm truncate">${localeBadge}${full.name} <span class="text-xs text-muted">(${type})</span></div><div class="flex items-center gap-2"><button class="edit-file-btn text-sm px-2 py-1 bg-blue-600 text-white rounded" data-id="${f.id}">Edit</button><button class="remove-file-btn text-sm px-2 py-1 bg-red-500 text-white rounded" data-id="${f.id}">Remove</button></div></div>${notePreview}</div>`;
        }).join('') || '<div class="text-muted text-sm">No files attached.</div>';
        // attach handlers
        filesListEl.querySelectorAll('.edit-file-btn').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const id = parseInt(e.currentTarget.dataset.id);
                // use full library item for edits
                const full = getLibraryItems().find(x => x.id === id);
                if (!full) return;
                if (full.isNote) {
                    // open caption modal pre-filled for editing existing note
                    openCaptionModal(category.id, full.content || '', full.id);
                    return;
                } else if (full.isImage) {
                    // inline image rename modal
                    let modal = document.getElementById('image-rename-modal');
                    if (!modal) {
                        modal = document.createElement('div');
                        modal.id = 'image-rename-modal';
                        modal.className = 'modal-backdrop';
                        modal.innerHTML = `
                            <div class="modal">
                                <h2 class="text-lg font-semibold mb-2">Rename image</h2>
                                <input id="image-rename-input" class="w-full p-2 border rounded" />
                                <div class="mt-3 text-right"><button id="image-rename-cancel" class="btn-secondary px-3 py-1 rounded">Cancel</button><button id="image-rename-save" class="bg-blue-600 text-white px-3 py-1 rounded">Save</button></div>
                            </div>`;
                        document.body.appendChild(modal);
                        modal.querySelector('#image-rename-cancel').addEventListener('click', () => modal.classList.remove('active'));
                    }
                    modal.querySelector('#image-rename-input').value = full.name || '';
                    modal.classList.add('active');
                    modal.querySelector('#image-rename-save').onclick = async () => {
                        const newName = modal.querySelector('#image-rename-input').value;
                        if (!newName) { modal.classList.remove('active'); return; }
                        try {
                            full.name = newName;
                            await dbSave(STORES.IMAGES, { id: full.id, isImage: true, name: newName, dataUrl: full.dataUrl, categoryName: full.categoryName });
                            showToast('Image renamed');
                        } catch (err) { console.error('Failed to rename image', err); showToast('Failed to rename image'); }
                        modal.classList.remove('active');
                        renderFilesListForCategory(category.id);
                        updatePreview(category.id);
                    };
                    return;
                } else if (full.isParsedFile) {
                    // open modal table editor
                    openParsedEditor(full, category);
                    return; // modal will handle save and re-render
                }
                renderFilesListForCategory(category.id);
                updatePreview(category.id);
            });
        });
        filesListEl.querySelectorAll('.remove-file-btn').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const id = parseInt(e.currentTarget.dataset.id);
                // remove from category.files
                category.files = (category.files || []).filter(x => x.id !== id);
                // remove from library and DB
                removeLibraryItemById(id);
                await dbDelete(STORES.NOTES, id).catch(()=>{});
                await dbDelete(STORES.IMAGES, id).catch(()=>{});
                await dbDelete(STORES.FILES, id).catch(()=>{});
            // rebuild category.data from remaining files
        rebuildCategoryData(category);
        renderFilesListForCategory(category.id);
        updatePreview(category.id);
            });
        });
        // locale badge click handler: inline select to change per-file locale
        filesListEl.querySelectorAll('.locale-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = parseInt(e.currentTarget.dataset.id, 10);
                const current = (getLibraryItems().find(i => i.id === id) || {})._locale || 'uk';
                // create inline select
                const sel = document.createElement('select');
                sel.className = 'text-xs p-1 rounded border';
                ['uk','us','auto'].forEach(opt => {
                    const o = document.createElement('option'); o.value = opt; o.textContent = opt.toUpperCase();
                    if (opt === current) o.selected = true;
                    sel.appendChild(o);
                });
                // replace button with select
                e.currentTarget.replaceWith(sel);
                sel.focus();
                const finish = async (newVal) => {
                    try {
                        const lib = getLibraryItems().find(i => i.id === id);
                        if (!lib) return;
                        lib._locale = newVal;
                        await dbSave(STORES.FILES, { ...lib, _locale: lib._locale });
                        // rebuild and refresh UI
                        rebuildCategoryData(category);
                        updatePreview(category.id);
                        renderFilesListForCategory(category.id);
                    } catch (err) {
                        console.error('Failed to update locale', err);
                        // restore original button on failure
                        renderFilesList();
                    }
                };
                sel.addEventListener('change', (ev) => finish(ev.target.value));
                sel.addEventListener('blur', () => finish(sel.value));
            });
        });
        // view-note buttons
        filesListEl.querySelectorAll('.view-note-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = parseInt(e.currentTarget.dataset.id, 10);
                const lib = getLibraryItems().find(i => i.id === id);
                const modal = createSimpleModal('view-note-modal', lib ? lib.name : 'Note');
                const pre = document.createElement('div'); pre.style.whiteSpace = 'pre-wrap'; pre.className = 'p-2'; pre.textContent = lib ? lib.content : '';
                modal.setBody(pre);
                modal.setFooter(`<button class="btn-secondary px-3 py-1 rounded" onclick="document.getElementById('view-note-modal').classList.remove('active')">Close</button>`);
                modal.open();
            });
        });
    }
    renderFilesList();

    // Rebuild category.data by aggregating all parsed file rows currently attached to the category
    function rebuildCategoryData(category) {
        category.data = [];
        (category.files || []).forEach(fileRef => {
            const lib = getLibraryItems().find(i => i.id === fileRef.id);
            if (lib && lib.isParsedFile && Array.isArray(lib.dataRows)) {
                const visible = lib.dataRows.filter(r => !r.__hidden).map(r => ({ ...r, __sourceLocale: lib._locale || 'uk', __sourceFileId: lib.id }));
                category.data.push(...visible);
            }
        });
    }

    // Modal table editor for parsed files
    function openParsedEditor(parsedFile, category) {
        // parsedFile: object from library with { id, name, dataRows, headers }
        const modal = document.createElement('div');
        modal.className = 'fixed inset-0 bg-black bg-opacity-50 flex items-start justify-center z-50 p-4';
        modal.innerHTML = `
            <div class="bg-white rounded shadow-lg max-w-4xl w-full max-h-[80vh] overflow-auto p-4">
                <div class="flex items-center justify-between mb-3">
                    <div class="flex items-center gap-3">
                      <h3 class="text-lg font-semibold">Edit ${escapeHtml(parsedFile.name)}</h3>
                      <label class="text-sm text-subtle">Locale:</label>
                      <select id="parsed-locale" class="text-sm p-1 rounded border">
                        <option value="uk">UK (DD/MM/YYYY)</option>
                        <option value="us">US (MM/DD/YYYY)</option>
                        <option value="auto">Auto (try UK)</option>
                      </select>
                    </div>
                    <div class="flex gap-2">
                        <button id="parsed-cancel" class="px-3 py-1 border rounded">Cancel</button>
                        <button id="parsed-save" class="px-3 py-1 bg-blue-600 text-white rounded">Save</button>
                    </div>
                </div>
                <div class="overflow-auto">
                    <table id="parsed-table" class="min-w-full border-collapse">
                        <thead></thead>
                        <tbody></tbody>
                    </table>
                </div>
            </div>
        `;
        document.body.appendChild(modal);

        const thead = modal.querySelector('#parsed-table thead');
        const tbody = modal.querySelector('#parsed-table tbody');

        const headers = parsedFile.headers && parsedFile.headers.length ? parsedFile.headers : (parsedFile.dataRows && parsedFile.dataRows.length ? Object.keys(parsedFile.dataRows[0]) : []);

    // render header (add Hide column)
        const trh = document.createElement('tr');
        const thHide = document.createElement('th');
        thHide.className = 'border px-2 py-1 bg-gray-100 text-center text-xs';
        thHide.textContent = 'Hide';
        trh.appendChild(thHide);
        headers.forEach(h => {
            const th = document.createElement('th');
            th.className = 'border px-2 py-1 text-left bg-gray-100';
            th.textContent = h;
            trh.appendChild(th);
        });
        thead.appendChild(trh);

    // set locale select to file preference if present
    const localeSelect = modal.querySelector('#parsed-locale');
    if (parsedFile._locale) localeSelect.value = parsedFile._locale;

        // render rows
        const rows = parsedFile.dataRows || [];
        rows.forEach((row, rIdx) => {
            const tr = document.createElement('tr');
            // hide checkbox
            const tdHide = document.createElement('td');
            tdHide.className = 'border px-2 py-1 text-center';
            const hideCb = document.createElement('input');
            hideCb.type = 'checkbox';
            hideCb.className = 'row-hide-checkbox';
            hideCb.dataset.row = rIdx;
            hideCb.checked = !!row.__hidden;
            tdHide.appendChild(hideCb);
            tr.appendChild(tdHide);
            headers.forEach(h => {
                const td = document.createElement('td');
                td.className = 'border px-2 py-1';
                const input = document.createElement('input');
                input.value = row[h] != null ? row[h] : '';
                input.className = 'w-full border-0 outline-none';
                input.dataset.col = h;
                input.dataset.row = rIdx;
                td.appendChild(input);
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        // allow adding/removing rows quickly
        const addRowBtn = document.createElement('button');
        addRowBtn.textContent = 'Add row';
        addRowBtn.className = 'mt-2 px-2 py-1 border rounded';
        addRowBtn.addEventListener('click', () => {
            const newRow = {};
            headers.forEach(h => newRow[h] = '');
            rows.push(newRow);
            // append to table
            const tr = document.createElement('tr');
            headers.forEach(h => {
                const td = document.createElement('td');
                td.className = 'border px-2 py-1';
                const input = document.createElement('input');
                input.value = '';
                input.className = 'w-full border-0 outline-none';
                input.dataset.col = h;
                input.dataset.row = rows.length - 1;
                td.appendChild(input);
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        modal.querySelector('.bg-white').appendChild(addRowBtn);

        modal.querySelector('#parsed-cancel').addEventListener('click', () => {
            document.body.removeChild(modal);
        });

        modal.querySelector('#parsed-save').addEventListener('click', async () => {
            // read table into dataRows and include hide checkbox state
            const inputs = Array.from(tbody.querySelectorAll('input')).filter(i => !i.classList.contains('row-hide-checkbox'));
            const newRows = [];
            const grouped = {};
            inputs.forEach(inp => {
                const r = parseInt(inp.dataset.row, 10);
                const c = inp.dataset.col;
                grouped[r] = grouped[r] || {};
                grouped[r][c] = inp.value;
            });
            // attach hide flags
            const hideCbs = Array.from(tbody.querySelectorAll('.row-hide-checkbox'));
            hideCbs.forEach(cb => {
                const r = parseInt(cb.dataset.row, 10);
                grouped[r] = grouped[r] || {};
                if (cb.checked) grouped[r].__hidden = true;
            });
            // interpret ambiguous date strings according to chosen locale
            const chosenLocale = (modal.querySelector('#parsed-locale') || { value: 'uk' }).value || 'uk';
            Object.keys(grouped).forEach(rk => {
                const rowObj = grouped[rk];
                Object.keys(rowObj).forEach(col => {
                    const val = rowObj[col];
                    if (typeof val === 'string') {
                        const parsedIso = parseDateByLocale(val, chosenLocale);
                        // debug: show parsing step for ambiguous values
                        if (parsedIso) {
                            console.debug('parseDateByLocale =>', { col, val, chosenLocale, parsedIso });
                            rowObj[col] = parsedIso;
                        } else {
                            // also log when value looks like a slashed date but wasn't parsed
                            if (/^\s*\d{1,2}\/\d{1,2}\/\d{2,4}\s*$/.test(val)) console.debug('parseDateByLocale failed to parse slashed date', { col, val, chosenLocale });
                        }
                    }
                });
            });
            Object.keys(grouped).sort((a,b)=>a-b).forEach(k => newRows.push(grouped[k]));

            // save back to library and DB
            parsedFile.dataRows = newRows;
            // persist chosen locale
            parsedFile._locale = (modal.querySelector('#parsed-locale') || { value: 'uk' }).value || 'uk';
            parsedFile.headers = headers;
            try {
                await dbSave(STORES.FILES, { id: parsedFile.id, name: parsedFile.name, isParsedFile: true, dataRows: newRows, headers: headers, categoryName: parsedFile.categoryName });
            } catch (err) {
                console.error('Failed saving parsed file', err);
                const el = document.getElementById('toast-container'); if (el) { const d = document.createElement('div'); d.className='toast'; d.textContent='Failed to save file changes.'; el.appendChild(d); setTimeout(()=>d.remove(),3000); }
                return; // don't close modal on failure
            }

            // rebuild category data and update preview
            rebuildCategoryData(category);
            updatePreview(category.id);

            // update any in-memory category references that used this file's rows
            document.body.removeChild(modal);
            // re-render files list to reflect any changes
            renderFilesListForCategory(category.id);
        });
    }

    function escapeHtml(s) {
        return String(s).replace(/[&<>"]/g, function (m) { return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]; });
    }

    populateColumnsSelector(category.id);
    section.querySelector('.add-filter-btn').addEventListener('click', () => addFilterRule(category.id));
    updatePreview(category.id);
}

export function populateColumnsSelector(categoryId) {
    const cat = categories.find(c => c.id === categoryId);
    const container = document.querySelector(`#custom-${categoryId} .columns-container`);
    if (!cat || !container) return;
    cat.selectedColumns = [...cat.headers];
    container.innerHTML = cat.headers.map(h => `
        <div class="column-item flex items-center bg-surface p-2 rounded-md border-base border" draggable="true" data-column-name="${h}">
            <svg class="w-5 h-5 text-slate-400 mr-2 drag-handle" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" d="M3.75 6.75h16.5M3.75 12h16.5m-16.5 5.25h16.5" /></svg>
            <input type="checkbox" id="col-${categoryId}-${h}" value="${h}" checked class="h-4 w-4 rounded">
            <label for="col-${categoryId}-${h}" class="ml-3 block text-sm">${h}</label>
        </div>
    `).join('');
    container.querySelectorAll('input[type=checkbox]').forEach(cb => {
        cb.addEventListener('change', () => {
            cat.selectedColumns = [...container.querySelectorAll('input:checked')].map(c => c.value);
            updatePreview(categoryId);
        });
    });
}

export function addFilterRule(categoryId) {
    const cat = categories.find(c => c.id === categoryId);
    const container = document.querySelector(`#custom-${categoryId} .filter-rules-container`);
    if (!cat || !container) return;
    const ruleId = `rule-${Date.now()}`;
    const rule = { id: ruleId, column: cat.headers[0] || '', operator: 'contains', value: '' };
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
            <option value="numeric_gt">&gt;</option>
            <option value="numeric_gte">&gt;=</option>
            <option value="numeric_lt">&lt;</option>
            <option value="numeric_lte">&lt;=</option>
        </select>
        <input type="text" data-type="value" class="w-full rounded-md input-base text-sm p-2">
        <button data-action="remove" class="p-1 text-red-500">&times;</button>
    `;
    container.appendChild(ruleEl);
    ruleEl.addEventListener('change', (e) => { rule[e.target.dataset.type] = e.target.value; updatePreview(categoryId); });
    ruleEl.addEventListener('keyup', (e) => { if (e.target.dataset.type === 'value') { rule.value = e.target.value; updatePreview(categoryId); } });
    ruleEl.querySelector('[data-action=remove]').addEventListener('click', () => { cat.filterRules = cat.filterRules.filter(r => r.id !== ruleId); ruleEl.remove(); updatePreview(categoryId); });
}

export function updatePreview(categoryId) {
    const cat = categories.find(c => c.id === categoryId);
    const section = document.querySelector(`#custom-${categoryId}`);
    if(!section || !cat) return;
    const filteredData = cat.data.filter(row => {
        return (cat.filterRules || []).every(rule => {
            if (!rule.value) return true;
            const raw = row[rule.column];
            // Use the locale-formatted display string for matching so user-entered slashed dates (DD/MM) match what they see
            const displayStr = String(formatByLocale(raw == null ? '' : raw, row && row.__sourceLocale ? row.__sourceLocale : undefined)).toLowerCase();
            const filterValStr = String(rule.value).toLowerCase();
            // strict numeric detection: only treat as numeric if both sides are pure numbers (no slashes/letters)
            const numericRe = /^\s*-?\d+(?:\.\d+)?\s*$/;
            const isRowNumeric = numericRe.test(String(raw == null ? '' : raw));
            const isFilterNumeric = numericRe.test(String(rule.value));
            const bothNumeric = isRowNumeric && isFilterNumeric;
            if (rule.operator === 'contains') return displayStr.includes(filterValStr);
            if (rule.operator === 'not_contains') return !displayStr.includes(filterValStr);
            if (rule.operator === 'equals') return displayStr === filterValStr;
            if (bothNumeric) {
                const maybeRowNum = Number(String(raw).trim());
                const maybeFilterNum = Number(String(rule.value).trim());
                if (rule.operator === 'numeric_gt') return maybeRowNum > maybeFilterNum;
                if (rule.operator === 'numeric_gte') return maybeRowNum >= maybeFilterNum;
                if (rule.operator === 'numeric_lt') return maybeRowNum < maybeFilterNum;
                if (rule.operator === 'numeric_lte') return maybeRowNum <= maybeFilterNum;
            }
            // fallback: string compare for the numeric ops
            if (rule.operator === 'numeric_gt') return rowValStr > filterValStr;
            if (rule.operator === 'numeric_gte') return rowValStr >= filterValStr;
            if (rule.operator === 'numeric_lt') return rowValStr < filterValStr;
            if (rule.operator === 'numeric_lte') return rowValStr <= filterValStr;
            return true;
        });
    });
    const head = section.querySelector('.preview-head');
    const body = section.querySelector('.preview-body');
    head.innerHTML = `<tr>${(cat.selectedColumns||[]).map(h => `<th class="px-6 py-3">${h}</th>`).join('')}</tr>`;
    body.innerHTML = filteredData.slice(0,50).map(row => `
        <tr class="border-b border-base hover:bg-surface-hover">
            ${(cat.selectedColumns||[]).map(h => `<td class="px-6 py-4">${formatByLocale(row[h], row.__sourceLocale)}</td>`).join('')}
        </tr>
    `).join('');
    const countEl = section.querySelector('.preview-row-count'); if (countEl) countEl.textContent = `${filteredData.length} rows`;
}

export function createPdf() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'landscape' });
    let y = 20;
    // Draw logo if present (stored as dataURL in localStorage)
    const logoData = localStorage.getItem('APP_LOGO_DATAURL');
    const pageW = doc.internal.pageSize.getWidth();
    const margin = 14;
    let headerRightX = margin;
    let headerBottomY = y;
    let logoDrawW = 0, logoDrawH = 0;
    if (logoData) {
        try {
            const imgProps = doc.getImageProperties(logoData);
            const maxH = 18; // mm
            const ratio = imgProps.width / imgProps.height;
            logoDrawW = maxH * ratio;
            logoDrawH = maxH;
            // draw logo at left
            doc.addImage(logoData, 'PNG', margin, y - 6, logoDrawW, logoDrawH);
            headerRightX = margin + logoDrawW + 8; // place title to the right of logo with a gap
            headerBottomY = Math.max(headerBottomY, y - 6 + logoDrawH);
        } catch (err) {
            console.warn('Failed to draw logo on PDF', err);
            headerRightX = margin;
        }
    }

    const reportTitle = localStorage.getItem('APP_REPORT_TITLE') || 'Report';
    const reportSubtitle = localStorage.getItem('APP_REPORT_SUBTITLE') || '';
    // available width for title (to the right of logo)
    const availableW = pageW - margin - headerRightX;
    doc.setFontSize(18);
    const titleLines = doc.splitTextToSize(reportTitle, availableW);
    doc.text(titleLines, headerRightX, y);
    // estimate title/subtitle block height (convert points to mm roughly: 1pt = 0.352777 mm)
    const ptsToMm = (pt) => pt * 0.352777;
    const titleFontPt = 18; // we set 18pt
    const titleLineHeight = titleLines.length * ptsToMm(titleFontPt * 1.15);
    let titleBlockH = titleLineHeight;
    y += 0; // don't advance y yet; we'll compute header height below
    if (reportSubtitle) {
        doc.setFontSize(11);
        const subLines = doc.splitTextToSize(reportSubtitle, availableW);
        doc.text(subLines, headerRightX, y + ptsToMm(titleFontPt * 1.15));
        const subFontPt = 11;
        const subBlockH = subLines.length * ptsToMm(subFontPt * 1.15);
        titleBlockH += subBlockH;
    }
    // header height is the max of logo height and text block height
    const headerHeight = Math.max(logoDrawH || 0, titleBlockH);
    const smallGap = 6; // mm
    y = margin + headerHeight + smallGap;
    categories.forEach(cat => {
        doc.setFontSize(14); doc.text(cat.name, 14, y); y += 7;
        const filteredData = (cat.data || []).filter(row => (cat.filterRules||[]).every(rule => {
            if (!rule.value) return true;
            const raw = row[rule.column];
            const rv = String(raw == null ? '' : raw).toLowerCase();
            return rv.includes(String(rule.value).toLowerCase());
        }));

        const selCols = (cat.selectedColumns||[]);
        // Separate image column (case-insensitive match) from textual columns
        const imageColIndex = selCols.findIndex(h => h && h.toString().toLowerCase() === 'image');
        const textCols = selCols.filter((_, i) => i !== imageColIndex);

        // Render textual table if there are text columns
        if (textCols.length) {
            const body = filteredData.map(row => textCols.map(h => String(formatByLocale(row[h] || '', row.__sourceLocale))));
            doc.autoTable({ startY: y, head: [textCols], body: body, theme: 'striped' });
            y = doc.autoTable.previous.finalY + 8;
        }

        // Collect image items (resolve library images) and render them larger under the category
        if (imageColIndex !== -1) {
            const images = [];
            filteredData.forEach(row => {
                const rawVal = row[selCols[imageColIndex]];
                if (!rawVal) return;
                const imgItem = getLibraryItems().find(i => i.isImage && (i.name === rawVal || i.id === rawVal));
                if (imgItem && imgItem.dataUrl) images.push({ dataUrl: imgItem.dataUrl, name: imgItem.name });
            });

            if (images.length) {
                // Page measurements (mm)
                const pageW = doc.internal.pageSize.getWidth();
                const pageH = doc.internal.pageSize.getHeight();
                const margin = 14;
                const maxW = pageW - margin * 2;
                const maxH = pageH - margin * 2 - 20; // leave space for header/title

                images.forEach((img, idx) => {
                    // If not enough vertical space for a reasonably sized image, add a new page
                    if (y + 10 >= pageH - margin) {
                        doc.addPage();
                        y = margin;
                    }
                    // draw caption
                    doc.setFontSize(12);
                    doc.text(img.name || `${cat.name} image ${idx+1}`, margin, y);
                    y += 6;
                    try {
                        const props = doc.getImageProperties(img.dataUrl);
                        const ratio = props.width / props.height;
                        let drawW = maxW;
                        let drawH = drawW / ratio;
                        if (drawH > maxH) {
                            drawH = maxH;
                            drawW = drawH * ratio;
                        }
                        // If image would overflow current page vertically, start a new page
                        if (y + drawH > pageH - margin) {
                            doc.addPage();
                            y = margin + 6; // header space
                        }
                        const x = (pageW - drawW) / 2;
                        doc.addImage(img.dataUrl, props.fileType ? props.fileType.replace('image/','').toUpperCase() : 'PNG', x, y, drawW, drawH);
                        y += drawH + 10;
                    } catch (err) {
                        console.warn('Failed to draw gallery image', err);
                    }
                });
                // leave some breathing room before next category
                y += 6;
            }
        }
        // ensure if we approach the page bottom we add a new page
        const pageHCheck = doc.internal.pageSize.getHeight();
        if (y > pageHCheck - 40) { doc.addPage(); y = 20; }
    });
    doc.save('report.pdf');
}

// Format value for UK display: DD/MM/YYYY for dates, otherwise string default
function formatToUK(val) {
    if (val == null) return '';
    // Date object
    if (val instanceof Date) {
        const d = val.getDate().toString().padStart(2,'0');
        const m = (val.getMonth()+1).toString().padStart(2,'0');
        const y = val.getFullYear();
        return `${d}/${m}/${y}`;
    }
    // ISO-like YYYY-MM-DD
    if (typeof val === 'string') {
        const iso = val.match(/^\s*(\d{4})-(\d{2})-(\d{2})(?:T.*)?\s*$/);
        if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;
        // numeric slashed date, assume UK day/month/year
        const sl = val.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
        if (sl) {
            const day = sl[1].padStart(2,'0');
            const mon = sl[2].padStart(2,'0');
            const year = sl[3].length === 2 ? (parseInt(sl[3],10) > 50 ? '19'+sl[3] : '20'+sl[3]) : sl[3];
            return `${day}/${mon}/${year}`;
        }
        // dashed numeric date like DD-MM-YYYY
        const dash = val.match(/^\s*(\d{1,2})-(\d{1,2})-(\d{2,4})\s*$/);
        if (dash) {
            const day = dash[1].padStart(2,'0');
            const mon = dash[2].padStart(2,'0');
            const year = dash[3].length === 2 ? (parseInt(dash[3],10) > 50 ? '19'+dash[3] : '20'+dash[3]) : dash[3];
            return `${day}/${mon}/${year}`;
        }
        // text month formats '1 Jan 2025'
        const m2 = val.match(/^\s*(\d{1,2})\s+([A-Za-z]{3,})\.?\s*,?\s*(\d{4})\s*$/);
        if (m2) {
            const day = m2[1].padStart(2,'0');
            const monStr = m2[2].toLowerCase();
            const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
            const mi = months.indexOf(monStr.slice(0,3));
            if (mi !== -1) return `${day}/${String(mi+1).padStart(2,'0')}/${m2[3]}`;
        }
    }
    return String(val);
}
    // Parse a date-like string according to locale preference and return canonical DD/MM/YYYY or null
    export function parseDateByLocale(str, locale = 'uk') {
        if (!str || typeof str !== 'string') return null;
        const s = str.trim();
        // already ISO-like -> convert to DD/MM/YYYY
        const isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        const toDD = (y, m, d) => `${String(d).padStart(2,'0')}/${String(m).padStart(2,'0')}/${String(y).padStart(4,'0')}`;
        const normYear = (y) => { if (y < 100) return y + (y > 50 ? 1900 : 2000); return y; };
        const valid = (y, m, d) => { const dt = new Date(y, m - 1, d); return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d; };

        if (isoMatch) {
            const y = parseInt(isoMatch[1], 10);
            const m = parseInt(isoMatch[2], 10);
            const d = parseInt(isoMatch[3], 10);
            if (valid(y, m, d)) return toDD(y, m, d);
        }

        const loc = (locale || 'uk').toString().toLowerCase();

        // Numeric slashed dates like 03/04/2024
        const slashMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
        if (slashMatch) {
            let a = parseInt(slashMatch[1], 10);
            let b = parseInt(slashMatch[2], 10);
            let y = normYear(parseInt(slashMatch[3], 10));

            const tryUk = () => { // DD/MM/YYYY -> day=a, month=b
                if (valid(y, b, a)) return toDD(y, b, a);
                return null;
            };
            const tryUs = () => { // MM/DD/YYYY -> month=a, day=b
                if (valid(y, a, b)) return toDD(y, a, b);
                return null;
            };

            if (loc === 'uk') {
                const r = tryUk();
                if (r) return r;
                return tryUs(); // fallback if UK parse invalid
            }
            if (loc === 'us') {
                const r = tryUs();
                if (r) return r;
                return tryUk();
            }

            // auto: if one part > 12 it's the day; otherwise try both preferring UK
            if (a > 12 && b <= 12) {
                return tryUk();
            }
            if (b > 12 && a <= 12) {
                return tryUs();
            }
            // ambiguous (both <=12): prefer UK but validate
            const rUk = tryUk();
            if (rUk) return rUk;
            const rUs = tryUs();
            if (rUs) return rUs;
            return null;
        }

        // Formats like '12 Jan 2025' or 'Jan 12, 2025'
        const dayMonthName = s.match(/^\s*(\d{1,2})\s+([A-Za-z]{3,})\.?\,?\s*(\d{4})\s*$/);
        if (dayMonthName) {
            const day = parseInt(dayMonthName[1], 10);
            const monStr = dayMonthName[2].toLowerCase().slice(0,3);
            const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
            const mi = months.indexOf(monStr);
            if (mi !== -1) {
                const month = mi + 1;
                const year = parseInt(dayMonthName[3], 10);
                if (valid(year, month, day)) return toDD(year, month, day);
            }
        }
        const monthDayName = s.match(/^\s*([A-Za-z]{3,})\.?\,?\s*(\d{1,2}),?\s*(\d{4})\s*$/);
        if (monthDayName) {
            const monStr = monthDayName[1].toLowerCase().slice(0,3);
            const day = parseInt(monthDayName[2], 10);
            const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
            const mi = months.indexOf(monStr);
            if (mi !== -1) {
                const month = mi + 1;
                const year = parseInt(monthDayName[3], 10);
                if (valid(year, month, day)) return toDD(year, month, day);
            }
        }

        // Fallback: try Date.parse (last resort) — return DD/MM/YYYY if successful
        const parsed = Date.parse(s);
        if (!Number.isNaN(parsed)) {
            const dt = new Date(parsed);
            return toDD(dt.getFullYear(), dt.getMonth() + 1, dt.getDate());
        }
        return null;
    }
export { formatToUK };

// Format value according to locale (row-level locale). Falls back to UK formatting.
export function formatByLocale(val, locale) {
    const loc = locale || 'uk';
    if (val == null) return '';
    if (loc === 'us') {
        // US format: MM/DD/YYYY
        // handle ISO input (YYYY-MM-DD)
        if (typeof val === 'string') {
            const iso = val.match(/^\s*(\d{4})-(\d{2})-(\d{2})(?:T.*)?\s*$/);
            if (iso) return `${iso[2]}/${iso[3]}/${iso[1]}`;
            // canonical stored format is DD/MM/YYYY (slashes)
            const slash = val.match(/^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/);
            if (slash) {
                const day = slash[1].padStart(2,'0');
                const mon = slash[2].padStart(2,'0');
                const y = slash[3].length === 2 ? (parseInt(slash[3],10) > 50 ? '19'+slash[3] : '20'+slash[3]) : slash[3];
                // interpret stored DD/MM/YYYY and display as MM/DD/YYYY for US
                return `${mon}/${day}/${y}`;
            }
        }
        // fallback to generic string
        return String(val);
    }
    // default: UK
    return formatToUK(val);
}
