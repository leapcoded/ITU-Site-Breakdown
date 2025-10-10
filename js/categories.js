// Category management module
export let categories = [];

// Library holds saved items (notes, images, parsed files)
export let libraryFiles = [];

export function addCategory(category) {
    categories.push(category);
}

export function removeCategory(categoryId) {
    categories = categories.filter(c => c.id !== categoryId);
}

export function getCategory(categoryId) {
    return categories.find(c => c.id === categoryId);
}

export function updateCategory(categoryId, updates) {
    const cat = getCategory(categoryId);
    if (cat) Object.assign(cat, updates);
}

export function addLibraryItem(item) {
    libraryFiles.push(item);
}

export function removeLibraryItemById(id) {
    libraryFiles = libraryFiles.filter(i => i.id !== id);
}

export function getLibraryItems() {
    return libraryFiles;
}

export function clearLibrary() {
    libraryFiles.length = 0;
}
