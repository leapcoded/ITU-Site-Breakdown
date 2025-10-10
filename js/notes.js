// Notes management module
export function addNoteToCategory(category, noteObj) {
    if (!category.notes) category.notes = [];
    category.notes.push(noteObj);
}

export function getNotes(category) {
    return category.notes || [];
}
