// Image management module
export function addImageToCategory(category, imageObj) {
    if (!category.images) category.images = [];
    category.images.push(imageObj);
}

export function removeImageFromCategory(category, imageId) {
    if (!category.images) return;
    category.images = category.images.filter(img => img.id !== imageId);
}

export function getImages(category) {
    return category.images || [];
}
