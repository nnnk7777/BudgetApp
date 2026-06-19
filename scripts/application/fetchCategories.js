function fetchCategories() {
    const categoryNames = categories.map(c => c.name);

    return JSON.stringify(categoryNames);
}
