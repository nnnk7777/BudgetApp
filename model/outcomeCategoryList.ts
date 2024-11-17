import { Category } from "./category";
import { CategoryList } from "./categoryList";
import categories from "../config/categories.js";

export class OutcomeCategoryList implements CategoryList {
    private list: Category[];

    constructor() {
        this.list = categories.map((item) => {
            return new Category(item.name, item.color, item.textColor);
        });
    }

    public getCategoryList(): Category[] {
        return this.list;
    }

    public getCategoryNames(): string[] {
        return this.list.map((category) => category.getCategoryName());
    }
}
