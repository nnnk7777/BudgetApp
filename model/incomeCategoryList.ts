import { Category } from "./category";
import { CategoryList } from "./categoryList";

export class IncomeCategoryList implements CategoryList {
  private list: Category[];

  constructor() {
    this.list = [
      new Category("給料", "#f58f7f"),
      new Category("割り勘・立て替え", "#f5d97f"),
      new Category("売却", "#7fb8f5"),
      new Category("その他", "#ededed"),
    ];
  }

  public getCategoryList(): Category[] {
    return this.list;
  }

  public getCategoryNames(): string[] {
    return this.list.map((category) => category.getCategoryName());
  }
}
