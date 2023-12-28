import { Category } from "./category";
import { CategoryList } from "./categoryList";

export class OutcomeCategoryList implements CategoryList {
  private list: Category[];

  constructor() {
    this.list = [
      new Category("コンビニ・お菓子", "#ffbb4d"),
      new Category("日用品・消耗品", "#d5ff9e"),
    ];
  }
  public getCategoryList(): Category[] {
    return this.list;
  }

  public getCategoryNames(): string[] {
    return this.list.map((category) => category.getCategoryName());
  }
}
