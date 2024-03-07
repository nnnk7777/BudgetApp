import { Category } from "./category";
import { CategoryList } from "./categoryList";

export class OutcomeCategoryList implements CategoryList {
  private list: Category[];

  constructor() {
    this.list = [
      new Category("コンビニ・お菓子", "#ffbb4d"),
      new Category("日用品・消耗品", "#d5ff9e"),
      new Category("健康・美容", "#f5f36c"),
      new Category("漫画・本", "#e3c852"),
      new Category("交通費", "#adfff6"),
      new Category("スーパー・食品など", "#f7c3ef"),
      new Category("ふるさと納税", "#d2691e"),
      new Category("外食", "#ff9e9e"),
      new Category("映画館・美術館など", "#7d50cc", "#ffffff"),
      new Category("服", "#65dbcc"),
      new Category("植物", "#2fba25", "#ffffff"),
      new Category("銭湯", "#368696", "#ffffff"),
      new Category("ゲーム・娯楽", "#6b97ff", "#ffffff"),
      new Category("アプリ・サブスク", "#4039bf", "#ffffff"),
      new Category("デバイス", "#6a8ead", "#ffffff"),
      new Category("家具・家電", "#5e2975", "#ffffff"),
      new Category("プレゼント・お土産", "#cfaee6"),
      new Category("家賃", "#f0819f"),
      new Category("公共料金など", "#545454", "#ffffff"),
      new Category("その他", "#c7c7c7"),
    ];
  }
  public getCategoryList(): Category[] {
    return this.list;
  }

  public getCategoryNames(): string[] {
    return this.list.map((category) => category.getCategoryName());
  }
}
