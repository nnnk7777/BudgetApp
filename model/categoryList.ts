import { Category } from "./category";

export interface CategoryList {
  getCategoryList(): Category[];
  getCategoryNames(): string[];
}
