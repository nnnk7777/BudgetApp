export class Category {
  private category: string;
  private backgroundColor: string;
  private fontColor: string | null;

  constructor(
    category: string,
    backgroundColor: string,
    fontColor: string = "#000"
  ) {
    this.category = category;
    this.backgroundColor = backgroundColor;
    this.fontColor = fontColor;
  }

  public getCategoryName(): string {
    return this.category;
  }

  public getBackgroundColor(): string {
    return this.backgroundColor;
  }

  public getFontColor(): string | null {
    return this.fontColor;
  }
}
