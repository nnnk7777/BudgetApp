export class AccountingItem {
  private date: string;
  private title: string;
  private price: number;

  constructor(title: string, price: number, date: string = "") {
    this.title = title;
    this.price = price;
    this.date = date;
  }

  getTitle(): string {
    return this.title;
  }

  getPrice(): number {
    return this.price;
  }

  getDate(): string {
    return this.date;
  }
}
