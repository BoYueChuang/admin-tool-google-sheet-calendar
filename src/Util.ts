import moment from "moment";

export { default as moment } from "moment";

export type Moment = moment.Moment;

export function getContinuousRowsBothEnds(cell: GoogleAppsScript.Spreadsheet.Range): [number, number] {
  const nextDataCell = cell.getNextDataCell(SpreadsheetApp.Direction.DOWN);
  const firstRow = cell.getRow();
  const lastRow = nextDataCell.getValue() ? nextDataCell.getRow() : firstRow;
  return [firstRow, lastRow];
}

export function mapGetOrInit<T>(map: Map<string, T>, key: string, init: () => T): T {
  if (!map.has(key)) {
    map.set(key, init());
  }
  return map.get(key);
}

export function mapDump<T>(map: Map<string, T>): string {
  return [...map.entries()].map(([k, v]) => `[${k} => ${v}]`).join(",");
}

export class Index<T> extends Map<string, T> {
  ensureKey(key: Moment | string): string {
    if (moment.isMoment(key)) {
      key = key.format("YYYY-MM-DD");
    }
    return key;
  }

  get(key: Moment | string): T {
    return super.get(this.ensureKey(key));
  }

  getOrDefault(supplier: () => T, key1: Moment | string, key2?: Moment): T {
    let key = this.ensureKey(key1);
    if (key2 !== undefined) {
      key += `/${this.ensureKey(key2)}`;
    }
    if (!this.has(key)) {
      this.set(key, supplier());
    }
    return this.get(key);
  }

  set(key: Moment | string, value: T): this {
    return super.set(this.ensureKey(key), value);
  }
}
