import { mapDump } from "./Util"

const EVENT_TYPES_SHEET_NAME = "Setup_Event Types";
const EVENT_TYPES_SHEET_DATA_ANCHOR = "A2";
const EVENT_TYPES_SHEET_DATA_HEIGHT = 29;

const index = new Map<string, EventType>();

export class EventType {
  readonly id: string;
  readonly name: string;
  readonly code: string;

  constructor(id: string, name: string, code: string) {
    this.id = id;
    this.name = name;
    this.code = code;
  }

  toString(): string {
    return `${this.id}`;
  }
}

export default {
  load: (): void => {
    const anchor = getAnchor();
    const types = getTypesData(anchor);
    for (var row = 0; row <= EVENT_TYPES_SHEET_DATA_HEIGHT; row++) {
      const id = anchor.offset(row, 0).getA1Notation();
      const name = types[row][0];
      if (!name) {
        break;
      }
      const eventType = new EventType(id, name, name.substring(0, 1));
      index.set(id, eventType);
      index.set(name, eventType);
    }
  },
  get: (idOrName: string): EventType => {
    return index.get(idOrName);
  },
  dump: (): string => {
    return mapDump(index);
  },
} as const;

function getAnchor(): GoogleAppsScript.Spreadsheet.Range {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVENT_TYPES_SHEET_NAME);
  return sheet.getRange(EVENT_TYPES_SHEET_DATA_ANCHOR);
}

function getTypesData(anchor: GoogleAppsScript.Spreadsheet.Range): string[][] {
  const row = anchor.getRow();
  const col = anchor.getColumn();
  const range = anchor.getSheet().getRange(row, col, EVENT_TYPES_SHEET_DATA_HEIGHT, 1);
  return range.getDisplayValues();
}
