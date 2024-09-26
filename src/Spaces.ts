import { mapDump } from "./Util"

const SPACES_SHEET_NAME = "Setup_Spaces";
const SPACES_SHEET_DATA_ANCHOR = "B2";
const SPACES_SHEET_DATA_WIDTH = 6;
const SPACES_SHEET_DATA_HEIGHT = 29;

const index = new Map<string, Space>();
const groups = new Map<string, Space[]>();
const calendars = new Map<string, string>();

export class Space {
  readonly id: string;
  readonly location: string;
  readonly name: string;

  constructor(id: string, location: string, name: string) {
    this.id = id;
    this.location = location;
    this.name = name;
  }

  toString(): string {
    return this.id;
  }

  static keyOf(location: string, name: string): string {
    return `${location}/${name}`;
  }
}

export default {
  load: (): void => {
    const anchor = getAnchor();
    const data = getSpacesData(anchor);
    for (var loc = 0; loc < SPACES_SHEET_DATA_WIDTH; loc++) {
      const location = data[0][loc];
      if (!location) {
        break;
      }
      const calendar = data[1][loc];
      const groupAllSpaces = data[2][loc];
      const list = [];
      for (var row = 3; row < SPACES_SHEET_DATA_HEIGHT; row++) {
        const name = data[row][loc];
        if (!name) {
          break;
        }
        const id = anchor.offset(row, loc).getA1Notation();
        const space = new Space(id, location, name);
        list.push(space);
        index.set(id, space);
        index.set(Space.keyOf(location, name), space);
      }
      groups.set(Space.keyOf(location, groupAllSpaces), list);
      calendars.set(location, calendar);
    }
  },
  get: get,
  getGroup: (location: string, name: string): Space[] => {
    return groups.get(Space.keyOf(location, name));
  },
  getDropdownTemplate: (location: string): GoogleAppsScript.Spreadsheet.Range => {
    const anchor = getAnchor();
    const data = getLocationsData(anchor);
    for (var loc = 0; loc < SPACES_SHEET_DATA_WIDTH; loc++) {
      if (data[loc] === location) {
        return anchor.offset(SPACES_SHEET_DATA_HEIGHT, loc);
      }
    }
    return null;
  },
  getCalendarNames: (): Map<string, string> => {
    return calendars;
  },
  dump: (): string => {
    return `${mapDump(index)}\n${mapDump(groups)}`;
  },
} as const;

function get(id: string): Space;
function get(location: string, name: string): Space;
function get(idOrLocation: string, name?: string): Space {
  if (name !== undefined) {
    return index.get(Space.keyOf(idOrLocation, name));
  }
  return index.get(idOrLocation);
}

function getAnchor(): GoogleAppsScript.Spreadsheet.Range {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SPACES_SHEET_NAME);
  return sheet.getRange(SPACES_SHEET_DATA_ANCHOR);
}

function getSpacesData(anchor: GoogleAppsScript.Spreadsheet.Range): string[][] {
  const row = anchor.getRow();
  const col = anchor.getColumn();
  const range = anchor.getSheet().getRange(row, col, SPACES_SHEET_DATA_HEIGHT, SPACES_SHEET_DATA_WIDTH);
  return range.getDisplayValues();
}

function getLocationsData(anchor: GoogleAppsScript.Spreadsheet.Range): string[] {
  const row = anchor.getRow();
  const col = anchor.getColumn();
  const range = anchor.getSheet().getRange(row, col, 1, SPACES_SHEET_DATA_WIDTH);
  return range.getDisplayValues()[0];
}
