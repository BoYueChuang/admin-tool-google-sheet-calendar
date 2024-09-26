import EventTypes, { EventType } from "./EventTypes";
import Spaces, { Space } from "./Spaces";
import { getContinuousRowsBothEnds, mapDump, moment, Moment } from "./Util";

export const EVENTS_SHEET_NAME = "Event List";
const EVENTS_SHEET_DATA_ANCHOR = "A4";
const EVENTS_SHEET_CHECK_ANCHOR = "L4";
const EVENTS_SHEET_WEEK_ANCHOR = "E4";
const EVENTS_SHEET_DATA_WIDTH = 10;

const index = new Map<string, Event>();

export class Event {
  readonly id: string;
  readonly name: string;
  readonly type: EventType;
  readonly spaces: Space[];
  readonly times: EventTime[];
  readonly owner: string;
  readonly wholeDay: boolean;

  constructor(id: string, name: string, type: EventType, spaces: Space[], times: EventTime[], owner: string, wholeDay: boolean) {
    this.id = id;
    this.name = name;
    this.type = type;
    this.spaces = spaces;
    this.times = times;
    this.owner = owner;
    this.wholeDay = wholeDay;
  }

  toString(): string {
    return `{id=${this.id},type=${this.type},spaces=[${this.spaces}],times=[${this.times}]}`;
  }
}

class EventTime {
  readonly date: Moment;
  readonly begin: Moment;
  readonly end: Moment;

  constructor(date: Moment, begin: Moment, end: Moment) {
    this.date = date;
    this.begin = begin;
    this.end = end;
  }

  toString(): string {
    return `${this.date.format("YYYY-MM-DD")}/${this.begin.format("HH")}~${this.end.format("HH")}`;
  }

  static explode(d1: string, d2: string, t1: string, t2: string): EventTime[] {
    const beginDate = moment(d1, "MM/DD/YYYY");
    const endDate = moment(d2, "MM/DD/YYYY");
    const begin = t1 ? moment(t1, "LTS") : moment("00:00", "hh:mm");
    const end = t2 ? moment(t2, "LTS") : moment("23:59", "hh:mm");
    const result = [];
    if (beginDate.isValid()) {
      result.push(new EventTime(beginDate, begin, end));
      let date = beginDate.clone().add(1, "days");
      while (endDate.isSameOrAfter(date)) {
        result.push(new EventTime(date, begin, end));
        date = date.clone().add(1, "days");
      }
    }
    return result;
  }
}

export default {
  load: (): void => {
    const anchor = getAnchor();
    const data = getEventsData(anchor);
    for (var row = 0; row < data.length; row++) {
      const id = anchor.offset(row, 0).getA1Notation();
      const [name, typeString, d1, d2, _, t1, t2, owner, location, spacesString] = data[row];
      const type = EventTypes.get(typeString);
      const spaces = getSpaces(location, spacesString);
      const times = EventTime.explode(d1, d2, t1, t2);
      const event = new Event(id, name, type, spaces, times, owner, !t1 && !t2);
      index.set(id, event);
    }
  },
  get: (id: string): Event => {
    return index.get(id);
  },
  forEach: (cb: (e: Event) => void): void => {
    index.forEach((event) => cb(event));
  },
  refreshSheet: (): void => {
    const anchor = getAnchor();
    const sheet = anchor.getSheet();
    const checkAnchor = sheet.getRange(EVENTS_SHEET_CHECK_ANCHOR);
    const weekAnchor = sheet.getRange(EVENTS_SHEET_WEEK_ANCHOR);
    const [first, last] = getContinuousRowsBothEnds(anchor);
    const height = last - first;
    for (var row = 0; row <= height; row++) {
      const checkCell = checkAnchor.offset(row, 0);
      const weekCell = weekAnchor.offset(row, 0);
      checkCell.setFormula(`=HOPE_CALENDAR_CHECK()`);
      weekCell.setFormula(`=WEEKNUM(${weekCell.offset(0, -2).getA1Notation()})`);
    }
  },
  dump: (): string => {
    return mapDump(index);
  },
} as const;

function getAnchor(): GoogleAppsScript.Spreadsheet.Range {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EVENTS_SHEET_NAME);
  return sheet.getRange(EVENTS_SHEET_DATA_ANCHOR);
}

function getEventsData(anchor: GoogleAppsScript.Spreadsheet.Range): string[][] {
  const row = anchor.getRow();
  const col = anchor.getColumn();
  const [first, last] = getContinuousRowsBothEnds(anchor);
  const height = last - first + 1;
  return anchor.getSheet().getRange(row, col, height, EVENTS_SHEET_DATA_WIDTH).getDisplayValues();
}

function getSpaces(location: string, spacesString: string): Space[] {
  const spaceNames = spacesString.split(",").map((s) => s.trim());
  const result = [];
  for (const name of spaceNames) {
    const allSpaces = Spaces.getGroup(location, name); // all-spaces is the only space group now
    if (allSpaces) {
      return allSpaces;
    }
    result.push(Spaces.get(location, name));
  }
  return result;
}
