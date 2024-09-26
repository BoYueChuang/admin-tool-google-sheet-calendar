import Spaces from "./Spaces";
import { Index, moment, Moment } from "./Util";
import { Schedule } from "./Schedules";

const CALENDAR_SHEET_RANGE = "A1:AE97";
const CALENDAR_SHEET_ANCHOR_TOP = 2;
const CALENDAR_SHEET_ANCHOR_LEFT = 2;
const CALENDAR_SHEET_ANCHOR_RIGHT = 17;

const STYLE_EVENT_TITLE = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("#434343").build();
const STYLE_EVENT_TIME = SpreadsheetApp.newTextStyle().setForegroundColor("#4a86e8").build();
const STYLE_EVENT_SPACES = SpreadsheetApp.newTextStyle().setForegroundColor("#434343").build();
const STYLE_EVENT_OWNER = SpreadsheetApp.newTextStyle().setForegroundColor("#999999").build();

const EMPTY_CELL_DATA = { userEnteredValue: { stringValue: "" } };

interface DatePosition {
  row: number,
  column: number,
  dayOffset: number,
}

class DateRecord {
  index: Set<string> = new Set();
  entries: Entry[] = [];
}

const datePositions = new Index<DatePosition>();
const dateRecords = new Index<DateRecord>();
const batchUpdateData = new Map<string, Map<string, any[]>>();

interface RichTextPart {
  text: string,
  style: GoogleAppsScript.Spreadsheet.TextStyle,
}

class RichText {
  parts: RichTextPart[] = [];

  toString(): string {
    return this.parts.map((p) => p.text).join(" ");
  }

  setStyle(builder: GoogleAppsScript.Spreadsheet.RichTextValueBuilder, begin: number): void {
    let lo = begin;
    for (var i = 0; i < this.parts.length; i++) {
      const part = this.parts[i];
      const hi = lo + part.text.length;
      builder.setTextStyle(lo, hi, part.style);
      lo = hi + 1;
    }
  }
}

export class Entry {
  schedule: Schedule;
  warning: boolean;

  constructor(schedule: Schedule, warning: boolean) {
    this.schedule = schedule;
    this.warning = warning;
  }

  toRichText(): RichText {
    const result = new RichText();
    const warn = this.warning ? "⚠️" : "";
    const name = this.schedule.event.name;
    const code = this.schedule.event.type.code;
    result.parts.push({ text: `${warn}[${code}] ${name}`, style: STYLE_EVENT_TITLE });
    if (!this.schedule.event.wholeDay) {
      const begin = this.schedule.begin.format("HH:mm");
      const end = this.schedule.end.format("HH:mm");
      result.parts.push({ text: `${begin}-${end}`, style: STYLE_EVENT_TIME });
    }
    const spaces = this.schedule.event.spaces.map((s) => s.name).join(",");
    result.parts.push({ text: spaces, style: STYLE_EVENT_SPACES });
    if (this.schedule.event.owner) {
      result.parts.push({ text: this.schedule.event.owner, style: STYLE_EVENT_OWNER });
    }
    return result;
  }
}

export default {
  load: (): void => {
    Spaces.getCalendarNames().forEach((_, location) => {
      batchUpdateData.set(location, new Map());
    });
    const anyCalendar = Spaces.getCalendarNames().values().next().value;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(anyCalendar);
    digest(sheet.getRange(CALENDAR_SHEET_RANGE).getValues());
  },
  set: (date: Moment, schedule: Schedule, warning: boolean): void => {
    const thatDay = dateRecords.getOrDefault(() => new DateRecord(), date);
    if (!thatDay.index.has(schedule.event.id)) { // set multi-space event only once
      thatDay.index.add(schedule.event.id);
      thatDay.entries.push(new Entry(schedule, warning));
    }
  },
  render: (): void => {
    dateRecords.forEach((values, dateKey) => {
      if (!datePositions.has(dateKey)) {
        throw new Error("請確認每個活動時間都是在此次計畫年度當中");
      }
      const pos = datePositions.get(dateKey);
      const loc = new Index<Entry[]>();
      values.entries.forEach((entry) => {
        const location = entry.schedule.event.spaces[0].location;
        loc.getOrDefault(() => [], location).push(entry);
      });
      loc.forEach((entries, location) => {
        const cell = toCellData(toRichTextValue(entries));
        batchUpdateData.get(location).get(`${pos.row}:${pos.column}`)[pos.dayOffset] = cell;
      });
    });
    const ssid = SpreadsheetApp.getActiveSpreadsheet().getId();
    Spaces.getCalendarNames().forEach((calendar, location) => {
      const sid = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(calendar).getSheetId();
      const req: GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest = {
        requests: [],
      }
      batchUpdateData.get(location).forEach((data, index) => {
        req.requests.push({ updateCells: updateCells(sid, index, data) });
      });
      const res = Sheets.Spreadsheets.batchUpdate(req, ssid);
      if (res) {
        Logger.log(res);
      }
    });
  },
} as const;

function digest(data: any[][]): void {
  for (var row = CALENDAR_SHEET_ANCHOR_TOP; row < data.length; row++) {
    if (data[row][CALENDAR_SHEET_ANCHOR_LEFT - 1] === "Sunday") {
      digestMonth(row - 1, CALENDAR_SHEET_ANCHOR_LEFT - 1, data); // from 1-based to 0-based
      digestMonth(row - 1, CALENDAR_SHEET_ANCHOR_RIGHT - 1, data); // from 1-based to 0-based
    }
  }
}

function digestMonth(row: number, col: number, data: any[][]): void {
  const thisMonth = moment(data[row][col], "MMMM").month();
  row += 2;
  while (row < data.length && data[row][col]) {
    const dataRow = row + 1;
    batchUpdateData.forEach((calendarData, _) => {
      calendarData.set(`${dataRow}:${col}`, Array(14).fill(EMPTY_CELL_DATA)); // note the merged cells
    });
    let c = col;
    for (var day = 0; day < 7; day++) {
      const date = moment(data[row][c]);
      if (date.month() === thisMonth) {
        datePositions.set(date, { row: dataRow, column: col, dayOffset: day * 2 }); // note the merged cells
      }
      c += 2;
    }
    row += 2;
  }
}

function toRichTextValue(entries: Entry[]): GoogleAppsScript.Spreadsheet.RichTextValue {
  const richTexts = entries.map((e) => e.toRichText());
  const text = richTexts.map((r) => r.toString()).join("\n");
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  const beginOf = [...text].flatMap((c, i) => (c === "\n" ? i + 1 : []));
  beginOf.unshift(0);
  for (var i = 0; i < beginOf.length; i++) {
    richTexts[i].setStyle(builder, beginOf[i]);
  }
  return builder.build();
}

interface CellData {
  userEnteredValue: {
    stringValue: string,
  },
  textFormatRuns: TextFormatRun[],
}

interface TextFormatRun {
  startIndex: number,
  format: {
    foregroundColorStyle: ColorStyle,
    fontFamily: string,
    fontSize: number,
    bold: boolean,
    italic: boolean,
    strikethrough: boolean,
    underline: boolean,
  },
}

interface ColorStyle {
  rgbColor: {
    red: number,
    green: number,
    blue: number,
  },
}

function toCellData(richText: GoogleAppsScript.Spreadsheet.RichTextValue): CellData {
  return {
    userEnteredValue: { stringValue: richText.getText() },
    textFormatRuns: richText.getRuns().map(toTextFormatRun),
  };
}

function toTextFormatRun(run: GoogleAppsScript.Spreadsheet.RichTextValue): TextFormatRun {
  const style = run.getTextStyle();
  return {
    startIndex: run.getStartIndex(),
    format: {
      foregroundColorStyle: toColorStyle(style.getForegroundColorObject()),
      fontFamily: style.getFontFamily(),
      fontSize: style.getFontSize(),
      bold: style.isBold(),
      italic: style.isItalic(),
      strikethrough: style.isStrikethrough(),
      underline: style.isUnderline(),
    },
  };
}

function toColorStyle(colorObj: GoogleAppsScript.Spreadsheet.Color): ColorStyle {
  const c = colorObj?.asRgbColor();
  return !c ? null : {
    rgbColor: {
      red: c.getRed() / 255,
      green: c.getGreen() / 255,
      blue: c.getBlue() / 255,
    },
  };
}

function updateCells(sheet: number, index: string, data: any[]): GoogleAppsScript.Sheets.Schema.UpdateCellsRequest {
  const [row, col] = index.split(":");
  return {
    rows: [{ values: data }],
    fields: "*",
    start: {
      sheetId: sheet,
      rowIndex: parseInt(row),
      columnIndex: parseInt(col),
    },
  }
}
