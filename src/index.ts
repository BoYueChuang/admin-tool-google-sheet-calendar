import EventTypes from "./EventTypes";
import Spaces from "./Spaces";
import Events, { EVENTS_SHEET_NAME } from "./Events";
import Schedules from "./Schedules";
import Calendar from "./Calendar";
import { Index } from "./Util";

export function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('The Hope Calendar')
    .addItem('Refresh Events', 'refreshEvents')
    .addItem('Refresh Calendar', 'refreshCalendar')
    .addToUi();
}

export function refreshEvents(): void {
  Events.refreshSheet();
}

export function refreshCalendar(): void {
  EventTypes.load();
  Spaces.load();
  Events.load();
  Events.forEach(Schedules.set);
  // Logger.log(EventTypes.dump());
  // Logger.log(Spaces.dump());
  // Logger.log(Events.dump());
  // Logger.log(Schedules.dump());
  Calendar.load();
  const hasConflict = new Index<Set<string>>();
  Schedules.forEach((d, _, ss) => {
    Schedules.computeConflicts(ss).forEach((_, eventId) => {
      hasConflict.getOrDefault(() => new Set(), d).add(eventId);
    });
  });
  Schedules.forEach((d, _, es) => {
    es.forEach((e) => Calendar.set(d, e, hasConflict.get(d)?.has(e.event.id)));
  });
  Calendar.render();
}

export function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  const { range, value } = e;
  if (range.getSheet().getName() === EVENTS_SHEET_NAME && range.getColumn() === 9) {
    const selectedLocation = value;
    setSpaceDropdown(range.offset(0, 1), selectedLocation);
  }
}

function setSpaceDropdown(cell: GoogleAppsScript.Spreadsheet.Range, location: string): void {
  const dropdown = Spaces.getDropdownTemplate(location);
  if (dropdown) {
    dropdown.copyTo(cell);
  } else {
    cell.clear();
    cell.clearDataValidations();
  }
}
