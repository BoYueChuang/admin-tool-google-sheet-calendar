import { Event } from "./Events";
import Spaces, { Space } from "./Spaces";
import { mapDump, Index, Moment, moment } from "./Util";

const index = new Index<Index<Schedule[]>>();

export class Schedule {
  readonly event: Event;
  readonly begin: Moment;
  readonly end: Moment;

  constructor(event: Event, being: Moment, end: Moment) {
    this.event = event;
    this.begin = being;
    this.end = end;
  }

  compare(that: Schedule): number {
    return this.compareTypeId(that) || this.compareEventId(that);
  }

  compareTypeId(that: Schedule): number {
    return parseInt(that.event.type.id.substring(1)) - parseInt(this.event.type.id.substring(1));
  }

  compareEventId(that: Schedule): number {
    return parseInt(that.event.id.substring(1)) - parseInt(this.event.id.substring(1));
  }

  conflict(that: Schedule): boolean {
    return this.begin.isBefore(that.end) || this.end.isAfter(that.begin);
  }

  toString(): string {
    return `{event=${this.event.id},time=${this.begin.format("HH")}~${this.end.format("HH")}}`;
  }
}

export default {
  set: (event: Event): void => {
    event.times.forEach((time) => {
      event.spaces.forEach((space) => {
        index
          .getOrDefault(() => new Index(), time.date)
          .getOrDefault(() => [], space.id)
          .push(new Schedule(event, time.begin, time.end));
      });
    });
  },
  forEach: (cb: (d: Moment, s: Space, es: Schedule[]) => void): void => {
    index.forEach((schedules, date) => {
      schedules.forEach((eventSchedule, spaceId) => {
        cb(moment(date, "YYYY-MM-DD"), Spaces.get(spaceId), eventSchedule);
      });
    });
  },
  computeConflicts: (schedules: Schedule[]): Map<string, string[]> => {
    const result = new Index<string[]>();
    schedules.sort((a, b) => a.compare(b));
    for (var lo = 0; lo < schedules.length; lo++) {
      for (var hi = lo + 1; hi < schedules.length; hi++) {
        if (schedules[lo].conflict(schedules[hi])) {
          result.getOrDefault(() => [], schedules[lo].event.id).push(schedules[hi].event.id);
        }
      }
    }
    return result;
  },
  dump(): string {
    return [...index.entries()].map(([date, spaces]) => `[${date} => [${mapDump(spaces)}]]`).join("\n");
  },
} as const;
