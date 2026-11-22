import { describe, expect, test } from "bun:test";
import {
  findSlots,
  formatLocalMinute,
  freeBusyIntervals,
  localMidnight,
  mergeIntervals,
  parseBusinessHours,
} from "../src/tools/freeslots.ts";

const hours = parseBusinessHours("09:00-18:00");

describe("free-slot arithmetic", () => {
  test("business hours are validated", () => {
    expect(parseBusinessHours("09:00-18:00")).toEqual({
      startHour: 9,
      startMinute: 0,
      endHour: 18,
      endMinute: 0,
    });
    expect(() => parseBusinessHours("18:00-09:00")).toThrow();
    expect(() => parseBusinessHours("9-18")).toThrow();
    expect(() => parseBusinessHours("25:00-26:00")).toThrow();
  });

  test("a free/busy string becomes merged intervals", () => {
    const start = new Date(2026, 3, 20, 0, 0, 0);
    const intervals = freeBusyIntervals("002200", start, 30);
    expect(intervals).toHaveLength(1);
    expect(formatLocalMinute(intervals[0]!.start)).toBe("2026-04-20T01:00");
    expect(formatLocalMinute(intervals[0]!.end)).toBe("2026-04-20T02:00");
  });

  test("adjacent and overlapping intervals merge", () => {
    const base = new Date(2026, 3, 20, 9, 0, 0);
    const at = (minutes: number) => new Date(base.getTime() + minutes * 60_000);
    const merged = mergeIntervals([
      { start: at(60), end: at(120) },
      { start: at(0), end: at(30) },
      { start: at(30), end: at(60) },
      { start: at(180), end: at(240) },
      { start: at(240), end: at(300) },
    ]);
    expect(merged).toHaveLength(2);
    expect(formatLocalMinute(merged[0]!.start)).toBe("2026-04-20T09:00");
    expect(formatLocalMinute(merged[0]!.end)).toBe("2026-04-20T11:00");
    expect(formatLocalMinute(merged[1]!.start)).toBe("2026-04-20T12:00");
    expect(formatLocalMinute(merged[1]!.end)).toBe("2026-04-20T14:00");
  });

  test("finds slots around a busy block on a Monday", () => {
    const now = new Date(2026, 3, 20, 9, 0, 0);
    const slots = findSlots({
      now,
      daysAhead: 1,
      durationMinutes: 30,
      businessHours: hours,
      weekdaysOnly: true,
      stepMinutes: 30,
      limit: 4,
      busy: [
        { start: new Date(2026, 3, 20, 10, 0, 0), end: new Date(2026, 3, 20, 11, 0, 0) },
      ],
    });
    expect(slots.map((s) => s.start)).toEqual([
      "2026-04-20T09:00",
      "2026-04-20T09:30",
      "2026-04-20T11:00",
      "2026-04-20T11:30",
    ]);
  });

  test("skips weekends when weekdays_only is true", () => {
    const saturday = new Date(2026, 3, 25, 9, 0, 0);
    const slots = findSlots({
      now: saturday,
      daysAhead: 3,
      durationMinutes: 30,
      businessHours: hours,
      weekdaysOnly: true,
      stepMinutes: 60,
      limit: 2,
      busy: [],
    });
    expect(slots.map((s) => s.start)).toEqual(["2026-04-27T09:00", "2026-04-27T10:00"]);
  });

  test("does not cross the end of the business day", () => {
    const now = new Date(2026, 3, 20, 17, 0, 0);
    const slots = findSlots({
      now,
      daysAhead: 2,
      durationMinutes: 60,
      businessHours: hours,
      weekdaysOnly: true,
      stepMinutes: 30,
      limit: 3,
      busy: [],
    });
    expect(slots.map((s) => s.start)).toEqual([
      "2026-04-20T17:00",
      "2026-04-21T09:00",
      "2026-04-21T09:30",
    ]);
  });

  test("an all-day busy interval removes the day", () => {
    const now = new Date(2026, 3, 20, 9, 0, 0);
    const slots = findSlots({
      now,
      daysAhead: 2,
      durationMinutes: 60,
      businessHours: hours,
      weekdaysOnly: true,
      stepMinutes: 60,
      limit: 2,
      busy: [
        { start: new Date(2026, 3, 20, 0, 0, 0), end: new Date(2026, 3, 21, 0, 0, 0) },
      ],
    });
    expect(slots.map((s) => s.start)).toEqual(["2026-04-21T09:00", "2026-04-21T10:00"]);
  });

  test("a slot that starts before now is not returned", () => {
    const now = new Date(2026, 3, 20, 9, 17, 0);
    const slots = findSlots({
      now,
      daysAhead: 1,
      durationMinutes: 30,
      businessHours: hours,
      weekdaysOnly: true,
      stepMinutes: 30,
      limit: 1,
      busy: [],
    });
    expect(slots[0]!.start).toBe("2026-04-20T09:30");
  });

  test("localMidnight keeps local date and zeroes the time", () => {
    const at = new Date(2026, 3, 20, 15, 30, 1, 500);
    const midnight = localMidnight(at);
    expect(midnight.getHours()).toBe(0);
    expect(midnight.getMinutes()).toBe(0);
    expect(midnight.getDate()).toBe(20);
  });
});
