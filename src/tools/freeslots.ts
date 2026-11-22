/**
 * Free-slot search.
 *
 * `find_freeslots` iterates the calendar and breaks early; it never uses
 * `Items.Restrict` together with `IncludeRecurrences` (that returns nothing).
 * The bridge returns busy intervals and this module does the slot arithmetic,
 * so the logic is testable without Outlook.
 */

export interface Interval {
  start: Date;
  end: Date;
}

export interface Slot {
  start: string;
  end: string;
  weekday: string;
}

export interface BusinessHours {
  startHour: number;
  startMinute: number;
  endHour: number;
  endMinute: number;
}

export interface SlotSearchOptions {
  now: Date;
  daysAhead: number;
  durationMinutes: number;
  businessHours: BusinessHours;
  weekdaysOnly: boolean;
  stepMinutes: number;
  limit: number;
  busy: Interval[];
}

const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

export function parseBusinessHours(value: string): BusinessHours {
  const match = /^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$/.exec(value.trim());
  if (!match) throw new Error(`invalid business_hours: ${value} (expected HH:MM-HH:MM)`);
  const [startHour, startMinute, endHour, endMinute] = match.slice(1).map(Number) as [number, number, number, number];
  if (startHour > 23 || endHour > 23 || startMinute > 59 || endMinute > 59) {
    throw new Error(`invalid business_hours: ${value}`);
  }
  if (endHour * 60 + endMinute <= startHour * 60 + startMinute) {
    throw new Error(`business_hours must end after it starts: ${value}`);
  }
  return { startHour, startMinute, endHour, endMinute };
}

export function formatLocalMinute(date: Date): string {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${date.getFullYear()}-${p(date.getMonth() + 1)}-${p(date.getDate())}T${p(date.getHours())}:${p(date.getMinutes())}`;
}

export function mergeIntervals(intervals: Interval[]): Interval[] {
  const sorted = [...intervals].sort((a, b) => a.start.getTime() - b.start.getTime());
  const merged: Interval[] = [];
  for (const interval of sorted) {
    const last = merged[merged.length - 1];
    if (last && interval.start.getTime() <= last.end.getTime()) {
      if (interval.end.getTime() > last.end.getTime()) last.end = interval.end;
    } else {
      merged.push({ start: interval.start, end: interval.end });
    }
  }
  return merged;
}

/** Decode an Exchange FreeBusy string: `0` free, anything else busy. */
export function freeBusyIntervals(busyString: string, start: Date, slotMinutes: number): Interval[] {
  const intervals: Interval[] = [];
  for (let i = 0; i < busyString.length; i++) {
    if (busyString[i] === "0") continue;
    const slotStart = new Date(start.getTime() + i * slotMinutes * 60_000);
    const slotEnd = new Date(slotStart.getTime() + slotMinutes * 60_000);
    const last = intervals[intervals.length - 1];
    if (last && last.end.getTime() === slotStart.getTime()) last.end = slotEnd;
    else intervals.push({ start: slotStart, end: slotEnd });
  }
  return intervals;
}

function overlaps(busy: Interval[], start: Date, end: Date): boolean {
  return busy.some((interval) => start.getTime() < interval.end.getTime() && end.getTime() > interval.start.getTime());
}

export function findSlots(options: SlotSearchOptions): Slot[] {
  const busy = mergeIntervals(options.busy);
  const slots: Slot[] = [];
  const endWindow = new Date(options.now.getTime() + options.daysAhead * 86_400_000);

  const atBusinessStart = (date: Date): Date => {
    const next = new Date(date);
    next.setHours(options.businessHours.startHour, options.businessHours.startMinute, 0, 0);
    return next;
  };
  const nextDay = (date: Date): Date => {
    const next = new Date(date);
    next.setDate(next.getDate() + 1);
    return atBusinessStart(next);
  };

  let current = new Date(options.now);
  current.setSeconds(0, 0);
  current.setMinutes(current.getMinutes() + ((options.stepMinutes - (current.getMinutes() % options.stepMinutes)) % options.stepMinutes));

  while (current.getTime() < endWindow.getTime() && slots.length < options.limit) {
    if (options.weekdaysOnly && (current.getDay() === 0 || current.getDay() === 6)) {
      current = nextDay(current);
      continue;
    }
    const businessStart = new Date(current);
    businessStart.setHours(options.businessHours.startHour, options.businessHours.startMinute, 0, 0);
    const businessEnd = new Date(current);
    businessEnd.setHours(options.businessHours.endHour, options.businessHours.endMinute, 0, 0);

    if (current.getTime() < businessStart.getTime()) {
      current = businessStart;
      continue;
    }
    const slotEnd = new Date(current.getTime() + options.durationMinutes * 60_000);
    if (slotEnd.getTime() > businessEnd.getTime()) {
      current = nextDay(current);
      continue;
    }
    if (!overlaps(busy, current, slotEnd)) {
      slots.push({
        start: formatLocalMinute(current),
        end: formatLocalMinute(slotEnd),
        weekday: WEEKDAYS[current.getDay()]!,
      });
    }
    current = new Date(current.getTime() + options.stepMinutes * 60_000);
  }
  return slots;
}

export function localMidnight(date: Date): Date {
  const midnight = new Date(date);
  midnight.setHours(0, 0, 0, 0);
  return midnight;
}
