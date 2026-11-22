import { describe, expect, test } from "bun:test";
import {
  DASL,
  and,
  bodyContains,
  escapeDasl,
  formatDaslDate,
  mailFilter,
  or,
  receivedSince,
  restrictExpression,
  searchClause,
  senderContains,
  subjectContains,
  unreadOnly,
} from "../src/outlook/dasl.ts";

describe("DASL filters", () => {
  test("dates use local short ISO without a zone", () => {
    const date = new Date(2026, 3, 20, 13, 5, 6);
    expect(formatDaslDate(date)).toBe("2026-04-20T13:05:06");
    expect(receivedSince(date)).toBe(`"${DASL.received}" >= '2026-04-20T13:05:06'`);
  });

  test("an ISO string with a zone is accepted", () => {
    const clause = receivedSince("2026-04-20T13:05:06Z");
    expect(clause).toContain(`"${DASL.received}" >= '`);
    expect(clause).not.toContain("Z'");
  });

  test("unread uses the DASL read property", () => {
    expect(unreadOnly()).toBe(`"${DASL.read}" = 0`);
  });

  test("single quotes are doubled and never close the literal", () => {
    expect(escapeDasl("O'Brien")).toBe("O''Brien");
    expect(subjectContains("O'Brien")).toBe(`"${DASL.subject}" LIKE '%O''Brien%'`);
  });

  test("scope clauses are parenthesised OR groups", () => {
    const clause = searchClause("budget", "all");
    expect(clause.startsWith("(")).toBe(true);
    expect(clause).toContain(DASL.subject);
    expect(clause).toContain(DASL.textDescription);
    expect(clause).toContain(DASL.fromName);
    expect(clause).toContain("OR");
  });

  test("sender searches both name and address", () => {
    const clause = senderContains("example.com");
    expect(clause).toContain(DASL.fromName);
    expect(clause).toContain(DASL.fromEmail);
  });

  test("only DASL property URIs are used", () => {
    const clause = searchClause("x", "all") + mailFilter({ sinceIso: "2026-04-20T00:00:00", unreadOnly: true });
    expect(clause).not.toMatch(/\[(ReceivedTime|Subject|From)\]/);
    expect(clause).toContain("urn:schemas:httpmail:");
  });

  test("filters combine with AND, and empty parts disappear", () => {
    expect(and("a", null, "", false, "b")).toBe("a AND b");
    expect(mailFilter({})).toBe("");
    expect(mailFilter({ unreadOnly: true })).toBe(unreadOnly());
  });

  test("OR with one clause needs no parentheses", () => {
    expect(or("a")).toBe("a");
    expect(or(null, "a")).toBe("a");
    expect(or("a", "b")).toBe("(a OR b)");
  });

  test("the restriction expression adds @SQL once", () => {
    expect(restrictExpression(subjectContains("x")).startsWith("@SQL=")).toBe(true);
  });

  test("an invalid date throws instead of producing a broken filter", () => {
    expect(() => receivedSince("not a date")).toThrow("invalid date");
  });

  test("body search uses textdescription, not an object property", () => {
    expect(bodyContains("x")).toContain(DASL.textDescription);
  });
});
