import { describe, expect, test } from "bun:test";
import {
  AllowlistError,
  checkRecipients,
  compileAllowlist,
  extractAddress,
  matchesAllowlist,
} from "../src/config/matcher.ts";

const match = (entries: string[], value: string): boolean => matchesAllowlist(compileAllowlist(entries), value);

describe("recipient allowlist", () => {
  test("an exact address is literal and case-insensitive", () => {
    expect(match(["User.Name@example.com"], "user.name@EXAMPLE.com")).toBe(true);
    expect(match(["user.name@example.com"], "other.name@example.com")).toBe(false);
    expect(match(["user.name@example.com"], "userXname@example.com")).toBe(false);
  });

  test("regex metacharacters in a plain entry stay literal", () => {
    expect(match(["a.b@example.com"], "a.b@example.com")).toBe(true);
    expect(match(["a.b@example.com"], "axb@example.com")).toBe(false);
    expect(match(["user+tag@example.com"], "user+tag@example.com")).toBe(true);
    expect(match(["user+tag@example.com"], "usertag@example.com")).toBe(false);
    expect(match(["user(1)@example.com"], "user(1)@example.com")).toBe(true);
  });

  test("a bare domain matches that domain only", () => {
    expect(match(["example.com"], "anyone@example.com")).toBe(true);
    expect(match(["example.com"], "anyone@mail.example.com")).toBe(false);
    expect(match(["@example.com"], "anyone@example.com")).toBe(true);
    expect(match(["example.com"], "anyone@example.org")).toBe(false);
  });

  test("*@domain matches any local part at that exact domain", () => {
    expect(match(["*@example.com"], "anyone@example.com")).toBe(true);
    expect(match(["*@example.com"], "first.last+tag@example.com")).toBe(true);
    expect(match(["*@example.com"], "anyone@mail.example.com")).toBe(false);
  });

  test("*.domain matches subdomains but not the apex", () => {
    expect(match(["*.example.com"], "anyone@mail.example.com")).toBe(true);
    expect(match(["*.example.com"], "anyone@a.b.example.com")).toBe(true);
    expect(match(["*.example.com"], "anyone@example.com")).toBe(false);
    expect(match(["*.example.com"], "anyone@notexample.com")).toBe(false);
  });

  test("/regex/ is case-insensitive and applies to the whole address", () => {
    expect(match(["/^first\\.last@example\\.com$/"], "First.Last@Example.com")).toBe(true);
    expect(match(["/@example\\.com$/"], "anyone@example.com")).toBe(true);
    expect(match(["/@example\\.com$/"], "anyone@example.org")).toBe(false);
  });

  test("an invalid regex refuses to compile", () => {
    expect(() => compileAllowlist(["/[/"])).toThrow(AllowlistError);
    expect(() => compileAllowlist(["/(/i"])).toThrow(AllowlistError);
    expect(() => compileAllowlist(["/abc"])).toThrow(AllowlistError);
    expect(() => compileAllowlist(["//"])).toThrow(AllowlistError);
  });

  test("invalid wildcard entries are rejected", () => {
    expect(() => compileAllowlist(["*@"])).toThrow(AllowlistError);
    expect(() => compileAllowlist(["*."])).toThrow(AllowlistError);
    expect(() => compileAllowlist(["*@*.example.com"])).toThrow(AllowlistError);
    expect(() => compileAllowlist(["foo bar"])).toThrow(AllowlistError);
  });

  test("an empty list is unrestricted", () => {
    const result = checkRecipients([], ["anyone@example.com", "else@example.org"]);
    expect(result.ok).toBe(true);
    expect(result.unrestricted).toBe(true);
    expect(result.rejected).toEqual([]);
  });

  test("To, Cc and Bcc addresses are all checked", () => {
    const result = checkRecipients(["example.com"], [
      "inside@example.com",
      "outside@example.org",
      "another@example.net",
    ]);
    expect(result.ok).toBe(false);
    expect(result.accepted).toEqual(["inside@example.com"]);
    expect(result.rejected).toEqual(["outside@example.org", "another@example.net"]);
  });

  test("display names are unwrapped before matching", () => {
    expect(extractAddress("Name <user@example.com>")).toBe("user@example.com");
    expect(match(["example.com"], "A Person <user@example.com>")).toBe(true);
    expect(match(["example.com"], "Name only")).toBe(false);
  });

  test("a group entry that matches any address accepts all of them", () => {
    expect(
      checkRecipients(["example.com", "/@contractor\\.example\\.net$/"], [
        "a@example.com",
        "b@contractor.example.net",
      ]).ok,
    ).toBe(true);
  });

  test("case is ignored on both sides", () => {
    expect(match(["EXAMPLE.COM"], "user@Example.Com")).toBe(true);
    expect(match(["User@Example.Com"], "USER@EXAMPLE.COM")).toBe(true);
  });
});
