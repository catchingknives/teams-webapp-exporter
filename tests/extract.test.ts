import { describe, it, expect } from "vitest";
import {
  getExtractionScript,
  getScrollDownExtractionScript,
} from "../lib/extract";

describe("getExtractionScript", () => {
  it("returns a string wrapping an async IIFE", () => {
    const script = getExtractionScript(7);
    expect(script).toMatch(/^\(async function\(\)/);
    expect(script).toMatch(/\)\(\)$/);
  });

  it("injects the days parameter correctly", () => {
    const script = getExtractionScript(30);
    expect(script).toContain("var needsScroll = 30 >= 0");
    expect(script).toContain("Date.now() - 30 * 86400000");
  });

  it("disables scrolling when days is -1", () => {
    const script = getExtractionScript(-1);
    expect(script).toContain("var needsScroll = -1 >= 0");
  });

  it("sets cutoffDate to null when days is 0 (all history)", () => {
    const script = getExtractionScript(0);
    expect(script).toContain("var cutoffDate = 0 > 0 ? new Date");
  });

  it("injects sinceTimestamp when provided", () => {
    const ts = "2025-06-01T00:00:00.000Z";
    const script = getExtractionScript(7, ts);
    expect(script).toContain(`new Date("${ts}")`);
  });

  it("sets sinceDate to null when no sinceTimestamp", () => {
    const script = getExtractionScript(7);
    expect(script).toContain("var sinceDate = null");
  });

  it("references key DOM selectors", () => {
    const script = getExtractionScript(7);
    expect(script).toContain("chat-pane-list");
    expect(script).toContain("message-body-");
    expect(script).toContain("timestamp-");
    expect(script).toContain("message-author-name");
  });
});

describe("getScrollDownExtractionScript", () => {
  it("returns a string wrapping an async IIFE", () => {
    const script = getScrollDownExtractionScript();
    expect(script).toMatch(/^\(async function\(\)/);
    expect(script).toMatch(/\)\(\)$/);
  });

  it("references chat-pane-list", () => {
    const script = getScrollDownExtractionScript();
    expect(script).toContain("chat-pane-list");
  });

  it("scrolls down (uses scrollTop +=)", () => {
    const script = getScrollDownExtractionScript();
    expect(script).toContain("scrollTop +=");
  });
});
