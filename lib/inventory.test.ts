import { describe, expect, it } from "vitest";
import { analyzeInventory, parseInventoryRows, summarizeResults } from "./inventory";

const headers = ["Product", "Current Stock", "Average Daily Sales", "Supplier Lead Time", "Unit Cost"];
const rows = [
  ["Wireless Mouse", 40, 7, 10, 12],
  ["Mechanical Keyboard", 150, 4, 8, 32],
  ["Webcam", 25, 2, 14, 28],
  ["Computer Monitor", 180, 2, 12, 145],
  ["USB-C Cable", 60, 6, 7, 8]
];

describe("inventory analysis", () => {
  it("matches all expected demonstration classifications", () => {
    const results = analyzeInventory(parseInventoryRows(headers, rows));
    expect(results.map(({ product, risk }) => [product, risk])).toEqual([
      ["Wireless Mouse", "Urgent"],
      ["Mechanical Keyboard", "Healthy"],
      ["Webcam", "Urgent"],
      ["Computer Monitor", "Overstocked"],
      ["USB-C Cable", "Reorder Soon"]
    ]);
    expect(results[0].suggestedOrder).toBe(128);
    expect(summarizeResults(results)).toEqual({
      urgent: 2,
      reorderSoon: 1,
      overstocked: 1,
      urgentOrderCost: 2404
    });
  });

  it("reports a missing required column", () => {
    expect(() => parseInventoryRows(headers.slice(0, 4), rows)).toThrow(
      "Missing required column: Unit Cost"
    );
  });

  it.each([
    [["", 1, 1, 1, 1], "Product must not be blank"],
    [["Bad Stock", "many", 1, 1, 1], "Current Stock must contain a number"],
    [["Negative", -1, 1, 1, 1], "Current Stock must be zero or greater"],
    [["No Sales", 1, 0, 1, 1], "Average Daily Sales must be greater than zero"]
  ])("rejects invalid row %#", (row, message) => {
    expect(() => parseInventoryRows(headers, [row])).toThrow(message as string);
  });

  it("rejects duplicate product names case-insensitively", () => {
    expect(() => parseInventoryRows(headers, [rows[0], ["wireless mouse", 2, 1, 1, 1]])).toThrow(
      "Duplicate product: wireless mouse"
    );
  });
});
