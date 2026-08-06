/**
 * Unit tests for excel-commands.ts using the mock framework.
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import {
	type MockWorkbookData,
	installExcelMock,
	getMutations,
	resetMutations,
	setWorkbookData,
} from "./excel-mock";

// ── Mock bridge ──────────────────────────────────────────────────

const mockBridge = {
	reportResult: async (
		_cmd: string,
		_success: boolean,
		_err?: string,
		_payload?: unknown,
	) => {},
};

vi.mock("./communication", () => ({
	reportResult: (
		cmd: string,
		success: boolean,
		err?: string,
		payload?: unknown,
	) => mockBridge.reportResult(cmd, success, err, payload),
	MCP_SERVER_URL: "http://127.0.0.1:3000",
}));

import { processCommand } from "./excel-commands";

// ── Test data ───────────────────────────────────────────────────

function makeTestWorkbook(): MockWorkbookData {
	return {
		sheets: [
			{
				name: "Q3 Forecast",
				values: [
					["Product", "Q1", "Q2", "Q3", "Total"],
					["Widget A", 100, 120, 150, null],
					["Widget B", 200, 180, 210, null],
					["Widget C", 50, 60, 70, null],
				],
				formulas: [
					[null, null, null, null, null],
					[null, null, null, null, "=SUM(B2:D2)"],
					[null, null, null, null, "=SUM(B3:D3)"],
					[null, null, null, null, "=SUM(B4:D4)"],
				],
				tables: [
					{
						name: "SalesData",
						range: "A1:D4",
						columns: ["Product", "Q1", "Q2", "Q3"],
					},
				],
			},
			{
				name: "Summary",
				values: [
					["Category", "Amount"],
					["Revenue", 5000],
					["Expenses", 3000],
				],
			},
		],
		namedRanges: [
			{ name: "TotalRevenue", refersTo: "Summary!$B$2" },
			{ name: "DataRange", refersTo: "'Q3 Forecast'!$A$1:$D$4" },
		],
	};
}

// ── Tests ────────────────────────────────────────────────────────

describe("excel-commands", () => {
	beforeEach(() => {
		installExcelMock();
		setWorkbookData(makeTestWorkbook());
		resetMutations();
	});

	// ── workbook map ────────────────────────────────────────────

	describe("excel_get_workbook_map", () => {
		it("returns sheet names and dimensions", async () => {
			const result = (await processCommand(
				"cmd-1",
				"excel_get_workbook_map",
				{},
			)) as any;

			expect(result.sheetCount).toBe(2);
			expect(result.sheets[0].name).toBe("Q3 Forecast");
			expect(result.sheets[1].name).toBe("Summary");
			expect(result.sheets[0].rowCount).toBeGreaterThanOrEqual(3);
		});

		it("includes tables when requested", async () => {
			const result = (await processCommand("cmd-2", "excel_get_workbook_map", {
				includeTables: true,
			})) as any;

			expect(result.sheets[0].tables).toBeDefined();
			expect(result.sheets[0].tables[0].name).toBe("SalesData");
			expect(result.sheets[0].tables[0].columns).toContain("Product");
		});

		it("includes named ranges when requested", async () => {
			const result = (await processCommand("cmd-3", "excel_get_workbook_map", {
				includeNamedRanges: true,
			})) as any;

			expect(result.namedRanges).toBeDefined();
			expect(result.namedRanges.length).toBe(2);
			expect(result.namedRanges[0].name).toBe("TotalRevenue");
		});

		it("omits tables and named ranges when disabled", async () => {
			const result = (await processCommand("cmd-4", "excel_get_workbook_map", {
				includeTables: false,
				includeNamedRanges: false,
			})) as any;

			expect(result.sheets[0].tables).toBeUndefined();
			expect(result.namedRanges).toBeUndefined();
		});
	});

	// ── read range ──────────────────────────────────────────────

	describe("excel_read_range", () => {
		it("reads values from a range", async () => {
			const result = (await processCommand("cmd-5", "excel_read_range", {
				sheetName: "Q3 Forecast",
				address: "A1:D4",
			})) as any;

			expect(result.values).toBeDefined();
			expect(result.values[0][0]).toBe("Product");
			expect(result.values[1][0]).toBe("Widget A");
			expect(result.rowCount).toBe(4);
			expect(result.columnCount).toBe(4);
		});

		it("reads formulas when includeFormulas is true", async () => {
			const result = (await processCommand("cmd-6", "excel_read_range", {
				sheetName: "Q3 Forecast",
				address: "E2:E4",
				includeFormulas: true,
			})) as any;

			expect(result.formulas).toBeDefined();
			expect(result.formulas[0][0]).toBe("=SUM(B2:D2)");
		});

		it("returns error for missing sheetName", async () => {
			const result = (await processCommand("cmd-7", "excel_read_range", {
				address: "A1:D4",
			})) as any;

			expect(result.error).toBeDefined();
			expect(result.error).toContain("sheetName");
		});

		it("returns error for missing address", async () => {
			const result = (await processCommand("cmd-8", "excel_read_range", {
				sheetName: "Q3 Forecast",
			})) as any;

			expect(result.error).toBeDefined();
			expect(result.error).toContain("address");
		});

		it("reads a single cell", async () => {
			const result = (await processCommand("cmd-9", "excel_read_range", {
				sheetName: "Q3 Forecast",
				address: "B2",
			})) as any;

			expect(result.values).toBeDefined();
			expect(result.values[0][0]).toBe(100);
		});
	});

	// ── write range ─────────────────────────────────────────────

	describe("excel_write_range", () => {
		it("writes values to a range", async () => {
			const result = (await processCommand("cmd-10", "excel_write_range", {
				sheetName: "Summary",
				address: "B2:B3",
				values: [[6000], [3500]],
			})) as any;

			expect(result.written).toBe(true);
			expect(result.rowsWritten).toBe(2);
			expect(result.totalCells).toBe(2);
			expect(result.undoable).toBe(true);
		});

		it("returns diff preview with before/after", async () => {
			const result = (await processCommand("cmd-11", "excel_write_range", {
				sheetName: "Summary",
				address: "B2:B3",
				values: [[6000], [3500]],
			})) as any;

			expect(result.diffPreview).toBeDefined();
			expect(result.diffPreview.length).toBeGreaterThan(0);
			expect(result.diffPreview[0].before).toBeDefined();
			expect(result.diffPreview[0].after).toBeDefined();
		});

		it("returns error for missing values", async () => {
			const result = (await processCommand("cmd-12", "excel_write_range", {
				sheetName: "Summary",
				address: "B2:B3",
			})) as any;

			expect(result.error).toBeDefined();
		});

		it("rejects ranges over 1M cells", async () => {
			const hugeValues = Array(1000).fill(Array(1001).fill(0));
			const result = (await processCommand("cmd-13", "excel_write_range", {
				sheetName: "Summary",
				address: "A1:ALL1000",
				values: hugeValues,
			})) as any;

			expect(result.errorCode).toBe("RANGE_TOO_LARGE");
		});

		it("tracks mutation in mock", async () => {
			await processCommand("cmd-14", "excel_write_range", {
				sheetName: "Summary",
				address: "B2:B3",
				values: [[6000], [3500]],
			});

			const mutations = getMutations();
			expect(mutations.length).toBe(1);
			expect(mutations[0].type).toBe("write_range");
			expect(mutations[0].sheet).toBe("Summary");
		});
	});

	// ── write formula ───────────────────────────────────────────

	describe("excel_write_formula", () => {
		it("writes a valid formula", async () => {
			const result = (await processCommand("cmd-15", "excel_write_formula", {
				sheetName: "Q3 Forecast",
				address: "E2",
				formula: "=SUM(B2:D2)",
			})) as any;

			expect(result.written).toBe(true);
			expect(result.formula).toBe("=SUM(B2:D2)");
			expect(result.undoable).toBe(true);
		});

		it("rejects formula not starting with =", async () => {
			const result = (await processCommand("cmd-16", "excel_write_formula", {
				sheetName: "Q3 Forecast",
				address: "E2",
				formula: "SUM(B2:D2)",
			})) as any;

			expect(result.errorCode).toBe("INVALID_FORMULA");
		});

		it("rejects formula with mismatched parentheses", async () => {
			const result = (await processCommand("cmd-17", "excel_write_formula", {
				sheetName: "Q3 Forecast",
				address: "E2",
				formula: "=SUM(B2:D2",
			})) as any;

			expect(result.errorCode).toBe("INVALID_FORMULA");
		});

		it("returns error for missing formula", async () => {
			const result = (await processCommand("cmd-18", "excel_write_formula", {
				sheetName: "Q3 Forecast",
				address: "E2",
			})) as any;

			expect(result.error).toContain("formula");
		});
	});

	// ── create table ────────────────────────────────────────────

	describe("excel_create_table", () => {
		it("creates a table from a range", async () => {
			const result = (await processCommand("cmd-19", "excel_create_table", {
				sheetName: "Summary",
				address: "A1:B3",
				hasHeaders: true,
			})) as any;

			expect(result.created).toBe(true);
			expect(result.tableName).toBeDefined();
			expect(result.hasHeaders).toBe(true);
			expect(result.undoable).toBe(true);
		});

		it("creates a table with custom name", async () => {
			const result = (await processCommand("cmd-20", "excel_create_table", {
				sheetName: "Summary",
				address: "A1:B3",
				tableName: "MyCustomTable",
				hasHeaders: true,
			})) as any;

			expect(result.created).toBe(true);
			expect(result.tableName).toBe("MyCustomTable");
		});

		it("returns error for missing address", async () => {
			const result = (await processCommand("cmd-21", "excel_create_table", {
				sheetName: "Summary",
			})) as any;

			expect(result.error).toContain("address");
		});

		it("tracks mutation in mock", async () => {
			await processCommand("cmd-22", "excel_create_table", {
				sheetName: "Summary",
				address: "A1:B3",
				hasHeaders: true,
			});

			const mutations = getMutations();
			expect(mutations.length).toBe(1);
			expect(mutations[0].type).toBe("create_table");
		});
	});

	// ── unknown command ─────────────────────────────────────────

	describe("unknown command", () => {
		it("returns error for unknown Excel command", async () => {
			const result = (await processCommand(
				"cmd-23",
				"excel_nonexistent",
				{},
			)) as any;

			expect(result.error).toContain("Unknown Excel command");
		});
	});
});
