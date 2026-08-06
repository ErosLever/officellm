/**
 * Excel command handler using Office JS API.
 *
 * Key Excel JS API patterns:
 * - Excel.run(async (context) => { ... }) for batched operations
 * - context.workbook.worksheets.getItem(name) for sheet access
 * - sheet.getRange("A1:D10") for range access
 * - range.values — 2D array for reading/writing cell values
 * - range.formulas — 2D array for reading/writing formulas
 * - range.text — 2D array of display strings (read-only)
 * - range.numberFormat — 2D array for number format strings
 * - Must sync() before reading any loaded property
 * - Tables: sheet.tables.add(address, hasHeaders) creates a ListObject
 */

/// <reference types="@types/office-js" />

import { reportResult } from "./communication";

export async function processCommand(
	commandId: string,
	commandName: string,
	args: unknown,
): Promise<unknown> {
	let result: unknown;
	let success = true;

	try {
		switch (commandName) {
			case "excel_get_workbook_map":
				result = await handleGetWorkbookMap(args);
				break;
			case "excel_read_range":
				result = await handleReadRange(args);
				break;
			case "excel_write_range":
				result = await handleWriteRange(args);
				break;
			case "excel_write_formula":
				result = await handleWriteFormula(args);
				break;
			case "excel_create_table":
				result = await handleCreateTable(args);
				break;
			case "excel_add_sheet":
				result = await handleAddSheet(args);
				break;
			case "excel_delete_sheet":
				result = await handleDeleteSheet(args);
				break;
			case "excel_rename_sheet":
				result = await handleRenameSheet(args);
				break;
			case "excel_sort_range":
				result = await handleSortRange(args);
				break;
			case "excel_filter_range":
				result = await handleFilterRange(args);
				break;
			case "excel_create_chart":
				result = await handleCreateChart(args);
				break;
			case "excel_get_charts":
				result = await handleGetCharts(args);
				break;
			case "excel_format_range":
				result = await handleFormatRange(args);
				break;
			case "excel_apply_conditional_formatting":
				result = await handleConditionalFormatting(args);
				break;
			case "excel_create_pivottable":
				result = await handleCreatePivotTable(args);
				break;
			case "excel_freeze_panes":
				result = await handleFreezePanes(args);
				break;
			case "excel_get_named_ranges":
				result = await handleGetNamedRanges(args);
				break;
			case "excel_add_named_range":
				result = await handleAddNamedRange(args);
				break;
			case "excel_add_data_validation":
				result = await handleAddDataValidation(args);
				break;
			case "excel_remove_data_validation":
				result = await handleRemoveDataValidation(args);
				break;
			case "excel_protect_sheet":
				result = await handleProtectSheet(args);
				break;
			case "excel_unprotect_sheet":
				result = await handleUnprotectSheet(args);
				break;
			case "excel_set_page_layout":
				result = await handleSetPageLayout(args);
				break;
			case "excel_get_page_layout":
				result = await handleGetPageLayout(args);
				break;
			default:
				result = { error: `Unknown Excel command: ${commandName}` };
		}

		if (result && typeof result === "object" && "error" in result) {
			success = false;
		}
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		console.error(`Excel command ${commandId} failed:`, errorMessage);
		success = false;
		result = { error: errorMessage };
	}

	await reportResult(commandId, success, undefined, result);
	return result;
}

// ── Helpers ─────────────────────────────────────────────────────

function runInExcel<T>(fn: (ctx: any) => Promise<T>): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const Excel: any = (globalThis as any).Excel;
		if (!Excel || typeof Excel.run !== "function") {
			reject(new Error("Excel.run() not available"));
			return;
		}
		Excel.run(async (ctx: any) => {
			resolve(await fn(ctx));
		}).catch(reject);
	});
}

// ── Read tools ──────────────────────────────────────────────────

async function handleGetWorkbookMap(args: unknown): Promise<unknown> {
	const config = args as {
		includeTables?: boolean;
		includeNamedRanges?: boolean;
	};
	const includeTables = config.includeTables ?? true;
	const includeNamedRanges = config.includeNamedRanges ?? true;

	return runInExcel(async (ctx) => {
		const worksheets = ctx.workbook.worksheets;
		worksheets.load("items");
		await ctx.sync();

		const sheets: Array<{
			name: string;
			position: number;
			visibility: string;
			usedRange: string;
			rowCount: number;
			columnCount: number;
			tables?: Array<{ name: string; range: string; columns: string[] }>;
		}> = [];

		for (let i = 0; i < worksheets.items.length; i++) {
			const ws = worksheets.items[i];
			ws.load("name,position,visibility");
		}
		await ctx.sync();

		for (const ws of worksheets.items) {
			const usedRange = ws.getUsedRange();
			usedRange.load("address,columnCount,rowCount");
			await ctx.sync();

			const sheetInfo: (typeof sheets)[0] = {
				name: ws.name,
				position: ws.position,
				visibility: String(ws.visibility || "Visible"),
				usedRange: usedRange.address
					? String(usedRange.address).split("!").pop() || ""
					: "",
				rowCount: usedRange.rowCount || 0,
				columnCount: usedRange.columnCount || 0,
			};

			if (includeTables) {
				const tables = ws.tables;
				tables.load("items");
				await ctx.sync();

				const tableList: Array<{
					name: string;
					range: string;
					columns: string[];
				}> = [];
				for (const t of tables.items) {
					t.load("name");
					t.range.load("address");
					await ctx.sync();

					const cols = t.columns;
					cols.load("items");
					await ctx.sync();

					const colNames: string[] = [];
					for (const col of cols.items) {
						col.load("name");
					}
					await ctx.sync();
					for (const col of cols.items) {
						colNames.push(String(col.name));
					}

					tableList.push({
						name: String(t.name),
						range: String(t.range.address).split("!").pop() || "",
						columns: colNames,
					});
				}
				sheetInfo.tables = tableList;
			}

			sheets.push(sheetInfo);
		}

		const result: Record<string, unknown> = {
			sheetCount: sheets.length,
			sheets,
		};

		if (includeNamedRanges) {
			const names = ctx.workbook.names;
			names.load("items");
			await ctx.sync();

			const namedRanges: Array<{ name: string; refersTo: string }> = [];
			for (const nr of names.items) {
				nr.load("name,refersTo");
			}
			await ctx.sync();
			for (const nr of names.items) {
				namedRanges.push({
					name: String(nr.name),
					refersTo: String(nr.refersTo),
				});
			}
			(result as any).namedRanges = namedRanges;
		}

		return result;
	});
}

async function handleReadRange(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		includeFormulas?: boolean;
		includeNumberFormats?: boolean;
	};
	const {
		sheetName = "",
		address = "",
		includeFormulas = true,
		includeNumberFormats = false,
	} = config;

	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		const propsToLoad = [
			"values",
			"text",
			"rowCount",
			"columnCount",
			"address",
		];
		if (includeFormulas) propsToLoad.push("formulas");
		if (includeNumberFormats) propsToLoad.push("numberFormat");

		range.load(propsToLoad.join(","));
		await ctx.sync();

		const rowCount = range.rowCount as number;
		const columnCount = range.columnCount as number;
		const totalCells = rowCount * columnCount;

		// Size check: 256KB max (~250K characters)
		const values = range.values as (string | number | boolean | null)[][];
		const serializedSize = JSON.stringify(values).length;
		const maxBytes = 256 * 1024;

		const result: Record<string, unknown> = {
			sheetName,
			address: String(range.address).split("!").pop() || address,
			rowCount,
			columnCount,
			totalCells,
			values,
			text: range.text,
		};

		if (includeFormulas) {
			result.formulas = range.formulas;
		}

		if (includeNumberFormats) {
			result.numberFormats = range.numberFormat;
		}

		if (serializedSize > maxBytes) {
			result.truncated = true;
			result.byteCount = serializedSize;
		}

		return result;
	});
}

// ── Write tools ─────────────────────────────────────────────────

async function handleWriteRange(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		values?: (string | number | boolean | null)[][];
	};
	const { sheetName = "", address = "", values = [] } = config;

	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };
	if (!values || values.length === 0)
		return { error: "values must be a non-empty 2D array" };

	// Range size check: reject > 1M cells
	const totalCells = values.reduce((acc, row) => acc + row.length, 0);
	if (totalCells > 1_000_000) {
		return {
			error: `Range too large: ${totalCells} cells. Maximum is 1,000,000.`,
			errorCode: "RANGE_TOO_LARGE",
		};
	}

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		// Read before-values for diff preview
		range.load("values");
		await ctx.sync();

		const beforeValues = range.values as (string | number | boolean | null)[][];

		// Write new values
		range.values = values;
		await ctx.sync();

		// Build diff preview (sample first 5 rows)
		const sampleRows = Math.min(values.length, 5);
		const diff: Array<{
			row: number;
			col: number;
			before: string;
			after: string;
		}> = [];

		for (let r = 0; r < sampleRows; r++) {
			for (let c = 0; c < values[r].length; c++) {
				const before = beforeValues?.[r]?.[c] ?? null;
				const after = values[r][c];
				if (String(before) !== String(after)) {
					diff.push({
						row: r,
						col: c,
						before: String(before ?? ""),
						after: String(after ?? ""),
					});
				}
			}
		}

		return {
			sheetName,
			address,
			rowsWritten: values.length,
			columnsWritten: values[0]?.length || 0,
			totalCells: values.reduce((acc, row) => acc + row.length, 0),
			written: true,
			diffPreview: diff,
			undoable: true,
		};
	});
}

async function handleWriteFormula(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		formula?: string;
	};
	const { sheetName = "", address = "", formula = "" } = config;

	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };
	if (!formula) return { error: "formula is required" };

	// Formula validation: must start with '='
	if (!formula.startsWith("=")) {
		return {
			error: `Invalid formula: must start with '='. Got: '${formula}'`,
			errorCode: "INVALID_FORMULA",
		};
	}

	// Basic syntax check: check for obvious mismatched parentheses
	let depth = 0;
	for (const ch of formula) {
		if (ch === "(") depth++;
		if (ch === ")") depth--;
		if (depth < 0) {
			return {
				error: `Invalid formula: unmatched closing parenthesis in '${formula}'`,
				errorCode: "INVALID_FORMULA",
			};
		}
	}
	if (depth !== 0) {
		return {
			error: `Invalid formula: unmatched opening parenthesis in '${formula}'`,
			errorCode: "INVALID_FORMULA",
		};
	}

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		// Build the formula array — one formula per row in the range
		range.load("rowCount,columnCount");
		await ctx.sync();

		const rowCount = range.rowCount as number;
		const colCount = range.columnCount as number;

		// If single cell, wrap formula; if range, repeat formula across cells
		const formulaArray: string[][] = [];
		for (let r = 0; r < rowCount; r++) {
			const row: string[] = [];
			for (let c = 0; c < colCount; c++) {
				row.push(formula);
			}
			formulaArray.push(row);
		}

		range.formulas = formulaArray;
		await ctx.sync();

		return {
			sheetName,
			address,
			formula,
			cellsWritten: rowCount * colCount,
			written: true,
			undoable: true,
		};
	});
}

async function handleCreateTable(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		tableName?: string;
		hasHeaders?: boolean;
	};
	const { sheetName = "", address = "", tableName, hasHeaders = true } = config;

	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const table = sheet.tables.add(address, hasHeaders);

		table.load("name");
		table.range.load("address");
		await ctx.sync();

		let finalName = String(table.name);

		// Rename if custom name provided
		if (tableName) {
			table.name = tableName;
			await ctx.sync();
			finalName = tableName;
		}

		return {
			sheetName,
			address,
			tableName: finalName,
			hasHeaders,
			created: true,
			undoable: true,
		};
	});
}

async function handleAddSheet(args: unknown): Promise<unknown> {
	const config = args as { name?: string; position?: number };
	return runInExcel(async (ctx) => {
		const ws = ctx.workbook.worksheets.add(config.name);
		if (config.position !== undefined) ws.position = config.position;
		ws.load("name,position");
		await ctx.sync();
		return {
			name: String(ws.name),
			position: ws.position,
			created: true,
			undoable: true,
		};
	});
}

async function handleDeleteSheet(args: unknown): Promise<unknown> {
	const config = args as { sheetName?: string };
	const { sheetName = "" } = config;
	if (!sheetName) return { error: "sheetName is required" };

	return runInExcel(async (ctx) => {
		const worksheets = ctx.workbook.worksheets;
		worksheets.load("items");
		await ctx.sync();
		if (worksheets.items.length <= 1) {
			return {
				error: "Cannot delete the last worksheet",
				errorCode: "CANNOT_DELETE_LAST_SHEET",
			};
		}
		const sheet = worksheets.getItem(sheetName);
		sheet.delete();
		await ctx.sync();
		return { sheetName, deleted: true, undoable: true };
	});
}

async function handleRenameSheet(args: unknown): Promise<unknown> {
	const config = args as { sheetName?: string; newName?: string };
	const { sheetName = "", newName = "" } = config;
	if (!sheetName) return { error: "sheetName is required" };
	if (!newName) return { error: "newName is required" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		sheet.name = newName;
		await ctx.sync();
		return { oldName: sheetName, newName, renamed: true, undoable: true };
	});
}

async function handleSortRange(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		criteria?: Array<{ column: number; ascending: boolean }>;
		hasHeader?: boolean;
	};
	const {
		sheetName = "",
		address = "",
		criteria = [],
		hasHeader = true,
	} = config;
	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };
	if (!criteria.length) return { error: "criteria must be non-empty" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);
		range.load("rowCount");
		await ctx.sync();
		if (range.rowCount <= 1)
			return { error: "Range too small to sort", errorCode: "RANGE_TOO_SMALL" };

		const sort = range.sort;
		const sortFields = criteria.map((c) => ({
			key: c.column,
			ascending: c.ascending,
		}));
		sort.apply({ fields: sortFields, hasHeader, matchCase: false });
		await ctx.sync();
		return {
			sheetName,
			address,
			sortedBy: criteria.length,
			criteria,
			undoable: true,
		};
	});
}

async function handleFilterRange(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		column?: number;
		criteria?: Record<string, unknown>;
		clearFilters?: boolean;
	};
	const {
		sheetName = "",
		address = "",
		column = 0,
		criteria,
		clearFilters = false,
	} = config;
	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		if (clearFilters) {
			range.filter?.clearAllFilters();
			await ctx.sync();
			return { sheetName, address, filtersCleared: true, undoable: true };
		}

		const filter = range.filter;
		if (!filter) {
			return {
				error: "No autofilter on this range. Apply filter first.",
				errorCode: "NO_AUTOFILTER",
			};
		}

		const filterCriteria = criteria || {};
		const value = (filterCriteria as any).value;
		const operator = (filterCriteria as any).operator;

		if (operator) {
			filter.apply({
				criteria: [{ filterOn: "custom", operator, values: [String(value)] }],
				columnIndex: column,
			});
		} else if (value !== undefined) {
			filter.apply({
				criteria: [{ filterOn: "value", values: [String(value)] }],
				columnIndex: column,
			});
		}
		await ctx.sync();
		return { sheetName, address, column, filtered: true, undoable: true };
	});
}

async function handleCreateChart(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		dataRange?: string;
		chartType?: string;
		title?: string;
	};
	const {
		sheetName = "",
		dataRange = "",
		chartType = "Column",
		title,
	} = config;
	if (!sheetName) return { error: "sheetName is required" };
	if (!dataRange) return { error: "dataRange is required" };

	const supportedTypes = [
		"Column",
		"Bar",
		"Line",
		"Pie",
		"Scatter",
		"Area",
		"Doughnut",
	];
	if (!supportedTypes.includes(chartType)) {
		return {
			error: `Unsupported chart type: ${chartType}. Supported: ${supportedTypes.join(", ")}`,
			errorCode: "UNSUPPORTED_CHART_TYPE",
		};
	}

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const dataRng = sheet.getRange(dataRange);
		dataRng.load("columnCount,rowCount");
		await ctx.sync();

		// Position chart to the right of data
		const chart = sheet.charts.add(chartType, dataRng, "Auto");
		chart.load("name");
		chart.load(["left", "top", "width", "height"]);
		await ctx.sync();

		// Auto-position: right of data range
		const dataEndCol = dataRng.columnCount;
		chart.left = dataEndCol * 80; // ~80px per column
		chart.top = 0;

		if (title) chart.title = { text: title };
		await ctx.sync();

		return {
			sheetName,
			dataRange,
			chartType,
			title: title || null,
			chartName: String(chart.name),
			created: true,
			undoable: true,
		};
	});
}

async function handleGetCharts(args: unknown): Promise<unknown> {
	const config = args as { sheetName?: string };
	const { sheetName } = config;

	return runInExcel(async (ctx) => {
		if (!sheetName) {
			(ctx.workbook.worksheets as any).load("items");
			await ctx.sync();
		}

		const charts: Array<{
			sheetName: string;
			chartType: string;
			title: string | null;
			dataRange: string;
			position: { top: number; left: number; width: number; height: number };
		}> = [];

		const sheets = sheetName
			? [ctx.workbook.worksheets.getItem(sheetName)]
			: (ctx.workbook.worksheets as any).items || [];
		for (const ws of sheets) {
			ws.load("name");
			ws.charts.load("items");
			await ctx.sync();

			for (const chart of ws.charts.items) {
				chart.load(["chartType", "name"]);
				chart.load(["left", "top", "width", "height"]);
				const dataRng = chart.getDataRange();
				dataRng.load("address");
				await ctx.sync();

				let chartTitle: string | null = null;
				try {
					chart.title.load("text");
					await ctx.sync();
					chartTitle = chart.title.text || null;
				} catch {
					chartTitle = null;
				}

				charts.push({
					sheetName: String(ws.name),
					chartType: String(chart.chartType),
					title: chartTitle,
					dataRange: String(dataRng.address).split("!").pop() || "",
					position: {
						top: chart.top,
						left: chart.left,
						width: chart.width,
						height: chart.height,
					},
				});
			}
		}

		return { totalCharts: charts.length, charts };
	});
}

async function handleFormatRange(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		font?: Record<string, unknown>;
		fill?: Record<string, unknown>;
		borders?: Record<string, unknown>;
		alignment?: Record<string, unknown>;
		numberFormat?: string;
	};
	const {
		sheetName = "",
		address = "",
		font,
		fill,
		borders,
		alignment,
		numberFormat,
	} = config;
	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		if (font) {
			const rf = range.format.font;
			if (font.name) rf.name = String(font.name);
			if (font.size) rf.size = Number(font.size);
			if (font.bold !== undefined) rf.bold = Boolean(font.bold);
			if (font.italic !== undefined) rf.italic = Boolean(font.italic);
			if (font.color) rf.color = String(font.color);
		}
		if (fill) {
			if (fill.color) range.format.fill.color = String(fill.color);
		}
		if (borders) {
			const bs = range.format.borders;
			const style = String(borders.style || "thin");
			const color = String(borders.color || "black");
			bs.getItem("EdgeTop").style = style;
			bs.getItem("EdgeTop").color = color;
			bs.getItem("EdgeBottom").style = style;
			bs.getItem("EdgeBottom").color = color;
			bs.getItem("EdgeLeft").style = style;
			bs.getItem("EdgeLeft").color = color;
			bs.getItem("EdgeRight").style = style;
			bs.getItem("EdgeRight").color = color;
		}
		if (alignment) {
			const ra = range.format.alignment;
			if (alignment.horizontal) ra.horizontal = String(alignment.horizontal);
			if (alignment.vertical) ra.vertical = String(alignment.vertical);
			if (alignment.wrapText !== undefined)
				ra.wrapText = Boolean(alignment.wrapText);
		}
		if (numberFormat) range.numberFormat = [[numberFormat]];

		await ctx.sync();
		return {
			sheetName,
			address,
			formatted: true,
			applied: {
				font: !!font,
				fill: !!fill,
				borders: !!borders,
				alignment: !!alignment,
				numberFormat: !!numberFormat,
			},
			undoable: true,
		};
	});
}

async function handleConditionalFormatting(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName?: string;
		address?: string;
		ruleType?: string;
		operator?: string;
		value?: string;
		format?: Record<string, unknown>;
		minColor?: string;
		maxColor?: string;
		iconSet?: string;
	};
	const {
		sheetName = "",
		address = "",
		ruleType = "",
		operator,
		value,
		format,
		minColor,
		maxColor,
		iconSet,
	} = config;
	if (!sheetName) return { error: "sheetName is required" };
	if (!address) return { error: "address is required" };
	if (!ruleType) return { error: "ruleType is required" };

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);
		const cf = range.conditionalFormats;

		if (ruleType === "cellValue") {
			const rule = cf.add(Excel.ConditionalFormatType.cellValue);
			rule.cellValue.format.fill.color = String(format?.fillColor || "red");
			rule.cellValue.rule = {
				operator: operator || "greaterThan",
				formula1: value || "0",
			};
		} else if (ruleType === "dataBar") {
			const rule = cf.add(Excel.ConditionalFormatType.dataBar);
			rule.dataBar.lowerBoundType = "Auto";
			rule.dataBar.upperBoundType = "Auto";
		} else if (ruleType === "colorScale") {
			const rule = cf.add(Excel.ConditionalFormatType.colorScale);
			rule.colorScale.criteria = [
				{ type: "min", color: minColor || "green" },
				{ type: "max", color: maxColor || "red" },
			];
		} else if (ruleType === "iconSet") {
			const rule = cf.add(Excel.ConditionalFormatType.iconSet);
			rule.iconSet.style = iconSet || "3TrafficLights";
		} else {
			return {
				error: `Unknown ruleType: ${ruleType}. Supported: cellValue, dataBar, colorScale, iconSet.`,
				errorCode: "INVALID_PARAMETER",
			};
		}

		await ctx.sync();
		return { sheetName, address, ruleType, applied: true, undoable: true };
	});
}

async function handleCreatePivotTable(args: unknown): Promise<unknown> {
	const config = args as {
		sourceRange?: string;
		name?: string;
		rows?: string[];
		columns?: string[];
		values?: Array<{ field: string; aggregation: string }>;
		destinationSheet?: string;
	};
	const {
		sourceRange = "",
		name,
		rows = [],
		columns = [],
		values = [],
		destinationSheet,
	} = config;
	if (!sourceRange) return { error: "sourceRange is required" };
	if (!rows.length) return { error: "rows must be non-empty" };
	if (!values.length) return { error: "values must be non-empty" };

	return runInExcel(async (ctx) => {
		// Parse source range (may include sheet name)
		const parts = sourceRange.split("!");
		const srcSheetName = parts.length > 1 ? parts[0] : undefined;
		const srcAddress = parts.length > 1 ? parts[1] : sourceRange;

		const srcSheet = srcSheetName
			? ctx.workbook.worksheets.getItem(srcSheetName)
			: ctx.workbook.worksheets.getActiveWorksheet();
		const srcRange = srcSheet.getRange(srcAddress);

		// Create destination sheet
		const destName = destinationSheet || "PivotTable";
		const destSheet = ctx.workbook.worksheets.add(destName);

		// Create pivot table
		const pivotTables = destSheet.pivotTables;
		const pt = pivotTables.add(
			name || "PivotTable1",
			srcRange,
			destSheet.getRange("A1"),
		);

		// Add row fields
		for (const rowField of rows) {
			pt.rowHierarchicalProperties.add(rowField);
		}

		// Add column fields
		for (const colField of columns || []) {
			pt.columnHierarchicalProperties.add(colField);
		}

		// Add value fields
		for (const vf of values) {
			pt.dataHierarchicalProperties.add(vf.field, vf.aggregation || "sum");
		}

		await ctx.sync();
		return {
			sourceRange,
			destinationSheet: destName,
			rows,
			columns,
			values,
			created: true,
			undoable: true,
		};
	});
}

// ── Phase 16: Excel Navigation ───────────────────────────────────

async function handleFreezePanes(args: unknown): Promise<unknown> {
	const config = args as { sheetName: string; at: string; action?: string };
	const { sheetName, at, action = "freeze" } = config;

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);

		if (action === "unfreeze") {
			sheet.freezePanes.unfreeze();
			await ctx.sync();
			return { unfrozen: true, sheetName };
		}

		// Freeze at the specified cell
		const freezeRange = sheet.getRange(at);
		sheet.freezePanes.freezeAt(freezeRange);
		await ctx.sync();

		return { frozen: true, at, sheetName, undoable: true };
	});
}

async function handleGetNamedRanges(_args: unknown): Promise<unknown> {
	return runInExcel(async (ctx) => {
		const names = ctx.workbook.names;
		names.load("items");
		await ctx.sync();

		const ranges: any[] = [];
		for (const item of names.items) {
			item.load(["name", "comment", "scope"]);
			const ref = item.getRangeOrNullObject();
			ref.load(["address"]);
		}
		await ctx.sync();

		for (const item of names.items) {
			const ref = item.getRangeOrNullObject();
			ranges.push({
				name: item.name,
				comment: item.comment || "",
				scope: item.scope || "workbook",
				address: ref.isNullObject ? null : ref.address,
			});
		}

		return { namedRanges: ranges, count: ranges.length };
	});
}

async function handleAddNamedRange(args: unknown): Promise<unknown> {
	const config = args as {
		name: string;
		sheetName: string;
		address: string;
		comment?: string;
	};
	const { name, sheetName, address, comment } = config;

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);

		ctx.workbook.names.add(name, range, comment || "");
		await ctx.sync();

		return {
			name,
			sheetName,
			address,
			comment: comment || "",
			created: true,
			undoable: true,
		};
	});
}

// ── Phase 17: Excel Data Validation ──────────────────────────────

async function handleAddDataValidation(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName: string;
		address: string;
		type: string;
		operator?: string;
		formula1?: string;
		formula2?: string;
		showInputMessage?: boolean;
		inputTitle?: string;
		inputMessage?: string;
		showErrorMessage?: boolean;
		errorTitle?: string;
		errorMessage?: string;
		errorStyle?: string;
	};
	const {
		sheetName,
		address,
		type,
		operator = "between",
		formula1,
		formula2,
		showInputMessage = true,
		inputTitle,
		inputMessage,
		showErrorMessage = true,
		errorTitle,
		errorMessage,
		errorStyle = "stop",
	} = config;

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);
		const dv = range.dataValidation;

		// Map type to Excel.DataValidationType
		const typeMap: Record<string, string> = {
			list: "list",
			wholeNumber: "wholeNumber",
			decimal: "decimal",
			date: "date",
			textLength: "textLength",
			custom: "custom",
		};

		dv.type = typeMap[type] || type;

		// Set rule based on type
		if (type === "list") {
			dv.rule = { list: { inCellDropDown: true, formula1: formula1 || "" } };
		} else if (type === "custom") {
			dv.rule = { custom: { formula1: formula1 || "" } };
		} else {
			// Number/date/textLength validation with operator
			const rule: any = { operator };
			if (formula1 !== undefined) rule.formula1 = formula1;
			if (formula2 !== undefined) rule.formula2 = formula2;
			dv.rule = rule;
		}

		// Input message
		if (showInputMessage) {
			dv.prompt = {
				show: true,
				title: inputTitle || "",
				message: inputMessage || "",
			};
		}

		// Error alert
		if (showErrorMessage) {
			dv.errorAlert = {
				show: true,
				style: errorStyle,
				title: errorTitle || "Invalid input",
				message: errorMessage || "Please enter a valid value.",
			};
		}

		await ctx.sync();
		return {
			sheetName,
			address,
			type,
			operator,
			applied: true,
			undoable: true,
		};
	});
}

async function handleRemoveDataValidation(args: unknown): Promise<unknown> {
	const config = args as { sheetName: string; address: string };
	const { sheetName, address } = config;

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const range = sheet.getRange(address);
		range.dataValidation.clear();
		await ctx.sync();

		return { sheetName, address, removed: true, undoable: true };
	});
}

// ── Phase 18: Excel Protection ───────────────────────────────────

async function handleProtectSheet(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName: string;
		password?: string;
		allowSort?: boolean;
		allowAutoFilter?: boolean;
		allowInsertRows?: boolean;
		allowInsertColumns?: boolean;
		allowDeleteRows?: boolean;
		allowDeleteColumns?: boolean;
		allowFormatCells?: boolean;
		allowPivotTables?: boolean;
	};
	const {
		sheetName,
		password,
		allowSort,
		allowAutoFilter,
		allowInsertRows,
		allowInsertColumns,
		allowDeleteRows,
		allowDeleteColumns,
		allowFormatCells,
		allowPivotTables,
	} = config;

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const options: any = {};
		if (allowSort !== undefined) options.allowSort = allowSort;
		if (allowAutoFilter !== undefined)
			options.allowAutoFilter = allowAutoFilter;
		if (allowInsertRows !== undefined)
			options.allowInsertRows = allowInsertRows;
		if (allowInsertColumns !== undefined)
			options.allowInsertColumns = allowInsertColumns;
		if (allowDeleteRows !== undefined)
			options.allowDeleteRows = allowDeleteRows;
		if (allowDeleteColumns !== undefined)
			options.allowDeleteColumns = allowDeleteColumns;
		if (allowFormatCells !== undefined)
			options.allowFormatCells = allowFormatCells;
		if (allowPivotTables !== undefined)
			options.allowPivotTables = allowPivotTables;
		sheet.protection.protect(options, password || "");
		await ctx.sync();
		return { sheetName, protected: true };
	});
}

async function handleUnprotectSheet(args: unknown): Promise<unknown> {
	const config = args as { sheetName: string; password?: string };
	const { sheetName, password } = config;

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		sheet.protection.unprotect(password || "");
		await ctx.sync();
		return { sheetName, unprotected: true };
	});
}

// ── Phase 18: Excel Page Layout ──────────────────────────────────

async function handleSetPageLayout(args: unknown): Promise<unknown> {
	const config = args as {
		sheetName: string;
		orientation?: string;
		paperSize?: string;
		printArea?: string;
		printTitleRows?: string;
		centerHorizontally?: boolean;
		centerVertically?: boolean;
	};
	const {
		sheetName,
		orientation,
		paperSize,
		printArea,
		printTitleRows,
		centerHorizontally,
		centerVertically,
	} = config;

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const layout = sheet.pageLayout;

		if (orientation) layout.orientation = orientation;
		if (paperSize) layout.paperSize = paperSize;
		if (printArea) layout.setPrintArea(printArea);
		if (printTitleRows) layout.setPrintTitleRows(printTitleRows);
		if (centerHorizontally !== undefined)
			layout.centerHorizontally = centerHorizontally;
		if (centerVertically !== undefined)
			layout.centerVertically = centerVertically;

		await ctx.sync();
		return { sheetName, set: true, undoable: true };
	});
}

async function handleGetPageLayout(args: unknown): Promise<unknown> {
	const config = args as { sheetName: string };
	const { sheetName } = config;

	return runInExcel(async (ctx) => {
		const sheet = ctx.workbook.worksheets.getItem(sheetName);
		const layout = sheet.pageLayout;
		layout.load([
			"orientation",
			"paperSize",
			"centerHorizontally",
			"centerVertically",
		]);
		await ctx.sync();

		return {
			sheetName,
			orientation: layout.orientation,
			paperSize: layout.paperSize,
			centerHorizontally: layout.centerHorizontally,
			centerVertically: layout.centerVertically,
		};
	});
}
