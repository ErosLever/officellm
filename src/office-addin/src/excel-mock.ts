/**
 * Mock framework for Excel JS API — mirrors the real Excel.run() pattern
 * for unit testing excel-commands.ts without a live Excel instance.
 *
 * Architecture: Same pattern as word-mock.ts and powerpoint-mock.ts.
 * - MockWorkbookData defines test data
 * - installExcelMock() patches window.Excel with a mock
 * - Mock tracks all mutations for assertions
 */

// ── Data types ──────────────────────────────────────────────────

export interface MockSheetData {
	name: string;
	/** 2D array of cell values (rows × columns). null = empty cell */
	values: (string | number | boolean | null)[][];
	/** 2D array of formula strings. null = no formula */
	formulas?: (string | null)[][];
	/** 2D array of number format strings. null = default */
	numberFormats?: (string | null)[][];
	/** Tables on this sheet */
	tables?: MockTableData[];
}

export interface MockTableData {
	name: string;
	/** Range address, e.g. "A1:D5" */
	range: string;
	/** Column headers */
	columns: string[];
}

export interface MockNamedRange {
	name: string;
	/** Sheet-scoped reference, e.g. "Sheet1!$A$1:$D$10" */
	refersTo: string;
}

export interface MockWorkbookData {
	sheets: MockSheetData[];
	namedRanges?: MockNamedRange[];
}

// ── Mock state ──────────────────────────────────────────────────

let _workbookData: MockWorkbookData = { sheets: [] };
let _mutations: Array<{ type: string; sheet: string; details: unknown }> = [];

export function getMutations() {
	return _mutations;
}

export function resetMutations() {
	_mutations = [];
}

export function setWorkbookData(data: MockWorkbookData) {
	_workbookData = data;
	_mutations = [];
}

export function getWorkbookData(): MockWorkbookData {
	return _workbookData;
}

// ── Helpers ─────────────────────────────────────────────────────

function parseAddress(address: string): {
	startRow: number;
	startCol: number;
	endRow: number;
	endCol: number;
} {
	// Parse A1-style address like "B2:D10"
	const match = address.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
	if (!match) throw new Error(`Invalid address: ${address}`);

	const colLetterToNum = (col: string) => {
		let n = 0;
		for (let i = 0; i < col.length; i++) {
			n = n * 26 + (col.charCodeAt(i) - 64);
		}
		return n - 1; // zero-based
	};

	const startCol = colLetterToNum(match[1].toUpperCase());
	const startRow = parseInt(match[2]) - 1; // zero-based
	const endCol = match[3] ? colLetterToNum(match[3].toUpperCase()) : startCol;
	const endRow = match[4] ? parseInt(match[4]) - 1 : startRow;

	return { startRow, startCol, endRow, endCol };
}

// ── Mock context ────────────────────────────────────────────────

function createMockContext() {
	const mockWorksheet = (sheetData: MockSheetData, index: number) => ({
		name: sheetData.name,
		position: index,
		visibility: "Visible",
		load(_props: string) {
			// In mock, properties are immediately available
		},

		getRange(address: string) {
			const { startRow, startCol, endRow, endCol } = parseAddress(address);
			const rangeObj = {
				address: `${sheetData.name}!${address}`,
				_loads: [] as string[],
				_loaded: {} as Record<string, unknown>,

				load(props: string) {
					this._loads.push(props);
				},

				get values() {
					const rows: (string | number | boolean | null)[][] = [];
					for (let r = startRow; r <= endRow; r++) {
						const row: (string | number | boolean | null)[] = [];
						for (let c = startCol; c <= endCol; c++) {
							row.push(sheetData.values?.[r]?.[c] ?? null);
						}
						rows.push(row);
					}
					return rows;
				},

				set values(newValues: (string | number | boolean | null)[][]) {
					_mutations.push({
						type: "write_range",
						sheet: sheetData.name,
						details: { address, values: newValues },
					});
					for (let r = 0; r < newValues.length; r++) {
						if (!sheetData.values[startRow + r])
							sheetData.values[startRow + r] = [];
						for (let c = 0; c < newValues[r].length; c++) {
							sheetData.values[startRow + r][startCol + c] = newValues[r][c];
						}
					}
				},

				get formulas() {
					const rows: (string | null)[][] = [];
					for (let r = startRow; r <= endRow; r++) {
						const row: (string | null)[] = [];
						for (let c = startCol; c <= endCol; c++) {
							const formula = sheetData.formulas?.[r]?.[c];
							if (formula) {
								row.push(formula);
							} else {
								row.push(null);
							}
						}
						rows.push(row);
					}
					return rows;
				},

				set formulas(newFormulas: (string | null)[][]) {
					_mutations.push({
						type: "write_formula",
						sheet: sheetData.name,
						details: { address, formulas: newFormulas },
					});
					if (!sheetData.formulas) sheetData.formulas = [];
					for (let r = 0; r < newFormulas.length; r++) {
						if (!sheetData.formulas[startRow + r])
							sheetData.formulas[startRow + r] = [];
						for (let c = 0; c < newFormulas[r].length; c++) {
							sheetData.formulas[startRow + r][startCol + c] =
								newFormulas[r][c];
						}
					}
				},

				get text() {
					const rows: string[][] = [];
					for (let r = startRow; r <= endRow; r++) {
						const row: string[] = [];
						for (let c = startCol; c <= endCol; c++) {
							const v = sheetData.values?.[r]?.[c];
							row.push(v != null ? String(v) : "");
						}
						rows.push(row);
					}
					return rows;
				},

				get numberFormat() {
					const rows: (string | null)[][] = [];
					for (let r = startRow; r <= endRow; r++) {
						const row: (string | null)[] = [];
						for (let c = startCol; c <= endCol; c++) {
							row.push(sheetData.numberFormats?.[r]?.[c] ?? null);
						}
						rows.push(row);
					}
					return rows;
				},

				set numberFormat(formats: (string | null)[][]) {
					if (!sheetData.numberFormats) sheetData.numberFormats = [];
					for (let r = 0; r < formats.length; r++) {
						if (!sheetData.numberFormats[startRow + r])
							sheetData.numberFormats[startRow + r] = [];
						for (let c = 0; c < formats[r].length; c++) {
							sheetData.numberFormats[startRow + r][startCol + c] =
								formats[r][c];
						}
					}
				},

				get rowCount() {
					return endRow - startRow + 1;
				},

				get columnCount() {
					return endCol - startCol + 1;
				},

				get cellCount() {
					return this.rowCount * this.columnCount;
				},
			};
			return rangeObj;
		},

		getUsedRange() {
			const maxRow = sheetData.values.length - 1;
			const maxCol = Math.max(...sheetData.values.map((r) => r.length)) - 1;
			return {
				address: `${sheetData.name}!A1:${String.fromCharCode(65 + maxCol)}${maxRow + 1}`,
				rowCount: maxRow + 1,
				columnCount: maxCol + 1,
				load() {},
			};
		},

		get tables() {
			return {
				items: (sheetData.tables || []).map((t) => ({
					name: t.name,
					range: { address: `${sheetData.name}!${t.range}`, load() {} },
					columns: {
						items: t.columns.map((c, i) => ({
							name: c,
							nameIndex: i,
							load() {},
						})),
						load() {},
					},
					load() {},
				})),
				load() {},
				add(address: string, hasHeaders: boolean) {
					const tableName = `Table${
						_workbookData.sheets.reduce(
							(acc, s) => acc + (s.tables?.length || 0),
							0,
						) + 1
					}`;
					_mutations.push({
						type: "create_table",
						sheet: sheetData.name,
						details: { address, hasHeaders, tableName },
					});
					if (!sheetData.tables) sheetData.tables = [];
					sheetData.tables.push({
						name: tableName,
						range: address,
						columns: [],
					});
					return {
						name: tableName,
						range: { address: `${sheetData.name}!${address}`, load() {} },
						load() {},
					};
				},
			};
		},
	});

	return {
		workbook: {
			worksheets: {
				items: _workbookData.sheets.map((s, i) => mockWorksheet(s, i)),
				getCount() {
					return { value: _workbookData.sheets.length, load() {} };
				},
				getItem(name: string) {
					const sheet = _workbookData.sheets.find(
						(s) => s.name.toLowerCase() === name.toLowerCase(),
					);
					if (!sheet) throw new Error(`Sheet '${name}' not found`);
					const idx = _workbookData.sheets.indexOf(sheet);
					return mockWorksheet(sheet, idx);
				},
				getActiveWorksheet() {
					return mockWorksheet(_workbookData.sheets[0], 0);
				},
				load() {},
			},
			tables: {
				items: _workbookData.sheets.flatMap(
					(s) =>
						s.tables?.map((t) => ({
							name: t.name,
							range: { address: `${s.name}!${t.range}` },
							worksheet: s.name,
							load() {},
						})) || [],
				),
				load() {},
			},
			names: {
				items: (_workbookData.namedRanges || []).map((nr) => ({
					name: nr.name,
					refersTo: nr.refersTo,
					load() {},
				})),
				load() {},
			},
			getSelectedRange() {
				// Return first cell of first sheet as "selection"
				return mockWorksheet(_workbookData.sheets[0], 0).getRange("A1");
			},
		},
		sync: async () => {
			// Process queued loads — in mock, data is immediately available
		},
	};
}

// ── Install mock ────────────────────────────────────────────────

export function installExcelMock() {
	const mockExcel = {
		run: async (fn: (ctx: any) => Promise<any>) => {
			const ctx = createMockContext();
			return fn(ctx);
		},
	};

	(globalThis as any).Excel = mockExcel;
	(globalThis as any).window = globalThis;
}

export function uninstallExcelMock() {
	delete (globalThis as any).Excel;
}
