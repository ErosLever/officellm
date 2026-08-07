/**
 * Word command handler using Office JS API.
 *
 * Key Word JS API patterns:
 * - Word.run(async (context) => { ... }) for batched operations
 * - context.document.body.paragraphs for paragraph access
 * - paragraph.load("text,style,outlineLevel,uniqueLocalId") for properties
 * - context.document.getSelection() for current selection
 * - range.insertComment(text) for comments
 * - Must sync() before reading any loaded property
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
			case "word_get_outline":
				result = await handleGetOutline(args);
				break;
			case "word_get_paragraphs":
				result = await handleGetParagraphs(args);
				break;
			case "word_get_selection":
				result = await handleGetSelection(args);
				break;
			case "word_search":
				result = await handleSearch(args);
				break;
			case "word_replace_text":
				result = await handleReplaceText(args);
				break;
			case "word_insert_text":
				result = await handleInsertText(args);
				break;
			case "word_add_comment":
				result = await handleAddComment(args);
				break;
			case "word_delete_paragraph":
				result = await handleDeleteParagraph(args);
				break;
			case "word_get_tracked_changes":
				result = await handleGetTrackedChanges(args);
				break;
			case "word_accept_all_changes":
				result = await handleAcceptAllChanges(args);
				break;
			case "word_reject_all_changes":
				result = await handleRejectAllChanges(args);
				break;
			case "word_get_tables":
				result = await handleGetTables(args);
				break;
			case "word_insert_table":
				result = await handleInsertTable(args);
				break;
			case "word_update_table_cell":
				result = await handleUpdateTableCell(args);
				break;
			case "word_add_table_rows":
				result = await handleAddTableRows(args);
				break;
			case "word_delete_table_row":
				result = await handleDeleteTableRow(args);
				break;
			case "word_delete_table_column":
				result = await handleDeleteTableColumn(args);
				break;
			case "word_add_table_column":
				result = await handleAddTableColumn(args);
				break;
			case "word_merge_table_cells":
				result = await handleMergeTableCells(args);
				break;
			case "word_split_table_cell":
				result = await handleSplitTableCell(args);
				break;
			case "word_copy_table_structure":
				result = await handleCopyTableStructure(args);
				break;
			case "word_set_table_format":
				result = await handleSetTableFormat(args);
				break;
			case "word_get_headers_footers":
				result = await handleGetHeadersFooters(args);
				break;
			case "word_set_header_footer":
				result = await handleSetHeaderFooter(args);
				break;
			case "word_replace_selection":
				result = await handleReplaceSelection(args);
				break;
			case "word_insert_image":
				result = await handleInsertImage(args);
				break;
			case "word_apply_style":
				result = await handleApplyStyle(args);
				break;
			case "word_get_sections":
				result = await handleGetSections(args);
				break;
			case "word_insert_list":
				result = await handleInsertList(args);
				break;
			case "word_find_replace":
				result = await handleFindReplace(args);
				break;
			case "word_get_bookmarks":
				result = await handleGetBookmarks(args);
				break;
			case "word_insert_bookmark":
				result = await handleInsertBookmark(args);
				break;
			case "word_delete_bookmark":
				result = await handleDeleteBookmark(args);
				break;
			case "word_goto_bookmark":
				result = await handleGotoBookmark(args);
				break;
			case "word_get_properties":
				result = await handleGetProperties(args);
				break;
			case "word_set_properties":
				result = await handleSetProperties(args);
				break;
			case "word_get_hyperlinks":
				result = await handleGetHyperlinks(args);
				break;
			case "word_insert_hyperlink":
				result = await handleInsertHyperlink(args);
				break;
			case "word_insert_footnote":
				result = await handleInsertFootnote(args);
				break;
			case "word_insert_endnote":
				result = await handleInsertEndnote(args);
				break;
			case "word_insert_field":
				result = await handleInsertField(args);
				break;
			case "word_get_content_controls":
				result = await handleGetContentControls(args);
				break;
			case "word_insert_content_control":
				result = await handleInsertContentControl(args);
				break;
			case "word_get_formatting":
				result = await handleGetFormatting(args);
				break;
			case "word_set_formatting":
				result = await handleSetFormatting(args);
				break;
			case "word_get_comments":
				result = await handleGetComments(args);
				break;
			case "word_edit_comment":
				result = await handleEditComment(args);
				break;
			case "word_resolve_comment":
				result = await handleResolveComment(args);
				break;
			case "word_delete_comment":
				result = await handleDeleteComment(args);
				break;
			case "word_reply_to_comment":
				result = await handleReplyToComment(args);
				break;
			case "word_edit_reply":
				result = await handleEditReply(args);
				break;
			case "word_delete_reply":
				result = await handleDeleteReply(args);
				break;
			case "word_get_styles":
				result = await handleGetStyles(args);
				break;
			case "word_modify_style":
				result = await handleModifyStyle(args);
				break;
			case "word_create_style":
				result = await handleCreateStyle(args);
				break;
			case "word_create_and_remap_style":
				result = await handleCreateAndRemapStyle(args);
				break;
			default:
				result = { error: `Unknown Word command: ${commandName}` };
		}

		if (result && typeof result === "object" && "error" in result) {
			success = false;
		}
	} catch (error) {
		let errorMessage: string;
		try {
			const e = error as any;
			const di = e?.debugInfo;
			errorMessage = di?.message || di?.errorLocation || JSON.stringify(di) || e?.message || e?.code || JSON.stringify(e) || String(error);
		} catch {
			errorMessage = String(error);
		}
		console.error(`Word command ${commandId} failed:`, JSON.stringify(error), error);
		success = false;
		result = { error: errorMessage };
	}

	const errorStr = (!success && result && typeof result === "object" && "error" in result)
		? (result as any).error as string
		: undefined;
	await reportResult(commandId, success, errorStr, result);
	return result;
}

// ── Helpers ─────────────────────────────────────────────────────

function runInWord<T>(fn: (ctx: any) => Promise<T>): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const Word: any = (window as any).Word;
		if (!Word || typeof Word.run !== "function") {
			reject(new Error("Word.run() not available"));
			return;
		}
		Word.run(async (ctx: any) => {
			resolve(await fn(ctx));
		}).catch(reject);
	});
}

// ── Read tools ──────────────────────────────────────────────────

async function handleGetOutline(args: unknown): Promise<unknown> {
	const config = args as { maxDepth?: number };
	const maxDepth = config.maxDepth ?? 3;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		// Load heading-relevant properties
		for (const p of paragraphs.items) {
			p.load("text,style,outlineLevel,uniqueLocalId");
		}
		await ctx.sync();

		const headings: Array<{
			index: number;
			level: number;
			text: string;
			style: string;
			id: string;
		}> = [];

		for (let i = 0; i < paragraphs.items.length; i++) {
			const p = paragraphs.items[i];
			const outlineLevel = String(p.outlineLevel || "");
			const style = String(p.style || "");

			// Extract numeric level from outlineLevel (e.g., "OutlineLevel1" → 1)
			let level = 0;
			const match = outlineLevel.match(/OutlineLevel(\d)/);
			if (match) {
				level = parseInt(match[1]);
			} else if (/Heading\s*(\d)/i.test(style)) {
				const headingMatch = style.match(/Heading\s*(\d)/i);
				if (headingMatch) level = parseInt(headingMatch[1]);
			}

			if (level > 0 && level <= maxDepth) {
				headings.push({
					index: i,
					level,
					text: String(p.text || "").trim(),
					style: String(p.style || ""),
					id: String(p.uniqueLocalId || ""),
				});
			}
		}

		return {
			documentName: "Document",
			totalParagraphs: paragraphs.items.length,
			headings,
		};
	});
}

async function handleGetParagraphs(args: unknown): Promise<unknown> {
	const config = args as { startIndex?: number; count?: number };
	const startIndex = config.startIndex ?? 0;
	const count = config.count ?? 50;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		const total = paragraphs.items.length;
		const end = Math.min(startIndex + count, total);
		const result: Array<{
			index: number;
			text: string;
			style: string;
			id: string;
		}> = [];

		for (let i = startIndex; i < end; i++) {
			paragraphs.items[i].load("text,style,uniqueLocalId");
		}
		await ctx.sync();

		for (let i = startIndex; i < end; i++) {
			const p = paragraphs.items[i];
			result.push({
				index: i,
				text: String(p.text || ""),
				style: String(p.style || ""),
				id: String(p.uniqueLocalId || ""),
			});
		}

		return { totalParagraphs: total, paragraphs: result };
	});
}

async function handleGetSelection(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		const selection = ctx.document.getSelection();
		selection.load("text");
		await ctx.sync();

		const text = String(selection.text || "");
		const paragraphs = selection.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		// Load paragraph context
		for (const p of paragraphs.items) {
			p.load("text,style");
		}
		await ctx.sync();

		const contextParagraphs = paragraphs.items.map((p: any, i: number) => ({
			index: i,
			text: String(p.text || ""),
			style: String(p.style || ""),
		}));

		return {
			type: text ? "text" : "empty",
			text,
			paragraphs: contextParagraphs,
		};
	});
}

async function handleSearch(args: unknown): Promise<unknown> {
	const config = args as { searchText?: string; matchCase?: boolean; snippetRadius?: number };
	const searchText = config.searchText ?? "";
	const matchCase = config.matchCase ?? false;
	const snippetRadius = config.snippetRadius ?? 60;

	if (!searchText) return { error: "searchText is required" };

	const escaped = searchText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
	const re = new RegExp(escaped, matchCase ? "g" : "gi");

	return runInWord(async (ctx) => {
		// Scan our own loaded paragraph text rather than Word's body.search(),
		// which can return duplicate/corrupted ranges from hidden field-code
		// runs (content controls, mail-merge fields, cross-references).
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		for (const p of paragraphs.items) p.load("text");
		await ctx.sync();

		const results: Array<{ paragraphIndex: number; snippet: string }> = [];

		paragraphs.items.forEach((p: any, paragraphIndex: number) => {
			const paraText = String(p.text || "");
			re.lastIndex = 0;
			let m: RegExpExecArray | null;
			while ((m = re.exec(paraText)) !== null) {
				const matchStart = m.index;
				const matchEnd = matchStart + m[0].length;
				const start = Math.max(0, matchStart - snippetRadius);
				const end = Math.min(paraText.length, matchEnd + snippetRadius);
				const before = start > 0 ? "…" + paraText.slice(start, matchStart) : paraText.slice(0, matchStart);
				const after = end < paraText.length ? paraText.slice(matchEnd, end) + "…" : paraText.slice(matchEnd, end);
				results.push({ paragraphIndex, snippet: before + "[" + m[0] + "]" + after });
				if (m[0].length === 0) re.lastIndex++;
			}
		});

		return { searchText, totalMatches: results.length, matches: results };
	});
}

// ── Write tools ─────────────────────────────────────────────────

async function handleReplaceText(args: unknown): Promise<unknown> {
	const config = args as {
		paragraphIndex?: number;
		oldText?: string;
		newText?: string;
	};
	const { paragraphIndex = 0, oldText = "", newText = "" } = config;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
			return {
				error: `Paragraph index ${paragraphIndex} out of range (0-${paragraphs.items.length - 1})`,
			};
		}

		const paragraph = paragraphs.items[paragraphIndex];
		paragraph.load("text");
		await ctx.sync();

		const currentText = String(paragraph.text || "");
		if (!currentText.includes(oldText)) {
			return {
				error: `Text '${oldText}' not found in paragraph ${paragraphIndex}`,
			};
		}

		// Use search within the paragraph to do targeted replace
		const searchResults = paragraph.search(oldText, { matchCase: true });
		searchResults.load("items");
		await ctx.sync();

		if (searchResults.items.length > 0) {
			searchResults.items[0].insertText(newText, "Replace");
			await ctx.sync();
		}

		return { paragraphIndex, oldText, newText, replaced: true, tracked: true };
	});
}

async function handleInsertText(args: unknown): Promise<unknown> {
	const config = args as {
		text?: string;
		insertLocation?: string;
		paragraphIndex?: number;
	};
	const { text = "", insertLocation = "end", paragraphIndex } = config;

	return runInWord(async (ctx) => {
		if (insertLocation === "end" || paragraphIndex === undefined) {
			// Insert at the end of the document
			ctx.document.body.insertParagraph(text, "End");
			await ctx.sync();
			return { text, insertLocation: "end", inserted: true, tracked: true };
		}

		// Insert relative to a specific paragraph
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
			return { error: `Paragraph index ${paragraphIndex} out of range` };
		}

		const paragraph = paragraphs.items[paragraphIndex];
		const location = insertLocation === "beforeParagraph" ? "Before" : "After";
		paragraph.insertParagraph(text, location);
		await ctx.sync();

		return {
			text,
			insertLocation,
			paragraphIndex,
			inserted: true,
			tracked: true,
		};
	});
}

async function handleAddComment(args: unknown): Promise<unknown> {
	const config = args as {
		commentText?: string;
		searchText?: string;
		paragraphIndex?: number;
		fromParagraph?: number;
		toParagraph?: number;
	};
	const { commentText = "", searchText, paragraphIndex, fromParagraph, toParagraph } = config;

	return runInWord(async (ctx) => {
		// 1. Phrase anchor — document-wide search
		if (searchText) {
			const results = ctx.document.body.search(searchText, { matchCase: false });
			results.load("items");
			await ctx.sync();
			if (results.items.length === 0) {
				return { error: `Text not found: "${searchText}"`, errorCode: "NOT_FOUND" };
			}
			results.items[0].insertComment(commentText);
			await ctx.sync();
			return { commentText, searchText, added: true };
		}

		// 2. Paragraph range — fromParagraph/toParagraph (paragraphIndex is shorthand for both)
		const paraFrom = fromParagraph ?? paragraphIndex;
		const paraTo = toParagraph ?? paraFrom;

		if (paraFrom !== undefined) {
			const paragraphs = ctx.document.body.paragraphs;
			paragraphs.load("items");
			await ctx.sync();

			const count = paragraphs.items.length;
			if (paraFrom < 0 || paraFrom >= count) {
				return { error: `fromParagraph ${paraFrom} out of range (0–${count - 1})` };
			}
			if (paraTo! < paraFrom || paraTo! >= count) {
				return { error: `toParagraph ${paraTo} out of range (must be >= fromParagraph and < ${count})` };
			}

			const startRange = paragraphs.items[paraFrom].getRange("Whole");
			if (paraTo === paraFrom) {
				startRange.insertComment(commentText);
			} else {
				const endRange = paragraphs.items[paraTo!].getRange("Whole");
				const spanRange = startRange.expandTo(endRange);
				spanRange.insertComment(commentText);
			}
			await ctx.sync();

			return { commentText, fromParagraph: paraFrom, toParagraph: paraTo, added: true };
		}

		// 3. Fallback — current selection
		const selection = ctx.document.getSelection();
		selection.insertComment(commentText);
		await ctx.sync();

		return { commentText, target: "selection", added: true };
	});
}

async function handleDeleteParagraph(args: unknown): Promise<unknown> {
	const config = args as { paragraphIndex?: number };
	const { paragraphIndex = 0 } = config;

	return runInWord(async (ctx) => {
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();

		if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
			return { error: `Paragraph index ${paragraphIndex} out of range` };
		}

		// Enable tracked changes for this mutation
		const savedMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = "TrackMineOnly";

		const paragraph = paragraphs.items[paragraphIndex];
		paragraph.delete();
		await ctx.sync();

		ctx.document.changeTrackingMode = savedMode;

		return { paragraphIndex, deleted: true, tracked: true };
	});
}

// ── Tracked Changes Tools ───────────────────────────────────────

async function handleGetTrackedChanges(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		const mode = ctx.document.changeTrackingMode;
		await ctx.sync();

		return {
			changeTrackingMode: mode,
			pendingChanges: 0, // Mock doesn't track real revision count
		};
	});
}

async function handleAcceptAllChanges(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		ctx.document.acceptAllChanges();
		await ctx.sync();

		return { accepted: true };
	});
}

async function handleRejectAllChanges(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		ctx.document.rejectAllChanges();
		await ctx.sync();

		return { rejected: true };
	});
}

// ── Word Structure: Tables ──────────────────────────────────

async function handleGetTables(args: unknown): Promise<unknown> {
	const config = args as { includeCellText?: boolean; maxRows?: number };
	const { includeCellText = true, maxRows = 50 } = config;

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();

		const result: Array<{
			index: number;
			rowCount: number;
			columnCount: number;
			cells?: string[][];
		}> = [];
		for (let i = 0; i < tables.items.length; i++) {
			const tbl = tables.items[i];
			tbl.load("rowCount,columnCount");
			await ctx.sync();

			const entry: (typeof result)[0] = {
				index: i,
				rowCount: tbl.rowCount,
				columnCount: tbl.columnCount,
			};

			if (includeCellText) {
				const rows = Math.min(tbl.rowCount, maxRows);
				const cells: string[][] = [];
				for (let r = 0; r < rows; r++) {
					const rowCells: string[] = [];
					for (let c = 0; c < tbl.columnCount; c++) {
						try {
							const cell = tbl.getCell(r, c);
							cell.value.load("text");
							await ctx.sync();
							rowCells.push(String((cell.value as any).text || ""));
						} catch {
							rowCells.push("[error]");
						}
					}
					cells.push(rowCells);
				}
				entry.cells = cells;
			}
			result.push(entry);
		}
		return { tableCount: result.length, tables: result };
	});
}

async function handleInsertTable(args: unknown): Promise<unknown> {
	const config = args as {
		rows?: number;
		columns?: number;
		afterParagraphIndex?: number;
		headerRow?: string[];
	};
	const { rows = 1, columns = 2, afterParagraphIndex = -1, headerRow } = config;

	return runInWord(async (ctx) => {
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (
			Word as any
		).ChangeTrackingMode.trackMineOnly;

		const body = ctx.document.body;
		let insertRange: any;

		if (afterParagraphIndex === -1) {
			insertRange = body.getRange("End");
		} else {
			const paras = body.paragraphs;
			paras.load("items");
			await ctx.sync();
			if (afterParagraphIndex < paras.items.length) {
				insertRange = paras.items[afterParagraphIndex].getRange("After");
			} else {
				insertRange = body.getRange("End");
			}
		}

		insertRange.insertTable(
			rows + (headerRow ? 1 : 0),
			columns,
			(Word as any).InsertLocation.after,
			headerRow || undefined,
		);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return {
			rows,
			columns,
			afterParagraphIndex,
			inserted: true,
			tracked: true,
		};
	});
}

async function handleUpdateTableCell(args: unknown): Promise<unknown> {
	const config = args as {
		tableIndex?: number;
		row?: number;
		column?: number;
		text?: string;
	};
	const { tableIndex = 0, row = 0, column = 0, text = "" } = config;

	return runInWord(async (ctx) => {
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (
			Word as any
		).ChangeTrackingMode.trackMineOnly;

		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();

		if (tableIndex >= tables.items.length) {
			ctx.document.changeTrackingMode = originalMode;
			await ctx.sync();
			return {
				error: `Table index ${tableIndex} out of bounds. Document has ${tables.items.length} tables.`,
				errorCode: "CELL_OUT_OF_BOUNDS",
			};
		}

		const tbl = tables.items[tableIndex];
		tbl.load("rowCount,columnCount");
		await ctx.sync();

		if (row >= tbl.rowCount || column >= tbl.columnCount) {
			ctx.document.changeTrackingMode = originalMode;
			await ctx.sync();
			return {
				error: `Cell (${row},${column}) out of bounds. Table is ${tbl.rowCount}x${tbl.columnCount}.`,
				errorCode: "CELL_OUT_OF_BOUNDS",
				details: { rowCount: tbl.rowCount, columnCount: tbl.columnCount },
			};
		}

		const cell = tbl.getCell(row, column);
		cell.value.text = text;
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { tableIndex, row, column, text, updated: true, tracked: true };
	});
}

// ── Word Table: Extended Operations ────────────────────────

async function handleAddTableRows(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { tableIndex = 0, insertAfterRow = -1, rows, headers } = config;
	if (!rows || !Array.isArray(rows)) return { error: "rows is required (array)", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();
		if ((tableIndex as number) >= tables.items.length)
			return { error: `Table index ${tableIndex} out of bounds`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const tbl = tables.items[tableIndex as number];
		tbl.load("rowCount,columnCount");
		await ctx.sync();

		// Normalise: array of objects → matrix
		let matrix: string[][];
		if (rows.length > 0 && typeof rows[0] === "object" && !Array.isArray(rows[0])) {
			const hdrs = (headers as string[]) ?? Object.keys(rows[0] as object);
			matrix = (rows as Record<string, unknown>[]).map(r => hdrs.map(h => String(r[h] ?? "")));
		} else {
			matrix = rows as string[][];
		}

		const location = (insertAfterRow as number) === -1 ? Word.InsertLocation.end : Word.InsertLocation.after;
		if ((insertAfterRow as number) >= 0) {
			const row = tbl.rows.getFirst();
			tbl.rows.load("items");
			await ctx.sync();
			if ((insertAfterRow as number) < tbl.rows.items.length) {
				tbl.rows.items[insertAfterRow as number].insertRows(Word.InsertLocation.after, matrix.length, matrix);
			} else {
				tbl.addRows(Word.InsertLocation.end, matrix.length, matrix);
			}
		} else {
			tbl.addRows(Word.InsertLocation.end, matrix.length, matrix);
		}
		await ctx.sync();
		return { tableIndex, rowsAdded: matrix.length };
	});
}

async function handleDeleteTableRow(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { tableIndex = 0, rowIndex } = config;
	if (rowIndex === undefined) return { error: "rowIndex is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();
		if ((tableIndex as number) >= tables.items.length)
			return { error: `Table index ${tableIndex} out of bounds`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const tbl = tables.items[tableIndex as number];
		tbl.load("rowCount");
		await ctx.sync();
		if ((rowIndex as number) >= tbl.rowCount)
			return { error: `Row ${rowIndex} out of bounds (rowCount=${tbl.rowCount})`, errorCode: "CELL_OUT_OF_BOUNDS" };

		tbl.deleteRows(rowIndex as number, 1);
		await ctx.sync();
		return { tableIndex, rowIndex, deleted: true };
	});
}

async function handleMergeTableCells(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { tableIndex = 0, topRow, firstColumn, bottomRow, lastColumn } = config;
	if (topRow === undefined || firstColumn === undefined || bottomRow === undefined || lastColumn === undefined)
		return { error: "topRow, firstColumn, bottomRow, lastColumn are required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();
		if ((tableIndex as number) >= tables.items.length)
			return { error: `Table index ${tableIndex} out of bounds`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const tbl = tables.items[tableIndex as number];
		tbl.mergeCells(topRow as number, firstColumn as number, bottomRow as number, lastColumn as number);
		await ctx.sync();
		return { tableIndex, merged: true };
	});
}

async function handleSplitTableCell(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { tableIndex = 0, row, column, rowCount = 1, columnCount = 2 } = config;
	if (row === undefined || column === undefined)
		return { error: "row and column are required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();
		if ((tableIndex as number) >= tables.items.length)
			return { error: `Table index ${tableIndex} out of bounds`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const tbl = tables.items[tableIndex as number];
		const cell = tbl.getCell(row as number, column as number);
		cell.split(rowCount as number, columnCount as number);
		await ctx.sync();
		return { tableIndex, row, column, split: true };
	});
}

// Border helpers for add/delete column smart adjustment
const OUTER_BORDER_LOCS = ["Top", "Bottom", "Left", "Right"] as const;
const ALL_BORDER_LOCS = [...OUTER_BORDER_LOCS, "InsideHorizontal", "InsideVertical"] as const;

async function readTableBorders(tbl: any, ctx: any): Promise<Record<string, {color: string; type: string; width: number}>> {
	const result: Record<string, any> = {};
	for (const loc of ALL_BORDER_LOCS) {
		const b = tbl.getBorder(loc as any);
		b.load("color,type,width");
		result[loc] = b;
	}
	await ctx.sync();
	return Object.fromEntries(ALL_BORDER_LOCS.map(loc => [loc, { color: result[loc].color, type: result[loc].type, width: result[loc].width }]));
}

function isOuterOnlyPattern(borders: Record<string, {type: string}>): boolean {
	const outerSet = OUTER_BORDER_LOCS.some(l => borders[l]?.type && borders[l].type !== "None" && borders[l].type !== "Mixed");
	const innerSet = ["InsideHorizontal","InsideVertical"].some(l => borders[l]?.type && borders[l].type !== "None" && borders[l].type !== "Mixed");
	return outerSet && !innerSet;
}

async function applyTableBorders(tbl: any, borders: Record<string, {color: string; type: string; width: number}>, ctx: any): Promise<void> {
	for (const [loc, props] of Object.entries(borders)) {
		const b = tbl.getBorder(loc as any);
		b.color = props.color;
		b.type = props.type;
		b.width = props.width;
	}
	await ctx.sync();
}

async function handleDeleteTableColumn(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { tableIndex = 0, columnIndex } = config;
	if (columnIndex === undefined) return { error: "columnIndex is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();
		if ((tableIndex as number) >= tables.items.length)
			return { error: `Table index ${tableIndex} out of bounds`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const tbl = tables.items[tableIndex as number];
		tbl.load("rowCount,columnCount");
		await ctx.sync();
		const colIdx = columnIndex as number;
		if (colIdx >= tbl.columnCount)
			return { error: `Column ${colIdx} out of bounds (columnCount=${tbl.columnCount})`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const borders = await readTableBorders(tbl, ctx);
		const outerOnly = isOuterOnlyPattern(borders);
		const isFirst = colIdx === 0;
		const isLast = colIdx === tbl.columnCount - 1;

		tbl.deleteColumns(colIdx, 1);
		await ctx.sync();

		// After deletion, restore outer border on the newly-exposed outer column
		if (outerOnly && (isFirst || isLast)) {
			const newBorders = { ...borders };
			if (isFirst) { newBorders["InsideVertical"] = { color: "Auto", type: "None", width: 0 }; }
			if (isLast)  { newBorders["InsideVertical"] = { color: "Auto", type: "None", width: 0 }; }
			await applyTableBorders(tbl, newBorders, ctx);
		}

		return { tableIndex, columnIndex: colIdx, deleted: true };
	});
}

async function handleAddTableColumn(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { tableIndex = 0, insertBeforeColumn = -1, values, copyFormatFrom = "left" } = config;

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();
		if ((tableIndex as number) >= tables.items.length)
			return { error: `Table index ${tableIndex} out of bounds`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const tbl = tables.items[tableIndex as number];
		tbl.load("rowCount,columnCount");
		await ctx.sync();

		const borders = await readTableBorders(tbl, ctx);
		const outerOnly = isOuterOnlyPattern(borders);
		const colIdx = insertBeforeColumn as number;
		const isFirst = colIdx === 0;
		const isLast = colIdx === -1 || colIdx >= tbl.columnCount;
		const cellValues = values ? (values as string[][]) : undefined;

		if (isFirst) {
			tbl.addColumns(Word.InsertLocation.start, 1, cellValues ?? null);
		} else {
			// Insert after the column before insertBeforeColumn
			const refCol = isLast ? tbl.columnCount - 1 : colIdx - 1;
			tbl.getCell(0, refCol).insertColumns(Word.InsertLocation.after, 1, cellValues ?? null);
		}
		await ctx.sync();

		if (outerOnly && (isFirst || isLast)) {
			await applyTableBorders(tbl, borders, ctx);
		}

		return { tableIndex, columnAdded: true };
	});
}

async function handleCopyTableStructure(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { tableIndex = 0, afterParagraphIndex = -1, includeHeaders = true, emptyRows = 1 } = config;

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();
		if ((tableIndex as number) >= tables.items.length)
			return { error: `Table index ${tableIndex} out of bounds`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const tbl = tables.items[tableIndex as number];
		tbl.load("rowCount,columnCount,headerRowCount");
		await ctx.sync();

		const numCols = tbl.columnCount;
		const headerCount = (includeHeaders && tbl.headerRowCount > 0) ? tbl.headerRowCount : 0;

		// Read header cell text
		let headerRow: string[] | undefined;
		if (headerCount > 0) {
			headerRow = [];
			for (let c = 0; c < numCols; c++) {
				const cell = tbl.getCell(0, c);
				cell.value.load("text");
				await ctx.sync();
				headerRow.push(String((cell.value as any).text ?? ""));
			}
		}

		// Read borders and cell padding
		const borders = await readTableBorders(tbl, ctx);
		const padLocs = ["Top","Left","Bottom","Right"] as const;
		const padProxies = padLocs.map(l => { const p = tbl.getCellPadding(l as any); p.load("value"); return p; });
		await ctx.sync();
		const padding = Object.fromEntries(padLocs.map((l,i) => [l, (padProxies[i] as any).value as number]));

		// Read column widths from row 0
		const colWidths: number[] = [];
		for (let c = 0; c < numCols; c++) {
			const cell = tbl.getCell(0, c);
			cell.load("columnWidth");
			await ctx.sync();
			colWidths.push(cell.columnWidth);
		}

		// Determine insert range
		const body = ctx.document.body;
		let insertRange: any;
		if ((afterParagraphIndex as number) === -1) {
			insertRange = body.getRange("End");
		} else {
			const paras = body.paragraphs;
			paras.load("items");
			await ctx.sync();
			const idx = afterParagraphIndex as number;
			insertRange = idx < paras.items.length ? paras.items[idx].getRange("After") : body.getRange("End");
		}

		const totalRows = headerCount + (emptyRows as number);
		const newTbl = insertRange.insertTable(totalRows, numCols, Word.InsertLocation.after, headerRow ?? null);
		await ctx.sync();

		// Apply borders
		await applyTableBorders(newTbl, borders, ctx);

		// Apply padding
		for (const [loc, val] of Object.entries(padding)) {
			newTbl.setCellPadding(loc as any, val);
		}
		await ctx.sync();

		// Apply column widths
		for (let c = 0; c < numCols; c++) {
			const cell = newTbl.getCell(0, c);
			cell.columnWidth = colWidths[c];
		}
		await ctx.sync();

		return { tableIndex, inserted: true, rows: totalRows, columns: numCols };
	});
}

async function handleSetTableFormat(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { tableIndex = 0 } = config;

	return runInWord(async (ctx) => {
		const tables = ctx.document.body.tables;
		tables.load("items");
		await ctx.sync();
		if ((tableIndex as number) >= tables.items.length)
			return { error: `Table index ${tableIndex} out of bounds`, errorCode: "CELL_OUT_OF_BOUNDS" };

		const tbl = tables.items[tableIndex as number];
		tbl.load("rowCount,columnCount,headerRowCount");
		await ctx.sync();

		const applied: string[] = [];

		if (config.headerRowCount !== undefined) { tbl.headerRowCount = config.headerRowCount as number; applied.push("headerRowCount"); }
		if (config.style !== undefined) { tbl.style = config.style as string; applied.push("style"); }
		if (config.horizontalAlignment !== undefined) { tbl.horizontalAlignment = config.horizontalAlignment as string; applied.push("horizontalAlignment"); }
		if (config.verticalAlignment !== undefined) { tbl.verticalAlignment = config.verticalAlignment as string; applied.push("verticalAlignment"); }
		for (const side of ["Top","Left","Bottom","Right"] as const) {
			const key = `cellPadding${side}`;
			if (config[key] !== undefined) { tbl.setCellPadding(side, config[key] as number); applied.push(key); }
		}

		if (config.columnWidths !== undefined) {
			const widths = config.columnWidths as number[];
			for (let c = 0; c < Math.min(widths.length, tbl.columnCount); c++) {
				tbl.getCell(0, c).columnWidth = widths[c];
			}
			applied.push("columnWidths");
		}

		await ctx.sync();

		// Row height
		if (config.minimumRowHeight !== undefined) {
			const h = config.minimumRowHeight as number;
			const rowIdx = config.rowIndex as number | undefined;
			tbl.rows.load("items");
			await ctx.sync();
			const rows = rowIdx !== undefined ? [tbl.rows.items[rowIdx]] : tbl.rows.items;
			for (const r of rows) { r.preferredHeight = h; }
			await ctx.sync();
			applied.push("minimumRowHeight");
		}

		// Column-level style/font
		if (config.columnIndex !== undefined) {
			const ci = config.columnIndex as number;
			tbl.rows.load("items");
			await ctx.sync();
			for (const r of tbl.rows.items) {
				if (ci >= tbl.columnCount) continue;
				r.cells.load("items");
				await ctx.sync();
				const targetCell = r.cells.items[ci];
				if (!targetCell) continue;
				const isHdr = r.isHeader;
				const styleName = isHdr
					? (config.headerStyle ?? config.paragraphStyle) as string | undefined
					: config.paragraphStyle as string | undefined;
				if (styleName) { targetCell.value.style = styleName; }
				if (config.bold !== undefined) targetCell.value.font.bold = config.bold as boolean;
				if (config.italic !== undefined) targetCell.value.font.italic = config.italic as boolean;
				if (config.color !== undefined) targetCell.value.font.color = config.color as string;
				if (config.alignment !== undefined) targetCell.value.paragraphs.getFirst().alignment = config.alignment as string;
			}
			await ctx.sync();
			applied.push("columnFormatting");
		}

		return { tableIndex, formatted: true, applied };
	});
}

// ── Word Structure: Headers/Footers ────────────────────────

async function handleGetHeadersFooters(args: unknown): Promise<unknown> {
	const config = args as { sectionIndex?: number };
	const { sectionIndex } = config;

	return runInWord(async (ctx) => {
		const sections = ctx.document.sections;
		sections.load("items");
		await ctx.sync();

		const targetSections =
			sectionIndex !== undefined
				? [sections.items[sectionIndex]]
				: sections.items;
		const result: Array<{
			sectionIndex: number;
			header?: string;
			footer?: string;
		}> = [];

		for (let i = 0; i < targetSections.length; i++) {
			const sec = targetSections[i];
			const entry: (typeof result)[0] = { sectionIndex: sectionIndex ?? i };

			try {
				const header = sec.getHeader("Default");
				header.load("text");
				await ctx.sync();
				entry.header = String(header.text || "");
			} catch {
				entry.header = "";
			}

			try {
				const footer = sec.getFooter("Default");
				footer.load("text");
				await ctx.sync();
				entry.footer = String(footer.text || "");
			} catch {
				entry.footer = "";
			}

			result.push(entry);
		}
		return { sectionCount: sections.items.length, sections: result };
	});
}

async function handleSetHeaderFooter(args: unknown): Promise<unknown> {
	const config = args as {
		sectionIndex?: number;
		type?: string;
		variant?: string;
		text?: string;
	};
	const {
		sectionIndex = 0,
		type = "header",
		variant = "default",
		text = "",
	} = config;

	return runInWord(async (ctx) => {
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (
			Word as any
		).ChangeTrackingMode.trackMineOnly;

		const sections = ctx.document.sections;
		sections.load("items");
		await ctx.sync();

		if (sectionIndex >= sections.items.length) {
			ctx.document.changeTrackingMode = originalMode;
			await ctx.sync();
			return {
				error: `Section ${sectionIndex} not found. Document has ${sections.items.length} sections.`,
				errorCode: "INVALID_PARAMETER",
			};
		}

		const sec = sections.items[sectionIndex];
		const bodyObj =
			type === "footer"
				? sec.getFooter(variant === "firstPage" ? "FirstPage" : "Default")
				: sec.getHeader(variant === "firstPage" ? "FirstPage" : "Default");
		bodyObj.insertText(text, (Word as any).InsertLocation.replace);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { sectionIndex, type, variant, text, set: true, tracked: true };
	});
}

// ── Word Structure: Selection & Insert ──────────────────────

async function handleReplaceSelection(args: unknown): Promise<unknown> {
	const config = args as { text?: string };
	const { text = "" } = config;

	return runInWord(async (ctx) => {
		const selection = ctx.document.getSelection();
		selection.load("text");
		await ctx.sync();

		if (!selection.text || selection.text.trim() === "") {
			return {
				error: "No text selected. Select text first.",
				errorCode: "EMPTY_SELECTION",
			};
		}

		const originalText = selection.text;
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (
			Word as any
		).ChangeTrackingMode.trackMineOnly;

		selection.insertText(text, (Word as any).InsertLocation.replace);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { originalText, newText: text, replaced: true, tracked: true };
	});
}

async function handleInsertImage(args: unknown): Promise<unknown> {
	const config = args as {
		imageBase64?: string;
		afterParagraphIndex?: number;
		width?: number;
		height?: number;
	};
	const { imageBase64 = "", afterParagraphIndex = -1, width, height } = config;

	if (!imageBase64)
		return { error: "imageBase64 is required", errorCode: "INVALID_PARAMETER" };

	// Size check: 10MB max
	const sizeBytes = Math.ceil((imageBase64.length * 3) / 4);
	if (sizeBytes > 10 * 1024 * 1024) {
		return {
			error: `Image too large: ${(sizeBytes / 1024 / 1024).toFixed(1)}MB. Max: 10MB.`,
			errorCode: "IMAGE_TOO_LARGE",
		};
	}

	return runInWord(async (ctx) => {
		const body = ctx.document.body;
		let insertRange: any;

		if (afterParagraphIndex === -1) {
			insertRange = body.getRange("End");
		} else {
			const paras = body.paragraphs;
			paras.load("items");
			await ctx.sync();
			insertRange =
				paras.items[
					Math.min(afterParagraphIndex, paras.items.length - 1)
				].getRange("After");
		}

		const image = insertRange.insertInlinePictureFromBase64(
			imageBase64,
			(Word as any).InsertLocation.after,
		);
		if (width) image.width = width;
		if (height) image.height = height;
		await ctx.sync();

		return {
			afterParagraphIndex,
			inserted: true,
			width: image.width,
			height: image.height,
		};
	});
}

// ── Word Structure: Styles & Lists ──────────────────────────

async function handleApplyStyle(args: unknown): Promise<unknown> {
	const config = args as {
		paragraphIndex?: number;
		styleName?: string;
		characterStyle?: string;
		searchText?: string;
		applyToSelection?: boolean;
	};
	const { paragraphIndex = 0, styleName, characterStyle, searchText, applyToSelection } = config;

	// Character style path
	if (characterStyle) {
		return runInWord(async (ctx) => {
			let range: any;
			if (searchText) {
				const results = ctx.document.body.search(searchText, { matchCase: false });
				results.load("items");
				await ctx.sync();
				if (results.items.length === 0)
					return { error: `Text not found: "${searchText}"`, errorCode: "NOT_FOUND" };
				range = results.items[0];
			} else if (applyToSelection) {
				range = ctx.document.getSelection();
			} else {
				return { error: "characterStyle requires searchText or applyToSelection:true", errorCode: "INVALID_PARAMETER" };
			}
			range.style = characterStyle;
			await ctx.sync();
			return { characterStyle, searchText: searchText ?? null, applyToSelection: applyToSelection ?? false, applied: true };
		});
	}

	// Paragraph style path
	if (!styleName)
		return { error: "styleName or characterStyle is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const paras = ctx.document.body.paragraphs;
		paras.load("items");
		await ctx.sync();

		if (paragraphIndex >= paras.items.length) {
			return {
				error: `Paragraph ${paragraphIndex} not found. Document has ${paras.items.length} paragraphs.`,
				errorCode: "INVALID_PARAMETER",
			};
		}

		const para = paras.items[paragraphIndex];
		para.style = styleName;
		await ctx.sync();

		return { paragraphIndex, styleName, applied: true };
	});
}

async function handleGetSections(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx) => {
		const sections = ctx.document.sections;
		sections.load("items");
		await ctx.sync();

		const result: Array<{
			index: number;
			differentFirstPage: boolean;
			differentOddAndEvenPages: boolean;
		}> = [];
		for (let i = 0; i < sections.items.length; i++) {
			const sec = sections.items[i];
			sec.load("differentFirstPage,differentOddAndEvenPages");
			await ctx.sync();
			result.push({
				index: i,
				differentFirstPage: sec.differentFirstPage || false,
				differentOddAndEvenPages: sec.differentOddAndEvenPages || false,
			});
		}
		return { sectionCount: result.length, sections: result };
	});
}

async function handleInsertList(args: unknown): Promise<unknown> {
	const config = args as {
		type?: string;
		items?: string[];
		afterParagraphIndex?: number;
	};
	const { type = "bulleted", items = [], afterParagraphIndex = -1 } = config;
	if (!items.length)
		return { error: "items must be non-empty", errorCode: "EMPTY_ITEMS" };

	return runInWord(async (ctx) => {
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = (
			Word as any
		).ChangeTrackingMode.trackMineOnly;

		const body = ctx.document.body;
		const bulletType =
			type === "numbered"
				? (Word as any).BulletType.numbered
				: (Word as any).BulletType.bulleted;

		// Build list text
		const listText = items.join("\r");
		let insertRange: any;

		if (afterParagraphIndex === -1) {
			insertRange = body.getRange("End");
		} else {
			const paras = body.paragraphs;
			paras.load("items");
			await ctx.sync();
			insertRange =
				paras.items[
					Math.min(afterParagraphIndex, paras.items.length - 1)
				].getRange("After");
		}

		insertRange.insertParagraph(listText, (Word as any).InsertLocation.after);
		await ctx.sync();

		// Apply list formatting to the inserted paragraph
		const insertedPara = insertRange.paragraphs.getLast();
		insertedPara.load("uniqueLocalId");
		await ctx.sync();

		// Split by \r and apply bullet/number formatting
		const listItems = insertedPara.split(["\r"]);
		listItems.load("items");
		await ctx.sync();

		for (const item of listItems.items) {
			item.startList(bulletType);
		}
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return {
			type,
			itemCount: items.length,
			afterParagraphIndex,
			inserted: true,
			tracked: true,
		};
	});
}

// ── Phase 14: Find & Replace ──────────────────────────────────────

async function handleFindReplace(args: unknown): Promise<unknown> {
	const config = args as {
		findText: string;
		replaceText?: string;
		matchCase?: boolean;
		matchWholeWord?: boolean;
		useWildcards?: boolean;
		previewOnly?: boolean;
		scopeFromParagraph?: number;
		scopeToParagraph?: number;
	};

	const {
		findText,
		replaceText = "",
		matchCase = false,
		matchWholeWord = false,
		useWildcards = false,
		previewOnly = false,
		scopeFromParagraph,
		scopeToParagraph,
	} = config;

	const Word: any = (window as any).Word;

	return runInWord(async (ctx: any) => {
		const originalMode = ctx.document.changeTrackingMode;

		// Build search options
		const searchOptions: any = {};
		if (matchCase) searchOptions.matchCase = true;
		if (matchWholeWord) searchOptions.matchWholeWord = true;
		if (useWildcards) searchOptions.matchWildcards = true;

		// Determine search scope
		let searchBody: any;
		if (scopeFromParagraph !== undefined || scopeToParagraph !== undefined) {
			const paras = ctx.document.body.paragraphs;
			paras.load("items");
			await ctx.sync();
			const from = scopeFromParagraph ?? 0;
			const to = scopeToParagraph ?? paras.items.length - 1;
			const startPara = paras.items[Math.max(0, from)];
			const endPara = paras.items[Math.min(to, paras.items.length - 1)];
			const startRange = startPara.getRange("Start");
			const endRange = endPara.getRange("End");
			searchBody = startRange.expandTo(endRange);
		} else {
			searchBody = ctx.document.body;
		}

		const searchResults = searchBody.search(findText, searchOptions);
		searchResults.load(["items"]);
		await ctx.sync();

		if (previewOnly) {
			const previews: any[] = [];
			for (let i = 0; i < Math.min(searchResults.items.length, 50); i++) {
				const range = searchResults.items[i];
				range.load(["text", "paragraphsUnique"]);
			}
			if (searchResults.items.length > 0) {
				await ctx.sync();
			}
			for (let i = 0; i < Math.min(searchResults.items.length, 50); i++) {
				const range = searchResults.items[i];
				previews.push({
					index: i,
					text: range.text,
				});
			}
			return {
				previewOnly: true,
				matchCount: searchResults.items.length,
				previews,
			};
		}

		// Perform replacements with tracked changes
		if (!previewOnly) {
			ctx.document.changeTrackingMode = (
				Word as any
			).ChangeTrackingMode.trackMineOnly;
		}

		const matchCount = searchResults.items.length;
		for (let i = 0; i < searchResults.items.length; i++) {
			searchResults.items[i].insertText(replaceText, "Replace");
		}
		await ctx.sync();

		if (!previewOnly) {
			ctx.document.changeTrackingMode = originalMode;
			await ctx.sync();
		}

		return {
			replacements: matchCount,
			findText,
			tracked: true,
		};
	});
}

// ── Phase 18: Bookmarks ──────────────────────────────────────────

async function handleGetBookmarks(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx: any) => {
		const bookmarks = ctx.document.bookmarks;
		bookmarks.load("items");
		await ctx.sync();

		const result: any[] = [];
		for (const bm of bookmarks.items) {
			bm.load(["name"]);
		}
		await ctx.sync();
		for (const bm of bookmarks.items) {
			result.push({ name: bm.name });
		}

		return { bookmarks: result, count: result.length };
	});
}

async function handleInsertBookmark(args: unknown): Promise<unknown> {
	const config = args as {
		name: string;
		fromParagraph: number;
		toParagraph?: number;
	};
	const { name, fromParagraph, toParagraph } = config;

	return runInWord(async (ctx: any) => {
		const paras = ctx.document.body.paragraphs;
		paras.load("items");
		await ctx.sync();

		const end = toParagraph ?? fromParagraph;
		const startPara = paras.items[Math.max(0, fromParagraph)];
		const endPara = paras.items[Math.min(end, paras.items.length - 1)];
		const startRange = startPara.getRange("Start");
		const endRange = endPara.getRange("End");
		const range = startRange.expandTo(endRange);
		range.insertBookmark(name);
		await ctx.sync();

		return { name, fromParagraph, toParagraph: end, inserted: true };
	});
}

async function handleDeleteBookmark(args: unknown): Promise<unknown> {
	const config = args as { name: string };
	const { name } = config;

	return runInWord(async (ctx: any) => {
		ctx.document.deleteBookmark(name);
		await ctx.sync();
		return { name, deleted: true };
	});
}

async function handleGotoBookmark(args: unknown): Promise<unknown> {
	const config = args as { name: string };
	const { name } = config;

	return runInWord(async (ctx: any) => {
		const range = ctx.document.getBookmarkRange(name);
		range.load("text");
		await ctx.sync();
		return { name, text: range.text };
	});
}

// ── Phase 18: Document Properties ────────────────────────────────

async function handleGetProperties(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx: any) => {
		const props = ctx.document.builtInDocumentProperties;
		props.load([
			"title",
			"author",
			"subject",
			"keywords",
			"category",
			"company",
			"manager",
			"comments",
			"creationDate",
			"lastSaveTime",
			"revisionNumber",
		]);
		await ctx.sync();

		return {
			title: props.title || "",
			author: props.author || "",
			subject: props.subject || "",
			keywords: props.keywords || "",
			category: props.category || "",
			company: props.company || "",
			manager: props.manager || "",
			comments: props.comments || "",
			creationDate: props.creationDate?.toISOString?.() || "",
			lastSaveTime: props.lastSaveTime?.toISOString?.() || "",
			revisionNumber: props.revisionNumber || "",
		};
	});
}

async function handleSetProperties(args: unknown): Promise<unknown> {
	const config = args as {
		title?: string;
		author?: string;
		subject?: string;
		keywords?: string;
		category?: string;
		company?: string;
		manager?: string;
		comments?: string;
	};
	return runInWord(async (ctx: any) => {
		const props = ctx.document.builtInDocumentProperties;
		if (config.title !== undefined) props.title = config.title;
		if (config.author !== undefined) props.author = config.author;
		if (config.subject !== undefined) props.subject = config.subject;
		if (config.keywords !== undefined) props.keywords = config.keywords;
		if (config.category !== undefined) props.category = config.category;
		if (config.company !== undefined) props.company = config.company;
		if (config.manager !== undefined) props.manager = config.manager;
		if (config.comments !== undefined) props.comments = config.comments;
		await ctx.sync();
		return { updated: true };
	});
}

// ── Phase 18: Hyperlinks ──────────────────────────────────────────

async function handleGetHyperlinks(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx: any) => {
		const hyperlinks = ctx.document.hyperlinks;
		hyperlinks.load("items");
		await ctx.sync();

		const result: any[] = [];
		for (const hl of hyperlinks.items) {
			hl.load(["address", "screenTip"]);
			const range = hl.getRange();
			range.load("text");
		}
		await ctx.sync();
		for (const hl of hyperlinks.items) {
			const range = hl.getRange();
			result.push({ address: hl.address, text: range.text });
		}

		return { hyperlinks: result, count: result.length };
	});
}

async function handleInsertHyperlink(args: unknown): Promise<unknown> {
	const config = args as { text: string; url: string; paragraphIndex?: number };
	const { text, url, paragraphIndex } = config;

	return runInWord(async (ctx: any) => {
		const Word: any = (window as any).Word;
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;

		let range: any;
		if (paragraphIndex !== undefined) {
			const paras = ctx.document.body.paragraphs;
			paras.load("items");
			await ctx.sync();
			const para =
				paras.items[Math.min(paragraphIndex, paras.items.length - 1)];
			range = para.getRange("End");
		} else {
			range = ctx.document.body.getRange("End");
		}

		const insertedRange = range.insertText(text, "After");
		insertedRange.hyperlink = url;
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();

		return { text, url, inserted: true, tracked: true };
	});
}

// ── Phase 18: Footnotes/Endnotes & Fields ────────────────────────

async function handleInsertFootnote(args: unknown): Promise<unknown> {
	const config = args as { paragraphIndex: number; text: string };
	return runInWord(async (ctx: any) => {
		const Word: any = (window as any).Word;
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;

		const paras = ctx.document.body.paragraphs;
		paras.load("items");
		await ctx.sync();
		const para =
			paras.items[Math.min(config.paragraphIndex, paras.items.length - 1)];
		const range = para.getRange("End");
		range.insertFootnote(config.text);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();
		return {
			paragraphIndex: config.paragraphIndex,
			inserted: true,
			tracked: true,
		};
	});
}

async function handleInsertEndnote(args: unknown): Promise<unknown> {
	const config = args as { paragraphIndex: number; text: string };
	return runInWord(async (ctx: any) => {
		const Word: any = (window as any).Word;
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;

		const paras = ctx.document.body.paragraphs;
		paras.load("items");
		await ctx.sync();
		const para =
			paras.items[Math.min(config.paragraphIndex, paras.items.length - 1)];
		const range = para.getRange("End");
		range.insertEndnote(config.text);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();
		return {
			paragraphIndex: config.paragraphIndex,
			inserted: true,
			tracked: true,
		};
	});
}

async function handleInsertField(args: unknown): Promise<unknown> {
	const config = args as { fieldType: string; paragraphIndex?: number };
	const { fieldType, paragraphIndex } = config;

	return runInWord(async (ctx: any) => {
		const Word: any = (window as any).Word;
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;

		let range: any;
		if (paragraphIndex !== undefined) {
			const paras = ctx.document.body.paragraphs;
			paras.load("items");
			await ctx.sync();
			range =
				paras.items[Math.min(paragraphIndex, paras.items.length - 1)].getRange(
					"End",
				);
		} else {
			range = ctx.document.body.getRange("End");
		}

		// Map field type to enum
		const fieldMap: Record<string, string> = {
			TableOfContents: "TableOfContents",
			PageNumber: "Page",
			NumPages: "NumPages",
			Date: "Date",
			Time: "Time",
			FileName: "FileName",
			Author: "Author",
		};
		const ft = fieldMap[fieldType] || fieldType;
		range.insertField("End", ft);
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();
		return { fieldType, inserted: true, tracked: true };
	});
}

// ── Phase 18: Content Controls ───────────────────────────────────

async function handleGetContentControls(_args: unknown): Promise<unknown> {
	return runInWord(async (ctx: any) => {
		const ccs = ctx.document.contentControls;
		ccs.load("items");
		await ctx.sync();

		const result: any[] = [];
		for (const cc of ccs.items) {
			cc.load(["title", "tag"]);
			const range = cc.getRange("Whole");
			range.load("text");
		}
		await ctx.sync();
		for (const cc of ccs.items) {
			const range = cc.getRange("Whole");
			result.push({
				title: cc.title || "",
				tag: cc.tag || "",
				text: range.text,
			});
		}

		return { contentControls: result, count: result.length };
	});
}

async function handleInsertContentControl(args: unknown): Promise<unknown> {
	const config = args as {
		title: string;
		tag?: string;
		fromParagraph: number;
		toParagraph?: number;
	};
	const { title, tag, fromParagraph, toParagraph } = config;

	return runInWord(async (ctx: any) => {
		const Word: any = (window as any).Word;
		const originalMode = ctx.document.changeTrackingMode;
		ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly;

		const paras = ctx.document.body.paragraphs;
		paras.load("items");
		await ctx.sync();

		const end = toParagraph ?? fromParagraph;
		const startPara = paras.items[Math.max(0, fromParagraph)];
		const endPara = paras.items[Math.min(end, paras.items.length - 1)];
		const startRange = startPara.getRange("Start");
		const endRange = endPara.getRange("End");
		const range = startRange.expandTo(endRange);

		const cc = range.insertContentControl();
		cc.title = title;
		if (tag) cc.tag = tag;
		await ctx.sync();

		ctx.document.changeTrackingMode = originalMode;
		await ctx.sync();
		return {
			title,
			tag: tag || "",
			fromParagraph,
			toParagraph: end,
			inserted: true,
			tracked: true,
		};
	});
}

// ── Formatting (read/write) ──────────────────────────────────────

async function handleGetFormatting(args: unknown): Promise<unknown> {
	const config = args as {
		paragraphIndex?: number;
		fromParagraph?: number;
		toParagraph?: number;
	};

	const from = config.fromParagraph ?? config.paragraphIndex ?? 0;

	return runInWord(async (ctx) => {
		const paras = ctx.document.body.paragraphs;
		paras.load("items");
		await ctx.sync();

		const total = paras.items.length;
		const to = Math.min(config.toParagraph ?? from, total - 1);

		if (from < 0 || from >= total) {
			return { error: `paragraphIndex ${from} out of range (0–${total - 1})`, errorCode: "INVALID_PARAMETER" };
		}

		const results: unknown[] = [];
		for (let i = from; i <= to; i++) {
			const p = paras.items[i];
			p.load("text,style,alignment,firstLineIndent,leftIndent,rightIndent,lineSpacing,spaceAfter,spaceBefore,outlineLevel");
			p.font.load("name,size,bold,italic,underline,color,strikeThrough,doubleStrikeThrough,subscript,superscript,highlightColor");
		}
		await ctx.sync();

		for (let i = from; i <= to; i++) {
			const p = paras.items[i];
			results.push({
				paragraphIndex: i,
				text: String(p.text ?? ""),
				paragraph: {
					style: String(p.style ?? ""),
					alignment: String(p.alignment ?? ""),
					firstLineIndent: p.firstLineIndent ?? 0,
					leftIndent: p.leftIndent ?? 0,
					rightIndent: p.rightIndent ?? 0,
					lineSpacing: p.lineSpacing ?? null,
					spaceBefore: p.spaceBefore ?? 0,
					spaceAfter: p.spaceAfter ?? 0,
					outlineLevel: p.outlineLevel ?? 0,
				},
				font: {
					name: String(p.font.name ?? ""),
					size: p.font.size ?? null,
					bold: p.font.bold ?? false,
					italic: p.font.italic ?? false,
					underline: String(p.font.underline ?? "None"),
					color: String(p.font.color ?? ""),
					strikeThrough: p.font.strikeThrough ?? false,
					doubleStrikeThrough: p.font.doubleStrikeThrough ?? false,
					subscript: p.font.subscript ?? false,
					superscript: p.font.superscript ?? false,
					highlightColor: String(p.font.highlightColor ?? ""),
				},
			});
		}

		return { paragraphs: results };
	});
}

async function handleSetFormatting(args: unknown): Promise<unknown> {
	const config = args as {
		paragraphIndex?: number;
		fromParagraph?: number;
		toParagraph?: number;
		applyToSelection?: boolean;
		// paragraph formatting
		style?: string;
		alignment?: string;
		firstLineIndent?: number;
		leftIndent?: number;
		rightIndent?: number;
		lineSpacing?: number;
		spaceBefore?: number;
		spaceAfter?: number;
		// font formatting
		fontName?: string;
		fontSize?: number;
		bold?: boolean;
		italic?: boolean;
		underline?: string;
		color?: string;
		strikeThrough?: boolean;
		doubleStrikeThrough?: boolean;
		subscript?: boolean;
		superscript?: boolean;
		highlightColor?: string;
	};

	return runInWord(async (ctx) => {
		let targets: any[];

		if (config.applyToSelection) {
			const sel = ctx.document.getSelection();
			const selParas = sel.paragraphs;
			selParas.load("items");
			await ctx.sync();
			targets = selParas.items;
		} else {
			const paras = ctx.document.body.paragraphs;
			paras.load("items");
			await ctx.sync();

			const total = paras.items.length;
			const from = config.fromParagraph ?? config.paragraphIndex ?? 0;
			const to = Math.min(config.toParagraph ?? from, total - 1);

			if (from < 0 || from >= total) {
				return { error: `paragraphIndex ${from} out of range (0–${total - 1})`, errorCode: "INVALID_PARAMETER" };
			}
			targets = paras.items.slice(from, to + 1);
		}

		for (const p of targets) {
			// Paragraph-level properties
			if (config.style !== undefined)           p.style = config.style;
			if (config.alignment !== undefined)       p.alignment = config.alignment;
			if (config.firstLineIndent !== undefined) p.firstLineIndent = config.firstLineIndent;
			if (config.leftIndent !== undefined)      p.leftIndent = config.leftIndent;
			if (config.rightIndent !== undefined)     p.rightIndent = config.rightIndent;
			if (config.lineSpacing !== undefined)     p.lineSpacing = config.lineSpacing;
			if (config.spaceBefore !== undefined)     p.spaceBefore = config.spaceBefore;
			if (config.spaceAfter !== undefined)      p.spaceAfter = config.spaceAfter;
			// Font properties
			if (config.fontName !== undefined)           p.font.name = config.fontName;
			if (config.fontSize !== undefined)           p.font.size = config.fontSize;
			if (config.bold !== undefined)               p.font.bold = config.bold;
			if (config.italic !== undefined)             p.font.italic = config.italic;
			if (config.underline !== undefined)          p.font.underline = config.underline;
			if (config.color !== undefined)              p.font.color = config.color;
			if (config.strikeThrough !== undefined)      p.font.strikeThrough = config.strikeThrough;
			if (config.doubleStrikeThrough !== undefined) p.font.doubleStrikeThrough = config.doubleStrikeThrough;
			if (config.subscript !== undefined)          p.font.subscript = config.subscript;
			if (config.superscript !== undefined)        p.font.superscript = config.superscript;
			if (config.highlightColor !== undefined)     p.font.highlightColor = config.highlightColor;
		}

		await ctx.sync();

		return {
			applied: true,
			paragraphCount: targets.length,
			applyToSelection: config.applyToSelection ?? false,
		};
	});
}

// ── Comments ─────────────────────────────────────────────────────

async function handleGetComments(args: unknown): Promise<unknown> {
	const config = args as { includeResolved?: boolean };
	const includeResolved = config.includeResolved ?? true;

	return runInWord(async (ctx) => {
		const comments = ctx.document.body.getComments();
		comments.load("items");
		await ctx.sync();

		// Load scalar properties on all comments
		for (const c of comments.items) {
			c.load("id,content,authorName,authorEmail,creationDate,resolved");
			c.replies.load("items");
		}
		await ctx.sync();

		// Load reply properties
		for (const c of comments.items) {
			for (const r of c.replies.items) {
				r.load("id,content,authorName,authorEmail,creationDate");
			}
			// Load anchor text
			c.getRange().load("text");
		}
		await ctx.sync();

		const result = [];
		for (const c of comments.items) {
			if (!includeResolved && c.resolved) continue;
			const anchorText = (() => {
				try { return String(c.getRange().text ?? "").slice(0, 80); }
				catch { return ""; }
			})();
			result.push({
				id: c.id,
				author: c.authorName,
				email: c.authorEmail,
				date: c.creationDate,
				text: c.content,
				resolved: c.resolved,
				anchorText,
				replies: c.replies.items.map((r: any) => ({
					id: r.id,
					author: r.authorName,
					email: r.authorEmail,
					date: r.creationDate,
					text: r.content,
				})),
			});
		}

		return { commentCount: result.length, comments: result };
	});
}

async function handleEditComment(args: unknown): Promise<unknown> {
	const config = args as { commentId?: string; text?: string };
	const { commentId = "", text = "" } = config;
	if (!commentId) return { error: "commentId is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const comments = ctx.document.body.getComments();
		comments.load("items");
		await ctx.sync();
		for (const c of comments.items) c.load("id");
		await ctx.sync();

		const comment = comments.items.find((c: any) => c.id === commentId);
		if (!comment) return { error: `Comment '${commentId}' not found`, errorCode: "NOT_FOUND" };

		comment.content = text;
		await ctx.sync();
		return { commentId, updated: true };
	});
}

async function handleResolveComment(args: unknown): Promise<unknown> {
	const config = args as { commentId?: string; resolved?: boolean };
	const { commentId = "", resolved = true } = config;
	if (!commentId) return { error: "commentId is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const comments = ctx.document.body.getComments();
		comments.load("items");
		await ctx.sync();
		for (const c of comments.items) c.load("id");
		await ctx.sync();

		const comment = comments.items.find((c: any) => c.id === commentId);
		if (!comment) return { error: `Comment '${commentId}' not found`, errorCode: "NOT_FOUND" };

		comment.resolved = resolved;
		await ctx.sync();
		return { commentId, resolved };
	});
}

async function handleDeleteComment(args: unknown): Promise<unknown> {
	const config = args as { commentId?: string };
	const { commentId = "" } = config;
	if (!commentId) return { error: "commentId is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const comments = ctx.document.body.getComments();
		comments.load("items");
		await ctx.sync();
		for (const c of comments.items) c.load("id");
		await ctx.sync();

		const comment = comments.items.find((c: any) => c.id === commentId);
		if (!comment) return { error: `Comment '${commentId}' not found`, errorCode: "NOT_FOUND" };

		comment.delete();
		await ctx.sync();
		return { commentId, deleted: true };
	});
}

async function handleReplyToComment(args: unknown): Promise<unknown> {
	const config = args as { commentId?: string; text?: string };
	const { commentId = "", text = "" } = config;
	if (!commentId) return { error: "commentId is required", errorCode: "INVALID_PARAMETER" };
	if (!text) return { error: "text is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const comments = ctx.document.body.getComments();
		comments.load("items");
		await ctx.sync();
		for (const c of comments.items) c.load("id");
		await ctx.sync();

		const comment = comments.items.find((c: any) => c.id === commentId);
		if (!comment) return { error: `Comment '${commentId}' not found`, errorCode: "NOT_FOUND" };

		const reply = comment.reply(text);
		await ctx.sync();
		reply.load("id");
		await ctx.sync();
		return { commentId, replyId: reply.id, added: true };
	});
}

async function handleEditReply(args: unknown): Promise<unknown> {
	const config = args as { commentId?: string; replyId?: string; text?: string };
	const { commentId = "", replyId = "", text = "" } = config;
	if (!commentId) return { error: "commentId is required", errorCode: "INVALID_PARAMETER" };
	if (!replyId) return { error: "replyId is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const comments = ctx.document.body.getComments();
		comments.load("items");
		await ctx.sync();
		for (const c of comments.items) {
			c.load("id");
			c.replies.load("items");
		}
		await ctx.sync();

		const comment = comments.items.find((c: any) => c.id === commentId);
		if (!comment) return { error: `Comment '${commentId}' not found`, errorCode: "NOT_FOUND" };

		for (const r of comment.replies.items) r.load("id");
		await ctx.sync();

		const reply = comment.replies.items.find((r: any) => r.id === replyId);
		if (!reply) return { error: `Reply '${replyId}' not found`, errorCode: "NOT_FOUND" };

		reply.content = text;
		await ctx.sync();
		return { commentId, replyId, updated: true };
	});
}

async function handleDeleteReply(args: unknown): Promise<unknown> {
	const config = args as { commentId?: string; replyId?: string };
	const { commentId = "", replyId = "" } = config;
	if (!commentId) return { error: "commentId is required", errorCode: "INVALID_PARAMETER" };
	if (!replyId) return { error: "replyId is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const comments = ctx.document.body.getComments();
		comments.load("items");
		await ctx.sync();
		for (const c of comments.items) {
			c.load("id");
			c.replies.load("items");
		}
		await ctx.sync();

		const comment = comments.items.find((c: any) => c.id === commentId);
		if (!comment) return { error: `Comment '${commentId}' not found`, errorCode: "NOT_FOUND" };

		for (const r of comment.replies.items) r.load("id");
		await ctx.sync();

		const reply = comment.replies.items.find((r: any) => r.id === replyId);
		if (!reply) return { error: `Reply '${replyId}' not found`, errorCode: "NOT_FOUND" };

		reply.delete();
		await ctx.sync();
		return { commentId, replyId, deleted: true };
	});
}

// ── Styles ────────────────────────────────────────────────────────

function applyFontPropsToStyle(styleFont: any, cfg: Record<string, unknown>): void {
	if (cfg.fontName !== undefined)           styleFont.name = cfg.fontName;
	if (cfg.fontSize !== undefined)           styleFont.size = cfg.fontSize;
	if (cfg.bold !== undefined)               styleFont.bold = cfg.bold;
	if (cfg.italic !== undefined)             styleFont.italic = cfg.italic;
	if (cfg.underline !== undefined)          styleFont.underline = cfg.underline;
	if (cfg.color !== undefined)              styleFont.color = cfg.color;
	if (cfg.strikeThrough !== undefined)      styleFont.strikeThrough = cfg.strikeThrough;
	if (cfg.doubleStrikeThrough !== undefined) styleFont.doubleStrikeThrough = cfg.doubleStrikeThrough;
	if (cfg.subscript !== undefined)          styleFont.subscript = cfg.subscript;
	if (cfg.superscript !== undefined)        styleFont.superscript = cfg.superscript;
}

function applyParaPropsToStyle(stylePF: any, cfg: Record<string, unknown>): void {
	if (cfg.alignment !== undefined)       stylePF.alignment = cfg.alignment;
	if (cfg.firstLineIndent !== undefined) stylePF.firstLineIndent = cfg.firstLineIndent;
	if (cfg.leftIndent !== undefined)      stylePF.leftIndent = cfg.leftIndent;
	if (cfg.rightIndent !== undefined)     stylePF.rightIndent = cfg.rightIndent;
	if (cfg.lineSpacing !== undefined)     stylePF.lineSpacing = cfg.lineSpacing;
	if (cfg.spaceBefore !== undefined)     stylePF.spaceBeforeParagraph = cfg.spaceBefore;
	if (cfg.spaceAfter !== undefined)      stylePF.spaceAfterParagraph = cfg.spaceAfter;
}

async function handleGetStyles(args: unknown): Promise<unknown> {
	const config = args as { type?: string; inUseOnly?: boolean };
	const { type, inUseOnly = false } = config;

	return runInWord(async (ctx) => {
		const styles = ctx.document.getStyles();
		styles.load("items");
		await ctx.sync();

		for (const s of styles.items) {
			s.load("name,nameLocal,type,builtIn,inUse,styleId");
			s.font.load("name,size,bold,italic,underline,color,strikeThrough,doubleStrikeThrough,subscript,superscript");
			try { s.paragraphFormat.load("alignment,firstLineIndent,leftIndent,rightIndent,lineSpacing,spaceBeforeParagraph,spaceAfterParagraph"); }
			catch { /* paragraphFormat not available for character styles */ }
		}
		await ctx.sync();

		const result = [];
		for (const s of styles.items) {
			if (inUseOnly && !s.inUse) continue;
			const styleName = (s as any).name || (s as any).nameLocal || "";
			const styleType = (s as any).type || "";
			if (type && styleType.toLowerCase() !== type.toLowerCase()) continue;

			const entry: Record<string, unknown> = {
				name: styleName,
				styleId: (s as any).styleId,
				type: styleType,
				builtIn: s.builtIn,
				inUse: s.inUse,
				font: {
					name: s.font.name,
					size: s.font.size,
					bold: s.font.bold,
					italic: s.font.italic,
					underline: s.font.underline,
					color: s.font.color,
					strikeThrough: s.font.strikeThrough,
					doubleStrikeThrough: s.font.doubleStrikeThrough,
					subscript: s.font.subscript,
					superscript: s.font.superscript,
				},
			};
			try {
				entry.paragraphFormat = {
					alignment: s.paragraphFormat.alignment,
					firstLineIndent: s.paragraphFormat.firstLineIndent,
					leftIndent: s.paragraphFormat.leftIndent,
					rightIndent: s.paragraphFormat.rightIndent,
					lineSpacing: s.paragraphFormat.lineSpacing,
					spaceBefore: s.paragraphFormat.spaceBeforeParagraph,
					spaceAfter: s.paragraphFormat.spaceAfterParagraph,
				};
			} catch { /* character styles have no paragraphFormat */ }
			result.push(entry);
		}

		return { styleCount: result.length, styles: result };
	});
}

async function handleModifyStyle(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { styleName } = config;
	if (!styleName || typeof styleName !== "string")
		return { error: "styleName is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		const styles = ctx.document.getStyles();
		styles.load("items");
		await ctx.sync();
		for (const s of styles.items) s.load("name,nameLocal");
		await ctx.sync();

		const style = styles.items.find((s: any) => (s.name || s.nameLocal) === styleName);
		if (!style) return { error: `Style '${styleName}' not found`, errorCode: "NOT_FOUND" };

		applyFontPropsToStyle(style.font, config);
		try { applyParaPropsToStyle(style.paragraphFormat, config); }
		catch { /* character style — no paragraphFormat */ }

		await ctx.sync();
		return { styleName, modified: true };
	});
}

async function handleCreateStyle(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { styleName, styleType = "Paragraph", baseStyle } = config;
	if (!styleName || typeof styleName !== "string")
		return { error: "styleName is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		let style: any;
		try {
			style = ctx.document.addStyle(styleName as string, styleType as string);
		} catch (e) {
			return { error: `Could not create style: ${e instanceof Error ? e.message : String(e)}`, errorCode: "HOST_NOT_AVAILABLE" };
		}

		if (baseStyle && typeof baseStyle === "string") {
			try { style.baseStyle = baseStyle; } catch { /* ignore if base style not found */ }
		}

		applyFontPropsToStyle(style.font, config);
		try { applyParaPropsToStyle(style.paragraphFormat, config); }
		catch { /* character style — no paragraphFormat */ }

		await ctx.sync();
		return { styleName, styleType, baseStyle: baseStyle ?? null, created: true };
	});
}

async function handleCreateAndRemapStyle(args: unknown): Promise<unknown> {
	const config = args as Record<string, unknown>;
	const { baseStyleName, newStyleName } = config;
	if (!baseStyleName || typeof baseStyleName !== "string")
		return { error: "baseStyleName is required", errorCode: "INVALID_PARAMETER" };
	if (!newStyleName || typeof newStyleName !== "string")
		return { error: "newStyleName is required", errorCode: "INVALID_PARAMETER" };

	return runInWord(async (ctx) => {
		// Create new style inheriting from base
		const newStyle: any = ctx.document.addStyle(newStyleName, "Paragraph");
		newStyle.baseStyle = baseStyleName;
		applyFontPropsToStyle(newStyle.font, config);
		try { applyParaPropsToStyle(newStyle.paragraphFormat, config); } catch { /* ignore */ }
		try {
			await ctx.sync();
		} catch (e) {
			const msg = (e as any)?.debugInfo?.message || (e instanceof Error ? e.message : String(e));
			return { error: `Could not create style '${newStyleName}': ${msg}`, errorCode: "HOST_NOT_AVAILABLE" };
		}

		// Remap all paragraphs that use baseStyleName to the new style
		const paragraphs = ctx.document.body.paragraphs;
		paragraphs.load("items");
		await ctx.sync();
		for (const p of paragraphs.items) p.load("style");
		await ctx.sync();

		let remapped = 0;
		for (const p of paragraphs.items) {
			if (p.style === baseStyleName) {
				p.style = newStyleName;
				remapped++;
			}
		}
		await ctx.sync();

		return { newStyleName, baseStyleName, remapped };
	});
}
