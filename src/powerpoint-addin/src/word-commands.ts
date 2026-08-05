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
			default:
				result = { error: `Unknown Word command: ${commandName}` };
		}

		if (result && typeof result === "object" && "error" in result) {
			success = false;
		}
	} catch (error) {
		const errorMessage = error instanceof Error ? error.message : String(error);
		console.error(`Word command ${commandId} failed:`, errorMessage);
		success = false;
		result = { error: errorMessage };
	}

	await reportResult(commandId, success, undefined, result);
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
	const config = args as { searchText?: string; matchCase?: boolean };
	const searchText = config.searchText ?? "";
	const matchCase = config.matchCase ?? false;

	if (!searchText) return { error: "searchText is required" };

	return runInWord(async (ctx) => {
		const searchResults = ctx.document.body.search(searchText, { matchCase });
		searchResults.load("items");
		await ctx.sync();

		const results: Array<{ index: number; text: string; matchText: string }> =
			[];

		// Load parent paragraphs for context
		for (let i = 0; i < searchResults.items.length; i++) {
			const range = searchResults.items[i];
			range.load("text");
		}
		await ctx.sync();

		for (let i = 0; i < searchResults.items.length; i++) {
			const range = searchResults.items[i];
			results.push({
				index: i,
				text: String(range.text || ""),
				matchText: String(range.text || ""),
			});
		}

		return {
			searchText,
			totalMatches: results.length,
			matches: results,
		};
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
	const config = args as { commentText?: string; paragraphIndex?: number };
	const { commentText = "", paragraphIndex } = config;

	return runInWord(async (ctx) => {
		if (paragraphIndex !== undefined) {
			// Comment on a specific paragraph
			const paragraphs = ctx.document.body.paragraphs;
			paragraphs.load("items");
			await ctx.sync();

			if (paragraphIndex < 0 || paragraphIndex >= paragraphs.items.length) {
				return { error: `Paragraph index ${paragraphIndex} out of range` };
			}

			const paragraph = paragraphs.items[paragraphIndex];
			const range = paragraph.getRange("Whole");
			range.insertComment(commentText);
			await ctx.sync();

			return { commentText, paragraphIndex, added: true };
		}

		// Comment on current selection
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
	const config = args as { paragraphIndex?: number; styleName?: string };
	const { paragraphIndex = 0, styleName = "" } = config;
	if (!styleName)
		return { error: "styleName is required", errorCode: "INVALID_PARAMETER" };

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
