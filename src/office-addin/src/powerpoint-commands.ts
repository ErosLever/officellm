/**
 * PowerPoint command handler using Office JS API.
 *
 * Critical API rules (verified by testing):
 * - getTextFrameOrNullObject() creates a NEW object each call — MUST store references
 * - ctx.load(tf, "isNullObject,textRange/text") — load on the context
 * - isNullObject must be explicitly loaded — not available by default
 * - slide.load("shapes/items/$none") loads collection items without properties
 * - shape.load("id,name,type,left,top,width,height,rotation") — comma-separated for direct props
 * - MUST sync() before reading ANY loaded property
 * - PowerPoint.run() rejects on failure — .catch() on the returned promise
 * - shape.getImageAsBase64() and slide.getImageAsBase64() return ClientResult<string> — sync() first
 * - shape.fill and textRange.font are nested objects — load via slash paths
 */

/// <reference types="@types/office-js" />

import { reportResult } from "./communication";

// ── Command dispatch ────────────────────────────────────────────

export async function processCommand(
	commandId: string,
	commandName: string,
	args: unknown,
): Promise<unknown> {
	let result: unknown;
	let success = true;

	try {
		switch (commandName) {
			// Read tools
			case "powerpoint_get_deck_outline":
				result = await handleGetDeckOutline(args);
				break;
			case "powerpoint_get_slide":
				result = await handleGetSlide(args);
				break;
			case "powerpoint_get_slide_image":
				result = await handleGetSlideImage(args);
				break;
			case "powerpoint_get_shape_image":
				result = await handleGetShapeImage(args);
				break;
			case "powerpoint_get_table":
				result = await handleGetTable(args);
				break;
			case "powerpoint_get_shape_paragraphs":
				result = await handleGetShapeParagraphs(args);
				break;
			case "powerpoint_get_shape_text_markdown":
				result = await handleGetShapeTextMarkdown(args);
				break;
			case "powerpoint_get_slide_text_markdown":
				result = await handleGetSlideTextMarkdown(args);
				break;
			case "powerpoint_get_selection":
				result = await handleGetSelection(args);
				break;
			case "powerpoint_get_speaker_notes":
				result = await handleGetSpeakerNotes(args);
				break;

			// Write tools
			case "powerpoint_update_shape_text":
				result = await handleUpdateShapeText(args);
				break;
			case "powerpoint_update_shape_properties":
				result = await handleUpdateShapeProperties(args);
				break;
			case "powerpoint_update_text_range_properties":
				result = await handleUpdateTextRangeProperties(args);
				break;
			case "powerpoint_insert_paragraph":
				result = await handleInsertParagraph(args);
				break;
			case "powerpoint_delete_paragraph":
				result = await handleDeleteParagraph(args);
				break;
			case "powerpoint_set_shape_text_markdown":
				result = await handleSetShapeTextMarkdown(args);
				break;
			case "powerpoint_update_speaker_notes":
				result = await handleUpdateSpeakerNotes(args);
				break;

			// Shape CRUD
			case "powerpoint_add_textbox":
				result = await handleAddTextbox(args);
				break;
			case "powerpoint_add_image":
				result = await handleAddImage(args);
				break;
			case "powerpoint_add_table":
				result = await handleAddTable(args);
				break;
			case "powerpoint_delete_shape":
				result = await handleDeleteShape(args);
				break;

			// Slide management
			case "powerpoint_add_slide":
				result = await handleAddSlide(args);
				break;
			case "powerpoint_set_slide_layout":
				result = await handleSetSlideLayout(args);
				break;
			case "powerpoint_delete_slide":
				result = await handleDeleteSlide(args);
				break;
			case "powerpoint_move_slide":
				result = await handleMoveSlide(args);
				break;
			case "powerpoint_duplicate_slide":
				result = await handleDuplicateSlide(args);
				break;
			case "powerpoint_export_slide_internal":
				result = await handleExportSlideInternal(args);
				break;
			case "powerpoint_import_slide_internal":
				result = await handleImportSlideInternal(args);
				break;

			// Phase 18: Tags & Metadata
			case "powerpoint_get_tags":
				result = await handleGetTags(args);
				break;
			case "powerpoint_set_tag":
				result = await handleSetTag(args);
				break;
			case "powerpoint_delete_slides_by_tag":
				result = await handleDeleteSlidesByTag(args);
				break;

			// Phase 18: Shape Formatting
			case "powerpoint_set_shape_fill":
				result = await handleSetShapeFill(args);
				break;
			case "powerpoint_set_shape_line":
				result = await handleSetShapeLine(args);
				break;
			case "powerpoint_set_shape_rotation":
				result = await handleSetShapeRotation(args);
				break;

			// Phase 18: Geometric Shapes & Lines
			case "powerpoint_add_geometric_shape":
				result = await handleAddGeometricShape(args);
				break;
			case "powerpoint_add_line":
				result = await handleAddLine(args);
				break;

			// Phase 18: Slide Merge
			case "powerpoint_insert_slides_from_file":
				result = await handleInsertSlidesFromFile(args);
				break;

			// Phase 18: Layouts & Theme
			case "powerpoint_get_layouts":
				result = await handleGetLayouts(args);
				break;
			case "powerpoint_get_theme_colors":
				result = await handleGetThemeColors(args);
				break;
			case "powerpoint_group_shapes":
				result = await handleGroupShapes(args);
				break;
			case "powerpoint_ungroup_shape":
				result = await handleUngroupShape(args);
				break;
			default:
				result = { error: `Unknown command: ${commandName}` };
		}

		if (result && typeof result === "object" && "error" in result) {
			success = false;
		}
	} catch (error) {
		const officeError = error as any;
		let errorMessage = error instanceof Error ? error.message : String(error);
		if (officeError?.debugInfo) {
			errorMessage += ` | debugInfo: ${JSON.stringify(officeError.debugInfo)}`;
		}
		console.error(`Command ${commandId} failed:`, errorMessage);
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

const TITLE_SKIP_PATTERNS = [
	/slide\s*number/i,
	/footer/i,
	/header/i,
	/date/i,
	/background/i,
];

function isContentShape(name: string): boolean {
	return !TITLE_SKIP_PATTERNS.some((p) => p.test(name));
}

/**
 * Wraps PowerPoint.run() in a Promise that properly rejects on errors.
 */
function runInPowerPoint<T>(fn: (ctx: any) => Promise<T>): Promise<T> {
	return new Promise<T>((resolve, reject) => {
		const PowerPoint: any = (window as any).PowerPoint;
		if (!PowerPoint || typeof PowerPoint.run !== "function") {
			reject(new Error("PowerPoint.run() not available"));
			return;
		}

		PowerPoint.run(async (ctx: any) => {
			resolve(await fn(ctx));
		}).catch(reject);
	});
}

function safeStr(val: any, fallback = ""): string {
	return val != null ? String(val) : fallback;
}

function safeNum(val: any, fallback = 0): number {
	return typeof val === "number" ? val : fallback;
}

const PARAGRAPH_ALIGNMENT_NAMES = [
	"Left",
	"Center",
	"Right",
	"Justify",
	"JustifyLow",
	"Distributed",
	"ThaiDistributed",
];

const BULLET_TYPE_NAMES = ["Unsupported", "None", "Numbered", "Unnumbered"];

const BULLET_STYLE_NAMES = [
	"Unsupported",
	"AlphabetLowercasePeriod",
	"AlphabetUppercasePeriod",
	"ArabicNumeralParenthesisRight",
	"ArabicNumeralPeriod",
	"RomanLowercaseParenthesesBoth",
	"RomanLowercaseParenthesisRight",
	"RomanLowercasePeriod",
	"RomanUppercasePeriod",
	"AlphabetLowercaseParenthesesBoth",
	"AlphabetLowercaseParenthesisRight",
	"AlphabetUppercaseParenthesesBoth",
	"AlphabetUppercaseParenthesisRight",
	"ArabicNumeralParenthesesBoth",
	"ArabicNumeralPlain",
	"RomanUppercaseParenthesesBoth",
	"RomanUppercaseParenthesisRight",
	"SimplifiedChinesePlain",
	"SimplifiedChinesePeriod",
	"CircleNumberDoubleBytePlain",
	"CircleNumberWideDoubleByteWhitePlain",
	"CircleNumberWideDoubleByteBlackPlain",
	"TraditionalChinesePlain",
	"TraditionalChinesePeriod",
	"ArabicAlphabetDash",
	"ArabicAbjadDash",
	"HebrewAlphabetDash",
	"KanjiKoreanPlain",
	"KanjiKoreanPeriod",
	"ArabicDoubleBytePlain",
	"ArabicDoubleBytePeriod",
	"ThaiAlphabetPeriod",
	"ThaiAlphabetParenthesisRight",
	"ThaiAlphabetParenthesesBoth",
	"ThaiNumeralPeriod",
	"ThaiNumeralParenthesisRight",
	"ThaiNumeralParenthesesBoth",
	"HindiAlphabetPeriod",
	"HindiNumeralPeriod",
	"KanjiSimplifiedChineseDoubleBytePeriod",
	"HindiNumeralParenthesisRight",
	"HindiAlphabet1Period",
];

// The 16 Latin-representable BulletStyle values, mapped to a (family, wrapper)
// pair — the single source of truth used by both paragraphsToMarkdown (encoder)
// and markdownToParagraphSpecs (decoder) so they can't drift apart. The other
// 25 styles (non-Latin scripts, circled-digit glyphs, Dash wrappers) have no
// entry here and fall back to plain "${count}." rendering / are unreachable
// from the decoder.
export type BulletFamily = "arabic" | "alphaLower" | "alphaUpper" | "romanLower" | "romanUpper";
export type BulletWrapper = "plain" | "period" | "parenRight" | "parenBoth";

export const BULLET_STYLE_INFO: Record<string, { family: BulletFamily; wrapper: BulletWrapper }> = {
	ArabicNumeralPlain: { family: "arabic", wrapper: "plain" },
	ArabicNumeralPeriod: { family: "arabic", wrapper: "period" },
	ArabicNumeralParenthesisRight: { family: "arabic", wrapper: "parenRight" },
	ArabicNumeralParenthesesBoth: { family: "arabic", wrapper: "parenBoth" },
	AlphabetLowercasePeriod: { family: "alphaLower", wrapper: "period" },
	AlphabetLowercaseParenthesisRight: { family: "alphaLower", wrapper: "parenRight" },
	AlphabetLowercaseParenthesesBoth: { family: "alphaLower", wrapper: "parenBoth" },
	AlphabetUppercasePeriod: { family: "alphaUpper", wrapper: "period" },
	AlphabetUppercaseParenthesisRight: { family: "alphaUpper", wrapper: "parenRight" },
	AlphabetUppercaseParenthesesBoth: { family: "alphaUpper", wrapper: "parenBoth" },
	RomanLowercasePeriod: { family: "romanLower", wrapper: "period" },
	RomanLowercaseParenthesisRight: { family: "romanLower", wrapper: "parenRight" },
	RomanLowercaseParenthesesBoth: { family: "romanLower", wrapper: "parenBoth" },
	RomanUppercasePeriod: { family: "romanUpper", wrapper: "period" },
	RomanUppercaseParenthesisRight: { family: "romanUpper", wrapper: "parenRight" },
	RomanUppercaseParenthesesBoth: { family: "romanUpper", wrapper: "parenBoth" },
};

/** 1=a, 2=b, ..., 26=z, 27=aa, 28=ab, ... (bijective base-26, no zero digit). */
export function toAlphabetCounter(n: number, uppercase: boolean): string {
	let s = "";
	let rem = n;
	while (rem > 0) {
		const digit = (rem - 1) % 26;
		s = String.fromCharCode(97 + digit) + s;
		rem = Math.floor((rem - 1) / 26);
	}
	return uppercase ? s.toUpperCase() : s;
}

/**
 * Roman numeral using only i/v/x, range 1-19 — matches the decoder's
 * supported marker range. l/c/d/m are never treated as Roman (see
 * markdownToParagraphSpecs), so nothing above 19 needs to be representable.
 */
export function toRomanNumeral(n: number, uppercase: boolean): string {
	const tens = Math.floor(n / 10);
	const ones = n % 10;
	const onesMap = ["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix"];
	const s = (tens > 0 ? "x" : "") + onesMap[ones];
	return uppercase ? s.toUpperCase() : s;
}

// Family switch defaults to plain numeric ("arabic" and any unrecognized
// family) rather than roman — roman is the more surprising silent fallback,
// numeric is the safer "did nothing weird" behavior.
export function formatBulletMarker(family: BulletFamily, wrapper: BulletWrapper, n: number): string {
	const core =
		family === "alphaLower"
			? toAlphabetCounter(n, false)
			: family === "alphaUpper"
				? toAlphabetCounter(n, true)
				: family === "romanLower"
					? toRomanNumeral(n, false)
					: family === "romanUpper"
						? toRomanNumeral(n, true)
						: String(n);
	if (wrapper === "plain") return core;
	if (wrapper === "period") return core + ".";
	if (wrapper === "parenRight") return core + ")";
	return "(" + core + ")";
}

// Live Office.js (observed on Mac desktop, PowerPointApi 1.10) returns some
// enum-typed read properties (paragraphFormat.horizontalAlignment,
// bulletFormat.type, bulletFormat.style) as the 0-based index of the enum
// member's declaration order in @types/office-js — e.g. "2" for
// BulletType.numbered — instead of the named string ("Numbered") the type
// declarations promise. Map the numeric-string form back to its name; pass
// through anything already named, empty, or out of range unchanged.
function normalizeEnumString(val: string, names: string[]): string {
	if (/^\d+$/.test(val)) {
		const index = Number(val);
		if (index >= 0 && index < names.length) return names[index];
	}
	return val;
}

/**
 * Splits shape text into paragraphs on `\r` (paragraph boundary; `\n` is a
 * soft line break within a paragraph). `length` excludes the trailing `\r`.
 */
export function splitParagraphs(
	text: string,
): { text: string; start: number; length: number }[] {
	// Each paragraph's start is the previous paragraph's end plus the `\r`
	// delimiter that was consumed by split() — reduce carries that running
	// offset forward instead of a mutable loop counter.
	return text
		.split("\r")
		.reduce<{ text: string; start: number; length: number }[]>((acc, part) => {
			const prev = acc[acc.length - 1];
			const start = prev ? prev.start + prev.length + 1 : 0;
			return [...acc, { text: part, start, length: part.length }];
		}, []);
}

const TEXT_RANGE_PROP_PATH = [
	"font/name",
	"font/size",
	"font/bold",
	"font/italic",
	"font/color",
	"font/underline",
	"font/strikethrough",
	"font/doubleStrikethrough",
	"font/allCaps",
	"font/smallCaps",
	"font/subscript",
	"font/superscript",
	"paragraphFormat/horizontalAlignment",
	"paragraphFormat/indentLevel",
	"paragraphFormat/bulletFormat/type",
	"paragraphFormat/bulletFormat/style",
	"paragraphFormat/bulletFormat/visible",
].join(",");

type RangeProperties = {
	font: Record<string, unknown>;
	paragraphFormat: Record<string, unknown>;
};

// Reads a loaded TextRange (whole-shape range or a getSubstring() result)
// into the same flat font/paragraphFormat shape handleUpdateShapeProperties
// accepts, so getter output and setter input line up field-for-field.
function extractRangeProperties(range: any): RangeProperties {
	return {
		font: {
			fontName: safeStr(range.font.name),
			fontSize: safeNum(range.font.size),
			bold: !!range.font.bold,
			italic: !!range.font.italic,
			color: safeStr(range.font.color),
			underline: safeStr(range.font.underline),
			strikethrough: !!range.font.strikethrough,
			doubleStrikethrough: !!range.font.doubleStrikethrough,
			allCaps: !!range.font.allCaps,
			smallCaps: !!range.font.smallCaps,
			subscript: !!range.font.subscript,
			superscript: !!range.font.superscript,
		},
		paragraphFormat: {
			horizontalAlignment: normalizeEnumString(
				safeStr(range.paragraphFormat.horizontalAlignment),
				PARAGRAPH_ALIGNMENT_NAMES,
			),
			indentLevel: safeNum(range.paragraphFormat.indentLevel),
			bulletType: normalizeEnumString(
				safeStr(range.paragraphFormat.bulletFormat.type),
				BULLET_TYPE_NAMES,
			),
			bulletStyle: normalizeEnumString(
				safeStr(range.paragraphFormat.bulletFormat.style),
				BULLET_STYLE_NAMES,
			),
			bulletVisible: !!range.paragraphFormat.bulletFormat.visible,
		},
	};
}

// Declarative setters keyed by the same config property names used in
// RangeProperties/extractRangeProperties, so handleUpdateTextRangeProperties
// can apply "only the keys present in config" without an if-chain per field.
const FONT_SETTERS: Record<string, (font: any, value: any) => void> = {
	fontName: (font, v) => (font.name = v),
	fontSize: (font, v) => (font.size = v),
	bold: (font, v) => (font.bold = v),
	italic: (font, v) => (font.italic = v),
	color: (font, v) => (font.color = v),
	underline: (font, v) => (font.underline = v),
	strikethrough: (font, v) => (font.strikethrough = v),
	doubleStrikethrough: (font, v) => (font.doubleStrikethrough = v),
	allCaps: (font, v) => (font.allCaps = v),
	smallCaps: (font, v) => (font.smallCaps = v),
	subscript: (font, v) => (font.subscript = v),
	superscript: (font, v) => (font.superscript = v),
};

const PARAGRAPH_FORMAT_SETTERS: Record<string, (pf: any, value: any) => void> = {
	horizontalAlignment: (pf, v) => (pf.horizontalAlignment = v),
	indentLevel: (pf, v) => (pf.indentLevel = v),
};

const BULLET_FORMAT_SETTERS: Record<string, (bf: any, value: any) => void> = {
	bulletType: (bf, v) => (bf.type = v),
	bulletStyle: (bf, v) => (bf.style = v),
	bulletVisible: (bf, v) => (bf.visible = v),
};

/**
 * Applies every setter whose config key is present (!== undefined) to
 * `target`, returning the applied key names — replaces a per-field
 * `if (config.x !== undefined) { ...; updated.push("x") }` chain.
 */
function applyDefinedProperties(
	config: Record<string, unknown>,
	target: any,
	setters: Record<string, (target: any, value: any) => void>,
): string[] {
	return Object.keys(setters).filter((key) => config[key] !== undefined) // only requested keys
		.map((key) => {
			setters[key](target, config[key]);
			return key;
		});
}

/**
 * True if `config` defines at least one key from any of the given setter
 * maps — used to gate a load+sync+apply block behind "was anything actually
 * requested" without hand-listing the same setter keys in an OR-chain.
 */
function hasAnyDefinedKey(
	config: Record<string, unknown>,
	...setterMaps: Record<string, (...args: any[]) => void>[]
): boolean {
	return setterMaps.some((setters) => Object.keys(setters).some((key) => config[key] !== undefined));
}

/**
 * Applies font/paragraphFormat/bulletFormat overrides to a text range —
 * shared by every handler that restyles a `PowerPoint.TextRange` from a flat
 * config object (font/paragraphFormat/bulletFormat setter keys).
 *
 * `bulletFormat.visible` is written in its own sync, after `type`/`style`:
 * live PowerPoint silently drops the whole bulletFormat write (reverting to
 * Unnumbered) when `visible` is queued in the same batch as `type`/`style`
 * before a single sync — confirmed via manual testing against a live
 * document, not reproducible against the mock. Splitting the sync fixes it.
 */
async function applyRangeProperties(
	config: Record<string, unknown>,
	range: PowerPoint.TextRange,
	ctx: PowerPoint.RequestContext,
): Promise<string[]> {
	const { bulletVisible, ...rest } = config;

	const updated = [
		...applyDefinedProperties(config, range.font, FONT_SETTERS),
		...applyDefinedProperties(config, range.paragraphFormat, PARAGRAPH_FORMAT_SETTERS),
		...applyDefinedProperties(rest, range.paragraphFormat.bulletFormat, BULLET_FORMAT_SETTERS),
	];

	if (bulletVisible !== undefined) {
		await ctx.sync();
		updated.push(
			...applyDefinedProperties({ bulletVisible }, range.paragraphFormat.bulletFormat, BULLET_FORMAT_SETTERS),
		);
	}

	return updated;
}

/**
 * Diffs `full` against `base`, keeping only the categories/keys whose values
 * differ. Used to shrink a paragraph's complete properties down to just its
 * overrides relative to the shape's defaultProperties.
 */
function diffRangeProperties(
	full: RangeProperties,
	base: RangeProperties,
): Record<string, Record<string, unknown>> {
	// Two-level reduce: outer walks font/paragraphFormat, inner walks each
	// category's keys, keeping only ones whose stringified value changed.
	// Categories/keys that fully match `base` are omitted entirely rather
	// than kept as empty objects/false-y placeholders.
	return (["font", "paragraphFormat"] as const).reduce<Record<string, Record<string, unknown>>>(
		(diff, category) => {
			const catDiff = Object.keys(full[category]).reduce<Record<string, unknown>>(
				(acc, key) => {
					const same =
						JSON.stringify(full[category][key]) === JSON.stringify(base[category][key]);
					return same ? acc : { ...acc, [key]: full[category][key] };
				},
				{},
			);
			return Object.keys(catDiff).length > 0 ? { ...diff, [category]: catDiff } : diff;
		},
		{},
	);
}

/**
 * Slide indices are 1-based at the tool boundary (matching the PowerPoint UI's own
 * slide numbering) but 0-based internally (OfficeJS's pres.slides.items[i]). Convert
 * only at the point of native indexing — everywhere else the value stays 1-based.
 */
function toZeroBasedSlideIndex(oneBased: number): number {
	return oneBased - 1;
}
function toOneBasedSlideIndex(zeroBased: number): number {
	return zeroBased + 1;
}

/**
 * Resolves a 1-based slide index to a live Slide object.
 * Shared by every handler that operates on a single slide (no shape).
 */
async function resolveSlide(
	ctx: PowerPoint.RequestContext,
	slideIndex: number,
): Promise<{ slide: PowerPoint.Slide } | { error: string }> {
	const pres = ctx.presentation;
	pres.load("slides");
	await ctx.sync();

	if (slideIndex < 1 || slideIndex > pres.slides.items.length) {
		return { error: `Slide index ${slideIndex} out of range (1-${pres.slides.items.length})` };
	}

	return { slide: pres.slides.items[toZeroBasedSlideIndex(slideIndex)] };
}

/**
 * Resolves a 1-based slide index + shapeId (id or name) to a live Shape object.
 * Shared by every handler that operates on a single named/id'd shape.
 */
async function resolveShape(
	ctx: PowerPoint.RequestContext,
	slideIndex: number,
	shapeId: string,
): Promise<{ shape: PowerPoint.Shape } | { error: string }> {
	const pres = ctx.presentation;
	pres.load("slides");
	await ctx.sync();

	if (slideIndex < 1 || slideIndex > pres.slides.items.length) {
		return { error: `Slide index ${slideIndex} out of range (1-${pres.slides.items.length})` };
	}

	const slide = pres.slides.items[toZeroBasedSlideIndex(slideIndex)];
	slide.load("shapes/items/$none");
	await ctx.sync();

	for (const s of slide.shapes.items) {
		s.load("id,name");
	}
	await ctx.sync();

	const shape = slide.shapes.items.find(
		(s: any) => safeStr(s.id) === shapeId || safeStr(s.name) === shapeId,
	);
	return shape ? { shape } : { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
}

/**
 * Resolves the inherited font size for a shape whose own text range reports
 * PowerPoint's mixed/unresolved-size sentinel (0 after safeNum) — e.g. an
 * unfilled placeholder with no run-level font override, which renders at a
 * size set on the slide layout instead of the shape itself. Matches the
 * shape to its layout counterpart by placeholder type (e.g. "Title") and
 * reads that layout shape's own font size. Returns 0 if the shape isn't a
 * placeholder, has no layout counterpart, or the layout shape's size is
 * itself unresolved.
 *
 * Layout shapes are probed one at a time (rather than batched into a single
 * load/sync) because `placeholderFormat` throws a GeneralException for any
 * shape that isn't a placeholder (e.g. a decorative background shape) —
 * batching would fail the whole sync as soon as one non-placeholder shape
 * was hit.
 */
async function resolveInheritedFontSize(
	ctx: PowerPoint.RequestContext,
	shape: PowerPoint.Shape,
): Promise<number> {
	let placeholderType = "";
	try {
		const pf = shape.placeholderFormat;
		pf.load("type");
		await ctx.sync();
		placeholderType = safeStr(pf.type);
	} catch {
		return 0;
	}
	if (!placeholderType) return 0;

	const layout = shape.getParentSlide().layout;
	layout.load("shapes/items/$none");
	await ctx.sync();

	for (const layoutShape of layout.shapes.items) {
		let layoutType = "";
		try {
			const pf = layoutShape.placeholderFormat;
			pf.load("type");
			await ctx.sync();
			layoutType = safeStr(pf.type);
		} catch {
			continue;
		}
		if (layoutType !== placeholderType) continue;

		const tf = layoutShape.getTextFrameOrNullObject();
		ctx.load(tf, "isNullObject,textRange/font/size");
		await ctx.sync();
		return tf.isNullObject ? 0 : safeNum(tf.textRange.font.size);
	}

	return 0;
}

// ── Read tools ──────────────────────────────────────────────────

async function handleGetDeckOutline(args: unknown): Promise<unknown> {
	const config = args as { startSlide?: number; endSlide?: number };

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const totalSlides = pres.slides.items.length;

		const from = Math.max(0, toZeroBasedSlideIndex(config.startSlide ?? 1));
		const to = Math.min(totalSlides - 1, toZeroBasedSlideIndex(config.endSlide ?? totalSlides));
		if (from > to) {
			return { error: `Invalid slide range: startSlide (${toOneBasedSlideIndex(from)}) is after endSlide (${toOneBasedSlideIndex(to)}).` };
		}
		const indices: number[] = [];
		for (let i = from; i <= to; i++) indices.push(i);

		// Load shape items
		for (const i of indices) {
			pres.slides.items[i].load("shapes/items/$none");
		}
		await ctx.sync();

		// Load direct shape properties: id, name, type, left, top, width, height
		for (const i of indices) {
			for (const s of pres.slides.items[i].shapes.items) {
				s.load("id,name,type,left,top,width,height");
			}
		}
		await ctx.sync();

		// Build slide list
		const slideList = [];
		for (const i of indices) {
			const shapes = pres.slides.items[i].shapes.items;

			// Load text for title detection
			const tfs: any[] = [];
			for (const s of shapes) {
				const tf = s.getTextFrameOrNullObject();
				ctx.load(tf, "isNullObject,textRange/text");
				tfs.push(tf);
			}
			await ctx.sync();

			let title = "";
			const shapeData = [];
			for (let j = 0; j < shapes.length; j++) {
				const s = shapes[j];
				const tf = tfs[j];
				let text = "";
				if (!tf.isNullObject && tf.textRange?.text) {
					text = String(tf.textRange.text).trim();
				}

				const shapeName = safeStr(s.name);
				if (!title && text && isContentShape(shapeName)) {
					title = text;
				}

				shapeData.push({
					id: safeStr(s.id),
					name: shapeName,
					type: safeStr(s.type),
					left: safeNum(s.left),
					top: safeNum(s.top),
					width: safeNum(s.width),
					height: safeNum(s.height),
					text,
				});
			}

			slideList.push({
				index: toOneBasedSlideIndex(i),
				title: title || `Slide ${toOneBasedSlideIndex(i)}`,
				shapes: shapeData,
			});
		}

		return { documentName: "Presentation", totalSlides, slides: slideList };
	});
}

async function handleGetSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const slideIndex = config.slideIndex ?? 1;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;

		// Step 1: Load shape items + applied layout
		slide.load("shapes/items/$none");
		slide.layout.load("id,name");
		await ctx.sync();

		// Step 2: Load direct properties
		for (const s of slide.shapes.items) {
			s.load("id,name,type,left,top,width,height,rotation");
		}
		await ctx.sync();

		// Step 3: Load text frames + font + fill
		const textFrameData: any[] = [];
		const fillData: any[] = [];
		for (const s of slide.shapes.items) {
			// Text frame
			const tf = s.getTextFrameOrNullObject();
			ctx.load(
				tf,
				"isNullObject,textRange/text,textRange/font/name,textRange/font/size,textRange/font/bold,textRange/font/italic,textRange/font/color",
			);
			textFrameData.push(tf);

			// Fill
			ctx.load(s.fill, "foregroundColor,transparency");
			fillData.push(s.fill);
		}
		await ctx.sync();

		// Step 3.5: PowerPoint's whole-range font.size getter can report the
		// mixed/unresolved sentinel (0 after safeNum) even when every
		// character in the range shares the same real, explicit size (seen
		// live: an all-35pt title still reported the sentinel for the whole
		// range). Sample a single character for any text frame hitting the
		// sentinel before falling back to the shape's layout-inherited size.
		const sampleByIndex = new Map<number, any>();
		for (let i = 0; i < textFrameData.length; i++) {
			const tf = textFrameData[i];
			if (tf.isNullObject || !tf.textRange?.font) continue;
			const text = safeStr(tf.textRange.text);
			if (safeNum(tf.textRange.font.size) === 0 && text.length > 0) {
				const sample = tf.textRange.getSubstring(0, 1);
				ctx.load(sample, "font/size");
				sampleByIndex.set(i, sample);
			}
		}
		if (sampleByIndex.size > 0) await ctx.sync();

		// Step 4: Build result
		let slideTitle = "";
		const shapeList = [];

		for (let i = 0; i < slide.shapes.items.length; i++) {
			const s = slide.shapes.items[i];
			const tf = textFrameData[i];
			const fill = fillData[i];

			let text = "";
			let font: Record<string, unknown> | undefined;
			if (!tf.isNullObject) {
				text = safeStr(tf.textRange?.text);
				if (tf.textRange?.font) {
					let size = safeNum(tf.textRange.font.size);
					if (size === 0) {
						const sample = sampleByIndex.get(i);
						size = sample ? safeNum(sample.font.size) : 0;
					}
					if (size === 0) {
						size = await resolveInheritedFontSize(ctx, s);
					}
					font = {
						name: safeStr(tf.textRange.font.name),
						size,
						bold: !!tf.textRange.font.bold,
						italic: !!tf.textRange.font.italic,
						color: safeStr(tf.textRange.font.color),
					};
				}
			}

			const shapeName = safeStr(s.name);
			if (!slideTitle && text.trim() && isContentShape(shapeName)) {
				slideTitle = text.trim();
			}

			shapeList.push({
				id: safeStr(s.id),
				name: shapeName,
				type: safeStr(s.type),
				left: safeNum(s.left),
				top: safeNum(s.top),
				width: safeNum(s.width),
				height: safeNum(s.height),
				rotation: safeNum(s.rotation),
				text,
				font,
				fillColor: safeStr(fill.foregroundColor),
				fillTransparency: safeNum(fill.transparency),
			});
		}

		return {
			slideIndex,
			title: slideTitle || `Slide ${slideIndex}`,
			shapes: shapeList,
			layout: { id: safeStr(slide.layout.id), name: safeStr(slide.layout.name) },
		};
	});
}

async function handleGetSlideImage(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		width?: number;
		height?: number;
	};
	const slideIndex = config.slideIndex ?? 1;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;
		const imageOptions: any = {};
		if (config.width) imageOptions.width = config.width;
		if (config.height) imageOptions.height = config.height;

		const imageResult = slide.getImageAsBase64(imageOptions);
		await ctx.sync();

		return {
			slideIndex,
			image: `data:image/png;base64,${imageResult.value}`,
		};
	});
}

async function handleGetShapeImage(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		width?: number;
		height?: number;
	};
	const slideIndex = config.slideIndex ?? 1;
	const shapeId = config.shapeId ?? "";

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;
		slide.load("shapes/items/$none");
		await ctx.sync();

		// Find shape by ID or name
		let targetShape: any = null;
		for (const s of slide.shapes.items) {
			s.load("id,name");
		}
		await ctx.sync();

		for (const s of slide.shapes.items) {
			if (safeStr(s.id) === shapeId || safeStr(s.name) === shapeId) {
				targetShape = s;
				break;
			}
		}

		if (!targetShape) {
			return { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
		}

		const imageOptions: any = {};
		if (config.width) imageOptions.width = config.width;
		if (config.height) imageOptions.height = config.height;

		const imageResult = targetShape.getImageAsBase64(imageOptions);
		await ctx.sync();

		return {
			slideIndex,
			shapeId,
			image: `data:image/png;base64,${imageResult.value}`,
		};
	});
}

async function handleGetTable(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string };
	const slideIndex = config.slideIndex ?? 1;
	const shapeId = config.shapeId ?? "";

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;
		slide.load("shapes/items/$none");
		await ctx.sync();

		// Find the shape
		let targetShape: any = null;
		for (const s of slide.shapes.items) {
			s.load("id,name,type");
		}
		await ctx.sync();

		for (const s of slide.shapes.items) {
			if (safeStr(s.id) === shapeId || safeStr(s.name) === shapeId) {
				targetShape = s;
				break;
			}
		}

		if (!targetShape) {
			return { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
		}

		// Get the table
		const table = targetShape.getTable();
		table.load("rowCount,columnCount");
		await ctx.sync();

		const rows = table.rowCount;
		const cols = table.columnCount;

		// Read all cells in a single sync
		const cellRefs: any[] = [];
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				const cell = table.getCellOrNullObject(r, c);
				cell.load("text");
				cellRefs.push(cell);
			}
		}
		await ctx.sync();

		// Fill in the cells array
		const cells: string[][] = [];
		let idx = 0;
		for (let r = 0; r < rows; r++) {
			const row: string[] = [];
			for (let c = 0; c < cols; c++) {
				const cell = cellRefs[idx++];
				row.push(safeStr(cell.text));
			}
			cells.push(row);
		}

		return {
			slideIndex,
			shapeId,
			rowCount: rows,
			columnCount: cols,
			cells,
		};
	});
}

/**
 * Core of handleGetShapeParagraphs, operating on an already-resolved shape
 * rather than re-resolving by slideIndex/shapeId — shared with the shape
 * snapshot logic (powerpoint_copy_shape), which already holds a live shape
 * reference (including group-child shapes not addressable via resolveShape).
 */
async function getShapeParagraphsData(
	shape: PowerPoint.Shape,
	ctx: PowerPoint.RequestContext,
	label = "Shape",
): Promise<{ error: string } | {
	fullText: string;
	defaultProperties: RangeProperties;
	propertyGroups: { groupId: number; properties: Record<string, unknown> }[];
	paragraphs: { index: number; text: string; start: number; length: number; groupId: number }[];
}> {
	const tf = shape.getTextFrameOrNullObject();
	ctx.load(tf, "isNullObject,textRange/text");
	await ctx.sync();

	if (tf.isNullObject) {
		return { error: `${label} does not support text (type: image/table/etc)` };
	}

	const fullText = safeStr(tf.textRange.text);
	const paragraphSpans = splitParagraphs(fullText);

	// Batch: load the whole-range properties (the defaultProperties
	// baseline) plus one getSubstring() range per paragraph, all in a
	// single sync — mirrors handleGetTable's create-N-then-one-sync shape.
	ctx.load(tf.textRange, TEXT_RANGE_PROP_PATH);
	const paragraphRanges = paragraphSpans.map((span) => {
		const range = tf.textRange.getSubstring(span.start, span.length);
		ctx.load(range, TEXT_RANGE_PROP_PATH);
		return range;
	});
	await ctx.sync();

	const defaultProperties = extractRangeProperties(tf.textRange);

	// PowerPoint's whole-range font.size getter can report the
	// mixed/unresolved sentinel (0 after safeNum) even when every
	// character in the range shares the same real, explicit size (seen
	// live: an all-35pt title still reported the sentinel for the whole
	// range and for a getSubstring() spanning its full length). Sample a
	// single character — which reliably resolves — before falling back
	// to the shape's layout-inherited size.
	const needsDefaultSample = safeNum((defaultProperties.font as { fontSize?: number }).fontSize) === 0 && fullText.length > 0;
	const defaultSample = needsDefaultSample ? tf.textRange.getSubstring(0, 1) : null;
	if (defaultSample) ctx.load(defaultSample, "font/size");

	const paragraphSamples = paragraphRanges.map((range, index) => {
		const props = extractRangeProperties(range);
		const needsSample = safeNum((props.font as { fontSize?: number }).fontSize) === 0 && paragraphSpans[index].length > 0;
		if (!needsSample) return null;
		const sample = range.getSubstring(0, 1);
		ctx.load(sample, "font/size");
		return sample;
	});
	if (defaultSample || paragraphSamples.some(Boolean)) await ctx.sync();

	if (defaultSample) {
		const sampledSize = safeNum(defaultSample.font.size);
		defaultProperties.font.fontSize = sampledSize > 0 ? sampledSize : await resolveInheritedFontSize(ctx, shape);
	}

	// A paragraph reporting the same mixed/unresolved sentinel (0) as the
	// shape's own whole-range size has no run-level override either, so it
	// inherits the same resolved size — without this, diffRangeProperties
	// would surface a spurious fontSize:0 override once defaultProperties
	// has been corrected above.
	const paragraphProperties = paragraphRanges.map((range, index) => {
		const props = extractRangeProperties(range);
		if (safeNum((props.font as { fontSize?: number }).fontSize) === 0) {
			const sample = paragraphSamples[index];
			const sampledSize = sample ? safeNum(sample.font.size) : 0;
			props.font.fontSize = sampledSize > 0 ? sampledSize : defaultProperties.font.fontSize;
		}
		return props;
	});

	// Dedup by content: identical diffs (even for non-adjacent
	// paragraphs) share one groupId/propertyGroups entry.
	const groupIdByDiffKey = new Map<string, number>();
	const propertyGroups: { groupId: number; properties: Record<string, unknown> }[] = [];
	const paragraphs = paragraphSpans.map((span, index) => {
		const diff = diffRangeProperties(paragraphProperties[index], defaultProperties);
		const diffKey = JSON.stringify(diff);
		let groupId = groupIdByDiffKey.get(diffKey);
		if (groupId === undefined) {
			groupId = propertyGroups.length;
			groupIdByDiffKey.set(diffKey, groupId);
			propertyGroups.push({ groupId, properties: diff });
		}
		return { index, text: span.text, start: span.start, length: span.length, groupId };
	});

	return {
		fullText,
		defaultProperties,
		propertyGroups,
		paragraphs,
	};
}

async function handleGetShapeParagraphs(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string };
	const slideIndex = config.slideIndex ?? 1;
	const shapeId = config.shapeId ?? "";

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveShape(ctx, slideIndex, shapeId);
		if ("error" in resolved) return resolved;
		const { shape } = resolved;

		const data = await getShapeParagraphsData(shape, ctx, `Shape '${shapeId}'`);
		if ("error" in data) return data;

		return { slideIndex, shapeId, ...data };
	});
}

export type InlineRun = {
	start: number;
	length: number;
	bold: boolean;
	italic: boolean;
	strikethrough: boolean;
};

// Relative priority of each marker character in the fixed nesting order:
// bold outermost, italic next, strikethrough innermost. Reading an opening
// marker sequence left-to-right (e.g. "*_~") visits strictly increasing
// priorities; the matching close is that sequence reversed ("~_*").
const MARKER_PRIORITY: Record<string, number> = { "*": 0, "_": 1, "~": 2 };

function escapeMarkers(s: string): string {
	return s.replace(/[*_~]/g, (m) => "\\" + m);
}

function unescapeMarkers(s: string): string {
	return s.replace(/\\([*_~])/g, "$1");
}

/**
 * Inverse of `parseInlineMarkers`: wraps each run's slice of `text` in
 * Slack-style markers (`*bold*`, `_italic_`, `~strike~`), nested in the
 * fixed order bold > italic > strikethrough (e.g. bold+italic+strikethrough
 * -> "*_~text~_*"), and backslash-escapes literal `*`/`_`/`~` everywhere
 * else so they round-trip through `parseInlineMarkers` as plain characters.
 * `runs` must be sorted by `start` and non-overlapping.
 */
export function wrapInlineMarkers(text: string, runs: InlineRun[]): string {
	let out = "";
	let cursor = 0;
	for (const run of runs) {
		out += escapeMarkers(text.slice(cursor, run.start));
		let wrapped = escapeMarkers(text.slice(run.start, run.start + run.length));
		if (run.strikethrough) wrapped = `~${wrapped}~`;
		if (run.italic) wrapped = `_${wrapped}_`;
		if (run.bold) wrapped = `*${wrapped}*`;
		out += wrapped;
		cursor = run.start + run.length;
	}
	out += escapeMarkers(text.slice(cursor));
	return out;
}

// Finds the first occurrence of `marker` at or after `from` that isn't
// preceded by a backslash escape.
function findUnescapedMarker(text: string, from: number, marker: string): number {
	let idx = text.indexOf(marker, from);
	while (idx !== -1 && text[idx - 1] === "\\") {
		idx = text.indexOf(marker, idx + 1);
	}
	return idx;
}

/**
 * Inverse of `wrapInlineMarkers`: strips Slack-style inline markers
 * (`*bold*`, `_italic_`, `~strike~`, nestable in the fixed bold > italic >
 * strikethrough order) out of `text`, returning the plain unmarked text plus
 * the runs those markers described (offsets relative to the returned text).
 * Backslash-escaped marker characters (`\*`, `\_`, `\~`) are unescaped and
 * never treated as delimiters. A marker sequence that doesn't follow the
 * fixed nesting order, or has no matching close later in the string, is not
 * an error — it's left as plain literal characters, mirroring this file's
 * existing "don't hard-error on ambiguous input" posture.
 */
export function parseInlineMarkers(text: string): { text: string; runs: InlineRun[] } {
	const runs: InlineRun[] = [];
	let out = "";
	let i = 0;

	while (i < text.length) {
		const ch = text[i];

		if (ch === "\\" && i + 1 < text.length && "*_~".includes(text[i + 1])) {
			out += text[i + 1];
			i += 2;
			continue;
		}

		if (ch !== "*" && ch !== "_" && ch !== "~") {
			out += ch;
			i += 1;
			continue;
		}

		// Greedily capture the longest strictly-increasing-priority run of
		// marker chars starting here (e.g. "*_~"), then try progressively
		// shorter prefixes until one has a matching unescaped close later
		// in the string.
		let openEnd = i + 1;
		while (
			openEnd < text.length &&
			"*_~".includes(text[openEnd]) &&
			MARKER_PRIORITY[text[openEnd]] > MARKER_PRIORITY[text[openEnd - 1]]
		) {
			openEnd++;
		}

		let matched = false;
		for (let len = openEnd - i; len >= 1 && !matched; len--) {
			const open = text.slice(i, i + len);
			const close = [...open].reverse().join("");
			const closeIndex = findUnescapedMarker(text, i + len, close);
			if (closeIndex === -1) continue;

			const inner = unescapeMarkers(text.slice(i + len, closeIndex));
			runs.push({
				start: out.length,
				length: inner.length,
				bold: open.includes("*"),
				italic: open.includes("_"),
				strikethrough: open.includes("~"),
			});
			out += inner;
			i = closeIndex + close.length;
			matched = true;
		}

		if (!matched) {
			out += ch;
			i += 1;
		}
	}

	return { text: out, runs };
}

/**
 * Groups shapes into reading-order rows: sorted top-to-bottom, with shapes
 * whose `top` differs by no more than 5% of slideHeight treated as the same
 * row (tolerates slight vertical misalignment), and left-to-right within
 * each row. Returns arrays of original-array indices, one per row. Falls
 * back to a fixed 10pt tolerance when slideHeight is unavailable (<= 0).
 */
export function groupShapesIntoRows(shapes: { top: number; left: number }[], slideHeight: number): number[][] {
	const tolerance = slideHeight > 0 ? 0.05 * slideHeight : 10;
	const order = shapes.map((_, i) => i).sort((a, b) => shapes[a].top - shapes[b].top);

	const rows: number[][] = [];
	let currentRow: number[] = [];
	let rowAnchorTop = 0;
	for (const i of order) {
		if (currentRow.length === 0 || Math.abs(shapes[i].top - rowAnchorTop) <= tolerance) {
			if (currentRow.length === 0) rowAnchorTop = shapes[i].top;
			currentRow.push(i);
		} else {
			rows.push(currentRow);
			currentRow = [i];
			rowAnchorTop = shapes[i].top;
		}
	}
	if (currentRow.length > 0) rows.push(currentRow);

	for (const row of rows) row.sort((a, b) => shapes[a].left - shapes[b].left);
	return rows;
}

/**
 * A paragraph's effective font size >= 32pt renders as a markdown heading.
 * `fontSize <= 0` means PowerPoint reported a mixed/unresolved size across the
 * range (its sentinel for "not uniform" — e.g. an unfilled placeholder with no
 * run-level override, inheriting from the layout/master) rather than an
 * actually tiny font, so it never qualifies as a heading either. See
 * `isFootnoteShape` for the same guard applied to the footnote check.
 */
export function isHeadingSize(fontSize: number): boolean {
	return fontSize > 0 && fontSize >= 32;
}

/**
 * A shape qualifies as a footnote when its effective font size is <= 12pt
 * AND its bottom edge falls within the last 15% of the slide's height.
 * `fontSize <= 0` means PowerPoint reported a mixed/unknown size across the
 * range (its sentinel for "not uniform"), not an actually tiny font, so it
 * never qualifies.
 */
export function isFootnoteShape(shape: { top: number; height: number; fontSize: number }, slideHeight: number): boolean {
	return shape.fontSize > 0 && shape.fontSize <= 12 && slideHeight > 0 && shape.top + shape.height >= 0.85 * slideHeight;
}

/**
 * Renders paragraphs as an indented markdown-like list, one marker per
 * paragraph, indented by its effective indentLevel (a paragraph's own
 * override if present in propertyGroups, else the shape's defaultProperties
 * value). Numbered paragraphs (bulletType "Numbered") get "1.", "2." markers
 * that count contiguous numbered paragraphs at the same indent level —
 * interrupted by a shallower paragraph or a non-numbered paragraph at the
 * same level, both of which restart the count at 1. Non-numbered paragraphs
 * get bulletChar. Pure and standalone so a future markdown->paragraphs parser
 * (for a "set text from markdown" tool) can sit next to it and reuse the
 * same indent-unit/marker conventions.
 */
type BulletLevelState = { count: number; wasNumbered: boolean; style: string };

/**
 * Builds one paragraph's bullet/indent markdown line (e.g. "  1. text") and
 * advances `levelState` (mutated in place) for numbered-list continuation.
 * Factored out of `paragraphsToMarkdown` so `shapeParagraphsToSlideMarkdown`
 * can reuse the same bullet/indent/numbering logic per paragraph while
 * deciding on a different prefix (heading/footnote) for some paragraphs.
 */
function paragraphLineToMarkdown(
	p: { text: string; groupId: number },
	runs: InlineRun[],
	diffByGroupId: Map<number, Record<string, any>>,
	defaultProperties: RangeProperties,
	bulletChar: string,
	levelState: BulletLevelState[],
): string {
	const diff = diffByGroupId.get(p.groupId) ?? {};
	const defaultPf = defaultProperties.paragraphFormat as {
		indentLevel?: number;
		bulletType?: string;
		bulletStyle?: string;
	};
	const indentLevel = safeNum(diff.paragraphFormat?.indentLevel, defaultPf.indentLevel);
	const bulletType = safeStr(diff.paragraphFormat?.bulletType, defaultPf.bulletType);
	const bulletStyle = safeStr(diff.paragraphFormat?.bulletStyle, defaultPf.bulletStyle);
	const isNumbered = bulletType === "Numbered";

	levelState.length = Math.min(levelState.length, indentLevel + 1);
	const prev = levelState[indentLevel];
	const count = isNumbered && prev?.wasNumbered && prev.style === bulletStyle ? prev.count + 1 : 1;
	levelState[indentLevel] = { count, wasNumbered: isNumbered, style: bulletStyle };

	const info = BULLET_STYLE_INFO[bulletStyle];
	const marker = !isNumbered ? bulletChar : info ? formatBulletMarker(info.family, info.wrapper, count) : `${count}.`;
	const text = wrapInlineMarkers(p.text, runs);
	return "  ".repeat(indentLevel) + marker + " " + text;
}

export function paragraphsToMarkdown(
	paragraphs: { text: string; groupId: number }[],
	propertyGroups: { groupId: number; properties: Record<string, any> }[],
	defaultProperties: RangeProperties,
	options: { bulletChar?: string } = {},
	wordRuns: InlineRun[][] = [],
): string {
	const bulletChar = options.bulletChar ?? "-";
	const diffByGroupId = new Map(propertyGroups.map((g) => [g.groupId, g.properties]));

	// Running per-indent-level numbering state, indexed by level. Truncated
	// whenever a shallower paragraph is seen, so a deeper list always
	// restarts at 1 the next time that level is reused. A bulletStyle change
	// at the same level also restarts the count — same "restart on
	// interruption" semantics already applied to indentLevel changes.
	const levelState: BulletLevelState[] = [];

	return paragraphs
		.map((p, index) => paragraphLineToMarkdown(p, wordRuns[index] ?? [], diffByGroupId, defaultProperties, bulletChar, levelState))
		.join("\n");
}

/**
 * Combines a shape's paragraphs into the slide-level markdown block for that
 * shape, as used by `powerpoint_get_slide_text_markdown`: each paragraph's
 * *effective* font size (its own diff override, else the shape's default)
 * decides whether that paragraph renders as a heading (`# `, >=32pt, list
 * formatting ignored), a footnote (`[^1] ` prefix on the normal bullet line,
 * when `isFootnote` — computed once for the whole shape — is true), or a
 * plain bullet/indent line.
 */
export function shapeParagraphsToSlideMarkdown(
	paragraphs: { text: string; groupId: number }[],
	propertyGroups: { groupId: number; properties: Record<string, any> }[],
	defaultProperties: RangeProperties,
	isFootnote: boolean,
	options: { bulletChar?: string } = {},
	wordRuns: InlineRun[][] = [],
): string {
	const bulletChar = options.bulletChar ?? "-";
	const diffByGroupId = new Map(propertyGroups.map((g) => [g.groupId, g.properties]));
	const defaultFontSize = safeNum((defaultProperties.font as { fontSize?: number } | undefined)?.fontSize, 0);
	const levelState: BulletLevelState[] = [];

	return paragraphs
		.map((p, index) => {
			const diff = diffByGroupId.get(p.groupId) ?? {};
			const fontSize = safeNum(diff.font?.fontSize, defaultFontSize);
			const runs = wordRuns[index] ?? [];

			if (isHeadingSize(fontSize)) {
				levelState.length = 0;
				return "# " + wrapInlineMarkers(p.text, runs);
			}

			const line = paragraphLineToMarkdown(p, runs, diffByGroupId, defaultProperties, bulletChar, levelState);
			return isFootnote ? "[^1] " + line : line;
		})
		.join("\n");
}

/**
 * Reads per-word bold/italic/strikethrough for every paragraph in `tf`'s text
 * and merges adjacent words sharing identical flags into runs, for
 * `paragraphsToMarkdown`'s `wordRuns` param. Batches every word's
 * `getSubstring` load behind a single `ctx.sync()`, mirroring
 * `handleGetShapeParagraphs`'s paragraph-level batching.
 */
async function loadWordRuns(
	tf: PowerPoint.TextFrame,
	paragraphSpans: { text: string; start: number; length: number }[],
	ctx: PowerPoint.RequestContext,
): Promise<InlineRun[][]> {
	const wordSpansByParagraph = paragraphSpans.map((span) =>
		[...span.text.matchAll(/\S+/g)].map((m) => ({ start: m.index as number, length: m[0].length })),
	);

	const wordRangesByParagraph = wordSpansByParagraph.map((wordSpans, pIndex) =>
		wordSpans.map((word) => {
			const range = tf.textRange.getSubstring(paragraphSpans[pIndex].start + word.start, word.length);
			ctx.load(range, "font/bold,font/italic,font/strikethrough");
			return range;
		}),
	);
	await ctx.sync();

	return wordSpansByParagraph.map((wordSpans, pIndex) => {
		const runs: InlineRun[] = [];
		let lastWordIndex = -1;
		wordSpans.forEach((word, wIndex) => {
			const range = wordRangesByParagraph[pIndex][wIndex];
			const bold = !!range.font.bold;
			const italic = !!range.font.italic;
			const strikethrough = !!range.font.strikethrough;
			const last = runs[runs.length - 1];
			if (
				last &&
				lastWordIndex === wIndex - 1 &&
				last.bold === bold &&
				last.italic === italic &&
				last.strikethrough === strikethrough
			) {
				last.length = word.start + word.length - last.start;
				lastWordIndex = wIndex;
			} else if (bold || italic || strikethrough) {
				runs.push({ start: word.start, length: word.length, bold, italic, strikethrough });
				lastWordIndex = wIndex;
			}
		});
		return runs;
	});
}

async function handleGetShapeTextMarkdown(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string; bulletChar?: string };
	const slideIndex = config.slideIndex ?? 1;
	const shapeId = config.shapeId ?? "";
	const bulletChar = config.bulletChar ?? "-";

	if (bulletChar !== "-" && bulletChar !== "*") {
		return { error: `bulletChar must be '-' or '*', got '${bulletChar}'` };
	}

	const result = (await handleGetShapeParagraphs({ slideIndex, shapeId })) as any;
	if (result.error) return result;

	const wordRuns = await runInPowerPoint(async (ctx) => {
		const resolved = await resolveShape(ctx, slideIndex, shapeId);
		if ("error" in resolved) return [];
		const tf = resolved.shape.getTextFrameOrNullObject();
		ctx.load(tf, "isNullObject");
		await ctx.sync();
		if (tf.isNullObject) return [];
		return loadWordRuns(tf, splitParagraphs(result.fullText), ctx);
	});

	const markdown = paragraphsToMarkdown(
		result.paragraphs,
		result.propertyGroups,
		result.defaultProperties,
		{ bulletChar },
		wordRuns,
	);

	return { slideIndex, shapeId, markdown };
}

async function handleGetSlideTextMarkdown(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; bulletChar?: string };
	const slideIndex = config.slideIndex ?? 1;
	const bulletChar = config.bulletChar ?? "-";

	if (bulletChar !== "-" && bulletChar !== "*") {
		return { error: `bulletChar must be '-' or '*', got '${bulletChar}'` };
	}

	const prep = await runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		pres.pageSetup.load("slideHeight");
		await ctx.sync();

		if (slideIndex < 1 || slideIndex > pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range (1-${pres.slides.items.length})` };
		}

		const slideHeight = safeNum(pres.pageSetup.slideHeight);
		const slide = pres.slides.items[toZeroBasedSlideIndex(slideIndex)];
		slide.load("shapes/items/$none");
		await ctx.sync();

		for (const s of slide.shapes.items) {
			s.load("id,name,left,top,width,height");
		}
		await ctx.sync();

		const tfs: any[] = [];
		for (const s of slide.shapes.items) {
			const tf = s.getTextFrameOrNullObject();
			ctx.load(tf, "isNullObject,textRange/text");
			tfs.push(tf);
		}
		await ctx.sync();

		const shapes = slide.shapes.items
			.map((s: any, i: number) => ({
				shapeId: safeStr(s.id),
				left: safeNum(s.left),
				top: safeNum(s.top),
				height: safeNum(s.height),
				hasText: !tfs[i].isNullObject && !!safeStr(tfs[i].textRange?.text).trim(),
			}))
			.filter((s: { hasText: boolean }) => s.hasText);

		return { slideHeight, shapes };
	});
	if ("error" in prep) return prep;
	const { slideHeight, shapes } = prep;

	const rows = groupShapesIntoRows(shapes, slideHeight);
	const orderedShapes = rows.flat().map((i) => shapes[i]);

	const blocks: string[] = [];
	for (const shape of orderedShapes) {
		const result = (await handleGetShapeParagraphs({ slideIndex, shapeId: shape.shapeId })) as any;
		if (result.error) continue;

		const wordRuns = await runInPowerPoint(async (ctx) => {
			const resolved = await resolveShape(ctx, slideIndex, shape.shapeId);
			if ("error" in resolved) return [];
			const tf = resolved.shape.getTextFrameOrNullObject();
			ctx.load(tf, "isNullObject");
			await ctx.sync();
			if (tf.isNullObject) return [];
			return loadWordRuns(tf, splitParagraphs(result.fullText), ctx);
		});

		const defaultFontSize = safeNum((result.defaultProperties.font as { fontSize?: number } | undefined)?.fontSize, 0);
		const isFootnote = isFootnoteShape({ top: shape.top, height: shape.height, fontSize: defaultFontSize }, slideHeight);

		blocks.push(
			shapeParagraphsToSlideMarkdown(result.paragraphs, result.propertyGroups, result.defaultProperties, isFootnote, { bulletChar }, wordRuns),
		);
	}

	return { slideIndex, markdown: blocks.join("\n\n") };
}

type LevelSlot = { family: BulletFamily; wrapper: BulletWrapper; lastCore: string };

function findBulletStyleName(family: BulletFamily, wrapper: BulletWrapper): string | undefined {
	return Object.keys(BULLET_STYLE_INFO).find(
		(name) => BULLET_STYLE_INFO[name].family === family && BULLET_STYLE_INFO[name].wrapper === wrapper,
	);
}

// Classifies an unwrapped marker core (the token with its period/parens
// stripped) into a BulletFamily, given the previous numbered line's state at
// the same indent level (only consulted to resolve the i/I ambiguity — see
// markdownToParagraphSpecs). Returns undefined for anything unrecognized.
function classifyMarkerCore(core: string, slot: LevelSlot | undefined): BulletFamily | undefined {
	// Capped at 1-2 digits (1-99) — a realistic ceiling for a slide list,
	// same spirit as capping Roman numerals at 1-19 below.
	if (/^\d{1,2}$/.test(core)) return "arabic";

	if (core.length === 1) {
		if (!/^[a-zA-Z]$/.test(core)) return undefined;
		const lower = core.toLowerCase();
		const isLower = core === lower;
		if (lower === "l" || lower === "c" || lower === "d" || lower === "m") {
			return isLower ? "alphaLower" : "alphaUpper";
		}
		if (lower === "v" || lower === "x") {
			return isLower ? "romanLower" : "romanUpper";
		}
		if (lower === "i") {
			const continuesAlphabetRun =
				!!slot &&
				(slot.family === "alphaLower" || slot.family === "alphaUpper") &&
				slot.lastCore === (isLower ? "h" : "H");
			if (continuesAlphabetRun) return isLower ? "alphaLower" : "alphaUpper";
			return isLower ? "romanLower" : "romanUpper";
		}
		return isLower ? "alphaLower" : "alphaUpper";
	}

	// Multi-character: only Roman numerals (1-19, i/v/x only) have a valid
	// multi-letter marker in this scheme — no case mixing (no /i flag), so
	// e.g. "iV" matches neither and falls through to the error below.
	if (/^x?(ix|iv|v?i{0,3})$/.test(core)) return "romanLower";
	if (/^X?(IX|IV|V?I{0,3})$/.test(core)) return "romanUpper";
	return undefined;
}

type ParagraphSpec = {
	text: string;
	indentLevel: number;
	bulletType: "None" | "Numbered" | "Unnumbered";
	bulletStyle?: string;
	bulletVisible?: true;
	runs?: InlineRun[];
};

// Marker token capped at 7 chars — enough for the longest realistic marker,
// "(xviii)" (parenBoth-wrapped Roman 18), with no wasted unbounded \S+.
const MARKER_TOKEN_RE = /^(\S{1,7}) (.*)$/;
// Explicit optional prefix "(" / suffix ")"/"." groups instead of four
// separate wrapper alternatives — narrows directly to the four wrapper
// shapes plus "plain", and rejects any other combo (e.g. an unclosed "(l" or
// a nonsensical "(l.") as not marker-shaped.
const WRAP_MATCH_RE = /^(\()?(\w{1,5})(\)|\.)?$/;

function splitIndent(line: string): { indentLevel: number; rest: string } {
	const match = line.match(/^((?:  )*)(.*)$/) as [string, string, string];
	return { indentLevel: match[1].length / 2, rest: match[2] };
}

function wrapperOf(prefix: string | undefined, suffix: string | undefined): BulletWrapper | undefined {
	if (prefix) return suffix === ")" ? "parenBoth" : undefined;
	if (suffix === ")") return "parenRight";
	if (suffix === ".") return "period";
	if (!suffix) return "plain";
	return undefined;
}

// Shape-only check for the lookahead corroboration: does this token merely
// look like a marker (bullet char, or wrapper-shaped core), without
// classifying its core into a real BulletFamily. Used only to decide whether
// an ambiguous single-letter token on the *current* line should be trusted.
function looksMarkerShaped(token: string): boolean {
	if (token === "-" || token === "*") return true;
	const wrapMatch = token.match(WRAP_MATCH_RE);
	if (!wrapMatch) return false;
	return wrapperOf(wrapMatch[1], wrapMatch[3]) !== undefined;
}

// Looks at the next non-blank line (regardless of its indent level) and
// reports whether it also looks marker-shaped — corroborating evidence that
// the current line's ambiguous token is really a list marker and not prose.
function nextLineCorroborates(lines: string[], fromIndex: number): boolean {
	for (let j = fromIndex; j < lines.length; j++) {
		const { rest } = splitIndent(lines[j]);
		if (rest === "") continue;
		const tokenMatch = rest.match(MARKER_TOKEN_RE);
		return !!tokenMatch && looksMarkerShaped(tokenMatch[1]);
	}
	return false;
}

/**
 * Parses the markdown flavor produced by paragraphsToMarkdown back into
 * per-paragraph specs (text/indentLevel/bulletType/bulletStyle/bulletVisible).
 * Pure and standalone — the "set text from markdown" handler wraps this and
 * applies the specs to a real shape.
 *
 * Lines that don't look like a real list marker (or whose marker-shaped
 * token can't be corroborated as intentional — see below) fall back to
 * plain prose paragraphs (`bulletType: "None"`) rather than erroring, so
 * prose can be interleaved with list items. The only remaining hard error is
 * a marker-shaped token that appears *inside* an already-established
 * numbered run (same wrapper, same indent level) but doesn't fit the run's
 * pattern — there, silently falling back to plain text would likely not be
 * what the caller intended.
 */
export function markdownToParagraphSpecs(markdown: string): ParagraphSpec[] | { error: string } {
	const lines = markdown.split("\n");
	const specs: ParagraphSpec[] = [];

	// Sequence-aware state, mirroring paragraphsToMarkdown's levelState, but
	// only needs to remember enough to resolve i/I: per indent level, the
	// family/wrapper/lastCore of the immediately preceding numbered line at
	// that level, within the current contiguous numbered run.
	const levelSlots: (LevelSlot | undefined)[] = [];

	for (let i = 0; i < lines.length; i++) {
		const lineNo = i + 1;
		const { indentLevel, rest } = splitIndent(lines[i]);
		levelSlots.length = Math.min(levelSlots.length, indentLevel + 1);

		const asPlain = () => {
			levelSlots[indentLevel] = undefined;
			const { text: cleanText, runs } = parseInlineMarkers(rest);
			specs.push({ text: cleanText, indentLevel, bulletType: "None", runs });
		};

		const tokenMatch = rest.match(MARKER_TOKEN_RE);
		if (!tokenMatch) {
			asPlain();
			continue;
		}
		const [, token, text] = tokenMatch;

		if (token === "-" || token === "*") {
			levelSlots[indentLevel] = undefined;
			const { text: cleanText, runs } = parseInlineMarkers(text);
			specs.push({ text: cleanText, indentLevel, bulletType: "Unnumbered", bulletVisible: true, runs });
			continue;
		}

		const wrapMatch = token.match(WRAP_MATCH_RE);
		const wrapper = wrapMatch ? wrapperOf(wrapMatch[1], wrapMatch[3]) : undefined;
		const core = wrapMatch?.[2];
		const runSlot = levelSlots[indentLevel];
		const insideMatchingRun = !!runSlot && runSlot.wrapper === wrapper;

		if (!wrapper || core === undefined) {
			// Not even marker-shaped — plain text, unless we're inside a
			// run expecting this wrapper (can't happen here since wrapper
			// is undefined, so this is always a safe plain fallback).
			asPlain();
			continue;
		}

		const family = classifyMarkerCore(core, runSlot);
		const bulletStyle = family && findBulletStyleName(family, wrapper);
		if (!family || !bulletStyle) {
			if (insideMatchingRun) {
				return {
					error: `Marker '${token}' at line ${lineNo} continues a numbered list but does not match any supported bullet style.`,
				};
			}
			asPlain();
			continue;
		}

		const isExemptArabic = family === "arabic" && core === "1";
		const corroborated = isExemptArabic || insideMatchingRun || nextLineCorroborates(lines, i + 1);
		if (!corroborated) {
			asPlain();
			continue;
		}

		levelSlots[indentLevel] = { family, wrapper, lastCore: core };
		const { text: cleanText, runs } = parseInlineMarkers(text);
		specs.push({ text: cleanText, indentLevel, bulletType: "Numbered", bulletStyle, bulletVisible: true, runs });
	}

	return specs;
}

async function handleGetSelection(_args: unknown): Promise<unknown> {
	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;

		// Try to get selected text range
		try {
			const textRange = pres.getSelectedTextRange();
			textRange.load(
				"text,font/name,font/size,font/bold,font/italic,font/color",
			);

			const parentTf = textRange.getParentTextFrame();
			const parentShape = parentTf.getParentShape();
			const parentSlide = parentShape.getParentSlide();
			parentShape.load("id,name");
			parentSlide.load("id");
			await ctx.sync();

			return {
				type: "text",
				text: safeStr(textRange.text),
				font: {
					name: safeStr(textRange.font.name),
					size: safeNum(textRange.font.size),
					bold: !!textRange.font.bold,
					italic: !!textRange.font.italic,
					color: safeStr(textRange.font.color),
				},
				shapeId: safeStr(parentShape.id),
				shapeName: safeStr(parentShape.name),
				slideId: safeStr(parentSlide.id),
			};
		} catch {
			// No text selected — try shapes
		}

		// Try to get selected shapes
		try {
			const selectedShapes = pres.getSelectedShapes();
			selectedShapes.load("items");
			await ctx.sync();

			if (selectedShapes.items.length > 0) {
				for (const s of selectedShapes.items) {
					s.load("id,name,type");
				}
				await ctx.sync();

				return {
					type: "shapes",
					shapes: selectedShapes.items.map((s: any) => ({
						id: safeStr(s.id),
						name: safeStr(s.name),
						type: safeStr(s.type),
					})),
				};
			}
		} catch {
			// No shapes selected
		}

		return { type: "none" };
	});
}

async function handleGetSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; slideRange?: string };
	const slideIndex = config.slideIndex;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		// Determine which slides to read
		let indices: number[] = [];
		if (slideIndex != null) {
			indices = [toZeroBasedSlideIndex(slideIndex)];
		} else if (config.slideRange) {
			const match = config.slideRange.match(/^(\d+)-(\d+)$/);
			if (match) {
				const from = toZeroBasedSlideIndex(parseInt(match[1]));
				const to = toZeroBasedSlideIndex(parseInt(match[2]));
				for (let i = from; i <= to && i < pres.slides.items.length; i++) {
					indices.push(i);
				}
			} else {
				return {
					error: `Invalid slideRange format: '${config.slideRange}'. Use '2-5'.`,
				};
			}
		} else {
			// Default: all slides
			for (let i = 0; i < pres.slides.items.length; i++) {
				indices.push(i);
			}
		}

		// Read notes for each slide
		const notes: Array<{ slideIndex: number; notes: string }> = [];
		const notesTfs: any[] = [];

		for (const idx of indices) {
			if (idx < 0 || idx >= pres.slides.items.length) {
				notes.push({ slideIndex: toOneBasedSlideIndex(idx), notes: "" });
				notesTfs.push(null);
				continue;
			}

			const slide = pres.slides.items[idx];
			const notesSlide = slide.getNotesSlideOrNullObject();
			const tf = notesSlide.textFrame;
			ctx.load(tf, "isNullObject,textRange/text");
			notesTfs.push(tf);
		}
		await ctx.sync();

		for (let i = 0; i < indices.length; i++) {
			const tf = notesTfs[i];
			let noteText = "";
			if (tf && !tf.isNullObject && tf.textRange?.text) {
				noteText = String(tf.textRange.text).trim();
			}
			notes.push({ slideIndex: toOneBasedSlideIndex(indices[i]), notes: noteText });
		}

		return { notes };
	});
}

// ── Write tools ─────────────────────────────────────────────────

async function handleUpdateShapeText(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		text?: string;
	};
	const { slideIndex = 1, shapeId = "", text = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveShape(ctx, slideIndex, shapeId);
		if ("error" in resolved) return resolved;
		const { shape } = resolved;

		// Get text frame — check if it's a text-bearing shape
		const tf = shape.getTextFrameOrNullObject();
		ctx.load(tf, "isNullObject");
		await ctx.sync();

		if (tf.isNullObject) {
			return {
				error: `Shape '${shapeId}' does not support text (type: image/table/etc)`,
			};
		}

		// Set the text
		tf.textRange.text = text;
		await ctx.sync();

		return { slideIndex, shapeId, newText: text };
	});
}

async function handleUpdateShapeProperties(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
		rotation?: number;
		fontName?: string;
		fontSize?: number;
		bold?: boolean;
		italic?: boolean;
		color?: string;
		underline?: string;
		strikethrough?: boolean;
		doubleStrikethrough?: boolean;
		allCaps?: boolean;
		smallCaps?: boolean;
		subscript?: boolean;
		superscript?: boolean;
		horizontalAlignment?: string;
		indentLevel?: number;
		bulletType?: string;
		bulletStyle?: string;
		bulletVisible?: boolean;
		textMarginTop?: number;
		textMarginBottom?: number;
		textMarginLeft?: number;
		textMarginRight?: number;
		autoSizeSetting?: string;
		wordWrap?: boolean;
		verticalAlignment?: string;
	};
	const { slideIndex = 1, shapeId = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveShape(ctx, slideIndex, shapeId);
		if ("error" in resolved) return resolved;
		const { shape } = resolved;

		const updated: string[] = [];

		// Position and size
		if (config.left !== undefined) {
			shape.left = config.left;
			updated.push("left");
		}
		if (config.top !== undefined) {
			shape.top = config.top;
			updated.push("top");
		}
		if (config.width !== undefined) {
			shape.width = config.width;
			updated.push("width");
		}
		if (config.height !== undefined) {
			shape.height = config.height;
			updated.push("height");
		}
		if (config.rotation !== undefined) {
			shape.rotation = config.rotation;
			updated.push("rotation");
		}

		// Font, paragraph/bullet, and text frame properties
		if (
			hasAnyDefinedKey(config, FONT_SETTERS, PARAGRAPH_FORMAT_SETTERS, BULLET_FORMAT_SETTERS) ||
			config.textMarginTop !== undefined ||
			config.textMarginBottom !== undefined ||
			config.textMarginLeft !== undefined ||
			config.textMarginRight !== undefined ||
			config.autoSizeSetting !== undefined ||
			config.wordWrap !== undefined ||
			config.verticalAlignment !== undefined
		) {
			const tf = shape.getTextFrameOrNullObject();
			ctx.load(tf, "isNullObject");
			await ctx.sync();

			if (!tf.isNullObject) {
				const font = tf.textRange.font;
				if (config.fontName !== undefined) {
					font.name = config.fontName;
					updated.push("fontName");
				}
				if (config.fontSize !== undefined) {
					font.size = config.fontSize;
					updated.push("fontSize");
				}
				if (config.bold !== undefined) {
					font.bold = config.bold;
					updated.push("bold");
				}
				if (config.italic !== undefined) {
					font.italic = config.italic;
					updated.push("italic");
				}
				if (config.color !== undefined) {
					font.color = config.color;
					updated.push("color");
				}
				if (config.underline !== undefined) {
					font.underline = config.underline as any;
					updated.push("underline");
				}
				if (config.strikethrough !== undefined) {
					font.strikethrough = config.strikethrough;
					updated.push("strikethrough");
				}
				if (config.doubleStrikethrough !== undefined) {
					font.doubleStrikethrough = config.doubleStrikethrough;
					updated.push("doubleStrikethrough");
				}
				if (config.allCaps !== undefined) {
					font.allCaps = config.allCaps;
					updated.push("allCaps");
				}
				if (config.smallCaps !== undefined) {
					font.smallCaps = config.smallCaps;
					updated.push("smallCaps");
				}
				if (config.subscript !== undefined) {
					font.subscript = config.subscript;
					updated.push("subscript");
				}
				if (config.superscript !== undefined) {
					font.superscript = config.superscript;
					updated.push("superscript");
				}

				const pf = tf.textRange.paragraphFormat;
				if (config.horizontalAlignment !== undefined) {
					pf.horizontalAlignment = config.horizontalAlignment as any;
					updated.push("horizontalAlignment");
				}
				if (config.indentLevel !== undefined) {
					pf.indentLevel = config.indentLevel;
					updated.push("indentLevel");
				}

				const bf = pf.bulletFormat;
				if (config.bulletType !== undefined) {
					bf.type = config.bulletType as any;
					updated.push("bulletType");
				}
				if (config.bulletStyle !== undefined) {
					bf.style = config.bulletStyle as any;
					updated.push("bulletStyle");
				}
				if (config.bulletVisible !== undefined) {
					bf.visible = config.bulletVisible;
					updated.push("bulletVisible");
				}

				if (config.textMarginTop !== undefined) {
					tf.topMargin = config.textMarginTop;
					updated.push("textMarginTop");
				}
				if (config.textMarginBottom !== undefined) {
					tf.bottomMargin = config.textMarginBottom;
					updated.push("textMarginBottom");
				}
				if (config.textMarginLeft !== undefined) {
					tf.leftMargin = config.textMarginLeft;
					updated.push("textMarginLeft");
				}
				if (config.textMarginRight !== undefined) {
					tf.rightMargin = config.textMarginRight;
					updated.push("textMarginRight");
				}
				if (config.autoSizeSetting !== undefined) {
					tf.autoSizeSetting = config.autoSizeSetting as any;
					updated.push("autoSizeSetting");
				}
				if (config.wordWrap !== undefined) {
					tf.wordWrap = config.wordWrap;
					updated.push("wordWrap");
				}
				if (config.verticalAlignment !== undefined) {
					tf.verticalAlignment = config.verticalAlignment as any;
					updated.push("verticalAlignment");
				}
			}
		}

		await ctx.sync();

		return { slideIndex, shapeId, updated };
	});
}

async function handleUpdateTextRangeProperties(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		start?: number;
		length?: number;
		expectedText?: string;
		[key: string]: unknown;
	};
	const { slideIndex = 1, shapeId = "", start = 0, length = 0 } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveShape(ctx, slideIndex, shapeId);
		if ("error" in resolved) return resolved;
		const { shape } = resolved;

		const tf = shape.getTextFrameOrNullObject();
		ctx.load(tf, "isNullObject,textRange/text");
		await ctx.sync();

		if (tf.isNullObject) {
			return {
				error: `Shape '${shapeId}' does not support text (type: image/table/etc)`,
			};
		}

		const fullText = safeStr(tf.textRange.text);
		if (start < 0 || start + length > fullText.length) {
			return { error: `start/length out of range for shape text (length ${fullText.length})` };
		}

		const range = tf.textRange.getSubstring(start, length);
		ctx.load(range, `text,${TEXT_RANGE_PROP_PATH}`);
		await ctx.sync();

		if (config.expectedText !== undefined && safeStr(range.text) !== config.expectedText) {
			return {
				error:
					"Text at [start,length) does not match expectedText — shape text may have changed since it was last read.",
			};
		}

		const updated = await applyRangeProperties(config, range, ctx);

		await ctx.sync();

		return { slideIndex, shapeId, start, length, updated };
	});
}

async function handleInsertParagraph(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		text?: string;
		position?: string;
		refParagraphStart?: number;
		refParagraphLength?: number;
		refParagraphText?: string;
		[key: string]: unknown;
	};
	const {
		slideIndex = 1,
		shapeId = "",
		text = "",
		position = "end",
		refParagraphStart = 0,
		refParagraphLength = 0,
	} = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveShape(ctx, slideIndex, shapeId);
		if ("error" in resolved) return resolved;
		const { shape } = resolved;

		const tf = shape.getTextFrameOrNullObject();
		ctx.load(tf, "isNullObject,textRange/text");
		await ctx.sync();

		if (tf.isNullObject) {
			return {
				error: `Shape '${shapeId}' does not support text (type: image/table/etc)`,
			};
		}

		const fullText = safeStr(tf.textRange.text);

		if (position === "before" || position === "after") {
			if (refParagraphStart < 0 || refParagraphStart + refParagraphLength > fullText.length) {
				return {
					error: `refParagraphStart/refParagraphLength out of range for shape text (length ${fullText.length})`,
				};
			}
		} else if (position !== "start" && position !== "end") {
			return { error: `Invalid position '${position}' — expected start|end|before|after` };
		}

		if (config.refParagraphText !== undefined && (position === "before" || position === "after")) {
			const refRange = tf.textRange.getSubstring(refParagraphStart, refParagraphLength);
			ctx.load(refRange, "text");
			await ctx.sync();

			if (safeStr(refRange.text) !== config.refParagraphText) {
				return {
					error:
						"Text at [refParagraphStart,refParagraphLength) does not match refParagraphText — shape text may have changed since it was last read.",
				};
			}
		}

		// The separator's placement relative to `text` — not the numeric
		// value of `insertAt` — is what determines the splice: "before"/"start"
		// always land right before an existing paragraph (so the new text is
		// followed by "\r"), "after"/"end" always land right after one (so the
		// new text is preceded by "\r"). An empty shape has neither.
		let insertAt: number;
		let insertText: string;
		if (position === "start") {
			insertAt = 0;
			insertText = fullText.length === 0 ? text : text + "\r";
		} else if (position === "end") {
			insertAt = fullText.length;
			insertText = fullText.length === 0 ? text : "\r" + text;
		} else if (position === "before") {
			insertAt = refParagraphStart;
			insertText = text + "\r";
		} else {
			insertAt = refParagraphStart + refParagraphLength;
			insertText = "\r" + text;
		}

		const insertionPoint = tf.textRange.getSubstring(insertAt, 0);
		insertionPoint.text = insertText;
		await ctx.sync();

		const newStart = insertText.startsWith("\r") ? insertAt + 1 : insertAt;
		const newLength = text.length;

		let updated: string[] = [];
		if (hasAnyDefinedKey(config, FONT_SETTERS, PARAGRAPH_FORMAT_SETTERS, BULLET_FORMAT_SETTERS)) {
			const range = tf.textRange.getSubstring(newStart, newLength);
			ctx.load(range, TEXT_RANGE_PROP_PATH);
			await ctx.sync();

			updated = await applyRangeProperties(config, range, ctx);

			await ctx.sync();
		}

		return { slideIndex, shapeId, start: newStart, length: newLength, updated };
	});
}

async function handleDeleteParagraph(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		paragraphStart?: number;
		paragraphLength?: number;
		expectedText?: string;
	};
	const { slideIndex = 1, shapeId = "", paragraphStart = 0, paragraphLength = 0 } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveShape(ctx, slideIndex, shapeId);
		if ("error" in resolved) return resolved;
		const { shape } = resolved;

		const tf = shape.getTextFrameOrNullObject();
		ctx.load(tf, "isNullObject,textRange/text");
		await ctx.sync();

		if (tf.isNullObject) {
			return {
				error: `Shape '${shapeId}' does not support text (type: image/table/etc)`,
			};
		}

		const fullText = safeStr(tf.textRange.text);
		if (paragraphStart < 0 || paragraphStart + paragraphLength > fullText.length) {
			return {
				error: `paragraphStart/paragraphLength out of range for shape text (length ${fullText.length})`,
			};
		}

		const range = tf.textRange.getSubstring(paragraphStart, paragraphLength);
		ctx.load(range, "text");
		await ctx.sync();

		if (config.expectedText !== undefined && safeStr(range.text) !== config.expectedText) {
			return {
				error:
					"Text at [paragraphStart,paragraphLength) does not match expectedText — shape text may have changed since it was last read.",
			};
		}

		// Deleting only [paragraphStart, paragraphStart+paragraphLength) would
		// leave the paragraph's bounding "\r" behind, splitting an empty
		// paragraph in two. Consume one adjacent "\r" — the following one if
		// this isn't the last paragraph, otherwise the preceding one — so the
		// paragraph count actually drops by one. A lone remaining paragraph
		// has no adjacent "\r" at all; its text is simply emptied, matching
		// how PowerPoint text frames always keep at least one paragraph.
		let delStart = paragraphStart;
		let delEnd = paragraphStart + paragraphLength;
		if (delEnd < fullText.length && fullText[delEnd] === "\r") {
			delEnd += 1;
		} else if (delStart > 0 && fullText[delStart - 1] === "\r") {
			delStart -= 1;
		}

		const deleteRange = tf.textRange.getSubstring(delStart, delEnd - delStart);
		deleteRange.text = "";
		await ctx.sync();

		return { slideIndex, shapeId, paragraphStart, paragraphLength, deleted: true };
	});
}

async function handleSetShapeTextMarkdown(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		shapeId?: string;
		markdown?: string;
	};
	const { slideIndex = 1, shapeId = "", markdown = "" } = config;

	const specs = markdownToParagraphSpecs(markdown);
	if ("error" in specs) return specs;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveShape(ctx, slideIndex, shapeId);
		if ("error" in resolved) return resolved;
		const { shape } = resolved;

		const tf = shape.getTextFrameOrNullObject();
		ctx.load(tf, "isNullObject");
		await ctx.sync();

		if (tf.isNullObject) {
			return {
				error: `Shape '${shapeId}' does not support text (type: image/table/etc)`,
			};
		}

		const fullText = specs.map((s) => s.text).join("\r");
		tf.textRange.text = fullText;
		await ctx.sync();

		const spans = splitParagraphs(fullText);

		for (let i = 0; i < specs.length; i++) {
			const spec = specs[i];
			const span = spans[i];
			const range = tf.textRange.getSubstring(span.start, span.length);
			ctx.load(range, TEXT_RANGE_PROP_PATH);
			await ctx.sync();

			await applyRangeProperties(
				{
					indentLevel: spec.indentLevel,
					bulletType: spec.bulletType,
					bulletStyle: spec.bulletStyle,
				},
				range,
				ctx,
			);
			await ctx.sync();

			for (const run of spec.runs ?? []) {
				const wordRange = tf.textRange.getSubstring(span.start + run.start, run.length);
				ctx.load(wordRange, TEXT_RANGE_PROP_PATH);
				await ctx.sync();

				await applyRangeProperties(
					{
						bold: run.bold || undefined,
						italic: run.italic || undefined,
						strikethrough: run.strikethrough || undefined,
					},
					wordRange,
					ctx,
				);
				await ctx.sync();
			}
		}

		return { slideIndex, shapeId, paragraphCount: specs.length };
	});
}

async function handleUpdateSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; notes?: string };
	const { slideIndex = 1, notes = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;
		const notesSlide = (slide as any).getNotesSlideOrNullObject();
		const tf = notesSlide.textFrame;
		ctx.load(tf, "isNullObject,textRange/text");
		await ctx.sync();

		if (tf.isNullObject) {
			return { error: `Slide ${slideIndex} does not have a notes pane` };
		}

		tf.textRange.text = notes;
		await ctx.sync();

		return { slideIndex, newNotes: notes };
	});
}

// ── Shape CRUD ──────────────────────────────────────────────────

async function handleAddTextbox(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		text?: string;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
	};
	const {
		slideIndex = 1,
		text = "",
		left = 100,
		top = 100,
		width = 300,
		height = 100,
	} = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;
		const textbox = slide.shapes.addTextBox(text, { left, top, width, height });
		textbox.load("id,name");
		await ctx.sync();

		return {
			slideIndex,
			shapeId: safeStr(textbox.id),
			name: safeStr(textbox.name),
		};
	});
}

async function handleAddImage(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		imageBase64?: string;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
	};
	const { slideIndex = 1, imageBase64 = "", left = 100, top = 100 } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;

		// Strip data URI prefix if present
		const base64Data = imageBase64.replace(/^data:image\/[^;]+;base64,/, "");

		const options: any = { left, top };
		if (config.width !== undefined) options.width = config.width;
		if (config.height !== undefined) options.height = config.height;

		// Office.js has no shape-creation primitive for pictures (no addPicture/addImage
		// on PowerPoint.ShapeCollection) — the only image API is ShapeFill.setImage(),
		// which paints a picture fill onto an existing shape. Use a borderless rectangle.
		const picture: any = slide.shapes.addGeometricShape("Rectangle", options);
		picture.fill.setImage(base64Data);
		picture.lineFormat.visible = false;
		picture.load("id,name");
		await ctx.sync();

		return {
			slideIndex,
			shapeId: safeStr(picture.id),
			name: safeStr(picture.name),
		};
	});
}

async function handleAddTable(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex?: number;
		rows?: number;
		columns?: number;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
	};
	const {
		slideIndex = 1,
		rows = 2,
		columns = 2,
		left = 100,
		top = 100,
	} = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;

		const options: any = { left, top };
		if (config.width !== undefined) options.width = config.width;
		if (config.height !== undefined) options.height = config.height;

		const table = slide.shapes.addTable(rows, columns, options);
		table.load("id,name");
		await ctx.sync();

		return {
			slideIndex,
			shapeId: safeStr(table.id),
			name: safeStr(table.name),
		};
	});
}

async function handleDeleteShape(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string };
	const { slideIndex = 1, shapeId = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;
		slide.load("shapes/items/$none");
		await ctx.sync();

		for (const s of slide.shapes.items) {
			s.load("id,name");
		}
		await ctx.sync();

		const shape = slide.shapes.items.find(
			(s: any) => safeStr(s.id) === shapeId || safeStr(s.name) === shapeId,
		);
		if (!shape) {
			return { error: `Shape '${shapeId}' not found on slide ${slideIndex}` };
		}

		shape.delete();
		await ctx.sync();

		return { slideIndex, shapeId, deleted: true };
	});
}

// ── Slide management ────────────────────────────────────────────

async function handleAddSlide(args: unknown): Promise<unknown> {
	const config = args as { atIndex?: number; layoutId?: string; slideMasterId?: string };

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const options: any = {};
		if (config.atIndex !== undefined) {
			options.index = toZeroBasedSlideIndex(config.atIndex);
		}
		if (config.layoutId !== undefined) {
			options.layoutId = config.layoutId;
		}
		if (config.slideMasterId !== undefined) {
			options.slideMasterId = config.slideMasterId;
		}

		pres.slides.add(options);

		// Reload to get the new slide
		pres.load("slides");
		await ctx.sync();

		// The new slide is at the specified index (or end)
		const newIndex =
			config.atIndex !== undefined
				? Math.min(toZeroBasedSlideIndex(config.atIndex), pres.slides.items.length - 1)
				: pres.slides.items.length - 1;

		const newSlide = pres.slides.items[newIndex];
		newSlide.load("id");
		await ctx.sync();

		return {
			slideIndex: toOneBasedSlideIndex(newIndex),
			slideId: safeStr(newSlide.id),
		};
	});
}

async function handleDeleteSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const { slideIndex = 1 } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;
		slide.delete();
		await ctx.sync();

		return { slideIndex, deleted: true };
	});
}

async function handleMoveSlide(args: unknown): Promise<unknown> {
	const config = args as { fromIndex?: number; toIndex?: number };
	const { fromIndex = 1, toIndex = 1 } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, fromIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;
		slide.load("id");
		await ctx.sync();

		const slideId = safeStr(slide.id);
		slide.moveTo(toZeroBasedSlideIndex(toIndex));
		await ctx.sync();

		return { fromIndex, toIndex, slideId };
	});
}

async function handleDuplicateSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; targetIndex?: number };
	const { slideIndex = 1, targetIndex } = config;

	return runInPowerPoint(async (ctx) => {
		const PowerPoint: any = (window as any).PowerPoint;
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const slideCount = pres.slides.items.length;
		if (slideIndex < 1 || slideIndex > slideCount) {
			return { error: `Slide index ${slideIndex} out of range (deck has ${slideCount} slides)` };
		}

		// Validate before mutating: once the duplicate is inserted the deck has slideCount + 1
		// slides, so valid target positions are 1..slideCount+1 inclusive (slideCount+1 itself
		// means "append after the new last slide"). Checking this up front avoids leaving a stray
		// duplicate behind when the request is rejected.
		if (targetIndex !== undefined && (targetIndex < 1 || targetIndex > slideCount + 1)) {
			return { error: `targetIndex ${targetIndex} out of range (deck has ${slideCount} slides; valid range is 1-${slideCount + 1})` };
		}

		const sourceSlide = pres.slides.items[toZeroBasedSlideIndex(slideIndex)];
		sourceSlide.load("id");
		await ctx.sync();

		const exportResult = sourceSlide.exportAsBase64();
		await ctx.sync();

		const options: any = {
			formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
			targetSlideId: sourceSlide.id,
		};

		pres.insertSlidesFromBase64(exportResult.value, options);
		await ctx.sync();

		pres.load("slides");
		await ctx.sync();

		const newSlideIndex = toZeroBasedSlideIndex(slideIndex) + 1;
		const newSlide = pres.slides.items[newSlideIndex];
		newSlide.load("id");
		await ctx.sync();
		const newSlideId = safeStr(newSlide.id);

		let finalIndex = newSlideIndex;
		if (targetIndex !== undefined && targetIndex !== toOneBasedSlideIndex(newSlideIndex)) {
			finalIndex = Math.max(
				0,
				Math.min(toZeroBasedSlideIndex(targetIndex), pres.slides.items.length - 1),
			);
			newSlide.moveTo(finalIndex);
			await ctx.sync();
		}

		return {
			sourceIndex: slideIndex,
			newSlideIndex: toOneBasedSlideIndex(finalIndex),
			newSlideId,
		};
	});
}

// Internal-only: exports a single slide as base64 for cross-document duplication.
// Not part of the public tool list — invoked by the server when
// powerpoint_duplicate_slide targets a different instance.
async function handleExportSlideInternal(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const { slideIndex = 1 } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide: sourceSlide } = resolved;
		const exportResult = sourceSlide.exportAsBase64();
		await ctx.sync();

		return { base64: exportResult.value };
	});
}

// Internal-only: inserts a slide exported by handleExportSlideInternal into this
// presentation, at targetIndex if given, otherwise appended at the end.
// Not part of the public tool list — invoked by the server when
// powerpoint_duplicate_slide targets a different instance.
async function handleImportSlideInternal(args: unknown): Promise<unknown> {
	const config = args as { base64?: string; targetIndex?: number };
	const { base64, targetIndex } = config;

	if (!base64) {
		return { error: "Missing base64 slide data" };
	}

	return runInPowerPoint(async (ctx) => {
		const PowerPoint: any = (window as any).PowerPoint;
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const beforeCount = pres.slides.items.length;

		// Validate before mutating: once the slide is inserted the deck has beforeCount + 1
		// slides, so valid target positions are 1..beforeCount+1 inclusive (beforeCount+1 itself
		// means "append after the new last slide"). Checking this up front avoids leaving a
		// stray imported slide behind when the request is rejected.
		if (targetIndex !== undefined && (targetIndex < 1 || targetIndex > beforeCount + 1)) {
			return { error: `targetIndex ${targetIndex} out of range (target deck has ${beforeCount} slides; valid range is 1-${beforeCount + 1})` };
		}

		const options: any = {
			formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
		};
		if (beforeCount > 0) {
			const lastSlide = pres.slides.items[beforeCount - 1];
			lastSlide.load("id");
			await ctx.sync();
			options.targetSlideId = lastSlide.id;
		}

		pres.insertSlidesFromBase64(base64, options);
		await ctx.sync();

		pres.load("slides");
		await ctx.sync();

		const newSlideIndex = beforeCount;
		const newSlide = pres.slides.items[newSlideIndex];
		newSlide.load("id");
		await ctx.sync();
		const newSlideId = safeStr(newSlide.id);

		let finalIndex = newSlideIndex;
		if (targetIndex !== undefined && targetIndex !== toOneBasedSlideIndex(newSlideIndex)) {
			finalIndex = Math.max(
				0,
				Math.min(toZeroBasedSlideIndex(targetIndex), pres.slides.items.length - 1),
			);
			newSlide.moveTo(finalIndex);
			await ctx.sync();
		}

		return { newSlideIndex: toOneBasedSlideIndex(finalIndex), newSlideId };
	});
}

// ── Phase 18: Tags & Metadata ────────────────────────────────────

async function handleGetTags(args: unknown): Promise<unknown> {
	const config = args as {
		target?: string;
		slideIndex?: number;
		shapeId?: string;
	};
	const { target = "presentation", slideIndex, shapeId } = config;

	return runInPowerPoint(async (ctx) => {
		let tagTarget: any;

		if (target === "slide" && slideIndex !== undefined) {
			tagTarget = ctx.presentation.slides.getItemAt(toZeroBasedSlideIndex(slideIndex));
		} else if (target === "shape" && slideIndex !== undefined && shapeId) {
			const slide = ctx.presentation.slides.getItemAt(toZeroBasedSlideIndex(slideIndex));
			tagTarget = slide.shapes.getItem(shapeId);
		} else {
			tagTarget = ctx.presentation;
		}

		tagTarget.tags.load("items");
		await ctx.sync();

		const tags: Record<string, string> = {};
		for (const tag of tagTarget.tags.items) {
			tag.load(["key", "value"]);
		}
		await ctx.sync();
		for (const tag of tagTarget.tags.items) {
			tags[tag.key] = tag.value;
		}

		return { target, tags, count: Object.keys(tags).length };
	});
}

async function handleSetTag(args: unknown): Promise<unknown> {
	const config = args as {
		key: string;
		value: string;
		target?: string;
		slideIndex?: number;
		shapeId?: string;
	};
	const { key, value, target = "presentation", slideIndex, shapeId } = config;

	return runInPowerPoint(async (ctx) => {
		let tagTarget: any;

		if (target === "slide" && slideIndex !== undefined) {
			tagTarget = ctx.presentation.slides.getItemAt(toZeroBasedSlideIndex(slideIndex));
		} else if (target === "shape" && slideIndex !== undefined && shapeId) {
			const slide = ctx.presentation.slides.getItemAt(toZeroBasedSlideIndex(slideIndex));
			tagTarget = slide.shapes.getItem(shapeId);
		} else {
			tagTarget = ctx.presentation;
		}

		tagTarget.tags.add(key, value);
		await ctx.sync();

		return { key, value, target, set: true };
	});
}

async function handleDeleteSlidesByTag(args: unknown): Promise<unknown> {
	const config = args as { key: string; value?: string };
	const { key, value } = config;

	return runInPowerPoint(async (ctx) => {
		const slides = ctx.presentation.slides;
		slides.load("items");
		await ctx.sync();

		const toDelete: any[] = [];
		for (const slide of slides.items) {
			slide.tags.load("items");
		}
		await ctx.sync();

		for (const slide of slides.items) {
			for (const tag of slide.tags.items) {
				tag.load(["key", "value"]);
			}
		}
		await ctx.sync();

		for (const slide of slides.items) {
			for (const tag of slide.tags.items) {
				if (tag.key === key && (value === undefined || tag.value === value)) {
					toDelete.push(slide);
					break;
				}
			}
		}

		for (const slide of toDelete) {
			slide.delete();
		}
		await ctx.sync();

		return { deletedCount: toDelete.length, key, value: value || "any" };
	});
}

// ── Phase 18: Shape Fill/Line/Rotation ───────────────────────────

async function handleSetShapeFill(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex: number;
		shapeId: string;
		fillType?: string;
		color?: string;
		transparency?: number;
		imageBase64?: string;
	};
	const {
		slideIndex,
		shapeId,
		fillType = "solid",
		color,
		transparency,
		imageBase64,
	} = config;

	return runInPowerPoint(async (ctx) => {
		const slide = ctx.presentation.slides.getItemAt(
			toZeroBasedSlideIndex(slideIndex),
		);
		const shape = slide.shapes.getItem(shapeId);
		const fill = shape.fill;

		if (fillType === "none") {
			fill.clear();
		} else if (fillType === "image" && imageBase64) {
			fill.setImage(imageBase64);
		} else if (color) {
			fill.setSolidColor(color);
		}

		if (transparency !== undefined) {
			fill.transparency = transparency;
		}

		await ctx.sync();
		return { slideIndex, shapeId, fillType, undoable: true };
	});
}

async function handleSetShapeLine(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex: number;
		shapeId: string;
		color?: string;
		width?: number;
		style?: string;
		visible?: boolean;
	};
	const { slideIndex, shapeId, color, width, style, visible = true } = config;

	return runInPowerPoint(async (ctx) => {
		const slide = ctx.presentation.slides.getItemAt(
			toZeroBasedSlideIndex(slideIndex),
		);
		const shape = slide.shapes.getItem(shapeId);
		const line = shape.lineFormat;

		if (color) line.color = color;
		if (width !== undefined) line.weight = width;
		if (style) line.style = style;
		line.visible = visible;

		await ctx.sync();
		return { slideIndex, shapeId, undoable: true };
	});
}

async function handleSetShapeRotation(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex: number;
		shapeId: string;
		degrees: number;
	};
	const { slideIndex, shapeId, degrees } = config;

	return runInPowerPoint(async (ctx) => {
		const slide = ctx.presentation.slides.getItemAt(
			toZeroBasedSlideIndex(slideIndex),
		);
		const shape = slide.shapes.getItem(shapeId);
		shape.rotation = degrees;
		await ctx.sync();

		return { slideIndex, shapeId, degrees, undoable: true };
	});
}

// ── Phase 18: Geometric Shapes & Lines ──────────────────────────

async function handleAddGeometricShape(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex: number;
		shapeType: string;
		left?: number;
		top?: number;
		width?: number;
		height?: number;
	};
	const {
		slideIndex,
		shapeType,
		left = 100,
		top = 100,
		width = 100,
		height = 100,
	} = config;

	return runInPowerPoint(async (ctx) => {
		const PowerPoint: any = (window as any).PowerPoint;
		const slide = ctx.presentation.slides.getItemAt(
			toZeroBasedSlideIndex(slideIndex),
		);
		const shape = slide.shapes.addGeometricShape(
			PowerPoint.GeometricShapeType[shapeType] || shapeType,
			{ left, top, width, height },
		);
		shape.load(["id", "name"]);
		await ctx.sync();

		return {
			slideIndex,
			shapeId: shape.id,
			shapeName: shape.name,
			shapeType,
			undoable: true,
		};
	});
}

async function handleAddLine(args: unknown): Promise<unknown> {
	const config = args as {
		slideIndex: number;
		startX: number;
		startY: number;
		endX: number;
		endY: number;
		connectorType?: string;
	};
	const {
		slideIndex,
		startX,
		startY,
		endX,
		endY,
		connectorType = "straight",
	} = config;

	return runInPowerPoint(async (ctx) => {
		const PowerPoint: any = (window as any).PowerPoint;
		const slide = ctx.presentation.slides.getItemAt(
			toZeroBasedSlideIndex(slideIndex),
		);
		const connector = PowerPoint.ConnectorType[connectorType] || connectorType;
		const shape = slide.shapes.addLine(connector, {
			left: startX,
			top: startY,
			width: endX - startX,
			height: endY - startY,
		});
		shape.load(["id", "name"]);
		await ctx.sync();

		return {
			slideIndex,
			shapeId: shape.id,
			shapeName: shape.name,
			undoable: true,
		};
	});
}

// ── Phase 18: Slide Merge ─────────────────────────────────────────

async function handleInsertSlidesFromFile(args: unknown): Promise<unknown> {
	const config = args as {
		base64File: string;
		insertAfterSlideIndex?: number;
		slideIndexes?: string;
		formatting?: string;
	};
	const {
		base64File,
		insertAfterSlideIndex,
		slideIndexes,
		formatting = "useDestinationTheme",
	} = config;

	return runInPowerPoint(async (ctx) => {
		const PowerPoint: any = (window as any).PowerPoint;
		const pres = ctx.presentation;

		const options: any = {};
		if (formatting === "keepSourceFormatting") {
			options.formatting =
				PowerPoint.InsertSlideFormatting.keepSourceFormatting;
		} else {
			options.formatting = PowerPoint.InsertSlideFormatting.useDestinationTheme;
		}

		if (insertAfterSlideIndex !== undefined) {
			pres.slides.load("items");
			await ctx.sync();
			const targetSlide =
				pres.slides.items[toZeroBasedSlideIndex(insertAfterSlideIndex)];
			if (targetSlide) {
				targetSlide.load("id");
				await ctx.sync();
				options.targetSlideId = targetSlide.id;
			}
		}

		if (slideIndexes) {
			options.sourceSlideIds = slideIndexes
				.split(",")
				.map((s: string) => String(toZeroBasedSlideIndex(parseInt(s.trim(), 10))));
		}

		pres.insertSlidesFromBase64(base64File, options);
		await ctx.sync();

		return { inserted: true, formatting, undoable: true };
	});
}

// ── Phase 18: Layouts & Theme ─────────────────────────────────────

async function handleGetLayouts(_args: unknown): Promise<unknown> {
	return runInPowerPoint(async (ctx) => {
		const masters = ctx.presentation.slideMasters;
		masters.load("items");
		await ctx.sync();

		const layouts: any[] = [];
		for (const master of masters.items) {
			master.load("name");
			master.layouts.load("items");
		}
		await ctx.sync();

		for (const master of masters.items) {
			for (const layout of master.layouts.items) {
				layout.load(["id", "name"]);
			}
		}
		await ctx.sync();

		for (const master of masters.items) {
			for (const layout of master.layouts.items) {
				layouts.push({ id: layout.id, name: layout.name, master: master.name });
			}
		}

		return { layouts, count: layouts.length };
	});
}

async function handleSetSlideLayout(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; layoutId: string; slideMasterId?: string };
	const { slideIndex = 1, layoutId, slideMasterId } = config;

	return runInPowerPoint(async (ctx) => {
		const resolved = await resolveSlide(ctx, slideIndex);
		if ("error" in resolved) return resolved;
		const { slide } = resolved;

		let layoutObj: any;

		if (slideMasterId !== undefined) {
			const master = ctx.presentation.slideMasters.getItemOrNullObject(slideMasterId);
			master.load("isNullObject");
			await ctx.sync();
			if (master.isNullObject) {
				return { error: `Slide master '${slideMasterId}' not found` };
			}
			layoutObj = master.layouts.getItemOrNullObject(layoutId);
			layoutObj.load("isNullObject");
			await ctx.sync();
			if (layoutObj.isNullObject) {
				return { error: `Layout '${layoutId}' not found on slide master '${slideMasterId}'` };
			}
		} else {
			const masters = ctx.presentation.slideMasters;
			masters.load("items");
			await ctx.sync();

			for (const master of masters.items) {
				master.layouts.load("items");
			}
			await ctx.sync();

			const candidates: any[] = [];
			for (const master of masters.items) {
				for (const layout of master.layouts.items) {
					layout.load(["id"]);
					candidates.push(layout);
				}
			}
			await ctx.sync();

			layoutObj = candidates.find((l) => l.id === layoutId);
			if (!layoutObj) {
				return { error: `Layout '${layoutId}' not found on any slide master` };
			}
		}

		slide.applyLayout(layoutObj);
		await ctx.sync();

		return { slideIndex, layoutId, slideMasterId: slideMasterId ?? null, applied: true, undoable: true };
	});
}

async function handleGetThemeColors(_args: unknown): Promise<unknown> {
	return runInPowerPoint(async (ctx) => {
		const masters = ctx.presentation.slideMasters;
		masters.load("items");
		await ctx.sync();

		if (masters.items.length === 0) return { colors: {}, count: 0 };

		const master = masters.items[0];
		const theme = master.themeColorScheme;
		theme.load([
			"name",
			"dark1",
			"light1",
			"dark2",
			"light2",
			"accent1",
			"accent2",
			"accent3",
			"accent4",
			"accent5",
			"accent6",
			"hyperlink",
			"followedHyperlink",
		]);
		await ctx.sync();

		return {
			themeName: theme.name,
			colors: {
				dark1: theme.dark1,
				light1: theme.light1,
				dark2: theme.dark2,
				light2: theme.light2,
				accent1: theme.accent1,
				accent2: theme.accent2,
				accent3: theme.accent3,
				accent4: theme.accent4,
				accent5: theme.accent5,
				accent6: theme.accent6,
				hyperlink: theme.hyperlink,
				followedHyperlink: theme.followedHyperlink,
			},
		};
	});
}

async function handleGroupShapes(args: unknown): Promise<unknown> {
	const config = args as { slideIndex: number; shapeIds: string };
	const { slideIndex, shapeIds } = config;

	return runInPowerPoint(async (ctx) => {
		const slide = ctx.presentation.slides.getItemAt(
			toZeroBasedSlideIndex(slideIndex),
		);
		const ids = shapeIds.split(",").map((s: string) => s.trim());
		const shapes: any[] = [];
		for (const id of ids) {
			shapes.push(slide.shapes.getItem(id));
		}

		const group = slide.shapes.addGroup(shapes);
		group.load(["id", "name"]);
		await ctx.sync();

		return {
			slideIndex,
			groupId: group.id,
			groupName: group.name,
			undoable: true,
		};
	});
}

async function handleUngroupShape(args: unknown): Promise<unknown> {
	const config = args as { slideIndex: number; shapeId: string };
	const { slideIndex, shapeId } = config;

	return runInPowerPoint(async (ctx) => {
		const slide = ctx.presentation.slides.getItemAt(
			toZeroBasedSlideIndex(slideIndex),
		);
		const shape = slide.shapes.getItem(shapeId);
		shape.group.ungroup();
		await ctx.sync();

		return { slideIndex, shapeId, ungrouped: true, undoable: true };
	});
}
