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
		const errorMessage = error instanceof Error ? error.message : String(error);
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
 */
function applyRangeProperties(config: Record<string, unknown>, range: PowerPoint.TextRange): string[] {
	return [
		...applyDefinedProperties(config, range.font, FONT_SETTERS),
		...applyDefinedProperties(config, range.paragraphFormat, PARAGRAPH_FORMAT_SETTERS),
		...applyDefinedProperties(config, range.paragraphFormat.bulletFormat, BULLET_FORMAT_SETTERS),
	];
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
 * Resolves a slide index + shapeId (id or name) to a live Shape object.
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

	if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
		return { error: `Slide index ${slideIndex} out of range` };
	}

	const slide = pres.slides.items[slideIndex];
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

// ── Read tools ──────────────────────────────────────────────────

async function handleGetDeckOutline(args: unknown): Promise<unknown> {
	const config = args as { startSlide?: number; endSlide?: number };

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const totalSlides = pres.slides.items.length;

		const from = Math.max(0, config.startSlide ?? 0);
		const to = Math.min(totalSlides - 1, config.endSlide ?? totalSlides - 1);
		if (from > to) {
			return { error: `Invalid slide range: startSlide (${from}) is after endSlide (${to}).` };
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
				index: i,
				title: title || `Slide ${i + 1}`,
				shapes: shapeData,
			});
		}

		return { documentName: "Presentation", totalSlides, slides: slideList };
	});
}

async function handleGetSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const slideIndex = config.slideIndex ?? 0;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return {
				error: `Slide index ${slideIndex} out of range (0-${pres.slides.items.length - 1})`,
			};
		}

		const slide = pres.slides.items[slideIndex];

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
					font = {
						name: safeStr(tf.textRange.font.name),
						size: safeNum(tf.textRange.font.size),
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
			title: slideTitle || `Slide ${slideIndex + 1}`,
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
	const slideIndex = config.slideIndex ?? 0;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
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
	const slideIndex = config.slideIndex ?? 0;
	const shapeId = config.shapeId ?? "";

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
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
	const slideIndex = config.slideIndex ?? 0;
	const shapeId = config.shapeId ?? "";

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
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
		const cellTexts: any[] = [];
		for (let r = 0; r < rows; r++) {
			for (let c = 0; c < cols; c++) {
				const cell = table.getCell(r, c);
				const tf = cell.textFrame;
				ctx.load(tf, "textRange/text");
				cellTexts.push(tf);
			}
		}
		await ctx.sync();

		// Fill in the cells array
		const cells: string[][] = [];
		let idx = 0;
		for (let r = 0; r < rows; r++) {
			const row: string[] = [];
			for (let c = 0; c < cols; c++) {
				const tf = cellTexts[idx++];
				row.push(safeStr(tf.textRange?.text));
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

async function handleGetShapeParagraphs(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string };
	const slideIndex = config.slideIndex ?? 0;
	const shapeId = config.shapeId ?? "";

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

		// Dedup by content: identical diffs (even for non-adjacent
		// paragraphs) share one groupId/propertyGroups entry.
		const groupIdByDiffKey = new Map<string, number>();
		const propertyGroups: { groupId: number; properties: Record<string, unknown> }[] = [];
		const paragraphs = paragraphSpans.map((span, index) => {
			const diff = diffRangeProperties(extractRangeProperties(paragraphRanges[index]), defaultProperties);
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
			slideIndex,
			shapeId,
			fullText,
			defaultProperties,
			propertyGroups,
			paragraphs,
		};
	});
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
export function paragraphsToMarkdown(
	paragraphs: { text: string; groupId: number }[],
	propertyGroups: { groupId: number; properties: Record<string, any> }[],
	defaultProperties: RangeProperties,
	options: { bulletChar?: string } = {},
): string {
	const bulletChar = options.bulletChar ?? "-";
	const diffByGroupId = new Map(propertyGroups.map((g) => [g.groupId, g.properties]));

	// Running per-indent-level numbering state, indexed by level. Truncated
	// whenever a shallower paragraph is seen, so a deeper list always
	// restarts at 1 the next time that level is reused.
	const levelState: { count: number; wasNumbered: boolean }[] = [];

	return paragraphs
		.map((p) => {
			const diff = diffByGroupId.get(p.groupId) ?? {};
			const defaultPf = defaultProperties.paragraphFormat as { indentLevel?: number; bulletType?: string };
			const indentLevel = safeNum(diff.paragraphFormat?.indentLevel, defaultPf.indentLevel);
			const bulletType = safeStr(diff.paragraphFormat?.bulletType, defaultPf.bulletType);
			const isNumbered = bulletType === "Numbered";

			levelState.length = Math.min(levelState.length, indentLevel + 1);
			const prev = levelState[indentLevel];
			const count = isNumbered && prev?.wasNumbered ? prev.count + 1 : 1;
			levelState[indentLevel] = { count, wasNumbered: isNumbered };

			const marker = isNumbered ? `${count}.` : bulletChar;
			return "  ".repeat(indentLevel) + marker + " " + p.text;
		})
		.join("\n");
}

async function handleGetShapeTextMarkdown(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; shapeId?: string; bulletChar?: string };
	const slideIndex = config.slideIndex ?? 0;
	const shapeId = config.shapeId ?? "";
	const bulletChar = config.bulletChar ?? "-";

	if (bulletChar !== "-" && bulletChar !== "*") {
		return { error: `bulletChar must be '-' or '*', got '${bulletChar}'` };
	}

	const result = (await handleGetShapeParagraphs({ slideIndex, shapeId })) as any;
	if (result.error) return result;

	const markdown = paragraphsToMarkdown(result.paragraphs, result.propertyGroups, result.defaultProperties, {
		bulletChar,
	});

	return { slideIndex, shapeId, markdown };
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
			indices = [slideIndex];
		} else if (config.slideRange) {
			const match = config.slideRange.match(/^(\d+)-(\d+)$/);
			if (match) {
				const from = parseInt(match[1]);
				const to = parseInt(match[2]);
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
				notes.push({ slideIndex: idx, notes: "" });
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
			notes.push({ slideIndex: indices[i], notes: noteText });
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
	const { slideIndex = 0, shapeId = "", text = "" } = config;

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
	const { slideIndex = 0, shapeId = "" } = config;

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
	const { slideIndex = 0, shapeId = "", start = 0, length = 0 } = config;

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

		const updated = applyRangeProperties(config, range);

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
		slideIndex = 0,
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

			updated = applyRangeProperties(config, range);

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
	const { slideIndex = 0, shapeId = "", paragraphStart = 0, paragraphLength = 0 } = config;

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

async function handleUpdateSpeakerNotes(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; notes?: string };
	const { slideIndex = 0, notes = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		const notesSlide = slide.getNotesSlideOrNullObject();
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
		slideIndex = 0,
		text = "",
		left = 100,
		top = 100,
		width = 300,
		height = 100,
	} = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
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
	const { slideIndex = 0, imageBase64 = "", left = 100, top = 100 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];

		// Strip data URI prefix if present
		const base64Data = imageBase64.replace(/^data:image\/[^;]+;base64,/, "");

		const options: any = { left, top };
		if (config.width !== undefined) options.width = config.width;
		if (config.height !== undefined) options.height = config.height;

		const picture = slide.shapes.addPicture(base64Data, options);
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
		slideIndex = 0,
		rows = 2,
		columns = 2,
		left = 100,
		top = 100,
	} = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];

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
	const { slideIndex = 0, shapeId = "" } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
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
			options.index = config.atIndex;
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
				? Math.min(config.atIndex, pres.slides.items.length - 1)
				: pres.slides.items.length - 1;

		const newSlide = pres.slides.items[newIndex];
		newSlide.load("id");
		await ctx.sync();

		return {
			slideIndex: newIndex,
			slideId: safeStr(newSlide.id),
		};
	});
}

async function handleDeleteSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const { slideIndex = 0 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const slide = pres.slides.items[slideIndex];
		slide.delete();
		await ctx.sync();

		return { slideIndex, deleted: true };
	});
}

async function handleMoveSlide(args: unknown): Promise<unknown> {
	const config = args as { fromIndex?: number; toIndex?: number };
	const { fromIndex = 0, toIndex = 0 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (fromIndex < 0 || fromIndex >= pres.slides.items.length) {
			return { error: `fromIndex ${fromIndex} out of range` };
		}

		const slide = pres.slides.items[fromIndex];
		slide.load("id");
		await ctx.sync();

		const slideId = safeStr(slide.id);
		slide.moveTo(toIndex);
		await ctx.sync();

		return { fromIndex, toIndex, slideId };
	});
}

async function handleDuplicateSlide(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number; targetIndex?: number };
	const { slideIndex = 0, targetIndex } = config;

	return runInPowerPoint(async (ctx) => {
		const PowerPoint: any = (window as any).PowerPoint;
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const slideCount = pres.slides.items.length;
		if (slideIndex < 0 || slideIndex >= slideCount) {
			return { error: `Slide index ${slideIndex} out of range (deck has ${slideCount} slides)` };
		}

		// Validate before mutating: once the duplicate is inserted the deck has slideCount + 1
		// slides, so valid target positions are 0..slideCount inclusive (slideCount itself means
		// "append after the new last slide"). Checking this up front avoids leaving a stray
		// duplicate behind when the request is rejected.
		if (targetIndex !== undefined && (targetIndex < 0 || targetIndex > slideCount)) {
			return { error: `targetIndex ${targetIndex} out of range (deck has ${slideCount} slides; valid range is 0-${slideCount})` };
		}

		const sourceSlide = pres.slides.items[slideIndex];
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

		const newSlideIndex = slideIndex + 1;
		const newSlide = pres.slides.items[newSlideIndex];
		newSlide.load("id");
		await ctx.sync();
		const newSlideId = safeStr(newSlide.id);

		let finalIndex = newSlideIndex;
		if (targetIndex !== undefined && targetIndex !== newSlideIndex) {
			finalIndex = Math.max(
				0,
				Math.min(targetIndex, pres.slides.items.length - 1),
			);
			newSlide.moveTo(finalIndex);
			await ctx.sync();
		}

		return {
			sourceIndex: slideIndex,
			newSlideIndex: finalIndex,
			newSlideId,
		};
	});
}

// Internal-only: exports a single slide as base64 for cross-document duplication.
// Not part of the public tool list — invoked by the server when
// powerpoint_duplicate_slide targets a different instance.
async function handleExportSlideInternal(args: unknown): Promise<unknown> {
	const config = args as { slideIndex?: number };
	const { slideIndex = 0 } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		if (slideIndex < 0 || slideIndex >= pres.slides.items.length) {
			return { error: `Slide index ${slideIndex} out of range` };
		}

		const sourceSlide = pres.slides.items[slideIndex];
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
		// slides, so valid target positions are 0..beforeCount inclusive (beforeCount itself
		// means "append after the new last slide"). Checking this up front avoids leaving a
		// stray imported slide behind when the request is rejected.
		if (targetIndex !== undefined && (targetIndex < 0 || targetIndex > beforeCount)) {
			return { error: `targetIndex ${targetIndex} out of range (target deck has ${beforeCount} slides; valid range is 0-${beforeCount})` };
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
		if (targetIndex !== undefined && targetIndex !== newSlideIndex) {
			finalIndex = Math.max(
				0,
				Math.min(targetIndex, pres.slides.items.length - 1),
			);
			newSlide.moveTo(finalIndex);
			await ctx.sync();
		}

		return { newSlideIndex: finalIndex, newSlideId };
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
			tagTarget = ctx.presentation.slides.getItemAt(slideIndex);
		} else if (target === "shape" && slideIndex !== undefined && shapeId) {
			const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
			tagTarget = ctx.presentation.slides.getItemAt(slideIndex);
		} else if (target === "shape" && slideIndex !== undefined && shapeId) {
			const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
		const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
		const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
		const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
		const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
		const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
			const targetSlide = pres.slides.items[insertAfterSlideIndex];
			if (targetSlide) {
				targetSlide.load("id");
				await ctx.sync();
				options.targetSlideId = targetSlide.id;
			}
		}

		if (slideIndexes) {
			options.sourceSlideIds = slideIndexes
				.split(",")
				.map((s: string) => s.trim());
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
	const { slideIndex = 0, layoutId, slideMasterId } = config;

	return runInPowerPoint(async (ctx) => {
		const pres = ctx.presentation;
		pres.load("slides");
		await ctx.sync();

		const slideCount = pres.slides.items.length;
		if (slideIndex < 0 || slideIndex >= slideCount) {
			return { error: `Slide index ${slideIndex} out of range (deck has ${slideCount} slides)` };
		}

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

		const slide = pres.slides.items[slideIndex];
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
		const slide = ctx.presentation.slides.getItemAt(slideIndex);
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
		const slide = ctx.presentation.slides.getItemAt(slideIndex);
		const shape = slide.shapes.getItem(shapeId);
		shape.group.ungroup();
		await ctx.sync();

		return { slideIndex, shapeId, ungrouped: true, undoable: true };
	});
}
