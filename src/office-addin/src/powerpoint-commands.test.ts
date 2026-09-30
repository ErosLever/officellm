/**
 * Unit tests for powerpoint-commands.ts using the mock framework.
 *
 * Tests all 17 command handlers with simulated Office JS API.
 * The mock faithfully replicates load()/sync() patterns.
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { PowerPointMock, type MockPresentationData } from "./powerpoint-mock";

// ── Mock bridge (vi.mock is hoisted, so use module-level reference) ──
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

// Import AFTER vi.mock (vitest hoists the mock above this)
import {
	processCommand,
	toAlphabetCounter,
	toRomanNumeral,
	markdownToParagraphSpecs,
	wrapInlineMarkers,
	parseInlineMarkers,
	groupShapesIntoRows,
	isHeadingSize,
	isFootnoteShape,
} from "./powerpoint-commands";

// ── Test data ───────────────────────────────────────────────────

function makeTestDeck(): MockPresentationData {
	return {
		slideWidth: 960,
		slideHeight: 540,
		slides: [
			{
				id: "slide_0",
				shapes: [
					{
						id: "s1",
						name: "Title 1",
						type: "TextBox",
						text: "Quarterly Results",
						left: 50,
						top: 30,
						width: 600,
						height: 60,
						font: { name: "Calibri", size: 36, bold: true },
					},
					{
						id: "s2",
						name: "Content 1",
						type: "TextBox",
						text: "Revenue grew 15%",
						left: 50,
						top: 120,
						width: 600,
						height: 200,
					},
					{
						id: "s3",
						name: "Picture 1",
						type: "Image",
						left: 50,
						top: 350,
						width: 300,
						height: 200,
					},
					{
						id: "s4",
						name: "Slide Number Placeholder",
						type: "TextBox",
						text: "1",
						left: 700,
						top: 500,
						width: 30,
						height: 20,
						font: { size: 10 },
					},
					{
						id: "s7",
						name: "Bullet List 1",
						type: "TextBox",
						// 5 paragraphs: 0/2/4 share default formatting, 1 has a
						// distinct indentLevel, 3 has a distinct font color —
						// exercises dedup-by-content and non-adjacent grouping.
						text: "Alpha\rBeta\rGamma\rDelta\rEpsilon",
						left: 50,
						top: 250,
						width: 600,
						height: 300,
						font: { name: "Calibri", size: 24, color: "#000000" },
						paragraphFormat: { horizontalAlignment: "Left", indentLevel: 0, bulletType: "Unnumbered" },
						paragraphs: [
							undefined,
							{ paragraphFormat: { indentLevel: 2 } },
							undefined,
							{ font: { color: "#FF0000" } },
							undefined,
						] as any,
					},
				],
				notes: "Talk about revenue growth",
			},
			{
				id: "slide_1",
				shapes: [
					{
						id: "s5",
						name: "Title 1",
						type: "TextBox",
						text: "Pricing Table",
						left: 50,
						top: 30,
						width: 600,
						height: 60,
					},
					{
						id: "s6",
						name: "Table 1",
						type: "Table",
						left: 50,
						top: 120,
						width: 600,
						height: 200,
						tableCells: [
							["Product", "Price", "Stock"],
							["Widget A", "$10", "100"],
							["Widget B", "$20", "50"],
						],
					},
				],
				notes: "",
			},
			{ id: "slide_2", shapes: [], notes: "" },
		],
	};
}

// ── Test harness ─────────────────────────────────────────────────

let mock: PowerPointMock;

beforeEach(() => {
	// Node.js doesn't have `window` — polyfill it
	(globalThis as any).window = globalThis;

	mock = new PowerPointMock(makeTestDeck());
	mock.install();
	mockBridge.reportResult = mock.mockReportResult;
});

// Need to re-import after mock — but vitest hoists vi.mock automatically

const BULLET_STYLE_MARKDOWN_CASES: { style: string; markers: string[] }[] = [
	{ style: "ArabicNumeralPlain", markers: ["1", "2", "3"] },
	{ style: "ArabicNumeralPeriod", markers: ["1.", "2.", "3."] },
	{ style: "ArabicNumeralParenthesisRight", markers: ["1)", "2)", "3)"] },
	{ style: "ArabicNumeralParenthesesBoth", markers: ["(1)", "(2)", "(3)"] },
	{ style: "AlphabetLowercasePeriod", markers: ["a.", "b.", "c."] },
	{ style: "AlphabetLowercaseParenthesisRight", markers: ["a)", "b)", "c)"] },
	{ style: "AlphabetLowercaseParenthesesBoth", markers: ["(a)", "(b)", "(c)"] },
	{ style: "AlphabetUppercasePeriod", markers: ["A.", "B.", "C."] },
	{ style: "AlphabetUppercaseParenthesisRight", markers: ["A)", "B)", "C)"] },
	{ style: "AlphabetUppercaseParenthesesBoth", markers: ["(A)", "(B)", "(C)"] },
	{ style: "RomanLowercasePeriod", markers: ["i.", "ii.", "iii."] },
	{ style: "RomanLowercaseParenthesisRight", markers: ["i)", "ii)", "iii)"] },
	{ style: "RomanLowercaseParenthesesBoth", markers: ["(i)", "(ii)", "(iii)"] },
	{ style: "RomanUppercasePeriod", markers: ["I.", "II.", "III."] },
	{ style: "RomanUppercaseParenthesisRight", markers: ["I)", "II)", "III)"] },
	{ style: "RomanUppercaseParenthesesBoth", markers: ["(I)", "(II)", "(III)"] },
];

// ── READ TOOLS ───────────────────────────────────────────────────

describe("powerpoint_get_deck_outline", () => {
	it("returns all slides with shape properties", async () => {
		const result = (await processCommand(
			"cmd-1",
			"powerpoint_get_deck_outline",
			{},
		)) as any;

		expect(result.totalSlides).toBe(3);
		expect(result.slides[0].title).toBe("Quarterly Results");
		expect(result.slides[0].shapes).toHaveLength(5);
		expect(result.slides[0].shapes[0].type).toBe("TextBox");
		expect(result.slides[0].shapes[0].left).toBe(50);
		expect(result.slides[0].shapes[0].width).toBe(600);
	});

	it("skips non-content shapes for title detection", async () => {
		const result = (await processCommand(
			"cmd-2",
			"powerpoint_get_deck_outline",
			{},
		)) as any;
		// "Slide Number Placeholder" should be skipped for title
		expect(result.slides[0].title).toBe("Quarterly Results");
	});

	it("falls back to 'Slide N' for slides without titles", async () => {
		const result = (await processCommand(
			"cmd-3",
			"powerpoint_get_deck_outline",
			{},
		)) as any;
		expect(result.slides[2].title).toBe("Slide 3");
	});

	it("scopes to a startSlide/endSlide range without touching other slides", async () => {
		const result = (await processCommand(
			"cmd-3b",
			"powerpoint_get_deck_outline",
			{ startSlide: 2, endSlide: 2 },
		)) as any;

		expect(result.totalSlides).toBe(3);
		expect(result.slides).toHaveLength(1);
		expect(result.slides[0].index).toBe(2);
		expect(result.slides[0].title).toBe("Pricing Table");
	});

	it("clamps an out-of-range endSlide to the last slide", async () => {
		const result = (await processCommand(
			"cmd-3c",
			"powerpoint_get_deck_outline",
			{ startSlide: 3, endSlide: 99 },
		)) as any;

		expect(result.slides).toHaveLength(1);
		expect(result.slides[0].index).toBe(3);
	});

	it("errors when startSlide is after endSlide", async () => {
		const result = (await processCommand(
			"cmd-3d",
			"powerpoint_get_deck_outline",
			{ startSlide: 3, endSlide: 1 },
		)) as any;

		expect(result.error).toBeDefined();
	});
});

describe("powerpoint_get_slide", () => {
	it("returns full shape properties for a slide", async () => {
		const result = (await processCommand("cmd-4", "powerpoint_get_slide", {
			slideIndex: 1,
		})) as any;

		expect(result.slideIndex).toBe(1);
		expect(result.shapes).toHaveLength(5);

		const title = result.shapes[0];
		expect(title.id).toBe("s1");
		expect(title.name).toBe("Title 1");
		expect(title.type).toBe("TextBox");
		expect(title.left).toBe(50);
		expect(title.top).toBe(30);
		expect(title.width).toBe(600);
		expect(title.height).toBe(60);
		expect(title.rotation).toBe(0);
		expect(title.text).toBe("Quarterly Results");
	});

	it("reports an empty layout when no layout has been applied", async () => {
		const result = (await processCommand("cmd-4b", "powerpoint_get_slide", {
			slideIndex: 1,
		})) as any;

		expect(result.layout).toEqual({ id: "", name: "" });
	});

	it("reflects the layout applied via powerpoint_set_slide_layout", async () => {
		await processCommand("cmd-4c", "powerpoint_set_slide_layout", {
			slideIndex: 1,
			layoutId: "layout_1",
		});

		const result = (await processCommand("cmd-4d", "powerpoint_get_slide", {
			slideIndex: 1,
		})) as any;

		expect(result.layout).toEqual({ id: "layout_1", name: "Title and Content" });
	});

	it("includes font properties for text shapes", async () => {
		const result = (await processCommand("cmd-5", "powerpoint_get_slide", {
			slideIndex: 1,
		})) as any;
		const title = result.shapes[0];

		expect(title.font).toBeDefined();
		expect(title.font.name).toBe("Calibri");
		expect(title.font.size).toBe(36);
		expect(title.font.bold).toBe(true);
	});

	it("returns error for out-of-range index", async () => {
		const result = (await processCommand("cmd-6", "powerpoint_get_slide", {
			slideIndex: 99,
		})) as any;
		expect(result.error).toContain("out of range");
	});

	it("resolves a placeholder's inherited font size from its layout when the shape reports the mixed/unresolved size sentinel", async () => {
		// Mirrors live Office.js: an unfilled Title placeholder with no
		// run-level font override reports textRange.font.size as the
		// mixed/unresolved sentinel (0 after safeNum), even though it
		// visually renders at the size set on the slide layout's own Title
		// placeholder.
		mock = new PowerPointMock({
			slideWidth: 960,
			slideHeight: 540,
			slides: [
				{
					id: "slide_0",
					layoutId: "layout_title",
					shapes: [
						{
							id: "sTitle",
							name: "Title 1",
							type: "TextBox",
							text: "Third parties: CVEs",
							placeholderType: "Title",
						},
					],
					notes: "",
				},
			],
			slideMasters: [
				{
					id: "master_0",
					name: "Office Theme",
					layouts: [
						{
							id: "layout_title",
							name: "Title and Content",
							shapes: [
								{
									id: "layoutTitle",
									name: "Title Placeholder 1",
									type: "TextBox",
									text: "",
									placeholderType: "Title",
									font: { size: 35 },
								},
							],
						},
					],
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand("cmd-6b", "powerpoint_get_slide", {
			slideIndex: 1,
		})) as any;

		expect(result.shapes[0].font.size).toBe(35);
	});

	it("leaves a non-placeholder shape's unresolved font size at 0 rather than throwing", async () => {
		mock = new PowerPointMock({
			slideWidth: 960,
			slideHeight: 540,
			slides: [
				{
					id: "slide_0",
					layoutId: "layout_title",
					shapes: [
						{
							id: "sPlain",
							name: "Plain Box",
							type: "TextBox",
							text: "Not a placeholder",
						},
					],
					notes: "",
				},
			],
			slideMasters: [
				{
					id: "master_0",
					name: "Office Theme",
					layouts: [
						{
							id: "layout_title",
							name: "Title and Content",
							shapes: [
								{
									id: "layoutDecor",
									name: "Decoration",
									type: "TextBox",
									text: "",
								},
							],
						},
					],
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand("cmd-6c", "powerpoint_get_slide", {
			slideIndex: 1,
		})) as any;

		expect(result.shapes[0].font.size).toBe(0);
	});

	it("samples a single character's font size when the whole-range getter reports the mixed sentinel despite a uniform real size", async () => {
		// Mirrors a live Office.js quirk (seen on an all-35pt title): the
		// whole-range textRange.font.size reports the mixed/unresolved
		// sentinel even though every character in the range shares one real,
		// explicit size. The layout's own placeholder has a DIFFERENT size
		// (25) than the shape's real per-character size (35), so if the
		// sampling step regressed back to jumping straight to layout
		// inheritance, this test would catch it by observing 25 instead of
		// the true 35.
		mock = new PowerPointMock({
			slideWidth: 960,
			slideHeight: 540,
			slides: [
				{
					id: "slide_0",
					layoutId: "layout_title",
					shapes: [
						{
							id: "sTitle",
							name: "Title 1",
							type: "TextBox",
							text: "Third parties: CVEs",
							placeholderType: "Title",
							paragraphs: [
								{ runs: [{ start: 0, length: 1, font: { size: 35 } }] },
							],
						},
					],
					notes: "",
				},
			],
			slideMasters: [
				{
					id: "master_0",
					name: "Office Theme",
					layouts: [
						{
							id: "layout_title",
							name: "Title and Content",
							shapes: [
								{
									id: "layoutTitle",
									name: "Title Placeholder 1",
									type: "TextBox",
									text: "",
									placeholderType: "Title",
									font: { size: 25 },
								},
							],
						},
					],
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand("cmd-6e", "powerpoint_get_slide", {
			slideIndex: 1,
		})) as any;

		expect(result.shapes[0].font.size).toBe(35);
	});
});

describe("powerpoint_get_slide_image", () => {
	it("returns base64 PNG image", async () => {
		const result = (await processCommand(
			"cmd-7",
			"powerpoint_get_slide_image",
			{ slideIndex: 1 },
		)) as any;

		expect(result.slideIndex).toBe(1);
		expect(result.image).toMatch(/^data:image\/png;base64,/);
	});
});

describe("powerpoint_get_shape_image", () => {
	it("returns image for a shape by ID", async () => {
		const result = (await processCommand(
			"cmd-8",
			"powerpoint_get_shape_image",
			{ slideIndex: 1, shapeId: "s3" },
		)) as any;

		expect(result.slideIndex).toBe(1);
		expect(result.shapeId).toBe("s3");
		expect(result.image).toMatch(/^data:image\/png;base64,/);
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand(
			"cmd-9",
			"powerpoint_get_shape_image",
			{ slideIndex: 1, shapeId: "nonexistent" },
		)) as any;
		expect(result.error).toContain("not found");
	});
});

describe("powerpoint_get_table", () => {
	it("returns table cells as 2D array", async () => {
		const result = (await processCommand("cmd-10", "powerpoint_get_table", {
			slideIndex: 2,
			shapeId: "s6",
		})) as any;

		expect(result.rowCount).toBe(3);
		expect(result.columnCount).toBe(3);
		expect(result.cells[0]).toEqual(["Product", "Price", "Stock"]);
		expect(result.cells[1]).toEqual(["Widget A", "$10", "100"]);
		expect(result.cells[2]).toEqual(["Widget B", "$20", "50"]);
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand("cmd-11", "powerpoint_get_table", {
			slideIndex: 2,
			shapeId: "nope",
		})) as any;
		expect(result.error).toContain("not found");
	});
});

describe("powerpoint_get_shape_paragraphs", () => {
	it("returns fullText and defaultProperties matching the shape's whole-range values", async () => {
		const result = (await processCommand(
			"cmd-gp1",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(result.fullText).toBe("Alpha\rBeta\rGamma\rDelta\rEpsilon");
		expect(result.defaultProperties.font).toMatchObject({
			fontName: "Calibri",
			fontSize: 24,
			color: "#000000",
		});
		expect(result.defaultProperties.paragraphFormat).toMatchObject({
			horizontalAlignment: "Left",
			indentLevel: 0,
		});
	});

	it("groups non-adjacent paragraphs with identical formatting under the same groupId", async () => {
		const result = (await processCommand(
			"cmd-gp2",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		const [p0, p1, p2, p3, p4] = result.paragraphs;
		expect(p0.groupId).toBe(p2.groupId);
		expect(p0.groupId).toBe(p4.groupId);
		expect(p1.groupId).not.toBe(p0.groupId);
		expect(p3.groupId).not.toBe(p0.groupId);
		expect(p1.groupId).not.toBe(p3.groupId);
	});

	it("gives a paragraph matching defaultProperties exactly an empty properties diff", async () => {
		const result = (await processCommand(
			"cmd-gp3",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		const group0 = result.propertyGroups.find(
			(g: any) => g.groupId === result.paragraphs[0].groupId,
		);
		expect(group0.properties).toEqual({});
	});

	it("isolates a distinct indentLevel override to only its own group", async () => {
		const result = (await processCommand(
			"cmd-gp4",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		const group1 = result.propertyGroups.find(
			(g: any) => g.groupId === result.paragraphs[1].groupId,
		);
		expect(group1.properties).toEqual({
			paragraphFormat: { indentLevel: 2 },
		});
	});

	it("isolates a distinct font color override to only its own group", async () => {
		const result = (await processCommand(
			"cmd-gp5",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		const group3 = result.propertyGroups.find(
			(g: any) => g.groupId === result.paragraphs[3].groupId,
		);
		expect(group3.properties).toEqual({ font: { color: "#FF0000" } });
	});

	it("reports each paragraph's text and start/length spans", async () => {
		const result = (await processCommand(
			"cmd-gp6",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(result.paragraphs.map((p: any) => p.text)).toEqual([
			"Alpha",
			"Beta",
			"Gamma",
			"Delta",
			"Epsilon",
		]);
		expect(result.paragraphs[1].start).toBe(6); // "Alpha\r".length
		expect(result.paragraphs[1].length).toBe(4); // "Beta".length
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand(
			"cmd-gp7",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "nope" },
		)) as any;
		expect(result.error).toContain("not found");
	});

	it("returns error for shape with no text frame", async () => {
		const result = (await processCommand(
			"cmd-gp8",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s3" }, // s3 is an Image
		)) as any;
		expect(result.error).toContain("does not support text");
	});

	it("normalizes live Office.js's numeric-string enum encoding back to named values", async () => {
		// Observed on live Office.js (Mac desktop, PowerPointApi 1.10):
		// horizontalAlignment/bulletFormat.type/bulletFormat.style come back as
		// the 0-based index into the enum's declaration order (e.g. "2" for
		// BulletType.numbered) instead of the named string the @types/office-js
		// declarations promise. Simulate that raw encoding via fixture input.
		mock = new PowerPointMock({
			slides: [
				{
					id: "slide_0",
					shapes: [
						{
							id: "sE",
							name: "Enum Encoding",
							type: "TextBox",
							text: "Only",
							paragraphFormat: {
								horizontalAlignment: "0",
								bulletType: "2",
								bulletStyle: "4",
							},
						},
					],
					notes: "",
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-gp9",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "sE" },
		)) as any;

		expect(result.defaultProperties.paragraphFormat).toMatchObject({
			horizontalAlignment: "Left",
			bulletType: "Numbered",
			bulletStyle: "ArabicNumeralPeriod",
		});
	});

	it("resolves defaultProperties.font.fontSize from the layout when the shape's own size is the mixed/unresolved sentinel", async () => {
		mock = new PowerPointMock({
			slideWidth: 960,
			slideHeight: 540,
			slides: [
				{
					id: "slide_0",
					layoutId: "layout_title",
					shapes: [
						{
							id: "sTitle",
							name: "Title 1",
							type: "TextBox",
							text: "Third parties: CVEs",
							placeholderType: "Title",
						},
					],
					notes: "",
				},
			],
			slideMasters: [
				{
					id: "master_0",
					name: "Office Theme",
					layouts: [
						{
							id: "layout_title",
							name: "Title and Content",
							shapes: [
								{
									id: "layoutTitle",
									name: "Title Placeholder 1",
									type: "TextBox",
									text: "",
									placeholderType: "Title",
									font: { size: 35 },
								},
							],
						},
					],
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-gp10",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "sTitle" },
		)) as any;

		expect(result.defaultProperties.font.fontSize).toBe(35);
	});

	it("leaves defaultProperties.font.fontSize at 0 when the layout's placeholder counterpart is itself unresolved", async () => {
		mock = new PowerPointMock({
			slideWidth: 960,
			slideHeight: 540,
			slides: [
				{
					id: "slide_0",
					layoutId: "layout_title",
					shapes: [
						{
							id: "sTitle",
							name: "Title 1",
							type: "TextBox",
							text: "Third parties: CVEs",
							placeholderType: "Title",
						},
					],
					notes: "",
				},
			],
			slideMasters: [
				{
					id: "master_0",
					name: "Office Theme",
					layouts: [
						{
							id: "layout_title",
							name: "Title and Content",
							shapes: [
								{
									id: "layoutTitle",
									name: "Title Placeholder 1",
									type: "TextBox",
									text: "",
									placeholderType: "Title",
								},
							],
						},
					],
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-gp11",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "sTitle" },
		)) as any;

		expect(result.defaultProperties.font.fontSize).toBe(0);
	});

	it("samples a single character's font size when the whole-range and full-length-substring getters both report the mixed sentinel despite a uniform real size", async () => {
		// Same live Office.js quirk as the powerpoint_get_slide regression
		// test above, exercised through the paragraph-diffing path: the
		// layout's placeholder counterpart has a DIFFERENT size (25) than
		// the shape's real per-character size (35), so a regression back to
		// layout-inheritance-first would surface as 25 instead of 35 here.
		mock = new PowerPointMock({
			slideWidth: 960,
			slideHeight: 540,
			slides: [
				{
					id: "slide_0",
					layoutId: "layout_title",
					shapes: [
						{
							id: "sTitle",
							name: "Title 1",
							type: "TextBox",
							text: "Third parties: CVEs",
							placeholderType: "Title",
							paragraphs: [
								{ runs: [{ start: 0, length: 1, font: { size: 35 } }] },
							],
						},
					],
					notes: "",
				},
			],
			slideMasters: [
				{
					id: "master_0",
					name: "Office Theme",
					layouts: [
						{
							id: "layout_title",
							name: "Title and Content",
							shapes: [
								{
									id: "layoutTitle",
									name: "Title Placeholder 1",
									type: "TextBox",
									text: "",
									placeholderType: "Title",
									font: { size: 25 },
								},
							],
						},
					],
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-gp12",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "sTitle" },
		)) as any;

		expect(result.defaultProperties.font.fontSize).toBe(35);
		const group0 = result.propertyGroups.find(
			(g: any) => g.groupId === result.paragraphs[0].groupId,
		);
		expect(group0.properties).toEqual({});
	});
});

describe("powerpoint_get_shape_text_markdown", () => {
	it("renders each paragraph as a '- ' bullet indented by its indentLevel", async () => {
		const result = (await processCommand(
			"cmd-md1",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(result.markdown).toBe(
			["- Alpha", "    - Beta", "- Gamma", "- Delta", "- Epsilon"].join("\n"),
		);
	});

	it("uses '*' when bulletChar is '*'", async () => {
		const result = (await processCommand(
			"cmd-md2",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "s7", bulletChar: "*" },
		)) as any;

		expect(result.markdown.split("\n")[0]).toBe("* Alpha");
	});

	it("rejects a bulletChar other than '-' or '*'", async () => {
		const result = (await processCommand(
			"cmd-md3",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "s7", bulletChar: "+" },
		)) as any;

		expect(result.error).toContain("bulletChar");
	});

	it("numbers contiguous 'Numbered' paragraphs within an indent level, restarting after an interruption", async () => {
		mock = new PowerPointMock({
			slides: [
				{
					id: "slide_0",
					shapes: [
						{
							id: "sN",
							name: "Numbered List",
							type: "TextBox",
							text: "One\rTwo\rNested\rThree\rFour",
							paragraphFormat: { indentLevel: 0, bulletType: "Numbered" },
							paragraphs: [
								undefined,
								undefined,
								{ paragraphFormat: { indentLevel: 1, bulletType: "Numbered" } },
								{ paragraphFormat: { bulletType: "Unnumbered" } },
								undefined,
							] as any,
						},
					],
					notes: "",
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-md4",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "sN" },
		)) as any;

		expect(result.markdown).toBe(
			["1. One", "2. Two", "  1. Nested", "- Three", "1. Four"].join("\n"),
		);
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand(
			"cmd-md5",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "nope" },
		)) as any;
		expect(result.error).toContain("not found");
	});

	it("returns error for shape with no text frame", async () => {
		const result = (await processCommand(
			"cmd-md6",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "s3" },
		)) as any;
		expect(result.error).toContain("does not support text");
	});

	it.each(BULLET_STYLE_MARKDOWN_CASES)(
		"renders $style with its real bullet markers",
		async ({ style, markers }) => {
			mock = new PowerPointMock({
				slides: [
					{
						id: "slide_0",
						shapes: [
							{
								id: "sStyled",
								name: "Styled List",
								type: "TextBox",
								text: "One\rTwo\rThree",
								paragraphFormat: { indentLevel: 0, bulletType: "Numbered", bulletStyle: style },
							},
						],
						notes: "",
					},
				],
			});
			mock.install();
			mockBridge.reportResult = mock.mockReportResult;

			const result = (await processCommand(
				"cmd-md-style",
				"powerpoint_get_shape_text_markdown",
				{ slideIndex: 1, shapeId: "sStyled" },
			)) as any;

			expect(result.markdown).toBe(
				[`${markers[0]} One`, `${markers[1]} Two`, `${markers[2]} Three`].join("\n"),
			);
		},
	);

	it("restarts the count when bulletStyle changes at the same indent level", async () => {
		mock = new PowerPointMock({
			slides: [
				{
					id: "slide_0",
					shapes: [
						{
							id: "sSwitch",
							name: "Switching List",
							type: "TextBox",
							text: "One\rTwo\rThree",
							paragraphFormat: {
								indentLevel: 0,
								bulletType: "Numbered",
								bulletStyle: "AlphabetLowercasePeriod",
							},
							paragraphs: [
								undefined,
								undefined,
								{ paragraphFormat: { bulletStyle: "RomanLowercasePeriod" } },
							] as any,
						},
					],
					notes: "",
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-md-switch",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "sSwitch" },
		)) as any;

		expect(result.markdown).toBe(["a. One", "b. Two", "i. Three"].join("\n"));
	});

	it("falls back to '1.'/'2.' for an out-of-scope bulletStyle", async () => {
		mock = new PowerPointMock({
			slides: [
				{
					id: "slide_0",
					shapes: [
						{
							id: "sUnsupported",
							name: "Unsupported Style List",
							type: "TextBox",
							text: "One\rTwo",
							paragraphFormat: {
								indentLevel: 0,
								bulletType: "Numbered",
								bulletStyle: "CircleNumberDoubleBytePlain",
							},
						},
					],
					notes: "",
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-md-fallback",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "sUnsupported" },
		)) as any;

		expect(result.markdown).toBe(["1. One", "2. Two"].join("\n"));
	});

	it("renders a bold run and a separate italic run in the same paragraph", async () => {
		mock = new PowerPointMock({
			slides: [
				{
					id: "slide_0",
					shapes: [
						{
							id: "sInline",
							name: "Inline Formatted",
							type: "TextBox",
							text: "say hello world now",
							paragraphs: [
								{
									runs: [
										{ start: 4, length: 5, font: { bold: true } },
										{ start: 10, length: 5, font: { italic: true } },
									],
								},
							] as any,
						},
					],
					notes: "",
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-md-inline",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "sInline" },
		)) as any;

		expect(result.markdown).toBe("say *hello* _world_ now");
	});
});

describe("bullet marker helper functions", () => {
	it("toAlphabetCounter wraps past 26 (z -> aa -> ab)", () => {
		expect(toAlphabetCounter(1, false)).toBe("a");
		expect(toAlphabetCounter(26, false)).toBe("z");
		expect(toAlphabetCounter(27, false)).toBe("aa");
		expect(toAlphabetCounter(28, false)).toBe("ab");
		expect(toAlphabetCounter(27, true)).toBe("AA");
	});

	it("toRomanNumeral spot checks (4, 9, 19)", () => {
		expect(toRomanNumeral(4, false)).toBe("iv");
		expect(toRomanNumeral(9, false)).toBe("ix");
		expect(toRomanNumeral(19, false)).toBe("xix");
		expect(toRomanNumeral(19, true)).toBe("XIX");
	});
});

describe("groupShapesIntoRows", () => {
	const slideHeight = 540;

	it("groups shapes with exactly equal top into one row", () => {
		const shapes = [{ top: 100, left: 50 }, { top: 100, left: 10 }];
		expect(groupShapesIntoRows(shapes, slideHeight)).toEqual([[1, 0]]);
	});

	it("treats slight misalignment within 5% of slideHeight as the same row", () => {
		// 5% of 540 = 27, so tops 100 and 120 (diff 20) are same row
		const shapes = [{ top: 100, left: 50 }, { top: 120, left: 10 }];
		expect(groupShapesIntoRows(shapes, slideHeight)).toEqual([[1, 0]]);
	});

	it("keeps clearly different tops in separate rows", () => {
		// diff 200 >> 27 tolerance
		const shapes = [{ top: 30, left: 0 }, { top: 350, left: 0 }, { top: 500, left: 0 }];
		expect(groupShapesIntoRows(shapes, slideHeight)).toEqual([[0], [1], [2]]);
	});

	it("orders shapes left-to-right within a row", () => {
		const shapes = [
			{ top: 100, left: 300 },
			{ top: 105, left: 10 },
			{ top: 102, left: 150 },
		];
		expect(groupShapesIntoRows(shapes, slideHeight)).toEqual([[1, 2, 0]]);
	});

	it("returns rows top-to-bottom, ordered by first-seen top in each row", () => {
		const shapes = [
			{ top: 350, left: 0 },
			{ top: 30, left: 0 },
			{ top: 500, left: 0 },
		];
		expect(groupShapesIntoRows(shapes, slideHeight)).toEqual([[1], [0], [2]]);
	});

	it("returns an empty array for empty input", () => {
		expect(groupShapesIntoRows([], slideHeight)).toEqual([]);
	});

	it("falls back to a fixed tolerance when slideHeight is unavailable", () => {
		const shapes = [{ top: 100, left: 0 }, { top: 105, left: 0 }, { top: 200, left: 0 }];
		expect(groupShapesIntoRows(shapes, 0)).toEqual([[0, 1], [2]]);
	});
});

describe("isHeadingSize", () => {
	it("31pt is not a heading, 32pt is", () => {
		expect(isHeadingSize(31)).toBe(false);
		expect(isHeadingSize(32)).toBe(true);
	});

	it("does not qualify when fontSize is 0 (PowerPoint's mixed/unresolved-size sentinel)", () => {
		expect(isHeadingSize(0)).toBe(false);
	});
});

describe("isFootnoteShape", () => {
	const slideHeight = 540;

	it("qualifies a small-font shape whose bottom edge is in the last 15% of slide height", () => {
		// last 15% starts at 459; bottom edge here is 470
		expect(isFootnoteShape({ top: 450, height: 20, fontSize: 12 }, slideHeight)).toBe(true);
	});

	it("does not qualify when the bottom edge is just outside the last 15%", () => {
		// bottom edge 458 < 459 threshold
		expect(isFootnoteShape({ top: 438, height: 20, fontSize: 12 }, slideHeight)).toBe(false);
	});

	it("does not qualify when font size exceeds 12pt, even at the bottom", () => {
		expect(isFootnoteShape({ top: 500, height: 20, fontSize: 13 }, slideHeight)).toBe(false);
	});

	it("does not qualify when fontSize is 0 (PowerPoint's mixed-size sentinel)", () => {
		expect(isFootnoteShape({ top: 500, height: 20, fontSize: 0 }, slideHeight)).toBe(false);
	});

	it("does not qualify when slideHeight is unavailable", () => {
		expect(isFootnoteShape({ top: 500, height: 20, fontSize: 12 }, 0)).toBe(false);
	});
});

describe("powerpoint_get_slide_text_markdown", () => {
	it("combines every text-bearing shape in reading order, rendering headings and footnotes", async () => {
		const result = (await processCommand(
			"cmd-sm1",
			"powerpoint_get_slide_text_markdown",
			{ slideIndex: 1 },
		)) as any;

		const blocks = result.markdown.split("\n\n");
		expect(blocks).toEqual([
			"# *Quarterly Results*",
			"Revenue grew 15%",
			["- Alpha", "    - Beta", "- Gamma", "- Delta", "- Epsilon"].join("\n"),
			"[^1] 1",
		]);
	});

	it("skips shapes with no text frame or empty text", async () => {
		const result = (await processCommand(
			"cmd-sm2",
			"powerpoint_get_slide_text_markdown",
			{ slideIndex: 1 },
		)) as any;

		expect(result.markdown).not.toContain("undefined");
		expect(result.markdown.split("\n\n")).toHaveLength(4);
	});

	it("uses '*' when bulletChar is '*'", async () => {
		const result = (await processCommand(
			"cmd-sm3",
			"powerpoint_get_slide_text_markdown",
			{ slideIndex: 1, bulletChar: "*" },
		)) as any;

		expect(result.markdown).toContain("* Alpha");
	});

	it("rejects a bulletChar other than '-' or '*'", async () => {
		const result = (await processCommand(
			"cmd-sm4",
			"powerpoint_get_slide_text_markdown",
			{ slideIndex: 1, bulletChar: "+" },
		)) as any;

		expect(result.error).toContain("bulletChar");
	});

	it("returns error for out-of-range slideIndex", async () => {
		const result = (await processCommand(
			"cmd-sm5",
			"powerpoint_get_slide_text_markdown",
			{ slideIndex: 99 },
		)) as any;

		expect(result.error).toContain("out of range");
	});

	it("does not treat a shape with mixed per-paragraph font sizes (reported as size 0) as a footnote", async () => {
		// No shape-level font.size is set, so the mock's whole-range size is 0 —
		// PowerPoint's own sentinel for "not uniform across the range" — even
		// though every individual paragraph has a real, large font size.
		mock = new PowerPointMock({
			slideWidth: 960,
			slideHeight: 540,
			slides: [
				{
					id: "slide_0",
					shapes: [
						{
							id: "sMixed",
							name: "Mixed Sizes",
							type: "TextBox",
							text: "One\rTwo",
							left: 50,
							top: 480,
							width: 600,
							height: 40,
							paragraphs: [{ font: { size: 28 } }, { font: { size: 24 } }] as any,
						},
					],
					notes: "",
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-sm7",
			"powerpoint_get_slide_text_markdown",
			{ slideIndex: 1 },
		)) as any;

		expect(result.markdown).not.toContain("[^1]");
	});

	it("returns an empty markdown string for a slide with no text-bearing shapes", async () => {
		const result = (await processCommand(
			"cmd-sm6",
			"powerpoint_get_slide_text_markdown",
			{ slideIndex: 3 },
		)) as any;

		expect(result.markdown).toBe("");
	});

	it("renders a placeholder title as a heading using its layout-inherited font size", async () => {
		mock = new PowerPointMock({
			slideWidth: 960,
			slideHeight: 540,
			slides: [
				{
					id: "slide_0",
					layoutId: "layout_title",
					shapes: [
						{
							id: "sTitle",
							name: "Title 1",
							type: "TextBox",
							text: "Third parties: CVEs",
							placeholderType: "Title",
						},
					],
					notes: "",
				},
			],
			slideMasters: [
				{
					id: "master_0",
					name: "Office Theme",
					layouts: [
						{
							id: "layout_title",
							name: "Title and Content",
							shapes: [
								{
									id: "layoutTitle",
									name: "Title Placeholder 1",
									type: "TextBox",
									text: "",
									placeholderType: "Title",
									font: { size: 35 },
								},
							],
						},
					],
				},
			],
		});
		mock.install();
		mockBridge.reportResult = mock.mockReportResult;

		const result = (await processCommand(
			"cmd-sm8",
			"powerpoint_get_slide_text_markdown",
			{ slideIndex: 1 },
		)) as any;

		expect(result.markdown).toBe("# Third parties: CVEs");
	});
});

describe("powerpoint_get_selection", () => {
	it("returns none when nothing selected", async () => {
		const result = (await processCommand(
			"cmd-12",
			"powerpoint_get_selection",
			{},
		)) as any;
		expect(result.type).toBe("none");
	});
});

describe("powerpoint_get_speaker_notes", () => {
	it("returns notes for a specific slide", async () => {
		const result = (await processCommand(
			"cmd-13",
			"powerpoint_get_speaker_notes",
			{ slideIndex: 1 },
		)) as any;
		expect(result.notes).toHaveLength(1);
		expect(result.notes[0].notes).toBe("Talk about revenue growth");
	});

	it("returns empty string for slides without notes", async () => {
		const result = (await processCommand(
			"cmd-14",
			"powerpoint_get_speaker_notes",
			{ slideIndex: 2 },
		)) as any;
		expect(result.notes[0].notes).toBe("");
	});
});

// ── WRITE TOOLS ──────────────────────────────────────────────────

describe("powerpoint_update_shape_text", () => {
	it("updates text on a shape", async () => {
		const result = (await processCommand(
			"cmd-15",
			"powerpoint_update_shape_text",
			{
				slideIndex: 1,
				shapeId: "s1",
				text: "Q4 Results",
			},
		)) as any;

		expect(result.newText).toBe("Q4 Results");
		expect(result.shapeId).toBe("s1");
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand(
			"cmd-16",
			"powerpoint_update_shape_text",
			{
				slideIndex: 1,
				shapeId: "nope",
				text: "test",
			},
		)) as any;

		expect(result.error).toContain("not found");
	});

	it("returns error for non-text shape", async () => {
		const result = (await processCommand(
			"cmd-17",
			"powerpoint_update_shape_text",
			{
				slideIndex: 1,
				shapeId: "s3",
				text: "test", // s3 is an Image
			},
		)) as any;

		expect(result.error).toContain("does not support text");
	});
});

describe("powerpoint_update_shape_properties", () => {
	it("updates position properties", async () => {
		const result = (await processCommand(
			"cmd-18",
			"powerpoint_update_shape_properties",
			{
				slideIndex: 1,
				shapeId: "s1",
				left: 100,
				top: 50,
			},
		)) as any;

		expect(result.updated).toContain("left");
		expect(result.updated).toContain("top");
	});

	it("updates font properties", async () => {
		const result = (await processCommand(
			"cmd-19",
			"powerpoint_update_shape_properties",
			{
				slideIndex: 1,
				shapeId: "s1",
				fontSize: 48,
				bold: true,
			},
		)) as any;

		expect(result.updated).toContain("fontSize");
		expect(result.updated).toContain("bold");
	});

	it("updates font decoration properties", async () => {
		const result = (await processCommand(
			"cmd-19b",
			"powerpoint_update_shape_properties",
			{
				slideIndex: 1,
				shapeId: "s1",
				underline: "Single",
				strikethrough: true,
				allCaps: true,
			},
		)) as any;

		expect(result.updated).toContain("underline");
		expect(result.updated).toContain("strikethrough");
		expect(result.updated).toContain("allCaps");
	});

	it("updates paragraph and bullet properties", async () => {
		const result = (await processCommand(
			"cmd-19c",
			"powerpoint_update_shape_properties",
			{
				slideIndex: 1,
				shapeId: "s1",
				bulletType: "Numbered",
				bulletVisible: true,
				horizontalAlignment: "Center",
				indentLevel: 2,
			},
		)) as any;

		expect(result.updated).toContain("bulletType");
		expect(result.updated).toContain("bulletVisible");
		expect(result.updated).toContain("horizontalAlignment");
		expect(result.updated).toContain("indentLevel");
	});

	it("updates text frame properties", async () => {
		const result = (await processCommand(
			"cmd-19d",
			"powerpoint_update_shape_properties",
			{
				slideIndex: 1,
				shapeId: "s1",
				textMarginTop: 10,
				wordWrap: false,
				verticalAlignment: "Middle",
			},
		)) as any;

		expect(result.updated).toContain("textMarginTop");
		expect(result.updated).toContain("wordWrap");
		expect(result.updated).toContain("verticalAlignment");
	});

	it("does not update paragraph/font-decoration properties when shape has no text frame", async () => {
		const result = (await processCommand(
			"cmd-19e",
			"powerpoint_update_shape_properties",
			{
				slideIndex: 1,
				shapeId: "s3",
				bulletType: "Numbered",
				underline: "Single",
			},
		)) as any;

		expect(result.updated).not.toContain("bulletType");
		expect(result.updated).not.toContain("underline");
	});
});

describe("powerpoint_update_text_range_properties", () => {
	it("updates indentLevel and a font property on only the targeted paragraph", async () => {
		const before = (await processCommand(
			"cmd-utr-setup",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		const p2 = before.paragraphs[2]; // "Gamma", starts in the default group

		const result = (await processCommand(
			"cmd-utr1",
			"powerpoint_update_text_range_properties",
			{
				slideIndex: 1,
				shapeId: "s7",
				start: p2.start,
				length: p2.length,
				expectedText: p2.text,
				indentLevel: 3,
				bold: true,
			},
		)) as any;

		expect(result.updated).toContain("indentLevel");
		expect(result.updated).toContain("bold");

		const after = (await processCommand(
			"cmd-utr-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		// Only paragraph 2's group changed; its siblings (0, 4) stay grouped together.
		expect(after.paragraphs[2].groupId).not.toBe(before.paragraphs[0].groupId);
		expect(after.paragraphs[0].groupId).toBe(after.paragraphs[4].groupId);

		const changedGroup = after.propertyGroups.find(
			(g: any) => g.groupId === after.paragraphs[2].groupId,
		);
		expect(changedGroup.properties).toEqual({
			font: { bold: true },
			paragraphFormat: { indentLevel: 3 },
		});
	});

	it("returns error when expectedText does not match the current text at [start,length)", async () => {
		const result = (await processCommand(
			"cmd-utr2",
			"powerpoint_update_text_range_properties",
			{
				slideIndex: 1,
				shapeId: "s7",
				start: 0,
				length: 5, // "Alpha"
				expectedText: "Wrong",
				bold: true,
			},
		)) as any;

		expect(result.error).toContain("expectedText");
	});

	it("returns error when start/length are out of range for the shape text", async () => {
		const result = (await processCommand(
			"cmd-utr3",
			"powerpoint_update_text_range_properties",
			{
				slideIndex: 1,
				shapeId: "s7",
				start: 0,
				length: 9999,
				bold: true,
			},
		)) as any;

		expect(result.error).toContain("out of range");
	});

	it("returns error for shape with no text frame", async () => {
		const result = (await processCommand(
			"cmd-utr4",
			"powerpoint_update_text_range_properties",
			{
				slideIndex: 1,
				shapeId: "s3", // s3 is an Image
				start: 0,
				length: 1,
				bold: true,
			},
		)) as any;

		expect(result.error).toContain("does not support text");
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand(
			"cmd-utr5",
			"powerpoint_update_text_range_properties",
			{
				slideIndex: 1,
				shapeId: "nope",
				start: 0,
				length: 1,
				bold: true,
			},
		)) as any;

		expect(result.error).toContain("not found");
	});
});

describe("powerpoint_insert_paragraph", () => {
	it("inserts at the start, shifting every existing paragraph down by one", async () => {
		const result = (await processCommand("cmd-ip1", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			text: "Zero",
			position: "start",
		})) as any;

		expect(result.start).toBe(0);
		expect(result.length).toBe(4);
		expect(result.updated).toEqual([]);

		const after = (await processCommand(
			"cmd-ip1-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(after.paragraphs.map((p: any) => p.text)).toEqual([
			"Zero",
			"Alpha",
			"Beta",
			"Gamma",
			"Delta",
			"Epsilon",
		]);
	});

	it("inserts at the end, appending a new last paragraph", async () => {
		const result = (await processCommand("cmd-ip2", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			text: "Zeta",
			position: "end",
		})) as any;

		const after = (await processCommand(
			"cmd-ip2-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(after.paragraphs.map((p: any) => p.text)).toEqual([
			"Alpha",
			"Beta",
			"Gamma",
			"Delta",
			"Epsilon",
			"Zeta",
		]);
		expect(after.paragraphs[5].start).toBe(result.start);
		expect(after.paragraphs[5].length).toBe(result.length);
	});

	it("inserts before/after a reference paragraph without disturbing its neighbors", async () => {
		const before = (await processCommand(
			"cmd-ip3-setup",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		const gamma = before.paragraphs[2]; // "Gamma"

		await processCommand("cmd-ip3", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			text: "Middle",
			position: "before",
			refParagraphStart: gamma.start,
			refParagraphLength: gamma.length,
			refParagraphText: gamma.text,
		});

		const after = (await processCommand(
			"cmd-ip3-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(after.paragraphs.map((p: any) => p.text)).toEqual([
			"Alpha",
			"Beta",
			"Middle",
			"Gamma",
			"Delta",
			"Epsilon",
		]);
		// Neighboring paragraphs' formatting is untouched — "Beta" (index 1)
		// still shares its group with "Delta" wasn't moved, and "Gamma"
		// (now index 3) still carries the default formatting it had before.
		expect(after.paragraphs[3].groupId).toBe(after.paragraphs[0].groupId);
	});

	it("applies inline formatting to the newly inserted paragraph only", async () => {
		const result = (await processCommand("cmd-ip4", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			text: "Styled",
			position: "end",
			indentLevel: 3,
			color: "#00FF00",
		})) as any;

		expect(result.updated).toContain("indentLevel");
		expect(result.updated).toContain("color");

		const after = (await processCommand(
			"cmd-ip4-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		const newParagraph = after.paragraphs[after.paragraphs.length - 1];
		expect(newParagraph.text).toBe("Styled");
		const group = after.propertyGroups.find(
			(g: any) => g.groupId === newParagraph.groupId,
		);
		expect(group.properties).toEqual({
			font: { color: "#00FF00" },
			paragraphFormat: { indentLevel: 3 },
		});
	});

	it("inherits the spliced-into paragraph's formatting when no formatting params are given", async () => {
		const before = (await processCommand(
			"cmd-ip5-setup",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		const delta = before.paragraphs[3]; // "Delta", has a distinct font color override

		const result = (await processCommand("cmd-ip5", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			text: "AfterDelta",
			position: "after",
			refParagraphStart: delta.start,
			refParagraphLength: delta.length,
		})) as any;

		expect(result.updated).toEqual([]);

		const after = (await processCommand(
			"cmd-ip5-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		const newParagraph = after.paragraphs[4];
		expect(newParagraph.text).toBe("AfterDelta");
		expect(newParagraph.groupId).toBe(after.paragraphs[3].groupId);
	});

	it("returns error when refParagraphStart/refParagraphLength are out of range", async () => {
		const result = (await processCommand("cmd-ip6", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			text: "X",
			position: "after",
			refParagraphStart: 0,
			refParagraphLength: 9999,
		})) as any;

		expect(result.error).toContain("out of range");
	});

	it("returns error when refParagraphText does not match the current text at the reference range", async () => {
		const result = (await processCommand("cmd-ip7", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			text: "X",
			position: "before",
			refParagraphStart: 0,
			refParagraphLength: 5,
			refParagraphText: "Wrong",
		})) as any;

		expect(result.error).toContain("refParagraphText");
	});

	it("returns error for shape with no text frame", async () => {
		const result = (await processCommand("cmd-ip8", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s3", // s3 is an Image
			text: "X",
			position: "end",
		})) as any;

		expect(result.error).toContain("does not support text");
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand("cmd-ip9", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "nope",
			text: "X",
			position: "end",
		})) as any;

		expect(result.error).toContain("not found");
	});

	it("returns error for invalid position", async () => {
		const result = (await processCommand("cmd-ip10", "powerpoint_insert_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			text: "X",
			position: "middle",
		})) as any;

		expect(result.error).toContain("Invalid position");
	});
});

describe("powerpoint_delete_paragraph", () => {
	it("deletes a middle paragraph, shifting later paragraphs up", async () => {
		const before = (await processCommand(
			"cmd-dp1-setup",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		const beta = before.paragraphs[1]; // "Beta", has a distinct indentLevel override

		const result = (await processCommand("cmd-dp1", "powerpoint_delete_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			paragraphStart: beta.start,
			paragraphLength: beta.length,
			expectedText: beta.text,
		})) as any;

		expect(result.deleted).toBe(true);

		const after = (await processCommand(
			"cmd-dp1-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(after.paragraphs.map((p: any) => p.text)).toEqual([
			"Alpha",
			"Gamma",
			"Delta",
			"Epsilon",
		]);
		// "Delta" (now index 2) keeps its distinct font-color override —
		// deleting Beta's slot must not have shifted the wrong paragraph's
		// formatting.
		const delta = after.paragraphs[2];
		const group = after.propertyGroups.find((g: any) => g.groupId === delta.groupId);
		expect(group.properties).toEqual({ font: { color: "#FF0000" } });
	});

	it("deletes the first paragraph", async () => {
		const before = (await processCommand(
			"cmd-dp2-setup",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		const alpha = before.paragraphs[0];

		await processCommand("cmd-dp2", "powerpoint_delete_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			paragraphStart: alpha.start,
			paragraphLength: alpha.length,
		});

		const after = (await processCommand(
			"cmd-dp2-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(after.paragraphs.map((p: any) => p.text)).toEqual([
			"Beta",
			"Gamma",
			"Delta",
			"Epsilon",
		]);
	});

	it("deletes the last paragraph", async () => {
		const before = (await processCommand(
			"cmd-dp3-setup",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		const epsilon = before.paragraphs[4];

		await processCommand("cmd-dp3", "powerpoint_delete_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			paragraphStart: epsilon.start,
			paragraphLength: epsilon.length,
		});

		const after = (await processCommand(
			"cmd-dp3-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(after.paragraphs.map((p: any) => p.text)).toEqual([
			"Alpha",
			"Beta",
			"Gamma",
			"Delta",
		]);
	});

	it("empties the text when deleting the only remaining paragraph", async () => {
		await processCommand("cmd-dp4a", "powerpoint_update_shape_text", {
			slideIndex: 1,
			shapeId: "s7",
			text: "Solo",
		});

		await processCommand("cmd-dp4", "powerpoint_delete_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			paragraphStart: 0,
			paragraphLength: 4,
		});

		const after = (await processCommand(
			"cmd-dp4-verify",
			"powerpoint_get_shape_paragraphs",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;

		expect(after.paragraphs.map((p: any) => p.text)).toEqual([""]);
	});

	it("returns error when start/length are out of range", async () => {
		const result = (await processCommand("cmd-dp5", "powerpoint_delete_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			paragraphStart: 0,
			paragraphLength: 9999,
		})) as any;

		expect(result.error).toContain("out of range");
	});

	it("returns error when expectedText does not match the current text at the range", async () => {
		const result = (await processCommand("cmd-dp6", "powerpoint_delete_paragraph", {
			slideIndex: 1,
			shapeId: "s7",
			paragraphStart: 0,
			paragraphLength: 5,
			expectedText: "Wrong",
		})) as any;

		expect(result.error).toContain("expectedText");
	});

	it("returns error for shape with no text frame", async () => {
		const result = (await processCommand("cmd-dp7", "powerpoint_delete_paragraph", {
			slideIndex: 1,
			shapeId: "s3", // s3 is an Image
			paragraphStart: 0,
			paragraphLength: 1,
		})) as any;

		expect(result.error).toContain("does not support text");
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand("cmd-dp8", "powerpoint_delete_paragraph", {
			slideIndex: 1,
			shapeId: "nope",
			paragraphStart: 0,
			paragraphLength: 1,
		})) as any;

		expect(result.error).toContain("not found");
	});
});

describe("powerpoint_set_shape_text_markdown", () => {
	it.each(BULLET_STYLE_MARKDOWN_CASES)(
		"round-trips $style through set then get",
		async ({ markers }) => {
			const markdown = [`${markers[0]} One`, `${markers[1]} Two`, `${markers[2]} Three`].join(
				"\n",
			);

			const setResult = (await processCommand(
				"cmd-smd-set",
				"powerpoint_set_shape_text_markdown",
				{ slideIndex: 1, shapeId: "s7", markdown },
			)) as any;
			expect(setResult.paragraphCount).toBe(3);

			const getResult = (await processCommand(
				"cmd-smd-get",
				"powerpoint_get_shape_text_markdown",
				{ slideIndex: 1, shapeId: "s7" },
			)) as any;
			expect(getResult.markdown).toBe(markdown);
		},
	);

	it("sets Unnumbered bullets for '-' and '*' markers", async () => {
		const markdown = ["- One", "* Two"].join("\n");

		await processCommand("cmd-smd-un-set", "powerpoint_set_shape_text_markdown", {
			slideIndex: 1,
			shapeId: "s7",
			markdown,
		});

		const getResult = (await processCommand(
			"cmd-smd-un-get",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		expect(getResult.markdown).toBe(["- One", "- Two"].join("\n"));
	});

	it("sets indentLevel from 2-space nesting", async () => {
		const markdown = ["1. One", "  a. Two", "    i. Three", "    ii. Four"].join("\n");

		await processCommand("cmd-smd-indent-set", "powerpoint_set_shape_text_markdown", {
			slideIndex: 1,
			shapeId: "s7",
			markdown,
		});

		const getResult = (await processCommand(
			"cmd-smd-indent-get",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		expect(getResult.markdown).toBe(markdown);
	});

	it("applies bold/italic runs to the right words and round-trips through get", async () => {
		const markdown = "- say *hello* _world_ now";

		await processCommand("cmd-smd-inline-set", "powerpoint_set_shape_text_markdown", {
			slideIndex: 1,
			shapeId: "s7",
			markdown,
		});

		const getResult = (await processCommand(
			"cmd-smd-inline-get",
			"powerpoint_get_shape_text_markdown",
			{ slideIndex: 1, shapeId: "s7" },
		)) as any;
		expect(getResult.markdown).toBe(markdown);
	});

	it("returns error for missing shape, without mutating", async () => {
		const result = (await processCommand("cmd-smd-nf", "powerpoint_set_shape_text_markdown", {
			slideIndex: 1,
			shapeId: "nope",
			markdown: "1. One",
		})) as any;

		expect(result.error).toContain("not found");
	});

	it("returns error for shape with no text frame", async () => {
		const result = (await processCommand("cmd-smd-notf", "powerpoint_set_shape_text_markdown", {
			slideIndex: 1,
			shapeId: "s3", // s3 is an Image
			markdown: "1. One",
		})) as any;

		expect(result.error).toContain("does not support text");
	});
});

describe("markdownToParagraphSpecs", () => {
	it("resolves 'i'/'I' as Roman when starting a fresh numbered run", () => {
		const result = markdownToParagraphSpecs(["i. One", "ii. Two"].join("\n")) as any;
		expect(result.error).toBeUndefined();
		expect(result[0].bulletStyle).toBe("RomanLowercasePeriod");
		expect(result[1].bulletStyle).toBe("RomanLowercasePeriod");
	});

	it("resolves 'i'/'I' as alphabetic immediately after 'h'/'H' in the same run", () => {
		const result = markdownToParagraphSpecs(
			["f. One", "g. Two", "h. Three", "i. Four"].join("\n"),
		) as any;
		expect(result.error).toBeUndefined();
		expect(result.map((s: any) => s.bulletStyle)).toEqual([
			"AlphabetLowercasePeriod",
			"AlphabetLowercasePeriod",
			"AlphabetLowercasePeriod",
			"AlphabetLowercasePeriod",
		]);
	});

	it("resolves 'i' as Roman again when the alphabet run is severed by a bullet", () => {
		const result = markdownToParagraphSpecs(
			["h. One", "- Two", "i. Three", "ii. Four"].join("\n"),
		) as any;
		expect(result.error).toBeUndefined();
		expect(result[0].bulletStyle).toBe("AlphabetLowercasePeriod");
		expect(result[1].bulletType).toBe("Unnumbered");
		expect(result[2].bulletStyle).toBe("RomanLowercasePeriod");
	});

	it("treats 'v'/'x' as always Roman even right after an alphabetic run", () => {
		const result = markdownToParagraphSpecs(["u. One", "v. Two"].join("\n")) as any;
		expect(result.error).toBeUndefined();
		expect(result[0].bulletStyle).toBe("AlphabetLowercasePeriod");
		expect(result[1].bulletStyle).toBe("RomanLowercasePeriod");
	});

	it("treats 'l'/'c'/'d'/'m' as always alphabetic even right after a Roman run", () => {
		const result = markdownToParagraphSpecs(["x. One", "l. Two"].join("\n")) as any;
		expect(result.error).toBeUndefined();
		expect(result[0].bulletStyle).toBe("RomanLowercasePeriod");
		expect(result[1].bulletStyle).toBe("AlphabetLowercasePeriod");
	});

	it("classifies multi-letter Roman markers unambiguously regardless of context", () => {
		const result = markdownToParagraphSpecs(
			["i. One", "ii. Two", "iii. Three", "iv. Four"].join("\n"),
		) as any;
		expect(result.error).toBeUndefined();
		expect(result.map((s: any) => s.bulletStyle)).toEqual([
			"RomanLowercasePeriod",
			"RomanLowercasePeriod",
			"RomanLowercasePeriod",
			"RomanLowercasePeriod",
		]);
	});

	it("falls back to plain text for an unrecognized marker (bare multi-letter, no wrapper)", () => {
		const result = markdownToParagraphSpecs("abc One") as any;
		expect(result.error).toBeUndefined();
		expect(result[0]).toMatchObject({ text: "abc One", indentLevel: 0, bulletType: "None" });
	});

	it("falls back to plain text for a mixed-case multi-letter Roman token", () => {
		const result = markdownToParagraphSpecs("iV. One") as any;
		expect(result.error).toBeUndefined();
		expect(result[0]).toMatchObject({ text: "iV. One", indentLevel: 0, bulletType: "None" });
	});

	it("falls back to plain text for a line with no marker/space separator", () => {
		const result = markdownToParagraphSpecs("NoMarkerHere") as any;
		expect(result.error).toBeUndefined();
		expect(result[0]).toMatchObject({ text: "NoMarkerHere", indentLevel: 0, bulletType: "None" });
	});

	it("round-trips prose → list → prose", () => {
		const result = markdownToParagraphSpecs(
			["Intro paragraph.", "1. First item", "2. Second item", "Closing paragraph."].join("\n"),
		) as any;
		expect(result.error).toBeUndefined();
		expect(result.map((s: any) => s.bulletType)).toEqual(["None", "Numbered", "Numbered", "None"]);
		expect(result[0].text).toBe("Intro paragraph.");
		expect(result[3].text).toBe("Closing paragraph.");
	});

	it("treats a standalone bare arabic '1' as a real marker with no corroboration needed", () => {
		const result = markdownToParagraphSpecs("1. Solo item") as any;
		expect(result.error).toBeUndefined();
		expect(result[0]).toMatchObject({ bulletType: "Numbered", bulletStyle: "ArabicNumeralPeriod" });
	});

	it("falls back to plain text for a standalone non-1 arabic marker", () => {
		const result = markdownToParagraphSpecs("2. Solo item") as any;
		expect(result.error).toBeUndefined();
		expect(result[0]).toMatchObject({ text: "2. Solo item", bulletType: "None" });
	});

	it("corroborates a two-line alpha run via mutual lookahead", () => {
		const result = markdownToParagraphSpecs(["a. First", "b. Second"].join("\n")) as any;
		expect(result.error).toBeUndefined();
		expect(result.map((s: any) => s.bulletType)).toEqual(["Numbered", "Numbered"]);
		expect(result[0].bulletStyle).toBe("AlphabetLowercasePeriod");
		expect(result[1].bulletStyle).toBe("AlphabetLowercasePeriod");
	});

	it("falls back to plain text for an uncorroborated ambiguous single letter", () => {
		const result = markdownToParagraphSpecs("i. do stuff") as any;
		expect(result.error).toBeUndefined();
		expect(result[0]).toMatchObject({ text: "i. do stuff", bulletType: "None" });
	});

	it("errors on a malformed marker inside an established numbered run", () => {
		const result = markdownToParagraphSpecs(["a. First", "b. Second", "zz. Oops"].join("\n")) as any;
		expect(result.error).toContain("zz.");
	});

	it("corroborates via lookahead even when the next line uses a different marker style", () => {
		const result = markdownToParagraphSpecs(["i. Something", "(a) Next"].join("\n")) as any;
		expect(result.error).toBeUndefined();
		expect(result[0].bulletType).toBe("Numbered");
		expect(result[0].bulletStyle).toBe("RomanLowercasePeriod");
	});
});

describe("wrapInlineMarkers", () => {
	it("wraps a single bold word", () => {
		expect(wrapInlineMarkers("hello world", [{ start: 0, length: 5, bold: true, italic: false, strikethrough: false }])).toBe(
			"*hello* world",
		);
	});

	it("wraps a merged multi-word bold run as one span", () => {
		expect(
			wrapInlineMarkers("say hello world now", [
				{ start: 4, length: 11, bold: true, italic: false, strikethrough: false },
			]),
		).toBe("say *hello world* now");
	});

	it("nests bold+italic as *_text_*", () => {
		expect(wrapInlineMarkers("hello", [{ start: 0, length: 5, bold: true, italic: true, strikethrough: false }])).toBe(
			"*_hello_*",
		);
	});

	it("nests bold+italic+strikethrough as *_~text~_*", () => {
		expect(
			wrapInlineMarkers("hello", [{ start: 0, length: 5, bold: true, italic: true, strikethrough: true }]),
		).toBe("*_~hello~_*");
	});

	it("escapes literal marker characters in plain text", () => {
		expect(wrapInlineMarkers("a*b_c~d", [])).toBe("a\\*b\\_c\\~d");
	});

	it("escapes literal marker characters inside a formatted run", () => {
		expect(wrapInlineMarkers("a*b", [{ start: 0, length: 3, bold: true, italic: false, strikethrough: false }])).toBe(
			"*a\\*b*",
		);
	});
});

describe("parseInlineMarkers", () => {
	it("parses a single bold word", () => {
		const result = parseInlineMarkers("*hello* world");
		expect(result.text).toBe("hello world");
		expect(result.runs).toEqual([{ start: 0, length: 5, bold: true, italic: false, strikethrough: false }]);
	});

	it("parses a merged bold run spanning multiple words", () => {
		const result = parseInlineMarkers("say *hello world* now");
		expect(result.text).toBe("say hello world now");
		expect(result.runs).toEqual([{ start: 4, length: 11, bold: true, italic: false, strikethrough: false }]);
	});

	it("parses nested bold+italic", () => {
		const result = parseInlineMarkers("*_hello_*");
		expect(result.text).toBe("hello");
		expect(result.runs).toEqual([{ start: 0, length: 5, bold: true, italic: true, strikethrough: false }]);
	});

	it("parses full triple nesting bold+italic+strikethrough", () => {
		const result = parseInlineMarkers("*_~hello~_*");
		expect(result.text).toBe("hello");
		expect(result.runs).toEqual([{ start: 0, length: 5, bold: true, italic: true, strikethrough: true }]);
	});

	it("unescapes literal marker characters", () => {
		const result = parseInlineMarkers("a\\*b\\_c\\~d");
		expect(result.text).toBe("a*b_c~d");
		expect(result.runs).toEqual([]);
	});

	it("parses a lone marker whose content contains a different marker char literally", () => {
		// "_*text*_" opens with "_" (italic); "*" doesn't extend the open
		// sequence since its priority isn't higher, so only "_" is tried as
		// an opener. Its matching close is the trailing "_", and the "*"
		// characters in between are literal (not re-scanned as bold markers).
		const result = parseInlineMarkers("_*text*_");
		expect(result.text).toBe("*text*");
		expect(result.runs).toEqual([{ start: 0, length: 6, bold: false, italic: true, strikethrough: false }]);
	});

	it("treats an adjacent marker combo with no valid close as literal characters", () => {
		const result = parseInlineMarkers("*_text~");
		expect(result.text).toBe("*_text~");
		expect(result.runs).toEqual([]);
	});

	it("treats an unclosed marker as a literal character", () => {
		const result = parseInlineMarkers("*no close here");
		expect(result.text).toBe("*no close here");
		expect(result.runs).toEqual([]);
	});

	it("round-trips through wrapInlineMarkers", () => {
		const text = "say hello world now";
		const runs = [{ start: 4, length: 11, bold: true, italic: true, strikethrough: false }];
		const wrapped = wrapInlineMarkers(text, runs);
		const parsed = parseInlineMarkers(wrapped);
		expect(parsed.text).toBe(text);
		expect(parsed.runs).toEqual(runs);
	});

	it("round-trips escaped literals through wrapInlineMarkers", () => {
		const text = "a*b_c~d plain";
		const wrapped = wrapInlineMarkers(text, []);
		const parsed = parseInlineMarkers(wrapped);
		expect(parsed.text).toBe(text);
		expect(parsed.runs).toEqual([]);
	});
});

describe("powerpoint_update_speaker_notes", () => {
	it("writes notes to a slide", async () => {
		const result = (await processCommand(
			"cmd-20",
			"powerpoint_update_speaker_notes",
			{
				slideIndex: 1,
				notes: "New speaker notes",
			},
		)) as any;

		expect(result.newNotes).toBe("New speaker notes");
	});
});

// ── SHAPE CRUD ───────────────────────────────────────────────────

describe("powerpoint_add_textbox", () => {
	it("creates a new text box", async () => {
		const result = (await processCommand("cmd-21", "powerpoint_add_textbox", {
			slideIndex: 1,
			text: "Hello!",
			left: 100,
			top: 200,
			width: 300,
			height: 50,
		})) as any;

		expect(result.slideIndex).toBe(1);
		expect(result.name).toBeTruthy();
		expect(mock.data.slides[0].shapes).toHaveLength(6); // was 5
	});
});

describe("powerpoint_add_image", () => {
	it("creates a new image shape", async () => {
		const result = (await processCommand("cmd-22", "powerpoint_add_image", {
			slideIndex: 1,
			imageBase64: "data:image/png;base64,abc123",
			left: 100,
			top: 200,
		})) as any;

		expect(result.slideIndex).toBe(1);
		expect(result.name).toBeTruthy();
	});
});

describe("powerpoint_add_table", () => {
	it("creates a new table shape", async () => {
		const result = (await processCommand("cmd-23", "powerpoint_add_table", {
			slideIndex: 1,
			rows: 3,
			columns: 4,
			left: 100,
			top: 200,
		})) as any;

		expect(result.slideIndex).toBe(1);
		expect(result.name).toBeTruthy();
	});
});

describe("powerpoint_delete_shape", () => {
	it("removes a shape from the slide", async () => {
		const result = (await processCommand("cmd-24", "powerpoint_delete_shape", {
			slideIndex: 1,
			shapeId: "s2",
		})) as any;

		expect(result.deleted).toBe(true);
	});
});

// ── SLIDE MANAGEMENT ─────────────────────────────────────────────

describe("powerpoint_add_slide", () => {
	it("adds a new slide at the end", async () => {
		const result = (await processCommand(
			"cmd-25",
			"powerpoint_add_slide",
			{},
		)) as any;

		expect(result.slideIndex).toBe(4);
		expect(mock.data.slides).toHaveLength(4); // was 3
	});

	it("adds a new slide at specific index", async () => {
		const result = (await processCommand("cmd-26", "powerpoint_add_slide", {
			atIndex: 2,
		})) as any;

		expect(result.slideIndex).toBe(2);
		expect(mock.data.slides).toHaveLength(4);
	});

	it("passes layoutId through to the new slide", async () => {
		const result = (await processCommand("cmd-26b", "powerpoint_add_slide", {
			layoutId: "layout_1",
		})) as any;

		expect(result.slideIndex).toBe(4);
		expect(mock.data.slides[3].layoutId).toBe("layout_1");
	});

	it("passes slideMasterId through alongside layoutId", async () => {
		await processCommand("cmd-26c", "powerpoint_add_slide", {
			layoutId: "layout_2",
			slideMasterId: "master_0",
		});

		expect(mock.data.slides.at(-1)?.layoutId).toBe("layout_2");
	});
});

describe("powerpoint_get_layouts", () => {
	it("returns layouts from the default slide master fixture", async () => {
		const result = (await processCommand(
			"cmd-layouts-1",
			"powerpoint_get_layouts",
			{},
		)) as any;

		expect(result.count).toBe(3);
		expect(result.layouts[0]).toEqual({
			id: "layout_0",
			name: "Title Slide",
			master: "Office Theme",
		});
	});
});

describe("powerpoint_set_slide_layout", () => {
	it("applies a layout found via unscoped search across masters", async () => {
		const result = (await processCommand(
			"cmd-layout-1",
			"powerpoint_set_slide_layout",
			{ slideIndex: 1, layoutId: "layout_1" },
		)) as any;

		expect(result.applied).toBe(true);
		expect(result.slideIndex).toBe(1);
		expect(result.layoutId).toBe("layout_1");
		expect(mock.data.slides[0].layoutId).toBe("layout_1");
	});

	it("applies a layout when slideMasterId is given", async () => {
		const result = (await processCommand(
			"cmd-layout-2",
			"powerpoint_set_slide_layout",
			{ slideIndex: 2, layoutId: "layout_2", slideMasterId: "master_0" },
		)) as any;

		expect(result.applied).toBe(true);
		expect(result.slideMasterId).toBe("master_0");
		expect(mock.data.slides[1].layoutId).toBe("layout_2");
	});

	it("errors when layoutId is not found on any master", async () => {
		const result = (await processCommand(
			"cmd-layout-3",
			"powerpoint_set_slide_layout",
			{ slideIndex: 1, layoutId: "nonexistent_layout" },
		)) as any;

		expect(result.error).toContain("not found");
	});

	it("errors when slideMasterId is not found", async () => {
		const result = (await processCommand(
			"cmd-layout-4",
			"powerpoint_set_slide_layout",
			{ slideIndex: 1, layoutId: "layout_0", slideMasterId: "nonexistent_master" },
		)) as any;

		expect(result.error).toContain("not found");
	});

	it("errors on out-of-range slideIndex", async () => {
		const result = (await processCommand(
			"cmd-layout-5",
			"powerpoint_set_slide_layout",
			{ slideIndex: 99, layoutId: "layout_0" },
		)) as any;

		expect(result.error).toContain("out of range");
	});
});

describe("powerpoint_delete_slide", () => {
	it("deletes a slide", async () => {
		const result = (await processCommand("cmd-27", "powerpoint_delete_slide", {
			slideIndex: 1,
		})) as any;

		expect(result.deleted).toBe(true);
	});
});

describe("powerpoint_move_slide", () => {
	it("moves a slide to new position", async () => {
		const result = (await processCommand("cmd-28", "powerpoint_move_slide", {
			fromIndex: 1,
			toIndex: 3,
		})) as any;

		expect(result.fromIndex).toBe(1);
		expect(result.toIndex).toBe(3);
	});
});

describe("powerpoint_duplicate_slide", () => {
	it("duplicates a slide right after the source by default", async () => {
		const result = (await processCommand(
			"cmd-29",
			"powerpoint_duplicate_slide",
			{ slideIndex: 1 },
		)) as any;

		expect(result.sourceIndex).toBe(1);
		expect(result.newSlideIndex).toBe(2);
		expect(mock.data.slides).toHaveLength(4); // was 3
	});

	it("moves the duplicate to a specific target index", async () => {
		const result = (await processCommand(
			"cmd-30",
			"powerpoint_duplicate_slide",
			{ slideIndex: 1, targetIndex: 4 },
		)) as any;

		expect(result.sourceIndex).toBe(1);
		expect(result.newSlideIndex).toBe(4);
		expect(mock.data.slides).toHaveLength(4);
	});

	it("errors on out-of-range slideIndex", async () => {
		const result = (await processCommand(
			"cmd-31",
			"powerpoint_duplicate_slide",
			{ slideIndex: 99 },
		)) as any;

		expect(result.error).toContain("out of range");
	});

	it("accepts targetIndex equal to slide count (append at end)", async () => {
		// Deck starts with 3 slides; targetIndex == slideCount + 1 (4) means "append
		// after the new last slide" once the duplicate is inserted — must be valid.
		const result = (await processCommand(
			"cmd-32",
			"powerpoint_duplicate_slide",
			{ slideIndex: 1, targetIndex: 4 },
		)) as any;

		expect(result.error).toBeUndefined();
		expect(result.newSlideIndex).toBe(4);
		expect(mock.data.slides).toHaveLength(4);
	});

	it("errors on out-of-range targetIndex without mutating the deck", async () => {
		const result = (await processCommand(
			"cmd-33",
			"powerpoint_duplicate_slide",
			{ slideIndex: 1, targetIndex: 99 },
		)) as any;

		expect(result.error).toContain("out of range");
		// The deck must be untouched — no stray duplicate left behind.
		expect(mock.data.slides).toHaveLength(3);
	});

	it("errors on negative targetIndex without mutating the deck", async () => {
		const result = (await processCommand(
			"cmd-34",
			"powerpoint_duplicate_slide",
			{ slideIndex: 1, targetIndex: -1 },
		)) as any;

		expect(result.error).toContain("out of range");
		expect(mock.data.slides).toHaveLength(3);
	});
});

describe("powerpoint_import_slide_internal", () => {
	it("errors on out-of-range targetIndex without mutating the deck", async () => {
		const result = (await processCommand(
			"cmd-35",
			"powerpoint_import_slide_internal",
			{ base64: "ZmFrZQ==", targetIndex: 99 },
		)) as any;

		expect(result.error).toContain("out of range");
		expect(mock.data.slides).toHaveLength(3);
	});

	it("accepts targetIndex equal to slide count (append at end)", async () => {
		const result = (await processCommand(
			"cmd-36",
			"powerpoint_import_slide_internal",
			{ base64: "ZmFrZQ==", targetIndex: 4 },
		)) as any;

		expect(result.error).toBeUndefined();
		expect(mock.data.slides).toHaveLength(4);
	});
});

// ── ERROR HANDLING ───────────────────────────────────────────────

describe("unknown command", () => {
	it("returns error for unknown command", async () => {
		const result = (await processCommand(
			"cmd-99",
			"nonexistent_command",
			{},
		)) as any;
		expect(result.error).toContain("Unknown command");
	});
});

describe("command dispatch reports success/failure", () => {
	it("reports success for valid read", async () => {
		mock.reset();
		await processCommand("cmd-r1", "powerpoint_get_deck_outline", {});

		expect(mock.reportResultCalls).toHaveLength(1);
		expect(mock.reportResultCalls[0].success).toBe(true);
	});

	it("reports failure for error result", async () => {
		mock.reset();
		await processCommand("cmd-r2", "powerpoint_get_slide", { slideIndex: 99 });

		expect(mock.reportResultCalls).toHaveLength(1);
		expect(mock.reportResultCalls[0].success).toBe(false);
	});
});
