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
import { processCommand } from "./powerpoint-commands";

// ── Test data ───────────────────────────────────────────────────

function makeTestDeck(): MockPresentationData {
	return {
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
		expect(result.slides[0].shapes).toHaveLength(4);
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
});

describe("powerpoint_get_slide", () => {
	it("returns full shape properties for a slide", async () => {
		const result = (await processCommand("cmd-4", "powerpoint_get_slide", {
			slideIndex: 0,
		})) as any;

		expect(result.slideIndex).toBe(0);
		expect(result.shapes).toHaveLength(4);

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

	it("includes font properties for text shapes", async () => {
		const result = (await processCommand("cmd-5", "powerpoint_get_slide", {
			slideIndex: 0,
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
});

describe("powerpoint_get_slide_image", () => {
	it("returns base64 PNG image", async () => {
		const result = (await processCommand(
			"cmd-7",
			"powerpoint_get_slide_image",
			{ slideIndex: 0 },
		)) as any;

		expect(result.slideIndex).toBe(0);
		expect(result.image).toMatch(/^data:image\/png;base64,/);
	});
});

describe("powerpoint_get_shape_image", () => {
	it("returns image for a shape by ID", async () => {
		const result = (await processCommand(
			"cmd-8",
			"powerpoint_get_shape_image",
			{ slideIndex: 0, shapeId: "s3" },
		)) as any;

		expect(result.slideIndex).toBe(0);
		expect(result.shapeId).toBe("s3");
		expect(result.image).toMatch(/^data:image\/png;base64,/);
	});

	it("returns error for missing shape", async () => {
		const result = (await processCommand(
			"cmd-9",
			"powerpoint_get_shape_image",
			{ slideIndex: 0, shapeId: "nonexistent" },
		)) as any;
		expect(result.error).toContain("not found");
	});
});

describe("powerpoint_get_table", () => {
	it("returns table cells as 2D array", async () => {
		const result = (await processCommand("cmd-10", "powerpoint_get_table", {
			slideIndex: 1,
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
			slideIndex: 1,
			shapeId: "nope",
		})) as any;
		expect(result.error).toContain("not found");
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
			{ slideIndex: 0 },
		)) as any;
		expect(result.notes).toHaveLength(1);
		expect(result.notes[0].notes).toBe("Talk about revenue growth");
	});

	it("returns empty string for slides without notes", async () => {
		const result = (await processCommand(
			"cmd-14",
			"powerpoint_get_speaker_notes",
			{ slideIndex: 1 },
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
				slideIndex: 0,
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
				slideIndex: 0,
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
				slideIndex: 0,
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
				slideIndex: 0,
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
				slideIndex: 0,
				shapeId: "s1",
				fontSize: 48,
				bold: true,
			},
		)) as any;

		expect(result.updated).toContain("fontSize");
		expect(result.updated).toContain("bold");
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
			slideIndex: 0,
			text: "Hello!",
			left: 100,
			top: 200,
			width: 300,
			height: 50,
		})) as any;

		expect(result.slideIndex).toBe(0);
		expect(result.name).toBeTruthy();
		expect(mock.data.slides[0].shapes).toHaveLength(5); // was 4
	});
});

describe("powerpoint_add_image", () => {
	it("creates a new image shape", async () => {
		const result = (await processCommand("cmd-22", "powerpoint_add_image", {
			slideIndex: 0,
			imageBase64: "data:image/png;base64,abc123",
			left: 100,
			top: 200,
		})) as any;

		expect(result.slideIndex).toBe(0);
		expect(result.name).toBeTruthy();
	});
});

describe("powerpoint_add_table", () => {
	it("creates a new table shape", async () => {
		const result = (await processCommand("cmd-23", "powerpoint_add_table", {
			slideIndex: 0,
			rows: 3,
			columns: 4,
			left: 100,
			top: 200,
		})) as any;

		expect(result.slideIndex).toBe(0);
		expect(result.name).toBeTruthy();
	});
});

describe("powerpoint_delete_shape", () => {
	it("removes a shape from the slide", async () => {
		const result = (await processCommand("cmd-24", "powerpoint_delete_shape", {
			slideIndex: 0,
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

		expect(result.slideIndex).toBe(3);
		expect(mock.data.slides).toHaveLength(4); // was 3
	});

	it("adds a new slide at specific index", async () => {
		const result = (await processCommand("cmd-26", "powerpoint_add_slide", {
			atIndex: 1,
		})) as any;

		expect(result.slideIndex).toBe(1);
		expect(mock.data.slides).toHaveLength(4);
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
			fromIndex: 0,
			toIndex: 2,
		})) as any;

		expect(result.fromIndex).toBe(0);
		expect(result.toIndex).toBe(2);
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
