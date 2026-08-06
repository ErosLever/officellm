/**
 * Unit tests for word-commands.ts using the mock framework.
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import { WordMock, type MockDocumentData } from "./word-mock";

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

import { processCommand } from "./word-commands";

// ── Test data ───────────────────────────────────────────────────

function makeTestDoc(): MockDocumentData {
	return {
		paragraphs: [
			{
				text: "Introduction",
				style: "Heading 1",
				outlineLevel: "OutlineLevel1",
			},
			{
				text: "This document outlines our strategy for Q4.",
				style: "Normal",
				outlineLevel: "OutlineLevelBodyText",
			},
			{
				text: "Market Analysis",
				style: "Heading 2",
				outlineLevel: "OutlineLevel2",
			},
			{
				text: "The market grew by 15% year over year.",
				style: "Normal",
				outlineLevel: "OutlineLevelBodyText",
			},
			{
				text: "Revenue grew 20% in the APAC region.",
				style: "Normal",
				outlineLevel: "OutlineLevelBodyText",
			},
			{ text: "Conclusion", style: "Heading 1", outlineLevel: "OutlineLevel1" },
			{
				text: "We recommend proceeding with the expansion plan.",
				style: "Normal",
				outlineLevel: "OutlineLevelBodyText",
			},
		],
		selectedText: "market grew by 15%",
		comments: [],
		changeTrackingMode: "Off",
		changeLog: [],
	};
}

let mock: WordMock;

beforeEach(() => {
	(globalThis as any).window = globalThis;
	mock = new WordMock(makeTestDoc());
	mock.install();
	mockBridge.reportResult = mock.mockReportResult;
});

// ── READ TOOLS ──────────────────────────────────────────────────

describe("word_get_outline", () => {
	it("returns headings up to maxDepth", async () => {
		const result = (await processCommand("c1", "word_get_outline", {
			maxDepth: 2,
		})) as any;

		expect(result.headings).toHaveLength(3); // 2x Heading 1 + 1x Heading 2
		expect(result.headings[0].text).toBe("Introduction");
		expect(result.headings[0].level).toBe(1);
		expect(result.headings[1].text).toBe("Market Analysis");
		expect(result.headings[1].level).toBe(2);
		expect(result.headings[2].text).toBe("Conclusion");
	});

	it("includes all headings when maxDepth is 3", async () => {
		const result = (await processCommand("c2", "word_get_outline", {})) as any;
		expect(result.headings).toHaveLength(3); // 2 Heading 1 + 1 Heading 2
	});

	it("returns totalParagraphs", async () => {
		const result = (await processCommand("c3", "word_get_outline", {})) as any;
		expect(result.totalParagraphs).toBe(7);
	});
});

describe("word_get_paragraphs", () => {
	it("returns paragraphs from startIndex with count", async () => {
		const result = (await processCommand("c4", "word_get_paragraphs", {
			startIndex: 0,
			count: 3,
		})) as any;

		expect(result.totalParagraphs).toBe(7);
		expect(result.paragraphs).toHaveLength(3);
		expect(result.paragraphs[0].text).toBe("Introduction");
		expect(result.paragraphs[0].style).toBe("Heading 1");
	});

	it("returns all paragraphs by default", async () => {
		const result = (await processCommand(
			"c5",
			"word_get_paragraphs",
			{},
		)) as any;
		expect(result.paragraphs).toHaveLength(7);
	});
});

describe("word_get_selection", () => {
	it("returns selected text with paragraph context", async () => {
		const result = (await processCommand(
			"c6",
			"word_get_selection",
			{},
		)) as any;

		expect(result.type).toBe("text");
		expect(result.text).toBe("market grew by 15%");
		expect(result.paragraphs).toBeDefined();
	});
});

describe("word_search", () => {
	it("finds matching text", async () => {
		const result = (await processCommand("c7", "word_search", {
			searchText: "Revenue",
		})) as any;

		expect(result.totalMatches).toBeGreaterThanOrEqual(1);
		expect(result.searchText).toBe("Revenue");
	});

	it("returns 0 for no matches", async () => {
		const result = (await processCommand("c8", "word_search", {
			searchText: "xyzzy",
		})) as any;
		expect(result.totalMatches).toBe(0);
	});
});

// ── WRITE TOOLS ─────────────────────────────────────────────────

describe("word_replace_text", () => {
	it("replaces text in a paragraph", async () => {
		const result = (await processCommand("c9", "word_replace_text", {
			paragraphIndex: 3,
			oldText: "15%",
			newText: "20%",
		})) as any;

		expect(result.replaced).toBe(true);
		// Verify the paragraph text actually changed
		expect(mock.data.paragraphs[3].text).toBe(
			"The market grew by 20% year over year.",
		);
	});

	it("returns error if oldText not found", async () => {
		const result = (await processCommand("c10", "word_replace_text", {
			paragraphIndex: 3,
			oldText: "nonexistent",
			newText: "x",
		})) as any;

		expect(result.error).toContain("not found");
	});

	it("returns error for out-of-range index", async () => {
		const result = (await processCommand("c11", "word_replace_text", {
			paragraphIndex: 99,
			oldText: "x",
			newText: "y",
		})) as any;

		expect(result.error).toContain("out of range");
	});
});

describe("word_insert_text", () => {
	it("inserts at end of document", async () => {
		const result = (await processCommand("c12", "word_insert_text", {
			text: "New final paragraph",
			insertLocation: "end",
		})) as any;

		expect(result.inserted).toBe(true);
		// Verify the paragraph was actually added
		expect(mock.data.paragraphs).toHaveLength(8);
		expect(mock.data.paragraphs[7].text).toBe("New final paragraph");
	});

	it("inserts after a specific paragraph", async () => {
		const result = (await processCommand("c13", "word_insert_text", {
			text: "Inserted after intro",
			insertLocation: "afterParagraph",
			paragraphIndex: 0,
		})) as any;

		expect(result.inserted).toBe(true);
		// Verify the paragraph was inserted at the right position
		expect(mock.data.paragraphs).toHaveLength(8);
		expect(mock.data.paragraphs[1].text).toBe("Inserted after intro");
	});
});

describe("word_add_comment", () => {
	it("adds comment to selection", async () => {
		const result = (await processCommand("c14", "word_add_comment", {
			commentText: "Check this claim",
		})) as any;

		expect(result.added).toBe(true);
		// Verify comment was recorded
		expect(mock.data.comments).toHaveLength(1);
		expect(mock.data.comments[0].text).toBe("Check this claim");
	});

	it("adds comment to a specific paragraph", async () => {
		const result = (await processCommand("c15", "word_add_comment", {
			commentText: "Reword this",
			paragraphIndex: 4,
		})) as any;

		expect(result.added).toBe(true);
		// Verify comment was recorded with paragraph index
		expect(mock.data.comments).toHaveLength(1);
		expect(mock.data.comments[0].text).toBe("Reword this");
		expect(mock.data.comments[0].paragraphIndex).toBe(4);
	});
});

describe("word_delete_paragraph", () => {
	it("deletes a paragraph by index", async () => {
		const result = (await processCommand("c16", "word_delete_paragraph", {
			paragraphIndex: 3,
		})) as any;

		expect(result.deleted).toBe(true);
		// Verify paragraph was actually removed
		expect(mock.data.paragraphs).toHaveLength(6);
		expect(mock.data.paragraphs[3].text).toBe(
			"Revenue grew 20% in the APAC region.",
		);
	});

	it("returns error for out-of-range", async () => {
		const result = (await processCommand("c17", "word_delete_paragraph", {
			paragraphIndex: 99,
		})) as any;

		expect(result.error).toContain("out of range");
	});
});

// ── TRACKED CHANGES ────────────────────────────────────────────

describe("word_replace_text tracked changes", () => {
	it("returns tracked: true for replace", async () => {
		const result = (await processCommand("c20", "word_replace_text", {
			paragraphIndex: 3,
			oldText: "15%",
			newText: "20%",
		})) as any;

		expect(result.replaced).toBe(true);
		expect(result.tracked).toBe(true);
	});
});

describe("word_insert_text tracked changes", () => {
	it("returns tracked: true for insert at end", async () => {
		const result = (await processCommand("c21", "word_insert_text", {
			text: "Tracked insertion",
			insertLocation: "end",
		})) as any;

		expect(result.inserted).toBe(true);
		expect(result.tracked).toBe(true);
	});
});

describe("word_delete_paragraph tracked changes", () => {
	it("returns tracked: true for delete", async () => {
		const result = (await processCommand("c22", "word_delete_paragraph", {
			paragraphIndex: 3,
		})) as any;

		expect(result.deleted).toBe(true);
		expect(result.tracked).toBe(true);
	});
});

describe("word_get_tracked_changes", () => {
	it("returns change tracking mode and pending count", async () => {
		const result = (await processCommand(
			"c23",
			"word_get_tracked_changes",
			{},
		)) as any;

		expect(result.changeTrackingMode).toBeDefined();
		expect(typeof result.pendingChanges).toBe("number");
	});
});

describe("word_accept_all_changes", () => {
	it("accepts all changes and clears changeLog", async () => {
		// First make a change
		await processCommand("c24", "word_replace_text", {
			paragraphIndex: 3,
			oldText: "15%",
			newText: "20%",
		});

		const result = (await processCommand(
			"c25",
			"word_accept_all_changes",
			{},
		)) as any;

		expect(result.accepted).toBe(true);
		expect(mock.data.changeLog).toHaveLength(0);
	});
});

describe("word_reject_all_changes", () => {
	it("rejects all changes and restores original text", async () => {
		// First make a change
		await processCommand("c26", "word_replace_text", {
			paragraphIndex: 3,
			oldText: "15%",
			newText: "20%",
		});

		const result = (await processCommand(
			"c27",
			"word_reject_all_changes",
			{},
		)) as any;

		expect(result.rejected).toBe(true);
		expect(mock.data.changeLog).toHaveLength(0);
		expect(mock.data.paragraphs[3].text).toBe(
			"The market grew by 15% year over year.",
		);
	});
});

// ── ERROR HANDLING ───────────────────────────────────────────────

describe("unknown command", () => {
	it("returns error for unknown Word command", async () => {
		const result = (await processCommand("c99", "word_nonexistent", {})) as any;
		expect(result.error).toContain("Unknown Word command");
	});
});
