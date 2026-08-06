/**
 * Unit tests for outlook-commands.ts using the mock framework.
 */

import { describe, it, expect, beforeEach, vi } from "vitest";
import {
	type MockMailboxData,
	installOutlookMock,
	setMailboxData,
	generateConfirmationToken,
} from "./outlook-mock";
import { getMutations, resetMutations } from "./outlook-commands";
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

import { processCommand } from "./outlook-commands";

// ── Test data ───────────────────────────────────────────────────

function makeTestMailbox(): MockMailboxData {
	return {
		currentItem: {
			itemType: "Message",
			subject: "Q4 Budget Review - Action Required",
			body: "Hi team,\n\nPlease review the attached Q4 budget proposal before Friday's meeting.\n\nKey items:\n- Marketing spend increased 15%\n- New headcount for engineering\n- Office renovation deferred to Q1\n\nLet me know if you have questions.\n\nBest,\nAlice",
			bodyHtml:
				"<p>Hi team,</p><p>Please review the attached Q4 budget proposal before Friday's meeting.</p>",
			sender: { name: "Alice Smith", address: "alice@example.com" },
			to: [
				{ name: "Bob Jones", address: "bob@example.com" },
				{ name: "Carol White", address: "carol@example.com" },
			],
			cc: [{ name: "Dave Brown", address: "dave@example.com" }],
			received: "2026-05-16T09:00:00Z",
			attachments: [
				{ name: "Q4_Budget.xlsx", size: 45000, type: "File" },
				{ name: "Summary.pdf", size: 120000, type: "File" },
			],
			categories: ["Important"],
			conversationId: "conv-budget-001",
			itemId: "item-budget-001",
		},
		availableCategories: ["Important", "Project Alpha", "Follow Up", "Review"],
	};
}

// ── Tests ────────────────────────────────────────────────────────

describe("outlook-commands", () => {
	beforeEach(() => {
		installOutlookMock();
		setMailboxData(makeTestMailbox());
		resetMutations();
	});

	// ── get current item ────────────────────────────────────────

	describe("outlook_get_current_item", () => {
		it("returns email metadata", async () => {
			const result = (await processCommand(
				"cmd-1",
				"outlook_get_current_item",
				{},
			)) as any;

			expect(result.subject).toBe("Q4 Budget Review - Action Required");
			expect(result.sender.name).toBe("Alice Smith");
			expect(result.sender.address).toBe("alice@example.com");
			expect(result.to.length).toBe(2);
			expect(result.cc.length).toBe(1);
			expect(result.itemId).toBe("item-budget-001");
		});

		it("includes body text by default", async () => {
			const result = (await processCommand(
				"cmd-2",
				"outlook_get_current_item",
				{},
			)) as any;

			expect(result.body).toBeDefined();
			expect(result.body.length).toBeGreaterThan(0);
			expect(result.body).toContain("budget proposal");
		});

		it("omits body when includeBody is false", async () => {
			const result = (await processCommand(
				"cmd-3",
				"outlook_get_current_item",
				{
					includeBody: false,
				},
			)) as any;

			expect(result.body).toBeUndefined();
		});

		it("includes attachment metadata", async () => {
			const result = (await processCommand(
				"cmd-4",
				"outlook_get_current_item",
				{},
			)) as any;

			expect(result.attachments).toBeDefined();
			expect(result.attachments.length).toBe(2);
			expect(result.attachments[0].name).toBe("Q4_Budget.xlsx");
			expect(result.attachments[0].size).toBe(45000);
		});

		it("includes categories", async () => {
			const result = (await processCommand(
				"cmd-5",
				"outlook_get_current_item",
				{},
			)) as any;

			expect(result.categories).toContain("Important");
		});
	});

	// ── summarize thread ─────────────────────────────────────────

	describe("outlook_summarize_thread", () => {
		it("returns structured summary", async () => {
			const result = (await processCommand(
				"cmd-6",
				"outlook_summarize_thread",
				{},
			)) as any;

			expect(result.subject).toBe("Q4 Budget Review - Action Required");
			expect(result.conversationId).toBe("conv-budget-001");
			expect(result.timeline).toBeDefined();
			expect(result.timeline.length).toBeGreaterThan(0);
			expect(result.keyDecisions).toBeDefined();
			expect(result.actionItems).toBeDefined();
		});

		it("respects maxMessages parameter", async () => {
			const result = (await processCommand(
				"cmd-7",
				"outlook_summarize_thread",
				{
					maxMessages: 10,
				},
			)) as any;

			expect(result.maxMessages).toBe(10);
		});
	});

	// ── draft reply ──────────────────────────────────────────────

	describe("outlook_draft_reply", () => {
		it("creates a draft reply with tone", async () => {
			const result = (await processCommand("cmd-8", "outlook_draft_reply", {
				tone: "formal",
			})) as any;

			expect(result.draftCreated).toBe(true);
			expect(result.tone).toBe("formal");
			expect(result.neverAutoSent).toBe(true);
			expect(result.subject).toContain("RE:");
		});

		it("accepts key points", async () => {
			const result = (await processCommand("cmd-9", "outlook_draft_reply", {
				tone: "concise",
				keyPoints: ["Approve budget", "Schedule follow-up"],
			})) as any;

			expect(result.keyPoints).toContain("Approve budget");
			expect(result.keyPoints).toContain("Schedule follow-up");
		});

		it("rejects invalid tone", async () => {
			const result = (await processCommand("cmd-10", "outlook_draft_reply", {
				tone: "sarcastic",
			})) as any;

			expect(result.error).toBeDefined();
			expect(result.error).toContain("Invalid tone");
		});

		it("tracks mutation in mock", async () => {
			await processCommand("cmd-11", "outlook_draft_reply", {
				tone: "friendly",
			});

			const mutations = getMutations();
			expect(mutations.length).toBe(1);
			expect(mutations[0].type).toBe("draft_reply");
		});
	});

	// ── apply category ───────────────────────────────────────────

	describe("outlook_apply_category", () => {
		it("applies a category", async () => {
			const result = (await processCommand("cmd-12", "outlook_apply_category", {
				categoryName: "Project Alpha",
			})) as any;

			expect(result.applied).toBe(true);
			expect(result.categoryName).toBe("Project Alpha");
		});

		it("returns error for missing categoryName", async () => {
			const result = (await processCommand(
				"cmd-13",
				"outlook_apply_category",
				{},
			)) as any;

			expect(result.error).toContain("categoryName");
		});

		it("tracks mutation in mock", async () => {
			await processCommand("cmd-14", "outlook_apply_category", {
				categoryName: "Follow Up",
			});

			const mutations = getMutations();
			expect(mutations.length).toBe(1);
			expect(mutations[0].type).toBe("apply_category");
		});
	});

	// ── send message ─────────────────────────────────────────────

	describe("outlook_send_message", () => {
		it("rejects send without confirmation token", async () => {
			const result = (await processCommand(
				"cmd-15",
				"outlook_send_message",
				{},
			)) as any;

			expect(result.error).toBeDefined();
			expect(result.errorCode).toBe("CONFIRMATION_REQUIRED");
			expect(result.requiresConfirmation).toBe(true);
		});

		it("rejects invalid confirmation token", async () => {
			const result = (await processCommand("cmd-16", "outlook_send_message", {
				confirmationToken: "bad_token",
			})) as any;

			expect(result.errorCode).toBe("INVALID_TOKEN");
		});

		it("sends with valid confirmation token", async () => {
			const token = generateConfirmationToken("draft-001");
			const result = (await processCommand("cmd-17", "outlook_send_message", {
				confirmationToken: token,
				messageId: "draft-001",
			})) as any;

			expect(result.sent).toBe(true);
			expect(result.messageId).toBe("draft-001");
		});

		it("tracks mutation in mock", async () => {
			const token = generateConfirmationToken("draft-001");
			await processCommand("cmd-18", "outlook_send_message", {
				confirmationToken: token,
			});

			const mutations = getMutations();
			expect(mutations.length).toBe(1);
			expect(mutations[0].type).toBe("send_message");
		});
	});

	// ── unknown command ─────────────────────────────────────────

	describe("unknown command", () => {
		it("returns error for unknown Outlook command", async () => {
			const result = (await processCommand(
				"cmd-19",
				"outlook_nonexistent",
				{},
			)) as any;

			expect(result.error).toContain("Unknown Outlook command");
		});
	});
});
