/**
 * Mock framework for Outlook JS API — mirrors the real Office.context.mailbox
 * callback-based pattern for unit testing outlook-commands.ts.
 *
 * Architecture: Same pattern as excel-mock.ts.
 * - MockMailboxData defines test data
 * - installOutlookMock() patches window.Office with a mock
 * - Mock tracks all mutations for assertions
 */

// ── Data types ──────────────────────────────────────────────────

export interface MockEmailItem {
	itemType: "Message";
	subject: string;
	body: string;
	bodyHtml?: string;
	sender: { name: string; address: string };
	to: Array<{ name: string; address: string }>;
	cc: Array<{ name: string; address: string }>;
	received: string; // ISO date
	attachments?: Array<{ name: string; size: number; type: string }>;
	categories?: string[];
	conversationId?: string;
	itemId?: string;
}

export interface MockMailboxData {
	currentItem: MockEmailItem;
	threadItems?: MockEmailItem[];
	availableCategories?: string[];
}

// ── Mock state ──────────────────────────────────────────────────

let _mailboxData: MockMailboxData = {
	currentItem: {
		itemType: "Message",
		subject: "Test Email",
		body: "This is a test email body.",
		sender: { name: "Alice Smith", address: "alice@example.com" },
		to: [{ name: "Bob Jones", address: "bob@example.com" }],
		cc: [],
		received: "2026-05-16T09:00:00Z",
	},
};

let _mutations: Array<{ type: string; details: unknown }> = [];
const _confirmationTokens: Map<string, string> = new Map();

export function getMutations() {
	return _mutations;
}

export function resetMutations() {
	_mutations = [];
	_confirmationTokens.clear();
}

export function setMailboxData(data: MockMailboxData) {
	_mailboxData = data;
	_mutations = [];
	_confirmationTokens.clear();
}

export function getMailboxData(): MockMailboxData {
	return _mailboxData;
}

export function generateConfirmationToken(messageId: string): string {
	const token = `confirm_${Date.now()}_${Math.random().toString(36).slice(2)}`;
	_confirmationTokens.set(token, messageId);
	return token;
}

// ── Mock Office.context.mailbox ─────────────────────────────────

function createMockMailbox() {
	// Use getters to dynamically reference _mailboxData so changes after install work
	const getItem = () => _mailboxData.currentItem;

	return {
		get item() {
			const item = getItem();
			return {
				get itemType() {
					return item.itemType;
				},
				get subject() {
					return item.subject;
				},
				get itemId() {
					return item.itemId || "item-001";
				},

				body: {
					getAsync(
						coercionType: string,
						options: any,
						callback?: (result: any) => void,
					) {
						const cb = typeof options === "function" ? options : callback;
						const body =
							coercionType === "Html"
								? item.bodyHtml || `<p>${item.body}</p>`
								: item.body;
						cb?.({ status: "succeeded", value: body });
					},
				},

				get sender() {
					return {
						emailAddress: item.sender.address,
						displayName: item.sender.name,
					};
				},

				get to() {
					return item.to.map((r) => ({
						emailAddress: r.address,
						displayName: r.name,
					}));
				},

				get cc() {
					return item.cc.map((r) => ({
						emailAddress: r.address,
						displayName: r.name,
					}));
				},

				get dateTimeCreated() {
					return item.received;
				},
				get dateTimeModified() {
					return item.received;
				},

				get attachments() {
					return (item.attachments || []).map((a) => ({
						name: a.name,
						size: a.size,
						attachmentType: a.type,
					}));
				},

				get categories() {
					return item.categories || [];
				},
				get conversationId() {
					return item.conversationId || "conv-001";
				},

				loadPropertiesForIds(
					_ids: string[],
					_options: any,
					_callback?: (result: any) => void,
				) {
					const cb = typeof _options === "function" ? _options : _callback;
					cb?.({ status: "succeeded", value: null });
				},

				getAllInternetHeadersAsync(
					options: any,
					callback?: (result: any) => void,
				) {
					const cb = typeof options === "function" ? options : callback;
					cb?.({ status: "succeeded", value: {} });
				},
			};
		},

		get itemId() {
			return getItem().itemId || "item-001";
		},
		diagnostics: {
			hostName: "Outlook",
			hostVersion: "16.0",
		},

		userProfile: {
			emailAddress: "user@example.com",
			displayName: "Test User",
		},

		getCallbackTokenAsync(options: any, callback?: (result: any) => void) {
			const cb = typeof options === "function" ? options : callback;
			cb?.({ status: "succeeded", value: "mock-token-123" });
		},
	};
}

// ── Install mock ────────────────────────────────────────────────

export function installOutlookMock() {
	const mockOffice = {
		context: {
			mailbox: createMockMailbox(),
			requirements: {
				isSetSupported: () => true,
			},
		},
		CoercionType: { Text: "Text", Html: "Html" },
		AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
		HostType: { Outlook: "Outlook" },
		MailboxEnums: {
			BodyMode: {
				FullBody: "FullBody",
				HostConfig: "HostConfig",
			},
			ItemType: { Message: "Message", Appointment: "Appointment" },
		},
		onReady: (cb: any) => {
			if (cb) cb({ host: "Outlook", platform: "PC" });
		},
	};

	(globalThis as any).Office = mockOffice;
	(globalThis as any).window = globalThis;
}

export function uninstallOutlookMock() {
	delete (globalThis as any).Office;
}
