/**
 * Main entry point for the Office LLM Harness PowerPoint Add-in.
 * Registers with the MCP server, polls for commands, and manages the task pane UI.
 */

/// <reference types="@types/office-js" />

import {
	registerWithMcp,
	startHeartbeat,
	pollForCommands,
	reportResult,
	getOfficeState,
	connectSignalR,
	setCommandHandler,
	onConnectionStateChange,
} from "./communication";
import { processCommand as processPptCommand } from "./powerpoint-commands";
import { processCommand as processWordCommand } from "./word-commands";
import { processCommand as processExcelCommand } from "./excel-commands";
import { processCommand as processOutlookCommand } from "./outlook-commands";

// --- Host-aware command dispatch ---
const HOST_DISPATCH: Record<string, typeof processPptCommand> = {
	powerpoint_: processPptCommand,
	word_: processWordCommand,
	excel_: processExcelCommand,
	outlook_: processOutlookCommand,
};

function processCommand(
	commandId: string,
	commandName: string,
	args: unknown,
): Promise<unknown> {
	// Shared tools handled before host dispatch
	if (commandName === "office_export_document") {
		return handleExportDocument(args);
	}

	for (const [prefix, handler] of Object.entries(HOST_DISPATCH)) {
		if (commandName.startsWith(prefix)) {
			return handler(commandId, commandName, args);
		}
	}
	return Promise.resolve({ error: `Unknown host for command: ${commandName}` });
}

// --- State ---
let instanceId: string | null = null;

// --- Initialization ---
let isInitialized = false;

Office.onReady((info) => {
	console.log(`Office ready: ${info.host}`);
	updateOfficeStatus(true);

	// Register with MCP server (only once)
	if (!isInitialized) {
		isInitialized = true;
		initWithMcp();
	}
});

function updateMcpStatus(connected: boolean, text?: string): void {
	const el = document.getElementById("mcpStatus");
	if (!el) return;
	el.className = connected ? "badge connected" : "badge disconnected";
	el.textContent = text ?? (connected ? "Connected" : "Disconnected");
}

async function initWithMcp(): Promise<void> {
	updateMcpStatus(false, "Connecting...");
	try {
		const state = await getOfficeState();
		instanceId = await registerWithMcp(state.app, state.documentName);
		updateMcpStatus(true);
		updateContextDisplay(state);

		// Start heartbeat
		startHeartbeat(10000);

		// Set up command handler for SignalR
		setCommandHandler(async (commandId, commandName, args) => {
			addLogEntry(`[SignalR] Executing: ${commandName}`);
			const result = await processCommand(commandId, commandName, args);
			addLogEntry(`[SignalR] Command ${commandId} completed`);
			return result;
		});

		// Monitor connection state
		onConnectionStateChange((state) => {
			if (state === "connected") {
				updateMcpStatus(true, "Connected (SignalR)");
			} else if (state === "reconnecting") {
				updateMcpStatus(false, "Reconnecting...");
			} else {
				updateMcpStatus(true, "Connected (HTTP polling)");
			}
		});

		// Connect SignalR (falls back to HTTP polling automatically)
		await connectSignalR();

		// Always start HTTP polling as fallback (commands may arrive while reconnecting)
		setInterval(() => processPendingCommands(), 2000);

		console.log(`Add-in registered as: ${instanceId}`);
		addLogEntry(`Registered as ${instanceId}`);

		// Process any already-queued commands
		await processPendingCommands();
	} catch (error) {
		console.error("Registration failed:", error);
		updateMcpStatus(false, "Failed");
		addLogEntry(
			`Registration failed: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

// --- Diff Preview & Confirmation (removed) ---

// --- UI Updates ---

function updateOfficeStatus(ready: boolean): void {
	const el = document.getElementById("officeStatus");
	if (!el) return;
	el.className = ready ? "badge connected" : "badge disconnected";
	el.textContent = ready ? "Ready" : "Not Ready";
}

function updateContextDisplay(state: import("./communication").OfficeState): void {
	const section = document.getElementById("contextSection");
	const rows = document.getElementById("contextRows");
	const subtitle = document.getElementById("hostSubtitle");
	if (!section || !rows) return;

	section.style.display = "block";

	const host = state.app.toLowerCase();

	if (subtitle) {
		subtitle.textContent = `${state.app} · MCP Bridge`;
	}

	const items: Array<[string, string]> = [
		[host === "word" ? "Document" : host === "excel" ? "Workbook" : "Presentation",
		 state.documentName || "Untitled"],
	];

	if (host === "word") {
		items.push(["Paragraphs", String(state.paragraphCount)]);
		items.push(["Words", String(state.wordCount)]);
	} else if (host === "excel") {
		items.push(["Active Sheet", state.activeSheetName || "—"]);
		items.push(["Sheets", String(state.sheetCount)]);
	} else {
		items.push(["Slide", state.currentSlideIndex >= 0 ? `Slide ${state.currentSlideIndex + 1}` : "—"]);
		items.push(["Slides", `${state.slideCount} slide${state.slideCount !== 1 ? "s" : ""}`]);
	}

	rows.innerHTML = items
		.map(([label, value]) =>
			`<div class="status-row">
				<span class="status-label">${label}</span>
				<span class="status-value">${value}</span>
			</div>`)
		.join("");
}

function addLogEntry(message: string): void {
	const log = document.getElementById("activityLog");
	if (!log) return;

	const loading = log.querySelector(".loading");
	if (loading) loading.remove();

	const entry = document.createElement("div");
	entry.className = "log-entry";
	const time = new Date().toLocaleTimeString();
	entry.textContent = `[${time}] ${message}`;

	log.insertBefore(entry, log.firstChild);

	// Keep only last 20 entries
	while (log.children.length > 20) {
		log.removeChild(log.lastChild!);
	}
}

// --- Command Processing ---

let isProcessingCommands = false;

async function processPendingCommands(): Promise<void> {
	if (!instanceId || isProcessingCommands) return;

	isProcessingCommands = true;
	try {
		const commands = await pollForCommands();
		for (const cmd of commands) {
			addLogEntry(`Executing: ${cmd.command}`);
			try {
				const result = await processCommand(cmd.id, cmd.command, cmd.args);
				const success = !(result && typeof result === "object" && "error" in (result as any));
				const error = success ? undefined : ((result as any).error as string);
				await reportResult(cmd.id, success, error, result);
				addLogEntry(`Command ${cmd.id} completed`);
			} catch (err) {
				const errMsg = err instanceof Error ? err.message : String(err);
				await reportResult(cmd.id, false, errMsg, null);
				addLogEntry(`Command ${cmd.id} failed: ${errMsg}`);
			}
		}
	} catch (error) {
		addLogEntry(
			`Command poll error: ${error instanceof Error ? error.message : String(error)}`,
		);
	} finally {
		isProcessingCommands = false;
	}
}

// --- Actions ---

async function refreshContext(): Promise<void> {
	addLogEntry("Refreshing context...");

	try {
		const state = await getOfficeState();
		updateContextDisplay(state);
		addLogEntry(`Context refreshed: ${state.documentName}`);

		// Re-register with updated info
		if (instanceId) {
			instanceId = await registerWithMcp(state.app, state.documentName);
			addLogEntry(`Re-registered as ${instanceId}`);
		}
	} catch (error) {
		addLogEntry(
			`Error: ${error instanceof Error ? error.message : String(error)}`,
		);
	}
}

// --- Actions ---

// --- Expose functions for HTML onclick handlers ---
declare global {
	interface Window {
		refreshContext: () => Promise<void>;
	}
}

window.refreshContext = refreshContext;

// --- Start ---

console.log("Office LLM Harness PowerPoint Add-in loaded");
console.log("Office LLM Harness PowerPoint Add-in loaded");

// ── Document Export (Phase 13) ──────────────────────────────────

async function handleExportDocument(args: unknown): Promise<unknown> {
	const config = args as { format?: string; maxSizeMB?: number };
	const { format = "pdf", maxSizeMB = 50 } = config;

	const Office: any = (window as any).Office;
	if (!Office?.context?.document) {
		return {
			error: "Document context not available",
			errorCode: "HOST_NOT_AVAILABLE",
		};
	}

	// Outlook doesn't support getFileAsync
	const host = Office?.context?.mailbox ? "outlook" : "";
	if (host === "outlook") {
		return {
			error: "Outlook does not support document export",
			errorCode: "HOST_NOT_SUPPORTED",
		};
	}

	const fileType = format === "native" ? "compressed" : "pdf";

	return new Promise((resolve) => {
		Office.context.document.getFileAsync(
			fileType,
			{ sliceSize: 65536 },
			(result: any) => {
				if (result.status === "succeeded") {
					const file = result.value;
					const sliceCount = file.sliceCount;
					const fileSize = file.size;

					if (fileSize > maxSizeMB * 1024 * 1024) {
						file.closeAsync(() => {});
						resolve({
							error: `File too large: ${(fileSize / 1024 / 1024).toFixed(1)}MB. Max: ${maxSizeMB}MB.`,
							errorCode: "FILE_TOO_LARGE",
							sizeBytes: fileSize,
						});
						return;
					}

					const slices: string[] = [];
					let received = 0;

					function getSlice() {
						file.getSliceAsync(received, (sliceResult: any) => {
							if (sliceResult.status === "succeeded") {
								// Convert byte array to base64
								const bytes = sliceResult.value.data;
								let binary = "";
								for (let i = 0; i < bytes.length; i++) {
									binary += String.fromCharCode(bytes[i]);
								}
								slices.push(btoa(binary));
								received++;

								if (received === sliceCount) {
									const base64 = slices.join("");
									file.closeAsync(() => {});

									resolve({
										base64,
										format,
										sizeBytes: fileSize,
										mimeType:
											format === "pdf"
												? "application/pdf"
												: "application/octet-stream",
										exported: true,
									});
								} else {
									getSlice();
								}
							} else {
								file.closeAsync(() => {});
								resolve({
									error:
										sliceResult.error?.message || "Slice extraction failed",
									errorCode: "HOST_NOT_AVAILABLE",
								});
							}
						});
					}

					getSlice();
				} else {
					resolve({
						error: result.error?.message || "getFileAsync failed",
						errorCode: "HOST_NOT_AVAILABLE",
					});
				}
			},
		);
	});
}
