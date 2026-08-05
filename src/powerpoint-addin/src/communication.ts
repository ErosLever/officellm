/**
 * Communication module for the Office JS Add-in.
 * Connects TO the MCP server — registers as an instance,
 * polls for commands, and reports results.
 * Supports both SignalR (WebSocket) and HTTP polling fallback.
 */

import * as signalR from "@microsoft/signalr";

export const MCP_SERVER_URL = "http://127.0.0.1:3000";

// Injected by webpack DefinePlugin at build time from .env MCP_API_TOKEN.
declare const __MCP_TOKEN__: string;
const MCP_TOKEN: string = __MCP_TOKEN__;

/** Returns the Authorization header object for fetch calls. */
function authHeaders(): Record<string, string> {
	return { Authorization: `Bearer ${MCP_TOKEN}` };
}

// --- State ---
let instanceId: string | null = null;
let hubConnection: signalR.HubConnection | null = null;
let _registeredAppName: string = "Unknown";
let _registeredDocumentName: string = "(active)";

export type ConnectionState = "connected" | "reconnecting" | "fallback";

let _connectionState: ConnectionState = "fallback";
let _onConnectionStateChange: ((state: ConnectionState) => void) | null = null;

/**
 * Sets a callback for connection state changes (used by task pane UI).
 */
export function onConnectionStateChange(
	callback: (state: ConnectionState) => void,
): void {
	_onConnectionStateChange = callback;
}

function setConnectionState(state: ConnectionState): void {
	_connectionState = state;
	_onConnectionStateChange?.(state);
}

/**
 * Returns current connection state.
 */
export function getConnectionState(): ConnectionState {
	return _connectionState;
}
// ============================================================
// INSTANCE REGISTRATION & HEARTBEAT
// ============================================================

/**
 * Registers this add-in instance with the MCP server.
 * Returns the assigned instance ID.
 */
export async function registerWithMcp(
	appName: string,
	documentName: string,
): Promise<string> {
	_registeredAppName = appName;
	_registeredDocumentName = documentName;
	const response = await fetch(`${MCP_SERVER_URL}/instances/register`, {
		method: "POST",
		headers: { "Content-Type": "application/json", ...authHeaders() },
		body: JSON.stringify({ appName, documentName }),
	});

	if (!response.ok) {
		throw new Error(`Registration failed: HTTP ${response.status}`);
	}

	const data = await response.json();
	instanceId = data.instanceId ?? "";
	console.log(`Registered with MCP server: ${instanceId}`);
	return instanceId!;
}

// ============================================================
// SIGNALR CONNECTION
// ============================================================

export type CommandHandler = (
	commandId: string,
	commandName: string,
	args: unknown,
) => Promise<unknown>;

let _commandHandler: CommandHandler | null = null;

/**
 * Sets the command handler for incoming SignalR commands.
 */
export function setCommandHandler(handler: CommandHandler): void {
	_commandHandler = handler;
}

/**
 * Connects to the SignalR hub for real-time command delivery.
 * Falls back to HTTP polling if WebSocket fails.
 */
export async function connectSignalR(): Promise<void> {
	if (!instanceId) return;

	const connection = new signalR.HubConnectionBuilder()
		.withUrl(`${MCP_SERVER_URL}/hubs/commands?access_token=${MCP_TOKEN}`)
		.withAutomaticReconnect([0, 2000, 5000, 10000, 30000])
		.configureLogging(signalR.LogLevel.Warning)
		.build();

	// Handle incoming commands from server
	connection.on(
		"ExecuteCommand",
		async (commandId: string, commandName: string, args: unknown) => {
			console.log(`SignalR: Received command ${commandName} (${commandId})`);
			if (_commandHandler) {
				try {
					const result = await _commandHandler(commandId, commandName, args);
					// Report result back via SignalR
					const success = !(
						result &&
						typeof result === "object" &&
						"error" in (result as any)
					);
					const error = success ? undefined : ((result as any).error as string);
					await connection.invoke(
						"ReportResult",
						commandId,
						success,
						error,
						result,
					);
				} catch (err) {
					const errMsg = err instanceof Error ? err.message : String(err);
					await connection.invoke(
						"ReportResult",
						commandId,
						false,
						errMsg,
						null,
					);
				}
			}
		},
	);

	connection.onreconnecting(() => {
		console.log("SignalR: Reconnecting...");
		setConnectionState("reconnecting");
	});

	connection.onreconnected(() => {
		console.log("SignalR: Reconnected");
		setConnectionState("connected");
		// Re-join the instance group after reconnect
		connection
			.invoke("JoinGroup", instanceId)
			.catch((err: unknown) =>
				console.warn("SignalR: Failed to rejoin group:", err),
			);
	});

	connection.onclose(() => {
		console.log("SignalR: Connection closed");
		setConnectionState("fallback");
	});

	try {
		await connection.start();
		await connection.invoke("JoinGroup", instanceId);
		hubConnection = connection;
		setConnectionState("connected");
		console.log(`SignalR: Connected and joined group ${instanceId}`);
	} catch (err) {
		console.warn(
			"SignalR: Connection failed, falling back to HTTP polling:",
			err,
		);
		hubConnection = null;
		setConnectionState("fallback");
	}
}

/**
 * Sends a heartbeat to keep this instance alive.
 * Should be called periodically (e.g., every 10 seconds).
 */
export async function sendHeartbeat(): Promise<void> {
	if (!instanceId) return;

	try {
		await fetch(`${MCP_SERVER_URL}/instances/${instanceId}/heartbeat`, {
			method: "POST",
			headers: { "Content-Type": "application/json", ...authHeaders() },
			body: JSON.stringify({ appName: _registeredAppName, documentName: _registeredDocumentName }),
		});
	} catch (error) {
		console.warn("Heartbeat failed:", error);
	}
}

/**
 * Starts periodic heartbeat polling.
 */
export function startHeartbeat(intervalMs = 10000): void {
	sendHeartbeat(); // Send immediately
	setInterval(sendHeartbeat, intervalMs);
}

// ============================================================
// COMMAND POLLING
// ============================================================

interface PendingCommand {
	id: string;
	command: string;
	args?: unknown;
}

/**
 * Polls the MCP server for pending commands.
 * Returns an array of unclaimed commands.
 */
export async function pollForCommands(): Promise<PendingCommand[]> {
	if (!instanceId) return [];

	try {
		const response = await fetch(
			`${MCP_SERVER_URL}/instances/${instanceId}/commands`,
			{ headers: authHeaders() },
		);
		if (!response.ok) return [];

		const data = await response.json();
		return data.commands || [];
	} catch (error) {
		console.warn("Command poll failed:", error);
		return [];
	}
}

/**
 * Reports a command result back to the MCP server.
 * Tries SignalR first, falls back to HTTP.
 */
export async function reportResult(
	commandId: string,
	success: boolean,
	error?: string,
	payload?: unknown,
): Promise<void> {
	if (!instanceId) return;

	// Try SignalR first (instant)
	if (
		hubConnection &&
		hubConnection.state === signalR.HubConnectionState.Connected
	) {
		try {
			await hubConnection.invoke(
				"ReportResult",
				commandId,
				success,
				error,
				payload,
			);
			return;
		} catch (err) {
			console.warn("SignalR result report failed, falling back to HTTP:", err);
		}
	}

	const body: Record<string, unknown> = {
		commandId,
		success,
	};
	if (error) body.error = error;
	if (payload) body.payload = payload;

	const url = `${MCP_SERVER_URL}/instances/${instanceId}/result`;
	const MAX_RETRIES = 3;

	for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
		try {
			const response = await fetch(url, {
				method: "POST",
				headers: { "Content-Type": "application/json", ...authHeaders() },
				body: JSON.stringify(body),
			});
			if (response.ok) return;
			if (attempt === MAX_RETRIES) {
				console.warn(
					`Result report failed: HTTP ${response.status} after ${MAX_RETRIES} attempts`,
				);
			}
		} catch (err) {
			if (attempt === MAX_RETRIES) {
				console.warn(
					`Result report failed after ${MAX_RETRIES} attempts:`,
					err,
				);
			} else {
				await new Promise((r) => setTimeout(r, 500 * attempt));
			}
		}
	}
}

// ============================================================
// OFFICE STATE RETRIEVAL
// ============================================================

export interface OfficeState {
	app: string;
	documentName: string;
	slideCount: number;
	currentSlideIndex: number;
}

/**
 * Gets basic Office state information.
 */
export function getOfficeState(): Promise<OfficeState> {
	return new Promise((resolve) => {
		Office.onReady((info) => {
			// Try to get the real document name from Office context
			let documentName = "Untitled";
			try {
				// eslint-disable-next-line @typescript-eslint/no-explicit-any
				const doc: any = (window as any).Office?.context?.document;
				if (doc?.url) {
					// Extract filename from URL path
					try {
						documentName = decodeURIComponent(
							doc.url.split("/").pop() || "Untitled",
						);
					} catch {
						documentName = doc.url;
					}
				}
			} catch {
				// Office context not available yet
			}

			resolve({
				app: (info.host as unknown as string) || "Unknown",
				documentName,
				slideCount: 0,
				currentSlideIndex: 0,
			});
		});
	});
}
