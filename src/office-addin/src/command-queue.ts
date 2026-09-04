/**
 * Serializes command execution so at most one Office JS mutation is in flight
 * at a time for a given add-in instance.
 *
 * Commands can arrive via two independent paths — a SignalR push per command, and
 * the HTTP-polling fallback loop — and office_batch_call on the MCP server fires
 * several tool calls concurrently, pushing multiple SignalR events at once. Without
 * this queue, two commands (e.g. two powerpoint_duplicate_slide calls) can overlap:
 * each reads the pre-mutation slide count in its own PowerPoint.run() before either
 * has written, producing a lost update. Routing every command through runSerialized
 * guarantees they execute one at a time, in arrival order, regardless of which path
 * they arrived on.
 */

let queue: Promise<unknown> = Promise.resolve();

export function runSerialized<T>(fn: () => Promise<T>): Promise<T> {
	const result = queue.then(fn, fn);
	queue = result.then(
		() => undefined,
		() => undefined,
	);
	return result;
}
