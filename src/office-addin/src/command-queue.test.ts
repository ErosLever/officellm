import { describe, it, expect, beforeEach, vi } from "vitest";
import { runSerialized } from "./command-queue";

describe("runSerialized", () => {
	beforeEach(() => {
		vi.resetModules();
	});

	it("runs overlapping calls one at a time, in arrival order", async () => {
		const active: number[] = [];
		let maxConcurrent = 0;
		const order: number[] = [];

		function makeTask(id: number, delayMs: number) {
			return () =>
				new Promise<void>((resolve) => {
					active.push(id);
					maxConcurrent = Math.max(maxConcurrent, active.length);
					setTimeout(() => {
						order.push(id);
						active.splice(active.indexOf(id), 1);
						resolve();
					}, delayMs);
				});
		}

		// Task 1 is slower than task 2 — if they were not serialized, task 2
		// would finish first. Serialization must force task 1 to finish first.
		const p1 = runSerialized(makeTask(1, 30));
		const p2 = runSerialized(makeTask(2, 5));

		await Promise.all([p1, p2]);

		expect(maxConcurrent).toBe(1);
		expect(order).toEqual([1, 2]);
	});

	it("continues processing the queue after a task rejects", async () => {
		const order: number[] = [];

		const p1 = runSerialized(() => Promise.reject(new Error("boom")));
		const p2 = runSerialized(
			() =>
				new Promise<void>((resolve) => {
					order.push(2);
					resolve();
				}),
		);

		await expect(p1).rejects.toThrow("boom");
		await expect(p2).resolves.toBeUndefined();
		expect(order).toEqual([2]);
	});

	it("does not let a later call start before an earlier call's mutation is visible", async () => {
		// Simulates two "slide count" reads racing a mutation: without serialization,
		// task B would read the stale count before task A's increment lands.
		let sharedCount = 0;
		const observedByB: number[] = [];

		const taskA = () =>
			new Promise<void>((resolve) => {
				setTimeout(() => {
					sharedCount += 1;
					resolve();
				}, 10);
			});

		const taskB = () =>
			new Promise<void>((resolve) => {
				observedByB.push(sharedCount);
				resolve();
			});

		await Promise.all([runSerialized(taskA), runSerialized(taskB)]);

		expect(observedByB).toEqual([1]);
	});
});
