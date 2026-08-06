/**
 * Quick interactive test for PowerPoint JS API.
 * Paste into the browser DevTools console while the task pane is open.
 */

/* eslint-disable no-undef */
// @ts-nocheck

async function testPowerPointAPI() {
	console.log("=== PowerPoint JS API Test ===");

	if (!PowerPoint || typeof PowerPoint.run !== "function") {
		console.error("PowerPoint.run not available!");
		return;
	}

	await PowerPoint.run(async (context) => {
		const presentation = context.presentation;

		// Test 1: Load slides
		console.log("\n--- Test 1: Load slides ---");
		presentation.load("slides");
		await context.sync();
		console.log(`Slides count: ${presentation.slides.items.length}`);

		// Test 2: Load shapes items on first slide
		console.log("\n--- Test 2: Load shapes on slide 0 ---");
		if (presentation.slides.items.length > 0) {
			const slide = presentation.slides.items[0];
			slide.load("shapes/items/$none");
			await context.sync();
			console.log(`Shapes count on slide 0: ${slide.shapes.items.length}`);

			// Test 3: Load id and name per shape
			console.log("\n--- Test 3: Load id, name per shape ---");
			for (const shape of slide.shapes.items) {
				shape.load("id");
				shape.load("name");
			}
			await context.sync();
			for (const shape of slide.shapes.items) {
				console.log(`  shape: id=${shape.id}, name=${shape.name}`);
			}

			// Test 4: getTextFrameOrNullObject
			console.log("\n--- Test 4: getTextFrameOrNullObject ---");
			for (const shape of slide.shapes.items) {
				const tf = shape.getTextFrameOrNullObject();
				tf.load("textRange/text");
			}
			await context.sync();

			for (const shape of slide.shapes.items) {
				const tf = shape.getTextFrameOrNullObject();
				if (tf.isNullObject) {
					console.log(`  shape "${shape.name}" (${shape.id}): no textFrame`);
				} else {
					const text = tf.textRange?.text || "";
					console.log(
						`  shape "${shape.name}" (${shape.id}): "${text.substring(0, 80)}${text.length > 80 ? "..." : ""}"`,
					);
				}
			}

			// Test 5: Full deck outline
			console.log("\n--- Test 5: Full deck outline ---");
			for (const s of presentation.slides.items) {
				s.load("shapes/items/$none");
			}
			await context.sync();

			for (const s of presentation.slides.items) {
				for (const shape of s.shapes.items) {
					const tf = shape.getTextFrameOrNullObject();
					tf.load("textRange/text");
				}
			}
			await context.sync();

			for (let i = 0; i < presentation.slides.items.length; i++) {
				const s = presentation.slides.items[i];
				let title = "";
				for (const shape of s.shapes.items) {
					const tf = shape.getTextFrameOrNullObject();
					if (!tf.isNullObject && tf.textRange?.text) {
						const t = tf.textRange.text.trim();
						if (t && !title) title = t;
					}
				}
				console.log(`  Slide ${i}: "${title || "(no title)"}"`);
			}
		}

		console.log("\n=== All tests completed ===");
	});
}

testPowerPointAPI();
