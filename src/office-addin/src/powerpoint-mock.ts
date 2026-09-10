/**
 * PowerPoint JS API Mock Framework for unit testing.
 *
 * Faithfully replicates the Office JS API's queued-command + sync pattern:
 * - Properties are NOT available until ctx.sync() is called
 * - load() queues property names; sync() populates them from mock data
 * - getTextFrameOrNullObject() creates a NEW object each call
 * - getImageAsBase64() returns ClientResult<string> — sync() to read .value
 * - Setters (text, position, font) defer until sync()
 *
 * Usage:
 *   const mock = new PowerPointMock({ slides: [...] });
 *   mock.install();  // sets window.PowerPoint
 *   // ... run command handler ...
 *   mock.restore();  // cleanup
 */

// ── Types ───────────────────────────────────────────────────────

export interface MockShapeData {
	id: string;
	name: string;
	type?: string;
	left?: number;
	top?: number;
	width?: number;
	height?: number;
	rotation?: number;
	text?: string;
	font?: {
		name?: string;
		size?: number;
		bold?: boolean;
		italic?: boolean;
		color?: string;
		underline?: string;
		strikethrough?: boolean;
		doubleStrikethrough?: boolean;
		allCaps?: boolean;
		smallCaps?: boolean;
		subscript?: boolean;
		superscript?: boolean;
	};
	paragraphFormat?: {
		horizontalAlignment?: string;
		indentLevel?: number;
		bulletType?: string;
		bulletStyle?: string;
		bulletVisible?: boolean;
	};
	// Per-paragraph overrides, indexed by paragraph index — independent of
	// the flat font/paragraphFormat bags above, which back the whole-shape
	// textRange used by powerpoint_update_shape_properties.
	paragraphs?: Array<{
		font?: MockShapeData["font"];
		paragraphFormat?: MockShapeData["paragraphFormat"];
	}>;
	textFrame?: {
		marginTop?: number;
		marginBottom?: number;
		marginLeft?: number;
		marginRight?: number;
		autoSizeSetting?: string;
		wordWrap?: boolean;
		verticalAlignment?: string;
	};
	fillColor?: string;
	fillTransparency?: number;
	tableCells?: string[][];
}

export interface MockSlideData {
	id?: string;
	shapes: MockShapeData[];
	notes?: string;
	layoutId?: string;
}

export interface MockLayoutData {
	id: string;
	name: string;
}

export interface MockSlideMasterData {
	id: string;
	name: string;
	layouts: MockLayoutData[];
}

export interface MockPresentationData {
	slides: MockSlideData[];
	slideMasters?: MockSlideMasterData[];
}

const DEFAULT_SLIDE_MASTERS: MockSlideMasterData[] = [
	{
		id: "master_0",
		name: "Office Theme",
		layouts: [
			{ id: "layout_0", name: "Title Slide" },
			{ id: "layout_1", name: "Title and Content" },
			{ id: "layout_2", name: "Blank" },
		],
	},
];

// ── ClientResult ────────────────────────────────────────────────

class ClientResult<T> {
	value: T;
	constructor(val: T) {
		this.value = val;
	}
	_populate() {
		/* already populated */
	}
}

// ── PowerPointMock (top-level test utility) ─────────────────────

export class PowerPointMock {
	private _data: MockPresentationData;
	private _originalPowerPoint: any;
	private _originalOffice: any;
	reportResultCalls: Array<{
		commandId: string;
		success: boolean;
		error?: string;
		payload?: unknown;
	}> = [];

	constructor(data: MockPresentationData) {
		this._data = data;
	}

	get data() {
		return this._data;
	}
	get lastReport() {
		return this.reportResultCalls.at(-1) ?? null;
	}

	install() {
		this._originalPowerPoint = (globalThis as any).PowerPoint;
		this._originalOffice = (globalThis as any).Office;

		(globalThis as any).PowerPoint = {
			run: (fn: (ctx: any) => Promise<any>) => {
				const ctx = new MockContext(this._data);
				return fn(ctx);
			},
			ShapeType: {
				image: "Image",
				textBox: "TextBox",
				table: "Table",
				geometricShape: "GeometricShape",
				group: "Group",
				line: "Line",
			},
			InsertSlideFormatting: {
				keepSourceFormatting: "KeepSourceFormatting",
				useDestinationTheme: "UseDestinationTheme",
			},
		};

		(globalThis as any).Office = {
			onReady: (cb: (info: any) => void) => cb({ host: "PowerPoint" }),
			context: { document: { url: "test.pptx" } },
		};
	}

	restore() {
		(globalThis as any).PowerPoint = this._originalPowerPoint;
		(globalThis as any).Office = this._originalOffice;
	}

	mockReportResult = async (
		commandId: string,
		success: boolean,
		error?: string,
		payload?: unknown,
	) => {
		this.reportResultCalls.push({ commandId, success, error, payload });
	};

	reset() {
		this.reportResultCalls = [];
	}
}

// ── Mock Sync Context ───────────────────────────────────────────

interface QueuedLoad {
	target: any;
	props: string[];
}

class MockContext {
	private _data: MockPresentationData;
	private _loads: QueuedLoad[] = [];
	private _pendingActions: Array<() => void> = [];
	presentation: MockPresentation;

	constructor(data: MockPresentationData) {
		this._data = data;
		if (!this._data.slideMasters) {
			this._data.slideMasters = DEFAULT_SLIDE_MASTERS;
		}
		this.presentation = new MockPresentation(this, data);
	}

	queueLoad(target: any, props: string[]) {
		this._loads.push({ target, props });
	}
	queueAction(action: () => void) {
		this._pendingActions.push(action);
	}

	// ctx.load(target, props) — context-level load (used for textFrame, fill, etc.)
	load(target: any, props: string) {
		const propList = props.split(",").map((p) => p.trim());
		this._loads.push({ target, props: propList });
	}

	async sync() {
		for (const a of this._pendingActions) a();
		this._pendingActions = [];
		for (const l of this._loads) {
			if (l.target._populate) l.target._populate(l.props);
		}
		this._loads = [];
	}
}

// ── Mock Presentation ───────────────────────────────────────────

class MockPresentation {
	private _ctx: MockContext;
	slides: MockSlideCollection;
	slideMasters: MockSlideMasterCollection;

	constructor(ctx: MockContext, data: MockPresentationData) {
		this._ctx = ctx;
		this.slides = new MockSlideCollection(ctx, data.slides);
		this.slideMasters = new MockSlideMasterCollection(ctx, data.slideMasters ?? []);
	}

	load(prop: string) {
		if (prop === "slides") this._ctx.queueLoad(this.slides, ["items"]);
	}

	getSelectedShapes() {
		return new MockSelectedShapes();
	}
	getSelectedTextRange() {
		throw new Error("No text selected");
	}

	insertSlidesFromBase64(_base64File: string, options?: any) {
		this._ctx.queueAction(() => {
			const data = this.slides._data;
			const targetSlideId = options?.targetSlideId;
			const insertAt = targetSlideId
				? data.findIndex((s) => s.id === targetSlideId) + 1
				: 0;
			const newSlide: MockSlideData = {
				id: `slide_dup_${data.length}`,
				shapes: [],
				notes: "",
			};
			data.splice(insertAt, 0, newSlide);
			this.slides._rebuild();
		});
	}
}

// ── Mock Slide Layout ────────────────────────────────────────────

class MockSlideLayout {
	private _ctx: MockContext;
	private _data: MockLayoutData;
	private _loaded = new Set<string>();
	private _id = "";
	private _name = "";
	isNullObject = false;

	constructor(ctx: MockContext, data: MockLayoutData, isNullObject = false) {
		this._ctx = ctx;
		this._data = data;
		this.isNullObject = isNullObject;
	}

	get id() {
		return this._id;
	}
	get name() {
		return this._name;
	}

	load(props: string | string[]) {
		const propList = Array.isArray(props)
			? props
			: props.split(",").map((p) => p.trim());
		for (const p of propList) this._loaded.add(p);
		this._ctx.queueLoad(this, propList);
	}

	_populate(props: string[]) {
		const l = this._loaded;
		if (l.has("id") || props.includes("id")) this._id = this._data.id;
		if (l.has("name") || props.includes("name")) this._name = this._data.name;
	}
}

// ── Mock Slide Layout Collection ─────────────────────────────────

class MockSlideLayoutCollection {
	private _ctx: MockContext;
	private _data: MockLayoutData[];
	items: MockSlideLayout[];

	constructor(ctx: MockContext, data: MockLayoutData[]) {
		this._ctx = ctx;
		this._data = data;
		this.items = data.map((d) => new MockSlideLayout(ctx, d));
	}

	load(props: string | string[]) {
		const propList = Array.isArray(props)
			? props
			: props.split(",").map((p) => p.trim());
		if (propList.includes("items")) this._ctx.queueLoad(this, ["items"]);
	}

	getItem(key: string): MockSlideLayout {
		const found = this.items.find((l) => (l as any)._data.id === key);
		if (!found) throw new Error(`Layout '${key}' not found`);
		return found;
	}

	getItemAt(index: number): MockSlideLayout {
		const found = this.items[index];
		if (!found) throw new Error(`Layout index ${index} out of range`);
		return found;
	}

	getItemOrNullObject(key: string): MockSlideLayout {
		const found = this.items.find((l) => (l as any)._data.id === key);
		if (found) return found;
		return new MockSlideLayout(this._ctx, { id: "", name: "" }, true);
	}

	_populate() {
		/* items already populated */
	}
}

// ── Mock Slide Master ─────────────────────────────────────────────

class MockSlideMaster {
	private _ctx: MockContext;
	private _data: MockSlideMasterData;
	private _loaded = new Set<string>();
	private _id = "";
	private _name = "";
	layouts: MockSlideLayoutCollection;
	isNullObject = false;

	constructor(
		ctx: MockContext,
		data: MockSlideMasterData,
		isNullObject = false,
	) {
		this._ctx = ctx;
		this._data = data;
		this.isNullObject = isNullObject;
		this.layouts = new MockSlideLayoutCollection(ctx, data.layouts ?? []);
	}

	get id() {
		return this._id;
	}
	get name() {
		return this._name;
	}

	load(props: string | string[]) {
		const propList = Array.isArray(props)
			? props
			: props.split(",").map((p) => p.trim());
		for (const p of propList) this._loaded.add(p);
		this._ctx.queueLoad(this, propList);
	}

	_populate(props: string[]) {
		const l = this._loaded;
		if (l.has("id") || props.includes("id")) this._id = this._data.id;
		if (l.has("name") || props.includes("name")) this._name = this._data.name;
	}
}

// ── Mock Slide Master Collection ──────────────────────────────────

class MockSlideMasterCollection {
	private _ctx: MockContext;
	private _data: MockSlideMasterData[];
	items: MockSlideMaster[];

	constructor(ctx: MockContext, data: MockSlideMasterData[]) {
		this._ctx = ctx;
		this._data = data;
		this.items = data.map((d) => new MockSlideMaster(ctx, d));
	}

	load(props: string | string[]) {
		const propList = Array.isArray(props)
			? props
			: props.split(",").map((p) => p.trim());
		if (propList.includes("items")) this._ctx.queueLoad(this, ["items"]);
	}

	getItem(key: string): MockSlideMaster {
		const found = this.items.find((m) => (m as any)._data.id === key);
		if (!found) throw new Error(`Slide master '${key}' not found`);
		return found;
	}

	getItemAt(index: number): MockSlideMaster {
		const found = this.items[index];
		if (!found) throw new Error(`Slide master index ${index} out of range`);
		return found;
	}

	getItemOrNullObject(key: string): MockSlideMaster {
		const found = this.items.find((m) => (m as any)._data.id === key);
		if (found) return found;
		return new MockSlideMaster(
			this._ctx,
			{ id: "", name: "", layouts: [] },
			true,
		);
	}

	_populate() {
		/* items already populated */
	}
}

// ── Mock Applied Slide Layout (slide.layout) ────────────────────

class MockAppliedLayout {
	private _ctx: MockContext;
	private _slideData: MockSlideData;
	private _loaded = new Set<string>();
	private _id = "";
	private _name = "";

	constructor(ctx: MockContext, slideData: MockSlideData) {
		this._ctx = ctx;
		this._slideData = slideData;
	}

	get id() {
		return this._id;
	}
	get name() {
		return this._name;
	}

	load(props: string | string[]) {
		const propList = Array.isArray(props)
			? props
			: props.split(",").map((p) => p.trim());
		for (const p of propList) this._loaded.add(p);
		this._ctx.queueLoad(this, propList);
	}

	_populate(props: string[]) {
		const l = this._loaded;
		const layoutId = this._slideData.layoutId ?? "";
		const masters = (this._ctx.presentation.slideMasters as any)
			._data as MockSlideMasterData[];
		let name = "";
		for (const master of masters) {
			const found = master.layouts.find((ly) => ly.id === layoutId);
			if (found) {
				name = found.name;
				break;
			}
		}
		if (l.has("id") || props.includes("id")) this._id = layoutId;
		if (l.has("name") || props.includes("name")) this._name = name;
	}
}

// ── Mock Slide Collection ───────────────────────────────────────

class MockSlideCollection {
	private _ctx: MockContext;
	_data: MockSlideData[];
	items: MockSlide[];

	constructor(ctx: MockContext, data: MockSlideData[]) {
		this._ctx = ctx;
		this._data = data;
		this.items = data.map((s, i) => new MockSlide(ctx, s, i));
	}

	add(options?: any) {
		const newSlide: MockSlideData = {
			id: `slide_${this._data.length}`,
			shapes: [],
			notes: "",
			layoutId: options?.layoutId,
		};
		this._data.push(newSlide);
		this.items.push(new MockSlide(this._ctx, newSlide, this._data.length - 1));
	}

	_rebuild() {
		this.items = this._data.map((s, i) => new MockSlide(this._ctx, s, i));
	}

	_populate() {
		/* items already populated */
	}
}

// ── Mock Slide ──────────────────────────────────────────────────

class MockSlide {
	private _ctx: MockContext;
	private _data: MockSlideData;
	private _index: number;
	private _idBacking: string = "";
	shapes: MockShapeCollection;
	layout: MockAppliedLayout;

	constructor(ctx: MockContext, data: MockSlideData, index: number) {
		this._ctx = ctx;
		this._data = data;
		this._index = index;
		this._idBacking = data.id ?? `slide_${index}`;
		this.shapes = new MockShapeCollection(ctx, data.shapes, this);
		this.layout = new MockAppliedLayout(ctx, data);
	}

	get id() {
		return this._idBacking;
	}

	load(prop: string) {
		if (prop === "shapes/items/$none")
			this._ctx.queueLoad(this.shapes, ["items"]);
	}

	_populate() {}

	getImageAsBase64(options?: any): ClientResult<string> {
		const fakePng =
			"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
		const result = new ClientResult(`data:image/png;base64,${fakePng}`);
		this._ctx.queueLoad(result, []);
		return result;
	}

	exportAsBase64(): ClientResult<string> {
		const result = new ClientResult("fake-slide-export-base64");
		this._ctx.queueLoad(result, []);
		return result;
	}

	delete() {
		this._ctx.queueAction(() => {
			this._data.id = "__deleted__";
		});
	}

	moveTo(newIndex: number) {
		this._ctx.queueAction(() => {
			this._data.id = `moved_to_${newIndex}`;
		});
	}

	applyLayout(layout: MockSlideLayout) {
		this._ctx.queueAction(() => {
			this._data.layoutId = (layout as any)._data.id;
		});
	}

	getNotesSlide() {
		return new MockNotesSlide(this._ctx, this._data);
	}
	getNotesSlideOrNullObject() {
		return new MockNotesSlide(this._ctx, this._data);
	}
}

// ── Mock Notes Slide ────────────────────────────────────────────

class MockNotesSlide {
	private _ctx: MockContext;
	private _data: MockSlideData;
	textFrame: MockTextFrame;
	isNullObject = false;

	constructor(ctx: MockContext, data: MockSlideData) {
		this._ctx = ctx;
		this._data = data;
		this.textFrame = new MockTextFrame(
			ctx,
			{ text: data.notes ?? "", font: {} },
			true,
		);
	}

	load(props: string) {
		this._ctx.queueLoad(
			this.textFrame,
			props.split(",").map((p) => p.trim()),
		);
	}
}

// ── Mock Shape Collection ───────────────────────────────────────

class MockShapeCollection {
	private _ctx: MockContext;
	private _data: MockShapeData[];
	private _slide: MockSlide;
	items: MockShape[];

	constructor(ctx: MockContext, data: MockShapeData[], slide: MockSlide) {
		this._ctx = ctx;
		this._data = data;
		this._slide = slide;
		this.items = data.map((d) => new MockShape(ctx, d, this));
	}

	addTextBox(text: string, options: any): MockShape {
		const d: MockShapeData = {
			id: `shape_${this._data.length}`,
			name: `TextBox ${this._data.length}`,
			type: "TextBox",
			text,
			left: options.left,
			top: options.top,
			width: options.width,
			height: options.height,
		};
		this._data.push(d);
		const s = new MockShape(this._ctx, d, this);
		this.items.push(s);
		return s;
	}

	addPicture(base64: string, options: any): MockShape {
		const d: MockShapeData = {
			id: `shape_${this._data.length}`,
			name: `Picture ${this._data.length}`,
			type: "Image",
			left: options.left,
			top: options.top,
			width: options.width,
			height: options.height,
		};
		this._data.push(d);
		const s = new MockShape(this._ctx, d, this);
		this.items.push(s);
		return s;
	}

	addTable(rows: number, columns: number, options: any): MockShape {
		const cells: string[][] = Array.from({ length: rows }, () =>
			Array(columns).fill(""),
		);
		const d: MockShapeData = {
			id: `shape_${this._data.length}`,
			name: `Table ${this._data.length}`,
			type: "Table",
			left: options.left,
			top: options.top,
			width: options.width,
			height: options.height,
			tableCells: cells,
		};
		this._data.push(d);
		const s = new MockShape(this._ctx, d, this);
		this.items.push(s);
		return s;
	}

	/** Remove shape from collection */
	removeShape(shape: MockShape) {
		const idx = this.items.indexOf(shape);
		if (idx >= 0) this.items.splice(idx, 1);
	}
}

// ── Mock Shape ──────────────────────────────────────────────────

class MockShape {
	private _ctx: MockContext;
	private _data: MockShapeData;
	private _collection: MockShapeCollection;
	private _loaded = new Set<string>();

	// Backing fields — populated by _populate after sync()
	private _id = "";
	private _name = "";
	private _type = "TextBox";
	private _left = 0;
	private _top = 0;
	private _width = 0;
	private _height = 0;
	private _rotation = 0;
	fill: MockFill;

	constructor(
		ctx: MockContext,
		data: MockShapeData,
		collection: MockShapeCollection,
	) {
		this._ctx = ctx;
		this._data = data;
		this._collection = collection;
		this.fill = new MockFill(ctx, data);
	}

	// Getters
	get id() {
		return this._id;
	}
	get name() {
		return this._name;
	}
	get type() {
		return this._type;
	}
	get left() {
		return this._left;
	}
	get top() {
		return this._top;
	}
	get width() {
		return this._width;
	}
	get height() {
		return this._height;
	}
	get rotation() {
		return this._rotation;
	}

	// Setters — defer until sync
	set left(v) {
		this._ctx.queueAction(() => {
			this._data.left = v;
			this._left = v;
		});
	}
	set top(v) {
		this._ctx.queueAction(() => {
			this._data.top = v;
			this._top = v;
		});
	}
	set width(v) {
		this._ctx.queueAction(() => {
			this._data.width = v;
			this._width = v;
		});
	}
	set height(v) {
		this._ctx.queueAction(() => {
			this._data.height = v;
			this._height = v;
		});
	}
	set rotation(v) {
		this._ctx.queueAction(() => {
			this._data.rotation = v;
			this._rotation = v;
		});
	}

	load(propString: string) {
		const props = propString.split(",").map((p) => p.trim());
		for (const p of props) this._loaded.add(p);
		this._ctx.queueLoad(this, props);
	}

	_populate(props: string[]) {
		const l = this._loaded;
		if (l.has("id") || props.includes("id")) this._id = this._data.id;
		if (l.has("name") || props.includes("name")) this._name = this._data.name;
		if (l.has("type") || props.includes("type"))
			this._type = this._data.type ?? "TextBox";
		if (l.has("left") || props.includes("left"))
			this._left = this._data.left ?? 0;
		if (l.has("top") || props.includes("top")) this._top = this._data.top ?? 0;
		if (l.has("width") || props.includes("width"))
			this._width = this._data.width ?? 0;
		if (l.has("height") || props.includes("height"))
			this._height = this._data.height ?? 0;
		if (l.has("rotation") || props.includes("rotation"))
			this._rotation = this._data.rotation ?? 0;
	}

	getTextFrameOrNullObject(): MockTextFrame {
		const hasText = this._data.text !== undefined;
		// getTextFrameOrNullObject() creates a new backing-data object each
		// call (per this file's header comment), but `paragraphs` must be
		// the SAME array across calls for setter mutations made via one
		// call's getSubstring() to be visible to a later call's getter —
		// so vivify it on the persistent MockShapeData (this._data), not on
		// the fresh copy below, before handing out a reference to it.
		this._data.paragraphs ??= [];
		return new MockTextFrame(
			this._ctx,
			{
				text: this._data.text ?? "",
				font: this._data.font,
				paragraphFormat: this._data.paragraphFormat,
				textFrame: this._data.textFrame,
				paragraphs: this._data.paragraphs,
			},
			hasText,
		);
	}

	getTable(): MockTable {
		return new MockTable(this._ctx, this._data);
	}

	getImageAsBase64(options?: any): ClientResult<string> {
		const fakePng =
			"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
		const result = new ClientResult(`data:image/png;base64,${fakePng}`);
		this._ctx.queueLoad(result, []);
		return result;
	}

	delete() {
		this._ctx.queueAction(() => {
			this._collection.removeShape(this);
		});
	}
}

// ── Mock Fill ───────────────────────────────────────────────────

class MockFill {
	private _data: MockShapeData;
	foregroundColor = "";
	transparency = 0;

	constructor(_ctx: MockContext, data: MockShapeData) {
		this._data = data;
	}
	_populate() {
		this.foregroundColor = this._data.fillColor ?? "";
		this.transparency = this._data.fillTransparency ?? 0;
	}
}

// ── Mock TextFrame ──────────────────────────────────────────────

type TextFrameBackingData = {
	text: string;
	font?: MockShapeData["font"];
	paragraphFormat?: MockShapeData["paragraphFormat"];
	textFrame?: MockShapeData["textFrame"];
	paragraphs?: MockShapeData["paragraphs"];
};

class MockTextFrame {
	private _ctx: MockContext;
	private _data: TextFrameBackingData;
	isNullObject: boolean;
	textRange: MockTextRange;

	private _topMargin = 0;
	private _bottomMargin = 0;
	private _leftMargin = 0;
	private _rightMargin = 0;
	private _autoSizeSetting = "AutoSizeNone";
	private _wordWrap = true;
	private _verticalAlignment = "Top";

	constructor(ctx: MockContext, data: TextFrameBackingData, hasText: boolean) {
		this._ctx = ctx;
		this._data = data;
		this.isNullObject = !hasText;
		this.textRange = new MockTextRange(ctx, data);
	}

	get topMargin() {
		return this._topMargin;
	}
	get bottomMargin() {
		return this._bottomMargin;
	}
	get leftMargin() {
		return this._leftMargin;
	}
	get rightMargin() {
		return this._rightMargin;
	}
	get autoSizeSetting() {
		return this._autoSizeSetting;
	}
	get wordWrap() {
		return this._wordWrap;
	}
	get verticalAlignment() {
		return this._verticalAlignment;
	}

	set topMargin(v: number) {
		this._ctx.queueAction(() => {
			if (!this._data.textFrame) this._data.textFrame = {};
			this._data.textFrame.marginTop = v;
			this._topMargin = v;
		});
	}
	set bottomMargin(v: number) {
		this._ctx.queueAction(() => {
			if (!this._data.textFrame) this._data.textFrame = {};
			this._data.textFrame.marginBottom = v;
			this._bottomMargin = v;
		});
	}
	set leftMargin(v: number) {
		this._ctx.queueAction(() => {
			if (!this._data.textFrame) this._data.textFrame = {};
			this._data.textFrame.marginLeft = v;
			this._leftMargin = v;
		});
	}
	set rightMargin(v: number) {
		this._ctx.queueAction(() => {
			if (!this._data.textFrame) this._data.textFrame = {};
			this._data.textFrame.marginRight = v;
			this._rightMargin = v;
		});
	}
	set autoSizeSetting(v: string) {
		this._ctx.queueAction(() => {
			if (!this._data.textFrame) this._data.textFrame = {};
			this._data.textFrame.autoSizeSetting = v;
			this._autoSizeSetting = v;
		});
	}
	set wordWrap(v: boolean) {
		this._ctx.queueAction(() => {
			if (!this._data.textFrame) this._data.textFrame = {};
			this._data.textFrame.wordWrap = v;
			this._wordWrap = v;
		});
	}
	set verticalAlignment(v: string) {
		this._ctx.queueAction(() => {
			if (!this._data.textFrame) this._data.textFrame = {};
			this._data.textFrame.verticalAlignment = v;
			this._verticalAlignment = v;
		});
	}

	_populate(props: string[]) {
		const tfData = this._data.textFrame;
		this._topMargin = tfData?.marginTop ?? 0;
		this._bottomMargin = tfData?.marginBottom ?? 0;
		this._leftMargin = tfData?.marginLeft ?? 0;
		this._rightMargin = tfData?.marginRight ?? 0;
		this._autoSizeSetting = tfData?.autoSizeSetting ?? "AutoSizeNone";
		this._wordWrap = tfData?.wordWrap ?? true;
		this._verticalAlignment = tfData?.verticalAlignment ?? "Top";
		this.textRange._populate(props);
	}
}

// ── Mock TextRange ──────────────────────────────────────────────

class MockTextRange {
	private _ctx: MockContext;
	private _data: TextFrameBackingData;
	private _text: string;
	font: MockFont;
	paragraphFormat: MockParagraphFormat;

	constructor(ctx: MockContext, data: TextFrameBackingData) {
		this._ctx = ctx;
		this._data = data;
		this._text = data.text;
		this.font = new MockFont(ctx, data);
		this.paragraphFormat = new MockParagraphFormat(ctx, data);
	}

	get text() {
		return this._text;
	}
	set text(v: string) {
		const data = this._data;
		this._ctx.queueAction(() => {
			data.text = v;
		});
	}

	_populate(props: string[]) {
		this.font._populate(props);
		this.paragraphFormat._populate(props);
	}

	/**
	 * Mirrors PowerPoint.TextRange.getSubstring(start, length) for the
	 * single-paragraph-range case — the only shape getSubstring() calls
	 * take in this codebase (see splitParagraphs in powerpoint-commands.ts).
	 * The paragraph index is just the count of `\r` delimiters before
	 * `start`; the substring itself is a direct slice, no re-splitting of
	 * the whole text needed.
	 */
	getSubstring(start: number, length: number): MockTextRange {
		const index = this._data.text.slice(0, start).split("\r").length - 1;

		if (!this._data.paragraphs) this._data.paragraphs = [];
		if (!this._data.paragraphs[index]) this._data.paragraphs[index] = {};
		const paragraphData = this._data.paragraphs[index];
		// Merge the shape's flat font/paragraphFormat underneath any
		// paragraph-specific override (fixture-provided, or set by an
		// earlier getSubstring-backed mutation) so a partial override
		// still inherits every other key from the shape's real values —
		// matching real PowerPoint, where an untouched sub-range inherits
		// the whole range's formatting rather than reverting to blank
		// defaults. Existing override keys win via spread order.
		paragraphData.font = { ...this._data.font, ...paragraphData.font };
		paragraphData.paragraphFormat = {
			...this._data.paragraphFormat,
			...paragraphData.paragraphFormat,
		};

		return new MockTextRange(this._ctx, {
			text: this._data.text.substring(start, start + length),
			font: paragraphData.font,
			paragraphFormat: paragraphData.paragraphFormat,
		});
	}
}

// ── Mock Font ───────────────────────────────────────────────────

class MockFont {
	private _ctx: MockContext;
	private _data: TextFrameBackingData;
	private _name = "";
	private _size = 0;
	private _bold = false;
	private _italic = false;
	private _color = "";
	private _underline = "None";
	private _strikethrough = false;
	private _doubleStrikethrough = false;
	private _allCaps = false;
	private _smallCaps = false;
	private _subscript = false;
	private _superscript = false;

	constructor(ctx: MockContext, data: TextFrameBackingData) {
		this._ctx = ctx;
		this._data = data;
	}

	get name() {
		return this._name;
	}
	get size() {
		return this._size;
	}
	get bold() {
		return this._bold;
	}
	get italic() {
		return this._italic;
	}
	get color() {
		return this._color;
	}
	get underline() {
		return this._underline;
	}
	get strikethrough() {
		return this._strikethrough;
	}
	get doubleStrikethrough() {
		return this._doubleStrikethrough;
	}
	get allCaps() {
		return this._allCaps;
	}
	get smallCaps() {
		return this._smallCaps;
	}
	get subscript() {
		return this._subscript;
	}
	get superscript() {
		return this._superscript;
	}

	set name(v) {
		this._ctx.queueAction(() => {
			if (this._data.font) this._data.font.name = v;
			this._name = v;
		});
	}
	set size(v) {
		this._ctx.queueAction(() => {
			if (this._data.font) this._data.font.size = v;
			this._size = v;
		});
	}
	set bold(v) {
		this._ctx.queueAction(() => {
			if (this._data.font) this._data.font.bold = v;
			this._bold = v;
		});
	}
	set italic(v) {
		this._ctx.queueAction(() => {
			if (this._data.font) this._data.font.italic = v;
			this._italic = v;
		});
	}
	set color(v) {
		this._ctx.queueAction(() => {
			if (this._data.font) this._data.font.color = v;
			this._color = v;
		});
	}
	set underline(v: string) {
		this._ctx.queueAction(() => {
			if (!this._data.font) this._data.font = {};
			this._data.font.underline = v;
			this._underline = v;
		});
	}
	set strikethrough(v: boolean) {
		this._ctx.queueAction(() => {
			if (!this._data.font) this._data.font = {};
			this._data.font.strikethrough = v;
			this._strikethrough = v;
		});
	}
	set doubleStrikethrough(v: boolean) {
		this._ctx.queueAction(() => {
			if (!this._data.font) this._data.font = {};
			this._data.font.doubleStrikethrough = v;
			this._doubleStrikethrough = v;
		});
	}
	set allCaps(v: boolean) {
		this._ctx.queueAction(() => {
			if (!this._data.font) this._data.font = {};
			this._data.font.allCaps = v;
			this._allCaps = v;
		});
	}
	set smallCaps(v: boolean) {
		this._ctx.queueAction(() => {
			if (!this._data.font) this._data.font = {};
			this._data.font.smallCaps = v;
			this._smallCaps = v;
		});
	}
	set subscript(v: boolean) {
		this._ctx.queueAction(() => {
			if (!this._data.font) this._data.font = {};
			this._data.font.subscript = v;
			this._subscript = v;
		});
	}
	set superscript(v: boolean) {
		this._ctx.queueAction(() => {
			if (!this._data.font) this._data.font = {};
			this._data.font.superscript = v;
			this._superscript = v;
		});
	}

	_populate(props: string[]) {
		const f = this._data.font;
		this._name = f?.name ?? "";
		this._size = f?.size ?? 0;
		this._bold = f?.bold ?? false;
		this._italic = f?.italic ?? false;
		this._color = f?.color ?? "";
		this._underline = f?.underline ?? "None";
		this._strikethrough = f?.strikethrough ?? false;
		this._doubleStrikethrough = f?.doubleStrikethrough ?? false;
		this._allCaps = f?.allCaps ?? false;
		this._smallCaps = f?.smallCaps ?? false;
		this._subscript = f?.subscript ?? false;
		this._superscript = f?.superscript ?? false;
	}
}

// ── Mock ParagraphFormat / BulletFormat ──────────────────────────

class MockBulletFormat {
	private _ctx: MockContext;
	private _data: TextFrameBackingData;
	private _type = "None";
	private _style = "None";
	private _visible = false;

	constructor(ctx: MockContext, data: TextFrameBackingData) {
		this._ctx = ctx;
		this._data = data;
	}

	get type() {
		return this._type;
	}
	get style() {
		return this._style;
	}
	get visible() {
		return this._visible;
	}

	set type(v: string) {
		this._ctx.queueAction(() => {
			if (!this._data.paragraphFormat) this._data.paragraphFormat = {};
			this._data.paragraphFormat.bulletType = v;
			this._type = v;
		});
	}
	set style(v: string) {
		this._ctx.queueAction(() => {
			if (!this._data.paragraphFormat) this._data.paragraphFormat = {};
			this._data.paragraphFormat.bulletStyle = v;
			this._style = v;
		});
	}
	set visible(v: boolean) {
		this._ctx.queueAction(() => {
			if (!this._data.paragraphFormat) this._data.paragraphFormat = {};
			this._data.paragraphFormat.bulletVisible = v;
			this._visible = v;
		});
	}

	_populate() {
		const pf = this._data.paragraphFormat;
		this._type = pf?.bulletType ?? "None";
		this._style = pf?.bulletStyle ?? "None";
		this._visible = pf?.bulletVisible ?? false;
	}
}

class MockParagraphFormat {
	private _ctx: MockContext;
	private _data: TextFrameBackingData;
	private _horizontalAlignment = "Left";
	private _indentLevel = 0;
	bulletFormat: MockBulletFormat;

	constructor(ctx: MockContext, data: TextFrameBackingData) {
		this._ctx = ctx;
		this._data = data;
		this.bulletFormat = new MockBulletFormat(ctx, data);
	}

	get horizontalAlignment() {
		return this._horizontalAlignment;
	}
	get indentLevel() {
		return this._indentLevel;
	}

	set horizontalAlignment(v: string) {
		this._ctx.queueAction(() => {
			if (!this._data.paragraphFormat) this._data.paragraphFormat = {};
			this._data.paragraphFormat.horizontalAlignment = v;
			this._horizontalAlignment = v;
		});
	}
	set indentLevel(v: number) {
		this._ctx.queueAction(() => {
			if (!this._data.paragraphFormat) this._data.paragraphFormat = {};
			this._data.paragraphFormat.indentLevel = v;
			this._indentLevel = v;
		});
	}

	_populate(props: string[]) {
		const pf = this._data.paragraphFormat;
		this._horizontalAlignment = pf?.horizontalAlignment ?? "Left";
		this._indentLevel = pf?.indentLevel ?? 0;
		this.bulletFormat._populate();
	}
}

// ── Mock Table ──────────────────────────────────────────────────

class MockTable {
	private _ctx: MockContext;
	private _data: MockShapeData;
	rowCount = 0;
	columnCount = 0;

	constructor(ctx: MockContext, data: MockShapeData) {
		this._ctx = ctx;
		this._data = data;
	}

	load(props: string) {
		const cells = this._data.tableCells;
		if (cells) {
			this.rowCount = cells.length;
			this.columnCount = cells[0]?.length ?? 0;
		}
	}

	getCell(row: number, col: number): MockTableCell {
		return new MockTableCell(this._ctx, this._data, row, col);
	}
}

class MockTableCell {
	private _ctx: MockContext;
	private _data: MockShapeData;
	private _row: number;
	private _col: number;
	textFrame: MockTextFrame;

	constructor(ctx: MockContext, data: MockShapeData, row: number, col: number) {
		this._ctx = ctx;
		this._data = data;
		this._row = row;
		this._col = col;
		const cellText = data.tableCells?.[row]?.[col] ?? "";
		this.textFrame = new MockTextFrame(ctx, { text: cellText, font: {} }, true);
	}
}

// ── Mock SelectedShapes ─────────────────────────────────────────

class MockSelectedShapes {
	items: any[] = [];
	load() {}
}
