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
	};
	fillColor?: string;
	fillTransparency?: number;
	tableCells?: string[][];
}

export interface MockSlideData {
	id?: string;
	shapes: MockShapeData[];
	notes?: string;
}

export interface MockPresentationData {
	slides: MockSlideData[];
}

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

	constructor(ctx: MockContext, data: MockPresentationData) {
		this._ctx = ctx;
		this.slides = new MockSlideCollection(ctx, data.slides);
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
}

// ── Mock Slide Collection ───────────────────────────────────────

class MockSlideCollection {
	private _ctx: MockContext;
	private _data: MockSlideData[];
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
		};
		this._data.push(newSlide);
		this.items.push(new MockSlide(this._ctx, newSlide, this._data.length - 1));
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

	constructor(ctx: MockContext, data: MockSlideData, index: number) {
		this._ctx = ctx;
		this._data = data;
		this._index = index;
		this._idBacking = data.id ?? `slide_${index}`;
		this.shapes = new MockShapeCollection(ctx, data.shapes, this);
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
		return new MockTextFrame(
			this._ctx,
			{ text: this._data.text ?? "", font: this._data.font },
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

class MockTextFrame {
	private _ctx: MockContext;
	private _data: { text: string; font?: MockShapeData["font"] };
	isNullObject: boolean;
	textRange: MockTextRange;

	constructor(
		ctx: MockContext,
		data: { text: string; font?: MockShapeData["font"] },
		hasText: boolean,
	) {
		this._ctx = ctx;
		this._data = data;
		this.isNullObject = !hasText;
		this.textRange = new MockTextRange(ctx, data);
	}

	_populate(props: string[]) {
		this.textRange._populate(props);
	}
}

// ── Mock TextRange ──────────────────────────────────────────────

class MockTextRange {
	private _ctx: MockContext;
	private _data: { text: string; font?: MockShapeData["font"] };
	private _text: string;
	font: MockFont;

	constructor(
		ctx: MockContext,
		data: { text: string; font?: MockShapeData["font"] },
	) {
		this._ctx = ctx;
		this._data = data;
		this._text = data.text;
		this.font = new MockFont(ctx, data);
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
	}
}

// ── Mock Font ───────────────────────────────────────────────────

class MockFont {
	private _ctx: MockContext;
	private _data: { text: string; font?: MockShapeData["font"] };
	private _name = "";
	private _size = 0;
	private _bold = false;
	private _italic = false;
	private _color = "";

	constructor(
		ctx: MockContext,
		data: { text: string; font?: MockShapeData["font"] },
	) {
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

	_populate(props: string[]) {
		const f = this._data.font;
		this._name = f?.name ?? "";
		this._size = f?.size ?? 0;
		this._bold = f?.bold ?? false;
		this._italic = f?.italic ?? false;
		this._color = f?.color ?? "";
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
