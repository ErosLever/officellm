/**
 * Word JS API Mock Framework for unit testing.
 *
 * Simulates Word.run(), context.sync(), paragraphs, selection, search, and comments.
 */

export interface MockParagraphData {
	text: string;
	style?: string;
	outlineLevel?: string; // "OutlineLevel1" - "OutlineLevel9" or "OutlineLevelBodyText"
	uniqueLocalId?: string;
}

export interface MockDocumentData {
	paragraphs: MockParagraphData[];
	selectedText?: string;
	comments: Array<{ text: string; paragraphIndex?: number }>;
	changeTrackingMode: "Off" | "TrackAll" | "TrackMineOnly";
	changeLog: Array<{
		type: string;
		paragraphIndex: number;
		oldText?: string;
		newText?: string;
	}>;
}

export class WordMock {
	private _data: MockDocumentData;
	private _originalWord: any;
	private _originalOffice: any;
	reportResultCalls: Array<{
		commandId: string;
		success: boolean;
		error?: string;
		payload?: unknown;
	}> = [];

	constructor(data: MockDocumentData) {
		this._data = data;
	}

	get data() {
		return this._data;
	}
	get lastReport() {
		return this.reportResultCalls.at(-1) ?? null;
	}

	install() {
		this._originalWord = (globalThis as any).Word;
		this._originalOffice = (globalThis as any).Office;

		(globalThis as any).Word = {
			run: (fn: (ctx: any) => Promise<any>) => {
				const ctx = new WordMockContext(this._data);
				return fn(ctx);
			},
		};

		(globalThis as any).Office = {
			onReady: (cb: (info: any) => void) => cb({ host: "Word" }),
			context: { document: { url: "test.docx" } },
		};

		(globalThis as any).window = globalThis;
	}

	restore() {
		(globalThis as any).Word = this._originalWord;
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

	acceptAllChanges() {
		// Clear the change log — changes are accepted (applied)
		this._data.changeLog = [];
	}

	rejectAllChanges() {
		// Revert all tracked changes back to original state
		for (const change of [...this._data.changeLog].reverse()) {
			if (change.type === "replace" && change.oldText !== undefined) {
				this._data.paragraphs[change.paragraphIndex].text = change.oldText;
			} else if (change.type === "insert") {
				this._data.paragraphs.splice(change.paragraphIndex, 1);
			} else if (change.type === "delete" && change.oldText !== undefined) {
				this._data.paragraphs.splice(change.paragraphIndex, 0, {
					text: change.oldText,
					style: "Normal",
					outlineLevel: "OutlineLevelBodyText",
				});
			}
		}
		this._data.changeLog = [];
	}

	reset() {
		this.reportResultCalls = [];
	}
}

// ── Mock Context ────────────────────────────────────────────────

class WordMockContext {
	private _data: MockDocumentData;
	private _loads: Array<{ target: any; props: string[] }> = [];
	private _pendingActions: Array<() => void> = [];
	document: MockDocument;

	constructor(data: MockDocumentData) {
		this._data = data;
		this.document = new MockDocument(this, data);
	}

	queueLoad(target: any, props: string[]) {
		this._loads.push({ target, props });
	}
	queueAction(action: () => void) {
		this._pendingActions.push(action);
	}

	load(target: any, props: string) {
		this._loads.push({ target, props: props.split(",").map((p) => p.trim()) });
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

// ── Mock Document ───────────────────────────────────────────────

class MockDocument {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	body: MockBody;
	private _changeTrackingMode: string;

	constructor(ctx: WordMockContext, data: MockDocumentData) {
		this._ctx = ctx;
		this._data = data;
		this.body = new MockBody(ctx, data);
		this._changeTrackingMode = data.changeTrackingMode || "Off";
	}

	get changeTrackingMode() {
		return this._changeTrackingMode;
	}
	set changeTrackingMode(mode: string) {
		this._changeTrackingMode = mode;
	}

	getSelection() {
		return new MockSelection(this._ctx, this._data);
	}

	acceptAllChanges() {
		this._data.changeLog = [];
	}

	rejectAllChanges() {
		// Revert all tracked changes
		for (const change of [...this._data.changeLog].reverse()) {
			if (change.type === "replace" && change.oldText !== undefined) {
				this._data.paragraphs[change.paragraphIndex].text = change.oldText;
			} else if (change.type === "insert") {
				this._data.paragraphs.splice(change.paragraphIndex, 1);
			} else if (change.type === "delete" && change.oldText !== undefined) {
				this._data.paragraphs.splice(change.paragraphIndex, 0, {
					text: change.oldText,
					style: "Normal",
					outlineLevel: "OutlineLevelBodyText",
				});
			}
		}
		this._data.changeLog = [];
	}
}

// ── Mock Body ───────────────────────────────────────────────────

class MockBody {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	paragraphs: MockParagraphCollection;

	constructor(ctx: WordMockContext, data: MockDocumentData) {
		this._ctx = ctx;
		this._data = data;
		this.paragraphs = new MockParagraphCollection(ctx, data);
	}

	insertParagraph(text: string, location: string) {
		this._ctx.queueAction(() => {
			if (location === "End") {
				this._data.paragraphs.push({
					text,
					style: "Normal",
					outlineLevel: "OutlineLevelBodyText",
				});
			} else {
				this._data.paragraphs.unshift({
					text,
					style: "Normal",
					outlineLevel: "OutlineLevelBodyText",
				});
			}
		});
	}

	search(searchText: string, options: any) {
		return new MockSearchResults(this._ctx, this._data, searchText, options);
	}
}

// ── Mock Paragraph Collection ───────────────────────────────────

class MockParagraphCollection {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	items: MockParagraph[];

	constructor(ctx: WordMockContext, data: MockDocumentData) {
		this._ctx = ctx;
		this._data = data;
		this.items = data.paragraphs.map(
			(p, i) => new MockParagraph(ctx, data, p, i),
		);
	}

	load(props: string) {
		this._ctx.queueLoad(
			this,
			props.split(",").map((p: string) => p.trim()),
		);
	}

	_populate(props: string[]) {
		/* items already populated */
	}
}

// ── Mock Paragraph ──────────────────────────────────────────────

class MockParagraph {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	private _para: MockParagraphData;
	private _index: number;
	private _loaded = new Set<string>();

	// Backing fields (populated after sync)
	private _text = "";
	private _style = "";
	private _outlineLevel = "";
	private _uniqueLocalId = "";

	constructor(
		ctx: WordMockContext,
		data: MockDocumentData,
		para: MockParagraphData,
		index: number,
	) {
		this._ctx = ctx;
		this._data = data;
		this._para = para;
		this._index = index;
		this._uniqueLocalId = para.uniqueLocalId ?? `para_${index}`;
	}

	get text() {
		return this._text;
	}
	get style() {
		return this._style;
	}
	get outlineLevel() {
		return this._outlineLevel;
	}
	get uniqueLocalId() {
		return this._uniqueLocalId;
	}

	load(propString: string) {
		const props = propString.split(",").map((p) => p.trim());
		for (const p of props) this._loaded.add(p);
		this._ctx.queueLoad(this, props);
	}

	_populate(props: string[]) {
		const l = this._loaded;
		if (l.has("text") || props.includes("text")) this._text = this._para.text;
		if (l.has("style") || props.includes("style"))
			this._style = this._para.style ?? "Normal";
		if (l.has("outlineLevel") || props.includes("outlineLevel"))
			this._outlineLevel = this._para.outlineLevel ?? "OutlineLevelBodyText";
		if (l.has("uniqueLocalId") || props.includes("uniqueLocalId"))
			this._uniqueLocalId = this._para.uniqueLocalId ?? `para_${this._index}`;
	}

	getRange(_location: string) {
		return new MockRange(
			this._ctx,
			this._para.text,
			(newText) => {
				this._para.text = newText;
			},
			(commentText) => {
				this._data.comments.push({
					text: commentText,
					paragraphIndex: this._index,
				});
			},
		);
	}

	insertParagraph(text: string, location: string) {
		this._ctx.queueAction(() => {
			const newPara: MockParagraphData = {
				text,
				style: "Normal",
				outlineLevel: "OutlineLevelBodyText",
			};
			if (location === "After") {
				this._data.paragraphs.splice(this._index + 1, 0, newPara);
			} else {
				this._data.paragraphs.splice(this._index, 0, newPara);
			}
		});
	}

	search(searchText: string, options: any) {
		return new MockSearchResults(
			this._ctx,
			this._data,
			searchText,
			options,
			this._para.text,
			(oldText: string, newText: string) => {
				const original = this._para.text;
				this._para.text = this._para.text.replace(oldText, newText);
				this._data.changeLog.push({
					type: "replace",
					paragraphIndex: this._index,
					oldText: original,
					newText: this._para.text,
				});
			}, // propagate replace to paragraph data + log change
		);
	}

	delete() {
		this._ctx.queueAction(() => {
			this._data.paragraphs.splice(this._index, 1);
		});
	}
}

// ── Mock Selection ──────────────────────────────────────────────

class MockSelection {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	private _text: string;
	paragraphs: MockSelectionParagraphs;

	constructor(ctx: WordMockContext, data: MockDocumentData) {
		this._ctx = ctx;
		this._data = data;
		this._text = data.selectedText ?? "";
		this.paragraphs = new MockSelectionParagraphs(ctx, data);
	}

	get text() {
		return this._text;
	}

	load(props: string) {
		this._ctx.queueLoad(
			this,
			props.split(",").map((p) => p.trim()),
		);
	}

	_populate() {
		/* text already set */
	}

	insertComment(commentText: string) {
		this._ctx.queueAction(() => {
			this._data.comments.push({ text: commentText });
		});
	}
}

class MockSelectionParagraphs {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	items: any[];

	constructor(ctx: WordMockContext, data: MockDocumentData) {
		this._ctx = ctx;
		this._data = data;
		// Return all paragraphs as context (simplified)
		this.items = data.paragraphs.map(
			(p, i) => new MockParagraph(ctx, data, p, i),
		);
	}

	load(props: string) {
		// Queue load for each item
		for (const item of this.items) {
			item.load(props);
		}
	}
}

// ── Mock Search Results ─────────────────────────────────────────

class MockSearchResults {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	private _searchText: string;
	private _options: any;
	private _scopeText?: string;
	items: MockRange[];

	constructor(
		ctx: WordMockContext,
		data: MockDocumentData,
		searchText: string,
		options: any,
		scopeText?: string,
		onReplace?: (oldText: string, newText: string) => void,
	) {
		this._ctx = ctx;
		this._data = data;
		this._searchText = searchText;
		this._options = options;
		this._scopeText = scopeText;

		// Find matches
		const matchCase = options?.matchCase ?? false;
		const scope = scopeText ?? data.paragraphs.map((p) => p.text).join("\n");
		const searchIn = matchCase ? scope : scope.toLowerCase();
		const needle = matchCase ? searchText : searchText.toLowerCase();

		this.items = [];
		let pos = 0;
		while (true) {
			const idx = searchIn.indexOf(needle, pos);
			if (idx === -1) break;
			this.items.push(
				new MockRange(
					ctx,
					scope.substring(idx, idx + searchText.length),
					onReplace,
					undefined, // no comment callback for search results
				),
			);
			pos = idx + 1;
		}
	}

	load(props: string) {
		for (const item of this.items) {
			item.load(props);
		}
	}
}

// ── Mock Range ──────────────────────────────────────────────────

class MockRange {
	private _ctx: WordMockContext;
	private _text: string;
	private _onReplace?: (oldText: string, newText: string) => void;
	private _onComment?: (text: string) => void;

	constructor(
		ctx: WordMockContext,
		text: string,
		onReplace?: (oldText: string, newText: string) => void,
		onComment?: (text: string) => void,
	) {
		this._ctx = ctx;
		this._text = text;
		this._onReplace = onReplace;
		this._onComment = onComment;
	}

	get text() {
		return this._text;
	}

	load(_props: string) {
		// Already populated
	}

	insertText(text: string, location: string) {
		const oldText = this._text;
		this._ctx.queueAction(() => {
			if (location === "Replace") {
				this._text = text;
				if (this._onReplace) this._onReplace(oldText, text);
			}
		});
	}

	insertComment(commentText: string) {
		this._ctx.queueAction(() => {
			if (this._onComment) this._onComment(commentText);
		});
	}
}
