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

export interface MockInlinePictureData {
	base64: string;
	width: number;
	height: number;
	imageFormat?: string;
}

export interface MockCommentReplyData {
	id?: string;
	text: string;
	authorName?: string;
	authorEmail?: string;
	creationDate?: Date;
}

export interface MockCommentData {
	text: string;
	paragraphIndex?: number;
	id?: string;
	authorName?: string;
	authorEmail?: string;
	creationDate?: Date;
	resolved?: boolean;
	// Absolute document character offsets for the comment's anchor range.
	// When omitted, defaults to the whole span of `paragraphIndex`'s paragraph
	// (computed from cumulative paragraph text lengths, one separator char
	// between paragraphs — mirrors Word.Range.start/end well enough for tests).
	anchorStart?: number;
	anchorEnd?: number;
	// Document text under the comment's anchor range — distinct from `text`
	// (the comment's own content). Defaults to the anchored paragraph's text.
	anchorText?: string;
	replies?: MockCommentReplyData[];
}

export interface MockDocumentData {
	paragraphs: MockParagraphData[];
	selectedText?: string;
	comments: MockCommentData[];
	changeTrackingMode: "Off" | "TrackAll" | "TrackMineOnly";
	changeLog: Array<{
		type: string;
		paragraphIndex: number;
		oldText?: string;
		newText?: string;
	}>;
	inlinePictures?: MockInlinePictureData[];
}

// Cumulative absolute character offsets per paragraph — one separator
// character assumed between paragraphs, mirroring how Word.Range.start/end
// account for paragraph marks.
function computeParagraphOffsets(paragraphs: MockParagraphData[]): Array<{ start: number; end: number }> {
	const offsets: Array<{ start: number; end: number }> = [];
	let pos = 0;
	for (const p of paragraphs) {
		const start = pos;
		const end = start + p.text.length;
		offsets.push({ start, end });
		pos = end + 1;
	}
	return offsets;
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

	acceptAllRevisions() {
		// Clear the change log — changes are accepted (applied)
		this._data.changeLog = [];
	}

	rejectAllRevisions() {
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

	load(_props: string) {
		// Mock properties are always synchronously available; nothing to queue.
	}

	getSelection() {
		return new MockSelection(this._ctx, this._data);
	}

	acceptAllRevisions() {
		this._data.changeLog = [];
	}

	rejectAllRevisions() {
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
	inlinePictures: MockInlinePictureCollection;

	constructor(ctx: WordMockContext, data: MockDocumentData) {
		this._ctx = ctx;
		this._data = data;
		this.paragraphs = new MockParagraphCollection(ctx, data);
		this.inlinePictures = new MockInlinePictureCollection(ctx, data);
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

	getComments() {
		return new MockCommentCollection(this._ctx, this._data);
	}
}

// ── Mock Comments ────────────────────────────────────────────────

class MockCommentCollection {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	items: MockComment[];

	constructor(ctx: WordMockContext, data: MockDocumentData) {
		this._ctx = ctx;
		this._data = data;
		this.items = data.comments.map((c, i) => new MockComment(ctx, data, c, i));
	}

	load(_props: string) {
		this._ctx.queueLoad(this, []);
	}

	_populate(_props: string[]) {
		/* items already populated */
	}
}

class MockComment {
	private _ctx: WordMockContext;
	private _data: MockDocumentData;
	private _comment: MockCommentData;
	private _index: number;
	private _loaded = new Set<string>();

	private _id = "";
	private _content = "";
	private _authorName = "";
	private _authorEmail = "";
	private _creationDate: Date = new Date(0);
	private _resolved = false;
	replies: MockCommentReplyCollection;

	constructor(ctx: WordMockContext, data: MockDocumentData, comment: MockCommentData, index: number) {
		this._ctx = ctx;
		this._data = data;
		this._comment = comment;
		this._index = index;
		this._id = comment.id ?? `comment_${index}`;
		this.replies = new MockCommentReplyCollection(ctx, comment);
	}

	get id() {
		return this._id;
	}
	get content() {
		return this._content;
	}
	set content(text: string) {
		this._content = text;
		this._ctx.queueAction(() => {
			this._comment.text = text;
		});
	}
	get authorName() {
		return this._authorName;
	}
	get authorEmail() {
		return this._authorEmail;
	}
	get creationDate() {
		return this._creationDate;
	}
	get resolved() {
		return this._resolved;
	}
	set resolved(value: boolean) {
		this._resolved = value;
		this._ctx.queueAction(() => {
			this._comment.resolved = value;
		});
	}

	load(propString: string) {
		const props = propString.split(",").map((p) => p.trim());
		for (const p of props) this._loaded.add(p);
		this._ctx.queueLoad(this, props);
	}

	_populate(props: string[]) {
		const l = this._loaded;
		if (l.has("id") || props.includes("id")) this._id = this._comment.id ?? `comment_${this._index}`;
		if (l.has("content") || props.includes("content")) this._content = this._comment.text;
		if (l.has("authorName") || props.includes("authorName"))
			this._authorName = this._comment.authorName ?? "";
		if (l.has("authorEmail") || props.includes("authorEmail"))
			this._authorEmail = this._comment.authorEmail ?? "";
		if (l.has("creationDate") || props.includes("creationDate"))
			this._creationDate = this._comment.creationDate ?? new Date(0);
		if (l.has("resolved") || props.includes("resolved"))
			this._resolved = this._comment.resolved ?? false;
	}

	getRange() {
		const paragraphIndex = this._comment.paragraphIndex ?? 0;
		const para = this._data.paragraphs[paragraphIndex];
		const offsets = computeParagraphOffsets(this._data.paragraphs)[paragraphIndex] ?? { start: 0, end: 0 };
		const start = this._comment.anchorStart ?? offsets.start;
		const end = this._comment.anchorEnd ?? offsets.end;
		const anchorText = this._comment.anchorText ?? para?.text ?? "";
		return new MockRange(this._ctx, anchorText, undefined, undefined, start, end);
	}

	reply(text: string) {
		if (!this._comment.replies) this._comment.replies = [];
		const replyData: MockCommentReplyData = { text, id: `reply_${this._comment.replies.length}` };
		this._ctx.queueAction(() => {
			this._comment.replies!.push(replyData);
		});
		const newReply = new MockCommentReply(this._ctx, this._comment, replyData, this.replies.items.length);
		this.replies.items.push(newReply);
		return newReply;
	}

	delete() {
		this._ctx.queueAction(() => {
			const idx = this._data.comments.indexOf(this._comment);
			if (idx !== -1) this._data.comments.splice(idx, 1);
		});
	}
}

class MockCommentReplyCollection {
	private _ctx: WordMockContext;
	private _comment: MockCommentData;
	items: MockCommentReply[];

	constructor(ctx: WordMockContext, comment: MockCommentData) {
		this._ctx = ctx;
		this._comment = comment;
		this.items = (comment.replies ?? []).map((r, i) => new MockCommentReply(ctx, comment, r, i));
	}

	load(_props: string) {
		this._ctx.queueLoad(this, []);
	}

	_populate(_props: string[]) {
		/* items already populated */
	}
}

class MockCommentReply {
	private _ctx: WordMockContext;
	private _comment: MockCommentData;
	private _reply: MockCommentReplyData;
	private _index: number;
	private _loaded = new Set<string>();

	private _id = "";
	private _content = "";
	private _authorName = "";
	private _authorEmail = "";
	private _creationDate: Date = new Date(0);

	constructor(ctx: WordMockContext, comment: MockCommentData, reply: MockCommentReplyData, index: number) {
		this._ctx = ctx;
		this._comment = comment;
		this._reply = reply;
		this._index = index;
		this._id = reply.id ?? `reply_${index}`;
	}

	get id() {
		return this._id;
	}
	get content() {
		return this._content;
	}
	set content(text: string) {
		this._content = text;
		this._ctx.queueAction(() => {
			this._reply.text = text;
		});
	}
	get authorName() {
		return this._authorName;
	}
	get authorEmail() {
		return this._authorEmail;
	}
	get creationDate() {
		return this._creationDate;
	}

	load(propString: string) {
		const props = propString.split(",").map((p) => p.trim());
		for (const p of props) this._loaded.add(p);
		this._ctx.queueLoad(this, props);
	}

	_populate(props: string[]) {
		const l = this._loaded;
		if (l.has("id") || props.includes("id")) this._id = this._reply.id ?? `reply_${this._index}`;
		if (l.has("content") || props.includes("content")) this._content = this._reply.text;
		if (l.has("authorName") || props.includes("authorName"))
			this._authorName = this._reply.authorName ?? "";
		if (l.has("authorEmail") || props.includes("authorEmail"))
			this._authorEmail = this._reply.authorEmail ?? "";
		if (l.has("creationDate") || props.includes("creationDate"))
			this._creationDate = this._reply.creationDate ?? new Date(0);
	}

	delete() {
		this._ctx.queueAction(() => {
			const replies = this._comment.replies;
			if (!replies) return;
			const idx = replies.indexOf(this._reply);
			if (idx !== -1) replies.splice(idx, 1);
		});
	}
}

// ── Mock Inline Picture Collection ──────────────────────────────

class MockInlinePictureCollection {
	private _ctx: WordMockContext;
	items: MockInlinePicture[];

	constructor(ctx: WordMockContext, data: MockDocumentData) {
		this._ctx = ctx;
		this.items = (data.inlinePictures ?? []).map((p) => new MockInlinePicture(p));
	}

	load(_props: string) {
		this._ctx.queueLoad(this, []);
	}

	_populate(_props: string[]) {
		/* items already populated */
	}
}

class MockInlinePicture {
	private _pic: MockInlinePictureData;

	constructor(pic: MockInlinePictureData) {
		this._pic = pic;
	}

	get width() {
		return this._pic.width;
	}
	get height() {
		return this._pic.height;
	}
	get imageFormat() {
		return this._pic.imageFormat ?? "Png";
	}

	load(_props: string) {
		/* already populated */
	}

	getBase64ImageSrc() {
		return { value: this._pic.base64 };
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
		const offsets = computeParagraphOffsets(this._data.paragraphs)[this._index];
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
					anchorStart: offsets.start,
					anchorEnd: offsets.end,
				});
			},
			offsets.start,
			offsets.end,
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
	private _start: number;
	private _end: number;

	constructor(
		ctx: WordMockContext,
		text: string,
		onReplace?: (oldText: string, newText: string) => void,
		onComment?: (text: string) => void,
		start = 0,
		end = 0,
	) {
		this._ctx = ctx;
		this._text = text;
		this._onReplace = onReplace;
		this._onComment = onComment;
		this._start = start;
		this._end = end;
	}

	get text() {
		return this._text;
	}
	get start() {
		return this._start;
	}
	get end() {
		return this._end;
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

	expandTo(other: MockRange): MockRange {
		return new MockRange(
			this._ctx,
			this._text,
			this._onReplace,
			this._onComment,
			Math.min(this._start, other.start),
			Math.max(this._end, other.end),
		);
	}
}
