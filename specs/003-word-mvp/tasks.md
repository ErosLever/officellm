# Tasks: Word MVP Integration with Tracked Changes

**Input**: Design documents from `specs/003-word-mvp/`

**Prerequisites**: plan.md (required), spec.md (required)

**Tests**: TDD workflow — write tests first, then implement.

**Organization**: Tasks grouped by user story. Most foundational work is already done (8 Word tools, handlers, mocks, 128 tests). This plan adds tracked changes support and validates against real Word.

## Format: `[ID] [P?] [Story] Description`

- **[P]**: Can run in parallel (different files, no dependencies)
- **[Story]**: Which user story this task belongs to

## Path Conventions

```
src/mcp-server/Tools/           # C# backend
src/office-addin/src/       # TypeScript frontend
tests/mcp-server.Tests/         # C# tests
specs/003-word-mvp/             # Feature docs
.specify/memory/                # Constitution
```

---

## Phase 1: Spec & Documentation Alignment

**Purpose**: Update spec and constitution to reflect tracked changes approach before writing code

- [ ] T001 Update `specs/003-word-mvp/spec.md` — replace `word_rewrite_selection` with `word_replace_text` (tracked changes), remove confirmation gate references, add FR entries for `word_accept_all_changes`, `word_reject_all_changes`, `word_get_tracked_changes`, note PowerPoint has no tracked changes
- [ ] T002 [P] Update `.specify/memory/constitution.md` section III — amend "User Control & Transparency" to accept Word tracked changes as valid confirmation mechanism, add note that PowerPoint uses undo groups instead
- [ ] T003 [P] Update `AGENTS.md` — add Word JS API rules (changeTrackingMode pattern, load/sync rules, getTextFrame patterns), update tool count from 26 to 29, document tracked changes as default mutation mode

**Checkpoint**: Spec, constitution, and dev docs all reflect tracked changes approach

---

## Phase 2: Backend — New Tool Definitions (US1: Change Tracking Infrastructure)

**Goal**: Add 3 new Word tools to the C# MCP server and update all test assertions

**Independent Test**: `dotnet test` passes with 29 tools (was 26). All new tools appear in OpenAPI.

### Tests First

- [ ] T004 Add C# test `Mcp_ToolsList_Returns29Tools` in `tests/mcp-server.Tests/HttpEndpointTests.cs` — verify tool count is 29
- [ ] T005 [P] Add C# test `Mcp_NewTrackedChangeTools_RequireInstanceId` in `tests/mcp-server.Tests/HttpEndpointTests.cs` — [Theory] test for `word_accept_all_changes`, `word_reject_all_changes` requiring instanceId
- [ ] T006 [P] Update existing `OpenApi_ContainsAllToolEndpoints` in `tests/mcp-server.Tests/HttpEndpointTests.cs` — add 3 new Word endpoints to assertion

### Implementation

- [ ] T007 Add 3 new tool definitions to `src/mcp-server/Tools/McpToolEngine.cs` — `word_get_tracked_changes` (no instanceId required), `word_accept_all_changes` (instanceId required), `word_reject_all_changes` (instanceId required), with descriptions and empty inputSchema
- [ ] T008 Update `AddInCommands` HashSet in `src/mcp-server/Tools/McpToolEngine.cs` — add `word_get_tracked_changes`, `word_accept_all_changes`, `word_reject_all_changes`
- [ ] T009 Run `dotnet test tests/mcp-server.Tests/` — verify all C# tests pass (should be ~82)

**Checkpoint**: Backend recognizes 29 tools. C# tests green. OpenAPI includes all Word tools.

---

## Phase 3: Frontend — Tracked Changes in Word Handlers (US2: Mutation Tracking)

**Goal**: All Word mutation handlers enable changeTrackingMode before writing. 3 new handlers added.

**Independent Test**: `npm test` passes with mutation results including `tracked: true`. New handler tests for accept/reject.

### Tests First

- [ ] T010 [P] Add mock `changeTrackingMode` support to `src/office-addin/src/word-mock.ts` — add `changeTrackingMode` field to `MockDocumentData` interface, `changeLog` array for tracking mutations, `acceptAllChanges()` and `rejectAllChanges()` methods on `WordMock`
- [ ] T011 [P] Add test `word_replace_text returns tracked: true` in `src/office-addin/src/word-commands.test.ts` — verify mutation result includes `tracked: true` and mock changeLog records the change
- [ ] T012 [P] Add test `word_insert_text returns tracked: true` in `src/office-addin/src/word-commands.test.ts` — verify `tracked: true` in result
- [ ] T013 [P] Add test `word_delete_paragraph returns tracked: true` in `src/office-addin/src/word-commands.test.ts` — verify `tracked: true` in result
- [ ] T014 [P] Add test `word_get_tracked_changes returns mode and count` in `src/office-addin/src/word-commands.test.ts` — verify returns `changeTrackingMode`, `pendingChanges` count
- [ ] T015 [P] Add test `word_accept_all_changes clears changeLog` in `src/office-addin/src/word-commands.test.ts` — perform mutation, accept all, verify changeLog empty
- [ ] T016 [P] Add test `word_reject_all_changes restores original text` in `src/office-addin/src/word-commands.test.ts` — perform replace, reject all, verify paragraph text unchanged

### Implementation

- [ ] T017 Create `withChangeTracking<T>()` wrapper in `src/office-addin/src/word-commands.ts` — helper that saves changeTrackingMode, enables TrackMineOnly, runs mutation, restores original mode. Returns `{ tracked: true, ...result }`
- [ ] T018 Wrap `handleReplaceText` mutation with `withChangeTracking` in `src/office-addin/src/word-commands.ts`
- [ ] T019 [P] Wrap `handleInsertText` mutation with `withChangeTracking` in `src/office-addin/src/word-commands.ts`
- [ ] T020 [P] Wrap `handleDeleteParagraph` mutation with `withChangeTracking` in `src/office-addin/src/word-commands.ts`
- [ ] T021 Add `handleGetTrackedChanges` handler in `src/office-addin/src/word-commands.ts` — loads `document.changeTrackingMode`, counts pending revisions via `context.document.body.search` or mock equivalent
- [ ] T022 [P] Add `handleAcceptAllChanges` handler in `src/office-addin/src/word-commands.ts` — calls `context.document.acceptAllChanges()` (if API available) or iterates revisions
- [ ] T023 [P] Add `handleRejectAllChanges` handler in `src/office-addin/src/word-commands.ts` — calls `context.document.rejectAllChanges()` (if API available) or iterates revisions
- [ ] T024 Add 3 new `case` entries to `processCommand` switch in `src/office-addin/src/word-commands.ts` — `word_get_tracked_changes`, `word_accept_all_changes`, `word_reject_all_changes`
- [ ] T025 Run `cd src/office-addin && npm test` — verify all TS tests pass (should be ~55+)
- [ ] T026 Run `cd src/office-addin && npx webpack --mode production` — verify build succeeds

**Checkpoint**: All Word mutations return `tracked: true`. Accept/reject tools work in mock. TS tests green. Webpack builds.

---

## Phase 4: Manual Word Validation (US3: Real API Testing)

**Goal**: Validate all 11 Word tools against real Word JS API in a running Word instance

**Independent Test**: Each tool works correctly when invoked from Word task pane

- [ ] T027 Build and sideload add-in in Word — run `npx webpack --mode production`, open Word, insert add-in from shared folder, verify task pane loads and registers as `word_*` instance
- [ ] T028 Test `word_get_outline` in real Word — open multi-heading document, call tool, verify heading levels match document structure. Document any API differences.
- [ ] T029 [P] Test `word_get_paragraphs` in real Word — verify pagination (startIndex/count), paragraph text and style returned correctly
- [ ] T030 [P] Test `word_get_selection` in real Word — select text, call tool, verify selected text and paragraph context returned
- [ ] T031 [P] Test `word_search` in real Word — search for known text, verify match count and positions. Test case-sensitive search.
- [ ] T032 Test `word_replace_text` in real Word — replace text in a paragraph. Verify tracked change appears in Word Review pane. Verify changeTrackingMode save/restore works.
- [ ] T033 [P] Test `word_insert_text` in real Word — insert at end, insert after paragraph. Verify tracked change markers appear.
- [ ] T034 [P] Test `word_add_comment` in real Word — add comment on selection, add comment on paragraph. Verify Word comment appears in Review pane.
- [ ] T035 [P] Test `word_delete_paragraph` in real Word — delete a paragraph. Verify tracked deletion appears in Review pane.
- [ ] T036 Test `word_get_tracked_changes` in real Word — after mutations, call tool, verify it returns pending change count
- [ ] T037 Test `word_accept_all_changes` in real Word — after mutations, accept all, verify Review pane clears and changes applied
- [ ] T038 Test `word_reject_all_changes` in real Word — after mutations, reject all, verify Review pane clears and original text restored
- [ ] T039 Document real API differences in `specs/003-word-mvp/plan.md` — note any behavior differences between mock and real API. Fix handlers if needed.

**Checkpoint**: All 11 Word tools validated in real Word. Tracked changes confirmed working.

---

## Phase 5: Polish & Cross-Cutting

**Purpose**: Final cleanup and consistency across all hosts

- [ ] T040 [P] Fix any real API differences found during manual testing — update `src/office-addin/src/word-commands.ts` and `src/office-addin/src/word-mock.ts` to match real behavior
- [ ] T041 [P] Update `src/office-addin/src/word-mock.ts` — ensure mock `acceptAllChanges`/`rejectAllChanges` match real API behavior discovered during T037/T038
- [ ] T042 Run full test suite — `dotnet test` + `npm test` + `webpack build` — verify everything green
- [ ] T043 Merge `003-word-mvp` branch to master — `git checkout master && git merge 003-word-mvp`

**Checkpoint**: All tests green. Branch merged. Word MVP complete with tracked changes.

---

## Dependencies & Execution Order

### Phase Dependencies

```
Phase 1 (Spec/Docs) ──► Phase 2 (Backend) ──► Phase 3 (Frontend) ──► Phase 4 (Manual) ──► Phase 5 (Polish)
```

- **Phase 1**: No code dependencies. Do first to align docs.
- **Phase 2**: Depends on Phase 1 (spec defines tools). Backend must exist before frontend handlers.
- **Phase 3**: Depends on Phase 2 (tool definitions). Frontend handlers need backend to accept commands.
- **Phase 4**: Depends on Phase 3 (working code). Manual testing needs built artifacts.
- **Phase 5**: Depends on Phase 4 (real API findings drive fixes).

### Within Each Phase

- **Phase 2**: T004–T006 (tests) can be parallel. T007–T008 sequential. T009 validates.
- **Phase 3**: T010 (mock) first. T011–T016 (tests) parallel after T010. T017 (wrapper) first, then T018–T023 parallel. T024–T026 sequential validation.
- **Phase 4**: T027 first (sideload). T028–T031 (read tools) parallel. T032 first write test. T033–T035 parallel. T036–T038 sequential (build on mutations). T039 final.

### Parallel Opportunities

- T001, T002, T003 can all run in parallel (different files)
- T004, T005, T006 can run in parallel (different test methods, same file but no overlap)
- T011, T012, T013, T014, T015, T016 can all run in parallel (independent test cases)
- T018, T019, T020 can run in parallel (different handler functions)
- T022, T023 can run in parallel (different handler functions)
- T029, T030, T031 can run in parallel (independent read-only tools)
- T033, T034, T035 can run in parallel (independent write tools)

---

## Implementation Strategy

### MVP First (Phase 1–3 only)

1. Complete Phase 1: Spec alignment
2. Complete Phase 2: Backend (3 new tools)
3. Complete Phase 3: Frontend (tracked changes + new handlers)
4. **STOP**: Run automated tests — 82+ C#, 55+ TS, webpack clean
5. Ready for manual validation but code-complete

### Full Delivery (Phase 1–5)

1. Complete MVP (Phase 1–3)
2. Manual Word validation (Phase 4)
3. Fix real API issues (Phase 5)
4. Merge to master

---

## Notes

- PowerPoint has NO tracked changes API — mutations stay direct-write
- Word `changeTrackingMode` requires WordApi 1.4+ (Word 2019+, Microsoft 365)
- Constitution amendment (T002) is non-blocking but should be done before merge
- Manual testing (Phase 4) requires Windows machine with Word installed
- All `[P]` tasks are safe to parallelize — they touch different files or different functions
- TDD order: mock update → tests → implementation within each phase
