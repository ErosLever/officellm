# Implementation Plan: Word MVP Integration with Tracked Changes

**Branch**: `003-word-mvp` | **Date**: 2026-05-16 | **Spec**: `specs/003-word-mvp/spec.md`

**Input**: Feature specification from `specs/003-word-mvp/spec.md`

## Summary

Complete the Word MVP integration into the unified Office LLM Harness. The backend (8 Word tools), frontend (8 command handlers), mock framework, and tests are already implemented. The remaining work is: (1) add tracked changes support as the default mutation mode for Word, (2) validate all handlers against the real Word JS API via manual testing, (3) update the spec to reflect actual implementation decisions (direct-write replaced by tracked changes).

**Critical constraint**: PowerPoint has NO tracked changes API. For PowerPoint, mutations remain direct-write with undo grouped per `PowerPoint.run()` batch. Only Word gets tracked changes.

## Technical Context

**Language/Version**: TypeScript (ES2022, webpack + vitest) + C# .NET 8

**Primary Dependencies**: Office JS API (WordApi 1.4+ for changeTrackingMode), PowerPoint JS API, ASP.NET Core, xUnit, Vitest

**Storage**: In-memory (CommandStore, InstanceRegistry) + local JSONL audit log

**Testing**: xUnit (C#, 79 tests) + Vitest (TypeScript, 49 tests) + manual Word validation

**Target Platform**: Windows desktop (Office 2019+ / Microsoft 365)

**Project Type**: Office Add-in (unified JS) + MCP server (.NET 8)

**Performance Goals**: <5s for outline extraction on 200-page document (SC-001)

**Constraints**: Word changeTrackingMode requires WordApi 1.4+. PowerPoint has NO tracked changes API.

**Scale/Scope**: 26 tools total (18 PowerPoint + 8 Word + 1 shared). 128 automated tests.

## Constitution Check

| Principle                        | Status                | Notes                                                                                                                                                                                                                                                                  |
| -------------------------------- | --------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| I. Safety-First                  | ✅ PASS               | Tracked changes preserve document integrity. PowerPoint mutations grouped in undo batches.                                                                                                                                                                             |
| II. Local-Only Execution         | ✅ PASS               | MCP server on 127.0.0.1 only. No remote exposure.                                                                                                                                                                                                                      |
| III. User Control & Transparency | ✅ PASS (with update) | Tracked changes are Word-native undo. Constitution says "confirmation gate with before/after diffs" — tracked changes provide this via Word's native Review UI. Need to update constitution to accept tracked changes as an alternative to explicit confirmation gate. |
| IV. Minimal Surface Area         | ✅ PASS               | 8 focused Word tools with typed inputs.                                                                                                                                                                                                                                |
| V. Phased Delivery               | ✅ PASS               | Word MVP is independent phase.                                                                                                                                                                                                                                         |

**Gate result**: PASS with amendment. Constitution section III needs update to recognize tracked changes as a valid confirmation mechanism for Word.

### Complexity Tracking

| Violation                  | Why Needed                                    | Simpler Alternative Rejected Because                            |
| -------------------------- | --------------------------------------------- | --------------------------------------------------------------- |
| Constitution III amendment | Tracked changes provide native Word review UI | Custom confirmation gate duplicates Word's built-in Review mode |

## Project Structure

```text
src/
├── mcp-server/                    # .NET 8 MCP server
│   ├── Tools/
│   │   ├── McpToolEngine.cs       # Tool definitions + dispatch (26 tools)
│   │   └── OfficeTools.cs         # Audit logging only
│   └── ...
├── office-addin/              # Unified Office JS Add-in
│   ├── src/
│   │   ├── app.ts                 # Host dispatch + lifecycle
│   │   ├── communication.ts       # MCP registration/polling/reporting
│   │   ├── powerpoint-commands.ts # 17 PowerPoint handlers
│   │   ├── powerpoint-mock.ts     # PowerPoint mock framework
│   │   ├── word-commands.ts       # 8 Word handlers (NEEDS: tracked changes)
│   │   ├── word-mock.ts           # Word mock framework (NEEDS: changeTracking)
│   │   └── word-commands.test.ts  # 18 Word tests (NEEDS: tracked change tests)
│   ├── manifest.xml               # Unified (Presentation + Document + Workbook + Mailbox)
│   └── package.json               # webpack + vitest
tests/
└── mcp-server.Tests/              # 79 C# tests
specs/
└── 003-word-mvp/                  # This feature
```

## What's Already Done

| Component                | Status  | Details                                 |
| ------------------------ | ------- | --------------------------------------- |
| Backend tool definitions | ✅ Done | 8 Word tools in McpToolEngine.cs        |
| Backend dispatch         | ✅ Done | AddInCommands includes all 8 word\_\*   |
| Frontend handlers        | ✅ Done | 8 handlers in word-commands.ts          |
| Mock framework           | ✅ Done | word-mock.ts with load/sync pattern     |
| Unit tests               | ✅ Done | 18 tests in word-commands.test.ts       |
| C# tests                 | ✅ Done | 79 tests including Word tool validation |
| Manifest                 | ✅ Done | `<Host Name="Document"/>` included      |
| Host dispatch            | ✅ Done | HOST*DISPATCH map routes word*\*        |
| Adversarial review       | ✅ Done | All fixes applied (commit ac49225)      |

## What Needs To Be Done

### Phase 0: Research

#### R1: Word Tracked Changes API

**Decision**: Use `Word.ChangeTrackingMode.trackMineOnly` before mutations.

**API details** (from Microsoft docs):

- `context.document.changeTrackingMode = Word.ChangeTrackingMode.trackMineOnly` — enables tracking
- All subsequent `insertText()`, `delete()`, `insertParagraph()` operations create tracked changes
- `context.document.changeTrackingMode = Word.ChangeTrackingMode.off` — disables
- `range.getReviewedText(Word.ChangeTrackingVersion.original)` — get before-text
- `range.getReviewedText(Word.ChangeTrackingVersion.current)` — get after-text
- Requires WordApi 1.4+ (available in Word 2019+, Microsoft 365)

**Strategy**:

1. Before any Word mutation, enable `trackMineOnly`
2. Perform the mutation
3. Return `{ tracked: true }` in the result
4. User reviews in Word's native Review pane
5. Add a `word_accept_all_changes` and `word_reject_all_changes` tool for batch operations

#### R2: PowerPoint Tracked Changes Equivalent

**Decision**: PowerPoint has NO tracked changes API. Keep direct-write.

**Alternatives considered**:

- Comments-based change log: Too complex, poor UX
- Pre-mutation snapshots: Requires storing full slide state, impractical
- Custom undo groups: PowerPoint undo groups are per-`PowerPoint.run()` batch — already works

**Strategy**: Keep PowerPoint as direct-write. Each mutation tool already groups in a single `PowerPoint.run()` batch, which creates one undo group. User can Ctrl+Z to undo.

#### R3: Spec Updates Needed

The spec (specs/003-word-mvp/spec.md) describes tools that differ from implementation:

| Spec tool                   | Actual tool         | Delta                                  |
| --------------------------- | ------------------- | -------------------------------------- |
| `word_rewrite_selection`    | `word_replace_text` | Need to add tracked changes to replace |
| `word_insert_after_heading` | `word_insert_text`  | Generic insert is sufficient           |
| `word_add_review_comments`  | `word_add_comment`  | Functionally equivalent                |
| Confirmation gate           | Tracked changes     | Word native review instead             |

### Phase 1: Design & Contracts

#### D1: New Word Tools for Tracked Changes

Add 2 new tools:

| Tool                      | Purpose                                | Parameters |
| ------------------------- | -------------------------------------- | ---------- |
| `word_accept_all_changes` | Accept all tracked changes in document | (none)     |
| `word_reject_all_changes` | Reject all tracked changes in document | (none)     |

#### D2: Modified Word Handlers

Each Word mutation handler (`word_replace_text`, `word_insert_text`, `word_delete_paragraph`) needs to:

1. Save current `changeTrackingMode`
2. Set `changeTrackingMode = "TrackMineOnly"`
3. Perform the mutation
4. Restore original `changeTrackingMode`
5. Return `{ tracked: true, ...existingResult }`

#### D3: New Frontend Handler for Change Tracking State

Add `word_get_tracked_changes` tool that returns:

- Current change tracking mode
- Number of pending revisions
- List of revisions with author + type

#### D4: Mock Framework Update

`word-mock.ts` needs:

- `changeTrackingMode` on `MockDocument`
- `acceptAllChanges()` / `rejectAllChanges()` on `MockDocument`
- Mutation handlers check and report tracked state
- New tests for tracked change behavior

#### D5: C# Backend Update

`McpToolEngine.cs` needs:

- 2 new tool definitions: `word_accept_all_changes`, `word_reject_all_changes`
- Optional: `word_get_tracked_changes` for reading tracked change state
- Update `AddInCommands` hash set
- Update tool count tests

### Phase 2: Task Breakdown

#### Task 1: Update spec to match implementation

- Update `specs/003-word-mvp/spec.md` with actual tool names, tracked changes approach
- Remove confirmation gate references (replaced by Word tracked changes)
- Add tracked changes tools to FR section

#### Task 2: Add change tracking to Word handlers

- Modify `word_replace_text`, `word_insert_text`, `word_delete_paragraph` in `word-commands.ts`
- Wrap mutations with changeTrackingMode enable/restore
- Add `word_accept_all_changes`, `word_reject_all_changes`, `word_get_tracked_changes` handlers
- Return `{ tracked: true }` in mutation results

#### Task 3: Update mock framework

- Add `changeTrackingMode` to `MockDocumentData`
- Track mutations in `changeLog` array
- Implement `acceptAllChanges()` / `rejectAllChanges()`
- Ensure mock handlers respect tracking mode

#### Task 4: Update tests

- Add tests for tracked change behavior (mode save/restore, result flag)
- Add tests for accept/reject all
- Add tests for get_tracked_changes
- Verify write tests check `tracked: true` in results

#### Task 5: Backend updates

- Add 3 new tool definitions to `McpToolEngine.cs`
- Update `AddInCommands` set
- Update C# test assertions (tool count 26→29)

#### Task 6: Manual Word validation

- Load add-in in real Word
- Test each of the 11 Word tools
- Verify tracked changes appear in Word Review pane
- Document any API differences from mock behavior

#### Task 7: Constitution amendment

- Update constitution section III to recognize tracked changes as valid confirmation mechanism
- Add note about PowerPoint using undo groups instead

#### Task 8: Update AGENTS.md

- Add Word JS API rules (like PowerPoint rules already documented)
- Document changeTrackingMode pattern
- Update tool count references

## Contracts

### Word Mutation Tool Response Envelope

```typescript
interface WordMutationResult {
  // Existing fields (tool-specific)
  paragraphIndex?: number;
  oldText?: string;
  newText?: string;
  replaced?: boolean;
  inserted?: boolean;
  deleted?: boolean;

  // New tracked changes field
  tracked: boolean; // Always true for Word mutations
}
```

### New Tool Schemas

#### word_get_tracked_changes

```json
{
  "name": "word_get_tracked_changes",
  "inputSchema": {
    "type": "object",
    "properties": {},
    "required": []
  }
}
```

#### word_accept_all_changes

```json
{
  "name": "word_accept_all_changes",
  "inputSchema": {
    "type": "object",
    "properties": {},
    "required": ["instanceId"]
  }
}
```

#### word_reject_all_changes

```json
{
  "name": "word_reject_all_changes",
  "inputSchema": {
    "type": "object",
    "properties": {},
    "required": ["instanceId"]
  }
}
```
