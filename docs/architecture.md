# Architecture

## Runtime

Word hosts `taskpane.html` inside its native task pane. `Office.onReady` and the WordApi 1.3 requirement gate startup. The root page instead mounts a Tiptap editor. Both environments render the same React assistant and use the `DocumentAdapter` interface.

The Vite development server proxies `/api` to a loopback-only Express server. Its file denylist excludes private settings, environment files, certificates, keys, and Git metadata. Production serves only the built frontend and API from one HTTPS origin. No CORS allowlist is exposed. API requests require a non-simple client header and a loopback hostname; external origins are rejected. This protects against ordinary cross-site browser requests and DNS rebinding, but is not authentication against another process on the same computer. Do not expose this single-user service on a network.

Model profiles are loaded from an ignored local JSON file, updated via atomic rename, with writes serialized in-process. Public settings contain `hasKey`, not secrets. Remote HTTP profiles migrated from older versions are visible for repair but cannot make model requests.

## Request Lifecycle

1. The user authorizes sharing document context with the selected endpoint.
2. The client captures a document ID, revision, paragraph IDs, exact original texts, and an optional selection. Word captures a tracked Range for selection edits.
3. The client submits a bounded conversation and either document or selection context. Selection-only requests do not send unselected paragraphs.
4. The server sends a system instruction, conversation, untrusted document snapshot, and latest user request to a Chat Completions-compatible endpoint.
5. Public text streams over SSE. Provider reasoning fields and optional `<think>` envelopes are not displayed. Keepalive comments prevent idle intermediary timeouts; disconnect cancels the upstream fetch.
6. Agent document mode allows up to four sequential read-only `inspect_document` / `search_document` calls, with correlated assistant/tool messages, then a final answer or one `edit_document` / `replace_in_document` proposal. One 180-second deadline and bounded output/result sizes cover the turn. Schema, exact text, scope, counts and shape are checked before any proposal. Ask advertises no tools; selection mode exposes only a selection-edit schema. Tool availability is enforced again server-side. Truncated, malformed, unknown or over-budget calls cannot produce a proposal.
7. The user can choose review or live mode (live remains the default). Local toolbox replacements always stage a pending preview without contacting a model. Read-only tool completion is displayed separately from edit application. A proposal is not a success receipt.
8. The adapter validates document identity and revision again. ProseMirror applies one transaction; Word queues bottom-to-top paragraph operations and synchronizes through Office.js.
9. Only a successful adapter response marks a batch applied. Revision-guarded checkpoints enable rollback during the current open session.

## Operation Contract

`replace`, `insert`, `delete`, `format`, and `table` require `paragraphId` and exact `expectedText`. Paragraph IDs are snapshot-relative indices, not permanent document IDs. The full revision guard prevents index drift from applying stale edits. Multiple operations on the same paragraph in a batch are rejected. Table cells and nested browser paragraphs are protected from paragraph-wide operations.

`replace_text` adds count-guarded, case-sensitive literal substring replacement, optionally for one 1-based occurrence. The compiler requires a total count, refuses protected matches rather than skipping them, and caps batches at 60 paragraphs and 1,000 replacements. Browser replacements use original UTF-16 offsets in one bottom-to-top transaction; paragraphs with inline non-text nodes are protected. Replacement text inherits the first matched text run's marks, while surrounding runs remain untouched. Word preloads all search ranges with explicit literal options, escapes caret codes, checks exact counts/text and body revision, then queues writes. No sync is performed per occurrence. Word host behavior still requires native acceptance.

`replace_selection` is the only operation permitted for selection scope. Browser selections use immutable captured ProseMirror positions and a document fingerprint. Word selections use an explicitly tracked original Range. Reloaded Word selection proposals are rejected because their original Range no longer exists.

Revisions strictly fingerprint complete browser JSON or Word body OOXML, not only plain text. The former paragraph-equality fallback is removed: formatting or OOXML-only changes invalidate stale proposals even when snapshot text fields are equal. This can conservatively reject nonsemantic host OOXML changes; regeneration is safer than bypassing the guard. Fingerprints detect ordinary editing conflicts; they are not cryptographic authentication.

## Recovery Semantics

- Cancel stops generation; a batch already being synchronized may finish and remains explicitly visible.
- Stream errors retain any already visible answer and expose retry. Repeated upstream requests are never silently replayed after partial output.
- Pending proposals block further model turns for the same document until resolved.
- Manual changes after application block direct rollback. Use editor-native undo for later modifications first.
- Word does not provide a multi-operation transactional guarantee. A failing `context.sync` can partially execute; the UI reports this and directs the user to native Word undo. There is no unsafe automatic whole-body rollback over potentially concurrent user changes.
- Body OOXML rollback preserves the captured body package; headers, footers and other document regions are outside this editor's scope.
- Concurrent collaborators can race the last client check; Word version history is still required for important documents.

## Browser Document Conversion

Import validates ZIP directory sizes before Mammoth conversion and disables external-file access. DOMPurify sanitizes HTML before Tiptap insertion. Images and active embedded content are excluded. Export constructs an A4 document via `docx`, with actual heading styles and explicit fixed-width table columns and cell margins. It is not a round-trip renderer for all Office features.

## Verification Boundaries

Automated tests exercise real ProseMirror operations, mocked provider SSE streams, API request handling, configuration isolation, and DOCX conversion. Mocked Office tests verify adapter calls and conflict rules, not Microsoft Word's actual rendering or behavior. Manifest validation checks the XML contract, not Office.js runtime integration. Native task-pane loading and document reads were also checked in the installed Word application. See `verification.md` for completed checks and the remaining model-dependent Word acceptance.
