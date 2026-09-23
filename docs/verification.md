# Verification

## User-centered document tools — 2026-09-23

Work performed on `feat/user-centered-document-tools`, based on `9a4376e`, in the isolated `G:\test\WordAgent` checkout.

- Baseline before changes: 52 tests across 6 files passed.
- Updated suite: `npm test -- --maxWorkers=2` passed 101 tests across 8 files, adding 49 regression checks.
- `npm run build` passed TypeScript checking, Vite production bundling and tsup backend bundling.
- Changed code and documentation were formatted with the installed Prettier. No dependency or lockfile changes were needed.
- Coverage includes paginated literal search, scoped/count-checked replacement plans, protected targets, exact occurrence replacement and rollback in real ProseMirror, preserved surrounding marks/links, merged font attributes, strict rich-text/OOXML revision conflicts, mocked Word range preflight, bounded multi-turn tool calls, cancellation, malformed calls, selection scope, and cumulative SSE frame limits.
- React/jsdom checks exercise local search, preview staging, paragraph navigation, recipe selection, disabled controls, protected-match errors and complete change-card diffs. They assert local actions do not call `fetch`; they are not screenshot or native WebView2 acceptance.
- Model settings now have a separate Agent capability probe. Automated API tests verify the normal test sends no tools, the Agent probe sends no document context, forces the no-op probe function, recognizes a valid tool call, and redacts upstream error bodies.
- No paid model endpoint, personal settings, certificate trust, Word sideload or service restart was used for feature validation. Release publication uses separate transient credentials and does not store them in the repository. Node emitted environment proxy-agent experimental warnings; these were not test/build failures.

Native Word search semantics, formatting inheritance, mixed-format matches, protected documents and collaborative changes still need validation in a disposable real Word document. Strict OOXML guards can conservatively reject nonsemantic host changes; they are deliberately not bypassed by equal paragraph text. The toolbox does not add web search or factual/citation verification.

## Historical verification record

The sections below are retained from the original repository. They describe earlier acceptance work, not checks repeated for the document-tools change.

This record separates automated checks, browser acceptance, and native Word checks. A local deterministic provider was used only for acceptance testing; it is not an AI model.

## Automated Checks

- `npm test`: 49 tests in 5 files passed.
- `npm run build`: TypeScript, Vite frontend, and bundled Node backend passed.
- `npm run manifest:validate`: the Word add-in XML manifest passed validation.
- `npm audit`: no reported vulnerabilities at verification time. Audit results can change as advisories are published.
- Development HTTP checks: the workspace and task pane returned 200; private settings and certificate paths, including raw and encoded variants, returned 403.
- The frontend JavaScript bundle was checked against the configured API key values; no matching secrets were found. Values were not printed.
- Production entry point was started temporarily on HTTPS port 3003: workspace, task pane, API health, icons, and JavaScript assets passed. Missing client headers and external origins returned 403; private files returned 404. The temporary process was stopped after verification.

## Browser Acceptance

Verified with the real Tiptap editor and an explicitly labeled local test provider:

- Add a model through settings and select it for a conversation.
- Show the document-sharing confirmation before sending context.
- Stream an editing proposal without changing the original document in review mode.
- Apply a title change, then restore the original using rollback.
- Insert a real table in live mode and roll back the batch.
- Stop a slow request without applying an incomplete proposal.
- Clear test conversations and remove the temporary test model.
- Inspect desktop, mobile, and 320px narrow-pane layouts in light and dark themes.

The original four model profiles were preserved. The temporary provider was stopped and removed from settings. No mock provider is enabled in the normal application.

## Native Word Acceptance

The v2 manifest was sideloaded into the installed Microsoft Word desktop application. The task pane loaded, connected to the local service, read a disposable document, and displayed its real paragraph and heading outline.

Native Word AI writes were not verified end to end. The existing GPT profile failed authentication; the HY3 profile returned a missing endpoint or model error. The two migrated remote HTTP profiles are blocked until their endpoints are updated to HTTPS. A successful settings connection test alone also does not prove support for editing tool calls.

The Office adapter unit tests use a mocked host. They check the captured selection, exact revision checks, paragraph insertion order, rollback, and preservation of later manual formatting; they do not replace native Word acceptance.

## Remaining Acceptance

With a valid HTTPS model endpoint that supports streaming Chat Completions and tool calls, use a disposable Word document to verify:

1. Select a phrase, generate a selection edit, move the cursor elsewhere, and apply to the original selection.
2. Review and apply paragraph replacement, insertion, formatting, deletion, and a simple table.
3. Enable live application, check the completed batch, and roll it back.
4. Change the document manually while a proposal is pending and confirm that stale edits are rejected.
5. Save, close, and reopen the Word document to verify its content and formatting.

Browser DOCX import is deliberately not a lossless Word renderer. Protected content, complex layout, collaborative races, and partial Office synchronization failures remain subject to the limitations in the README.
