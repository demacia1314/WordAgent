# Verification

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
