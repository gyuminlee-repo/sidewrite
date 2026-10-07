# Sidewrite PR35 and PR36 independent native review

Reviewed 2026-10-07. Repository: https://github.com/gyuminlee-repo/sidewrite

## Conclusion

- PR35 at 4acd91b73dad9ebdc75a8fac5d73dcd519c5caa5: no additional blocking finding in the dependency change or public Office MCP compatibility coverage reviewed.
- PR36 at e01aecc9e34860288bac93023a4a23a09059efa2: one reproducible P2 must-fix in answer follow-through selection. Unchecked answers can disappear from the offered work after a bounded batch completes without edits.
- These are native review assessments. The automatic Codex review gate remains UNKNOWN. No gate result or cache was written. GitHub CI execution remains separately blocked according to the PR records.

## P2 Unchecked answers become unreachable after a candidate limit

Source: https://github.com/gyuminlee-repo/sidewrite/blob/e01aecc9e34860288bac93023a4a23a09059efa2/taskpane/shared/follow-places.js#L235-L255
Freshness check: https://github.com/gyuminlee-repo/sidewrite/blob/e01aecc9e34860288bac93023a4a23a09059efa2/taskpane/shared/follow-places.js#L375-L384

The selector adds answered paragraphs in question order, including already-seen answers unless that paragraph has a live edit. It then caps the list at 20 paragraphs and 40,000 characters. followState tests freshness only on that capped list. A successful turn that finds no edit needed leaves seen paragraphs ahead of the still-unchecked answers. The next check is hidden because every selected answer is already seen.

Reproductions using the production helper:

1. Create 21 answered questions under one comment, each naming a different paragraph. Give the comment no changes. Select its first batch of 20 paragraphs and record those answer IDs as seen, with no changes needed. Question a21 remains unseen, but followState returns offer:null and a line saying 20 places were checked.
2. Two answered paragraphs of 25,001 characters each reproduce the same loss through the character limit. The first is seen; a2 remains unseen; offer is null.

The supplied repro-pr36-answer-starvation.mjs reproduces both cases without a model, Office, or external services. It intentionally asserts the buggy result at the reviewed head. Run it with the path to a separate checkout of e01aecc9e34860288bac93023a4a23a09059efa2 as its argument. This documentation branch is based on main and does not contain PR36 implementation changes. The reproduction checks its initial candidate counts so an unrelated baseline without the feature cannot count as a successful reproduction.

Acceptance conditions for a fix:

- Unseen answers must remain reachable after a bounded batch, including when earlier answers need no edit.
- Retain size bounds and deduplicate a shared paragraph.
- Cover both the paragraph and character limits, repeated batches, an edited already-seen answer, and an answer saved while a turn is running.
- Merely computing freshness from an uncapped list is insufficient if every retry still selects only the same seen first batch. Selection must make progress toward unchecked answers.

No implementation change was made for this finding.

## PR36 coverage

Compared the actual PR delta against main 9b48302fb8de58c2d44b3b71823fc38b6269189c. All 21 changed local review files matched the remote diff blob prefixes. Review covered answer history preservation, same-answer save idempotency, changed answers during a turn, follow-through edit fences, candidate ordering, citation-rule placement, replay limit detection, retry conditions, and save/cleanup exit status.

Independently ran these focused tests on Node 24.19.0:

node --test tests/follow-places.test.mjs tests/follow-through.test.mjs tests/review-questions.test.mjs tests/replay-lib.test.mjs

Result: 99 tests, 98 passed, 0 failed, 1 skipped. The skip was PowerShell syntax because powershell.exe was unavailable. This is not the full JavaScript suite and not a real Word or model replay. The focused run reused existing installed dependencies, including SDK 1.31.0 and proxy-addr 2.0.8, rather than performing a clean install of PR36's older lockfile. It therefore does not validate that lockfile or replace the separate security update.

The existing author evidence remains distinct: 2,507 JavaScript tests, 13 Python tests, four real-Word model-free answer scenarios, one limited inference turn, and a partial manuscript replay of 20 ordinary plus eight answer turns before a usage-limit stop. Those executions were not repeated or independently established by this review.

Citation behavior was inspected at the prompt-rule and test level. The citation guard rule remains at its established index; the new paragraph-location instruction is appended. No new Word citation acceptance run was performed.

## PR35 coverage

PR: https://github.com/gyuminlee-repo/sidewrite/pull/35

All three changed files matched their remote diff blob prefixes. The dependency delta is scoped to @modelcontextprotocol/sdk 1.29.0 to 1.31.0 and proxy-addr 2.0.7 to 2.0.8, plus declared package metadata. Express remains 5.2.1. The existing Claude Agent SDK uses the same installed MCP SDK through deduplication.

Independently ran the four new tests/office-mcp.test.mjs protocol tests against installed SDK 1.31.0 on Node 24.19.0: four passed, zero failed, zero skipped. Coverage includes initialization, Word/Excel host-specific tool discovery, document routing, invalid-argument rejection, review-turn write guards, and bridge errors through the public Client and InMemoryTransport interfaces.

The existing author evidence of 2,474 passing tests and one skip on each of Node 20 and Node 22, plus audit zero, was read from the PR. Those full suites and the package audit were not rerun. No package installation was performed for this review. The four protocol tests use an in-memory transport and a stub bridge; they do not establish actual Office host behavior or safety of every external OAuth configuration.

## Scope and remaining limits

The read-only assessment made no GitHub changes, merge, deployment, installed service change, model API call, paid CLI call, or user-computer action. This documentation-only PR separately preserves the report and synthetic evidence. Review snapshots and synthetic fixtures were isolated. The separate pending legacy installer and Word undo patch is not part of PR35 or PR36 and has a separate review outcome.
