// Pure helpers of the acceptance runner: the Word COM plan built from
// revision.json, the expected pane texts of each step (wording of the
// stage-2 acceptance procedure), and the path rewrite that runs a
// feature_list verify_command against this run's files.

import { formatRationaleComment } from "../../daemon/review-tools.mjs";
import { normalizeParaText, placedRuns } from "../../daemon/revision-classify.mjs";

// The author's handling of the proposed changes in the acceptance procedure:
// R1.1 accepted, R1.2 rejected, R1.3 first change accepted then one inserted
// word edited, everything else accepted. R1.4 (the deletion of the flask
// sentence) is rejected, with a tracked insertion of the author pending in
// the same paragraph (siblingMatch).
export const DEFAULT_DECISIONS = {
  "R1.1": "accept",
  "R1.2": "reject",
  "R1.3": ["accept_edit", "accept"],
  "R1.4": "reject",
};

// One inserted word the author edits by hand, and what it becomes.
export const WORD_EDITS = {
  reading: "measuring",
  measuring: "recording",
  monitoring: "tracking",
  monitored: "tracked",
  followed: "tracked",
};

function decisionFor(decisions, ref, i) {
  const d = decisions[ref] ?? "accept";
  if (!Array.isArray(d)) return d;
  return d[Math.min(i, d.length - 1)];
}

// ops for word-com.ps1 apply. A change without a rationale comment cannot be
// found in Word and is listed in `skipped`.
export function buildComPlan(revision, decisions = DEFAULT_DECISIONS) {
  const ops = [];
  const skipped = [];
  for (const c of revision?.comments ?? []) {
    c.changes.forEach((ch, i) => {
      if (!ch.word_comment_id || ch.comment_missing || ch.comment_cleaned_at) {
        skipped.push(ch.id);
        return;
      }
      ops.push({
        change_id: ch.id,
        ref: c.ref,
        op: decisionFor(decisions, c.ref, i),
        kind: ch.kind,
        find: ch.find,
        text: ch.text,
        comment_text: formatRationaleComment(c.ref, ch.rationale, { lost: ch.lost_facts ?? [] }),
        replacements: WORD_EDITS,
      });
    });
  }
  return { ops, skipped };
}

// A replace whose find and text share a first or last word did not cover only
// the changed words (L3). Inserts and deletes are minimal by construction.
export function minimalFind(ch) {
  if (ch.kind !== "replace") return true;
  const w = (s) => String(s).trim().split(/\s+/).filter(Boolean);
  const a = w(ch.find);
  const b = w(ch.text);
  if (!a.length || !b.length) return true;
  return a[0] !== b[0] && a[a.length - 1] !== b[b.length - 1];
}

// Replace each `from` prefix with its `to`, longest first so a folder and a
// file inside it do not clash.
export function rewriteCommand(cmd, map) {
  const pairs = Object.entries(map).sort((x, y) => y[0].length - x[0].length);
  let out = "";
  for (let i = 0; i < cmd.length; ) {
    const hit = pairs.find(([from]) => from && cmd.startsWith(from, i));
    if (hit) {
      out += hit[1];
      i += hit[0].length;
    } else {
      out += cmd[i++];
    }
  }
  return out;
}

// What to do with the Word process the runner started, from the answer of
// word-com.ps1 quit. action: "none" | "kill" | "leave"; fail: the run fails.
//   kill   only when quit ran (so every open document was in the run folder
//          and was closed unsaved) and the process is still there.
//   leave  quit refused because a document outside the run folder is open
//          in that Word, or the open documents are unknown (the COM call
//          failed): unsaved work of the user may be in it, Word stays up.
export function cleanupDecision({ pid, preexisting = [], quit, runDir, alive = true }) {
  if (!pid || preexisting.includes(pid)) return { action: "none", fail: false, reason: "" };
  const dir =
    String(runDir ?? "")
      .replace(/\\+$/, "")
      .toLowerCase() + "\\";
  const docs = Array.isArray(quit?.documents) ? quit.documents : null;
  const outside = (docs ?? []).filter((d) => !String(d).toLowerCase().startsWith(dir));
  if (quit?.refused || outside.length) {
    return {
      action: "leave",
      fail: true,
      reason: `Word ${pid} left running: documents outside the run folder are open (${(quit?.outside ?? outside).join(", ")})`,
    };
  }
  if (quit?.ok === true && docs) {
    if (quit.exited === true) return { action: "none", fail: false, reason: "" };
    return { action: "kill", fail: false, reason: `Word ${pid} still alive after quit` };
  }
  if (!alive) return { action: "none", fail: false, reason: "" };
  return {
    action: "leave",
    fail: true,
    reason: `Word ${pid} left running: its open documents are unknown (${quit?.error ?? "no answer from quit"})`,
  };
}

// Features whose verify_command can pass only when something happened that
// the run cannot force. trigger: the daemon log line of that event.
export const CONDITIONS = {
  // A turn that ended on a Claude usage limit.
  "S2-27": { trigger: /^\[agent\] turn complete \(usage_limit\)/m },
};
const TURN_LOGGED = /^\[agent\] turn complete \(/m;

// Exit code of the run. features: [{ id, exit }] of the verify commands;
// conditional: ids given with --conditional; logs: text of the daemon logs
// this run wrote. A conditional feature that failed is
// SKIPPED-CONDITION-NOT-MET only when its trigger is provably absent: the
// logs show finished turns and none with the trigger. Otherwise it counts.
export function exitDecision({ failed, rows, features, featuresError, conditional = [], logs }) {
  const reasons = [];
  if (failed) reasons.push("the run stopped on an error");
  for (const r of rows ?? []) if (r.result === "FAIL") reasons.push(`step failed: ${r.step}`);
  if (featuresError) reasons.push(`verify commands did not run: ${featuresError}`);
  const text = typeof logs === "string" ? logs : "";
  const out = (features ?? []).map((f) => {
    if (f.exit === 0) return { ...f, status: "PASS" };
    const cond = conditional.includes(f.id) ? CONDITIONS[f.id] : null;
    if (cond && TURN_LOGGED.test(text) && !cond.trigger.test(text)) {
      return { ...f, status: "SKIPPED-CONDITION-NOT-MET" };
    }
    reasons.push(`verify command failed: ${f.id} (exit ${f.exit})`);
    return { ...f, status: "FAIL" };
  });
  return { code: reasons.length ? 1 : 0, reasons, features: out };
}

const has = (s, sub) => String(s ?? "").includes(sub);

// R1.1 all Accepted, R1.2 all Rejected, R1.3 first change Edited by you.
export function badgesMatch(view) {
  const badges = (ref) =>
    view.comments.find((c) => c.ref === ref)?.changes.map((c) => c.badge) ?? [];
  const a = badges("R1.1");
  const b = badges("R1.2");
  const c = badges("R1.3");
  return {
    pass:
      a.length > 0 &&
      a.every((x) => x === "Accepted") &&
      b.length > 0 &&
      b.every((x) => x === "Rejected") &&
      c[0] === "Edited by you",
    got: `R1.1 ${a.join("/")}; R1.2 ${b.join("/")}; R1.3 ${c.join("/")}`,
  };
}

// word-com.ps1 Norm: how the script prints a comment text.
export const norm = (s) =>
  String(s ?? "")
    .replace(/[\r\n\u000b\u0007]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

// C1: after the author accepts or rejects a change and presses Refresh, its
// rationale comment is resolved. revision: revision.json after that refresh;
// word: word-com.ps1 state (comments with `done` = Comment.Done); message:
// the pane's refresh message. Edited and pending changes keep an open comment.
export function resolvedMatch(revision, word, message = "") {
  const byText = new Map((word?.comments ?? []).map((c) => [norm(c.text), c]));
  const seen = [];
  for (const c of revision?.comments ?? []) {
    for (const ch of c.changes ?? []) {
      if (ch.state !== "accepted" && ch.state !== "rejected") continue;
      if (!ch.word_comment_id || ch.comment_missing || ch.comment_cleaned_at) continue;
      const text = formatRationaleComment(c.ref, ch.rationale, { lost: ch.lost_facts ?? [] });
      const com = byText.get(norm(text));
      seen.push({ id: ch.id, done: com ? com.done === true : null });
    }
  }
  const refused = has(message, "not marked resolved");
  return {
    pass: seen.length > 0 && seen.every((x) => x.done === true) && !refused,
    expect:
      "every accepted or rejected change has its rationale comment Done in Word; no 'not marked resolved' note",
    got:
      (seen
        .map((x) => `${x.id} ${x.done === null ? "no comment" : x.done ? "done" : "open"}`)
        .join("; ") || "no accepted or rejected change with a comment") +
      (refused ? "; pane: not marked resolved" : ""),
  };
}

// A change handled by the author while ANOTHER tracked change is pending in
// its paragraph: it must show its own state, not Pending (the paragraph's
// revision mark total used to decide), and its rationale comment must be
// resolved. Shown on the first change of `ref`, a pure deletion that was
// rejected or a pure insertion that was accepted; `hand` is the text of the
// author's own tracked insertion in that paragraph, still pending in Word.
// revision: revision.json after the refresh; view: the pane; word:
// word-com.ps1 state.
export function siblingMatch(revision, view, word, { ref, hand }) {
  const ch = revision?.comments?.find((c) => c.ref === ref)?.changes?.[0];
  const badge = view?.comments?.find((c) => c.ref === ref)?.changes?.[0]?.badge;
  const want =
    ch?.kind === "delete"
      ? ["rejected", "Rejected"]
      : ch?.kind === "insert"
        ? ["accepted", "Accepted"]
        : null;
  const sibling = (word?.revisions ?? []).some((r) => r.type === 1 && has(r.text, hand.trim()));
  const text = ch
    ? norm(formatRationaleComment(ref, ch.rationale, { lost: ch.lost_facts ?? [] }))
    : "";
  const com = (word?.comments ?? []).find((c) => norm(c.text) === text);
  return {
    pass: !!want && sibling && ch.state === want[0] && badge === want[1] && com?.done === true,
    expect: `${ref} first change (a pure deletion rejected, or a pure insertion accepted) reads its own state with the author's tracked "${hand.trim()}" pending in the paragraph, and its rationale comment is resolved`,
    got: `kind ${ch?.kind ?? "-"}; state ${ch?.state ?? "-"}; badge ${badge ?? "-"}; sibling revision pending ${sibling}; comment ${com ? (com.done ? "done" : "open") : "not found"}`,
  };
}

// The document of make_tracked_special.py as the pane's snapshot shows it in
// Word: every tracked run read from the paragraph OOXML (trackedRuns) must be
// found in the paragraph text Word reports, with no revision mark left over
// (placedRuns of the daemon, the function the classifier uses). paragraphs:
// snapshot paragraphs; expected: how many paragraphs hold a tracked run.
export function trackedPlaced(paragraphs, expected) {
  const tracked = (paragraphs ?? []).filter(
    (p) => (p.tracked?.length ?? 0) > 0 || (p.w_ins ?? 0) + (p.w_del ?? 0) > 0,
  );
  const bad = [];
  for (const p of tracked) {
    const placed = placedRuns(p, normalizeParaText(p.text));
    if (
      !placed ||
      placed.short.ins ||
      placed.short.del ||
      placed.runs.length !== p.tracked.length
    ) {
      const codes = (s) => [...String(s)].map((c) => c.codePointAt(0).toString(16)).join(" ");
      bad.push(
        `paragraph ${p.index}: text [${codes(p.text)}] runs ${JSON.stringify((p.tracked ?? []).map((r) => [r.kind, codes(r.text)]))}`,
      );
    }
  }
  return {
    pass: tracked.length > 0 && tracked.length === expected && bad.length === 0,
    expect: `${expected} paragraphs with a tracked run, each run found in the paragraph text Word reports`,
    got: bad.length ? bad.join("; ") : `${tracked.length} paragraphs with tracked runs, all placed`,
  };
}

// Judgments of one pane step from the runner's record of it. Each returns
// { pass, expect, got } where expect quotes the acceptance procedure.
export const JUDGES = {
  "open-review": (r, ctx) => ({
    pass: r.file === `Review file: ${ctx.revisionPath}`,
    expect: `Review file: ${ctx.revisionPath}`,
    got: r.file,
  }),
  split: (r) => ({
    pass:
      /^Split by /.test(r.message) && r.draft_refs.filter((x) => /^R1\.\d$/.test(x)).length === 5,
    expect: "Split by regex: 5 comment(s) (R1.1-R1.5 after the non-reviewer rows go)",
    got: `${r.message} refs=${r.draft_refs.join(",")}`,
  }),
  "edit-draft": (r) => {
    const refs = r.comments.map((c) => c.ref).join(",");
    const r13 = r.comments.find((c) => c.ref === "R1.3")?.reqs_line ?? "";
    return {
      pass:
        refs === "R1.1,R1.2,R1.3,R1.4,R1.5" &&
        has(
          r13,
          "Requirements: q1 say how growth was followed; q2 describe the synthetic carbon source as a novel substrate",
        ),
      expect: "cards R1.1-R1.5; R1.3 Requirements: q1 ...; q2 ...",
      got: `${refs} | ${r13}`,
    };
  },
  "ledger-add": (r) => ({
    pass:
      r.ledger.length === 3 &&
      has(r.ledger[0], "L1 instruction") &&
      has(r.ledger[1], "L2 forbid") &&
      has(r.ledger[2], "L3 instruction"),
    expect: "L1 instruction, L2 forbid, L3 instruction",
    got: r.ledger.join(" | "),
  }),
  "check-all": (r) => {
    const n = (id) => r.violations.filter((v) => v.startsWith(`${id}:`)).length;
    const want = ["R1.3 q1 has no change", "R1.3 q2 has no change"];
    return {
      pass:
        n("L1") === 2 &&
        n("L2") === 1 &&
        n("L3") === 0 &&
        want.every((w) => r.decisions.some((d) => d.startsWith(w))),
      expect: "violations L1 2, L2 1, L3 0; waiting: R1.3 q1, R1.3 q2",
      got: `L1 ${n("L1")}, L2 ${n("L2")}, L3 ${n("L3")}; ${r.decisions.join(" | ")}`,
    };
  },
  preview: (r) => ({
    pass: ["[L1]", "[L2]", "[L3]"].every((t) => has(r.prompt, t)),
    expect: "prompt shows [L1], [L2], [L3]",
    got: ["[L1]", "[L2]", "[L3]"].filter((t) => has(r.prompt, t)).join(" "),
  }),
  propose: (r, ctx) => {
    const re = new RegExp(
      `^${ctx.ref.replace(".", "\\.")} running \\d+:\\d\\d\\. \\d+ of \\d+ comments not done\\.$`,
    );
    const ticks = r.progress_samples.filter((s) => re.test(s));
    return {
      pass: !r.message_error && ticks.length >= 2 && ctx.turnCompleted === true,
      expect: `${ctx.ref} running m:ss. n of 5 comments not done. (changing), turn complete in the log, no error`,
      got: `${ticks.slice(0, 2).join(" -> ")} | turn_complete=${ctx.turnCompleted} | ${r.message}`,
    };
  },
  refresh: (r) => {
    const m = /Changes (\d+), expected tool comments (\d+), found (\d+)/.exec(r.message);
    return {
      pass: !!m && m[1] === m[2] && m[2] === m[3] && !has(r.message, "MISMATCH"),
      expect: "Changes N, expected tool comments N, found N. (no MISMATCH)",
      got: r.message,
    };
  },
  // A document in compatibility mode 14: Word refuses Comment.resolved. The
  // refresh still shows its counts, and one note for the handled changes.
  "refresh-mode14": (r) => {
    const m =
      /^Accepted (\d+), rejected (\d+), .*Changes (\d+), expected tool comments (\d+), found (\d+)/.exec(
        r.message,
      );
    const n = /(\d+) rationale comment\(s\) not marked resolved \(Word: /.exec(r.message);
    const handled = m ? Number(m[1]) + Number(m[2]) : 0;
    return {
      pass:
        !!m &&
        !!n &&
        handled > 0 &&
        Number(n[1]) === handled &&
        m[3] === m[4] &&
        m[4] === m[5] &&
        !has(r.message, "MISMATCH"),
      expect:
        "Accepted a, rejected r ... Changes N, expected tool comments N, found N. (a+r) rationale comment(s) not marked resolved (Word: ...)",
      got: r.message,
    };
  },
  // The badge part is what the procedure asks a person to confirm
  // (s2-list-matches); the outside-change part is L4.
  "refresh-after-user": (r) => {
    const b = badgesMatch(r);
    const at = r.external.findIndex((x) => has(x, "Lost: 0.82; 0.91"));
    return {
      pass: b.pass && has(r.message, "Outside changes:") && at >= 0,
      expect:
        "R1.1 Accepted, R1.2 Rejected, R1.3 first Edited by you; Outside changes: n; an outside item Lost: 0.82; 0.91",
      got: `${b.got}; outside item ${at + 1} of ${r.external.length}: ${r.external[at] ?? "-"}`,
    };
  },
  "clean-tool": (r) => ({
    pass: /Tool comments \d+ -> 0, your comments 1 -> 1 \(unchanged\)/.test(r.message),
    expect: "Tool comments N -> 0, your comments 1 -> 1 (unchanged)",
    got: r.message,
  }),
  "delete-all": (r) => ({
    pass:
      has(r.offer, "Back up and delete 2 comments (tool 0, yours 2)") &&
      has(r.message, "Backed up 2, deleted 2 of 2") &&
      !has(r.message, "MISMATCH"),
    expect: "Back up and delete 2 comments (tool 0, yours 2) then Backed up 2, deleted 2 of 2",
    got: `${r.offer} | ${r.message}`,
  }),
};

export function judge(name, record, ctx = {}) {
  const fn = JUDGES[name];
  if (!fn) return { pass: record.ok !== false, expect: "-", got: "-" };
  if (record.ok === false) return { pass: false, expect: "-", got: record.error };
  return fn(record, ctx);
}

// Fixed-width table of { step, pass, got } rows.
export function table(rows) {
  const w = Math.max(4, ...rows.map((r) => r.step.length));
  const line = (a, b, c) => `${a.padEnd(w)}  ${b.padEnd(7)}  ${c}`;
  return [
    line("step", "result", "detail"),
    ...rows.map((r) => line(r.step, r.result, String(r.got ?? "").slice(0, 160))),
  ].join("\n");
}
