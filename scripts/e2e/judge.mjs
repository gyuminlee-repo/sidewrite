// Pure helpers of the acceptance runner: the Word COM plan built from
// revision.json, the expected pane texts of each step (wording of the
// stage-2 acceptance procedure), and the path rewrite that runs a
// feature_list verify_command against this run's files.

import { formatRationaleComment } from "../../daemon/review-tools.mjs";

// The author's handling of the proposed changes in the acceptance procedure:
// R1.1 accepted, R1.2 rejected, R1.3 first change accepted then one inserted
// word edited, everything else accepted.
export const DEFAULT_DECISIONS = {
  "R1.1": "accept",
  "R1.2": "reject",
  "R1.3": ["accept_edit", "accept"],
};

// One inserted word the author edits by hand, and what it becomes.
export const WORD_EDITS = {
  reading: "measuring",
  measuring: "recording",
  monitoring: "tracking",
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
