"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.lastLookAheadRoles = void 0;
exports.classify = classify;
exports.matrixFromAsta = matrixFromAsta;
exports.approvalsFromAsta = approvalsFromAsta;
exports.overviewRows = overviewRows;
exports.matchProjectKey = matchProjectKey;
exports.lookAhead = lookAhead;
exports.lookAheadRoleRows = lookAheadRoleRows;
const roles_1 = require("./roles");
const asta_excel_1 = require("../parsers/asta_excel");
const match_1 = require("../parsers/match");
const overview_select_1 = require("./overview_select");
const STAGES = ['Drawings/Documents Collection', 'Application', 'Fee Payment', 'Issued'];
const iso = (s) => (s ? new Date(s + 'T00:00:00Z') : null);
// ---- Red Flag / Tracker from Asta (progress-based) ----
function classify(sch, keywords, byProgress = false) {
    const kws = keywords.map((k) => k.toLowerCase()).filter(Boolean);
    const matched = sch.activities.filter((a) => kws.some((k) => a.name.toLowerCase().includes(k) || (a.dept && a.dept.toLowerCase().includes(k))));
    if (!matched.length)
        return 'na';
    if (byProgress) {
        if (matched.every((a) => a.pct >= 100))
            return 'good';
        if (matched.every((a) => a.pct === 0))
            return 'red';
        return 'amber';
    }
    const rd = iso(sch.dataDate);
    const done = (a) => a.pct >= 100 || a.status === 'Completed';
    const overdue = (a) => a.finish && rd && iso(a.finish) < rd && !done(a);
    if (matched.every(done))
        return 'good';
    if (matched.some((a) => (a.critical || overdue(a)) && !done(a)))
        return 'red';
    if (matched.some((a) => a.pct > 0 || a.status === 'In Progress'))
        return 'amber';
    return 'good';
}
function matrixFromAsta(sch, cfg, code, byProgress) {
    const subs = cfg.red_flag_subs;
    const macro = cfg.red_flag_macro;
    const kw = cfg.column_keywords || {};
    const cells = subs.map((n) => classify(sch, kw[n] || [n.toLowerCase()], byProgress));
    return { project: sch.projectName, code, macro, subs, cells };
}
function approvalsFromAsta(sch, names) {
    return names.map((nm, i) => {
        const core = nm.replace(/\[.*?\]/g, '').trim().toLowerCase();
        const nkey = nm.toLowerCase().replace(/[^a-z0-9]/g, '');
        const matched = sch.activities.filter((a) => {
            const an = a.name.toLowerCase();
            if (a.approval && a.approval.toLowerCase().replace(/[^a-z0-9]/g, '') === nkey)
                return true;
            return core && an.includes(core);
        });
        const terminal = matched.filter((a) => /issue|sanction|certificate|approv/.test(a.name.toLowerCase()));
        let stage = -1;
        for (const a of matched)
            if (a.pct >= 100 || a.status === 'Completed') {
                const idx = STAGES.indexOf(a.approvalStatus);
                if (idx >= 0)
                    stage = Math.max(stage, idx);
                if (/issue|sanction|certificate/.test(a.name.toLowerCase()))
                    stage = Math.max(stage, 3);
            }
        if (terminal.length && terminal.every((a) => a.pct >= 100))
            stage = 3;
        const src = terminal.length ? terminal : matched;
        const maxDate = (f) => src.map(f).filter(Boolean).sort().slice(-1)[0] || null;
        return {
            sl: i + 1, approval: nm,
            drawings: stage >= 0, application: stage >= 1, fee: stage >= 2, issued: stage >= 3,
            base: maxDate((a) => a.baselineFinish), proj: maxDate((a) => a.finish),
        };
    });
}
// ---- Project Overview ----
// Rows come from the RULE-BASED selector (overview_select.ts), which classifies each WBS node by
// what it is — major section, deliverable, execution/approval package, mobilisation, sales — so the
// overview holds up regardless of how deeply a given project nests its WBS. ASTA already rolls up
// %, dates and baseline onto every summary row, so those values are used directly.
function overviewRows(sch) {
    const decisions = (0, overview_select_1.selectOverview)(sch);
    const S = (0, overview_select_1.sectionDepth)(Object.values(sch.wbs), sch.projectName);
    const day = 864e5;
    return decisions.filter((d) => d.selected).map((d) => {
        const w = d.node;
        const finish = w.finish ?? null, baseFinish = w.baselineFinish ?? null;
        const pct = Math.round(w.pct ?? 0);
        const delayDays = (finish && baseFinish)
            ? Math.max(0, Math.round((new Date(finish + 'T00:00:00Z').getTime() - new Date(baseFinish + 'T00:00:00Z').getTime()) / day))
            : 0;
        return {
            name: w.name, indent: Math.max(0, w.depth - S), summary: true, major: w.depth === S,
            category: d.category,
            start: w.start ?? null, finish,
            baseStart: w.baselineStart ?? w.start ?? null, baseFinish,
            pct, critical: false, milestone: !!w.milestone, duration: w.duration ?? 0,
            delayDays, done: pct >= 100 && !!finish,
        };
    });
}
function matchProjectKey(cfg, projectName) {
    let best = null, bestScore = 0;
    for (const key of Object.keys(cfg.approvals || {})) {
        const sc = (0, match_1.matchScore)(projectName, key);
        if (sc > bestScore) {
            bestScore = sc;
            best = key;
        }
    }
    return bestScore > 0 ? best : null;
}
// ---- 2-Week Look Ahead ----
function lookAhead(sch, reportDate) {
    // ---- 1. THE DATE FILTER, and nothing else ----
    // Inclusion is decided by ONE condition: does the activity's Actual/Projected FINISH fall inside the
    // next 14 calendar days? Progress plays no part in it.
    //
    // The previous rule filtered on the projected START and dropped anything already marked Completed,
    // which is why the slide filled up with 0% work: an activity that started weeks ago and finishes next
    // Tuesday — the very thing a look-ahead is for — was excluded because its start was behind us, and
    // partially-complete work was under-represented for the same reason.
    //
    // An activity at 100% whose finish is in the past cannot qualify: its finish is outside the window.
    // The reporting date is the one the rest of the deck uses; this does not introduce another.
    const rd = iso(reportDate || undefined) || iso(sch.dataDate) || new Date();
    const end = new Date(rd.getTime() + 14 * 864e5);
    //
    // TWO ways to qualify, and both come straight from the brief:
    //   • the finish falls INSIDE the window — the plain look-ahead case, whatever the progress; and
    //   • the work is STILL OPEN (< 100%) and due on or before the window's end — "progressed activities
    //     that are currently underway". An in-progress activity whose projected finish has already slipped
    //     past today is underway and still owed; leaving it out was what emptied the slide of progressed
    //     work. Only the finish DATE and the percentage are used — no new data source.
    //
    // A 100%-complete activity is now excluded BEFORE the date test, so neither an actual finish inside the
    // window nor an overdue baseline can bring it back.
    // Milestones dropped by RULE 0 below, listed for lookahead_roles_validation.md.
    const milestonesExcluded = [];
    const src = sch.activities.filter((a) => {
        // RULE 0 — A MILESTONE IS NOT LOOK-AHEAD WORK.
        // A milestone marks a moment, not something a person can be asked to do in the next fortnight, so
        // it does not belong on this slide even when its date falls inside the window.
        //
        // The test is the SOURCE's own identification and nothing else: BOTH of the export's duration
        // columns read zero — "Duration" AND "Duration [Elapsed day]" — which is what a milestone is in
        // the scheduling tool. ASTA keeps the two apart (working time against calendar time), so a row
        // can be zero in one and not the other; requiring both is the stricter reading. Where the export
        // carries no elapsed column the test uses Duration alone and the look-ahead report says so for
        // that run — see the note on `milestoneSource` in parsers/asta_excel.ts.
        //
        // The task's NAME is never consulted — a real activity called "… Signoff" or sitting under a
        // branch called "Milestone" stays if the export gives it a duration, and `milestoneSource` is
        // false whenever the Duration cell could not be read, so a malformed export cannot silently
        // delete real work.
        //
        // Milestones are NOT removed from the schedule: every other slide (Project Overview, Tracker &
        // Schedule, the phase dates) still sees them exactly as before. This is an exclusion from THIS
        // slide's row list only.
        //
        // Groups cannot be left empty by this: the person/subheading bands below are built FROM the rows
        // that survive, so a group whose only task was a milestone is never created in the first place.
        if (a.milestoneSource) {
            // Recorded only when it WOULD otherwise have reached the slide, so the report states the number
            // of rows this rule actually removed rather than how many milestones the export happens to hold.
            const mf = iso(a.finish);
            if ((a.pct || 0) < 100 && mf && mf <= end)
                milestonesExcluded.push(a.name);
            return false;
        }
        // RULE 1 — a finished activity is never look-ahead work.
        // 100% complete is excluded outright, whatever its dates say: a look-ahead is what still has to be
        // done, and a completed task inside the window was filling the slide with work nobody needs to act
        // on. This is checked FIRST, so no date rule can readmit it.
        if ((a.pct || 0) >= 100)
            return false;
        // RULE 2 — the date window, unchanged.
        const f = iso(a.finish);
        if (!f || f > end)
            return false; // beyond the window
        return true; // inside the window, or open and overdue
    });
    const wsub = (a) => {
        const p = (0, asta_excel_1.wbsPath)(sch, a.wbs).slice(2);
        const out = [];
        for (const s of p)
            if (s && !out.includes(s))
                out.push(s);
        return out.slice(-2).join(' | ');
    };
    // ---- 2. PERSON-WISE ARRANGEMENT ----
    // Priority order: person first, then progress. People keep the order the schedule presents them in —
    // first appearance, never alphabetical — and a person's work is never interleaved with anyone else's.
    //
    // Within a person, progressed work (% > 0) comes before zero-progress work. Subheadings are ordered
    // by whether they hold any progressed task, and inside a subheading the progressed tasks lead. That
    // is the one reading that satisfies both "progressed first" and "do not duplicate a subheading row":
    // a subheading holding both kinds appears ONCE, in the progressed part of that person's block, with
    // its own progressed tasks first. Splitting it would have printed the same subheading twice.
    // ---- roles and row creation ----
    // One row per TASK + PERSON: an activity naming two people produces two rows, the same person on the
    // same task never produces two. Accountable and Head come from the export, or from the responsible
    // person when the export leaves them blank. Task selection and the 14-day window above are untouched.
    const { rows: roles, stats } = (0, roles_1.roleRows)(src);
    // Carried into lookahead_roles_validation.md so the exclusions are stated, not silent.
    stats.milestonesExcluded = milestonesExcluded;
    stats.elapsedDurationColumn = sch.hasElapsedDuration === true;
    exports.lastLookAheadRoles = stats;
    // Tasks are grouped by the WHOLE hierarchy — Head, then Accountable, then Responsible — so the names
    // appear once as a grouped header above their block instead of on every row. Within a hierarchy the
    // existing WBS sub-band still separates the work, exactly as before.
    const byHier = new Map();
    for (const r of roles) { // insertion order = the schedule's own order
        const k = (0, roles_1.hierKey)(r);
        if (!byHier.has(k))
            byHier.set(k, []);
        byHier.get(k).push(r);
    }
    const progressed = (r) => (r.activity.pct || 0) > 0;
    // People keep the schedule's own order, with one exception: work with nobody named ("Unassigned") goes
    // LAST, so the named people read first and the unattributed tail sits at the end of the section.
    const isUnassigned = (k) => (byHier.get(k)[0].responsible || '') === 'Unassigned';
    const peopleOrder = [...byHier.keys()].filter((k) => !isUnassigned(k))
        .concat([...byHier.keys()].filter(isUnassigned));
    const groups = [];
    for (const k of peopleOrder) { // source order, Unassigned last
        const list = byHier.get(k);
        const first = list[0];
        // subheadings in the order they first appear for this person
        const bySub = new Map();
        for (const r of list) {
            const sub = wsub(r.activity);
            if (!bySub.has(sub))
                bySub.set(sub, []);
            bySub.get(sub).push(r);
        }
        // a subheading that holds progressed work comes first; order is otherwise as the schedule gives it
        const subs = [...bySub.keys()].sort((a, b) => {
            const pa = bySub.get(a).some(progressed) ? 0 : 1;
            const pb = bySub.get(b).some(progressed) ? 0 : 1;
            return pa - pb;
        });
        for (const sub of subs) {
            const rows = bySub.get(sub);
            // inside the subheading: progressed first, zero-progress after, each keeping schedule order
            const ordered = [...rows.filter(progressed), ...rows.filter((r) => !progressed(r))];
            groups.push({
                hier: (0, roles_1.hierarchy)(first), head: first.head, accountable: first.accountable,
                responsible: first.responsible, sub, rows: ordered,
            });
        }
    }
    return groups;
}
/** Role statistics from the last lookAhead() call — read by lookahead_roles_validation.md. */
exports.lastLookAheadRoles = null;
function lookAheadRoleRows(groups) { return groups.flatMap((g) => g.rows); }
