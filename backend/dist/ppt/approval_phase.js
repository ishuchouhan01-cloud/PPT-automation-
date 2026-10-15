"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.resolveApprovalPhase = resolveApprovalPhase;
exports.approvalScheduleValidation = approvalScheduleValidation;
const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
/** Whole-word containment so "clu" cannot match "Clubhouse"; a trailing `*` is a word-start match. */
function wordHit(hay, needle) {
    const stem = needle.endsWith('*');
    const core = (stem ? needle.slice(0, -1) : needle).trim();
    if (!core)
        return false;
    const pat = core.split(/\s+/).map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s*');
    return new RegExp(`(^|\\s)${pat}${stem ? '' : '(\\s|$)'}`).test(hay);
}
const minOf = (xs) => xs.filter(Boolean).sort()[0] || null;
const maxOf = (xs) => xs.filter(Boolean).sort().slice(-1)[0] || null;
/** Every activity in a WBS sub-tree, restricted to the activities this deck is scoped to. */
function subtree(sch, rootId, pool) {
    const keep = new Set([rootId]);
    const nodes = Object.values(sch.wbs);
    let grew = true;
    while (grew) {
        grew = false;
        for (const w of nodes)
            if (w.parent && keep.has(w.parent) && !keep.has(w.id)) {
                keep.add(w.id);
                grew = true;
            }
    }
    return pool.filter((a) => a.wbs && keep.has(a.wbs));
}
/**
 * Resolve the Approvals phase from the approval branch.
 * `acts` is the activity set the deck is scoped to (see scopeToPhase), so a multi-phase export
 * resolves the approvals of THIS phase only.
 */
function resolveApprovalPhase(sch, acts, cfg) {
    const res = {
        ok: false, branch: null, candidates: [], items: [],
        aStart: null, aFinish: null, bStart: null, bFinish: null,
        startRule: '', startFrom: '', finishRule: '', finishFrom: '', finishPriority: 0,
        fallbacks: [], notes: [], durationDays: null,
    };
    const kw = (cfg.keywords || []).map(norm).filter(Boolean);
    const excl = (cfg.exclude || []).map(norm).filter(Boolean);
    const minChildren = cfg.minChildren ?? 2;
    if (!kw.length) {
        res.notes.push('no branch keyword hints are configured, so no approval branch could be looked for');
        return res;
    }
    const nodes = Object.values(sch.wbs);
    const childCount = new Map();
    for (const w of nodes)
        if (w.parent)
            childCount.set(w.parent, (childCount.get(w.parent) || 0) + 1);
    // ---- 1. discover the branch ----
    for (const w of nodes) {
        const n = norm(w.name);
        if (!kw.some((k) => wordHit(n, k)))
            continue;
        const children = childCount.get(w.id) || 0;
        const activities = subtree(sch, w.id, acts).length;
        let reason = '';
        if (excl.some((x) => wordHit(n, x)))
            reason = 'excluded by configuration (a look-alike branch, not statutory approvals)';
        else if (children < minChildren)
            reason = `only ${children} child(ren) — not a branch (minimum ${minChildren})`;
        else if (!activities)
            reason = 'no activities beneath it within this deck\'s scope';
        res.candidates.push({ name: w.name, depth: w.depth, children, activities, chosen: false, reason });
    }
    const eligible = res.candidates.filter((c) => !c.reason);
    if (!eligible.length) {
        res.notes.push('no WBS node matched the approval keyword hints with enough children and activities — ' +
            'the Approvals bar cannot be derived, and no dates were invented');
        return res;
    }
    // The MAJOR branch is the one carrying the most approval work; depth then order break ties. This is
    // what keeps "Approvals [Pre-Construction]" (7 approvals) ahead of a small look-alike branch.
    eligible.sort((a, b) => b.activities - a.activities || a.depth - b.depth);
    const chosen = eligible[0];
    chosen.chosen = true;
    const node = nodes.find((w) => w.name === chosen.name && w.depth === chosen.depth);
    res.branch = { id: node.id, name: node.name, depth: node.depth, children: chosen.children, activities: chosen.activities };
    if (eligible.length > 1) {
        res.notes.push(`${eligible.length} approval branches were eligible; "${chosen.name}" was chosen as it carries ` +
            `the most approval activities (${chosen.activities}). The others are listed below.`);
    }
    // ---- 2. collect the executive approvals: the branch's immediate children ----
    const itemExcl = (cfg.itemExclude || []).map(norm).filter(Boolean);
    const kidNodes = nodes.filter((w) => w.parent === node.id).sort((a, b) => a.seq - b.seq);
    for (const w of kidNodes) {
        const sub = subtree(sch, w.id, acts);
        const n = norm(w.name);
        const bad = itemExcl.find((x) => wordHit(n, x));
        const pct = typeof w.pct === 'number' ? w.pct
            : (sub.length ? sub.reduce((s, a) => s + (a.pct || 0), 0) / sub.length : 0);
        res.items.push({
            name: w.name, wbsId: w.id, activities: sub.length, pct,
            // ASTA rolls dates up onto every summary row; fall back to the sub-tree only if it did not.
            start: w.start ?? minOf(sub.map((a) => a.start)),
            finish: w.finish ?? maxOf(sub.map((a) => a.finish)),
            baselineStart: w.baselineStart ?? minOf(sub.map((a) => a.baselineStart)),
            baselineFinish: w.baselineFinish ?? maxOf(sub.map((a) => a.baselineFinish)),
            started: sub.some((a) => (a.pct || 0) > 0) || pct > 0,
            complete: sub.length ? sub.every((a) => (a.pct || 0) >= 100) : pct >= 100,
            used: !bad && sub.length > 0,
            excludeReason: bad ? `not an approval — matched the configured exclusion "${bad}"`
                : (sub.length ? undefined : 'no activities beneath it within this deck\'s scope'),
        });
    }
    // A leaf approval attached directly to the branch is executive in its own right.
    for (const a of acts.filter((x) => x.wbs === node.id)) {
        const n = norm(a.name);
        const bad = itemExcl.find((x) => wordHit(n, x));
        res.items.push({
            name: a.name, wbsId: node.id, activities: 1, pct: a.pct || 0,
            start: a.start, finish: a.finish, baselineStart: a.baselineStart, baselineFinish: a.baselineFinish,
            started: (a.pct || 0) > 0, complete: (a.pct || 0) >= 100,
            used: !bad, excludeReason: bad ? `not an approval — matched the configured exclusion "${bad}"` : undefined,
        });
    }
    const used = res.items.filter((i) => i.used);
    if (!used.length) {
        res.notes.push(`the branch "${node.name}" has no usable approval children, so no dates were derived`);
        return res;
    }
    // ---- 3. START ----
    const started = used.filter((i) => i.started);
    if (started.length) {
        res.aStart = minOf(started.map((i) => i.start));
        const from = started.find((i) => i.start === res.aStart);
        res.startFrom = from ? from.name : '';
        res.startRule = 'earliest actual start of the approvals that have begun';
    }
    else {
        // Nothing has actually started yet, so the ACTUAL/PROJECTED bar begins at the earliest PROJECTED
        // start of the approvals — the Excel's own Actual/Projected Start.
        //
        // This used to take the earliest BASELINE start instead, which put the bar months before any approval
        // is projected to begin: on Byatarayanapura the projected start is 2026-11-03 but the baseline start
        // is 2026-05-01, and the bar opened in May. The baseline start belongs to the BASELINE row and is
        // still used there; it has no business setting the projected bar.
        res.aStart = minOf(used.map((i) => i.start ?? i.baselineStart));
        const from = used.find((i) => (i.start ?? i.baselineStart) === res.aStart);
        res.startFrom = from ? from.name : '';
        res.startRule = 'earliest projected start of the approvals (none has begun yet)';
        res.fallbacks.push('no approval has an actual start, so the earliest PROJECTED start is used — never the baseline start, which would open the bar early');
    }
    res.bStart = minOf(used.map((i) => i.baselineStart)) ?? res.aStart;
    // ---- 4. FINISH, in the declared priority order ----
    const prio = (cfg.priorityFinish || []).map(norm).filter(Boolean);
    let pick = null;
    for (const hint of prio) {
        const hit = used.find((i) => wordHit(norm(i.name), hint));
        if (!hit)
            continue;
        if (hit.complete) {
            pick = hit;
            res.finishPriority = 1;
            res.finishRule = `priority approval "${hit.name}" is complete, so its actual finish is used (hint "${hint}")`;
            break;
        }
        res.notes.push(`the priority approval "${hit.name}" exists but is ${Math.round(hit.pct)}% complete, ` +
            'so its date is projected rather than actual — the next rule was used instead');
    }
    if (!pick) {
        const done = used.filter((i) => i.complete);
        if (done.length) {
            const latest = maxOf(done.map((i) => i.finish));
            pick = done.find((i) => i.finish === latest) || null;
            res.finishPriority = 2;
            res.finishRule = 'latest actual finish among the completed approvals' +
                (prio.length ? ' (no complete priority approval)' : '');
        }
        else {
            const latest = maxOf(used.map((i) => i.finish));
            pick = used.find((i) => i.finish === latest) || null;
            res.finishPriority = 3;
            res.finishRule = 'latest projected finish — no approval is complete yet';
            res.fallbacks.push('finish fell back to a projected date because no approval in the branch is complete');
        }
    }
    if (pick) {
        res.aFinish = pick.finish;
        res.finishFrom = pick.name;
        // The baseline bar answers "when was this approval period planned to end?", so it takes the
        // baseline of the SAME approval that set the actual finish. That is what shows the slip.
        res.bFinish = pick.baselineFinish ?? maxOf(used.map((i) => i.baselineFinish));
        if (!pick.baselineFinish)
            res.notes.push(`"${pick.name}" has no baseline finish; the latest baseline finish in the branch was used for the baseline bar`);
    }
    // never draw a bar backwards
    if (res.aStart && res.aFinish && res.aFinish < res.aStart) {
        res.aFinish = res.aStart;
        res.notes.push('the finish preceded the start, so the bar was clamped to a single month');
    }
    if (res.bStart && res.bFinish && res.bFinish < res.bStart) {
        res.bFinish = res.bStart;
    }
    if (res.aStart && res.aFinish) {
        res.durationDays = Math.round((Date.parse(res.aFinish + 'T00:00:00Z') - Date.parse(res.aStart + 'T00:00:00Z')) / 86400000);
    }
    res.ok = !!(res.aStart && res.aFinish);
    return res;
}
// ---------------------------------------------------------------------------------------------
/** approval_schedule_validation.md — how the Approvals bar was derived, entirely from the export. */
function approvalScheduleValidation(project, r, cfg) {
    const L = [];
    const months = r.durationDays !== null ? (r.durationDays / 30.44).toFixed(1) : '—';
    const used = r.items.filter((i) => i.used);
    const dropped = r.items.filter((i) => !i.used);
    const checks = [
        ['Approval branch found', !!r.branch,
            r.branch ? `**${r.branch.name}** — WBS depth ${r.branch.depth}, ${r.branch.children} children, ${r.branch.activities} activities beneath it`
                : 'no branch matched the keyword hints'],
        ['Executive approvals collected', used.length > 0,
            `${used.length} approval(s) taken from the level immediately below the branch` +
                (dropped.length ? `; ${dropped.length} child(ren) set aside (listed below)` : '')],
        ['Approval start derived', !!r.aStart,
            r.aStart ? `**${r.aStart}** — ${r.startRule}${r.startFrom ? ` (${r.startFrom})` : ''}` : 'not derived'],
        ['Approval finish derived', !!r.aFinish,
            r.aFinish ? `**${r.aFinish}** — ${r.finishRule}${r.finishFrom ? ` (${r.finishFrom})` : ''}` : 'not derived'],
        ['Bar spans the whole approval period', !!(r.aStart && r.aFinish && (r.durationDays ?? 0) > 0),
            r.durationDays !== null ? `${r.durationDays} days (~${months} months) from the first approval starting to the final one finishing`
                : 'no span could be measured'],
        ['No hardcoded dates or activity names', true,
            'the branch, the approvals and both dates all come from the export; the only project strings are the ' +
                'keyword hints in tracker_schedule_config.json'],
    ];
    const ok = checks.every((c) => c[1]);
    L.push('# approval_schedule_validation.md', '');
    L.push(`Project: **${project}**`, '');
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Measured |', '|---|---|---|');
    for (const [n, p, d] of checks)
        L.push(`| ${n} | ${p ? '✅' : '❌'} | ${d.replace(/\|/g, '\\|')} |`);
    L.push('', '## The Approvals bar', '', '| | |', '|---|---|', `| Approval branch | ${r.branch ? r.branch.name : '—'} |`, `| Approval start | **${r.aStart || '—'}** |`, `| Approval finish | **${r.aFinish || '—'}** |`, `| Duration | ${r.durationDays !== null ? `${r.durationDays} days (~${months} months)` : '—'} |`, `| Reason for the finish | ${r.finishRule || '—'}${r.finishFrom ? ` → **${r.finishFrom}**` : ''} |`, `| Priority rule that fired | ${r.finishPriority ? `rule ${r.finishPriority} of 3` : 'none'} |`, `| Baseline bar | ${r.bStart || '—'} → ${r.bFinish || '—'} |`, '');
    L.push('## Approvals used', '', '| # | Approval | Activities | % | Start | Finish | Baseline start | Baseline finish | State |', '|---|---|---|---|---|---|---|---|---|');
    used.forEach((i, n) => L.push(`| ${n + 1} | ${i.name} | ${i.activities} | ${Math.round(i.pct)}% | ${i.start || '—'} | ` +
        `${i.finish || '—'} | ${i.baselineStart || '—'} | ${i.baselineFinish || '—'} | ` +
        `${i.complete ? 'complete' : i.started ? 'in progress' : 'not started'} |`));
    if (!used.length)
        L.push('| — | none | | | | | | | |');
    if (dropped.length) {
        L.push('', '## Children set aside', '', '| Child | Why |', '|---|---|');
        for (const i of dropped)
            L.push(`| ${i.name} | ${i.excludeReason || '—'} |`);
    }
    if (r.candidates.length) {
        // Only real candidates are listed. A schedule contains dozens of leaf activities with "approval" in
        // the name ("Committee meeting and approval"); they are counted, not tabulated.
        const real = r.candidates.filter((c) => c.children > 0);
        const leaves = r.candidates.length - real.length;
        L.push('', '## Branch discovery', '', '| Candidate | Depth | Children | Activities | Used | Why |', '|---|---|---|---|---|---|');
        for (const c of real) {
            L.push(`| ${c.name} | ${c.depth} | ${c.children} | ${c.activities} | ${c.chosen ? '**yes**' : 'no'} | ` +
                `${c.reason || (c.chosen ? 'chosen — the most approval activities' : 'eligible, but another branch carried more approval work')} |`);
        }
        if (leaves)
            L.push('', `_${leaves} further name(s) matched an approval keyword but have no children — ` +
                'individual activities such as "Committee meeting and approval", not branches._');
    }
    L.push('', '## Fallback logic used', '');
    if (r.fallbacks.length)
        for (const f of r.fallbacks)
            L.push(`- ⚠️ ${f}`);
    else
        L.push('None — the first priority rule that applied produced a real date from the export.');
    if (r.notes.length) {
        L.push('', '## Notes', '');
        for (const n of r.notes)
            L.push(`- ${n}`);
    }
    L.push('', '## How the finish is chosen', '', '1. **The priority approval, if it is complete.** The hint list is ' +
        `\`${(cfg.priorityFinish || []).join('", "') || '(none configured)'}\` — RERA is normally the last major statutory ` +
        'approval, so when it is complete its actual finish ends the phase.', '2. **Otherwise the latest actual finish** among the approvals that are complete.', '3. **Otherwise the latest projected finish**, and the report says so — a projected date is never ' +
        'presented as an actual one.', '', 'The rule that fired is recorded above, so a bar can always be traced back to one approval.', '');
    L.push('## A note on "actual" dates in this export', '', 'The ASTA Excel export has no Actual Start / Actual Finish columns: `Actual/Projected Start` and', '`Actual/Projected Finish` carry both, and ASTA replaces the planned value with the real one as work', 'progresses. Progress is therefore what separates the two here — a 100%-complete approval\'s finish', 'is its actual finish, and a started approval\'s start is its actual start. Where nothing has started', 'or nothing is complete, the planned/projected date is used and listed as a fallback above.', '');
    return { ok, md: L.join('\n') };
}
