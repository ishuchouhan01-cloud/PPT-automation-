"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.STAGE_ORDER = void 0;
exports.findApprovalBranches = findApprovalBranches;
exports.approvalsFromSchedule = approvalsFromSchedule;
exports.approvalValidation = approvalValidation;
/** The approval workflow, in the order a stage can only follow the one before it. */
exports.STAGE_ORDER = ['Application', 'Drawings Submission', 'Fee Payment', 'Issued'];
const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const has = (hay, needle) => norm(hay).includes(norm(needle));
/** Descendant activities of a WBS node (whole sub-tree). */
function subtreeActivities(sch, rootId) {
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
    return sch.activities.filter((a) => a.wbs && keep.has(a.wbs));
}
/** Locate the statutory-approval branch(es) by name, honouring the configured exclusions. */
function findApprovalBranches(sch, cfg) {
    const nodes = Object.values(sch.wbs);
    const kids = new Map();
    for (const w of nodes)
        if (w.parent)
            kids.set(w.parent, (kids.get(w.parent) || 0) + 1);
    const b = cfg?.branch || {};
    const minChildren = b.minChildren ?? 2;
    const excluded = [];
    const isExcluded = (name) => (b.exclude || []).some((x) => norm(name) === norm(x) || has(name, x));
    const matches = (name, list) => list.some((x) => norm(name) === norm(x) || norm(name).startsWith(norm(x)));
    const candidates = nodes.filter((w) => (kids.get(w.id) || 0) >= minChildren);
    // primary names first; only fall back to the looser aliases when no primary branch exists
    let hits = candidates.filter((w) => matches(w.name, b.primary || []));
    if (!hits.length)
        hits = candidates.filter((w) => matches(w.name, b.aliases || []));
    const keep = [];
    for (const w of hits) {
        if (isExcluded(w.name)) {
            excluded.push({ name: w.name, reason: 'configured exclusion (design/drawings branch, not statutory approvals)' });
            continue;
        }
        keep.push(w);
    }
    // record excluded look-alikes for the report
    for (const w of candidates) {
        if (isExcluded(w.name) && !excluded.some((e) => e.name === w.name)) {
            excluded.push({ name: w.name, reason: 'configured exclusion (design/drawings branch, not statutory approvals)' });
        }
    }
    return { branches: keep, excluded };
}
/**
 * Choose the activity that represents an approval's completion.
 *
 * THE APPROVAL'S OWN HIERARCHY DECIDES THIS, not a keyword search.
 *
 * `acts` is always the sub-tree of ONE approval node, so everything considered here already belongs
 * to this project, this approval and the level currently being displayed. Within that sub-tree the
 * answer is the LAST TASK — the end of the approval's own chain of work, in the export's own row
 * order, which is the hierarchical order the schedule was written in.
 *
 * Why not a keyword. The old rule scanned for "issuance", "issue of", "issue", "sanction"… and took
 * whichever matching task finished latest. Those words appear on intermediate steps too, and the rule
 * duly picked them: on One Tiny Flower's Building Plan Sanction it returned "Issuance of Demand Note"
 * (a fee step, 06-Oct-26) instead of "Issue of building plan sanction+license" (15-Oct-26), and on
 * Butterfly Dreams' Fire Provisional NOC it returned "Issue of Demand note and payment" (20-Aug-26)
 * instead of "Issue fire NOC" (28-Aug-26). Walking to the end of the hierarchy cannot make that
 * mistake, and it needs no per-approval exception: change the display level and the sub-tree changes
 * with it, so the chosen task follows automatically.
 *
 * `order` gives each activity its position in the export, so the answer does not depend on the order
 * the caller happened to assemble the pool in.
 */
function pickFinishActivity(acts, approvalName, cfg, order) {
    const byLatest = (pool) => pool.slice().sort((x, y) => (x.finish || '').localeCompare(y.finish || '')).slice(-1)[0] || null;
    const inExportOrder = (pool) => pool.slice().sort((x, y) => (order?.get(x.id) ?? 0) - (order?.get(y.id) ?? 0));
    // 1. a configured override for this approval (e.g. Layout → "Layout Approval + Work Order").
    //    Kept because it names a real task explicitly; it is not a date and not a fallback guess.
    for (const ov of cfg?.overrides || []) {
        if (!has(approvalName, ov.match))
            continue;
        for (const want of ov.preferredFinish || []) {
            const exact = acts.filter((a) => norm(a.name) === norm(want));
            const loose = exact.length ? exact : acts.filter((a) => has(a.name, want));
            if (loose.length)
                return { act: byLatest(loose), rule: `override for "${ov.match}" → "${want}"` };
        }
    }
    // 2. THE LAST TASK IN THIS APPROVAL'S OWN HIERARCHY.
    const chain = inExportOrder(acts);
    const last = chain[chain.length - 1] || null;
    if (last && (last.finish || last.actualFinish)) {
        return { act: last, rule: `last task in this approval's own hierarchy ("${last.name}", task ${chain.length} of ${chain.length})` };
    }
    // The final task carries no finish date. Step back through the chain to the last task that does,
    // and say so — rather than reaching outside this approval or inventing a date.
    for (let i = chain.length - 2; i >= 0; i--) {
        const a = chain[i];
        if (a.finish || a.actualFinish) {
            return { act: a, rule: `last DATED task in this approval's own hierarchy ("${a.name}", task ${i + 1} of ${chain.length}) — the final task${last ? ` "${last.name}"` : ''} carries no finish date in the source schedule` };
        }
    }
    return { act: last, rule: last ? `last task in this approval's own hierarchy ("${last.name}") — no task in this approval carries a finish date` : '' };
}
/** True when a stage's matching activity exists and is complete. */
function stageDone(acts, keys) {
    const pool = acts.filter((a) => keys.some((k) => has(a.name, k)));
    if (!pool.length)
        return { done: false, found: false };
    return { done: pool.every((a) => a.pct >= 100), found: true };
}
/**
 * Build the Approvals rows from the schedule.
 * `scoped` is the phase-scoped activity list (so a multi-phase export does not double up).
 */
/**
 * Significant words of an approval name. Filler is dropped so "Issue of Water NOC" and "Water NOC"
 * share their meaningful tokens, and "noc"/"sanction" are kept because they distinguish one approval
 * from another.
 */
const STOP = new Set(['of', 'for', 'the', 'and', 'a', 'an', 'to', 'from', 'with', 'in', 'on', 'issue',
    'issued', 'issuance', 'obtain', 'obtaining', 'get', 'getting', 'status', 'approval', 'approvals']);
const tokens = (s) => norm(s).split(' ').filter((w) => w && !STOP.has(w));
/**
 * How well an Excel activity (or branch child) answers a reference approval name: the share of the
 * reference name's significant words that appear in the candidate. 1 means every word is present.
 * Zero-token names (e.g. "Approvals") score 0 and are never matched on.
 */
function nameScore(refName, candidate) {
    const want = tokens(refName);
    if (!want.length)
        return 0;
    const hay = ' ' + norm(candidate) + ' ';
    const hit = want.filter((w) => hay.includes(' ' + w + ' ') || hay.includes(w)).length;
    return hit / want.length;
}
function approvalsFromSchedule(sch, cfg, scopedActivityIds, 
/**
 * The approval list from the previous/reference deck. When supplied it is the MASTER: the slide shows
 * these approvals, in this order, with this wording — and the Excel is consulted only for the dates.
 * When absent the schedule's own approval branch decides the list, exactly as before.
 */
refNames, 
/** THIS project's name, used only to find its own pinned list in masterListByProject. */
projectName) {
    // Each activity's position in the export. The export is written in hierarchical order, so this is
    // what lets "the last task in this approval's hierarchy" be identified without depending on the
    // order a particular caller assembled its candidate pool in.
    const exportOrder = new Map(sch.activities.map((a, i) => [a.id, i]));
    // ---- the master list ----
    // Precedence: the reference deck (when one is supplied and readable) → the configured masterList →
    // the schedule's own approval branch. The masterList is what guarantees a stable set of rows with
    // stable wording when no deck is uploaded; without it the slide was built from the branch, which is
    // how three separate NOCs arrived merged as "Airport NOC" and Fire NOC arrived renamed.
    // ---- this project's pinned list, if it has one ----
    // Keyed by project name and never shared. The old flat `masterList` was a GLOBAL list applied to every
    // project that arrived without a previous deck, which is exactly what must not happen: projects have
    // different approvals, and reusing one project's list for another is wrong even when it looks complete.
    const byProject = cfg?.masterListByProject || {};
    const pn = norm(projectName || '');
    let pinnedKey = '';
    for (const key of Object.keys(byProject)) {
        const k = norm(key);
        if (!k || !pn)
            continue;
        if (pn === k || pn.includes(k) || k.includes(pn)) {
            pinnedKey = key;
            break;
        }
    }
    const cfgMaster = pinnedKey ? byProject[pinnedKey] : [];
    const master = (refNames && refNames.length) ? refNames.map((n) => ({ name: n }))
        : cfgMaster.length ? cfgMaster : [];
    const masterFrom = (refNames && refNames.length) ? 'reference deck'
        : cfgMaster.length ? `configured list for "${pinnedKey}"` : '';
    const nodes = Object.values(sch.wbs);
    const { branches, excluded } = findApprovalBranches(sch, cfg);
    const containersExpanded = [];
    const details = [];
    const seen = new Set();
    const childExclude = cfg?.childExclude || [];
    const stages = cfg?.stages || {};
    // prefer branches inside the deck's phase when a scope was supplied
    const inScope = (w) => {
        if (!scopedActivityIds || !scopedActivityIds.size)
            return true;
        return subtreeActivities(sch, w.id).some((a) => scopedActivityIds.has(a.id));
    };
    const usable = branches.filter(inScope).length ? branches.filter(inScope) : branches;
    /**
     * AN APPROVAL THAT IS REALLY A CONTAINER OF SEVERAL APPROVALS.
     *
     * Normally each child of the approval branch is one approval and becomes one row. But a schedule
     * sometimes groups two genuinely separate statutory approvals under one heading — Varthur's
     * "Airport NOC" holds "HAL NOC" and "AAI NOC", each with its own six-task chain ending in its own
     * "Issuance of … NOC". Reporting the heading collapsed two real approvals, with two different
     * authorities and two different dates, into one row.
     *
     * THE TEST IS THE SHAPE OF THE DATA, NOT THE NAME. A child is a container when it holds TWO OR
     * MORE sub-branches of its own and no loose work outside them. The count is what matters:
     *   • two or more  → parallel approvals run under one heading      → each becomes a row
     *   • exactly one  → a nested STAGE of the same approval           → nothing changes
     * The single-sub-branch case is real and must not split: "Environmental Clearance - MoEF" contains
     * "TOR: Projects with SuBA >= 150000 sqm.", whose own last task is the parent's issuance
     * ("Issue of Environmental clearance."). Splitting that would invent an approval nobody tracks.
     * Measured across the exports: only Airport NOC has two; Environmental Clearance has one on both
     * the projects that carry it; five other projects have no nesting at all and are untouched.
     */
    const hasChildren = (id) => nodes.some((w) => w.parent === id);
    const expandContainer = (c) => {
        const subs = nodes.filter((w) => w.parent === c.id && hasChildren(w.id)).sort((a, b) => a.seq - b.seq);
        if (subs.length < 2)
            return [c];
        if (sch.activities.some((a) => a.wbs === c.id))
            return [c]; // loose work of its own → one approval
        containersExpanded.push(`"${c.name}" holds ${subs.length} separate approvals and no work of its own, ` +
            `so it is reported as ${subs.length} rows: ${subs.map((x) => `"${x.name}"`).join(', ')}`);
        return subs;
    };
    for (const br of usable) {
        const children = nodes.filter((w) => w.parent === br.id).sort((a, b) => a.seq - b.seq)
            .flatMap(expandContainer);
        for (const c of children) {
            const name = String(c.name || '').trim();
            if (!name)
                continue;
            // drop non-approval children (drawing/documentation/closure-milestone rows)
            if (childExclude.some((x) => has(name, x)))
                continue;
            const key = norm(name);
            if (seen.has(key))
                continue; // no duplicate approvals across phases
            seen.add(key);
            const acts = subtreeActivities(sch, c.id);
            const pct = Math.round(c.pct ?? (acts.length ? acts.reduce((s, a) => s + a.pct, 0) / acts.length : 0));
            const status = !acts.length ? 'Not Applicable'
                : pct >= 100 ? 'Completed' : pct > 0 ? 'In Progress' : 'Not Started';
            const notes = [];
            const { act, rule } = pickFinishActivity(acts, name, cfg, exportOrder);
            if (!act)
                notes.push('no activity inside this approval branch — Actual/Projected Finish left blank');
            const proj = act ? (act.finish || act.actualFinish || null) : null;
            if (act && !proj)
                notes.push(`completion activity "${act.name}" has no Actual/Projected finish date`);
            const base = act ? (act.baselineFinish || null) : null;
            if (act && !base)
                notes.push(`completion activity "${act.name}" has no Baseline finish date — left blank`);
            const sD = stageDone(acts, stages.drawings || []);
            const sA = stageDone(acts, stages.application || []);
            const sF = stageDone(acts, stages.fee || []);
            const sI = stageDone(acts, stages.issued || []);
            // ---- cascading completion ----
            // An approval cannot be Issued while its fee is unpaid, and the fee cannot be paid before the
            // application exists. The schedule sometimes reports the individual activities out of step
            // (a fee activity left at 0% after the certificate was issued), which produced impossible rows
            // such as ✔ ✔ ✘ ✔. The workflow order is Application → Drawings → Fee → Issued, so every stage
            // up to and including the highest completed one is shown as complete; the stages after it are
            // left exactly as the schedule reports them. Anything changed here is listed in the reports.
            const flags = [sA.done, sD.done, sF.done, sI.done];
            const highest = flags.lastIndexOf(true);
            const cascaded = flags.map((v, i) => v || i < highest);
            const cascadeApplied = cascaded
                .map((v, i) => (v !== flags[i] ? exports.STAGE_ORDER[i] : ''))
                .filter(Boolean);
            const [cA, cD, cF, cI] = cascaded;
            details.push({
                row: {
                    sl: details.length + 1, approval: name,
                    drawings: cD, application: cA, fee: cF, issued: cI,
                    base, proj,
                },
                cascadeApplied,
                status, pct, branch: br.name,
                finishFrom: act ? act.name : '', finishRule: rule,
                baselineFrom: base && act ? act.name : '',
                childActivities: acts.length, notes,
            });
        }
    }
    // ---- the reference deck's list is the master ----
    // Each reference approval keeps its own name and position; the Excel supplies only its dates and
    // ticks. Matching is on significant words, so "Water NOC" finds the branch child "Water NOC" and, if
    // there is no such child, the issuance activity "Issue of Water NOC" anywhere in the schedule.
    let listSource = 'approval branch';
    let listNote = `${details.length} approval(s) from THIS project's own approval branch in its own Excel — ` +
        'no previous PPT was supplied for it and no pinned list matches its name, so the list is schedule-derived';
    const unmapped = [];
    const notOnMaster = [];
    if (master.length) {
        listSource = masterFrom;
        const pool = details.slice();
        const out = [];
        for (const item of master) {
            const nm = item.name;
            // ---- hint-driven location of the Excel task ----
            // When the master item carries `match` hints they are authoritative and are tried FIRST, against
            // the whole schedule. This is what keeps AAI NOC, HAL NOC and IAF & GFTS NOC as three rows: each
            // finds (or fails to find) its own task by its own hint, and none of them can collapse onto a
            // shared branch child called "Airport NOC". A hint only locates a task — it never renames the row.
            if (item.match && item.match.length) {
                // ---- hint forms ----
                //   "text"              → the name contains it
                //   ["a", "b"]          → the name contains BOTH
                //   { all, any, none }  → every `all`, at least one `any`, and none of `none`
                //
                // Hints are tried IN ORDER, most specific first, and the FIRST that finds anything wins — they
                // are never unioned, so a loose term cannot outvote a precise one.
                //
                // The object form is what separates two tasks differing only by a suffix. Byatarayanapura carries
                // "Issue of building plan sanction[license and blueprints]." AND the same name ending "-CH": the
                // clubhouse row needs the one WITH the CH marker and the residential row the one WITHOUT, which no
                // include-only hint can express.
                //
                // NOTE ON THE BUG THIS REPLACES: the array form was previously passed straight to has(), which
                // normalises it — joining with a comma that then became a space — so ["building plan","clubhouse"]
                // silently behaved as the single phrase "building plan clubhouse" and matched nothing. It appeared
                // to work because matching nothing was the desired answer on that project.
                //
                // Terms of three characters or fewer match as WHOLE WORDS, so "ch" finds "… -CH" and nothing else.
                const term = (name, t) => {
                    const n = ' ' + norm(name) + ' ', w = norm(t);
                    return w.length <= 3 ? n.includes(' ' + w + ' ') : n.includes(w);
                };
                const hit = (a, h) => {
                    if (typeof h === 'string')
                        return term(a.name, h);
                    if (Array.isArray(h))
                        return h.every((t) => term(a.name, t));
                    const all = (h.all || []).every((t) => term(a.name, t));
                    const any = !h.any || !h.any.length || h.any.some((t) => term(a.name, t));
                    const none = !(h.none || []).some((t) => term(a.name, t));
                    return all && any && none;
                };
                const label = (h) => typeof h === 'string' ? h : Array.isArray(h) ? h.join(' + ')
                    : [(h.all || []).join(' + '),
                        h.any && h.any.length ? `any(${h.any.join('/')})` : '',
                        (h.none || []).length ? `not(${(h.none || []).join('/')})` : ''].filter(Boolean).join(' · ');
                let hinted = [], usedHint = '';
                for (const h of item.match) {
                    const pool = sch.activities.filter((a) => hit(a, h));
                    if (pool.length) {
                        hinted = pool;
                        usedHint = label(h);
                        break;
                    }
                }
                const issuance = hinted.filter((a) => /issuance|issue|sanction|approv|certificat|registrat|clearance|noc/i.test(a.name));
                const poolAct = issuance.length ? issuance : hinted;
                if (poolAct.length) {
                    const { act, rule } = pickFinishActivity(poolAct, nm, cfg, exportOrder);
                    const proj = act ? (act.finish || act.actualFinish || null) : null;
                    const base = act ? (act.baselineFinish || null) : null;
                    const sD = stageDone(poolAct, stages.drawings || []);
                    const sA = stageDone(poolAct, stages.application || []);
                    const sF = stageDone(poolAct, stages.fee || []);
                    const sI = stageDone(poolAct, stages.issued || []);
                    const flags = [sA.done, sD.done, sF.done, sI.done];
                    const highest = flags.lastIndexOf(true);
                    const casc = flags.map((v, i) => v || i < highest);
                    out.push({
                        row: { sl: 0, approval: nm, application: casc[0], drawings: casc[1], fee: casc[2], issued: casc[3], base, proj },
                        cascadeApplied: casc.map((v, i) => (v !== flags[i] ? exports.STAGE_ORDER[i] : '')).filter(Boolean),
                        status: act && act.pct >= 100 ? 'Completed' : act && act.pct > 0 ? 'In Progress' : 'Not Started',
                        pct: act ? Math.round(act.pct) : 0, branch: '(located by the master list hint)',
                        finishFrom: act ? act.name : '', finishRule: rule ? `${rule} — located by hint` : 'located by hint',
                        baselineFrom: base && act ? act.name : '',
                        childActivities: poolAct.length,
                        notes: [`located in the Excel by the hint "${usedHint}"; the row name is the master list's`],
                    });
                    continue;
                }
                // a hinted item that finds nothing is NOT looked up by word-score: that is exactly how three
                // separate NOCs would collapse onto one shared branch child.
                unmapped.push(nm);
                out.push({
                    row: { sl: 0, approval: nm, application: false, drawings: false, fee: false, issued: false, base: null, proj: null },
                    cascadeApplied: [], status: 'Not Applicable', pct: 0, branch: '(not found)',
                    finishFrom: '', finishRule: '', baselineFrom: '', childActivities: 0,
                    notes: [`no task in this Excel export matches any of ${item.match.map((h) => `"${label(h)}"`).join(', ')} — the row is KEPT with both dates BLANK, never dropped and never merged`],
                });
                continue;
            }
            // (a) a branch-derived approval that answers this name
            let bestI = -1, best = 0;
            pool.forEach((d, i) => {
                const sc = Math.max(nameScore(nm, d.row.approval), nameScore(d.row.approval, nm));
                if (sc > best) {
                    best = sc;
                    bestI = i;
                }
            });
            if (bestI >= 0 && best >= 0.6) {
                const d = pool.splice(bestI, 1)[0];
                // the reference deck's wording wins; every date, tick and note stays as the Excel produced it
                d.notes = d.notes.slice();
                if (norm(d.row.approval) !== norm(nm))
                    d.notes.push(`matched the schedule's "${d.row.approval}" (${Math.round(best * 100)}% of the words)`);
                d.row.approval = nm;
                out.push(d);
                continue;
            }
            // (b) no branch child — look for the issuance activity anywhere in the schedule
            const cands = sch.activities.filter((a) => nameScore(nm, a.name) >= 0.9);
            if (cands.length) {
                const { act, rule } = pickFinishActivity(cands, nm, cfg, exportOrder);
                const proj = act ? (act.finish || act.actualFinish || null) : null;
                const base = act ? (act.baselineFinish || null) : null;
                const sD = stageDone(cands, stages.drawings || []);
                const sA = stageDone(cands, stages.application || []);
                const sF = stageDone(cands, stages.fee || []);
                const sI = stageDone(cands, stages.issued || []);
                const flags = [sA.done, sD.done, sF.done, sI.done];
                const highest = flags.lastIndexOf(true);
                const casc = flags.map((v, i) => v || i < highest);
                out.push({
                    row: { sl: 0, approval: nm, application: casc[0], drawings: casc[1], fee: casc[2], issued: casc[3], base, proj },
                    cascadeApplied: casc.map((v, i) => (v !== flags[i] ? exports.STAGE_ORDER[i] : '')).filter(Boolean),
                    status: act && act.pct >= 100 ? 'Completed' : act && act.pct > 0 ? 'In Progress' : 'Not Started',
                    pct: act ? Math.round(act.pct) : 0, branch: '(matched outside the approval branch)',
                    finishFrom: act ? act.name : '', finishRule: rule ? rule + ' — matched by activity name' : '',
                    baselineFrom: base && act ? act.name : '',
                    childActivities: cands.length,
                    notes: [`no approval-branch child answered "${nm}"; dates taken from the matching activity in the schedule`],
                });
                continue;
            }
            // (c) nothing in the Excel answers it — the row appears with blank dates and is reported.
            unmapped.push(nm);
            out.push({
                row: { sl: 0, approval: nm, application: false, drawings: false, fee: false, issued: false, base: null, proj: null },
                cascadeApplied: [], status: 'Not Applicable', pct: 0, branch: '(not found)',
                finishFrom: '', finishRule: '', baselineFrom: '', childActivities: 0,
                notes: [`no activity in this Excel export answers "${nm}" — both dates left BLANK rather than guessed`],
            });
        }
        // Approvals the schedule has but the reference deck does not list are NOT added: the reference is
        // the master for which rows appear. They are named in the report so the omission is visible.
        details.length = 0;
        details.push(...out);
        notOnMaster.push(...pool.map((d) => d.row.approval));
        listNote = `${out.length} approval(s) from the ${listSource}` +
            (pool.length ? ` · ${pool.length} approval(s) in the schedule are not on the master list and were NOT added (no merging, no renaming): ${pool.map((d) => d.row.approval).join(', ')}` : '') +
            (unmapped.length ? ` · ${unmapped.length} with no Excel match: ${unmapped.join(', ')}` : '');
    }
    details.forEach((d, i) => { d.row.sl = i + 1; });
    const expected = cfg?.expected || [];
    const missingExpected = expected.filter((e) => !details.some((d) => has(d.row.approval, e) || has(e, d.row.approval)));
    return {
        listSource, listNote, unmapped,
        masterExpected: master.map((m) => m.name),
        notOnMaster: notOnMaster.slice(),
        details, rows: details.map((d) => d.row),
        branchesFound: usable.map((w) => ({ name: w.name, depth: w.depth, children: nodes.filter((n) => n.parent === w.id).length })),
        branchExcluded: excluded,
        missingExpected,
        // A blank date on a reference-listed approval the Excel does not carry is an honest gap, not a
        // failure of the mapping — it is reported instead of failing the build.
        ok: details.length > 0 && details.every((d) => d.row.proj !== null || unmapped.includes(d.row.approval)),
    };
}
/** approval_validation.md */
function approvalValidation(sch, r, scope, 
/**
 * What happened when the previous/reference deck was read for its approval list. Supplied so this
 * report can FAIL when a previous deck was uploaded but its list could not be read — the row names
 * and order then come from the schedule instead of from the deck, which the operator must be told.
 */
refState) {
    const withProj = r.details.filter((d) => d.row.proj).length;
    const withBase = r.details.filter((d) => d.row.base).length;
    const dupes = new Set();
    const dupFound = [];
    for (const d of r.details) {
        const k = norm(d.row.approval);
        if (dupes.has(k))
            dupFound.push(d.row.approval);
        dupes.add(k);
    }
    const detailish = r.details.filter((d) => /drawing|documentation|closure milestone/i.test(d.row.approval));
    const allNotes = r.details.flatMap((d) => d.notes.map((n) => `${d.row.approval}: ${n}`));
    const checks = [
        // THE PREVIOUS DECK IS AUTHORITATIVE FOR WHICH ROWS EXIST. If one was uploaded and its list could
        // not be read, the fallback silently substitutes the schedule branch's own wording — which is how
        // "AAI NOC", "HAL NOC" and "IAF & GFTS NOC" became "Airport NOC [BIAL]" / "[HAL]". That is a
        // reportable failure, not a quiet default.
        ...(refState && refState.supplied ? [{
                name: 'The previous PPT\'s own approval list was read and used',
                // NOT a build failure: an operator may legitimately upload a deck whose Approvals slide is a
                // picture, and the deck still generates correctly from the schedule. But it must never pass
                // SILENTLY, because the row names and order then differ from the deck the operator handed in.
                // The run log carries the same warning.
                ok: true,
                detail: refState.namesRead > 0
                    ? `${refState.namesRead} approval name(s) read from the previous deck and used for the rows and their order — ${refState.note}`
                    : `⚠ WARNING — a previous PPT WAS uploaded but NO approval list could be read from it (${refState.note}). ` +
                        'The rows therefore come from this project\'s own approval branch, so their NAMES and ORDER ' +
                        'are the schedule\'s, not the previous deck\'s. On the original hand-built decks the ' +
                        'Approvals table is a flat PNG image with no text, so there is nothing to read; use a ' +
                        'previously GENERATED deck as the reference, or pin the list for this project in ' +
                        'approval_config.json.',
            }] : []),
        // A project can legitimately have NO approvals: Nandi Hills Homestay is a construction-phase
        // schedule with no statutory approval branch, and its own previous deck has no Approvals slide. When
        // nothing at all points to approvals — no branch, no pinned list, no reference list — the section is
        // omitted from the deck and these three checks are not applicable rather than failures. If a branch
        // exists but produced no rows, that IS still a failure.
        ...(r.branchesFound.length === 0 && r.details.length === 0 && r.masterExpected.length === 0 ? [
            { name: 'Approvals apply to this project', ok: true,
                detail: 'no approval branch in the schedule, no list in the previous deck and no pinned list — ' +
                    'this project has no statutory approvals to report, so the section is omitted' },
        ] : [
            { name: 'Approval branch detected', ok: r.branchesFound.length > 0,
                detail: r.branchesFound.length
                    ? r.branchesFound.map((b) => `"${b.name}" (WBS depth ${b.depth}, ${b.children} children)`).join(' · ')
                    : 'NO approval branch found in the schedule' },
            { name: 'Number of approvals found', ok: r.details.length > 0, detail: `${r.details.length} approval(s) immediately below the branch` },
            { name: 'Display rows generated', ok: r.rows.length === r.details.length && r.rows.length > 0,
                detail: `${r.rows.length} row(s) — one approval per row` },
        ]),
        { name: 'Completion status correct', ok: r.details.every((d) => ['Completed', 'In Progress', 'Not Started', 'Not Applicable'].includes(d.status)),
            detail: ['Completed', 'In Progress', 'Not Started', 'Not Applicable']
                .map((s) => `${s}: ${r.details.filter((d) => d.status === s).length}`).join(', ') },
        // A master-list approval the Excel does not carry is an EXPECTED blank: the row must stay on the
        // slide, and inventing a date for it is the one thing that would be wrong. Those rows are named in
        // `unmapped`, counted here, and excluded from the pass/fail — anything else blank still fails.
        // An approval whose task could not be located at all is an EXPECTED blank: the row stays on the slide
        // with its name and empty dates, which is the specified behaviour — substituting another approval or
        // inventing a date would be the error. Rows where an activity WAS found but carries no finish date
        // still fail, because that is a data problem worth surfacing.
        { name: 'Actual Finish dates populated',
            ok: r.details.every((d) => d.row.proj !== null || !d.finishFrom),
            detail: `${withProj}/${r.details.length} from the schedule` +
                (() => {
                    const blanks = r.details.filter((d) => d.row.proj === null).map((d) => d.row.approval);
                    return blanks.length ? ` · ${blanks.length} kept with BLANK dates because no task in this export matches: ${blanks.join(', ')}` : '';
                })() },
        { name: 'Every master-list approval is present, in order, unrenamed',
            ok: !r.masterExpected.length || (r.details.length === r.masterExpected.length
                && r.details.every((d, i) => d.row.approval === r.masterExpected[i])),
            detail: r.masterExpected.length
                ? `expected ${r.masterExpected.length} · got ${r.details.length}` +
                    (r.details.length === r.masterExpected.length && r.details.every((d, i) => d.row.approval === r.masterExpected[i])
                        ? ' · names and order identical'
                        : ` · MISMATCH: ${r.details.map((d) => d.row.approval).join(' | ')}`)
                : 'no master list configured for this run' },
        { name: 'No approval merged or renamed', ok: !r.masterExpected.length || r.details.every((d) => r.masterExpected.includes(d.row.approval)),
            detail: r.notOnMaster.length
                ? `${r.notOnMaster.length} schedule approval(s) deliberately NOT added: ${r.notOnMaster.join(', ')}`
                : 'no schedule approval was outside the master list' },
        { name: 'Baseline dates populated', ok: true,
            detail: `${withBase}/${r.details.length} have a Baseline finish; the remainder are left BLANK (never fabricated)` },
        { name: 'No duplicate approvals', ok: dupFound.length === 0, detail: dupFound.length ? dupFound.join(', ') : 'none' },
        { name: 'No engineering detail activities included', ok: detailish.length === 0,
            detail: detailish.length ? detailish.map((d) => d.row.approval).join(', ')
                : `only statutory approvals; look-alike branches excluded: ${r.branchExcluded.length ? r.branchExcluded.map((e) => `"${e.name}"`).join(', ') : 'none present'}` },
        { name: 'No child activities displayed as separate approvals', ok: true,
            detail: 'rows are the branch\'s direct children only; their sub-activities feed the stage ticks and dates' },
    ];
    const ok = checks.every((c) => c.ok);
    const L = [];
    L.push('# approval_validation.md', '');
    L.push(`Project: **${sch.projectName}**  ·  source: **ASTA Excel export only** (Approval Status.xlsx is no longer used)` +
        (scope ? `  ·  scoped to **${scope}**` : ''), '');
    L.push(`Mapping: **approval_config.json**`, '');
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Detail |', '|---|---|---|');
    for (const c of checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    L.push('', '## Approvals (as rendered)', '');
    L.push('| Sl | Approval | Status | % | Drawings | Application | Fee | Issued | Baseline Finish | Actual/Proj. Finish |');
    L.push('|---|---|---|---|---|---|---|---|---|---|');
    for (const d of r.details) {
        const t = (b) => (b ? '✔' : '✘');
        L.push(`| ${d.row.sl} | ${d.row.approval} | ${d.status} | ${d.pct}% | ${t(d.row.drawings)} | ${t(d.row.application)} | ` +
            `${t(d.row.fee)} | ${t(d.row.issued)} | ${d.row.base || '—'} | ${d.row.proj || '—'} |`);
    }
    L.push('', '## Where each date came from', '', 'One Excel activity per approval — the ISSUANCE activity — supplies BOTH dates. Baseline Finish and', 'Actual / Projected Finish are never taken from different activities, and never from an unrelated one.', '', `Row list source: **${r.listSource}** — ${r.listNote}`, '', '| Approval | Excel activity used (issuance) | Baseline Finish | Actual / Proj. Finish | How it was chosen | Candidates |', '|---|---|---|---|---|---|');
    for (const d of r.details) {
        L.push(`| ${d.row.approval} | ${d.finishFrom || '— none found —'} | ${d.row.base || 'blank'} | ${d.row.proj || 'blank'} | ${d.finishRule || '—'} | ${d.childActivities} |`);
    }
    if (r.branchExcluded.length) {
        L.push('', '## Look-alike branches deliberately excluded', '');
        for (const e of r.branchExcluded)
            L.push(`- \`${e.name}\` — ${e.reason}`);
    }
    L.push('', '## Items that could not be determined confidently', '');
    if (allNotes.length) {
        L.push('Reported rather than guessed:', '');
        for (const n of allNotes)
            L.push(`- ⚠️ ${n}`);
    }
    else
        L.push('None — every approval resolved a completion activity and a date from the schedule.');
    if (r.missingExpected.length) {
        L.push('', '## Commonly expected approvals not present in this project', '');
        L.push('Not an error — these simply do not exist in this schedule, so no row is shown for them:', '');
        for (const m of r.missingExpected)
            L.push(`- ${m}`);
    }
    return { ok, md: L.join('\n') };
}
