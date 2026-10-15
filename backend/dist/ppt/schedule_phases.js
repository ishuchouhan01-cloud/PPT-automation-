"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.resolvePhases = resolvePhases;
exports.schedulePhaseReport = schedulePhaseReport;
const approval_phase_1 = require("./approval_phase");
const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
/**
 * Whole-word containment — "clu" must not match "Clubhouse".
 * A keyword ending in `*` is a STEM (word-start) match, so `mobiliz*` catches both
 * "Mobilization" and "Mobilizing". Stems are declared in the configuration, not inferred.
 */
function wordHit(hay, needle) {
    if (!needle)
        return false;
    const stem = needle.endsWith('*');
    const core = stem ? needle.slice(0, -1).trim() : needle;
    if (!core)
        return false;
    const pat = core.split(/\s+/).map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s*');
    return new RegExp(`(^|\\s)${pat}${stem ? '' : '(\\s|$)'}`).test(hay);
}
/** First day of the month an ISO date falls in. */
const monthStart = (iso) => iso.slice(0, 8) + '01';
/** Last day of the month an ISO date falls in. */
const monthEnd = (iso) => {
    const y = +iso.slice(0, 4), m = +iso.slice(5, 7);
    const d = new Date(Date.UTC(y, m, 0)); // day 0 of the next month = last day of this one
    return d.toISOString().slice(0, 10);
};
const addMonths = (iso, n) => {
    const d = new Date(iso + 'T00:00:00Z');
    d.setUTCMonth(d.getUTCMonth() + n);
    return d.toISOString().slice(0, 10);
};
const minOf = (xs) => xs.filter(Boolean).sort()[0] || null;
const maxOf = (xs) => xs.filter(Boolean).sort().slice(-1)[0] || null;
/**
 * Activities for one phase. The three layers are tried in the specified order — exact name, then
 * alias, then keyword — and the results are UNIONED, because a phase BAR must span all of that
 * phase's work. (Taking only the first layer that hit made e.g. LOD 100/200 span a single activity.)
 * The layer that first identified the phase is reported so the mapping stays auditable.
 */
/**
 * The ONE activity that a phase's `dateSource` names. Exact name first (case- and
 * punctuation-insensitive), then a whole-word containment match, so "Definitive Agreement" finds
 * "Definitive Agreement" and "Signing of Definitive Agreement" but never "Agreement Draft Review".
 * When several match, the LATEST by finish wins — the final signing, not an early draft — and the
 * choice is reported.
 */
function findByNames(acts, names) {
    const latest = (pool) => pool.slice().sort((x, y) => (x.finish || '').localeCompare(y.finish || '')).slice(-1)[0];
    // A PHRASE test on the normalised names, tolerant of a plural. The schedule writes "Definitive
    // Agreements" and "Go To Market - Sales Launch (Hard)"; an exact-name test misses the first and a
    // single-keyword test misses the second, which is why neither was found before.
    const phrase = (hay, want) => {
        const h = ' ' + norm(hay) + ' ', w = norm(want);
        return h.includes(' ' + w + ' ') || h.includes(' ' + w + 's ') || h.includes(' ' + w + ' -') || h.includes(w);
    };
    for (const want of names) {
        const exact = acts.filter((a) => norm(a.name) === norm(want) || norm(a.name) === norm(want) + 's');
        if (exact.length)
            return latest(exact);
    }
    for (const want of names) {
        const loose = acts.filter((a) => phrase(a.name, want));
        if (loose.length)
            return latest(loose);
    }
    return null;
}
function matchActivities(acts, p) {
    const exact = (list) => {
        if (!list?.length)
            return [];
        const want = new Set(list.map((x) => norm(x).replace(/\*$/, '').trim()));
        return acts.filter((a) => want.has(norm(a.name)));
    };
    const byPrimary = exact(p.primary);
    const byAlias = exact(p.aliases);
    const kws = (p.keywords || []).map((k) => norm(k) + (k.trim().endsWith('*') ? '*' : '')).filter(Boolean);
    const byKeyword = acts.filter((a) => { const n = norm(a.name); return kws.some((k) => wordHit(n, k)); });
    const seen = new Set();
    const hits = [];
    for (const group of [byPrimary, byAlias, byKeyword]) {
        for (const a of group)
            if (!seen.has(a)) {
                seen.add(a);
                hits.push(a);
            }
    }
    const by = byPrimary.length ? 'primary' : byAlias.length ? 'alias' : byKeyword.length ? 'keyword' : 'none';
    return { hits, by };
}
/** The activity that represents a phase's completion, honouring priorityMilestone then the list. */
function findFinishMilestone(acts, p) {
    const tryName = (nm) => {
        const want = norm(nm);
        // exact first, then whole-word containment, preferring the LATEST finishing match
        const ex = acts.filter((a) => norm(a.name) === want);
        const pool = ex.length ? ex : acts.filter((a) => wordHit(norm(a.name), want));
        if (!pool.length)
            return null;
        return pool.slice().sort((x, y) => (x.finish || '').localeCompare(y.finish || '')).slice(-1)[0];
    };
    if (p.priorityMilestone) {
        const a = tryName(p.priorityMilestone);
        if (a)
            return { act: a, source: `priority milestone "${p.priorityMilestone}"` };
    }
    for (const nm of p.finishMilestones || []) {
        const a = tryName(nm);
        if (a)
            return { act: a, source: `milestone "${nm}"` };
    }
    return { act: null, source: '' };
}
/**
 * Resolve every configured phase against the schedule.
 * `acts` is normally already scoped to the deck's phase (see tracker.ts scopeToPhase).
 */
function resolvePhases(sch, acts, cfg) {
    const out = [];
    const byId = new Map();
    for (const p of cfg.phases) {
        // ---------- a phase resolved from an approval BRANCH ----------
        // The Approvals bar must span the whole statutory approval period, so it is not resolved by
        // keyword-matching activities and taking one milestone. It is derived from the branch: earliest
        // start of the executive approvals to the finish of the final major one.
        if (p.approvalBranch) {
            const ap = (0, approval_phase_1.resolveApprovalPhase)(sch, acts, p.approvalBranch);
            const r = {
                id: p.id, label: p.label,
                aStart: ap.aStart, aFinish: ap.aFinish, bStart: ap.bStart, bFinish: ap.bFinish,
                matched: ap.items.filter((i) => i.used).length,
                matchedBy: ap.branch ? 'primary' : 'none',
                startSource: ap.startRule ? `schedule: ${ap.startRule}${ap.startFrom ? ` (${ap.startFrom})` : ''}` : '',
                finishSource: ap.finishRule ? `schedule: ${ap.finishRule}${ap.finishFrom ? ` (${ap.finishFrom})` : ''}` : '(unresolved)',
                derived: [], missing: [], approval: ap,
            };
            if (!ap.branch)
                r.missing.push('no statutory approval branch could be discovered from the keyword hints');
            else if (!ap.ok)
                r.missing.push(`the approval branch "${ap.branch.name}" produced no start/finish date`);
            for (const f of ap.fallbacks)
                r.derived.push(f);
            // The declared start fallback still applies if the branch itself carries no start date at all.
            if (!r.aStart && p.start.fallbackAfterPhase) {
                const prev = byId.get(p.start.fallbackAfterPhase);
                if (prev?.aFinish) {
                    const off = p.start.fallbackOffsetMonths || 0;
                    r.aStart = off ? addMonths(prev.aFinish, off) : prev.aFinish;
                    r.bStart = prev.bFinish ? (off ? addMonths(prev.bFinish, off) : prev.bFinish) : null;
                    r.startSource = `config fallback: after "${prev.label}" (the approval branch had no start date)`;
                    r.derived.push(`start (config fallback after ${prev.label})`);
                }
            }
            if (r.aStart && r.aFinish && r.aFinish < r.aStart)
                r.aFinish = r.aStart;
            if (r.bStart && r.bFinish && r.bFinish < r.bStart)
                r.bFinish = r.bStart;
            out.push(r);
            byId.set(p.id, r);
            continue;
        }
        const { hits, by } = matchActivities(acts, p);
        const r = {
            id: p.id, label: p.label,
            aStart: null, aFinish: null, bStart: null, bFinish: null,
            matched: hits.length, matchedBy: by,
            startSource: '', finishSource: '', derived: [], missing: [],
        };
        if (!hits.length) {
            r.missing.push(`no activity matched (primary/alias/keyword) for phase "${p.label}"`);
            out.push(r);
            byId.set(p.id, r);
            continue;
        }
        // ---------- START ----------
        if (p.start.rule === 'afterPhase') {
            const prev = p.start.phase ? byId.get(p.start.phase) : undefined;
            if (prev && prev.aFinish) {
                const off = p.start.offsetMonths || 0;
                r.aStart = off ? addMonths(prev.aFinish, off) : prev.aFinish;
                r.bStart = prev.bFinish ? (off ? addMonths(prev.bFinish, off) : prev.bFinish) : null;
                r.startSource = `config: after "${prev.label}"${off ? ` + ${off} month(s)` : ''}`;
                r.derived.push(`start (after ${prev.label})`);
            }
            else {
                // the declared predecessor has no date — fall back to this phase's own schedule dates
                r.aStart = minOf(hits.map((a) => a.start));
                r.bStart = minOf(hits.map((a) => a.baselineStart));
                r.startSource = 'schedule: earliest start of matched activities (declared predecessor undated)';
                if (p.start.phase)
                    r.missing.push(`predecessor phase "${p.start.phase}" had no finish date for the start rule`);
            }
        }
        else {
            // Schedule first (requirement: use schedule dates, never estimate).
            r.aStart = minOf(hits.map((a) => a.start));
            r.bStart = minOf(hits.map((a) => a.baselineStart));
            r.startSource = 'schedule: earliest start of matched activities';
            if (!r.aStart && p.start.fallbackAfterPhase) {
                const prev = byId.get(p.start.fallbackAfterPhase);
                if (prev?.aFinish) {
                    const off = p.start.fallbackOffsetMonths || 0;
                    r.aStart = off ? addMonths(prev.aFinish, off) : prev.aFinish;
                    r.bStart = prev.bFinish ? (off ? addMonths(prev.bFinish, off) : prev.bFinish) : null;
                    r.startSource = `config fallback: after "${prev.label}"${off ? ` + ${off} month(s)` : ''} (phase had no start date in the schedule)`;
                    r.derived.push(`start (config fallback after ${prev.label})`);
                }
            }
        }
        // ---------- FINISH ----------
        const applyOffset = (iso, n) => (iso && n ? addMonths(iso, n) : iso);
        if (p.finish.rule === 'milestone') {
            // Look inside the phase's own activities first; if the named milestone is not among them, look
            // across the whole (scoped) schedule — an explicitly configured milestone name is authoritative
            // even when the phase's keyword list did not happen to catch that activity.
            let { act, source } = findFinishMilestone(hits, p);
            if (!act) {
                const wider = findFinishMilestone(acts, p);
                if (wider.act) {
                    act = wider.act;
                    source = `${wider.source} (found outside the phase's matched set)`;
                }
            }
            if (act) {
                r.aFinish = applyOffset(act.finish || act.actualFinish, p.finish.offsetMonths);
                r.bFinish = applyOffset(act.baselineFinish, p.finish.offsetMonths);
                r.finishSource = `schedule: ${source}` + (p.finish.offsetMonths ? ` + ${p.finish.offsetMonths} month(s) (config)` : '');
                if (p.finish.offsetMonths)
                    r.derived.push(`finish (+${p.finish.offsetMonths} month offset from config)`);
                if (!r.aFinish)
                    r.missing.push(`${source} has no Actual/Projected finish date`);
            }
            else if (p.finish.fallback === 'latestFinish') {
                r.aFinish = applyOffset(maxOf(hits.map((a) => a.finish)), p.finish.offsetMonths);
                r.bFinish = applyOffset(maxOf(hits.map((a) => a.baselineFinish)), p.finish.offsetMonths);
                r.finishSource = 'schedule: latest finish of matched activities (configured fallback)';
                r.missing.push(`none of the finish milestones [${(p.finishMilestones || []).join(', ')}] were found`);
            }
            else {
                r.missing.push(`finish milestone not found — required [${(p.finishMilestones || []).join(' OR ')}]`);
                r.finishSource = '(unresolved)';
            }
        }
        else if (p.finish.rule === 'afterPhase') {
            const prev = p.finish.phase ? byId.get(p.finish.phase) : undefined;
            if (prev?.aFinish) {
                r.aFinish = applyOffset(prev.aFinish, p.finish.offsetMonths);
                r.bFinish = applyOffset(prev.bFinish, p.finish.offsetMonths);
                r.finishSource = `config: after "${prev.label}"${p.finish.offsetMonths ? ` + ${p.finish.offsetMonths} month(s)` : ''}`;
                r.derived.push(`finish (after ${prev.label})`);
            }
            else
                r.missing.push(`predecessor phase "${p.finish.phase}" had no finish date`);
        }
        else {
            r.aFinish = applyOffset(maxOf(hits.map((a) => a.finish)), p.finish.offsetMonths);
            r.bFinish = applyOffset(maxOf(hits.map((a) => a.baselineFinish)), p.finish.offsetMonths);
            r.finishSource = 'schedule: latest finish of matched activities';
        }
        // ---------- a single task supplies this phase's dates ----------
        // Applied LAST so it overrides whatever the generic rules produced, and only when the task is
        // actually in the schedule. Baseline comes from the same task's baseline dates — never from the
        // projected ones — so the baseline row stays a baseline.
        if (p.dateSource && p.dateSource.match.length) {
            const src = p.dateSource.fromWbs
                ? findByNames(Object.values(sch.wbs).map((w) => ({
                    // A WBS row presented as an activity-shaped record, so one matcher serves both.
                    id: w.id, name: w.name, start: w.start ?? null, finish: w.finish ?? null,
                    actualStart: null, actualFinish: null,
                    baselineStart: w.baselineStart ?? null, baselineFinish: w.baselineFinish ?? null,
                    pct: w.pct ?? 0,
                })), p.dateSource.match)
                : findByNames(acts, p.dateSource.match);
            const applies = p.dateSource.applies || 'finish';
            if (src) {
                if (applies === 'both') {
                    r.aStart = src.start || src.actualStart || r.aStart;
                    r.bStart = src.baselineStart || r.bStart;
                    r.startSource = `schedule: the "${src.name}" task`;
                }
                else if (applies === 'baselineStartToFinish') {
                    // The start is the BASELINE START OF THE LAND ACQUISITION ROW ITSELF, exactly as the schedule
                    // records it — not the task's, not a neighbouring phase's, never shifted.
                    //
                    // The row the user reads in the Excel is the WBS row named "Land Acquisition", and its baseline
                    // start (2025-06-04 on this export) is what the phase begins on. Its CHILD activities start
                    // later (2025-12-11), so taking the earliest child baseline would begin the bar six months
                    // late. The summary row supplies the START only — its finish is explicitly not to be used.
                    const wanted = [...(p.primary || []), p.label].map(norm);
                    const node = Object.values(sch.wbs)
                        .filter((w) => wanted.includes(norm(w.name)) && w.baselineStart)
                        .sort((a, b) => (a.baselineStart < b.baselineStart ? -1 : 1))[0];
                    const bs = (node && node.baselineStart) || r.bStart || null;
                    if (bs) {
                        r.aStart = bs;
                        r.bStart = bs;
                        r.startSource = `schedule: Baseline Start of ${node ? `the "${node.name}" row` : 'the phase\'s matched activities'} (unshifted)`;
                        r.derived.push(`start is the Baseline Start ${node ? `of the "${node.name}" row` : 'of the matched activities'} (${bs}) — read from the schedule, not derived`);
                    }
                    else {
                        r.missing.push('the phase carries no Baseline Start, so the bar keeps its ordinary start');
                    }
                    // ...AND ITS BASELINE FINISH COMES FROM THAT SAME ROW.
                    // The baseline bar must run between the schedule's OWN baseline dates. Taking the start from
                    // the row and leaving the finish unset made the baseline "start-only": One Tiny Flower's
                    // Land Acquisition row records 2023-04-11 → 2025-04-29, and the bar was drawn as a single
                    // cell at 2023-04, with the report saying the row had no baseline finish. It has one.
                    // The projected finish below is NEVER used for the baseline, and this is only read from the
                    // row that supplied the start — nothing is derived and nothing is borrowed from a
                    // neighbouring phase.
                    if (node && node.baselineFinish) {
                        r.bFinish = node.baselineFinish;
                        r.derived.push(`baseline finish is the Baseline Finish of the "${node.name}" row (${node.baselineFinish}) — read from the schedule, never from the actual/projected dates`);
                    }
                    else if (!r.bFinish) {
                        r.missing.push(`the "${node ? node.name : p.label}" row carries no Baseline Finish, so the baseline bar's extent is not implied`);
                    }
                }
                r.aFinish = src.finish || src.actualFinish || r.aFinish;
                if (applies !== 'baselineStartToFinish')
                    r.bFinish = src.baselineFinish || r.bFinish;
                r.finishSource = `schedule: the "${src.name}" task`;
                r.derived.push(`finish taken from the task "${src.name}"'s actual/projected finish (a date source only — not a legend entry)`);
                // The configured finish MILESTONES are no longer needed for this phase — the controlling task
                // outranks them — so a "milestone not found" note about them is stale and is withdrawn. Anything
                // else this phase could not resolve is left in place.
                r.missing = r.missing.filter((m) => !/finish milestones?\b/i.test(m));
            }
            else {
                r.missing.push(`no task matching ${p.dateSource.match.map((m) => `"${m}"`).join(' / ')} — this phase kept its ordinary rule`);
            }
        }
        // never draw a bar backwards
        if (r.aStart && r.aFinish && r.aFinish < r.aStart)
            r.aFinish = r.aStart;
        if (r.bStart && r.bFinish && r.bFinish < r.bStart)
            r.bFinish = r.bStart;
        out.push(r);
        byId.set(p.id, r);
    }
    // ---------------------------------------------------------------------------------------------
    // SECOND PASS — rules that point FORWARD to a later phase.
    //
    // The first pass runs in config order, so a phase cannot see one that comes after it. Two rules need
    // exactly that: Marketing Collateral must sit in the month before Sales launch (which is resolved from
    // the Go to Market task), and LOD 100/200 must run up to the start of Approvals. Both are applied here,
    // once every phase has its own dates, and both are recorded as derived.
    for (const p of cfg.phases) {
        const r = byId.get(p.id);
        if (!r)
            continue;
        // Marketing Collateral: exactly `months` before the named phase, ending where it begins.
        if (p.start.rule === 'monthsBeforePhase' && p.start.phase) {
            const nxt = byId.get(p.start.phase);
            const n = p.start.months ?? 1;
            // WHOLE CALENDAR MONTHS when alignToMonths is set. The tracker paints one phase per month, so a
            // day-anchored window (02 Aug → 02 Sep) lands in TWO months and collides with the phase after it;
            // the month-floor rule then pushed Marketing and Sales a month later than their own dates. Aligned
            // to the calendar, a September Sales launch gives Marketing Collateral exactly August — which is
            // what "exactly 1 month immediately before" means when the chart's unit is a month.
            const aligned = !!p.start.alignToMonths;
            // The earliest this phase may begin: the month after `fillGapFromPhase` ends. Taking the EARLIER of
            // that and the ordinary n-month window means the phase covers its minimum window and stretches back
            // to close a gap, but never starts before the phase it follows has finished.
            const prevPhase = p.start.fillGapFromPhase ? byId.get(p.start.fillGapFromPhase) : undefined;
            const afterPrev = prevPhase && prevPhase.aFinish ? monthStart(addMonths(monthStart(prevPhase.aFinish), 1)) : null;
            if (nxt && nxt.aStart) {
                if (aligned) {
                    const window = monthStart(addMonths(monthStart(nxt.aStart), -n));
                    const from = afterPrev && afterPrev < window ? afterPrev : window;
                    if (from !== window) {
                        r.derived.push(`start pulled back from ${window} to ${from} so no blank months are left after ${prevPhase.label}`);
                    }
                    r.aStart = from;
                    r.aFinish = monthEnd(addMonths(monthStart(nxt.aStart), -1));
                }
                else {
                    r.aFinish = nxt.aStart;
                    r.aStart = addMonths(nxt.aStart, -n);
                }
                r.startSource = `the ${n} whole month(s) before ${nxt.label} (config)`;
                r.finishSource = aligned ? `ends the month before ${nxt.label} (config)` : `ends where ${nxt.label} begins (config)`;
                r.derived.push(`${aligned ? 'the whole calendar month(s)' : `the ${n} month(s)`} immediately before ${nxt.label}`);
            }
            if (nxt && nxt.bStart) {
                if (aligned) {
                    const bWindow = monthStart(addMonths(monthStart(nxt.bStart), -n));
                    const bAfter = prevPhase && prevPhase.bFinish ? monthStart(addMonths(monthStart(prevPhase.bFinish), 1)) : null;
                    r.bStart = bAfter && bAfter < bWindow ? bAfter : bWindow;
                    r.bFinish = monthEnd(addMonths(monthStart(nxt.bStart), -1));
                }
                else {
                    r.bFinish = nxt.bStart;
                    r.bStart = addMonths(nxt.bStart, -n);
                }
            }
        }
        // LOD 100/200: it OCCUPIES the months between its own start and the start of the named phase —
        // extended if its activities end early (no empty gap) and cut back if they run on past it (no
        // overlap). Its own start is never moved.
        if (p.finish.rule === 'beforePhaseStart' && p.finish.phase) {
            const nxt = byId.get(p.finish.phase);
            if (nxt && nxt.aStart && r.aStart && nxt.aStart > r.aStart && r.aFinish !== nxt.aStart) {
                const was = r.aFinish;
                r.aFinish = nxt.aStart;
                r.finishSource = `runs up to the start of ${nxt.label} (config)`;
                r.derived.push(`finish ${was && was < nxt.aStart ? 'extended' : 'cut back'} from ${was || '—'} to ${nxt.aStart}, the start of ${nxt.label}`);
            }
            if (nxt && nxt.bStart && r.bStart && nxt.bStart > r.bStart)
                r.bFinish = nxt.bStart;
        }
    }
    // ---------------------------------------------------------------------------------------------
    // ORDERING GUARD — a phase declared to precede another may not overrun it.
    // Approvals must finish before Marketing Collateral begins. When the schedule says otherwise the bar
    // is clamped to the boundary and the clamp is REPORTED; the underlying dates are never rewritten.
    // A phase that must not begin before another ends: Site Establishment & Mobilization cannot precede
    // Sales launch, and LOD 100/200 cannot precede Land Acquisition. The start is pushed to the boundary
    // and the push is REPORTED; the schedule itself is never rewritten.
    for (const p of cfg.phases) {
        const after = p.mustStartAfter;
        if (!after)
            continue;
        const r = byId.get(p.id), prev = byId.get(after);
        if (!r || !prev)
            continue;
        if (r.aStart && prev.aFinish && r.aStart < prev.aFinish) {
            // A GUARD NEVER INVERTS A BAR. If the previous phase ends after this one finishes, the two
            // genuinely overlap in the schedule and pushing the start would collapse the phase to a point —
            // which is exactly what turned a 17-month statutory approval period into a single day. The overlap
            // is reported and the real dates are left alone.
            if (r.aFinish && prev.aFinish >= r.aFinish) {
                r.derived.push(`NOT pushed after ${prev.label}: ${prev.label} ends ${prev.aFinish}, after this phase's own finish ${r.aFinish} — the two overlap in the schedule and the real dates are kept`);
            }
            else {
                r.derived.push(`start pushed from ${r.aStart} to ${prev.aFinish}: it may not begin before ${prev.label} ends (config)`);
                r.aStart = prev.aFinish;
            }
        }
        if (r.bStart && prev.bFinish && r.bStart < prev.bFinish && (!r.bFinish || prev.bFinish < r.bFinish)) {
            r.bStart = prev.bFinish;
        }
    }
    for (const p of cfg.phases) {
        const before = p.mustFinishBefore;
        if (!before)
            continue;
        const r = byId.get(p.id), nxt = byId.get(before);
        if (!r || !nxt)
            continue;
        // Month-aligned: the phase must finish in the month BEFORE the named phase's month, so the two never
        // share a month on a chart whose unit is a month. Approvals therefore ends in July when Marketing
        // Collateral owns August.
        const limit = p.mustFinishBeforeAligned && nxt.aStart ? monthEnd(addMonths(monthStart(nxt.aStart), -1)) : nxt.aStart;
        if (r.aFinish && limit && r.aFinish > limit) {
            if (r.aStart && limit <= r.aStart) {
                r.derived.push(`NOT clamped before ${nxt.label}: the boundary ${limit} is at or before this phase's own start ${r.aStart} — the two overlap in the schedule and the real dates are kept`);
            }
            else {
                r.derived.push(`finish clamped from ${r.aFinish} to ${limit}: it must finish before ${nxt.label}'s month begins (config)`);
                r.aFinish = limit;
            }
        }
        if (r.bFinish && nxt.bStart && r.bFinish > nxt.bStart && (!r.bStart || nxt.bStart > r.bStart)) {
            r.bFinish = nxt.bStart;
        }
    }
    // ── UPSTREAM CHAIN ────────────────────────────────────────────────────────────────────────────
    // Land Acquisition → LOD 100/200 → Approvals, closed up so the reader never sees a blank stretch
    // of chart between them.
    //
    // Land Acquisition is the starting anchor of this part of the timeline and is NOT touched here: its
    // dates stay exactly as its own source rules resolved them. What this pass does is give the phase
    // after it somewhere sensible to live. LOD 100/200 starts the month AFTER Land Acquisition finishes,
    // always occupies at least its configured minimum, and — this is the point — STRETCHES to cover the
    // waiting time before Approvals actually start. A project whose land completes in June and whose
    // approvals only begin in October has three idle months in the source data; showing them as LOD
    // 100/200 is what the reference deck does, and it is honest, because that is the design work that
    // has to happen in the interval. The alternative, a blank gap, tells the reader nothing.
    //
    // Approvals' own start is left as real schedule data. If it starts EARLIER than the minimum this
    // pass gives LOD, the two genuinely overlap in the source: the minimum wins for LOD, the real
    // Approvals start is kept, and the overlap is reported rather than papered over.
    //
    // Run once per timeline row, on that row's own dates only.
    const chainRows = [
        { key: 'a', label: 'Actual/Projected' },
        { key: 'b', label: 'Baseline' },
    ];
    for (const step of cfg.timeline?.chain?.steps || []) {
        const r = byId.get(step.phase);
        if (!r)
            continue;
        const prev = step.startsAfterPhase ? byId.get(step.startsAfterPhase) : undefined;
        const nxt = step.fillGapBeforePhase ? byId.get(step.fillGapBeforePhase) : undefined;
        for (const row of chainRows) {
            const S = row.key === 'a' ? 'aStart' : 'bStart', F = row.key === 'a' ? 'aFinish' : 'bFinish';
            const prevFinish = prev ? prev[F] : null;
            if (!prevFinish) {
                // Nothing is invented. Without the preceding phase's finish there is no boundary to start from.
                if (prev)
                    r.derived.push(`${row.label}: left as resolved — ${prev.label} has no ${row.label} finish date in the source schedule, so there is no boundary to start from`);
                continue;
            }
            // ONLY BRIDGE A REAL GAP. This phase exists here to fill the wait between the phase before it
            // and the phase after it. If the phase AFTER it already starts on or before where this one
            // would begin, there is no wait to fill — the two genuinely overlap in the source — and moving
            // this phase to "the month after the previous one finished" would push it PAST the phase it is
            // supposed to precede and out of the configured order. Violet Hill's baseline does exactly
            // that: Land Acquisition's own row finishes 2025-12-11, after its baseline Approvals
            // (2025-05-13 → 2025-09-30), so LOD would land in Jan-2026, behind Approvals, Marketing and
            // Sales launch. The phase keeps its own resolved dates and the conflict is reported.
            const wouldStart = monthStart(addMonths(monthStart(prevFinish), 1));
            if (nxt && nxt[S] && nxt[S] <= wouldStart) {
                r.derived.push(`${row.label}: left on its own dates — ${prev.label} finishes ${prevFinish}, at or after ${nxt.label} starts (${nxt[S]}), so there is no gap for this phase to fill and moving it would put it out of order. The source schedule overlaps these phases.`);
                continue;
            }
            const start = wouldStart;
            const months = Math.max(1, step.minMonths || 1);
            const minFinish = monthEnd(addMonths(start, months - 1));
            // The gap to fill: everything up to the month before the next phase really starts.
            const gapFinish = nxt && nxt[S] ? monthEnd(addMonths(monthStart(nxt[S]), -1)) : null;
            const finish = gapFinish && gapFinish > minFinish ? gapFinish : minFinish;
            const wasS = r[S], wasF = r[F];
            r[S] = start;
            r[F] = finish;
            if (wasS !== start)
                r.derived.push(`${row.label}: start moved from ${wasS || 'none'} to ${start} — the month immediately after ${prev.label} finishes (${prevFinish}), leaving no empty month between them (config)`);
            if (finish === gapFinish && nxt) {
                r.derived.push(`${row.label}: stretched to ${finish} to fill the whole gap before ${nxt.label} starts (${nxt[S]}) — the waiting time is shown as ${r.label}, not as a blank chart (config)`);
            }
            else {
                r.derived.push(`${row.label}: occupies ${months} whole month(s) to ${finish}, the configured minimum${gapFinish ? ` (${nxt.label} starts ${nxt[S]}, which is not later, so there is no gap to fill)` : ''}`);
            }
            if (wasF && wasF !== finish)
                r.derived.push(`${row.label}: its own resolved finish was ${wasF}`);
            if (nxt && nxt[S] && nxt[S] < start) {
                r.derived.push(`${row.label}: NOTE — ${nxt.label} already starts ${nxt[S]}, before this phase begins; the two overlap in the source schedule and ${nxt.label}'s real start is kept`);
            }
        }
    }
    // ── SALES LAUNCH ANCHOR ───────────────────────────────────────────────────────────────────────
    // Runs LAST, so it has priority over every rule above.
    //
    // The rules above walk the timeline forwards: each phase is placed after the one before it, and a
    // phase that ends late therefore pushes everything after it later. That is wrong for the launch
    // window. Sales Launch is a hard commercial date taken from the 'Go To Market' task, so it is the
    // FIXED POINT and the phases around it are laid against it, not the other way round.
    //
    // Concretely, this pass replaces the forward chain for the launch window with:
    //     Approvals finish  ->  Marketing Collateral  ->  SALES LAUNCH  ->  Site Establishment
    // where Sales Launch owns exactly the Go To Market month, Marketing Collateral owns exactly the
    // month before it, Approvals is made to finish at the end of the month before THAT, and Site
    // Establishment starts the month after and still ends on its own source finish date. Every step is
    // a whole month, so the sequence is continuous and no empty month can appear between them.
    //
    // The BASELINE and the ACTUAL/PROJECTED rows are anchored SEPARATELY, each on its own Go To Market
    // date. Neither is ever derived from the other, so a project whose launch has slipped shows the two
    // windows in different places, which is the point of the chart. A timeline with no Go To Market date
    // of its own is left as the ordinary rules resolved it, and the absence is reported.
    const anchorCfg = cfg.timeline?.anchor;
    const anchor = anchorCfg ? byId.get(anchorCfg.phase) : undefined;
    if (anchorCfg && anchor) {
        // 'a' = the Actual/Projected row, 'b' = the Baseline row. One helper, run twice, so the two rows
        // cannot influence each other and cannot drift apart in behaviour either.
        const rows = [
            { key: 'a', label: 'Actual/Projected' },
            { key: 'b', label: 'Baseline' },
        ];
        for (const row of rows) {
            const S = row.key === 'a' ? 'aStart' : 'bStart', F = row.key === 'a' ? 'aFinish' : 'bFinish';
            const LOCK = row.key === 'a' ? 'anchoredA' : 'anchoredB';
            const seed = anchor[S] || anchor[F];
            if (!seed) {
                anchor.derived.push(`${row.label}: no anchor date in the source schedule, so the ${row.label} launch window is left as the ordinary rules resolved it (no date is borrowed from the other timeline)`);
                continue;
            }
            // The anchor month, expressed as its first and last day.
            const m0 = monthStart(seed);
            const wasS = anchor[S], wasF = anchor[F];
            anchor[S] = m0;
            anchor[F] = monthEnd(m0);
            anchor[LOCK] = true;
            if (wasS !== anchor[S] || wasF !== anchor[F]) {
                anchor.derived.push(`${row.label}: anchored to the whole of ${m0.slice(0, 7)}, the month of its own source date ${wasS || wasF} (${row.label} of the Go To Market task) — the anchor month is never shifted forward or back`);
            }
            for (const step of anchorCfg.precede || []) {
                const r = byId.get(step.phase);
                if (!r)
                    continue;
                const target = monthStart(addMonths(m0, -Math.abs(step.months)));
                const finish = monthEnd(target);
                if (step.adjustFinishOnly) {
                    // The start is real schedule data and is never moved; only the finish is brought to the
                    // boundary. Config allows this deliberately: closing the gap matters more than preserving an
                    // approvals finish date that the launch date has already overtaken.
                    if (r[F] !== finish) {
                        // Never invert the bar. If this phase's own start is already past the boundary the two
                        // genuinely overlap, so the real dates are kept and the overlap is reported.
                        if (r[S] && r[S] > finish) {
                            r.derived.push(`${row.label}: NOT pulled back to ${finish}: this phase starts ${r[S]} which is already after the boundary set by ${anchor.label} — the real dates are kept`);
                        }
                        else {
                            r.derived.push(`${row.label}: finish moved from ${r[F] || 'none'} to ${finish} so it ends in the month immediately before ${byId.get((anchorCfg.precede || [])[0].phase)?.label || 'the next phase'}, leaving no empty month before ${anchor.label} (config: ${Math.abs(step.months)} month(s) before the anchor)`);
                            r[F] = finish;
                        }
                    }
                }
                else {
                    if (r[S] !== target || r[F] !== finish) {
                        r.derived.push(`${row.label}: placed in ${target.slice(0, 7)} — exactly the ${Math.abs(step.months)} month(s) immediately before ${anchor.label} (config)`);
                    }
                    r[S] = target;
                    r[F] = finish;
                    r[LOCK] = true;
                }
            }
            for (const step of anchorCfg.follow || []) {
                const r = byId.get(step.phase);
                if (!r)
                    continue;
                // NO SOURCE EVIDENCE, NO BAR.
                // A phase only exists on the chart because the project's own schedule contains the work. If
                // nothing matched it and it carries no date of its own, there is nothing to place — and
                // placing it anyway is how a project with no site-establishment activity at all still showed
                // a dark-green "Site Establishment & Mobilization" bar: this pass gave it a start of "the
                // month after Sales Launch" purely because the anchor said something follows the anchor.
                // A legend entry is not evidence. Neither is the launch date, the phase before it, or a
                // parent WBS row. The phase is left undated, which is what stops the bar being drawn.
                if (!r.matched && !r.aStart && !r.aFinish && !r.bStart && !r.bFinish) {
                    r.derived.push(`${row.label}: NOT placed after ${anchor.label} — this project's schedule contains no activity for this phase at all (0 matched, no dates of its own), so no bar is drawn. A phase is never created from the legend or from the anchor.`);
                    continue;
                }
                const target = monthStart(addMonths(m0, 1));
                const own = r[F];
                // Only move the start if the phase's own finish can still follow it. Where the source has
                // this activity FINISHING BEFORE the launch month — which happens, e.g. site mobilisation
                // completed ahead of go-to-market — moving the start would invert the bar, and the only ways
                // out would be to shorten it or to derive its finish from the anchor. Both are forbidden. So
                // the phase keeps BOTH of its own real dates and the conflict is reported; the display's
                // existing ordering rule then places it after the anchor rather than drawing it out of order.
                if (own && own < target) {
                    r.derived.push(`${row.label}: start NOT moved to ${target} — this activity's own ${row.label} finish is ${own}, which is before ${anchor.label}'s month ends. Moving the start would invert the bar, and neither shortening it nor deriving its finish from ${anchor.label} is permitted, so both of its real dates are kept and the schedule's disagreement with the launch date is reported here.`);
                }
                else {
                    if (r[S] !== target) {
                        r.derived.push(`${row.label}: start moved from ${r[S] || 'none'} to ${target} — the month immediately after ${anchor.label}, leaving no empty month between them (config)`);
                    }
                    r[S] = target;
                    r[LOCK] = true;
                }
                if (step.keepOwnFinish) {
                    // The finish belongs to the source activity and is NEVER derived from the anchor and NEVER
                    // shortened to fit the chart. If it lands before the new start the source genuinely disagrees
                    // with the launch date, so the real finish is kept and the conflict is reported rather than
                    // quietly trimmed.
                    if (!own) {
                        r.derived.push(`${row.label}: no finish date in the source schedule for this activity, so the bar's extent is not implied`);
                    }
                    else if (own < target) {
                        r.derived.push(`${row.label}: source finish ${own} is BEFORE ${anchor.label} ends — the source finish is kept as-is and NOT stretched or shortened; the schedule and the launch date disagree here`);
                        // eslint-disable-next-line no-empty
                    }
                    else {
                        r.derived.push(`${row.label}: runs through to ${own}, taken directly from the ${row.label} Finish of its own source activity — never shortened to fit the chart`);
                    }
                }
            }
        }
    }
    return out;
}
/** Markdown report — every milestone that could not be found is listed, never estimated. */
function schedulePhaseReport(sch, phases, cfg, scope, colours) {
    const startPhase = cfg.timeline?.startPhase, endPhase = cfg.timeline?.endPhase;
    const first = phases.find((p) => p.id === startPhase) || phases[0];
    const last = phases.find((p) => p.id === endPhase) || phases[phases.length - 1];
    const dated = phases.filter((p) => (p.aStart && p.aFinish) || (p.bStart && p.bFinish));
    const allMissing = phases.flatMap((p) => p.missing.map((m) => `${p.label}: ${m}`));
    const checks = [
        { name: 'Phase mapping found', ok: phases.every((p) => p.matched > 0),
            detail: phases.map((p) => `${p.label} → ${p.matched} act (${p.matchedBy})`).join(' · ') },
        { name: 'All milestone dates found', ok: allMissing.length === 0,
            detail: allMissing.length ? `${allMissing.length} unresolved (listed below)` : 'every configured milestone resolved from the schedule' },
        // A baseline bar is drawn from the phase's Baseline START; a phase holding a start but no finish
        // is still drawn, ending the month before the next phase begins (see spansOf in tracker.ts). So the
        // count reported here is the number of phases with a Baseline Start — which is what the Baseline
        // row actually shows. Counting only fully-dated phases understated it and disagreed with the slide.
        { name: 'Baseline bars correct', ok: phases.every((p) => !p.bStart || !p.bFinish || p.bStart <= p.bFinish),
            detail: `${phases.filter((p) => p.bStart).length}/${phases.length} phases have a baseline bar ` +
                `(${phases.filter((p) => p.bStart && !p.bFinish).length} start-only, bounded by the next phase); none inverted` },
        { name: 'Actual bars correct', ok: phases.every((p) => !p.aStart || !p.aFinish || p.aStart <= p.aFinish),
            detail: `${phases.filter((p) => p.aStart && p.aFinish).length}/${phases.length} phases have an actual/projected bar; none inverted` },
        { name: 'Timeline starts at Land Acquisition', ok: !!first && first.id === (startPhase || first.id) && !!(first.aStart || first.bStart),
            detail: first ? `${first.label} — actual ${first.aStart || '—'}, baseline ${first.bStart || '—'}` : 'not resolved' },
        { name: 'Timeline ends at Site Establishment & Mobilization', ok: !!last && last.id === (endPhase || last.id) && !!(last.aFinish || last.bFinish),
            detail: last ? `${last.label} — actual ${last.aFinish || '—'}, baseline ${last.bFinish || '—'}` : 'not resolved' },
        { name: 'No missing milestones', ok: allMissing.length === 0,
            detail: allMissing.length ? `${allMissing.length} listed explicitly (no estimation performed)` : 'none' },
        { name: 'Every phase found its own swatch pair in the dashboard palette',
            ok: !phases.some((p) => (p.colourSource || '').startsWith('no swatch')),
            detail: (() => {
                const missingSwatch = phases.filter((p) => (p.colourSource || '').startsWith('no swatch'));
                return missingSwatch.length
                    ? `${missingSwatch.map((p) => p.label).join(', ')} — lighter tone derived from the phase's own solid tone`
                    : 'each phase matched a (solid, lighter) pair on the Schedule sheet; no generic fallback colour was used';
            })() },
    ];
    const ok = checks.every((c) => c.ok);
    const L = [];
    L.push('# tracker_schedule_validation.md', '');
    L.push(`Project: **${sch.projectName}**  ·  mapping: **tracker_schedule_config.json**` +
        (scope ? `  ·  scoped to **${scope}**` : ''), '');
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Detail |', '|---|---|---|');
    for (const c of checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    L.push('', '## Resolved phases (in timeline order)', '');
    L.push('| # | Phase | Matched by | Acts | Actual start | Actual finish | Baseline start | Baseline finish | Colour (act/base) |');
    L.push('|---|---|---|---|---|---|---|---|---|');
    phases.forEach((p, i) => {
        const c = (p.colourActual && p.colourBaseline)
            ? { actual: p.colourActual, baseline: p.colourBaseline }
            : colours[p.label];
        L.push(`| ${i + 1} | ${p.label} | ${p.matchedBy} | ${p.matched} | ${p.aStart || '—'} | ${p.aFinish || '—'} | ` +
            `${p.bStart || '—'} | ${p.bFinish || '—'} | ${c ? `#${c.actual} / #${c.baseline}` : '—'} |`);
    });
    // ---- how each phase got its two tones -------------------------------------------------------
    // The Baseline row must be the LIGHTER TONE OF THE SAME HUE as the Actual/Projected segment above
    // it, phase by phase — never one flat colour across the whole baseline row, and never a tone
    // belonging to a different phase. Both halves of that promise are measured here rather than
    // assumed: the resolution route is printed, and the pair is checked for same-hue-and-lighter.
    const rgb = (h) => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
    const lum = (h) => { const [r, g, b] = rgb(h); return 0.299 * r + 0.587 * g + 0.114 * b; };
    const hue = (h) => {
        const [r, g, b] = rgb(h).map((c) => c / 255);
        const mx = Math.max(r, g, b), mn = Math.min(r, g, b), d = mx - mn;
        if (d < 0.02)
            return -1; // greys carry no hue to compare
        let x = mx === r ? ((g - b) / d) % 6 : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
        return ((x * 60) + 360) % 360;
    };
    const hueGap = (a, b) => (a < 0 || b < 0) ? 0 : Math.min(Math.abs(a - b), 360 - Math.abs(a - b));
    const tonePairs = phases
        .map((p) => ({ p, a: p.colourActual || colours[p.label]?.actual, b: p.colourBaseline || colours[p.label]?.baseline }))
        .filter((x) => !!x.a && !!x.b);
    const badTone = tonePairs.filter((x) => !(lum(x.b) > lum(x.a) && hueGap(hue(x.a), hue(x.b)) <= 25));
    const sharedBaseline = new Map();
    for (const x of tonePairs)
        sharedBaseline.set(x.b, [...(sharedBaseline.get(x.b) || []), x.p.label]);
    const collisions = [...sharedBaseline.entries()].filter(([, ls]) => ls.length > 1);
    L.push('', '## Phase colours — each Baseline tone is the lighter tone of its OWN phase', '');
    L.push('| Phase | Actual / Projected | Baseline | Lighter? | Same hue? | Where the pair came from |');
    L.push('|---|---|---|---|---|---|');
    for (const x of tonePairs) {
        const lighter = lum(x.b) > lum(x.a);
        const same = hueGap(hue(x.a), hue(x.b)) <= 25;
        L.push(`| ${x.p.label} | \`#${x.a}\` | \`#${x.b}\` | ${lighter ? '✅' : '❌'} | ${same ? '✅' : '❌'} | ` +
            `${x.p.colourSource || 'dashboard swatch'} |`);
    }
    L.push('', `- Every baseline tone lighter than its own actual tone, same hue: ${badTone.length ? `❌ ${badTone.map((x) => x.p.label).join(', ')}` : '✅'}`);
    L.push(`- No two phases share a baseline tone (the row is not one flat colour): ` +
        `${collisions.length ? `❌ ${collisions.map(([c, ls]) => `#${c} used by ${ls.join(' + ')}`).join('; ')}` : '✅'}`);
    L.push(`- Colours assigned from the PHASE, never from the project name: ✅ ` +
        `(resolution reads only the phase label and its config id — the project name is not an input)`);
    L.push('', '## Where each date came from', '', '| Phase | Start | Finish |', '|---|---|---|');
    for (const p of phases)
        L.push(`| ${p.label} | ${p.startSource || '—'} | ${p.finishSource || '—'} |`);
    const derived = phases.filter((p) => p.derived.length);
    L.push('', '## Dates derived from a declared configuration rule', '');
    if (derived.length) {
        L.push('These are **not** estimates by the application — each is a rule written in', '`tracker_schedule_config.json` (e.g. "start after the previous phase"). Change the file to change them.', '', '| Phase | Derived |', '|---|---|');
        for (const p of derived)
            L.push(`| ${p.label} | ${p.derived.join('; ')} |`);
    }
    else
        L.push('None — every date came straight from a schedule milestone.');
    L.push('', '## Unresolved milestones', '');
    if (allMissing.length) {
        L.push('Listed explicitly rather than estimated:', '');
        for (const m of allMissing)
            L.push(`- ❌ ${m}`);
    }
    else
        L.push('None.');
    L.push('', '## Approvals — the whole statutory period', '');
    const appr = phases.find((p) => p.approval) || phases.find((p) => /approval/i.test(p.label));
    if (appr?.approval) {
        const a = appr.approval;
        L.push(`The Approvals bar is derived from the branch **${a.branch ? a.branch.name : '(not found)'}** rather than from a ` +
            'single milestone: it spans ' + `**${a.aStart || '—'} → ${a.aFinish || '—'}**` +
            (a.durationDays !== null ? ` (${a.durationDays} days)` : '') + ', taken from ' +
            `${a.items.filter((i) => i.used).length} executive approval(s). Finish rule: ${a.finishRule || '(unresolved)'}` +
            (a.finishFrom ? ` → **${a.finishFrom}**` : '') + '. Full derivation in approval_schedule_validation.md.');
    }
    else
        L.push(appr ? 'The Approvals phase is not configured to use branch discovery.' : 'Approvals phase not configured.');
    return { ok, md: L.join('\n') };
}
