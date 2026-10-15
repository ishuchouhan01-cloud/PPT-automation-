"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.VIEW_LABEL = exports.TRIAL_VIEWS = void 0;
exports.wbsEntries = wbsEntries;
exports.resolveOverviewView = resolveOverviewView;
exports.sampleCustomIds = sampleCustomIds;
exports.trialReport = trialReport;
const branch_1 = require("./branch");
exports.TRIAL_VIEWS = ['major', 'level1', 'level2', 'level3', 'level4', 'custom'];
exports.VIEW_LABEL = {
    default: 'Default (shipped selector)',
    major: 'Major Overview',
    level1: 'Level 1',
    level2: 'Level 2',
    level3: 'Level 3',
    level4: 'Level 4',
    custom: 'Custom',
};
/**
 * Flatten the project root's subtree into depth-annotated entries, in export row order.
 *
 * Zero-duration milestones are marked but kept in the structure, so a milestone that is somebody's
 * parent still links its children correctly; the views drop them when selecting.
 */
function wbsEntries(sch, shape, 
/**
 * OPTIONAL: restrict the tree to one branch of the project.
 *
 * The whole plan arrives in a SINGLE Excel export view, so when that file holds both a Design and a
 * Construction branch the two Project Overview slides are cut out of the one tree. A view applied to
 * "Project Overview - Construction" must therefore see only the Construction branch — otherwise
 * "Level 2" would count the Design rows too and the two slides would select from the same set.
 *
 * The project row is always kept as the tree's top, so the slide still opens with the project.
 */
branch) {
    const out = [];
    const root = sch.wbs[shape.rootId];
    if (!root)
        return out;
    const push = (n, depth, parent) => {
        const kids = (0, branch_1.childrenOf)(sch, n.id);
        out.push({
            id: n.id, name: n.name, depth, parent,
            hasChildren: kids.length > 0, seq: n.seq,
            milestone: n.milestone === true || n.duration === 0,
        });
        for (const k of kids)
            push(k, depth + 1, n.id);
    };
    if (branch && sch.wbs[branch.id]) {
        // The project row, then the branch as its only child, so relative depths match the unscoped case.
        const kidsOfRoot = (0, branch_1.childrenOf)(sch, root.id);
        out.push({
            id: root.id, name: root.name, depth: 0, parent: null,
            hasChildren: kidsOfRoot.length > 0, seq: root.seq,
            milestone: root.milestone === true || root.duration === 0,
        });
        push(sch.wbs[branch.id], 1, root.id);
    }
    else {
        push(root, 0, null);
    }
    return out.sort((a, b) => a.seq - b.seq);
}
/** Build a DbOverviewRow from a WBS node, carrying the export's own values through unchanged. */
function rowOf(sch, e) {
    const n = sch.wbs[e.id];
    const days = (() => {
        if (!n.start || !n.finish)
            return null;
        const d = (Date.parse(n.finish + 'T00:00:00Z') - Date.parse(n.start + 'T00:00:00Z')) / 86400000;
        return isNaN(d) ? null : Math.max(0, Math.round(d)) + 1;
    })();
    return {
        node: n,
        // The rendered indent/colour tier is the row's depth relative to the project row.
        level: e.depth,
        name: n.name,
        pct: typeof n.pct === 'number' ? n.pct : null,
        baselineStart: n.baselineStart ?? null,
        baselineFinish: n.baselineFinish ?? null,
        start: n.start ?? null,
        finish: n.finish ?? null,
        elapsedDays: days,
        elapsedMonths: days == null ? null : Math.round((days / 30.4375) * 100) / 100,
    };
}
/**
 * Close a selected set upwards: for every selected row, add the chain of ancestors up to the project
 * row, and nothing else.
 *
 * This is what makes a Custom selection readable without leaking. Selecting "Clubhouse" three levels
 * down brings in "Schematic Concept" and "Design" so the row can be placed in the plan — but it does
 * NOT bring in Schematic Concept's other children, or Design's other branches. Only the direct line.
 */
function withAncestors(selected, byId) {
    const ids = new Set(selected);
    let added = 0;
    for (const id of selected) {
        let cur = byId.get(id)?.parent || null;
        while (cur && !ids.has(cur)) {
            ids.add(cur);
            added++;
            cur = byId.get(cur)?.parent || null;
        }
    }
    return { ids, added };
}
function summarise(sch, entries, ids, added, note) {
    const chosen = entries.filter((e) => ids.has(e.id));
    const rows = chosen.map((e) => rowOf(sch, e));
    return {
        rows,
        parents: chosen.filter((e) => e.hasChildren).length,
        leaves: chosen.filter((e) => !e.hasChildren).length,
        ancestorsAdded: added,
        maxDepth: chosen.reduce((m, e) => Math.max(m, e.depth), 0),
        note,
    };
}
// ---------------------------------------------------------------------------------------------
// The views
// ---------------------------------------------------------------------------------------------
/**
 * Resolve a view to rows.
 *
 * `LEVEL n` is the Excel outline behaviour the brief asked for — "the same way Excel can
 * expand/collapse outline levels". Excel's outline buttons show every row down to level n, so
 * `level3` is every row at relative depth 0–3. It is NOT "the first N rows" and it is NOT "only the
 * rows exactly at depth n": the latter would hide a depth-2 phase that happens to have no depth-3
 * children, such as Nandi Hills' "Site Development", and the plan would then be missing a phase.
 *
 * `MAJOR OVERVIEW` is the management cut: the project row, its phases, and their sub-phases —
 * relative depth 0–2 — with individual detail tasks left out. On a deep tree that is a large
 * reduction; on the two reference projects, whose trees are broad rather than deep at the top, it
 * coincides with Level 2, and the trial report says so rather than implying a difference that is not
 * there.
 *
 * `CUSTOM` takes an explicit set of row ids and adds only the ancestors needed to place them.
 */
function resolveOverviewView(sch, shape, view, customIds = [], 
/** Restrict to one branch — see wbsEntries(). */
branch) {
    const entries = wbsEntries(sch, shape, branch);
    const byId = new Map(entries.map((e) => [e.id, e]));
    const live = entries.filter((e) => !e.milestone);
    const depthCut = (max, label) => {
        const ids = new Set(live.filter((e) => e.depth <= max).map((e) => e.id));
        // The project row is always present, so a view can never open on an orphaned child.
        const rootId = entries[0]?.id;
        if (rootId && !ids.has(rootId) && !byId.get(rootId).milestone)
            ids.add(rootId);
        const dropped = entries.filter((e) => e.milestone && e.depth <= max).map((e) => e.name);
        return summarise(sch, entries, ids, 0, `${label}: every row at relative depth 0–${max} of this project's own outline, which is what ` +
            `Excel's outline level ${max} button shows. Depth is measured from the project row ` +
            `("${entries[0]?.name || '?'}"), not from the absolute outline number, because the two exports ` +
            `nest the project row at different levels.` +
            (dropped.length ? ` Zero-duration milestone row(s) excluded: ${dropped.map((d) => `"${d}"`).join(', ')}.` : ''));
    };
    switch (view) {
        case 'major': {
            const r = depthCut(2, 'Major Overview');
            const tail = r.note.replace(/^Major Overview: /, '');
            return { ...r, note: `Major Overview — the management cut: the project row, its phases and their sub-phases, with ` +
                    `individual detail tasks excluded. That is relative depth 0–2, so on a tree that is broad ` +
                    `rather than deep at the top it selects the same rows as the Level 2 view; the two diverge ` +
                    `only where a project carries detail below depth 2. Mechanically: ${tail}` };
        }
        case 'level1': return depthCut(1, 'Level 1');
        case 'level2': return depthCut(2, 'Level 2');
        case 'level3': return depthCut(3, 'Level 3');
        case 'level4': return depthCut(4, 'Level 4');
        case 'custom': {
            const valid = customIds.filter((id) => byId.has(id) && !byId.get(id).milestone);
            const unknown = customIds.filter((id) => !byId.has(id));
            if (!valid.length) {
                return { rows: [], parents: 0, leaves: 0, ancestorsAdded: 0, maxDepth: 0,
                    note: 'Custom: nothing was selected, so no Project Overview was generated for this view.' +
                        (unknown.length ? ` ${unknown.length} id(s) did not match any row in this export.` : '') };
            }
            const { ids, added } = withAncestors(new Set(valid), byId);
            const r = summarise(sch, entries, ids, added, `Custom: ${valid.length} row(s) selected explicitly, plus ${added} ancestor row(s) added so ` +
                `each selection can be placed in the plan. Only the direct line above a selected row is ` +
                `added — sibling branches are never pulled in.` +
                (unknown.length ? ` ${unknown.length} requested id(s) are not in this export and were ignored.` : ''));
            return r;
        }
        default:
            return { rows: [], parents: 0, leaves: 0, ancestorsAdded: 0, maxDepth: 0,
                note: 'Default view — the shipped selector is used and this module is not involved.' };
    }
}
/**
 * A sensible Custom selection to demonstrate the view without the operator having to tick boxes.
 *
 * Chosen from the export itself: every phase (relative depth 1), every sub-phase (depth 2), and the
 * children of the FIRST sub-phase only. That produces exactly the shape the brief's example asks
 * for — some branches expanded, their siblings collapsed — and shows that selecting a deep row does
 * not expose unrelated siblings.
 */
function sampleCustomIds(sch, shape, branch) {
    const entries = wbsEntries(sch, shape, branch).filter((e) => !e.milestone);
    const picked = entries.filter((e) => e.depth >= 1 && e.depth <= 2).map((e) => e.id);
    const firstSub = entries.find((e) => e.depth === 2 && e.hasChildren);
    if (firstSub) {
        for (const e of entries)
            if (e.parent === firstSub.id)
                picked.push(e.id);
    }
    return picked;
}
function trialReport(project, code, source, rowsInExport, results) {
    const L = [];
    L.push('# Project Overview — WBS view trial', '');
    L.push(`**Project** ${project}${code ? ` (${code})` : ''}  `);
    L.push(`**Source** ${source} — ${rowsInExport} WBS rows in the export  `);
    L.push('**Status** TRIAL. The shipped Project Overview selector is unchanged and remains the ' +
        'default; these views are produced only when one is requested.', '');
    L.push('## What is being compared', '');
    L.push('Only WHICH activities appear. The table, the Gantt, the column set, the colours, the fonts, ' +
        'the row formatting, the calendar and the three bars are rendered by the same unchanged code in ' +
        'every view. No value is recalculated — every date, percentage and duration is the figure already ' +
        'in the export.', '');
    L.push('## Comparison', '');
    L.push('| View | Rows | Parent rows | Leaf activities | Pages | Deepest level | Gantt rows |');
    L.push('|---|---:|---:|---:|---:|---:|---:|');
    for (const r of results) {
        L.push(`| ${exports.VIEW_LABEL[r.view]} | ${r.rows} | ${r.parents} | ${r.leaves} | ${r.pages} | ` +
            `${r.maxDepth} | ${r.ganttRows}${r.ganttRows === r.rows ? '' : ' ⚠'} |`);
    }
    L.push('');
    const mismatch = results.filter((r) => r.ganttRows !== r.rows);
    L.push('## Table and Gantt agree', '');
    if (mismatch.length) {
        L.push(`**FAIL** — ${mismatch.map((m) => exports.VIEW_LABEL[m.view]).join(', ')} drew a different number of ` +
            'Gantt rows than table rows.', '');
    }
    else {
        L.push('**PASS** — every view drew exactly one Gantt row per table row, and no Gantt row for a ' +
            'hidden activity. This holds by construction, not by coincidence: the renderer draws each row\'s ' +
            'cells and its bars in the same loop, from the same measured geometry, so a row cannot exist in ' +
            'one and not the other.', '');
    }
    L.push('## Pagination', '');
    L.push('Deeper views run past one page. Each page keeps the same page size, the same table and Gantt ' +
        'widths, the same header and the same calendar — the date window is computed once across the whole ' +
        'view and pinned, so a given month sits at the same x on every page and the pages can be read side ' +
        'by side. Rows are never compressed to force a set onto fewer pages.', '');
    // ---- observations, drawn from the numbers rather than asserted ----
    const by = new Map(results.map((r) => [r.view, r]));
    const onePage = results.filter((r) => r.pages === 1 && r.rows > 0);
    const multi = results.filter((r) => r.pages > 1);
    L.push('## What the numbers show', '');
    if (onePage.length) {
        const widest = onePage.reduce((a, b) => (b.rows > a.rows ? b : a));
        L.push(`- **Fits one slide:** ${onePage.map((r) => `${exports.VIEW_LABEL[r.view]} (${r.rows})`).join(', ')}. ` +
            `The most detailed single-slide view here is **${exports.VIEW_LABEL[widest.view]}** at ${widest.rows} rows.`);
    }
    if (multi.length) {
        L.push(`- **Runs to several slides:** ${multi.map((r) => `${exports.VIEW_LABEL[r.view]} (${r.rows} rows / ` +
            `${r.pages} slides)`).join(', ')}. Readable, but no longer a single management page.`);
    }
    const maj = by.get('major'), l2 = by.get('level2');
    if (maj && l2 && maj.rows === l2.rows) {
        L.push(`- **Major Overview and Level 2 selected the same ${maj.rows} rows on this project.** That is ` +
            `not a bug: this tree is broad rather than deep at the top, so "phases and sub-phases" and ` +
            `"outline level 2" happen to describe the same set. They diverge on a project that carries ` +
            `detail tasks below depth 2.`);
    }
    const lvl = ['level1', 'level2', 'level3', 'level4'].map((v) => by.get(v)).filter(Boolean);
    const plateau = lvl.filter((r, i) => i > 0 && r.rows === lvl[i - 1].rows);
    if (plateau.length) {
        L.push(`- **${plateau.map((p) => exports.VIEW_LABEL[p.view]).join(' and ')} added nothing** over the level ` +
            `above, because this export is not that deep. The deepest row in it sits at relative level ` +
            `${Math.max(...results.map((r) => r.maxDepth))}.`);
    }
    const cus = by.get('custom');
    if (cus && cus.ancestorsAdded) {
        L.push(`- **Custom** kept ${cus.rows - cus.ancestorsAdded} explicitly chosen row(s) and added ` +
            `${cus.ancestorsAdded} parent row(s) so each could be placed in the plan — no sibling branch ` +
            `was pulled in.`);
    }
    L.push('');
    L.push('This is the trade-off to decide: a one-slide view a director can read at a glance, against a ' +
        'multi-slide view that shows the work packages a delivery meeting argues about. The numbers above ' +
        'are the evidence; the choice is a reporting decision, not a technical one.', '');
    L.push('## What each view did', '');
    for (const r of results) {
        L.push(`### ${exports.VIEW_LABEL[r.view]}`, '');
        L.push(`Rows **${r.rows}** · parent **${r.parents}** · leaf **${r.leaves}** · slides **${r.pages}** · ` +
            `deepest relative level **${r.maxDepth}**`, '');
        L.push(r.note, '');
    }
    return L.join('\n');
}
