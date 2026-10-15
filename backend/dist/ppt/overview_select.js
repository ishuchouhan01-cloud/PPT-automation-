"use strict";
// overview_select.ts — rule-based selection of the activities shown on Project Overview.
//
// WHY THIS EXISTS
// Different projects nest their WBS differently, so keying off "outline level 3" picked the wrong
// rows on some exports. This module classifies every WBS node by WHAT IT IS, not how deep it sits,
// and keeps the ones that represent major project progress.
//
// Priority order (as specified):
//   1. Major WBS Sections     — the top structural breakdown of the programme
//   2. Major Deliverables     — design/document packages that gate progress
//   3. Execution Packages     — construction / infrastructure / utilities work
//   4. Approval Packages      — statutory NOCs, sanctions, certifications
//   5. Site Mobilisation      — site establishment, mobilisation, security, P&M
//   6. Sales Activities       — launch, campaign, marketing collateral, experience centre
//
// Everything else (project/phase containers, milestone leaves, fine-grained detail) is rejected with
// a recorded reason, published to project_overview_selection.md.
Object.defineProperty(exports, "__esModule", { value: true });
exports.PRIORITY = void 0;
exports.isContainer = isContainer;
exports.sectionDepth = sectionDepth;
exports.selectOverview = selectOverview;
exports.selectionReport = selectionReport;
exports.PRIORITY = [
    'Major WBS Section', 'Major Deliverable', 'Execution Package',
    'Approval Package', 'Site Mobilisation', 'Sales Activity',
];
// Keyword vocabularies — deliberately broad so the rules generalise across projects.
const KW = [
    { cat: 'Approval Package', words: /\b(noc|approval|approvals|sanction|sanctioned|rera|cfe|cto|clearance|certificat|licen[cs]e|permit|khata|ekhata|clu|conversion|consent|moef|seac|environment|fire|height|airport|layout plan|development plan|building plan|occupancy|water supply|power|drainage)\b/i },
    { cat: 'Sales Activity', words: /\b(sales|launch|campaign|marketing|collateral|brochure|experience centre|experience center|exp home|edm|flyer|price book|inventory|walkthrough|pre[- ]?launch|booking|crm)\b/i },
    { cat: 'Site Mobilisation', words: /\b(mobilis|mobiliz|site establishment|site security|p&m|plant (and|&) machinery|labour camp|site office|boundary wall|site clearance|temporary)\b/i },
    { cat: 'Execution Package', words: /\b(construction|infrastructure|utilities|utility|execution|works|civil|structure|structural|mep|finishes|package|gfc|handover|delivery|excavation|foundation|shell)\b/i },
    { cat: 'Major Deliverable', words: /\b(design|drawing|drawings|concept|schematic|master ?plan|brief|estimate|boq|lod|pre construction|preconstruction|documents|document|tender|contract|business plan|land acquisition|acquisition|jda|mou|tsr|agreement|feasibility|survey|system setup|milestone)\b/i },
];
const norm = (s) => String(s || '').trim().toLowerCase();
/** Classify a node by name; undefined when nothing matches. */
function classify(name) {
    for (const cat of exports.PRIORITY) {
        const rule = KW.find((k) => k.cat === cat);
        if (rule && rule.words.test(name))
            return cat;
    }
    return undefined;
}
/**
 * A "container" row carries no progress meaning of its own: it echoes the project name, or is a bare
 * phase label such as "Ph 1 [21.76a]" / "Phase 2". Recognising these is what makes the section depth
 * detection generic — container levels repeat the project name, section levels do not.
 */
function isContainer(name, projectName) {
    const n = norm(name);
    const projTokens = norm(projectName).split(/[^a-z0-9]+/).filter((x) => x.length >= 4);
    if (projTokens.length && projTokens.every((tk) => n.includes(tk)))
        return true; // echoes project
    if (/^(ph|phase)\s*[0-9ivx]+/i.test(name.trim()))
        return true; // bare phase
    if (/^[a-z\s]*\bph(ase)?\s*[0-9]/i.test(name.trim()) && /\[.*\]/.test(name))
        return true;
    return false;
}
/**
 * Depth at which the MAJOR SECTIONS live. Container levels (project / phase chains) are skipped, so
 * this works whether a project nests two levels deep or five.
 */
function sectionDepth(nodes, projectName = '') {
    const byDepth = new Map();
    for (const w of nodes) {
        if (w.depth < 1)
            continue;
        (byDepth.get(w.depth) || byDepth.set(w.depth, []).get(w.depth)).push(w);
    }
    const depths = [...byDepth.keys()].sort((a, b) => a - b);
    // shallowest depth carrying at least 3 non-container rows = the section level
    for (const d of depths) {
        const real = (byDepth.get(d) || []).filter((w) => !isContainer(w.name, projectName));
        if (real.length >= 3)
            return d;
    }
    // fall back to the first branching level
    for (const d of depths)
        if ((byDepth.get(d) || []).length > 1)
            return d;
    return depths[0] ?? 1;
}
/**
 * Decide, for every WBS node, whether it belongs on the Project Overview.
 * Returns decisions in schedule order so the report mirrors the deck.
 */
function selectOverview(sch) {
    const all = Object.values(sch.wbs).sort((a, b) => a.seq - b.seq);
    const kids = new Map();
    for (const w of all)
        if (w.parent)
            kids.set(w.parent, (kids.get(w.parent) || 0) + 1);
    const hasChildren = (w) => (kids.get(w.id) || 0) > 0;
    const S = sectionDepth(all, sch.projectName);
    const out = [];
    const seenInParent = new Set();
    for (const w of all) {
        const name = String(w.name || '').trim();
        // ---- hard exclusions -------------------------------------------------
        if (!name) {
            out.push({ node: w, selected: false, reason: 'Unnamed WBS node' });
            continue;
        }
        if (w.depth < S) {
            out.push({ node: w, selected: false, reason: `Project/phase container (depth ${w.depth} < section depth ${S})` });
            continue;
        }
        // activities inside a "Milestone" group are represented by the Milestone row itself
        let cur = w.parent, underMilestone = false;
        while (cur && sch.wbs[cur]) {
            if (/milestone/i.test(sch.wbs[cur].name)) {
                underMilestone = true;
                break;
            }
            cur = sch.wbs[cur].parent;
        }
        if (underMilestone) {
            out.push({ node: w, selected: false, reason: 'Detail inside a Milestone group (Milestone summary shown instead)' });
            continue;
        }
        // de-duplicate repeated names within the same parent (ASTA repeats package names per block)
        const dedupeKey = `${w.parent}|${norm(name)}`;
        if (seenInParent.has(dedupeKey)) {
            out.push({ node: w, selected: false, reason: 'Duplicate package name within the same section' });
            continue;
        }
        // ---- rule-based inclusion -------------------------------------------
        let category;
        let reason = '';
        if (w.depth === S) {
            category = 'Major WBS Section';
            reason = `Rule 1 — major WBS section (top structural level, depth ${S})`;
        }
        else if (w.depth === S + 1) {
            const kw = classify(name);
            if (hasChildren(w)) {
                category = kw && kw !== 'Major WBS Section' ? kw : 'Major Deliverable';
                reason = `Rule 2 — summary package directly under a major section${kw ? ` (matched ${kw})` : ''}`;
            }
            else if (kw) {
                category = kw;
                reason = `Rule ${exports.PRIORITY.indexOf(kw) + 1} — recognised ${kw.toLowerCase()} by name`;
            }
            else {
                out.push({ node: w, selected: false, reason: 'Detail activity with no package/deliverable signature' });
                continue;
            }
        }
        else {
            // Deeper than S+1 is detail, not overview. (Kept explicit so the report explains every row.)
            out.push({ node: w, selected: false, reason: `Below overview level (depth ${w.depth} > ${S + 1})` });
            continue;
        }
        // A container row that slipped through (phase label at the section level) carries no meaning.
        if (isContainer(name, sch.projectName)) {
            out.push({ node: w, selected: false, reason: 'Project/phase container row (no progress meaning)' });
            continue;
        }
        seenInParent.add(dedupeKey);
        out.push({ node: w, selected: true, category, reason });
    }
    return out;
}
/** Markdown report of what the rules chose and why. */
function selectionReport(sch, decisions) {
    const all = Object.values(sch.wbs);
    const S = sectionDepth(all, sch.projectName);
    const sel = decisions.filter((d) => d.selected);
    const rej = decisions.filter((d) => !d.selected);
    const L = [];
    L.push('# project_overview_selection.md', '');
    L.push(`Project: **${sch.projectName}**  ·  WBS nodes: ${all.length}  ·  detected section depth: **${S}**`, '');
    L.push('Selection is rule-based, not depth-based: each WBS node is classified by what it *is*, so', 'the overview holds up across projects with different WBS structures.', '');
    L.push('## Priority rules', '');
    exports.PRIORITY.forEach((p, i) => L.push(`${i + 1}. ${p}`));
    L.push('', `## Selected activities (${sel.length})`, '', '| # | Activity | Depth | Category | Reason |', '|---|---|---|---|---|');
    sel.forEach((d, i) => L.push(`| ${i + 1} | ${d.node.name} | ${d.node.depth} | ${d.category} | ${d.reason} |`));
    const byReason = new Map();
    for (const d of rej) {
        const k = d.reason.replace(/\d+/g, 'N');
        byReason.set(k, (byReason.get(k) || 0) + 1);
    }
    L.push('', `## Rejected activities (${rej.length})`, '', '### Summary by reason', '', '| Reason | Count |', '|---|---|');
    [...byReason.entries()].sort((a, b) => b[1] - a[1]).forEach(([r, n]) => L.push(`| ${r} | ${n} |`));
    L.push('', '### Detail (first 200)', '', '| Activity | Depth | Reason |', '|---|---|---|');
    rej.slice(0, 200).forEach((d) => L.push(`| ${d.node.name} | ${d.node.depth} | ${d.reason} |`));
    if (rej.length > 200)
        L.push('', `_…and ${rej.length - 200} more._`);
    return L.join('\n');
}
