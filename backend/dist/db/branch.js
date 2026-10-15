"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.childrenOf = childrenOf;
exports.readExportShape = readExportShape;
exports.overviewRowsFor = overviewRowsFor;
const norm = (s) => String(s || '').toLowerCase();
/**
 * Branch classification.
 *
 * Deliberately narrow. A branch is only claimed when its name says so; anything else is reported as
 * unrecognised rather than being forced into one of the two buckets. Forcing would produce a slide
 * titled "Project Overview - Construction" over design work, which is worse than omitting it.
 */
function classify(name) {
    const n = norm(name);
    // "Pre-Construction Set [LOD 300]" and "Construction Documents [LOD 350]" are DESIGN deliverables
    // that merely contain the word "construction". They are level-3 children of "Design", never the
    // branch itself, but the test is written to be safe under either nesting.
    if (/\bpre-?construction\b/.test(n))
        return null;
    if (/\bconstruction documents?\b/.test(n))
        return null;
    if (/^construction\b/.test(n) || /\bconstruction$/.test(n))
        return 'construction';
    if (/^design\b/.test(n) || /\bdesign$/.test(n))
        return 'design';
    return null;
}
/** Children of a WBS node, in the export's own row order. */
function childrenOf(sch, id) {
    return Object.values(sch.wbs)
        .filter((n) => n.parent === id)
        .sort((a, b) => a.seq - b.seq);
}
/** A project code embedded in a row name: "0131 | Nandi Hills Homestay", "0116 - Meena and …". */
function codeIn(name) {
    const m = String(name || '').match(/^\s*([0-9]{3,4}(?:-[0-9]{1,2})?)\s*[|\-–:]/);
    return m ? m[1] : '';
}
/**
 * Inspect a parsed Design & Build export.
 *
 * Returns the row the Project Overview should start from and every reportable branch beneath it.
 * Never throws: an export with nothing recognisable comes back with an empty `branches` list and a
 * note saying so, and the engine then omits the affected sections.
 */
function readExportShape(sch, code) {
    const notes = [];
    const nodes = Object.values(sch.wbs).sort((a, b) => a.seq - b.seq);
    if (!nodes.length) {
        return { rootId: '', rootName: '', rootDepth: 0, code: '', branches: [],
            notes: ['The export contains no rows.'] };
    }
    // ---- the root row ----
    // Preferred: the topmost row whose name carries a project code. That is what both reference
    // Project Overview slides use as their first row.
    const wanted = String(code || '').replace(/^0+/, '');
    let root = nodes.find((n) => {
        const c = codeIn(n.name);
        if (!c)
            return false;
        return !wanted || c.replace(/^0+/, '') === wanted;
    });
    if (!root)
        root = nodes.find((n) => !!codeIn(n.name));
    if (root) {
        notes.push(`Project Overview root: row ${root.seq} "${root.name}" (depth ${root.depth}) — chosen ` +
            `because it carries the project code.`);
    }
    else {
        // No coded row. Fall back to the shallowest row that has children, so the tree still has a top.
        root = nodes.find((n) => childrenOf(sch, n.id).length > 0) || nodes[0];
        notes.push(`No row carries a project code, so the Project Overview root falls back to row ` +
            `${root.seq} "${root.name}" (depth ${root.depth}).`);
    }
    // ---- branches beneath the root ----
    // Only the root's own children are considered. A deeper row that happens to be called "Design"
    // is a sub-deliverable, not a reporting branch.
    const kids = childrenOf(sch, root.id);
    const branches = [];
    const unrecognised = [];
    for (const k of kids) {
        const kind = classify(k.name);
        if (!kind) {
            unrecognised.push(k.name);
            continue;
        }
        if (branches.some((b) => b.kind === kind)) {
            notes.push(`A second "${kind}" branch ("${k.name}") was found; the first one is used and this ` +
                `row is left to the schedule sections rather than creating a duplicate slide.`);
            continue;
        }
        branches.push({ kind, id: k.id, name: k.name, depth: k.depth, children: childrenOf(sch, k.id) });
    }
    if (branches.length) {
        notes.push(`Branches found: ${branches.map((b) => `"${b.name}" (${b.kind}, ${b.children.length} ` +
            `child row(s))`).join('; ')}.`);
    }
    else {
        notes.push('No "Design" or "Construction" branch under the root row — the Project Overview and ' +
            'schedule sections that depend on a branch are omitted for this export.');
    }
    if (unrecognised.length) {
        // These rows still appear on the Project Overview — it covers the root's whole subtree. They
        // simply do not NAME a section, which is all the branch classification decides.
        notes.push(`Root-level rows that do not name a reporting branch (they are still shown on the ` +
            `Project Overview, they just do not create a section of their own): ` +
            `${unrecognised.map((u) => `"${u}"`).join(', ')}.`);
    }
    return {
        rootId: root.id, rootName: root.name, rootDepth: root.depth,
        code: codeIn(root.name), branches, notes,
    };
}
/**
 * The "Duration [Elapsed day]" figure for a row.
 *
 * READ FROM THE EXPORT'S OWN DURATION COLUMN whenever the export has one. The slide's column is
 * headed "Duration [Elapsed day]", which is the name of a column in the export, so it has to BE that
 * column — not a second opinion about it.
 *
 * It used to be computed as (finish − start + 1), and that quietly disagreed with the source. On
 * Nandi Hills Homestay the two differed on 11 of 13 rows: Design read 588.31ed in the export and
 * printed 589ed; Approvals 470.79 printed 472; Master Archives 7 printed 8. The arithmetic is not
 * wrong — it counts calendar days inclusively — but ASTA's own elapsed duration is not that number,
 * and anyone checking the deck against the export sees a mismatch they cannot explain.
 *
 * Rounded ONCE, here, for display: 941.67 → 942. The unrounded value is kept on the node so this is
 * the only place that decides.
 *
 * FALLS BACK to counting the dates when the export has no readable duration for the row, so a row
 * that the source leaves blank still shows a sensible span rather than nothing.
 */
function elapsedDays(a, b, source) {
    if (source !== null && source !== undefined && Number.isFinite(source))
        return Math.round(source);
    if (!a || !b)
        return null;
    const d = (Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000;
    return isNaN(d) ? null : Math.max(0, Math.round(d)) + 1;
}
/**
 * Build the Project Overview rows for an export.
 *
 * WHAT THE SLIDE COVERS
 * ---------------------
 * The root's WHOLE subtree — not one branch. This was measured, not assumed: Meena's reference
 * Project Overview ends with "Master Archives [LOD 500]", and that row is a SIBLING of "Design"
 * (both are outline level 2 under the project row at level 1), which is why the reference paints it
 * the same colour as "Design" rather than the colour used for Design's children. A branch-scoped
 * slide would have dropped it.
 *
 * DEPTH IS CHOSEN, NOT FIXED
 * --------------------------
 * The reference slides show different depths because their trees are different sizes: NH Homestay's
 * Construction tree is small enough to list every activity (14 rows), Meena's Design tree is not.
 * So rows are added a whole level at a time while the total still fits `maxRows`, and the depth
 * actually used is reported. Nothing is ever truncated mid-level and no font is shrunk to force a
 * level in.
 *
 * Zero-duration milestone rows are excluded, as the reference does — NHHS's "Project Commencement"
 * is a 0-day milestone and does not appear on the reference slide; including it would add a row with
 * no bar and no dates to compare.
 *
 * A KNOWN, DELIBERATE DIFFERENCE FROM MEENA'S REFERENCE SLIDE
 * ----------------------------------------------------------
 * That slide is an ASTA screenshot in which "Master Archives [LOD 500]" happens to be COLLAPSED, so
 * it shows 8 rows. Its five children exist in the export and fit on the page, so they are rendered.
 * An ASTA window's expand/collapse state is not in the data and cannot be reconstructed from it —
 * guessing which branches to collapse would be inventing structure. The row count and depth are
 * stated in design_build_validation.md so the difference is visible rather than silent.
 */
function overviewRowsFor(sch, shape, maxRows, 
/**
 * OPTIONAL: restrict the tree to ONE branch of the project.
 *
 * The project's whole Design and Construction plan arrives in a SINGLE Excel export view, so when
 * that file carries both branches the two Project Overview slides have to be cut out of the one
 * tree — "Project Overview" from the Design branch, "Project Overview - Construction" from the
 * Construction branch. Without this they would both show the same mixed tree.
 *
 * Omitted, the whole root subtree is used. That is deliberate for an export with only ONE reportable
 * branch: Meena's Design export also carries "Master Archives [LOD 500]" as a SIBLING of "Design",
 * and its reference slide shows that row. Scoping to the branch would silently drop it.
 */
branch) {
    const root = sch.wbs[shape.rootId];
    const isMilestone = (n) => n.milestone === true || n.duration === 0;
    const mk = (n, level) => {
        const d = elapsedDays(n.start ?? null, n.finish ?? null, n.durationElapsed);
        return {
            node: n, level, name: n.name,
            pct: typeof n.pct === 'number' ? n.pct : null,
            baselineStart: n.baselineStart ?? null,
            baselineFinish: n.baselineFinish ?? null,
            start: n.start ?? null, finish: n.finish ?? null,
            elapsedDays: d,
            elapsedMonths: d == null ? null : Math.round((d / 30.4375) * 100) / 100,
        };
    };
    // Collect the subtree by level relative to the root. When a branch is given, the branch row itself
    // takes relative level 1 and its descendants follow, so the rendered indent and colour tier are the
    // same whether the export held one branch or two.
    const byLevel = [];
    const walk = (id, rel) => {
        for (const k of childrenOf(sch, id)) {
            (byLevel[rel] = byLevel[rel] || []).push(k);
            walk(k.id, rel + 1);
        }
    };
    if (branch) {
        const bn = sch.wbs[branch.id];
        if (bn) {
            byLevel[0] = [bn];
            walk(bn.id, 1);
        }
    }
    else if (root) {
        walk(root.id, 0);
    }
    const rootRow = root ? [mk(root, 0)] : [];
    let kept = [];
    let depthUsed = 0;
    const dropped = [];
    for (let rel = 0; rel < byLevel.length; rel++) {
        const level = byLevel[rel] || [];
        const add = level.filter((n) => {
            if (!isMilestone(n))
                return true;
            dropped.push(n.name);
            return false;
        });
        if (!add.length)
            continue;
        if (rootRow.length + kept.length + add.length > maxRows) {
            const rows = rootRow.concat(kept.slice().sort((a, b) => a.node.seq - b.node.seq));
            return {
                rows, depthUsed,
                note: `Rows shown to depth ${depthUsed} (${rows.length} row(s)). The next level would add ` +
                    `${add.length} more, taking the table past the ${maxRows}-row page limit, so it is not ` +
                    `expanded. No row is truncated and no font is shrunk to force it.` +
                    (dropped.length ? ` Zero-duration milestone row(s) omitted, as the reference does: ` +
                        `${dropped.map((d) => `"${d}"`).join(', ')}.` : ''),
            };
        }
        kept = kept.concat(add.map((n) => mk(n, rel + 1)));
        depthUsed = rel + 1;
    }
    // Schedule order across the whole subtree, so a parent is always immediately followed by its own
    // children rather than the rows arriving level by level.
    const rows = rootRow.concat(kept.slice().sort((a, b) => a.node.seq - b.node.seq));
    return {
        rows, depthUsed,
        note: `Rows shown to depth ${depthUsed} (${rows.length} row(s)) — the whole tree fits.` +
            (dropped.length ? ` Zero-duration milestone row(s) omitted, as the reference does: ` +
                `${dropped.map((d) => `"${d}"`).join(', ')}.` : ''),
    };
}
