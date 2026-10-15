"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.splitBlockName = splitBlockName;
exports.tidyBlock = tidyBlock;
exports.buildBlocks = buildBlocks;
exports.commonProjectName = commonProjectName;
exports.blockSplitSections = blockSplitSections;
exports.pageLabel = pageLabel;
exports.blockTitle = blockTitle;
/**
 * Split "Straight From The Heart [North Block]" into its project and its block.
 *
 * The qualifier is the LAST bracketed group, and it only counts as a block when it reads like one —
 * it has to contain a block-ish word. That test is deliberately narrow: exports use brackets for all
 * sorts of qualifiers ("Nandi Hills [MH]", "ItQE Ph 2c & 3a [T9, 10,11,12 & 31 Villas]",
 * "Site Establishment and Site Mobilization [FK 0]") and none of those is a reporting block. Getting
 * this wrong in the permissive direction would split an ordinary project in half, so when in doubt
 * this returns no block and the project behaves exactly as it always has.
 */
const BLOCK_WORDS = /\b(block|phase\s*block|sector|precinct|parcel|zone)\b/i;
function splitBlockName(name) {
    const raw = String(name || '').trim();
    const m = raw.match(/^(.*?)\s*\[([^\]]+)\]\s*$/);
    if (!m)
        return { project: raw, block: '' };
    const inner = m[2].trim();
    if (!BLOCK_WORDS.test(inner))
        return { project: raw, block: '' };
    return { project: m[1].trim(), block: inner };
}
/** Title Case a block label so "north block" and "NORTH BLOCK" read the same in a heading. */
function tidyBlock(label) {
    return label.replace(/\s+/g, ' ').trim()
        .replace(/\w\S*/g, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase());
}
const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, '');
/**
 * Build the block list from the parsed exports, in the order the PREVIOUS DECK presents them.
 *
 * The reference deck is the structure's source of truth, so if it shows South before North that is
 * the order the new deck uses, whichever order the exports happened to be supplied in. With no
 * reference — or a reference that names no block — the supplied order stands.
 */
function buildBlocks(schedules, refTitles = [], 
/**
 * Only titles beginning with one of these section names are used to decide the ORDER. Without it
 * the cover decides, and the cover lists the blocks in whatever order the project describes itself
 * ("North Block : Commercial Dev | 2 acres" above "South Block : Mixed Use Dev | 20 acres") while
 * the SECTIONS run South first. It is the sections that have to match.
 */
orderWithin = []) {
    const blocks = schedules.map((sch) => {
        const { project, block } = splitBlockName(sch.projectName);
        return { label: tidyBlock(block), project, rawName: sch.projectName, sch };
    });
    if (blocks.length < 2)
        return blocks;
    // First mention of each block label among the reference deck's SECTION slide titles.
    const scoped = orderWithin.length
        ? refTitles.filter((t) => orderWithin.some((sec) => norm(t).startsWith(norm(sec))))
        : refTitles;
    const hay = (scoped.length ? scoped : refTitles).map((t) => norm(t)).join(' | ');
    const firstAt = (label) => {
        const i = hay.indexOf(norm(label));
        return i < 0 ? Number.MAX_SAFE_INTEGER : i;
    };
    const ordered = blocks.slice();
    if (hay) {
        ordered.sort((a, b) => {
            const d = firstAt(a.label) - firstAt(b.label);
            return d !== 0 ? d : blocks.indexOf(a) - blocks.indexOf(b); // stable for anything unmentioned
        });
    }
    return ordered;
}
/**
 * The project name to use for the deck as a whole when several blocks are combined.
 *
 * It is the block-free name the exports agree on. This matters beyond the cover: the Summary
 * Dashboard holds ONE row for the project ("Straight from the Heart", code 0106), not one per block,
 * and Key Focus Area matches it STRICTLY. An export named "Straight From The Heart [North Block]"
 * therefore matches nothing and the section comes out empty — which is exactly what happened. When
 * the exports disagree on the block-free part, the first one's is used and the disagreement is
 * reported rather than guessed at.
 */
function commonProjectName(blocks) {
    const seen = [...new Set(blocks.map((b) => b.project).filter(Boolean))];
    return { name: seen[0] || (blocks[0]?.rawName ?? ''), agreed: seen.length <= 1, seen };
}
/**
 * WHICH SECTIONS THE PREVIOUS DECK REPORTS PER BLOCK.
 *
 * Read from the reference deck's own slide titles: a section is block-split when a title that starts
 * with that section's name also carries one of this project's block labels. Nothing is assumed — a
 * project whose previous deck keeps Approvals as one section gets one Approvals section.
 *
 * `sectionTitles` are the deck's own section names, e.g. "Project Overview", "Approvals Status",
 * "2 Week Look Ahead Plan".
 */
function blockSplitSections(refTitles, blocks, sectionTitles) {
    const out = new Set();
    if (blocks.length < 2 || !refTitles.length)
        return out;
    const labels = blocks.map((b) => norm(b.label)).filter(Boolean);
    for (const sec of sectionTitles) {
        const s = norm(sec);
        if (!s)
            continue;
        const hits = refTitles.filter((t) => norm(t).startsWith(s));
        // at least two of this project's blocks must be named across that section's own slides
        const named = new Set(labels.filter((l) => hits.some((t) => norm(t).includes(l))));
        if (named.size >= 2)
            out.add(sec);
    }
    return out;
}
/**
 * The page label for one slide of a section: "4" when the section is a single page, otherwise
 * "4A", "4B", … Letters run CONTINUOUSLY across the blocks, so South's three pages are 4A 4B 4C and
 * North's three carry on 4D 4E 4F, exactly as the previous deck numbers them.
 */
function pageLabel(sectionNo, indexAcrossSection, totalInSection) {
    if (totalInSection <= 1)
        return String(sectionNo);
    return `${sectionNo}${String.fromCharCode(65 + indexAcrossSection)}`;
}
/** "Project Overview" + "South Block" → "Project Overview [South Block]" (no suffix when unblocked). */
function blockTitle(base, label) {
    return label ? `${base} [${label}]` : base;
}
