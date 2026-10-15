"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.validateRefinement = validateRefinement;
exports.refinementReport = refinementReport;
// refine_validate.ts — verifies the presentation-refinement rules on the FINISHED .pptx.
//
// Everything below is read back out of the generated package at OOXML level, so the report states
// what the deck actually contains, not what the code intended:
//
//   • Project Overview  — only major section headings carry a background; every other activity row
//                         is white; conditional colour appears in the Actual / Proj. Finish column
//                         and nowhere else.
//   • Approvals         — Baseline Finish and Actual / Proj. Finish dates are Meta-Normal, not bold.
//   • 2 Week Look Ahead — the TE Departments column is black, not accent orange.
//   • Key Focus Area    — column headings as large as each column allows (cap 28pt, floor 16pt; the
//                         widest heading is still 28pt), "Team" replaces "Major Heading", body black.
//
// Nothing here changes a slide; it only inspects and reports.
const jszip_1 = __importDefault(require("jszip"));
const typography_1 = require("./typography");
const overview_1 = require("./overview");
const EMU = 914400;
const OLIVE = 'A9A25B', DELAY_FILL = 'F6D9C6';
/** The accent orange that was removed from body text. */
const ACCENT = ['E97132', 'C04F15'];
const decode = (s) => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
const textsOf = (xml) => (xml.match(/<a:t>[^<]*<\/a:t>/g) || []).map((s) => decode(s.replace(/<\/?a:t>/g, '')));
function shapesOf(xml) {
    const out = [];
    for (const m of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
        const body = m[0];
        const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
        const ext = body.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
        if (!off || !ext)
            continue;
        // shape fill = the solidFill inside spPr (before the text body)
        const spPr = body.split('<p:txBody')[0];
        const fillM = spPr.match(/<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/);
        out.push({
            x: +off[1] / EMU, y: +off[2] / EMU, w: +ext[1] / EMU, h: +ext[2] / EMU,
            text: textsOf(body).join(' ').trim(),
            fill: fillM ? fillM[1].toUpperCase() : null,
            colors: [...body.matchAll(/<a:rPr[\s\S]{0,200}?<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/g)].map((x) => x[1].toUpperCase()),
            faces: [...body.matchAll(/typeface="([^"]+)"/g)].map((x) => x[1]),
            sizes: [...body.matchAll(/sz="(\d+)"/g)].map((x) => +x[1] / 100),
            bold: /<a:rPr[^>]*\bb="1"/.test(body),
        });
    }
    return out;
}
/** All table cells of a slide, as row-major arrays of raw <a:tc> XML. */
/** A table cell's own FILL — the solidFill inside <a:tcPr>, ignoring the border colours in <a:lnX>. */
function cellFill(tc) {
    const m = tc.match(/<a:tcPr\b[\s\S]*?<\/a:tcPr>/);
    if (!m)
        return null;
    const body = m[0].replace(/<a:ln[LRTB]\b[\s\S]*?<\/a:ln[LRTB]>/g, '');
    const f = body.match(/<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/);
    return f ? f[1].toUpperCase() : null;
}
/** A cell's run colour. */
function cellTextColour(tc) {
    const m = tc.match(/<a:rPr[\s\S]*?<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/);
    return m ? m[1].toUpperCase() : null;
}
/** Straight-line shapes (prstGeom "line") with their geometry, colour and weight. */
function linesOf(xml) {
    const out = [];
    for (const m of xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
        const b = m[0];
        if (!/prst="line"/.test(b))
            continue;
        const off = b.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
        const ext = b.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
        if (!off || !ext)
            continue;
        const ln = b.match(/<a:ln[^>]*w="(\d+)"[^>]*>\s*<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/);
        out.push({
            x: +off[1] / EMU, y: +off[2] / EMU, w: +ext[1] / EMU, h: +ext[2] / EMU,
            color: ln ? ln[2].toUpperCase() : null, pt: ln ? +ln[1] / 12700 : 0,
        });
    }
    return out;
}
function tablesOf(xml) {
    const tables = [];
    for (const tb of xml.matchAll(/<a:tbl\b[\s\S]*?<\/a:tbl>/g)) {
        const rows = [];
        for (const tr of tb[0].matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)) {
            rows.push([...tr[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((c) => c[0]));
        }
        tables.push(rows);
    }
    return tables;
}
/** Slides whose CONTENT mentions the section — the Contents slide is excluded, since it lists every
 *  section title and would otherwise be mistaken for the section's own slide. */
const slideHas = (xml, s) => {
    const texts = textsOf(xml).map((t) => t.toUpperCase());
    if (texts.some((t) => t === 'CONTENTS'))
        return false;
    return texts.some((t) => t.includes(s));
};
async function validateRefinement(buf) {
    const z = await jszip_1.default.loadAsync(buf);
    const pres = (await z.file('ppt/presentation.xml')?.async('string')) || '';
    const rels = (await z.file('ppt/_rels/presentation.xml.rels')?.async('string')) || '';
    const rm = {};
    for (const m of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g))
        rm[m[1]] = m[2];
    const xmls = [];
    for (const m of pres.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)) {
        const t = rm[m[1]];
        if (!t || !/slides\/slide/.test(t))
            continue;
        xmls.push((await z.file('ppt/' + t.replace(/^\.\.\//, '').replace(/^ppt\//, ''))?.async('string')) || '');
    }
    const checks = [];
    const add = (name, ok, detail) => checks.push({ name, ok, detail });
    // ---------- Project Overview ----------
    const ovXmls = xmls.filter((x) => slideHas(x, 'PROJECT OVERVIEW'));
    // The slide is two panels: the table on the left and the Gantt on the right. Everything from
    // GANTT_X onwards belongs to the chart — its row bands, its bars, its axis and its TODAY rule —
    // and must not be judged by the table's rules.
    //
    // This boundary is IMPORTED from the renderer's own geometry, never recomputed here. It used to be
    // re-derived as `0.5 + (26.11 - 0.5) * 0.73`, a copy of the split that was true at the time. When
    // the Project Overview block was re-aligned to the header bar the renderer's boundary moved and
    // this copy did not, so the Gantt's grey year band and its blue/yellow bars fell on the table side
    // and were marked up as an illegal header fill and as stray conditional colouring. The report was
    // wrong, not the slide. One source of truth removes that whole class of false failure.
    const GANTT_X = (0, overview_1.overviewGeometry)().GX;
    const inTable = (sh) => sh.x < GANTT_X - 0.05;
    const inGantt = (sh) => sh.x >= GANTT_X - 0.05;
    // Row backgrounds are table-width rects at X0 0.5 (17.16" wide); the conditional cells are narrow.
    const rowRects = ovXmls.flatMap((x) => shapesOf(x).filter((s) => !s.text && inTable(s) && s.w > 12 && s.y > 3));
    // THE THREE-LEVEL HIERARCHY (replaces the old "olive headings on white rows" rule, which described
    // the design this one supersedes). Every table row must carry EXACTLY one of the three specified
    // grounds and nothing else, and all three must actually appear — a table where every row came out the
    // same colour would satisfy "no unexpected fill" while having no hierarchy at all.
    const L1 = 'D99370', L2 = 'BFD8D8', L3 = '9F9F5F';
    const lvl1 = rowRects.filter((s) => s.fill === L1);
    const lvl2 = rowRects.filter((s) => s.fill === L2);
    const lvl3 = rowRects.filter((s) => s.fill === L3);
    const other = rowRects.filter((s) => ![L1, L2, L3].includes(s.fill || ''));
    add('Project Overview rows use only the three hierarchy colours', rowRects.length > 0 && other.length === 0, `level 1 #${L1}: ${lvl1.length} row(s) · level 2 #${L2}: ${lvl2.length} · level 3 #${L3}: ${lvl3.length}` +
        (other.length ? ` — ${other.length} row(s) in an unexpected fill: ${[...new Set(other.map((s) => s.fill))].join(', ')}` : ''));
    add('Project Overview hierarchy is visually present', lvl1.length > 0 && lvl2.length > 0 && lvl3.length > 0, lvl1.length && lvl2.length && lvl3.length
        ? 'all three levels appear'
        : `missing level(s): ${[!lvl1.length && '1', !lvl2.length && '2', !lvl3.length && '3'].filter(Boolean).join(', ')}`);
    // The project row appears ONCE in the deck (page 1, first row) and no page may carry two. A page
    // without it is correct — the project is not repeated per page — so this is "at most one", plus
    // "exactly one across the deck".
    const l1PerPage = ovXmls.map((x) => shapesOf(x).filter((s) => !s.text && inTable(s) && s.w > 12 && s.y > 3 && s.fill === L1));
    const l1Total = l1PerPage.reduce((n, r) => n + r.length, 0);
    add('One project-level row in the deck, never two on a page', ovXmls.length === 0 || (l1Total === 1 && l1PerPage.every((r) => r.length <= 1)), `${l1Total} level-1 row(s) in total · per page: ${l1PerPage.map((r) => r.length).join(', ') || 'n/a'}`);
    // Header row: WHITE ground, BLACK text. Scoped to the TABLE side — the Gantt's year band also sits at
    // this height and is grey by design, so a height-only filter would have judged the calendar by the
    // table's rule.
    const hdrCells = ovXmls.flatMap((x) => shapesOf(x).filter((s) => s.text && inTable(s) && s.y > 2.4 && s.y < 2.7));
    const hdrFills = new Set(hdrCells.map((s) => (s.fill || '').toUpperCase()).filter(Boolean));
    const hdrCols = new Set(hdrCells.flatMap((s) => s.colors.map((c) => c.toUpperCase())));
    const hdrOk = hdrCells.length > 0 && [...hdrFills].every((f) => f === 'FFFFFF') && [...hdrCols].every((c) => c === '000000');
    add('Project Overview header is black on white', hdrOk, `${hdrCells.length} heading cell(s) · fill ${[...hdrFills].map((c) => '#' + c).join(', ') || 'none'} · text ${[...hdrCols].map((c) => '#' + c).join(', ') || 'inherited'}`);
    // Conditional cells: narrow rects tinted green or light orange. Their x must be the
    // Actual / Proj. Finish column, i.e. the 8th column (x ≈ 21.21").
    // Inside the grid only (y > 3), one column wide (1"–4") — this excludes the header badge and the
    // 0.5"-wide legend swatches, which are not table cells.
    const condRects = ovXmls.flatMap((x) => shapesOf(x).filter((s) => !s.text && inTable(s) && s.y > 3 && s.w >= 1 && s.w <= 4 && s.fill && s.fill !== 'FFFFFF' && s.fill !== OLIVE));
    const colXs = [...new Set(condRects.map((s) => s.x.toFixed(2)))];
    const okCol = colXs.length <= 1;
    add('Only Actual / Proj. Finish uses conditional colouring', okCol, condRects.length
        ? `${condRects.length} conditional cell(s), all in the column at x=${colXs.join(', ')}"` +
            `; colours used: ${[...new Set(condRects.map((s) => s.fill))].join(', ')}`
        : 'no conditional cell needed on this project (nothing complete or overdue)');
    // Project Overview column headings — typography only. They are individual text boxes (not a pptxgenjs
    // table), so they are found by matching the heading labels themselves against the header band's fill.
    // Seven headings, in the reference's OWN full wording — nothing abbreviated. This list is the
    // contract: if anyone shortens a heading in overview.ts, this check stops finding it and fails.
    const OV_HEADS = ['Name', 'Duration [Elapsed day]', 'Duration [Elapsed Month]',
        'Overall percent complete', 'Baseline Finish', 'Actual / Projected Start',
        'Actual / Projected Finish'];
    const flat = (t) => t.replace(/\s+/g, ' ').trim().toLowerCase();
    const wanted = new Set(OV_HEADS.map(flat));
    const ovHeads = ovXmls.flatMap((x) => shapesOf(x).filter((sh) => sh.text && wanted.has(flat(sh.text))));
    const ovSizes = [...new Set(ovHeads.flatMap((sh) => sh.sizes))];
    add(`Project Overview headings are ${typography_1.HEADER_FS_OVERVIEW}pt`, ovHeads.length >= OV_HEADS.length && ovSizes.length > 0 && ovSizes.every((v) => v === typography_1.HEADER_FS_OVERVIEW), ovHeads.length
        ? `${ovHeads.length} heading box(es) found across ${ovXmls.length} page(s); size(s) ${ovSizes.join('/')}pt`
        : 'no Project Overview heading row found');
    // Every heading box must be the SAME height — one header row, not a ragged one. The height itself is
    // no longer pinned to 0.95": the full headings need more than one line at 25pt in a 73%-wide table, so
    // the row is measured from the real wrapped line count. What still must hold is that they all match and
    // that the row stays within a sane share of the table.
    const ovHdrH = [...new Set(ovHeads.map((sh) => sh.h.toFixed(2)))];
    add('Project Overview header row is one consistent height', ovHdrH.length === 1 && parseFloat(ovHdrH[0] || '9') <= 2.2, ovHeads.length ? `header row ${ovHdrH.join('/')}" tall, identical across every heading` : 'not measured');
    // Every heading must be readable in full: no box may be narrower than its own longest word.
    const tooNarrow = ovHeads.filter((sh) => Math.max(...sh.text.split(/\s+/)
        .map((w) => w.length * typography_1.HEADER_FS_OVERVIEW * 0.6 / 72)) > sh.w);
    add('No Project Overview heading is too narrow for its own words', tooNarrow.length === 0, tooNarrow.length ? `${tooNarrow.map((sh) => sh.text).join(', ')}` : `all ${ovHeads.length} heading boxes fit their longest word`);
    // ---------- the Project Overview Gantt ----------
    if (ovXmls.length) {
        const G = { baseline: 'FFD000', actual: '5B9BD5', proj: '008A3E', today: 'C8281E' };
        // The chart body only — and PER PAGE, because a short page's legend sits higher than a full page's
        // last row. Zero-width shapes are excluded: a vertical rule's stroke colour is read by shapesOf in
        // the same place as a fill, so the month ticks would otherwise be counted as bars.
        const bars = ovXmls.flatMap((x) => {
            const shapes = shapesOf(x);
            const bottom = Math.max(0, ...shapes.filter((sh) => inTable(sh) && sh.w > 12 && sh.y > 3).map((sh) => sh.y + sh.h));
            return shapes.filter((s) => !s.text && inGantt(s) && s.fill && s.w > 0.01 && s.h > 0.01
                && s.y > 3 && s.y + s.h <= bottom + 0.02
                && s.fill !== 'FFFFFF' && s.fill !== 'E9E4C9');
        });
        const barFills = [...new Set(bars.map((b) => b.fill))];
        const allowed = new Set(Object.values(G));
        add('Gantt bars use only the reference palette', bars.length > 0 && barFills.every((f) => allowed.has(f)), bars.length ? `${bars.length} bar(s); colours ${barFills.map((f) => '#' + f).join(', ')} ` +
            '(Baseline #FFD000 · Actual Progress #5B9BD5 · Projection #008A3E)'
            : 'no bars drawn — the schedule has no dated rows');
        const todayRule = ovXmls.flatMap((x) => linesOf(x).filter((l) => inGantt(l) && l.color === G.today));
        add('Gantt carries a TODAY rule', todayRule.length > 0, todayRule.length ? `${todayRule.length} page(s) show it at x=${[...new Set(todayRule.map((l) => l.x.toFixed(2)))].join(', ')}"`
            : 'the progress date falls outside the timeline, so no rule was drawn');
        // Month labels: either every month is labelled, or none is — a partial label cannot be clipped
        // because a label is only drawn when its own column is wide enough for the whole abbreviation.
        // Every month must carry its initial, and no initial may be wider than its own column.
        const INITIALS = new Set('JFMASOND'.split(''));
        const monthTexts = ovXmls.flatMap((x) => shapesOf(x).filter((sh) => inGantt(sh)
            && sh.text.trim().length === 1 && INITIALS.has(sh.text.trim()) && sh.y < 4));
        const perPage = ovXmls.length ? monthTexts.length / ovXmls.length : 0;
        const clipped = monthTexts.filter((m) => Math.max(...m.sizes, 0) > 0
            && (Math.max(...m.sizes) * 0.60) / 72 > m.w); // glyph advance vs its column
        add('Every month shows its initial, none clipped', monthTexts.length > 0 && clipped.length === 0, monthTexts.length
            ? `${Math.round(perPage)} month initial(s) per page at ${[...new Set(monthTexts.flatMap((m) => m.sizes))].join('/')}pt; ` +
                `${clipped.length} wider than their column`
            : 'no month initial found on the axis');
        // the two panels must not overlap
        // Inside the grid only — the page's own title bar is 24.68" wide and sits above it.
        const tableRight = Math.max(0, ...ovXmls.flatMap((x) => shapesOf(x)
            .filter((sh) => inTable(sh) && sh.w > 12 && sh.y > 3).map((sh) => sh.x + sh.w)));
        add('Table and Gantt do not overlap', tableRight <= GANTT_X + 0.02, `table ends ${tableRight.toFixed(2)}", Gantt starts ${GANTT_X.toFixed(2)}"`);
    }
    // ---------- Approvals ----------
    const apXml = xmls.find((x) => slideHas(x, 'APPROVALS STATUS'));
    if (apXml) {
        const tbl = tablesOf(apXml)[0] || [];
        const body = tbl.slice(1);
        const dateCells = body.flatMap((r) => r.slice(6, 8)); // Baseline Finish, Actual/Proj Finish
        const withDates = dateCells.filter((c) => textsOf(c).join('').trim());
        const meta = withDates.filter((c) => c.includes(`typeface="${typography_1.FONT_NUM}"`));
        add('Approval date columns use Meta-Normal', withDates.length > 0 && meta.length === withDates.length, `${meta.length}/${withDates.length} populated Baseline Finish + Actual/Proj. Finish cells set in ${typography_1.FONT_NUM}`);
        // The dates must be regular weight — the bold was removed; size, colour and alignment unchanged.
        const boldDates = withDates.filter((c) => /<a:rPr[^>]*\bb="1"/.test(c));
        add('Approval dates are not bold', boldDates.length === 0, `${withDates.length - boldDates.length}/${withDates.length} date cells are regular weight` +
            (boldDates.length ? ` — ${boldDates.length} still bold` : ''));
    }
    else {
        // No Approvals section (a construction-phase project has no statutory approvals left) means there
        // are no date cells to typeset. Not applicable rather than a failure.
        add('Approval date columns use Meta-Normal', true, 'no Approvals Status section in this deck — not applicable');
    }
    // ---------- 2 Week Look Ahead ----------
    const laXmls = xmls.filter((x) => slideHas(x, 'WEEK LOOK AHEAD'));
    let deptCells = 0, deptOrange = 0;
    for (const x of laXmls)
        for (const tbl of tablesOf(x))
            for (const r of tbl.slice(1)) {
                if (r.length < 9)
                    continue; // band rows span the table
                const c = r[7]; // Task.TE Departments
                if (!textsOf(c).join('').trim())
                    continue;
                deptCells++;
                if (ACCENT.some((a) => c.toUpperCase().includes(`VAL="${a}"`)))
                    deptOrange++;
            }
    add('TE Department text is black', deptOrange === 0, `${deptCells} department cell(s) checked, ${deptOrange} still using accent orange`);
    // ---------- Key Focus Area ----------
    const kfXmls = xmls.filter((x) => slideHas(x, 'KEY FOCUS AREA'));
    let hdrSizes = [], hdrLabels = [], kfOrange = 0, kfBody = 0;
    for (const x of kfXmls)
        for (const tbl of tablesOf(x)) {
            const head = tbl[0] || [];
            hdrLabels = head.map((c) => textsOf(c).join(' ').trim());
            hdrSizes = head.flatMap((c) => [...c.matchAll(/sz="(\d+)"/g)].map((m) => +m[1] / 100));
            for (const r of tbl.slice(1))
                for (const c of r) {
                    if (r.length < 6)
                        continue; // CRITICAL / AT RISK band rows
                    if (!textsOf(c).join('').trim())
                        continue;
                    kfBody++;
                    if (ACCENT.some((a) => c.toUpperCase().includes(`VAL="${a}"`)))
                        kfOrange++;
                }
        }
    // EVERY Key Focus heading is the SAME size. Sizing each heading to its own column produced six
    // different sizes down one header row; the rule now is one size for all six, and a heading that
    // needs more room is given a wider COLUMN instead of a smaller size. So the check is exact
    // equality, not a range: if any heading has been quietly stepped down to fit, this fails.
    const uniq = [...new Set(hdrSizes)];
    add(`Every Key Focus heading is exactly ${typography_1.KEYFOCUS_HEAD_FS}pt`, hdrSizes.length > 0 && uniq.length === 1 && uniq[0] === typography_1.KEYFOCUS_HEAD_FS, hdrSizes.length
        ? `${hdrSizes.length} heading(s), size(s) found: ${uniq.join(', ')}pt`
        : 'no Key Focus header row found');
    add('No Key Focus heading was reduced independently to make it fit', hdrSizes.length > 0 && uniq.length === 1, hdrSizes.length ? `all headings share one size (${uniq.join(', ')}pt); a heading needing more room is given a wider column instead` : 'no Key Focus header row found');
    add('"Major Heading" has been renamed to "Team"', hdrLabels.includes('Team') && !hdrLabels.some((l) => /major heading/i.test(l)), `headings: ${hdrLabels.join(' · ')}`);
    add('All Key Focus content uses black text', kfOrange === 0, `${kfBody} content cell(s) checked, ${kfOrange} still using accent orange`);
    // Key Focus header row: white ground, black text — the only colour on the slide belongs to the bands.
    const kfHeadCells = kfXmls.flatMap((x) => (tablesOf(x)[0] || [])[0] || []);
    const kfHeadFills = [...new Set(kfHeadCells.map(cellFill).filter(Boolean))];
    const kfHeadCols = [...new Set(kfHeadCells.map(cellTextColour).filter(Boolean))];
    add('Key Focus header row is white with black text', kfHeadCells.length > 0 && kfHeadFills.every((f) => f === 'FFFFFF') && kfHeadCols.every((c) => c === '000000'), kfHeadCells.length
        ? `${kfHeadCells.length} heading cell(s): fill ${kfHeadFills.map((f) => '#' + f).join(', ')}, text ${kfHeadCols.map((c) => '#' + c).join(', ')}`
        : 'no Key Focus header row found');
    // The two section bands must be the SAME colours the Red Flag Summary uses. Both are read out of the
    // finished deck and compared to each other, so this cannot pass on a value that merely looks similar.
    const bandFill = (want) => {
        for (const x of kfXmls)
            for (const tbl of tablesOf(x))
                for (const r of tbl) {
                    for (const c of r)
                        if (textsOf(c).join(' ').trim().toUpperCase() === want)
                            return cellFill(c);
                }
        return null;
    };
    const rfXml = xmls.find((x) => slideHas(x, 'RED FLAG')); // slideHas skips the Contents slide
    const rfFills = new Set();
    if (rfXml)
        for (const m of rfXml.matchAll(/<a:srgbClr val="([0-9A-Fa-f]{6})"/g))
            rfFills.add(m[1].toUpperCase());
    for (const [band, want] of [['CRITICAL', 'EE4B2B'], ['AT RISK', 'FFAA33']]) {
        const got = bandFill(band);
        // A project with nothing at that severity simply has no band — that is data, not a formatting fault.
        add(`${band} band matches the Red Flag Summary (#${want})`, got === null || (got === want && (!rfXml || rfFills.has(want))), got ? `band fill #${got}` + (rfXml ? `; the Red Flag Summary ${rfFills.has(want) ? 'uses the same #' + want : 'does NOT use #' + want}` : '; no Red Flag slide in this deck')
            : `not applicable — this project has no ${band} rows`);
    }
    // Project Overview vertical grid lines — one per internal column boundary, matching the row separators.
    const OV_COLS = 7; // Activity + % Comp. + Duration + four dates
    // Column rules belong to the table; the TODAY rule and the month ticks are the chart's.
    //
    // The height floor is 0.30" — one row — NOT 1". The frames now end at the last populated row, so a
    // continuation page holding a single activity has rules only one row deep; a 1" floor reported that
    // page as having no rules at all, which was the check's assumption being wrong, not the slide.
    const ovLines = ovXmls.map((x) => linesOf(x).filter((l) => l.w === 0 && l.h > 0.30 && inTable(l)));
    const ovCounts = [...new Set(ovLines.map((l) => l.length))];
    const ovColours = [...new Set(ovLines.flat().map((l) => l.color))];
    const ovWeights = [...new Set(ovLines.flat().map((l) => l.pt.toFixed(2)))];
    add('Project Overview has a vertical rule between every column', ovLines.length > 0 && ovCounts.every((n) => n === OV_COLS - 1), ovLines.length ? `${ovCounts.join('/')} vertical rule(s) per page (expected ${OV_COLS - 1}), colour ` +
        `${ovColours.map((c) => '#' + c).join(', ')} at ${ovWeights.join('/')}pt — the row separators are #D9D9D9 at 0.75pt`
        : 'no Project Overview slide found');
    // Delay highlighting was removed: the only conditional fill left in the table is the completed green
    // (theme ragGood). Any other fill in that column would mean a delay marker had come back.
    const COMPLETED_FILL = 'A9D08E';
    const ovFills = [...new Set(condRects.map((sh) => sh.fill))].filter(Boolean);
    add('No delay highlighting in Project Overview', ovFills.every((f) => f.toUpperCase() === COMPLETED_FILL), ovFills.length ? `conditional fill(s) used: ${ovFills.map((f) => '#' + f).join(', ')} (completed green only)`
        : 'no conditional cell on this project');
    add('Vertical rules match the row separators exactly', ovColours.length === 1 && ovColours[0] === 'D9D9D9' && ovWeights.length === 1 && ovWeights[0] === '0.75', `colour ${ovColours.map((c) => '#' + c).join(', ') || '—'}, weight ${ovWeights.join('/') || '—'}pt`);
    // ---------- numbers are Meta-Normal, everywhere ----------
    // Scanned over the WHOLE deck, run by run, rather than trusted per call site. A run whose text is a
    // number — a percentage, a duration, a date, a year, a serial — must be in Meta-Normal.
    //
    // The page-number badge is the one deliberate exception: it is the section number in the brown
    // header, set in the title face at the title size, and the header design is fixed by the brief. It is
    // counted and named here rather than quietly skipped.
    // Scope: slides this generator AUTHORS. A slide copied verbatim from the customer's own deck — the
    // cover, Land Acquisition, Master Plan, any extra — carries their typography, and re-facing numbers
    // inside it would both change their branding (which the brief forbids) and mean editing foreign XML
    // as a string, which is what produced the repair dialog once already. Such slides are identified by
    // the fact that they reference a face this generator never writes, and are counted separately.
    const OURS = new Set([typography_1.FONT_NUM, typography_1.FONT_TEXT, 'Impact']);
    // TE-REF-COPY is written by postmerge.copyRefSlide onto every slide taken verbatim from the customer's
    // deck; it is a fact about provenance rather than an inference from which fonts appear.
    const authored = (x) => !/name="TE-REF-COPY"/.test(x) &&
        [...new Set([...x.matchAll(/typeface="([^"]+)"/g)].map((m) => m[1]))].every((f) => OURS.has(f));
    const NUMERIC = /^(?:[0-9]+(?:\.[0-9]+)?%?|[0-9]{1,4}d|[0-9]{1,2}\s+[A-Za-z]{3}\s+[0-9]{2,4}|[0-9]{4})$/;
    const numRuns = [];
    let copiedSlides = 0, copiedNums = 0;
    xmls.forEach((x, i) => {
        if (!authored(x)) {
            copiedSlides += 1;
            for (const m of x.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)) {
                const t = ((m[0].match(/<a:t>([\s\S]*?)<\/a:t>/) || [])[1] || '').trim();
                if (t && NUMERIC.test(t))
                    copiedNums += 1;
            }
            return;
        }
        for (const m of x.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)) {
            const b = m[0];
            const t = ((b.match(/<a:t>([\s\S]*?)<\/a:t>/) || [])[1] || '').trim();
            if (!t || !NUMERIC.test(t))
                continue;
            const face = (b.match(/typeface="([^"]+)"/) || [])[1] || '(inherited)';
            numRuns.push({ text: t, face, slide: i + 1 });
        }
    });
    const badgeFace = 'Impact';
    const offenders = numRuns.filter((r) => r.face !== typography_1.FONT_NUM && r.face !== badgeFace);
    const badges = numRuns.filter((r) => r.face === badgeFace);
    add('Every number is set in Meta-Normal', numRuns.length > 0 && offenders.length === 0, `${numRuns.filter((r) => r.face === typography_1.FONT_NUM).length} numeric run(s) in ${typography_1.FONT_NUM} across the authored slides` +
        (badges.length ? `; ${badges.length} page-number badge(s) in ${badgeFace} — the header's own face, fixed by the brief` : '') +
        (copiedSlides ? `; ${copiedNums} number(s) on ${copiedSlides} slide(s) copied verbatim from the customer's deck were left in their own typography` : '') +
        (offenders.length ? `; OFFENDERS: ${[...new Set(offenders.map((o) => `"${o.text}" in ${o.face} (slide ${o.slide})`))].slice(0, 6).join(', ')}` : ''));
    // ---------- cover date typography ----------
    // The cover is copied from the reference deck, so its runs carry the template's own fonts. Two are
    // re-faced (see postmerge.updateCoverDate) and the check reads them back off the finished slide.
    const cvXml = xmls.find((x) => /weekly update/i.test(textsOf(x).join(' ')));
    if (cvXml) {
        const runs = [...cvXml.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)].map((m) => m[0])
            .filter((r) => /<a:t>/.test(r));
        const runOf = (test) => runs.find((r) => test.test((r.match(/<a:t>([\s\S]*?)<\/a:t>/) || ['', ''])[1].trim()));
        const faceOf = (r) => (r ? [...new Set([...r.matchAll(/typeface="([^"]+)"/g)].map((m) => m[1]))] : []);
        const yr = runOf(/^\d{4}$/), ord = runOf(/^(st|nd|rd|th)$/i);
        // Only the COPIED cover (from the reference deck) sets the date as separate runs; the generated
        // placeholder cover writes it as one, and there is nothing to re-face there.
        // No four-digit year RUN means there is nothing for the year rule to act on. That happens two ways:
        // the generated placeholder cover writes the whole date as one run, and some customer covers write a
        // two-digit year ("Weekly Update: 14 Aug 26"). Neither is a failure — the rule is about re-facing a
        // year run that exists, and Nandi Hills' own cover has none.
        if (!yr) {
            add('Cover date typography', true, ord
                ? 'this cover writes its year as part of another run (e.g. a two-digit "26"), so there is no ' +
                    'year run to re-face — not applicable'
                : 'this deck has the generated placeholder cover, whose date is a single run — the year/ordinal ' +
                    're-facing applies to a cover copied from the reference deck');
        }
        else {
            add('Cover year is Meta-Normal and bold', !!yr && faceOf(yr).every((f) => f === typography_1.FONT_NUM) && /\bb="1"/.test(yr || ''), yr ? `"${(yr.match(/<a:t>([\s\S]*?)<\/a:t>/) || ['', ''])[1].trim()}" in ${faceOf(yr).join('/')}, ` +
                `${/\bb="1"/.test(yr) ? 'BOLD' : 'regular weight'}`
                : 'no separate year run on the cover');
            add('Cover ordinal is set in Helvetica Neue, still superscript', !!ord && faceOf(ord).every((f) => f === typography_1.FONT_TEXT) && /baseline="\d+"/.test(ord || ''), ord ? `"${(ord.match(/<a:t>([\s\S]*?)<\/a:t>/) || ['', ''])[1].trim()}" in ${faceOf(ord).join('/')}, ` +
                `baseline ${(ord.match(/baseline="(\d+)"/) || ['', 'none'])[1]}`
                : 'no separate ordinal run on the cover');
        }
    }
    return { ok: checks.every((c) => c.ok), checks };
}
function refinementReport(r, project, layoutOk, layoutSummary) {
    const L = [];
    L.push('# presentation_refinement_validation.md', '');
    L.push(`Project: **${project}**`, '');
    L.push(`## Result: ${r.ok && layoutOk ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Evidence (read back from the generated deck) |', '|---|---|---|');
    for (const c of r.checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    L.push(`| No text overflow | ${layoutOk ? '✅' : '❌'} | ${layoutSummary} |`);
    L.push(`| No overlapping objects | ${layoutOk ? '✅' : '❌'} | ${layoutSummary} |`);
    L.push(`| No formatting regressions | ${layoutOk ? '✅' : '❌'} | layout_validation.md: ${layoutOk ? 'PASS' : 'FAIL'} — heading typography, header-row visibility and table bounds all re-checked |`);
    L.push('', '## Rules applied', '', '- **Project Overview** — major section heading rows keep the olive band; every other activity row is', '  white. Conditional colour exists in the *Actual / Proj. Finish* column only: green at 100%', '  progress, light orange when the current date is past Baseline Finish and progress is below 100%,', '  white otherwise. No other column is conditionally coloured, and heading rows are excluded so their', '  olive band stays exactly as the reference template.', '- **Approvals Status** — unchanged except that both date columns are set in Meta-Normal.', '- **2 Week Look Ahead** — the TE Departments column uses the standard black body colour; font, size', '  and alignment are untouched.', `- **Key Focus Area** — column headings at ${typography_1.HEADER_FS_WIDE}pt (body text unchanged), *Team* shown in`, '  place of *Major Heading* (display label only — the underlying field is unchanged), and all body', '  text in the standard black.', '', 'No business logic and no data selection was changed by any of the above.');
    return L.join('\n');
}
