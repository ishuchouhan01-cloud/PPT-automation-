"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.keyFocusLayout = exports.kfDropped = exports.KF_ROW_CAP = exports.TABLE_FS = exports.badgeFs = exports.HDR = exports.HDR_BAR_FILL = exports.HDR_BADGE_FILL = exports.fmtDate = void 0;
exports.sectionHeader = sectionHeader;
exports.cover = cover;
exports.contents = contents;
exports.thankYou = thankYou;
exports.matrix = matrix;
exports.keyFocusCell = keyFocusCell;
exports.keyFocus = keyFocus;
exports.approvals = approvals;
exports.lookAheadSlides = lookAheadSlides;
exports.imageSlide = imageSlide;
exports.placeholder = placeholder;
const blocks_1 = require("./blocks");
const theme_1 = require("./theme");
const typography_1 = require("./typography");
const MO = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const fmtDate = (s) => {
    if (!s)
        return '';
    const d = new Date(s + 'T00:00:00Z');
    return `${String(d.getUTCDate()).padStart(2, '0')} ${MO[d.getUTCMonth()]} ${String(d.getUTCFullYear()).slice(2)}`;
};
exports.fmtDate = fmtDate;
// ---- ONE standard header for every generated slide ----
// The two template shades are accent·lumMod75 (lighter) and accent·lumMod50 (darker). They are now
// used the same way on every generated slide, with the pair SWAPPED from the original scheme:
//   • section-number box → the shade the page title used to have
//   • page-title bar     → the shade the number box used to have
const BAR = (t) => t.accentBar || t.accent; // lighter template shade
const ICON = (t) => t.accentIcon || t.accent; // darker template shade
/** Fill of the section-number box (standard across the deck). */
const HDR_BADGE_FILL = (t) => BAR(t);
exports.HDR_BADGE_FILL = HDR_BADGE_FILL;
/** Fill of the page-title bar (standard across the deck). */
const HDR_BAR_FILL = (t) => ICON(t);
exports.HDR_BAR_FILL = HDR_BAR_FILL;
// Geometry, identical on every slide: the badge is a fixed-width box and the title bar starts a clear
// 0.08" later, so a two-character label ("4A") can never touch or overlap the title.
exports.HDR = { badgeX: 0.57, badgeW: 0.78, barX: 1.43, barW: 24.68, y: 1.12, h: 1.04, textPad: 0.22 };
/**
 * Badge label size. It is the page title's size — the number belongs to the title and must read at the
 * same weight on every slide — and is reduced only if the label would not fit its fixed 0.78" box.
 * A two-character label ("4A") measures 0.59" at 40pt, so nothing is reduced in practice; the guard is
 * there for a hypothetical three-character label. Previously every 2-character label was set at 68% of
 * the title size, which is why 4A / 4B looked smaller than every other slide's number.
 */
const badgeFs = (label) => {
    const inner = exports.HDR.badgeW - 0.10; // the box, less a little breathing room
    const w = (0, typography_1.textWidth)(label, typography_1.TITLE_FS, typography_1.FONT_TITLE);
    return w <= inner ? typography_1.TITLE_FS : Math.max(typography_1.MIN_FS, Math.floor(typography_1.TITLE_FS * (inner / w)));
};
exports.badgeFs = badgeFs;
// HR-approved font roles: titles Impact, general text Helvetica Neue, all numbers Meta-Normal.
const FT = (_t) => typography_1.FONT_TEXT;
exports.TABLE_FS = typography_1.BODY_FS;
const HDR_FS = typography_1.HEADER_FS;
const T = (t) => ({
    head: { fill: { color: BAR(t) }, color: t.headerText, bold: true, fontFace: typography_1.FONT_TEXT, fontSize: HDR_FS },
    // Wide tables (Key Focus) use the larger heading size so the header row reads clearly against the
    // 28pt body text. Its ground is WHITE with black text, per the Total Environment style: on this
    // slide the only colour belongs to the CRITICAL / AT RISK section bands, so an orange header row
    // competed with them. Family, size, weight and alignment are unchanged.
    headWide: { fill: { color: 'FFFFFF' }, color: '000000', bold: true, fontFace: typography_1.FONT_TEXT, fontSize: typography_1.HEADER_FS_WIDE },
});
/** Table cell that renders numbers in Meta-Normal and everything else in Helvetica Neue. */
const cell = (v, extra = {}) => ({ text: v, options: { fontFace: (0, typography_1.fontFor)(v), ...extra } });
function sectionHeader(slide, title, num, t, 
/**
 * Fill for the title bar. Defaults to the deck-wide shade, so every other slide is untouched; the
 * 2 Week Look Ahead passes LOOKAHEAD.titleBar. Geometry, font, size, padding, text colour and the
 * numbered badge are all unchanged — only this one rectangle's fill differs.
 */
barFill, 
/**
 * Vertical position of the header band. Defaults to HDR.y, so every other slide is untouched; the
 * 2 Week Look Ahead passes a slightly higher value to sit closer to the wordmark and free vertical
 * space for its table. Height, width, x positions, font and padding are unchanged either way.
 */
yTop) {
    const HY = yTop ?? exports.HDR.y;
    const label = String(num);
    slide.addShape('rect', { x: exports.HDR.badgeX, y: HY, w: exports.HDR.badgeW, h: exports.HDR.h, fill: { color: (0, exports.HDR_BADGE_FILL)(t) }, line: { type: 'none' } });
    slide.addText(label, {
        x: exports.HDR.badgeX, y: HY, w: exports.HDR.badgeW, h: exports.HDR.h, align: 'center', valign: 'middle',
        color: t.headerText, fontFace: typography_1.FONT_TITLE, fontSize: (0, exports.badgeFs)(label), margin: 0, wrap: false,
    });
    slide.addShape('rect', { x: exports.HDR.barX, y: HY, w: exports.HDR.barW, h: exports.HDR.h, fill: { color: barFill || (0, exports.HDR_BAR_FILL)(t) }, line: { type: 'none' } });
    slide.addText(title.toUpperCase(), {
        x: exports.HDR.barX + exports.HDR.textPad, y: HY, w: exports.HDR.barW - exports.HDR.textPad * 2, h: exports.HDR.h, valign: 'middle',
        color: t.headerText, fontFace: typography_1.FONT_TITLE, fontSize: typography_1.TITLE_FS,
    });
}
function logos(slide, assets) {
    if (assets.wordmark)
        slide.addImage({ data: assets.wordmark, x: 0.57, y: 0.39, w: 3.47, h: 0.5 });
    if (assets.mascot)
        slide.addImage({ data: assets.mascot, x: 0.61, y: 13.74, w: 0.74, h: 0.79 });
}
function cover(pptx, sch, code, subtitle, reportDate, t, assets) {
    const s = pptx.addSlide();
    s.background = { color: t.accent };
    if (assets.cover)
        s.addImage({ data: assets.cover, x: 0, y: 0.18, w: 26.66, h: 15.02 });
    const title = code ? `${code} | ${sch.projectName}` : sch.projectName;
    s.addText([
        { text: title, options: { fontSize: 60, bold: true, color: 'FFFFFF', breakLine: true } },
        ...(subtitle ? [{ text: subtitle, options: { fontSize: 30, color: 'F0E0D0', breakLine: true } }] : []),
        { text: 'Weekly Update  ' + (0, exports.fmtDate)(reportDate), options: { fontSize: 34, bold: true, color: 'FFC000' } },
    ], { x: 0.4, y: 4.28, w: 16.5, h: 3.5, fontFace: typography_1.FONT_TEXT, valign: 'top' });
}
/**
 * The Contents slide.
 *
 * It is created EMPTY and returns a `fill(items)` callback, because the list has to be written from the
 * sections that actually got generated — and those are not known until the whole deck has been built.
 * pptxgenjs serialises at write() time, so a slide made early can still be populated later and keeps its
 * position (second, straight after the cover).
 *
 * Nothing here is hard-coded: the numbering is the index in the list it is given, so a section that was
 * not generated simply is not listed and the ones after it move up.
 */
function contents(pptx, t, assets) {
    const s = pptx.addSlide();
    s.addShape('rect', { x: 0.61, y: 1.12, w: 25.39, h: 1.04, fill: { color: (0, exports.HDR_BAR_FILL)(t) }, line: { type: 'none' } });
    s.addText('CONTENTS', { x: 0.9, y: 1.12, w: 24, h: 1.04, valign: 'middle', color: 'FFFFFF', fontFace: typography_1.FONT_TITLE, fontSize: typography_1.TITLE_FS });
    logos(s, assets);
    return (items) => {
        // Contents list: 40pt, anchored top-left directly under the header bar and flush with its left
        // edge (x 1.43) so it lines up with every other slide's content column.
        // Every number in the deck is set in Meta-Normal, so the list's own numbering is split into its own
        // run: "1." in Meta-Normal, the section title in Helvetica Neue. Size, colour and spacing unchanged.
        // Line spacing is 1.3 (was the implicit single spacing). It is set as a MULTIPLE rather than a
        // point value so it stays correct if the 40pt size is ever changed. Numbering, section names,
        // font family, size, colour, alignment and paraSpaceAfter are untouched, and `items` is still
        // whatever the engine's section register actually generated — nothing here is hard-coded.
        const LINE_SPACING = 1.3;
        s.addText(items.flatMap((it, i) => [
            { text: `${i + 1}.`, options: { fontFace: typography_1.FONT_NUM } },
            { text: `   ${it}`, options: { breakLine: true, paraSpaceAfter: 14 } },
        ]), { x: 1.43, y: 2.6, w: 18, h: 10.6, align: 'left', valign: 'top', fontFace: typography_1.FONT_TEXT, fontSize: 40, color: t.dark, lineSpacingMultiple: LINE_SPACING });
    };
}
function thankYou(pptx, t, assets) {
    const s = pptx.addSlide();
    if (assets.thanks)
        s.addImage({ data: assets.thanks, x: 0, y: 0, w: 26.69, h: 12.33 });
    else
        s.background = { color: t.accent };
    // "THANK YOU" in black, bottom-RIGHT — mirroring the TotalEnvironment wordmark at bottom-left
    // (wordmark sits at x 0.73 w 3.84 y 13.13, so this is right-aligned to the mirrored margin).
    s.addText('THANK YOU', {
        x: 26.67 - 0.73 - 10, y: 13.13, w: 10, h: 0.7,
        align: 'right', valign: 'middle', color: '000000', fontFace: typography_1.FONT_TEXT, fontSize: 40,
    });
    s.addText('www.total-environment.com  |  discover@total-environment.com  |  +91 99000 78000', { x: 0.34, y: 13.83, w: 16, h: 0.6, color: t.dark, fontFace: typography_1.FONT_TEXT, fontSize: 18 });
    if (assets.wordmark)
        s.addImage({ data: assets.wordmark, x: 0.73, y: 13.13, w: 3.84, h: 0.55 });
}
// ---- Indicator matrix (Red Flag = colour, Tracker = symbol) ----
// The rotated sub-headings use the SAME centred geometry as the Tracker slide (headingLayout): a
// rotated text box rotates its border and turns about its centre, so each cell outline is drawn as its
// own un-rotated shape and the label sits on top, centred both ways. Long labels break over two
// vertical lines and the header band grows to fit, so nothing overlaps or leaves the slide.
function matrix(pptx, data, num, title, mode, t, assets) {
    const s = pptx.addSlide();
    sectionHeader(s, title, num, t);
    logos(s, assets);
    const x0 = 1.43, TOTAL_W = 24.68;
    const top = 2.6, lead = 1.0, nameW = 3.2;
    const n = Math.max(data.subs.length, 1);
    const indW = (TOTAL_W - lead - nameW) / n;
    const GRID = 'D9D9D9';
    // Break a long heading into two vertical lines so the band stays a sensible height.
    const wrapName = (v) => {
        if (v.length <= 18)
            return v;
        const mid = Math.floor(v.length / 2);
        for (let d = 0; d < v.length; d++) {
            if (v[mid - d] === ' ')
                return `${v.slice(0, mid - d)}\n${v.slice(mid - d + 1)}`;
            if (v[mid + d] === ' ')
                return `${v.slice(0, mid + d)}\n${v.slice(mid + d + 1)}`;
        }
        return v;
    };
    const wrapped = data.subs.map(wrapName);
    const longest = wrapped.reduce((m, v) => Math.max(m, ...v.split('\n').map((l) => (0, typography_1.textWidth)(l, typography_1.MIN_FS, typography_1.FONT_TEXT))), 0);
    // Group band: tall enough for the longest group label, which may need two lines ("Site
    // Establishment"), so it can never spill into the sub-heading band below.
    const groupLines = Math.max(1, ...data.macro.map(([g, span]) => (0, typography_1.linesNeeded)(g, span * indW - 0.1, typography_1.MIN_FS, typography_1.FONT_TEXT)));
    const H_GROUP = Math.max(0.45, groupLines * (0, typography_1.lineHeight)(typography_1.MIN_FS) + 0.10);
    const H_SUB = Math.min(4.2, Math.max(1.9, longest + 0.38)); // band as TALL as the longest label is WIDE
    const H_ROW = 0.55;
    const hdrH = H_GROUP + H_SUB;
    const cy = top + hdrH;
    // frozen lead columns, spanning the whole header block
    s.addShape('rect', { x: x0, y: top, w: lead, h: hdrH, fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.5 } });
    const leadHb = (0, typography_1.headingLayout)(['Project Code'], [x0], top, lead, hdrH, typography_1.MIN_FS, typography_1.FONT_TEXT)[0];
    s.addText(leadHb.label, {
        x: leadHb.boxX, y: leadHb.boxY, w: leadHb.boxW, h: leadHb.boxH,
        align: 'center', valign: 'middle', rotate: 270, margin: 0, wrap: false,
        fontSize: typography_1.MIN_FS, bold: true, fontFace: typography_1.FONT_TEXT, color: t.dark,
    });
    s.addText('Projects', { x: x0 + lead, y: top, w: nameW, h: hdrH, align: 'center', valign: 'middle', fontSize: typography_1.MIN_FS, bold: true, fontFace: typography_1.FONT_TEXT, color: t.dark, line: { color: GRID, width: 0.5 } });
    // group header row — names, order and spans exactly as the dashboard has them
    let gx = x0 + lead + nameW;
    for (const [g, span] of data.macro) {
        const w = span * indW;
        s.addText(g, { x: gx, y: top, w, h: H_GROUP, align: 'center', valign: 'middle', fill: { color: GRID }, fontSize: typography_1.MIN_FS, bold: true, fontFace: typography_1.FONT_TEXT, color: t.dark, line: { color: 'FFFFFF', width: 0.5 }, margin: 1, wrap: true });
        gx += w;
    }
    // rotated sub-headings + the single project row of indicators
    const colXs = data.subs.map((_, i) => x0 + lead + nameW + i * indW);
    const heads = (0, typography_1.headingLayout)(wrapped, colXs, top + H_GROUP, indW, H_SUB, typography_1.MIN_FS, typography_1.FONT_TEXT);
    const symMap = { good: ['✔', '2E741C'], amber: ['✺', 'BF8600'], red: ['✘', 'C00000'], na: ['NA', t.dark] };
    // Colours come from the dashboard's own legend swatches when available (Red Flag Summary), so the
    // palette is never hardcoded; the template theme is the fallback.
    const fillMap = {
        good: data.palette?.good || t.ragGood,
        amber: data.palette?.amber || t.ragAmber,
        red: data.palette?.red || t.ragRed,
    };
    data.subs.forEach((_, i) => {
        const cx = colXs[i];
        s.addShape('rect', { x: cx, y: top + H_GROUP, w: indW, h: H_SUB, fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.5 } });
        const hb = heads[i];
        s.addText(hb.label, {
            x: hb.boxX, y: hb.boxY, w: hb.boxW, h: hb.boxH,
            align: 'center', valign: 'middle', rotate: 270, margin: 0, wrap: true,
            fontSize: typography_1.MIN_FS, fontFace: typography_1.FONT_TEXT, color: t.dark,
        });
        const rag = data.cells[i];
        if (mode === 'symbol') {
            const [sym, col] = symMap[rag] || ['', t.dark];
            s.addText(sym, { x: cx, y: cy, w: indW, h: H_ROW, align: 'center', valign: 'middle', margin: 0, wrap: false, fontSize: typography_1.MIN_FS, bold: true, color: col, fontFace: typography_1.FONT_TEXT, line: { color: GRID, width: 0.5 } });
        }
        else {
            // "NA" must stay on ONE line in these narrow columns, so margins are zero and wrap is off.
            s.addText(rag === 'na' ? 'NA' : '', { x: cx, y: cy, w: indW, h: H_ROW, align: 'center', valign: 'middle', margin: 0, wrap: false, fontSize: typography_1.MIN_FS, fontFace: typography_1.FONT_TEXT, color: t.dark, fill: fillMap[rag] ? { color: fillMap[rag] } : undefined, line: { color: GRID, width: 0.5 } });
        }
    });
    s.addText(data.code, { x: x0, y: cy, w: lead, h: H_ROW, align: 'center', valign: 'middle', fontSize: typography_1.MIN_FS, fontFace: typography_1.FONT_NUM, color: t.dark, line: { color: GRID, width: 0.5 } });
    s.addText(data.project, { x: x0 + lead, y: cy, w: nameW, h: H_ROW, valign: 'middle', margin: 4, fontSize: typography_1.MIN_FS, fontFace: typography_1.FONT_TEXT, color: t.dark, line: { color: GRID, width: 0.5 } });
    // legend
    const ragCol = {
        good: data.palette?.good || t.ragGood,
        amber: data.palette?.amber || t.ragAmber,
        red: data.palette?.red || t.ragRed,
        na: t.dark,
    };
    const legend = mode === 'symbol'
        ? [['✔ Completed', '2E741C'], ['✺ In Progress', 'BF8600'], ['✘ Not Started', 'C00000'], ['NA  Not Applicable', t.dark]]
        // Red Flag legend: the dashboard's own rows, in its own order, with its own swatch colours.
        : (data.legend && data.legend.length
            ? data.legend.filter(([, r]) => r !== 'na').map(([lab, r]) => [lab, ragCol[r] || t.dark])
            : [['Good', ragCol.good], ['At Risk', ragCol.amber], ['Critical', ragCol.red]]);
    let lx = 13;
    for (const [lab, col] of legend) {
        if (mode === 'color')
            s.addShape('rect', { x: lx, y: cy + 1.0, w: 0.5, h: 0.3, fill: { color: col }, line: { type: 'none' } });
        s.addText(lab, { x: lx + (mode === 'color' ? 0.6 : 0), y: cy + 0.95, w: 2.4, h: 0.4, fontSize: typography_1.HEADER_FS, color: mode === 'symbol' ? col : t.dark, fontFace: typography_1.FONT_TEXT });
        lx += 3;
    }
}
// ---- Key Focus ----
/** Row caps per group. More rows than this are dropped, keeping the highest-priority ones
 *  (source order = the dashboard's own priority ranking). The font is NEVER shrunk to fit more. */
exports.KF_ROW_CAP = { Critical: 24, 'At Risk': 28 };
/** Rows dropped by the caps on the last keyFocus() call — reported in layout_validation.md. */
exports.kfDropped = {};
/**
 * The Key Focus Area cell, with the source cell's own emphasis.
 *
 * The Summary Dashboard authors bold the part of a note that carries the news. Reading the cell as one
 * string threw that away, so every note arrived flat. `focusRuns` carries the cell's runs, and this
 * turns them into pptxgenjs runs — same order, same wording, each run keeping only its own bold flag.
 *
 * Deliberately conservative:
 *   • no bold in the source  → the plain string, exactly as before, so unformatted cells are untouched
 *   • bold in the source     → runs, so bold stays bold AND normal stays normal in the same cell
 * The cell is never bolded as a whole because part of it is bold, and nothing is re-worded: the text
 * is whatever Excel holds. Font, size, colour and alignment come from the table's own options, as
 * before — only the weight of individual runs is set here.
 */
function keyFocusCell(r) {
    const runs = r.focusRuns;
    if (!runs || !runs.length || !runs.some((x) => x.bold))
        return r.focus;
    return { text: runs.map((x) => ({ text: x.text, options: { bold: !!x.bold } })) };
}
exports.keyFocusLayout = null;
function keyFocus(pptx, sch, code, rows, num, t, assets) {
    const groups = ['Critical', 'At Risk'];
    const flat = [];
    let last = '';
    for (const g of groups) {
        // Cap each group without touching the priority logic: the incoming order already IS the
        // dashboard's priority ranking, so the highest-priority rows are simply the leading ones.
        const all = rows.filter((x) => x.group === g);
        const cap = exports.KF_ROW_CAP[g] ?? all.length;
        exports.kfDropped[g] = Math.max(0, all.length - cap);
        for (const r of all.slice(0, cap)) {
            if (r.group !== last) {
                flat.push({ band: r.group });
                last = r.group;
            }
            flat.push(r);
        }
    }
    // Project Code is gone (one project per deck); the freed width goes to the remaining columns.
    // Key Focus Area holds by far the most information on this slide, so it is given priority: it was
    // widened from 10.28" to 13.61" (+32%) and the width came out of Responsible / Accountable / Team /
    // Delay, each reduced by the SAME factor (0.698) so their relative proportions are preserved.
    // Projects keeps its width — the project name is a fixed string and already wraps at 3.40".
    //   Projects 3.40 · Responsible 2.09 · Accountable 2.09 · Team 2.37 · Key Focus 13.61 · Delay 1.12
    // The table width is UNCHANGED: the six values still sum to EXACTLY 24.68", and because this is a
    // single const consumed by both rowHeight() and every addTable() call, the widths are identical on
    // every row and on every continuation page. Nothing here changes the slide size or the margins.
    const header = ['Projects', 'Responsible', 'Accountable', 'Team', 'Key Focus Area', 'Delay\n[Days]'];
    // ONE HEADING SIZE FOR ALL SIX COLUMNS. Sizing each heading to its own column, as this did before,
    // produced six different sizes down one header row (28 / 23 / 23 / 28 / 28 / 20.5pt), which reads as
    // an accident rather than a design. The row is now uniform, and where a heading needs more room the
    // COLUMN is widened instead of the type being shrunk — see the width block below.
    const KF_HEAD_FS_ALL = typography_1.KEYFOCUS_HEAD_FS;
    // 0.30" rather than the 0.18" the row heights use: textWidth() is an average-character estimate and
    // the real advance widths come from the installed font, so a heading sized to within a hair of the
    // column still wrapped in some renderers ("Accountable" needed 1.88" of a 1.91" cell and broke).
    // The extra clearance is the tolerance for that difference.
    const HEAD_INSET = 0.30;
    // Project Code is gone (one project per deck); the freed width went to the remaining columns, with
    // Key Focus Area given priority because it holds by far the most information on the slide.
    //
    // Delay [Days] is now ~10% wider (1.12" → 1.23", +9.8%), and that width comes out of Responsible,
    // Accountable and Team — each reduced by the same factor (0.983) so their proportions are preserved.
    // Projects and Key Focus Area are not touched to pay for it.
    //   Projects 3.40 · Responsible 2.05 · Accountable 2.05 · Team 2.33 · Key Focus 13.62 · Delay 1.23
    //
    // Then, because the heading size is fixed, any column too narrow for its own longest heading word at
    // that size is WIDENED to fit, and the extra comes from Key Focus Area, which is the only column
    // with real slack. The total is corrected back to exactly 24.68" at the end, so the table width, the
    // slide width and the margins are unchanged, and every row and continuation page uses these same
    // widths because this is a single const consumed by rowHeight() and every addTable() call.
    const KF_BASE_W = [3.40, 2.05, 2.05, 2.33, 13.62, 1.23];
    const KF_DONOR = 4; // Key Focus Area — the only column with slack
    const KF_TABLE_W = 24.68; // must equal TABLE_W below
    const kfWidthNotes = [];
    const KF_COL_W = (() => {
        const w = KF_BASE_W.slice();
        let borrowed = 0;
        header.forEach((h, i) => {
            if (i === KF_DONOR)
                return;
            const longest = h.split(/[\s\n]+/).reduce((a, b) => (b.length > a.length ? b : a), '');
            const need = Math.ceil(((0, typography_1.textWidth)(longest, KF_HEAD_FS_ALL, typography_1.FONT_TEXT) + HEAD_INSET) * 100) / 100;
            if (need > w[i]) {
                kfWidthNotes.push(`${h.replace('\n', ' ')} widened ${w[i].toFixed(2)}"→${need.toFixed(2)}" so "${longest}" fits on one line at ${KF_HEAD_FS_ALL}pt`);
                borrowed += need - w[i];
                w[i] = need;
            }
        });
        if (borrowed)
            w[KF_DONOR] = Math.round((w[KF_DONOR] - borrowed) * 100) / 100;
        const total = w.reduce((a, b) => a + b, 0);
        w[KF_DONOR] = Math.round((w[KF_DONOR] + (KF_TABLE_W - total)) * 100) / 100;
        return w;
    })();
    const NCOL = KF_COL_W.length;
    // The header row must be tall enough for the column that needs the most lines — all at one size now.
    const headLines = header.map((h, i) => h.split('\n').reduce((n, part) => n + (0, typography_1.linesNeeded)(part, KF_COL_W[i] - 0.18, KF_HEAD_FS_ALL, typography_1.FONT_TEXT), 0));
    // ---- the printable content area ----
    // The squirrel sits at y 13.74; the table must finish ABOVE it with a small clear margin, and its
    // border may never cross that line. Nothing here changes the page margins.
    const X = 1.43, TABLE_W = 24.68, TOP = 2.55;
    const MASCOT_TOP = 13.74, BOTTOM_MARGIN = 0.15; // ≈ 14 px at 96 dpi
    const LIMIT = MASCOT_TOP - BOTTOM_MARGIN; // 13.59" — the lowest the table may reach
    // PowerPoint treats a row height as a minimum and the exact wrap depends on the installed font, so
    // pages are filled to 92% of the measured space; the reserve absorbs that drift.
    const FILL = 0.92;
    const AVAIL = (LIMIT - TOP) * FILL;
    const PAD = 0.14; // consistent vertical padding in every row
    // Was a flat two lines at HEADER_FS_WIDE. Now it is whatever the tallest heading actually needs at
    // its own measured size, with the previous two-line height kept as the floor so the header row never
    // becomes shorter than it was.
    const headerH = Math.max((0, typography_1.lineHeight)(KF_HEAD_FS_ALL) * 2 + PAD, Math.max(...headLines) * (0, typography_1.lineHeight)(KF_HEAD_FS_ALL) + PAD);
    const bandH = (0, typography_1.lineHeight)(typography_1.KEYFOCUS_FS) + PAD;
    /**
     * Height a data row needs. Text wraps on WORD boundaries inside its own column (fitCell/linesNeeded
     * never break a word), and the row grows to whatever the tallest cell needs — the font is never
     * reduced, and the slide is never made taller.
     */
    const rowHeight = (r) => {
        const cells = [
            [sch.projectName, KF_COL_W[0]], [r.responsible, KF_COL_W[1]], [r.accountable, KF_COL_W[2]],
            [r.major, KF_COL_W[3]], [r.focus, KF_COL_W[4]],
        ];
        let h = (0, typography_1.lineHeight)(typography_1.KEYFOCUS_FS) + PAD;
        for (const [txt, w] of cells) {
            const need = (0, typography_1.linesNeeded)(txt || '', w - 0.18, typography_1.KEYFOCUS_FS, typography_1.FONT_TEXT) * (0, typography_1.lineHeight)(typography_1.KEYFOCUS_FS) + PAD;
            if (need > h)
                h = need;
        }
        return h;
    };
    // ---- pagination on MEASURED heights: a row is never split across two slides ----
    const pages = [];
    let cur = [], hs = [], used = headerH;
    const flush = () => { if (cur.length) {
        pages.push({ items: cur, heights: hs });
        cur = [];
        hs = [];
        used = headerH;
    } };
    for (let i = 0; i < flat.length; i++) {
        const r = flat[i];
        const h = 'band' in r ? bandH : rowHeight(r);
        // A band immediately followed by its first row must not be orphaned at the foot of a page.
        const pairH = ('band' in r && flat[i + 1] && !('band' in flat[i + 1]))
            ? h + rowHeight(flat[i + 1]) : h;
        if (used + pairH > AVAIL && cur.length)
            flush();
        cur.push(r);
        hs.push(h);
        used += h;
    }
    flush();
    const layoutPages = [];
    (pages.length ? pages : [{ items: [], heights: [] }]).forEach((pg, pi) => {
        const s = pptx.addSlide();
        const title = 'Key Focus Area' + (pi ? ' (cont.)' : '');
        sectionHeader(s, title, num, t);
        logos(s, assets);
        const trows = [header.map((h) => ({ text: h, options: { ...T(t).headWide, fontSize: KF_HEAD_FS_ALL } }))];
        for (const r of pg.items) {
            // The section bands use the SAME two colours as the Red Flag Summary — t.ragRed (#EE4B2B) and
            // t.ragAmber (#FFAA33) are the very values that slide fills its red and amber cells with, so the
            // two slides cannot drift apart. They were previously E0502A / F2A600, which only looked similar.
            if ('band' in r) {
                trows.push([{ text: r.band.toUpperCase(), options: { colspan: NCOL, fill: { color: r.band === 'Critical' ? theme_1.STATUS.critical(t) : theme_1.STATUS.atRisk(t) }, color: 'FFFFFF', bold: true } }]);
                continue;
            }
            trows.push([sch.projectName, r.responsible, r.accountable,
                r.major,
                // black like every other cell (the accent-orange text was dropped), but the EMPHASIS the
                // dashboard author applied is kept: the cell is written as its own runs, each with that run's
                // bold flag, so "… GFC drawings received on 11 Aug'26 …" stays bold inside an otherwise normal
                // sentence. Only a cell with at least one bold run needs runs; anything else stays the plain
                // string it always was, so no unformatted cell is touched.
                keyFocusCell(r),
                { text: r.delay, options: { align: 'center', fontFace: typography_1.FONT_NUM } }]);
        }
        const rowHs = [headerH, ...pg.heights];
        s.addTable(trows, {
            x: X, y: TOP, w: TABLE_W, colW: KF_COL_W, rowH: rowHs,
            // Key Focus content is 24pt (KEYFOCUS_FS); the header row keeps HEADER_FS_WIDE.
            fontFace: typography_1.FONT_TEXT, fontSize: typography_1.KEYFOCUS_FS, valign: 'middle',
            border: { type: 'solid', color: 'D9D9D9', pt: 0.5 }, autoPage: false,
        });
        const bottom = TOP + rowHs.reduce((a, b) => a + b, 0);
        layoutPages.push({
            slide: pi + 1, title, rows: pg.items.filter((r) => !('band' in r)).length,
            bandRows: pg.items.filter((r) => 'band' in r).length,
            headerH, rowHs: pg.heights, tableTop: TOP, tableBottom: bottom,
            marginBelow: MASCOT_TOP - bottom, splitRow: false,
        });
    });
    exports.keyFocusLayout = { pages: layoutPages, limit: LIMIT, mascotTop: MASCOT_TOP, colW: KF_COL_W,
        headFs: KF_HEAD_FS_ALL, widthNotes: kfWidthNotes };
}
// ---- Approvals ----
function approvals(pptx, rows, num, t, assets, 
/**
 * True when the ROW ORDER is already authoritative and must not be re-sorted — i.e. the list came
 * from the previous/reference deck's own Approvals slide (or a pinned per-project list), where the
 * sequence is a business decision the deck records. False when the list was derived from the
 * schedule's approval branch, which has no meaningful order of its own; that case is sorted by
 * Actual / Projected Finish, as before.
 */
preserveOrder = false, 
/**
 * BLOCK-SPLIT SECTIONS ONLY — see lookAheadSlides(). A project reported in blocks shows one
 * Approvals slide per block under one section number: 7A [South Block], 7B [North Block].
 */
opts = {}) {
    const s = pptx.addSlide();
    sectionHeader(s, opts.title || 'Approvals Status', opts.letterTotal !== undefined ? (0, blocks_1.pageLabel)(num, opts.letterFrom ?? 0, opts.letterTotal) : num, t);
    logos(s, assets);
    const header = ['Sl.No', 'Approvals', 'Drawings /\nDoc. Submission', 'Application', 'Fee\nPayment', 'Issued', 'Baseline\nFinish', 'Actual /\nProj. Finish'];
    const tick = (v) => ({ text: v ? '✔' : '✘', options: { align: 'center', bold: true, color: v ? '2E741C' : 'C00000', fontSize: typography_1.BODY_FS } });
    // Stage headings are set larger than the data beneath them (same family, same fill, same row
    // geometry) so the four workflow stages read at a glance.
    const trows = [header.map((h) => ({ text: h, options: { ...T(t).head, fontSize: typography_1.HEADER_FS_APPROVAL } }))];
    // ---- chronological order by Actual / Projected Finish ----
    // The approvals read top to bottom in the order they are expected to complete, earliest first.
    // Sorted HERE, in the renderer, on the very value the table prints, so the order is the same
    // whichever way the list was assembled (the reference deck's own list, the per-project master list
    // in approval_config.json, or the project's approval branch in its own Excel). Sorting inside one
    // of those readers would leave the other two paths unsorted.
    //   · `proj` is an ISO date (YYYY-MM-DD), so a plain string comparison IS chronological.
    //   · Rows with no Actual / Projected Finish sort AFTER every dated row — they are not dropped.
    //   · Ties, and the undated tail, keep the list's original order (a stable sort on `sl`).
    //   · Baseline Finish is never used for ordering, and neither is the approval name.
    // Nothing is filtered: `sorted.length === rows.length` always.
    //
    // EXCEPT when the order is already authoritative. If the list came from the previous deck's own
    // Approvals slide, that sequence IS the answer and re-sorting it is wrong: the business reads the
    // rows in the order it maintains them, and a row whose current data is missing sorts to the bottom
    // of a date sort, which is exactly how AAI NOC and IAF & GFTS NOC ended up at the foot of the table
    // looking as though they had been dropped. In that case the rows are rendered in the order given.
    const sorted = preserveOrder ? rows.slice() : rows.slice().sort((a, b) => {
        const A = a.proj || '', B = b.proj || '';
        if (A && B)
            return A === B ? a.sl - b.sl : (A < B ? -1 : 1);
        if (A)
            return -1; // a is dated, b is not → a first
        if (B)
            return 1; // b is dated, a is not → b first
        return a.sl - b.sl; // both undated → original order
    });
    // Sl.No is renumbered so the column still reads 1, 2, 3 … down the page. The approval's identity is
    // its name, not its original row number, and a scrambled counter beside a sorted table reads as a
    // fault. No row is added, removed or merged.
    for (const [i, r] of sorted.entries()) {
        trows.push([
            // Sl.No is a number → Meta-Normal, like every number in the deck.
            { text: String(i + 1), options: { align: 'center', fontFace: typography_1.FONT_NUM } }, r.approval,
            tick(r.drawings), tick(r.application), tick(r.fee), tick(r.issued),
            // Baseline Finish / Actual-Proj. Finish are dates → Meta-Normal, like every date in the deck.
            { text: (0, exports.fmtDate)(r.base), options: { align: 'center', fontFace: typography_1.FONT_NUM } },
            // Dates are Meta-Normal, regular weight (the bold was dropped); size, colour, alignment and
            // the completed-date highlight are unchanged.
            { text: (0, exports.fmtDate)(r.proj), options: { align: 'center', fontFace: typography_1.FONT_NUM, fill: r.issued ? { color: 'C6E0B4' } : undefined } },
        ]);
    }
    s.addTable(trows, { x: 1.43, y: 2.7, w: 24.68, colW: [1.6, 7.6, 3.2, 2.8, 2.5, 2.2, 2.4, 2.38], fontFace: typography_1.FONT_TEXT, fontSize: exports.TABLE_FS, valign: 'middle', rowH: 0.62, border: { type: 'solid', color: 'D9D9D9', pt: 0.5 } });
}
// ---- 2-Week Look Ahead ----
function lookAheadSlides(pptx, groups, num, t, assets, 
/**
 * BLOCK-SPLIT SECTIONS ONLY. A project reported in blocks shows this section once per block under
 * ONE section number, with the pages lettered continuously — South 8A–8E, then North 8F–8G — and
 * the block named in the heading. `letterFrom` is how many pages the section already has from the
 * earlier blocks and `letterTotal` how many it will have altogether; `title` replaces the heading.
 * Omit all three and the behaviour is exactly as before: plain number, "(cont.)" on later pages.
 *
 * `sameHeadingEveryPage` — DESIGN & BUILD ONLY. Every page of the section carries the SAME heading,
 * with no "(cont.)" on the later ones. A continuation page is already obvious from the fact that
 * the table keeps going, and the suffix makes the second and third pages read as a different, lesser
 * section. Launch & Design does not pass it, so its headings are untouched.
 */
opts = {}) {
    // Header labels break at word boundaries so each reads on one or two clean lines at 18pt.
    // The "Line" column was removed — the rows are read top to bottom, so the number added nothing.
    // The table itself is unchanged: no role columns. The three roles are a GROUPED HEADER above each
    // block of tasks instead (Head → Accountable → Responsible, indented, equal neighbours merged), so
    // a name is written once per group rather than repeated on every row.
    const header = ['Name', 'Status', 'Overall %\ncomplete', 'Baseline\nFinish', 'Actual/Proj\nStart', 'Actual/Proj\nFinish', 'Task.TE\nDepartments', 'Notes'];
    const NCOL = header.length;
    const icon = (a) => a.pct >= 100 ? { text: '✔', c: '2E741C' } : a.pct > 0 ? { text: '◐', c: 'BF8600' } : { text: '✘', c: 'C00000' };
    const blocks = [];
    let line = 1, lastHier = '';
    for (const g of groups) {
        const key = g.hier.map((h) => h.name).join(' > ');
        // Person row: Head, Accountable and Responsible on ONE line, separated by " | " —
        // "Om Prakash | Aparna S | Ramya H".
        //
        // They used to occupy one indented row each. The separator asked for implies a single line, and the
        // example shows one, so the lines are joined instead of stacked. The grouping itself is untouched:
        // `g.hier` is still built by roles.hierarchy(), which merges adjacent equal roles first — so a task
        // whose Head, Accountable and Responsible are the same person still shows that name once, and a
        // two-level group still shows two names. Only the presentation changed, and the block keeps one
        // header row instead of up to three, which the pagination measures for itself.
        const hier = [{
                band: g.hier.map((h) => h.name).join(' | '),
                fill: theme_1.LOOKAHEAD.person, color: theme_1.LOOKAHEAD.bandText, level: 0,
            }];
        blocks.push({
            hierKey: key,
            // The hierarchy is carried on every block, but only PRINTED when the page has not shown it yet —
            // so a person is named once per page, and a page can never begin with their tasks unattributed.
            hier,
            // Summary activity / subheading row — classified exactly as before (g.sub); only the fill changes.
            sub: g.sub ? { band: g.sub, fill: theme_1.LOOKAHEAD.summary, color: theme_1.LOOKAHEAD.bandText } : null,
            tasks: g.rows.map((r) => ({ r, a: r.activity, line: line++ })),
        });
        if (key !== lastHier)
            lastHier = key;
    }
    // ---- geometry: the table must fit EXACTLY inside the printable area ----
    const X = 1.43, TABLE_W = 24.68; // 1.43 + 24.68 = 26.11" < 26.67" slide width
    // PowerPoint treats a row height as a MINIMUM, so a wrapped cell can end a little taller than we
    // measured. A small reserve keeps the last row (and the table's outline) inside the printable area.
    // PowerPoint treats a row height as a MINIMUM, and the exact wrap depends on the installed font, so
    // a row can render a little taller than measured. Pages are filled to 90% of the usable height and a
    // reserve is kept below, which keeps the last row and the table's outline inside the printable area.
    // ---- vertical layout, this slide only ----
    // The header sits 0.17" higher than the deck default (1.12" → 0.95"), just clear of the wordmark,
    // which ends at 0.89". The table follows it up (2.55" → 2.25"), closing the gap under the bar, and now
    // runs down to 14.45" instead of 13.60" — 0.55" clear of the 15.00" slide edge.
    //
    // The mascot sits at x 0.61–1.35, entirely to the LEFT of the table's 1.43" origin, so the extra depth
    // costs nothing: the two never touch.
    //
    // Net usable height: 9.69" → 11.35", about 17% more rows per page. The fill factor is 0.94 rather than
    // 0.90 because there is now a real margin below to absorb the renderer growing a row beyond its
    // declared minimum — the effect that clipped a row two releases ago.
    const HDR_Y = 0.95;
    const TOP = 2.25, BOTTOM = 14.45, SAFETY = 0.15;
    // 0.98 of the remaining height. Two cushions are still held back deliberately: the 0.15" reserve above
    // and the 0.55" between BOTTOM and the 15.00" slide edge. Row heights in the file are MINIMUMS — the
    // renderer grows a row whose text needs more — and that is what clipped a row two releases ago, so the
    // margin is insurance, not waste. Effective limit: 11.81", i.e. a table may reach y ≈ 14.06".
    const AVAIL = (BOTTOM - TOP - SAFETY) * 0.98;
    // Fixed widths for dates / progress / status / delay; Notes and Name absorb the remainder,
    // with Notes given the larger share because it holds the longest prose.
    // The width freed by "Line" is redistributed: the fixed columns each gain a little, and Name and
    // Notes absorb the rest, so no empty space is left and the table still ends exactly at 26.11".
    // FIXED COLUMN WIDTHS — the same numbers on every page of the section, and they do NOT depend on what
    // any row happens to contain. A page whose Notes are short must still line up with the page before it.
    //
    // Each fixed column is set to what its own HEADING needs at 21pt plus its widest datum at 19pt, and
    // nothing more; the surplus goes to Notes. Against the previous layout the fixed columns give up 2.67"
    // in total — Status 1.45→1.15, Overall % 2.15→1.60, each date 2.25→1.92, Departments 2.70→1.88 — and
    // Notes takes all of it: 6.40" → 9.07", 42% wider. Name is deliberately UNCHANGED at 5.23" so activity
    // names did not pay for the Notes column.
    // W_DEPT holds the word "Departments" on ONE line at 21pt. At 1.88" it wrapped to a third line ("s"
    // alone), which made the header row taller than the height the pagination had reserved for it and
    // pushed the last row of the page off the bottom edge — the clipped row and the stray "s" were one
    // fault, not two.
    const W_STATUS = 1.15, W_PCT = 1.60, W_DATE = 1.92, W_DEPT = 2.15;
    const W_NAME = 5.23;
    const W_NOTES = TABLE_W - (W_NAME + W_STATUS + W_PCT + W_DATE * 3 + W_DEPT);
    const colW = [W_NAME, W_STATUS, W_PCT, W_DATE, W_DATE, W_DATE, W_DEPT, W_NOTES];
    // Row height is measured from the WIDEST-wrapping cell (Name / Notes / Department) so text is
    // never clipped: grow the row first, and only fall back to a smaller font at the readable floor.
    // Row height only. The type size is FIXED at LOOKAHEAD_FS for every task on every page: a long name
    // wraps and its row grows to hold the wrapped lines, and no cell is ever set smaller. Previously the
    // smallest fitted size found in the row was applied to the WHOLE row, so one long task name shrank its
    // own dates and department too, and identical tasks looked different sizes on different pages.
    const rowFor = (a) => {
        const cands = [[a.name, W_NAME], [a.note || '', W_NOTES], [a.dept || '', W_DEPT]];
        let h = (0, typography_1.lineHeight)(typography_1.LOOKAHEAD_FS) + 0.12;
        for (const [txt, w] of cands) {
            const lines = (0, typography_1.linesNeeded)(txt || '', w - 0.14, typography_1.LOOKAHEAD_FS, typography_1.FONT_TEXT);
            const need = lines * (0, typography_1.lineHeight)(typography_1.LOOKAHEAD_FS) + 0.12;
            if (need > h)
                h = need;
        }
        return { rowH: h, fontSize: typography_1.LOOKAHEAD_FS };
    };
    // Header row height follows the heading size: two lines plus padding, so a 24pt heading is never
    // clipped and the row grows only by as much as the text needs.
    const HDR_H = (0, typography_1.lineHeight)(typography_1.HEADER_FS_LOOKAHEAD) * 2 + 0.20;
    const BAND_H = (0, typography_1.lineHeight)(typography_1.LOOKAHEAD_FS) + 0.12;
    // ---- BLOCK-based pagination, on measured heights ----
    //
    // The rule, in order:
    //   1. A block is placed whole or not at all. If the hierarchy + subheading + all its tasks do not fit
    //      in what is left of the page, the page ends and the block starts the next one.
    //   2. Nothing is ever left stranded: a page cannot end with hierarchy rows, nor with a subheading
    //      whose tasks are overleaf, because those rows are never placed without at least their whole
    //      block behind them.
    //   3. A block taller than a whole page has to be split — there is nowhere else for it to go. It is
    //      split only after filling a page, and each continuation repeats the hierarchy and the
    //      subheading so the reader is never shown orphaned tasks.
    //   4. Row heights and fonts are untouched. Nothing is squeezed to make a block fit.
    const pages = [];
    let cur = [], used = HDR_H, shownHier = '';
    const taskH = (t) => rowFor(t.a).rowH;
    const newPage = () => { if (cur.length)
        pages.push(cur); cur = []; used = HDR_H; shownHier = ''; };
    for (const b of blocks) {
        const hierH = b.hierKey === shownHier ? 0 : b.hier.length * BAND_H;
        const subH = b.sub ? BAND_H : 0;
        const tasksH = b.tasks.reduce((n, t) => n + taskH(t), 0);
        // 1 — the whole block, or a new page. Nothing is placed that cannot be completed here.
        if (cur.length && used + hierH + subH + tasksH > AVAIL)
            newPage();
        // 2 — place the block. The loop runs more than once only for a block too tall for any page.
        let placed = 0;
        for (;;) {
            const mark = cur.length, usedMark = used;
            if (b.hierKey !== shownHier) {
                for (const h of b.hier) {
                    cur.push(h);
                    used += BAND_H;
                }
                shownHier = b.hierKey;
            }
            if (b.sub) {
                cur.push(b.sub);
                used += BAND_H;
            }
            let any = false;
            while (placed < b.tasks.length) {
                const h = taskH(b.tasks[placed]);
                if (used + h > AVAIL)
                    break;
                cur.push(b.tasks[placed]);
                used += h;
                placed += 1;
                any = true;
            }
            // 3 — the header went on but no task could follow it. Take the header back off and break the
            // page: this is the case that used to leave a name, or a subheading, alone at the foot.
            if (!any && placed < b.tasks.length) {
                cur.length = mark;
                used = usedMark;
                shownHier = mark ? shownHier : '';
                if (!cur.length)
                    break; // a single task taller than an empty page — cannot happen
                newPage();
                continue;
            }
            if (placed >= b.tasks.length)
                break; // block complete
            newPage(); // 4 — genuinely too tall: continue overleaf with its header
        }
    }
    if (cur.length)
        pages.push(cur);
    (pages.length ? pages : [[]]).forEach((pg, pi) => {
        const s = pptx.addSlide();
        // The heading is the deck's STANDARD section header — same bar fill, same badge fill, same geometry
        // and typography as Approvals Status and every other section. The lighter #C65911 override added in
        // v2.26.0 is gone: no `barFill` argument means HDR_BAR_FILL, which is what the rest of the deck uses.
        // Unblocked: the number alone, with "(cont.)" marking a continuation — unchanged.
        // Blocked: the section's own letter for this page, and the block named in the heading.
        const blocked = opts.letterTotal !== undefined;
        const base = opts.title || '2 Week Look Ahead Plan';
        sectionHeader(s, (blocked || opts.sameHeadingEveryPage) ? base : base + (pi ? ' (cont.)' : ''), blocked ? (0, blocks_1.pageLabel)(num, (opts.letterFrom ?? 0) + pi, opts.letterTotal) : num, t, undefined, HDR_Y);
        logos(s, assets);
        const rowHs = [HDR_H];
        // Column headings: 24pt Helvetica Neue, REGULAR weight, white on the unchanged header fill.
        // Name and Notes stay left-aligned to match the template; the rest are centred.
        const trows = [header.map((h, ci) => ({
                text: h,
                options: {
                    ...T(t).head, bold: false, fontSize: typography_1.HEADER_FS_LOOKAHEAD,
                    align: (ci === 0 || ci === NCOL - 1) ? 'left' : 'center', valign: 'middle',
                },
            }))];
        for (const r of pg) {
            if (r.band) {
                // A hierarchy line is indented one step per level; the WBS sub-band (no level) is not.
                const indent = r.level === undefined ? 0 : r.level * 26;
                trows.push([{
                        text: r.band,
                        options: {
                            // The whole section is one size: the person and summary bands are LOOKAHEAD_FS too, and
                            // their hierarchy is carried by the bold weight and their own fill, not by being larger.
                            colspan: NCOL, fill: { color: r.fill }, color: r.color, bold: true, fontSize: typography_1.LOOKAHEAD_FS,
                            fontFace: typography_1.FONT_TEXT, margin: [2, 4, 2, 6 + indent],
                        },
                    }]);
                rowHs.push(BAND_H);
                continue;
            }
            const a = r.a;
            const ic = icon(a);
            const { rowH, fontSize } = rowFor(a);
            rowHs.push(rowH);
            const num_ = (v, extra = {}) => ({ text: v, options: { align: 'center', fontFace: typography_1.FONT_NUM, fontSize, ...extra } });
            trows.push([
                { text: a.name, options: { fontFace: typography_1.FONT_TEXT, fontSize, align: 'left' } },
                { text: ic.text, options: { align: 'center', bold: true, color: ic.c, fontSize, fontFace: typography_1.FONT_TEXT } },
                num_(a.pct + '%'),
                num_((0, exports.fmtDate)(a.baselineFinish)),
                num_((0, exports.fmtDate)(a.start)),
                num_((0, exports.fmtDate)(a.finish)),
                // same black body colour as every other content column (font/size/alignment unchanged)
                { text: a.dept, options: { color: t.dark, fontFace: typography_1.FONT_TEXT, fontSize } },
                { text: a.note, options: { fontFace: typography_1.FONT_TEXT, fontSize, align: 'left' } },
            ]);
        }
        // ---- the outer boundary is part of the TABLE, not a rectangle drawn over it ----
        // It used to be a separate `rect` sized to the SUM OF THE DECLARED ROW HEIGHTS. Those heights are
        // MINIMUMS: PowerPoint and LibreOffice both grow a row whose text needs more space, so the rendered
        // table is taller than the number the rectangle was built from — and the last row ended up BELOW the
        // border, outside the box. Measuring the file could never show it, because in the file the two agreed
        // exactly.
        //
        // Making the outer edges cell borders removes the possibility: the boundary is drawn by the same
        // cells as the content, so it follows the rendered rows however tall they turn out to be, closes at
        // the real bottom-left and bottom-right corners of the final row, and needs no empty row to close it.
        const OUT = { type: 'solid', color: (0, exports.HDR_BAR_FILL)(t), pt: 1.25 };
        const IN = { type: 'solid', color: 'D9D9D9', pt: 0.5 };
        const lastRow = trows.length - 1;
        trows.forEach((row, ri) => {
            row.forEach((cell, ci) => {
                // A band row is one colspan cell, so it is both the first and the last column.
                const spans = (cell.options && cell.options.colspan) ? cell.options.colspan : 1;
                const first = ci === 0;
                const last = ci + spans >= NCOL;
                cell.options = cell.options || {};
                // [top, right, bottom, left] — outer edges in the frame colour, everything inside unchanged.
                cell.options.border = [
                    ri === 0 ? OUT : IN,
                    last ? OUT : IN,
                    ri === lastRow ? OUT : IN,
                    first ? OUT : IN,
                ];
            });
        });
        s.addTable(trows, {
            x: X, y: TOP, w: TABLE_W, colW, rowH: rowHs,
            fontFace: typography_1.FONT_TEXT, fontSize: typography_1.LOOKAHEAD_FS, valign: 'middle',
            border: IN, autoPage: false,
        });
    });
    return { pages: Math.max(1, pages.length) };
}
// ---- Slide built from an uploaded image (e.g. an updated Land Acquisition sheet) ----
function imageSlide(pptx, title, num, img, t, assets) {
    const s = pptx.addSlide();
    sectionHeader(s, title, num, t);
    logos(s, assets);
    const png = img[0] === 0x89 && img[1] === 0x50;
    const data = `data:image/${png ? 'png' : 'jpeg'};base64,${img.toString('base64')}`;
    // fit inside the content area, preserving aspect ratio
    s.addImage({ data, x: 1.43, y: 2.55, w: 24.68, h: 11.0, sizing: { type: 'contain', w: 24.68, h: 11.0 } });
}
// ---- Static placeholder (Land Acquisition / Master Plan) ----
function placeholder(pptx, title, num, msg, t, assets) {
    const s = pptx.addSlide();
    sectionHeader(s, title, num, t);
    logos(s, assets);
    s.addText(msg, { x: 1.43, y: 3.2, w: 24, h: 2, fontSize: typography_1.BODY_FS, italic: true, color: '595959', fontFace: typography_1.FONT_TEXT });
}
