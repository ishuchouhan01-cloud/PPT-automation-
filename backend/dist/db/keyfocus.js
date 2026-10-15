"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.dbKeyFocus = dbKeyFocus;
const typography_1 = require("../ppt/typography");
const measure_1 = require("./measure");
const theme_1 = require("./theme");
const matrix_1 = require("./matrix");
/** The two name columns, wide enough for a full "Firstname Lastname" at the body size. */
const NAME_W = 4.6;
const COLS = [
    { label: 'Key Focus Area', w: matrix_1.DB_W - NAME_W * 2, align: 'left' },
    { label: 'Responsible', w: NAME_W, align: 'left' },
    { label: 'Accountable', w: NAME_W, align: 'left' },
];
/**
 * Render the Key Focus Area table.
 *
 * PAGINATION RULE: rows are dropped from the END of the lowest-priority band first, and the font is
 * never reduced to squeeze more in. A Key Focus slide exists to be read across a room; shrinking it
 * to fit one more line defeats the slide. Anything dropped is counted and reported.
 *
 * An empty band is never drawn — a "Critical" banner with nothing under it is an orphaned heading.
 */
function dbKeyFocus(slide, rows, bands, t) {
    const HDR_FS = theme_1.DB_FS.keyFocusHeader, BODY_FS = theme_1.DB_FS.keyFocus;
    const top = matrix_1.DB_TOP;
    // Header height from how the headings actually wrap in their own columns, not from break count.
    const hdrLines = Math.max(...COLS.map((c) => c.label.split('\n').reduce((n, seg) => n + (0, measure_1.dbLines)(seg, c.w - 0.14, HDR_FS, typography_1.FONT_TEXT), 0)));
    const H_HDR = Math.max(0.85, hdrLines * (0, typography_1.lineHeight)(HDR_FS) + 0.26);
    const H_BAND = (0, typography_1.lineHeight)(BODY_FS) + 0.20;
    // Measure every row against its widest wrapping column before deciding what fits.
    const focusW = COLS[0].w, respW = COLS[1].w, accW = COLS[2].w;
    const measure = (r) => (0, measure_1.dbRowHMax)([
        { text: r.focus, w: focusW }, { text: r.responsible, w: respW },
        { text: r.accountable, w: accW },
    ], BODY_FS, 0.56, 0.16);
    // Group in the dashboard's own band order, keeping each band's own row order.
    const grouped = bands
        .map((b) => ({ band: b, rows: rows.filter((r) => r.group.toLowerCase() === b.label.toLowerCase()) }))
        .filter((g) => g.rows.length > 0);
    const avail = matrix_1.DB_BOTTOM - (top + H_HDR) - 0.10;
    const dropped = [];
    let used = 0;
    const plan = [];
    for (const g of grouped) {
        if (used + H_BAND > avail) {
            dropped.push({ band: g.band.label, count: g.rows.length });
            continue;
        }
        plan.push({ band: g.band, h: H_BAND });
        used += H_BAND;
        let lost = 0;
        for (const r of g.rows) {
            const h = measure(r);
            if (used + h > avail) {
                lost++;
                continue;
            }
            plan.push({ row: r, h });
            used += h;
        }
        if (lost)
            dropped.push({ band: g.band.label, count: lost });
        // A banner with every row dropped would be an orphan; take it back out.
        if (plan[plan.length - 1]?.band === g.band) {
            plan.pop();
            used -= H_BAND;
        }
    }
    // ---- header ----
    let x = matrix_1.DB_X0;
    for (const c of COLS) {
        slide.addText(c.label, {
            x, y: top, w: c.w, h: H_HDR, align: 'center', valign: 'middle',
            fill: { color: 'FFFFFF' }, fontSize: HDR_FS, fontFace: typography_1.FONT_TEXT, color: '000000', bold: true,
            line: { color: theme_1.DB_GRID, width: 0.5 }, margin: 2, wrap: true,
        });
        x += c.w;
    }
    // ---- body ----
    let y = top + H_HDR;
    let rendered = 0;
    for (const item of plan) {
        if (item.band) {
            // Full-width banner in the dashboard's own colour, label left-aligned as the reference has it.
            slide.addText(item.band.label, {
                x: matrix_1.DB_X0, y, w: matrix_1.DB_W, h: item.h, valign: 'middle', margin: 6,
                fill: { color: item.band.color || t.ragAmber }, fontSize: BODY_FS, fontFace: typography_1.FONT_TEXT,
                color: '1A1A1A', line: { color: theme_1.DB_GRID, width: 0.5 },
            });
            y += item.h;
            continue;
        }
        const r = item.row;
        // The Key Focus cell keeps the emphasis the dashboard author gave it: where the source cell has
        // bold runs, the text is written as those runs so bold stays bold and normal stays normal inside
        // the one cell. A cell with no bold is written exactly as before, as a plain string.
        const focusVal = r.focusRuns && r.focusRuns.some((x) => x.bold)
            ? r.focusRuns.map((x) => ({ text: x.text, options: { bold: !!x.bold } }))
            : r.focus;
        const vals = [focusVal, r.responsible, r.accountable];
        let cx = matrix_1.DB_X0;
        COLS.forEach((c, ci) => {
            slide.addText(vals[ci], {
                x: cx, y, w: c.w, h: item.h, align: c.align, valign: 'middle', margin: 6, wrap: true,
                fill: { color: 'FFFFFF' }, fontSize: BODY_FS, fontFace: typography_1.FONT_TEXT, color: '1A1A1A',
                line: { color: theme_1.DB_GRID, width: 0.5 },
            });
            cx += c.w;
        });
        y += item.h;
        rendered++;
    }
    return { rendered, dropped, bands: grouped.map((g) => g.band.label), bottom: y };
}
