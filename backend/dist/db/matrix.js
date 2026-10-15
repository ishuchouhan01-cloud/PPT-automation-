"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.monthLabel = exports.DB_LEGEND_GAP = exports.DB_MASCOT_BOTTOM = exports.DB_RIGHT = exports.DB_BOTTOM = exports.DB_TOP = exports.DB_W = exports.DB_X0 = void 0;
exports.dbMatrix = dbMatrix;
exports.dbMatrixLegendSize = dbMatrixLegendSize;
exports.dbMatrixLegend = dbMatrixLegend;
exports.dbTimeline = dbTimeline;
exports.dbPhaseLegendSize = dbPhaseLegendSize;
exports.dbPhaseLegend = dbPhaseLegend;
const typography_1 = require("../ppt/typography");
const theme_1 = require("./theme");
/** Content column geometry, shared with every other Design & Build slide. */
exports.DB_X0 = 1.43;
exports.DB_W = 24.68;
/** First usable y below the standard section header. */
exports.DB_TOP = 2.45;
/** Hard bottom edge of the printable area. */
exports.DB_BOTTOM = 14.45;
/**
 * The two edges every Design & Build legend is pinned to.
 *
 * RIGHT — the section title bar's right edge. The bar is drawn at DB_X0 with width DB_W, so this is
 * the same number, not a second measurement that could drift from it.
 * BOTTOM — the foot of the squirrel, which logos() places at y 13.74 with height 0.79. Pinning the
 * legend's bottom here is what makes it sit level with the mascot across the slide.
 *
 * Both legends therefore anchor bottom-right and lay themselves out leftwards and upwards from the
 * corner, instead of being positioned by a hand-tuned offset per call site.
 */
exports.DB_RIGHT = exports.DB_X0 + exports.DB_W;
exports.DB_MASCOT_BOTTOM = 13.74 + 0.79;
/** One legend row, shared by both legends so they align when they sit side by side. */
const LEG_ROW_H = 0.42;
/** Clear space between two legends on the same slide. */
exports.DB_LEGEND_GAP = 0.5;
const MONTH_ABBR = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
/** "2025-04" → "Apr-25", as the reference labels its month columns. */
const monthLabel = (key) => {
    const [y, m] = key.split('-');
    const mi = Math.max(0, Math.min(11, parseInt(m, 10) - 1));
    return `${MONTH_ABBR[mi]}-${y.slice(2)}`;
};
exports.monthLabel = monthLabel;
/** Break a long heading over two lines at the space nearest the middle. */
function wrap2(v, limit = 18) {
    if (v.length <= limit)
        return v;
    const mid = Math.floor(v.length / 2);
    for (let d = 0; d < v.length; d++) {
        if (v[mid - d] === ' ')
            return `${v.slice(0, mid - d)}\n${v.slice(mid - d + 1)}`;
        if (v[mid + d] === ' ')
            return `${v.slice(0, mid + d)}\n${v.slice(mid + d + 1)}`;
    }
    return v;
}
/**
 * Draw one band matrix.
 *
 * `mode`
 *   'color'   — the cell fill IS the status (Red Flag Summary DB)
 *   'symbol'  — a glyph in the cell IS the status (Tracker DB)
 *
 * Both the glyphs and their colours come from the sheet's legend, so the slide cannot drift from the
 * workbook. Where the sheet gives no legend entry for a status, the cell is left blank rather than
 * guessed — a wrong tick on a status grid is a reporting error, not a cosmetic one.
 *
 * `top` lets the caller stack this above a timeline. `showLegend` is off for the Tracker, whose
 * legend the caller places once beside the phase legend, exactly as the reference does.
 */
function dbMatrix(slide, m, mode, t, opts = {}) {
    const top = opts.top ?? exports.DB_TOP;
    const H_ROW = opts.rowH ?? 0.62;
    const n = Math.max(1, m.cols.length);
    // NO LEAD BLOCK. The rotated "Project Code" column and the project-name column are gone.
    //
    // Both carried the same single value on every row of a one-project deck — the code is already on
    // the cover and in the file name, and the project name is the deck's subject — so between them
    // they spent 5.2" of a 24.68" band restating what the slide is about instead of showing status.
    // "Sl No." was already omitted for the same reason. The whole width now belongs to the indicator
    // columns, which is what the reader actually has to read across the room.
    const indW = exports.DB_W / n;
    const wrapped = m.cols.map((c) => wrap2(c));
    const longest = wrapped.reduce((mx, v) => Math.max(mx, ...v.split('\n').map((l) => (0, typography_1.textWidth)(l, theme_1.DB_FS.matrixCol, theme_1.FONT_TEXT))), 0);
    const bandLines = Math.max(1, ...m.bands.map(([g, span]) => (0, typography_1.linesNeeded)(g, span * indW - 0.1, theme_1.DB_FS.band, theme_1.FONT_TEXT)));
    const H_BAND = Math.max(0.5, bandLines * (0, typography_1.lineHeight)(theme_1.DB_FS.band) + 0.12);
    // The sub-heading band is as TALL as the longest label is WIDE, because the labels are rotated.
    const H_SUB = Math.min(4.6, Math.max(2.0, longest + 0.4));
    const hdrH = H_BAND + H_SUB;
    const rowY = top + hdrH;
    // ---- band row: the sheet's own bands, order and spans ----
    let bx = exports.DB_X0;
    for (const [label, span] of m.bands) {
        const w = span * indW;
        slide.addText(label, {
            x: bx, y: top, w, h: H_BAND, align: 'center', valign: 'middle',
            fill: { color: 'D9D9D9' }, fontSize: theme_1.DB_FS.band, fontFace: theme_1.FONT_TEXT, color: t.dark,
            line: { color: 'FFFFFF', width: 0.5 }, margin: 1, wrap: true,
        });
        bx += w;
    }
    // ---- rotated column headings + the one project row ----
    const colXs = m.cols.map((_, i) => exports.DB_X0 + i * indW);
    const heads = (0, typography_1.headingLayout)(wrapped, colXs, top + H_BAND, indW, H_SUB, theme_1.DB_FS.matrixCol, theme_1.FONT_TEXT);
    const bySym = new Map();
    const byFill = new Map();
    for (const l of m.legend) {
        if (l.symbol && !bySym.has(l.rag))
            bySym.set(l.rag, { glyph: l.symbol, color: l.symbolColor || t.dark });
        if (l.color && !byFill.has(l.rag))
            byFill.set(l.rag, l.color);
    }
    m.cols.forEach((_, i) => {
        const cx = colXs[i];
        slide.addShape('rect', { x: cx, y: top + H_BAND, w: indW, h: H_SUB, fill: { color: 'FFFFFF' }, line: { color: theme_1.DB_GRID, width: 0.5 } });
        const hb = heads[i];
        slide.addText(hb.label, {
            x: hb.boxX, y: hb.boxY, w: hb.boxW, h: hb.boxH,
            align: 'center', valign: 'middle', rotate: 270, margin: 0, wrap: true,
            fontSize: theme_1.DB_FS.matrixCol, fontFace: theme_1.FONT_TEXT, color: t.dark,
        });
        const rag = m.cells[i];
        // A cell with no status may still carry a marker the sheet meant to be read — "-" for "does not
        // apply to this project". Print the sheet's own text in that case rather than leaving a blank,
        // which would read as "not yet assessed".
        const marker = m.raw[i] && /^(-|--|–|n\/?a)$/i.test(m.raw[i]) ? m.raw[i] : '';
        if (mode === 'symbol') {
            const s = bySym.get(rag);
            slide.addText(s ? s.glyph : marker, {
                x: cx, y: rowY, w: indW, h: H_ROW, align: 'center', valign: 'middle', margin: 0, wrap: false,
                fontSize: theme_1.DB_FS.matrixRow, color: s ? s.color : t.dark, fontFace: theme_1.FONT_TEXT,
                line: { color: theme_1.DB_GRID, width: 0.5 },
            });
        }
        else {
            const fill = byFill.get(rag);
            // "NA" is a word, not a colour, and must stay on one line in a narrow column.
            slide.addText(rag === 'na' ? (marker || 'NA') : marker, {
                x: cx, y: rowY, w: indW, h: H_ROW, align: 'center', valign: 'middle', margin: 0, wrap: false,
                fontSize: theme_1.DB_FS.matrixCol, fontFace: theme_1.FONT_TEXT, color: t.dark,
                fill: fill && rag !== 'na' ? { color: fill } : undefined,
                line: { color: theme_1.DB_GRID, width: 0.5 },
            });
        }
    });
    if (opts.showLegend !== false)
        dbMatrixLegend(slide, m, mode, t);
    return { bottom: rowY + H_ROW, height: hdrH + H_ROW };
}
/** Footprint of the matrix legend, so a caller can place a second legend beside it. */
function dbMatrixLegendSize(m, mode) {
    const rows = m.legend.filter((l) => (mode === 'symbol' ? !!l.symbol : !!l.color));
    return { w: 0.95 + 2.6, h: rows.length * LEG_ROW_H, rows: rows.length };
}
/**
 * The matrix legend, exactly as the sheet lists it: its labels, its order, its swatches or glyphs.
 *
 * ANCHORED TO THE BOTTOM-RIGHT CORNER by default — right edge on DB_RIGHT (the title bar's right
 * edge), bottom edge on DB_MASCOT_BOTTOM (the foot of the squirrel) — and it lays itself out
 * leftwards and upwards from there. `right` moves that corner for a slide carrying two legends;
 * `bottom` is shared, so the two sit level with each other and with the mascot.
 */
function dbMatrixLegend(slide, m, mode, t, anchor = {}) {
    const rows = m.legend.filter((l) => (mode === 'symbol' ? !!l.symbol : !!l.color));
    const size = dbMatrixLegendSize(m, mode);
    const x = (anchor.right ?? exports.DB_RIGHT) - size.w;
    const y = (anchor.bottom ?? exports.DB_MASCOT_BOTTOM) - size.h;
    if (!rows.length)
        return y;
    const swW = 0.95, labW = 2.6, rowH = LEG_ROW_H;
    rows.forEach((l, i) => {
        const ry = y + i * rowH;
        if (mode === 'color') {
            slide.addShape('rect', { x, y: ry, w: swW, h: rowH, fill: { color: l.color }, line: { color: theme_1.DB_GRID, width: 0.5 } });
        }
        else {
            slide.addText(l.symbol, {
                x, y: ry, w: swW, h: rowH, align: 'center', valign: 'middle', margin: 0, wrap: false,
                fontSize: theme_1.DB_FS.legend, fontFace: theme_1.FONT_TEXT, color: l.symbolColor || t.dark,
                line: { color: theme_1.DB_GRID, width: 0.5 },
            });
        }
        slide.addText(l.label, {
            x: x + swW, y: ry, w: labW, h: rowH, valign: 'middle', margin: 5,
            fontSize: theme_1.DB_FS.legend, fontFace: theme_1.FONT_TEXT, color: t.dark,
            line: { color: theme_1.DB_GRID, width: 0.5 },
        });
    });
    return y + rows.length * rowH;
}
/**
 * Draw the month grid and the project's two bars.
 *
 * THE BAR IS THE SHEET'S PAINT, NOT A DATE CALCULATION. Each month cell is filled with exactly the
 * colour the business painted in the dashboard, which is the only reading that reproduces the
 * reference slide — the painted extents do not always agree with the Base Start / Finish columns
 * beside them, and re-deriving bars from those dates would silently overrule a human decision.
 *
 * Only the painted window is shown. NH Homestay's design row is painted Apr-25 → Sep-26 out of a
 * 2023–2027 grid; rendering all sixty months would shrink the bar to a sliver.
 */
function dbTimeline(slide, tl, t, opts = {}) {
    const top = opts.top ?? exports.DB_TOP;
    const H_ROW = opts.rowH ?? 0.5;
    if (tl.from < 0)
        return { bottom: top, months: 0 };
    const months = tl.months.slice(tl.from, tl.to + 1);
    const proj = tl.projected.slice(tl.from, tl.to + 1);
    const base = tl.baseline.slice(tl.from, tl.to + 1);
    // The Project Code and project-name cells are gone, exactly as in the matrix above — on a
    // one-project deck they repeated the deck's own subject on every row. What remains is the row-label
    // column, which does carry information ("Actual/Projected" against "Baseline") and cannot go.
    const labelW = 2.35;
    const gridX = exports.DB_X0 + labelW;
    const colW = (exports.DB_W - labelW) / months.length;
    // ---- year band: one cell per run of months in the same year ----
    const H_YEAR = 0.5;
    let i = 0;
    while (i < months.length) {
        const year = months[i].slice(0, 4);
        let j = i;
        while (j < months.length && months[j].slice(0, 4) === year)
            j++;
        slide.addText(year, {
            x: gridX + i * colW, y: top, w: (j - i) * colW, h: H_YEAR,
            align: 'center', valign: 'middle', fill: { color: 'FFF2CC' },
            fontSize: theme_1.DB_FS.year, fontFace: theme_1.FONT_NUM, color: t.dark,
            line: { color: theme_1.DB_GRID, width: 0.5 }, margin: 0,
        });
        i = j;
    }
    // ---- rotated month labels ----
    const labels = months.map(exports.monthLabel);
    const longest = labels.reduce((mx, l) => Math.max(mx, (0, typography_1.textWidth)(l, theme_1.DB_FS.month, theme_1.FONT_NUM)), 0);
    const H_MON = Math.max(1.15, longest + 0.32);
    const monY = top + H_YEAR;
    const colXs = months.map((_, k) => gridX + k * colW);
    const heads = (0, typography_1.headingLayout)(labels, colXs, monY, colW, H_MON, theme_1.DB_FS.month, theme_1.FONT_NUM);
    months.forEach((_, k) => {
        slide.addShape('rect', { x: colXs[k], y: monY, w: colW, h: H_MON, fill: { color: 'FFFFFF' }, line: { color: theme_1.DB_GRID, width: 0.5 } });
        const hb = heads[k];
        slide.addText(hb.label, {
            x: hb.boxX, y: hb.boxY, w: hb.boxW, h: hb.boxH,
            align: 'center', valign: 'middle', rotate: 270, margin: 0, wrap: false,
            fontSize: theme_1.DB_FS.month, fontFace: theme_1.FONT_NUM, color: t.dark,
        });
    });
    // ---- lead cells, spanning both bar rows ----
    const barTop = monY + H_MON;
    const pairH = H_ROW * 2;
    slide.addShape('rect', { x: exports.DB_X0, y: top, w: labelW, h: barTop - top, fill: { color: 'FFFFFF' }, line: { color: theme_1.DB_GRID, width: 0.5 } });
    // ---- the two bar rows ----
    const drawRow = (label, fills, y) => {
        slide.addText(label, {
            x: exports.DB_X0, y, w: labelW, h: H_ROW, valign: 'middle', margin: 5,
            fontSize: theme_1.DB_FS.month, fontFace: theme_1.FONT_TEXT, color: t.dark, line: { color: theme_1.DB_GRID, width: 0.5 },
        });
        fills.forEach((hex, k) => {
            slide.addShape('rect', {
                x: colXs[k], y, w: colW, h: H_ROW,
                fill: hex ? { color: hex } : { color: 'FFFFFF' },
                line: { color: theme_1.DB_GRID, width: 0.5 },
            });
        });
    };
    drawRow('Actual/Projected', proj, barTop);
    drawRow('Baseline', base, barTop + H_ROW);
    if (opts.showLegend !== false)
        dbPhaseLegend(slide, tl, t);
    return { bottom: barTop + pairH, months: months.length };
}
/** Footprint of the phase legend, so a caller can place a second legend beside it. */
function dbPhaseLegendSize(tl) {
    return { w: 1.25 * 2 + 3.5, h: tl.legend.length ? (tl.legend.length + 1) * LEG_ROW_H : 0 };
}
/**
 * The phase legend, exactly as the sheet lists it: [Projected swatch][Baseline swatch][label], in the
 * sheet's own order, with the sheet's own colours. Column captions match the reference.
 *
 * ANCHORED TO THE BOTTOM-RIGHT CORNER by default, like the matrix legend: right edge on the title
 * bar's right edge, bottom edge on the foot of the squirrel. See DB_RIGHT / DB_MASCOT_BOTTOM.
 */
function dbPhaseLegend(slide, tl, t, anchor = {}) {
    const size = dbPhaseLegendSize(tl);
    const x = (anchor.right ?? exports.DB_RIGHT) - size.w;
    const y = (anchor.bottom ?? exports.DB_MASCOT_BOTTOM) - size.h;
    if (!tl.legend.length)
        return y;
    const swW = 1.25, labW = 3.5, rowH = LEG_ROW_H;
    slide.addText('Projected', {
        x, y, w: swW, h: rowH, align: 'center', valign: 'middle', margin: 0,
        fontSize: theme_1.DB_FS.legend, fontFace: theme_1.FONT_TEXT, color: t.dark, line: { color: theme_1.DB_GRID, width: 0.5 },
    });
    slide.addText('Baseline', {
        x: x + swW, y, w: swW, h: rowH, align: 'center', valign: 'middle', margin: 0,
        fontSize: theme_1.DB_FS.legend, fontFace: theme_1.FONT_TEXT, color: t.dark, line: { color: theme_1.DB_GRID, width: 0.5 },
    });
    slide.addShape('rect', { x: x + swW * 2, y, w: labW, h: rowH, fill: { color: 'FFFFFF' }, line: { color: theme_1.DB_GRID, width: 0.5 } });
    tl.legend.forEach((p, i) => {
        const ry = y + (i + 1) * rowH;
        slide.addShape('rect', { x, y: ry, w: swW, h: rowH, fill: { color: p.projected }, line: { color: theme_1.DB_GRID, width: 0.5 } });
        slide.addShape('rect', { x: x + swW, y: ry, w: swW, h: rowH, fill: { color: p.baseline }, line: { color: theme_1.DB_GRID, width: 0.5 } });
        slide.addText(p.label, {
            x: x + swW * 2, y: ry, w: labW, h: rowH, valign: 'middle', margin: 5,
            fontSize: theme_1.DB_FS.legend, fontFace: theme_1.FONT_TEXT, color: t.dark, line: { color: theme_1.DB_GRID, width: 0.5 },
        });
    });
    return y + (tl.legend.length + 1) * rowH;
}
