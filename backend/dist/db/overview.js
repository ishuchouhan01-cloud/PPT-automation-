"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.dbProjectOverview = dbProjectOverview;
exports.overviewRowHeight = overviewRowHeight;
exports.overviewPageCapacity = overviewPageCapacity;
exports.planOverviewPages = planOverviewPages;
exports.overviewScale = overviewScale;
const typography_1 = require("../ppt/typography");
const measure_1 = require("./measure");
const slides_1 = require("../ppt/slides");
const theme_1 = require("./theme");
const matrix_1 = require("./matrix");
/**
 * Column headings, verbatim from the reference (with its own line breaks), at the 73 : 27 table-to-Gantt
 * proportion.
 *
 * The widths sum to 18.02 in of the 24.68 in content band — 73.0% — leaving 6.66 in for the Gantt,
 * 27.0%. They were previously 17.74 in (71.9 : 28.1), which read the same but was not the stated
 * proportion; rescaled here so the split is exactly the specified one.
 */
const COLS = [
    { label: 'Name', w: 6.40, align: 'left' },
    { label: 'Duration\n[Elapsed day]', w: 1.75, align: 'center' },
    { label: 'Duration\n[Elapsed\nMonth]', w: 1.75, align: 'center' },
    { label: 'Overall percent\ncomplete', w: 2.05, align: 'center' },
    { label: 'Baseline Finish', w: 1.97, align: 'center' },
    { label: 'Actual /\nProjected Start', w: 2.01, align: 'center' },
    { label: 'Actual / Projected\nFinish', w: 2.09, align: 'center' },
];
const TABLE_W = COLS.reduce((n, c) => n + c.w, 0);
const GANTT_W = matrix_1.DB_W - TABLE_W;
/** Fill and text colour for a hierarchy level. */
function levelStyle(level) {
    if (level < theme_1.OVERVIEW_LEVEL_FILL.length) {
        return { fill: theme_1.OVERVIEW_LEVEL_FILL[level], color: theme_1.OVERVIEW_LEVEL_TEXT[level], bold: true };
    }
    return { fill: theme_1.OVERVIEW_LEAF_FILL, color: theme_1.OVERVIEW_LEAF_TEXT, bold: false };
}
const MS_DAY = 86400000;
const parse = (s) => {
    if (!s)
        return null;
    const v = Date.parse(s + 'T00:00:00Z');
    return isNaN(v) ? null : v;
};
/** Month key sequence covering [from, to] inclusive. */
function monthSpan(fromMs, toMs) {
    const out = [];
    const d = new Date(fromMs);
    d.setUTCDate(1);
    const end = new Date(toMs);
    while (d.getUTCFullYear() * 12 + d.getUTCMonth() <= end.getUTCFullYear() * 12 + end.getUTCMonth()) {
        out.push(`${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}`);
        d.setUTCMonth(d.getUTCMonth() + 1);
    }
    return out;
}
/**
 * Render the Project Overview table and Gantt.
 *
 * ROW HEIGHTS ARE COMPUTED FROM WRAPPED TEXT, NOT ASSUMED. Declared table heights are minimums that
 * a renderer is free to grow, so every row is drawn as explicit shapes at measured heights. That is
 * why the Gantt bars stay aligned with their rows: both are positioned from the same measured
 * geometry rather than from a table the renderer might re-flow.
 */
function dbProjectOverview(slide, rows, t, 
/**
 * OPTIONAL, and additive: omitting it reproduces the previous behaviour exactly.
 *
 * `scale` pins the Gantt's date window instead of deriving it from THESE rows. Used when a view
 * spans several pages: every page must share one calendar, or the same activity would sit at a
 * different x on page 2 than on page 1 and the pages could not be read side by side.
 */
opts = {}) {
    const barless = [];
    const HDR_FS = theme_1.DB_FS.overviewHeader, BODY_FS = theme_1.DB_FS.overviewBody;
    // ---- header height from the tallest heading as it ACTUALLY wraps in its own column ----
    // Counting "\n" is not enough: a heading can wrap further inside a narrow column, and sizing the
    // row from the break count clips the last line.
    const hdrLines = Math.max(...COLS.map((c) => c.label.split('\n').reduce((n, seg) => n + (0, measure_1.dbLines)(seg, c.w - 0.14, HDR_FS, typography_1.FONT_TEXT), 0)));
    const H_HDR = Math.max(0.95, hdrLines * (0, typography_1.lineHeight)(HDR_FS) + 0.24);
    const top = matrix_1.DB_TOP;
    // ---- row heights: the Name column is the only one that wraps ----
    const nameW = COLS[0].w;
    const heights = rows.map((r) => {
        const indent = 0.18 * Math.min(r.level, 4);
        // Grow the row to fit the wrapped name; the font is never reduced, because a Project Overview
        // with two type sizes in one column reads as a mistake.
        return Math.max(0.62, (0, measure_1.dbRowH)(r.name, nameW - indent, BODY_FS, typography_1.FONT_TEXT, 0.20));
    });
    const total = heights.reduce((a, b) => a + b, 0);
    const avail = matrix_1.DB_BOTTOM - (top + H_HDR) - 0.12;
    // If the measured block is taller than the page, scale every row by one factor so the grid stays
    // regular. Row text is never clipped because the font is not reduced — the table simply gets no
    // more rows than overviewRowsFor() already limited it to.
    const k = total > avail ? avail / total : 1;
    const rowH = heights.map((h) => h * k);
    // ---- Gantt scale ----
    let lo = opts.scale ? opts.scale.lo : Infinity;
    let hi = opts.scale ? opts.scale.hi : -Infinity;
    if (!opts.scale) {
        for (const r of rows) {
            for (const v of [parse(r.start), parse(r.finish), parse(r.baselineStart), parse(r.baselineFinish)]) {
                if (v == null)
                    continue;
                lo = Math.min(lo, v);
                hi = Math.max(hi, v);
            }
        }
    }
    const haveScale = isFinite(lo) && isFinite(hi) && hi > lo;
    const gx = matrix_1.DB_X0 + TABLE_W;
    const months = haveScale ? monthSpan(lo, hi) : [];
    const scaleFrom = haveScale ? Date.parse(months[0] + '-01T00:00:00Z') : 0;
    const lastMonth = haveScale ? months[months.length - 1] : '';
    const scaleTo = haveScale
        ? Date.parse(`${lastMonth}-01T00:00:00Z`) + 31 * MS_DAY
        : 1;
    const px = (ms) => gx + ((ms - scaleFrom) / (scaleTo - scaleFrom)) * GANTT_W;
    // ---- header row ----
    let x = matrix_1.DB_X0;
    for (const c of COLS) {
        slide.addText(c.label, {
            x, y: top, w: c.w, h: H_HDR, align: 'center', valign: 'middle',
            fill: { color: 'FFFFFF' }, fontSize: HDR_FS, fontFace: typography_1.FONT_TEXT, color: '000000',
            line: { color: '000000', width: 0.75 }, margin: 2, wrap: true,
        });
        x += c.w;
    }
    // Gantt header: year band over month ticks, matching the reference's compact scale.
    slide.addShape('rect', { x: gx, y: top, w: GANTT_W, h: H_HDR, fill: { color: 'FFFFFF' }, line: { color: '000000', width: 0.75 } });
    if (haveScale) {
        const mw = GANTT_W / months.length;
        let i = 0;
        while (i < months.length) {
            const year = months[i].slice(0, 4);
            let j = i;
            while (j < months.length && months[j].slice(0, 4) === year)
                j++;
            slide.addText(year, {
                x: gx + i * mw, y: top + H_HDR * 0.34, w: (j - i) * mw, h: H_HDR * 0.32,
                align: 'center', valign: 'middle', margin: 0, wrap: false,
                fontSize: Math.max(14, theme_1.DB_FS.month - 2), fontFace: typography_1.FONT_NUM, color: t.dark,
            });
            // Year separator, as the reference draws it.
            if (i > 0) {
                slide.addShape('line', { x: gx + i * mw, y: top, w: 0, h: H_HDR, line: { color: 'C00000', width: 0.75 } });
            }
            i = j;
        }
        // Month initials only when there is room for them to be legible.
        if (mw >= 0.20) {
            months.forEach((mk, n) => {
                slide.addText((0, matrix_1.monthLabel)(mk).charAt(0), {
                    x: gx + n * mw, y: top + H_HDR * 0.68, w: mw, h: H_HDR * 0.3,
                    align: 'center', valign: 'middle', margin: 0, wrap: false,
                    fontSize: 13, fontFace: typography_1.FONT_NUM, color: '595959',
                });
            });
        }
    }
    else {
        slide.addText('No dated rows — no timeline can be drawn from this export', {
            x: gx, y: top, w: GANTT_W, h: H_HDR, align: 'center', valign: 'middle',
            fontSize: theme_1.DB_FS.legend, fontFace: typography_1.FONT_TEXT, color: '7F7F7F', margin: 2, wrap: true,
        });
    }
    // ---- body ----
    let y = top + H_HDR;
    rows.forEach((r, ri) => {
        const h = rowH[ri];
        const st = levelStyle(r.level);
        const indent = 0.18 * Math.min(r.level, 4);
        const vals = [
            r.name,
            r.elapsedDays == null ? '' : `${r.elapsedDays}ed`,
            r.elapsedMonths == null ? '' : `${r.elapsedMonths}em`,
            r.pct == null ? '' : `${r.pct}%`,
            (0, slides_1.fmtDate)(r.baselineFinish),
            (0, slides_1.fmtDate)(r.start),
            (0, slides_1.fmtDate)(r.finish),
        ];
        let cx = matrix_1.DB_X0;
        COLS.forEach((c, ci) => {
            const isName = ci === 0;
            slide.addText(vals[ci], {
                x: cx + (isName ? indent : 0), y, w: c.w - (isName ? indent : 0), h,
                align: c.align, valign: 'middle', margin: isName ? 6 : 3, wrap: true,
                fill: { color: st.fill }, fontSize: BODY_FS,
                // Numbers and dates in Meta-Normal, words in Helvetica Neue — the EDOS type rule.
                fontFace: isName ? typography_1.FONT_TEXT : typography_1.FONT_NUM,
                color: st.color, bold: st.bold,
                line: { color: theme_1.DB_GRID, width: 0.5 },
            });
            cx += c.w;
        });
        // Gantt cell background + the three bars.
        slide.addShape('rect', { x: gx, y, w: GANTT_W, h, fill: { color: 'FFFFFF' }, line: { color: theme_1.DB_GRID, width: 0.5 } });
        if (haveScale) {
            const s = parse(r.start), f = parse(r.finish), bf = parse(r.baselineFinish);
            const barH = Math.min(0.20, (h - 0.10) / 3);
            const gap = 0.03;
            const y0 = y + (h - (barH * 3 + gap * 2)) / 2;
            // Actual progress: from start to the completed fraction of the span.
            if (s != null && f != null && f > s) {
                const pct = Math.max(0, Math.min(100, r.pct ?? 0));
                if (pct > 0) {
                    const w = ((f - s) * (pct / 100)) / (scaleTo - scaleFrom) * GANTT_W;
                    if (w > 0.005) {
                        slide.addShape('rect', {
                            x: px(s), y: y0, w: Math.max(0.02, w), h: barH,
                            fill: { color: theme_1.OVERVIEW_BAR.actual }, line: { color: theme_1.OVERVIEW_BAR.outline, width: 0.5 },
                        });
                    }
                }
                // Projection: the whole actual/projected span.
                slide.addShape('rect', {
                    x: px(s), y: y0 + barH + gap, w: Math.max(0.02, px(f) - px(s)), h: barH,
                    fill: { color: theme_1.OVERVIEW_BAR.projection }, line: { color: theme_1.OVERVIEW_BAR.outline, width: 0.5 },
                });
            }
            else {
                barless.push({
                    name: r.name,
                    reason: !s && !f ? 'the export gives neither an actual/projected start nor finish'
                        : !s ? 'the export gives no actual/projected start'
                            : !f ? 'the export gives no actual/projected finish'
                                : 'the actual/projected finish is not after the start',
                });
            }
            // Baseline: the PLAN, drawn from baseline start to baseline finish.
            //
            // Not from the live start. Drawing it from the actual/projected start to the baseline finish
            // makes a slipped activity look as though it had always been planned to take that long —
            // "Structural Slab 1" appeared as a six-month baseline when the plan gave it six weeks. When
            // the export has no baseline start the bar is drawn as a short marker at the baseline finish
            // instead, which states the one date that is known without implying a duration.
            const bs = parse(r.baselineStart);
            if (bf != null) {
                const x1 = bs != null ? px(bs) : px(bf) - 0.08;
                slide.addShape('rect', {
                    x: Math.min(x1, px(bf)), y: y0 + (barH + gap) * 2,
                    w: Math.max(0.02, Math.abs(px(bf) - x1)), h: barH,
                    fill: { color: theme_1.OVERVIEW_BAR.baseline }, line: { color: theme_1.OVERVIEW_BAR.outline, width: 0.5 },
                });
            }
        }
        y += h;
    });
    // ---- legend, bottom-right, in the reference's order and wording ----
    const LEG = [
        ['Baseline', theme_1.OVERVIEW_BAR.baseline],
        ['Projection', theme_1.OVERVIEW_BAR.projection],
        ['Actual Progress', theme_1.OVERVIEW_BAR.actual],
    ];
    const ly = Math.min(y + 0.35, matrix_1.DB_BOTTOM + 0.30);
    let lx = matrix_1.DB_X0 + matrix_1.DB_W - 8.2;
    for (const [label, color] of LEG) {
        slide.addShape('rect', { x: lx, y: ly + 0.06, w: 1.00, h: 0.26, fill: { color }, line: { color: theme_1.OVERVIEW_BAR.outline, width: 0.5 } });
        slide.addText(label, {
            x: lx + 1.10, y: ly, w: 1.65, h: 0.4, valign: 'middle', margin: 0, wrap: false,
            fontSize: theme_1.DB_FS.legend, fontFace: typography_1.FONT_TEXT, color: t.dark,
        });
        lx += 2.80;
    }
    return {
        rendered: rows.length, barless,
        scale: haveScale ? { first: months[0], last: lastMonth } : null,
        bottom: y,
    };
}
// ---------------------------------------------------------------------------------------------
// Pagination
// ---------------------------------------------------------------------------------------------
/**
 * Height one Project Overview row needs, and the height available for rows on a page.
 *
 * Exported because the page planner and the renderer must agree to the inch. They measure the same
 * way — same font, same column width, same indent — so a page that the planner says holds nine rows
 * holds exactly nine when drawn.
 */
function overviewRowHeight(r) {
    const indent = 0.18 * Math.min(r.level, 4);
    return Math.max(0.62, (0, measure_1.dbRowH)(r.name, COLS[0].w - indent, theme_1.DB_FS.overviewBody, typography_1.FONT_TEXT, 0.20));
}
function overviewPageCapacity() {
    const HDR_FS = theme_1.DB_FS.overviewHeader;
    const hdrLines = Math.max(...COLS.map((c) => c.label.split('\n').reduce((n, seg) => n + (0, measure_1.dbLines)(seg, c.w - 0.14, HDR_FS, typography_1.FONT_TEXT), 0)));
    const H_HDR = Math.max(0.95, hdrLines * (0, typography_1.lineHeight)(HDR_FS) + 0.24);
    return matrix_1.DB_BOTTOM - (matrix_1.DB_TOP + H_HDR) - 0.12;
}
/**
 * Split rows into pages that each fit the printable area at full size.
 *
 * NOTHING IS SCALED DOWN TO MAKE A SET FIT. The single-page renderer compresses rows by one factor
 * if it is handed more than a page's worth, which is right for a capped set but wrong here: a Level 4
 * view can be four times a page, and squeezing it would produce unreadable slides. So the set is
 * broken into pages at the page's own capacity and every page is drawn at the normal size, with the
 * same header, the same column widths and the same calendar.
 */
function planOverviewPages(rows) {
    const avail = overviewPageCapacity();
    const pages = [];
    let cur = [], used = 0;
    for (const r of rows) {
        const h = overviewRowHeight(r);
        if (cur.length && used + h > avail) {
            pages.push(cur);
            cur = [];
            used = 0;
        }
        cur.push(r);
        used += h;
    }
    if (cur.length)
        pages.push(cur);
    return pages.length ? pages : [[]];
}
/** The date window covering every row, so all pages of a view share one calendar. */
function overviewScale(rows) {
    let lo = Infinity, hi = -Infinity;
    for (const r of rows) {
        for (const v of [parse(r.start), parse(r.finish), parse(r.baselineStart), parse(r.baselineFinish)]) {
            if (v == null)
                continue;
            lo = Math.min(lo, v);
            hi = Math.max(hi, v);
        }
    }
    return isFinite(lo) && isFinite(hi) && hi > lo ? { lo, hi } : undefined;
}
