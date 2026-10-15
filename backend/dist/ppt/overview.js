"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.overviewGeometry = overviewGeometry;
exports.projectOverview = projectOverview;
const derive_1 = require("./derive");
const blocks_1 = require("./blocks");
const slides_1 = require("./slides");
const typography_1 = require("./typography");
// ---- palette: the three-level row hierarchy, from the approved reference ----
// The level is derived from the table's OWN hierarchy — the project row, then `major` phase rows, then
// everything else — so no activity name appears anywhere in this file.
//
//   LEVEL 1  project / top-level      #D99370 ground, WHITE text
//   LEVEL 2  phase / second-level     #BFD8D8 ground, #D99370 text
//   LEVEL 3  task / activity          #9F9F5F ground, WHITE text
//
// The header row is deliberately NOT part of this scheme: white ground, black text, so the hierarchy
// starts below it and the headings stay legible against all three bands.
const LVL = [
    { fill: 'D99370', text: 'FFFFFF' }, // 1 — project
    { fill: 'BFD8D8', text: 'D99370' }, // 2 — phase
    { fill: '9F9F5F', text: 'FFFFFF' }, // 3 — task
];
const HDR_FILL = 'FFFFFF'; // table header ground
const HDR_TXT = '000000'; // table header text
const GRID = 'D9D9D9'; // light grid lines
// ---- Gantt palette, sampled from the approved reference image (not guessed) ----
const G_BASELINE = 'FFD000'; // Baseline
const G_ACTUAL = '5B9BD5'; // Actual Progress — a lighter blue, so it separates from the projection
const G_PROJ = '008A3E'; // Projection — a darker green, so it reads clearly under the blue
const G_TODAY = 'C8281E'; // the TODAY marker
const G_AXIS = 'BFBFBF'; // month ticks / chart frame
// The Tracker & Schedule calendar's year-band ground. Shared so the two calendars match.
const CAL_HEAD = 'D9D9D9';
const G_BAND = 'E9E4C9'; // the pale band behind a section-heading row, as in the reference
const MO3 = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
/**
 * THE horizontal geometry of every Project Overview page — table band, column boundaries and the
 * Gantt origin. Exported, and the ONE place these numbers exist: the renderer draws from it and
 * the validators measure against it, so the two can never drift apart (they did: the refinement
 * report recomputed the old 73:27 boundary of its own accord, and the moment the block moved it
 * started judging the Gantt's year band by the table's rules).
 *
 * Nothing here depends on a page's contents, so 4A, 4B, 4C … are identical BY CONSTRUCTION — same
 * left edge, same column x positions, same Gantt start, same total width. A continuation page
 * cannot come out horizontally different from the first one.
 *
 * LEFT EDGE — 1.43". The same x as the section header bar ("4A PROJECT OVERVIEW") and the badge
 * above it. It used to be 0.50", which left the table hanging 0.93" out to the left of its own
 * heading; that is what read as misaligned. The right edge is 26.11", the header bar's right edge,
 * so header, table and Gantt now share exactly one band.
 *
 * NAME COLUMN — the fixed columns are sized from their own headings and data; Name takes the rest
 * and then gives back a FIXED trim. A constant, not a measurement of the longest activity name, so
 * every project's Project Overview has the same column boundaries whatever its activities are
 * called. The 73:27 split is taken on the band it was calibrated against (0.50"→26.11"), NOT the
 * new narrower band: moving the block right to meet the header bar must not take a second, silent
 * bite out of Name on top of the declared trim. Name changes by exactly NAME_TRIM and no more; the
 * width that frees up goes to the Gantt.
 */
function overviewGeometry() {
    const X0 = 1.43, XEND = 26.11, CONTENT_W = XEND - X0;
    const SPLIT_BAND = 26.11 - 0.5; // the band the 73:27 proportions were set against
    const NAME_TRIM = 1.20; // fixed, every project, every page
    // Each fixed column is sized from the wider of (a) its own heading and (b) its data, so nothing is
    // clipped and no width is wasted — that is what makes seven columns and a Gantt fit side by side.
    // The headings are the reference's own, in full, nothing abbreviated; they are long, so each column
    // is sized by its longest WORD (the narrowest width at which it wraps without cutting a word). The
    // width budget stays at 25pt even though the headings print at 24pt: sizing the geometry from the
    // smaller type would have narrowed every date column, a dimension change nobody asked for.
    const HDR_METRIC_FS = 25;
    const hw = (v) => Math.max(...v.split(/[\s\n]+/).map((w) => (0, typography_1.textWidth)(w, HDR_METRIC_FS, typography_1.FONT_TEXT)));
    const dw = (v) => (0, typography_1.textWidth)(v, typography_1.OVERVIEW_FS, typography_1.FONT_NUM);
    const W_PCT = Math.max(hw('Overall percent complete'), dw('100%')) + 0.26;
    const W_DAY = Math.max(hw('Duration [Elapsed day]'), dw('1351ed')) + 0.26;
    const W_MON = Math.max(hw('Duration [Elapsed Month]'), dw('44.42em')) + 0.26;
    const W_DATE = Math.max(hw('Actual / Projected Finish'), dw('30 Mar 26')) + 0.26;
    const cols = [
        { label: 'Name', w: 0, al: 'left', num: false },
        { label: 'Duration\n[Elapsed day]', w: W_DAY, al: 'center', num: true },
        { label: 'Duration\n[Elapsed Month]', w: W_MON, al: 'center', num: true },
        { label: 'Overall percent\ncomplete', w: W_PCT, al: 'center', num: true },
        { label: 'Baseline Finish', w: W_DATE, al: 'center', num: true },
        { label: 'Actual /\nProjected Start', w: W_DATE, al: 'center', num: true },
        { label: 'Actual / Projected\nFinish', w: W_DATE, al: 'center', num: true },
    ];
    const fixed = cols.slice(1).reduce((a, c) => a + c.w, 0);
    cols[0].w = SPLIT_BAND * 0.73 - fixed - NAME_TRIM;
    const TABLE_W = fixed + cols[0].w;
    const GX = X0 + TABLE_W, GW = CONTENT_W - TABLE_W;
    let cx = X0;
    const xs = [];
    for (const c of cols) {
        xs.push(cx);
        cx += c.w;
    }
    return { X0, XEND, CONTENT_W, TABLE_W, GX, GW, cols, xs, HDR_METRIC_FS, NAME_TRIM };
}
function header(slide, t, page, title = 'Project Overview') {
    // Exactly the same header as every other generated slide: a fixed-width number box and a title bar
    // that starts after it, each its own text box. A two-character label ("4A") is set slightly smaller
    // so it always sits inside its box — the two can never touch, overlap or merge.
    (0, slides_1.sectionHeader)(slide, title, page, t);
}
function projectOverview(pptx, sch, t, assets, reportDate, 
/** The section number this slide carries. Passed in so it follows the deck's real section order
 *  rather than being fixed at 4 — a deck that omits an earlier section renumbers this one. */
sectionNo = 4, 
/**
 * OPTIONAL, additive, and used only by the Design & Build mode. Omitting it reproduces the Launch &
 * Design behaviour exactly — the rows are still selected by overviewRows() and the title is still
 * "Project Overview".
 *
 *   `rows`  — render THESE rows instead of selecting them from the schedule. Design & Build chooses
 *             its rows from its own WBS hierarchy and view level, so it needs to hand them over
 *             rather than have them re-derived. Everything downstream of the selection — the seven
 *             columns, the 73:27 split, the palette, the Gantt, the Today line, the legend, the
 *             pagination and the continuation pages — is this function's, unchanged. That is the
 *             point: one renderer, two callers, so the two cannot drift apart.
 *   `title` — the heading, for "Project Overview - Construction".
 */
/**
 *   `letterFrom` / `letterTotal` — page lettering across a section that several BLOCKS share.
 *      A section reported once per block is ONE section number whose pages letter continuously:
 *      South's three pages are 4A 4B 4C and North's carry on 4D 4E 4F. The caller passes how many
 *      pages the section already has (`letterFrom`) and how many it will have in total
 *      (`letterTotal`), because this call only knows about its own block. Omitting both reproduces
 *      the single-block behaviour exactly: one page → "4", several → 4A, 4B, …
 */
/**
 *   `nestingBars` — DESIGN & BUILD ONLY. Carry each summary row's own colour down the left edge of
 *      everything nested under it, the way Asta Powerproject draws a summary's bar over its tasks.
 *      Launch & Design does not pass it and is drawn exactly as before.
 */
opts = {}) {
    const injected = Array.isArray(opts.rows);
    const selected = injected ? opts.rows : (0, derive_1.overviewRows)(sch);
    // ---- the first displayed row: taken FROM the Excel hierarchy, not from the project name ----
    // This row heads the table, and it must be the real row that the displayed sections sit under.
    //
    // It used to be the export's ROOT WBS row, which is the bare project name. On most exports the
    // sections hang directly off the root, so the two are the same row and nothing changes. But some
    // exports carry a programme row between them — One Tiny Flower is
    //   depth 0  "One Tiny Flower"                            (the project)
    //   depth 1  "One Tiny Flower - Acquisition and Launch"    (the programme the sections belong to)
    //   depth 2  "Milestones", "Project Acquisition", …        (the displayed sections)
    // and the depth-1 row is the one the reader needs, because its figures (55% complete, 960 days,
    // finish 07-Dec-26) describe the work actually shown beneath it, whereas the root's (38%, 1106 days,
    // finish 30-Jun-27) describe a wider scope than the table lists. The depth-1 row was also being
    // dropped by the selector, which treats a row containing the whole project name as a container.
    //
    // So the row is now found by walking UP from the displayed sections: take the shallowest selected
    // row and use ITS PARENT. That is data-driven and needs no per-project rule — where the sections
    // already hang off the root, the parent IS the root and the previous behaviour is reproduced
    // exactly. The name is read from the Excel row; nothing is concatenated or constructed.
    const byName = new Map(Object.values(sch.wbs).map((w) => [w.name.trim(), w]));
    const depthOf = (w) => {
        let d = 0, c = w;
        while (c && c.parent) {
            c = sch.wbs[c.parent];
            d++;
        }
        return d;
    };
    let root;
    if (!injected) {
        // The shallowest section actually being displayed, in the export's own order.
        const topSection = selected
            .map((r) => byName.get(r.name.trim()))
            .filter((w) => !!w)
            .sort((a, b) => (depthOf(a) - depthOf(b)) || (a.seq - b.seq))[0];
        const parent = topSection && topSection.parent ? sch.wbs[topSection.parent] : undefined;
        root = parent || Object.values(sch.wbs)
            .filter((w) => !w.parent)
            .sort((a, b) => a.seq - b.seq)[0];
    }
    const projectRow = root ? {
        name: root.name, indent: 0, summary: true, major: true,
        pct: Math.round(root.pct ?? 0), duration: root.duration ?? 0,
        baseStart: root.baselineStart ?? null, baseFinish: root.baselineFinish ?? null,
        start: root.start ?? null, finish: root.finish ?? null,
        critical: false, milestone: false, delayDays: 0,
        done: (root.pct ?? 0) >= 100 && !!root.finish,
        category: 'the project itself',
    } : null;
    // ---- the Milestone section is not shown ----
    // Removed on request: the row itself and anything indented under it, so no gap is left where it was.
    // Everything else — data, order, hierarchy, calculations — is untouched.
    const isMilestone = (r) => r.major && /^milestones?$/i.test(r.name.trim());
    const kept = [];
    let skipping = false;
    for (const r of selected) {
        if (!injected && isMilestone(r)) {
            skipping = true;
            continue;
        } // drop the heading …
        if (!injected && skipping && !r.major)
            continue; // … and its own activities
        skipping = false;
        kept.push(r);
    }
    // Never twice: a selection that already contains the root drops the synthetic copy.
    const rows = projectRow && !kept.some((r) => r.name === projectRow.name)
        ? [projectRow, ...kept] : kept;
    /**
     * Which of the three hierarchy levels a row belongs to — read from the table's own structure, never
     * from a name. Level 1 is the project row (the root WBS node, which is `rows[0]` when present);
     * level 2 is a `major` phase / summary row; anything else is a level-3 task.
     */
    const projectName = projectRow ? projectRow.name
        : (injected && rows.length && rows[0].major ? rows[0].name : null);
    /**
     * THE THREE LEVEL COLOURS CYCLE. There are three, and a tree can be deeper than three, so they
     * repeat from the beginning rather than a fourth being invented or everything below level 3 being
     * painted alike:
     *
     *     level 1 → colour 1      level 4 → colour 1      level 7 → colour 1
     *     level 2 → colour 2      level 5 → colour 2      …
     *     level 3 → colour 3      level 6 → colour 3
     *
     * The tier comes from the row's own depth in the EXPORT (`level`), never from its position in the
     * table, so hiding a sibling cannot change the colour of anything that remains — which is exactly
     * what the Custom view needs. Rows that carry no `level` keep the original rule: the project row is
     * 1, a `major` phase/summary row is 2, anything else is 3. That is the Launch & Design path, and it
     * is untouched.
     */
    const TIERS = LVL.length; // three, today
    const levelOf = (r) => {
        if (typeof r.level === 'number' && r.level >= 1)
            return ((r.level - 1) % TIERS) + 1;
        return (projectName !== null && r.name === projectName && r.major) ? 1 : (r.major ? 2 : 3);
    };
    /** Left step per hierarchy level inside the Name cell. The ONLY thing that changes with depth in
     *  the table's geometry; column widths, the 73:27 split and every other column are untouched. */
    const INDENT_STEP = 0.30;
    const { X0, XEND, CONTENT_W, TABLE_W, GX, GW, cols, xs, HDR_METRIC_FS } = overviewGeometry();
    const TOP = 2.55, BOTTOM = 13.55; // stop just above the mascot
    // Header height: MEASURED, not assumed. Each heading is wrapped into its own column at 25pt and the
    // row is as deep as the deepest one needs, so no heading is clipped and none had to be shortened.
    const hdrLines = Math.max(...cols.map((c) => c.label.split('\n')
        .reduce((n, line) => n + (0, typography_1.linesNeeded)(line, c.w - 0.16, HDR_METRIC_FS, typography_1.FONT_TEXT), 0)));
    const HDR_H = Math.max(0.95, hdrLines * (0, typography_1.lineHeight)(HDR_METRIC_FS) + 0.12);
    // ---- the Gantt's timeline: ONE scale for every page ----
    // Month columns must be identical on 4A and 4B, so the span is computed across ALL rows once, from
    // the earliest baseline/actual start to the latest baseline/actual finish. Whole months only.
    const allDates = [];
    for (const r of rows)
        for (const d of [r.baseStart, r.baseFinish, r.start, r.finish])
            if (d)
                allDates.push(d);
    const sorted = allDates.slice().sort();
    const firstD = sorted[0] || null, lastD = sorted[sorted.length - 1] || null;
    const months = [];
    if (firstD && lastD) {
        const a = new Date(firstD + 'T00:00:00Z'), b = new Date(lastD + 'T00:00:00Z');
        for (let m = new Date(Date.UTC(a.getUTCFullYear(), a.getUTCMonth(), 1)); m <= b; m.setUTCMonth(m.getUTCMonth() + 1)) {
            months.push(new Date(m));
        }
    }
    const G_PAD = 0.10; // breathing room inside the chart frame
    const plotX = GX + G_PAD, plotW = Math.max(GW - G_PAD * 2, 1);
    const mW = months.length ? plotW / months.length : plotW;
    /** A date's x position on the shared month scale (day-accurate within its month). */
    const xOf = (iso) => {
        if (!iso || !months.length)
            return null;
        const d = new Date(iso + 'T00:00:00Z');
        const idx = (d.getUTCFullYear() - months[0].getUTCFullYear()) * 12 + (d.getUTCMonth() - months[0].getUTCMonth())
            + (d.getUTCDate() - 1) / 31;
        return plotX + Math.max(0, Math.min(months.length, idx)) * mW;
    };
    const todayIso = (reportDate || sch.dataDate || new Date().toISOString().slice(0, 10)).slice(0, 10);
    // ---- row heights: per row, never global; the type size never moves ----
    // The body size is FIXED at OVERVIEW_FS (24pt) for every cell. A long activity name is solved in this
    // order, which is the priority the brief sets: keep the size → wrap inside the Activity cell → let
    // THAT row grow by the extra line. No other row changes, and no other cell in the row changes size.
    const AVAIL = BOTTOM - TOP - HDR_H;
    const ROW_H = (0, typography_1.lineHeight)(typography_1.OVERVIEW_FS) + 0.10; // a one-line row
    const nameW = (r) => cols[0].w - r.indent * INDENT_STEP;
    /** How tall THIS row must be: one line, or more if its own activity name wraps. Nothing else. */
    const heightOf = (r) => Math.max(ROW_H, (0, typography_1.linesNeeded)(r.name, nameW(r) - 0.12, typography_1.OVERVIEW_FS, typography_1.FONT_TEXT) * (0, typography_1.lineHeight)(typography_1.OVERVIEW_FS) + 0.10);
    // The frame is a CONSTANT: header + the full usable height, identical on every page. Pages are then
    // filled by MEASURED row heights up to that height, so a page with a couple of two-line rows holds a
    // row or two fewer — but the table, the chart, the axis and the legend are in exactly the same place.
    const BODY_BOTTOM = TOP + HDR_H + AVAIL;
    // ---- pagination: measured heights, and a section is never left with 1–2 rows at the foot ----
    // A "section" is a phase heading row plus the activities beneath it. A phase may not BEGIN in the last
    // two rows of a page: if fewer than two of its rows would fit, the whole phase — heading and
    // activities — moves to the next page in one piece. (Threshold 2 on request; it was 3.) A phase longer
    // than a page is still split, but only after a full page of it.
    const KEEP = 2;
    const sections = [];
    for (const r of rows) {
        if (r.major || !sections.length)
            sections.push([r]);
        else
            sections[sections.length - 1].push(r);
    }
    const pages = [];
    let cur = [], used = 0;
    const flushPage = () => { if (cur.length)
        pages.push(cur); cur = []; used = 0; };
    for (const sec of sections) {
        // room for three of this section's rows, measured — not counted
        const needKeep = sec.slice(0, KEEP).reduce((n, r) => n + heightOf(r), 0);
        if (cur.length && used + Math.min(needKeep, sec.reduce((n, r) => n + heightOf(r), 0)) > AVAIL)
            flushPage();
        for (const r of sec) {
            const h = heightOf(r);
            if (cur.length && used + h > AVAIL)
                flushPage();
            cur.push(r);
            used += h;
        }
    }
    if (cur.length)
        pages.push(cur);
    // ---- NESTING COLOUR BARS (Design & Build) -------------------------------------------------
    // A summary row's colour is carried down the left of everything nested beneath it, as Asta
    // Powerproject draws a summary over its tasks: one continuous vertical bar in the PARENT's own
    // level colour, running from the foot of the parent row to the foot of its last descendant — that
    // is, until the next row at the same level or shallower.
    //
    // The runs are worked out ONCE over the whole row list, not per page, and then clipped to each
    // page. A phase whose tasks spill onto a continuation page therefore keeps its bar on that page
    // too; computing per page would have dropped the bar wherever the parent row itself was overleaf.
    // Depth comes from `indent`, the row's true WBS depth, so the bar is drawn at every level and each
    // level keeps its own colour.
    const pageStart = [];
    {
        let n = 0;
        for (const p of pages) {
            pageStart.push(n);
            n += p.length;
        }
    }
    const nestRuns = [];
    if (opts.nestingBars) {
        for (let i = 0; i < rows.length; i++) {
            const d = rows[i].indent ?? 0;
            let k = i + 1;
            while (k < rows.length && (rows[k].indent ?? 0) > d)
                k++;
            if (k > i + 1)
                nestRuns.push({ parent: i, from: i + 1, to: k - 1 });
        }
    }
    (pages.length ? pages : [[]]).forEach((chunk, pi) => {
        const s = pptx.addSlide();
        const from = opts.letterFrom ?? 0;
        const total = opts.letterTotal ?? pages.length;
        header(s, t, (0, blocks_1.pageLabel)(sectionNo, from + pi, total), opts.title);
        if (assets.wordmark)
            s.addImage({ data: assets.wordmark, x: 0.57, y: 0.39, w: 3.47, h: 0.5 });
        if (assets.mascot)
            s.addImage({ data: assets.mascot, x: 0.61, y: 13.74, w: 0.74, h: 0.79 });
        const bar = (0, slides_1.HDR_BAR_FILL)(t);
        // Each row's own height, and 24pt throughout — no cell is ever set smaller than the body size.
        const fits = chunk.map((r) => ({ rowH: heightOf(r), fontSize: typography_1.OVERVIEW_FS }));
        // header row — typography only. The size is the single change here: family, colour, fill,
        // alignment, margins, the header row height and every column width are exactly as they were.
        // Header row: WHITE ground, BLACK text — kept clear of the three hierarchy colours so the headings
        // read against all of them. Labels, wording, case, order, widths, height, alignment, margins, size
        // and the gridlines are all exactly as they were; only the two colours changed.
        // The heading's own alignment: every heading is centred in its cell, including Name — the body cells
        // keep their own alignment (Name left, the rest centre), which is why this is `'center'` here and
        // `c.al` on the data rows.
        cols.forEach((c, j) => s.addText(c.label, {
            x: xs[j], y: TOP, w: c.w, h: HDR_H, align: 'center', valign: 'middle',
            fill: { color: HDR_FILL }, color: HDR_TXT, bold: true,
            fontFace: typography_1.FONT_TEXT, fontSize: typography_1.HEADER_FS_OVERVIEW, margin: 3,
            line: { color: GRID, width: 0.75 },
        }));
        // ---------------- the Gantt ----------------
        // No heading above it: the area begins directly with the timeline, as requested.
        // 50 : 50 between the year row and the month row, so both read clearly. (Was 45 : 55.)
        const yearH = HDR_H * 0.50, monthH = HDR_H - yearH;
        // Year band, one cell per year, spanning exactly the months of that year that are on the axis.
        const years = [];
        months.forEach((m, i) => {
            const y = m.getUTCFullYear();
            const last = years[years.length - 1];
            if (last && last.y === y)
                last.n += 1;
            else
                years.push({ y, from: i, n: 1 });
        });
        // The year band follows the Tracker & Schedule calendar's own language — the same grey ground, black
        // numerals and white gridlines between years — so the two slides read as one scheduling system
        // rather than two chart styles. Only the width differs, because this one is compressed to 27%.
        for (const y of years) {
            s.addText(String(y.y), {
                x: plotX + y.from * mW, y: TOP, w: y.n * mW, h: yearH, align: 'center', valign: 'middle',
                fill: { color: CAL_HEAD }, color: '000000',
                // Pinned to 19pt: the calendar is explicitly out of scope, so the year label does not follow the
                // column headings down a point.
                fontFace: typography_1.FONT_NUM, fontSize: 19, margin: 0, wrap: false,
                line: { color: 'FFFFFF', width: 0.75 },
            });
        }
        // Month band: EVERY month carries its first letter — J F M A M J J A S O N D, repeating per year.
        //
        // A month column at 27% of the slide is about a sixth of an inch, so the letter is set at the
        // largest size that fits its own column with a hairline to spare. It is axis furniture, not body
        // text, so it is allowed below the body's readable floor; what it is NOT allowed to do is overflow
        // its column, which is what "no month letter is clipped" means. The size is reported.
        s.addShape('rect', { x: plotX, y: TOP + yearH, w: plotW, h: monthH, fill: { color: 'FFFFFF' }, line: { color: G_AXIS, width: 0.5 } });
        let moFs = typography_1.MIN_FS;
        while (moFs > 8 && (0, typography_1.textWidth)('W', moFs, typography_1.FONT_TEXT) > mW - 0.02)
            moFs -= 1;
        months.forEach((m, i) => {
            const mx = plotX + i * mW;
            // A cell per month, exactly as the Tracker draws its month band — the separation is what keeps a
            // run of initials from reading as one word.
            s.addShape('rect', { x: mx, y: TOP + yearH, w: mW, h: monthH, fill: { color: 'FFFFFF' }, line: { color: G_AXIS, width: 0.4 } });
            s.addText(MO3[m.getUTCMonth()].charAt(0), {
                x: mx, y: TOP + yearH, w: mW, h: monthH, align: 'center', valign: 'middle',
                color: '404040', fontFace: typography_1.FONT_TEXT, fontSize: moFs, margin: 0, wrap: false,
            });
        });
        let ry = TOP + HDR_H;
        /** Each drawn row's top and bottom on this page, so the nesting bars can span exact rows. */
        const rowTop = [], rowBot = [];
        chunk.forEach((r, i) => {
            const rowH = fits[i].rowH;
            const fs = fits[i].fontSize;
            rowTop[i] = ry;
            rowBot[i] = ry + rowH;
            // Row background and text: one of the three hierarchy styles, chosen from the row's OWN level.
            // Level 1 is the project row itself (the root WBS node, carried in as the first row); level 2 is
            // any `major` phase/summary row; everything else is level 3.
            const lvl = LVL[levelOf(r) - 1];
            const fill = lvl.fill;
            // With white rows the separators must be a light grey grid, or the table would lose its lines.
            s.addShape('rect', { x: X0, y: ry, w: TABLE_W, h: rowH, fill: { color: fill }, line: { color: GRID, width: 0.75 } });
            // ---- this row's Gantt band, at exactly the same y and height as the table row ----
            // The chart's background is always white. The table may band a section row olive; the Gantt does
            // not inherit it — a tint behind the bars competes with them and reads as a status it is not.
            s.addShape('rect', {
                x: GX, y: ry, w: GW, h: rowH,
                fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.5 },
            });
            // TWO LANES on ONE activity row, as the reference draws them:
            //   lane 0 — the PROJECTION (green) across the whole projected span, with the ACTUAL PROGRESS
            //            (blue) drawn OVER it from the same start, ending at the progress date. The blue
            //            therefore overlaps the green instead of sitting in a lane of its own.
            //   lane 1 — the BASELINE (yellow).
            // Nothing is moved to another row and no row is duplicated: three bars, one row, always.
            const barGap = rowH * 0.12;
            const barH = (rowH - barGap * 3) / 2;
            const laneY = (k) => ry + barGap + k * (barH + barGap);
            const draw = (from, to, k, colour) => {
                const x1 = xOf(from), x2 = xOf(to);
                if (x1 === null || x2 === null)
                    return;
                const w = Math.max(x2 - x1, 0.035); // a milestone still shows as a tick
                s.addShape('rect', {
                    x: x1, y: laneY(k), w: Math.min(w, plotX + plotW - x1), h: barH,
                    fill: { color: colour }, line: { type: 'none' },
                });
            };
            // Actual Progress is the completed part of the projected span: start + % of its duration.
            const progressEnd = () => {
                if (!r.start || !r.finish || r.pct <= 0)
                    return null;
                const a = Date.parse(r.start + 'T00:00:00Z'), b = Date.parse(r.finish + 'T00:00:00Z');
                if (!Number.isFinite(a) || !Number.isFinite(b) || b < a)
                    return null;
                return new Date(a + (b - a) * Math.min(r.pct, 100) / 100).toISOString().slice(0, 10);
            };
            draw(r.start, r.finish, 0, G_PROJ); // green: the full projected span
            draw(r.start, progressEnd(), 0, G_ACTUAL); // blue: OVER the green — same lane, same row
            draw(r.baseStart, r.baseFinish, 1, G_BASELINE); // yellow: the baseline, its own lane
            // Elapsed months are the elapsed days over the mean month length — a presentation of the same
            // number, not a second source of truth, so the two duration columns cannot disagree.
            const vals = [
                r.name,
                r.duration > 0 ? r.duration + 'ed' : '',
                r.duration > 0 ? (r.duration / 30.4375).toFixed(2) + 'em' : '',
                r.pct + '%',
                (0, slides_1.fmtDate)(r.baseFinish), (0, slides_1.fmtDate)(r.start), (0, slides_1.fmtDate)(r.finish),
            ];
            cols.forEach((c, j) => {
                // THE NAME COLUMN IS COLUMN 0. It is labelled "Name"; this tested for "Activity", a label
                // that does not exist in `cols`, so `isName` was never true — the indent was computed on
                // every row and then multiplied by nothing, and every level printed hard against the same
                // left edge. That is why the hierarchy read as flat however correct the depths were.
                // Keyed off the column INDEX now, so renaming the heading cannot silently break it again.
                const isName = j === 0;
                const ind = isName ? r.indent * INDENT_STEP : 0;
                const v = vals[j];
                // No conditional colour anywhere in the table. Delay highlighting was removed earlier, and the
                // green "100% complete" fill is now removed too: a completed activity uses the same background
                // as any other. Colour on this slide belongs to the section bands and to the Gantt bars.
                const cellFill = undefined;
                if (cellFill)
                    s.addShape('rect', { x: xs[j], y: ry, w: c.w, h: rowH, fill: { color: cellFill }, line: { color: GRID, width: 0.75 } });
                s.addText(v, {
                    x: xs[j] + ind, y: ry, w: Math.max(c.w - ind, 0.3), h: rowH,
                    align: c.al, valign: 'middle',
                    // ONLY the Activity cell may take a reduced size. Every other column stays at the body size —
                    // a long activity name must not shrink its row's dates or percentage.
                    fontSize: isName ? fs : typography_1.OVERVIEW_FS,
                    fontFace: c.num ? typography_1.FONT_NUM : (0, typography_1.fontFor)(v),
                    color: lvl.text, bold: levelOf(r) <= 2, // the level's own text colour
                    margin: 3, wrap: true,
                });
            });
            ry += rowH;
        });
        // ---- the nesting colour bars, drawn OVER the finished rows ----
        // Each run was computed over the whole row list; here it is clipped to the rows this page
        // actually drew, so a run that begins on the previous page still shows from the top of this one.
        // The bar sits in the indent step that separates the parent's text from its children's, which is
        // why it reads as the parent's line running down beside them rather than as a cell fill.
        if (opts.nestingBars && nestRuns.length) {
            const g0 = pageStart[pi] ?? 0, g1 = g0 + chunk.length - 1;
            for (const run of nestRuns) {
                const from = Math.max(run.from, g0), to = Math.min(run.to, g1);
                if (from > to)
                    continue; // no part of this run is on this page
                const yTop = rowTop[from - g0], yBot = rowBot[to - g0];
                if (yTop === undefined || yBot === undefined)
                    continue;
                const parent = rows[run.parent];
                const colour = LVL[levelOf(parent) - 1].fill;
                // A white parent (a level that falls on the leaf tier) has no line to carry down.
                if (!colour || colour.toUpperCase() === 'FFFFFF')
                    continue;
                // EACH BAR FILLS ITS WHOLE INDENT STEP, so consecutive levels' bars TOUCH and read as one
                // nested block of colour. They used to be 0.12" wide and centred inside the 0.30" step, which
                // left a 0.09" gap either side: at three or four levels deep that came out as a row of thin
                // separate stripes with white between them, which is what the slide looked wrong for. The bar
                // occupies exactly the gap between the parent's text and its children's, so it still cannot
                // reach any text.
                s.addShape('rect', {
                    x: X0 + (parent.indent ?? 0) * INDENT_STEP,
                    y: yTop, w: INDENT_STEP, h: yBot - yTop,
                    fill: { color: colour }, line: { type: 'none' },
                });
            }
        }
        // Vertical column separators. Drawn once per page, after the rows, so each is a single unbroken
        // line from the header band to the last row — the same light grey and the same 0.75pt weight as
        // the row separators, on the column boundaries the cells themselves were placed on, so they cannot
        // be a hairline out. The outer edges are left to the boundary rectangle below; nothing else is added.
        for (let j = 1; j < cols.length; j++) {
            // to the foot of the last populated row — the rules stop where the data stops
            s.addShape('line', { x: xs[j], y: TOP + HDR_H, w: 0, h: ry - (TOP + HDR_H), line: { color: GRID, width: 0.75 } });
        }
        // TABLE BOTTOM = GANTT BOTTOM = the foot of the LAST POPULATED ROW. No blank row and no bordered
        // empty area, and the two frames cannot disagree because both are drawn to the same `ry`.
        //
        // This does not conflict with "the same component size on every page": everything that defines the
        // component — its top edge, the column widths, the 73:27 split, the row height and therefore how many
        // rows a page holds — is identical page to page. Only the frame stops where the data stops.
        const gridBottom = ry;
        s.addShape('rect', { x: X0, y: TOP, w: TABLE_W, h: gridBottom - TOP, fill: { type: 'none' }, line: { color: bar, width: 1.25 } });
        s.addShape('rect', { x: GX, y: TOP, w: GW, h: gridBottom - TOP, fill: { type: 'none' }, line: { color: bar, width: 1.25 } });
        // TODAY: a vertical rule at the progress date, drawn last so it sits over the bars. Its position is
        // derived from the report date on the same month scale as the bars — never a fixed x.
        const tx = xOf(todayIso);
        if (tx !== null && tx > plotX && tx < plotX + plotW) {
            // From the month band down to the LAST ACTIVITY ROW and no further. The frame is a fixed size so
            // every page matches, but the rule belongs to the bars: it stops where the activities stop rather
            // than running on through the blank part of the grid.
            s.addShape('line', { x: tx, y: TOP + yearH, w: 0, h: Math.max(ry, TOP + HDR_H) - (TOP + yearH), line: { color: G_TODAY, width: 1.5 } });
            // The tag sits in the clear space ABOVE the axis, not inside the year band: printed over the band
            // it covered the year, which layout validation reported as merged text.
            const tw = (0, typography_1.textWidth)('TODAY', typography_1.MIN_FS, typography_1.FONT_TEXT) + 0.26;
            const TAG_H = 0.30;
            s.addText('TODAY', {
                x: Math.min(Math.max(tx - tw / 2, plotX), plotX + plotW - tw), y: TOP - TAG_H - 0.03, w: tw, h: TAG_H,
                align: 'center', valign: 'middle', margin: 0, wrap: false,
                fill: { color: G_TODAY }, color: 'FFFFFF', bold: true, fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS,
            });
        }
        // ---- legend: bottom-right, directly beneath the Gantt ----
        // Order as specified. Each item is measured, so the spacing between items is equal and the block
        // ends exactly on the content margin whatever the labels are.
        // Fixed position: measured from the page frame, not from the last row, so it does not move page to
        // page. "Completed" is not listed — the Gantt draws no completed bar, and the brief says to show it
        // only if that category is actually used.
        const ly = gridBottom + 0.16;
        const legend = [
            ['Baseline', G_BASELINE], ['Actual Progress', G_ACTUAL], ['Projection', G_PROJ],
        ];
        const SW = 0.42, SW_GAP = 0.14, ITEM_GAP = 0.45;
        const itemW = legend.map(([lab]) => SW + SW_GAP + (0, typography_1.textWidth)(lab, typography_1.HEADER_FS, typography_1.FONT_TEXT) + 0.06);
        const blockW = itemW.reduce((a, b) => a + b, 0) + ITEM_GAP * (legend.length - 1);
        let lx = XEND - blockW;
        legend.forEach(([lab, col], i) => {
            s.addShape('rect', { x: lx, y: ly, w: SW, h: 0.3, fill: { color: col }, line: { color: 'BFBFBF', width: 0.5 } });
            s.addText(lab, {
                x: lx + SW + SW_GAP, y: ly - 0.04, w: itemW[i] - SW - SW_GAP, h: 0.38,
                fontSize: typography_1.HEADER_FS, color: t.dark, fontFace: typography_1.FONT_TEXT, valign: 'middle', margin: 0, wrap: false,
            });
            lx += itemW[i] + ITEM_GAP;
        });
    });
    return { rows: rows.length, pages: Math.max(1, pages.length) };
}
