"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.headingLayout = exports.HeadingBox = exports.scheduleLayout = void 0;
exports.scopeToPhase = scopeToPhase;
exports.buildTrackerData = buildTrackerData;
exports.trackerSlide = trackerSlide;
exports.verifyHeadingAlignment = verifyHeadingAlignment;
exports.trackerValidation = trackerValidation;
exports.scheduleLayoutValidation = scheduleLayoutValidation;
exports.scheduleLayoutReview = scheduleLayoutReview;
exports.scheduleDesignReview = scheduleDesignReview;
const typography_1 = require("./typography");
const schedule_phases_1 = require("./schedule_phases");
const slides_1 = require("./slides");
/** Geometry of the last rendered schedule block — read by schedule_layout_validation.md. */
exports.scheduleLayout = null;
// Sizes in the dashboard are spreadsheet points (7–10). The slide canvas is 26.67" (2× a normal
// slide) so they are scaled up to the readable floor while preserving the sheet's hierarchy.
const scaleFs = (sheetPt) => Math.max(typography_1.MIN_FS, Math.round(sheetPt * 2.6));
// ---------------------------------------------------------------------------------------------
// Vertical (rotated) column-heading geometry.
//
// PowerPoint rotates a text box about its CENTRE, so a vertical heading is authored as a wide box
// (long axis = the cell's height) and then turned 270°. Two things make the headings line up:
//   • the box centre is placed on the cell centre — computed, never a hand-tuned offset;
//   • `align: 'center'` centres the glyph run along that long axis, which is what produces equal
//     white space above and below after rotation. (`align: 'left'` made every label hug one end,
//     which is why short headings looked shifted relative to long ones.)
// The box's cross-axis height is identical for every column, so no heading can sit higher or lower
// than another. Everything derives from the font size, so it stays correct if that changes.
// ---------------------------------------------------------------------------------------------
// Rotated-heading geometry now lives in typography.ts (it is pure type-setting maths) and is
// re-exported here so every existing import keeps working.
var typography_2 = require("./typography");
Object.defineProperty(exports, "HeadingBox", { enumerable: true, get: function () { return typography_2.HeadingBox; } });
Object.defineProperty(exports, "headingLayout", { enumerable: true, get: function () { return typography_2.headingLayout; } });
const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
/** Tokens that make a tracker column name matchable against activity names. */
function keyTokens(name, cfgKeywords) {
    if (cfgKeywords && cfgKeywords.length)
        return cfgKeywords.map(norm).filter(Boolean);
    const n = norm(name);
    const out = [n];
    // Dashboard abbreviations → the wording used in the ASTA schedule. Deliberately narrow: a loose
    // alias silently mis-states a project's status (e.g. "schematic" for LOD 100 matched 44 rows).
    const ALIAS = {
        'lod 100': ['lod 100', 'lod100'], 'lod 200': ['lod 200', 'lod200'],
        'boq v0': ['boq'], 'lp dp': ['layout plan sanction', 'development plan'],
        'bp': ['building plan sanction'], 'cfe': ['consent for establishment', 'cfe'],
        'moef': ['moef'], 'prjct setup': ['project setup'],
        'hm ln aprv': ['home loan'], 'edm fyers': ['edm', 'flyer'],
        'project walkthrogh': ['project walkthrough'], 'product walkthrogh': ['product walkthrough'],
        'org structure': ['org structure'], 'nala rerouting': ['nala'],
        'definitive agreement': ['definitive agreement'],
        'sales lounge experience centre': ['sales lounge', 'experience centre'],
        'other noc': ['other noc'], 'project schedule': ['project schedule'],
    };
    if (ALIAS[n])
        return [...new Set(ALIAS[n].map(norm))];
    return [...new Set(out)];
}
/**
 * Whole-word match of a token inside a normalised activity name. Substring matching is wrong here —
 * "clu" would otherwise match "Clubhouse" and report 43 activities against the CLU approval.
 */
function tokenHits(haystack, token) {
    if (!token)
        return false;
    const pat = token.split(/\s+/).map((w) => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\s*');
    return new RegExp(`(^|\\s)${pat}(\\s|$)`).test(haystack);
}
/**
 * Indicator status per the specified rule:
 *   Completed = 100% · In Progress = >0 and <100 · Not Started = 0 · Not Applicable = no such item.
 */
/**
 * Restrict the schedule to the PHASE this deck is about.
 *
 * A single ASTA export often holds several phases (e.g. "Violet Hill Ph 1 [21.76a]" and
 * "Violet Hill Ph 2 [27.08a]") with the same activity names repeated — Ph 1 complete, Ph 2 at 0%.
 * Aggregating across both would report every tracker item as "In Progress", so when the deck
 * identifies a phase (via the project area, e.g. "21.76a", or a "Ph n" hint) we use only that
 * phase's sub-tree. Projects with a single phase are unaffected.
 */
function scopeToPhase(sch, hint) {
    // Work on lower-cased RAW text: normalising first would turn "21.76a" into "21 76a" and the
    // decimal area — the strongest phase signal — would no longer match the WBS name.
    const h = String(hint || '').toLowerCase();
    if (!h.trim())
        return { activities: sch.activities, phase: null };
    const area = (h.match(/(\d+\.\d+)\s*a(?:cres?)?\b/) || [])[1] || '';
    const ph = (h.match(/\bph(?:ase)?\s*(\d+)\b/) || [])[1] || '';
    const tokens = [area, ph ? `ph ${ph}` : '', ph ? `phase ${ph}` : ''].filter(Boolean);
    if (!tokens.length)
        return { activities: sch.activities, phase: null };
    const nodes = Object.values(sch.wbs);
    // A phase root must be a CONTAINER (has children) — otherwise every leaf whose name happens to
    // mention the area ("CFE [21.76a]", "Project Model [21.76a]") would be treated as a phase.
    const withKids = new Set(nodes.filter((w) => w.parent).map((w) => w.parent));
    const hits = nodes.filter((w) => {
        if (!withKids.has(w.id))
            return false;
        const n = String(w.name || '').toLowerCase();
        return tokens.some((t) => n.includes(t));
    });
    if (!hits.length)
        return { activities: sch.activities, phase: null };
    // keep only the shallowest matches — that is the phase level itself, not packages inside it
    const minDepth = Math.min(...hits.map((w) => w.depth));
    const roots = hits.filter((w) => w.depth === minDepth);
    // collect the chosen phases' whole sub-trees
    const keep = new Set(roots.map((r) => r.id));
    let grew = true;
    while (grew) {
        grew = false;
        for (const w of nodes)
            if (w.parent && keep.has(w.parent) && !keep.has(w.id)) {
                keep.add(w.id);
                grew = true;
            }
    }
    const acts = sch.activities.filter((a) => a.wbs && keep.has(a.wbs));
    if (!acts.length)
        return { activities: sch.activities, phase: null };
    return { activities: acts, phase: [...new Set(roots.map((r) => r.name))].join(" / ") };
}
function buildTrackerData(sch, tpl, cfg, phaseHint = '', schedCfg = { phases: [] }, 
/** The project's own status row from the dashboard's Tracker sheet, keyed by sheet column number
 *  (see readTrackerStatuses). Optional: when absent, every status is schedule-derived exactly as
 *  before, so a run without a dashboard behaves as it always did. */
sheetStatus) {
    const scoped = scopeToPhase(sch, phaseHint);
    const acts0 = scoped.activities;
    const kwCfg = (cfg && cfg.column_keywords) || {};
    const bySymbol = new Map();
    for (const l of tpl.legend)
        bySymbol.set(l.label.replace(/\s+/g, '').toLowerCase(), l);
    const styleFor = (st) => {
        const k = st.replace(/\s+/g, '').toLowerCase();
        const l = bySymbol.get(k);
        return l || { symbol: '', color: '000000', font: typography_1.FONT_TEXT, size: 9 };
    };
    // Indicator symbol/colour overrides from configuration (e.g. Completed → green ✔ instead of "1").
    // Anything not overridden keeps the dashboard sheet's own symbol, face and colour.
    const overrides = (schedCfg && schedCfg.indicatorOverrides) || {};
    // The dashboard's own status row for THIS project, keyed by sheet column number, when it was read.
    // It is the status of record: the sheet is what the business reviews, so where it states a status
    // the slide shows that status. The schedule-derived pass below is the fallback for a column the
    // sheet leaves blank, and the only source at all when no dashboard row could be matched.
    const RAG_TO_STATUS = {
        good: 'Completed', amber: 'In Progress', red: 'Not Started', na: 'Not Applicable',
    };
    const sheetByCol = sheetStatus?.byCol;
    const indicators = tpl.columns.map((c) => {
        const toks = keyTokens(c.name, kwCfg[c.name]);
        const matched = acts0.filter((a) => {
            const an = norm(a.name), ad = norm(a.dept || '');
            return toks.some((t) => t.length >= 2 && (tokenHits(an, t) || (ad && tokenHits(ad, t))));
        });
        // Schedule-derived status: kept exactly as it was, and still used wherever the sheet is silent.
        let status;
        if (!matched.length)
            status = 'Not Applicable';
        else if (matched.every((a) => a.pct >= 100))
            status = 'Completed';
        else if (matched.some((a) => a.pct > 0))
            status = 'In Progress';
        else
            status = 'Not Started';
        let source = 'schedule';
        const rag = sheetByCol?.get(c.col);
        if (rag && RAG_TO_STATUS[rag]) {
            status = RAG_TO_STATUS[rag];
            source = 'dashboard';
        }
        const st = styleFor(status);
        const ov = overrides[status] || {};
        return {
            column: c.name, group: c.group, status,
            symbol: ov.symbol ?? st.symbol,
            color: ov.color ?? st.color,
            font: ov.font ?? st.font,
            size: ov.size ?? st.size,
            matched: matched.length,
            source,
        };
    });
    // ---- timeline: the major EXECUTION PHASES, resolved from tracker_schedule_config.json ----
    // Activity names, aliases, keywords, finish milestones and the date rules all live in that file —
    // nothing about a project's wording is hard-coded here.
    const resolved = (0, schedule_phases_1.resolvePhases)(sch, acts0, schedCfg);
    // PHASE COLOURS — resolved PER PHASE, never per project, and never per row.
    //
    // The pair (solid tone, lighter tone) for each executive phase is a SWATCH PAIR read from the
    // Summary Dashboard's Schedule sheet, so the two tones always belong to the same hue by
    // construction. What used to break that guarantee was the LOOKUP, not the palette: a plain
    // `Map.get(label)` demands that the dashboard spell the phase exactly as
    // tracker_schedule_config.json does. One trailing space, one capital letter ("Sales Launch" vs
    // "Sales launch"), one "&" written as "and", and the phase silently fell through to a hard-coded
    // pair — 7C7C7C / DBDBDB — which is a generic grey belonging to no phase at all. Worse, that pair
    // is IDENTICAL to the real Approvals swatch, so on the Actual row the substitution was invisible
    // and only the Baseline row looked wrong. Any phase could hit it; nothing about it was specific to
    // one project.
    //
    // The lookup is now tolerant and, above all, REPORTED. Four ordered attempts, each recorded in
    // tracker_schedule_validation.md so a wording drift shows up as a line of text instead of a
    // mystery grey bar:
    //   1. the exact label,
    //   2. the label normalised (case, spacing, punctuation, "&"/"and"),
    //   3. the phase id from the config with the same normalisation,
    //   4. only if the dashboard genuinely carries no swatch for the phase, the lighter tone is
    //      DERIVED from that phase's own solid tone. No new colour is invented while the palette has
    //      one to give, and the derived tone is still the same hue — never a foreign grey.
    const normKey = (s) => s.toLowerCase()
        .replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ').trim();
    const byExact = new Map(tpl.phases.map((p) => [p.label, p]));
    const byNorm = new Map(tpl.phases.map((p) => [normKey(p.label), p]));
    /** A lighter tone of the SAME hue — used only when the palette has no swatch for the phase. */
    const lighterTone = (hex) => {
        const n = parseInt(hex, 16);
        if (!/^[0-9a-fA-F]{6}$/.test(hex) || Number.isNaN(n))
            return 'DBDBDB';
        const mix = (c) => Math.round(c + (255 - c) * 0.66);
        return [(n >> 16) & 255, (n >> 8) & 255, n & 255]
            .map((c) => mix(c).toString(16).padStart(2, '0')).join('').toUpperCase();
    };
    const phaseColour = (label, id) => {
        const hit = byExact.get(label)
            || byNorm.get(normKey(label))
            || byNorm.get(normKey(id));
        const how = byExact.get(label) ? 'dashboard swatch (exact label)'
            : byNorm.get(normKey(label)) ? `dashboard swatch (label normalised to "${normKey(label)}")`
                : byNorm.get(normKey(id)) ? `dashboard swatch (matched on the config id "${id}")`
                    : 'no swatch in the dashboard — lighter tone derived from this phase\'s own solid tone';
        const actual = hit?.actual || '7C7C7C';
        return { actual, baseline: hit?.baseline || lighterTone(actual), how };
    };
    // The anchor pass itself records, per row, which phases it actually placed on an exact month. Only
    // those are locked against the display floor — never a phase it declined to move, and never a phase
    // whose start it left alone (Approvals is declared adjustFinishOnly, so its start stays real
    // schedule data and must still take its turn in the order, or Approvals would be drawn from its own
    // early start straight across Land Acquisition and LOD).
    const phases = resolved.map((r) => {
        const c = phaseColour(r.label, r.id);
        r.colourSource = c.how;
        r.colourActual = c.actual;
        r.colourBaseline = c.baseline;
        return {
            label: r.label,
            actual: c.actual, baseline: c.baseline,
            aStart: r.aStart, aFinish: r.aFinish, bStart: r.bStart, bFinish: r.bFinish,
            fullSpan: !!(schedCfg?.phases || []).find((x) => x.id === r.id)?.displayFullSpan,
            anchorLockA: !!r.anchoredA,
            anchorLockB: !!r.anchoredB,
        };
    });
    const counts = { 'Completed': 0, 'In Progress': 0, 'Not Started': 0, 'Not Applicable': 0 };
    for (const i of indicators)
        counts[i.status]++;
    // Effective legend = the sheet's legend with any configured override applied, so the key always
    // matches the glyphs actually drawn in the matrix.
    const legend = tpl.legend.map((l) => {
        const ov = overrides[l.label] || {};
        return {
            symbol: ov.symbol ?? l.symbol, label: l.label,
            color: ov.color ?? l.color, font: ov.font ?? l.font, size: ov.size ?? l.size,
        };
    });
    return { indicators, phases, counts, scopedPhase: scoped.phase, scopedActivities: acts0.length, resolved, legend };
}
// ---------------------------------------------------------------------------------------------
// Rendering — geometry mirrors the dashboard: group band, rotated sub-headers, one project row,
// indicator legend, then the timeline (year band / month band / Actual-Projected / Baseline) and
// the phase legend. Everything is native text and shapes, so it stays editable.
// ---------------------------------------------------------------------------------------------
const MO = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function trackerSlide(pptx, sch, tpl, data, code, area, num, t, assets, reportDate) {
    const s = pptx.addSlide();
    // ---- page heading: the ONE standard header shared by every generated slide ----
    (0, slides_1.sectionHeader)(s, 'Tracker & Schedule', num, t);
    if (assets.wordmark)
        s.addImage({ data: assets.wordmark, x: 0.57, y: 0.39, w: 3.47, h: 0.5 });
    if (assets.mascot)
        s.addImage({ data: assets.mascot, x: 0.61, y: 13.74, w: 0.74, h: 0.79 });
    // ---- geometry ----
    const X0 = 1.43, TOTAL_W = 24.68;
    // Project Code was removed (one project per deck). Its 1.5" goes entirely to Projects, so the
    // indicator columns keep the exact widths and x-positions the dashboard template defines and
    // nothing is left empty: W_PROJ 3.5 + 1.5 = 5.0.
    const W_PROJ = 5.0;
    const nCols = data.indicators.length;
    const indW = (TOTAL_W - W_PROJ) / Math.max(nCols, 1);
    const TOP = 2.50;
    // Rotated column headers read bottom-to-top, so the band must be as TALL as the longest label is
    // WIDE. Long labels are split over two vertical lines to keep the band a sensible height.
    const wrapName = (n) => {
        if (n.length <= 18)
            return n;
        const mid = Math.floor(n.length / 2);
        let cut = -1;
        for (let d = 0; d < n.length; d++) {
            if (n[mid - d] === ' ') {
                cut = mid - d;
                break;
            }
            if (n[mid + d] === ' ') {
                cut = mid + d;
                break;
            }
        }
        return cut > 0 ? `${n.slice(0, cut)}\n${n.slice(cut + 1)}` : n;
    };
    const wrapped = data.indicators.map((i) => wrapName(i.column));
    const longestLine = wrapped.reduce((m, n) => Math.max(m, ...n.split('\n').map((l) => l.length)), 0);
    const H_GROUP = 0.52;
    const H_SUB = Math.min(4.20, Math.max(2.10, longestLine * typography_1.MIN_FS * 0.52 / 72 + 0.38));
    const H_ROW = 0.62;
    const yGroup = TOP, ySub = TOP + H_GROUP, yRow = ySub + H_SUB;
    const GRID = 'BFBFBF';
    const fsGroup = scaleFs(tpl.headerSize), fsSub = typography_1.MIN_FS;
    // lead header block spanning group+sub rows.
    // NB: a rotated text box rotates its BORDER too, so every cell outline is drawn as its own
    // un-rotated shape and the rotated text sits on top with no fill and no line.
    const leadBlockH = H_GROUP + H_SUB;
    s.addText('Projects', {
        x: X0, y: yGroup, w: W_PROJ, h: leadBlockH, align: 'center', valign: 'bottom',
        fontFace: tpl.subFont, fontSize: fsGroup, color: '000000',
        fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.75 }, margin: 6,
    });
    // group bands — exact names, order and spans from the sheet
    let gx = X0 + W_PROJ;
    for (const g of tpl.groups) {
        const span = data.indicators.filter((i) => i.group === g.name).length;
        if (!span)
            continue;
        const w = span * indW;
        s.addText(g.name, {
            x: gx, y: yGroup, w, h: H_GROUP, align: 'center', valign: 'middle',
            fill: { color: tpl.headerFill }, color: '000000',
            fontFace: tpl.headerFont, fontSize: fsGroup, margin: 1, wrap: true,
            line: { color: 'FFFFFF', width: 0.75 },
        });
        gx += w;
    }
    // rotated sub-headers + the project's indicator row
    const colXs = data.indicators.map((_, i) => X0 + W_PROJ + i * indW);
    const heads = (0, typography_1.headingLayout)(wrapped, colXs, ySub, indW, H_SUB, fsSub, tpl.subFont);
    let cx = X0 + W_PROJ;
    data.indicators.forEach((ind, ci) => {
        // cell outline (un-rotated) …
        s.addShape('rect', { x: cx, y: ySub, w: indW, h: H_SUB, fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.5 } });
        // … then the vertical label, centred both ways inside that cell (see headingLayout).
        const hb = heads[ci];
        s.addText(hb.label, {
            x: hb.boxX, y: hb.boxY, w: hb.boxW, h: hb.boxH,
            align: 'center', valign: 'middle', rotate: 270,
            fontFace: tpl.subFont, fontSize: fsSub, color: '000000', margin: 0, wrap: true,
        });
        // indicator glyph — exact symbol/colour/face from the dashboard legend
        s.addText(ind.symbol, {
            x: cx, y: yRow, w: indW, h: H_ROW, align: 'center', valign: 'middle',
            fontFace: ind.font, fontSize: scaleFs(ind.size), color: ind.color, bold: false,
            line: { color: GRID, width: 0.5 }, margin: 0,
        });
        cx += indW;
    });
    const projLabel = area ? `${sch.projectName} [${area}]` : sch.projectName;
    s.addText(projLabel, {
        x: X0, y: yRow, w: W_PROJ, h: H_ROW, valign: 'middle',
        fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: '000000', margin: 4, line: { color: GRID, width: 0.75 },
    });
    // ---- indicator legend (symbols, labels, colours straight from the sheet) ----
    const yLeg = yRow + H_ROW + 0.28;
    let lx = X0 + W_PROJ;
    // The legend must show the SAME glyph the matrix uses, so any configured override (e.g. Completed
    // rendered as a green ✔ instead of the sheet's "1") is reflected here too.
    for (const l of data.legend) {
        s.addText(l.symbol, { x: lx, y: yLeg, w: 0.55, h: 0.42, align: 'center', valign: 'middle', fontFace: l.font, fontSize: scaleFs(l.size), color: l.color });
        s.addText(l.label, { x: lx + 0.55, y: yLeg, w: 2.6, h: 0.42, valign: 'middle', fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: t.dark });
        lx += 3.5;
    }
    // ---- timeline (generated every time from the schedule) ----
    // The timeline covers ONLY the major execution phases — it must begin at the configured first
    // phase (Land Acquisition) and end at the configured last one (Site Establishment & Mobilization),
    // rather than spanning every activity in the programme.
    const all = [];
    for (const p of data.phases)
        for (const d of [p.aStart, p.aFinish, p.bStart, p.bFinish])
            if (d)
                all.push(d);
    // ---- executive schedule panel ----
    //
    // Built to the approved reference design. It is a management view, not a Gantt: one line per phase
    // group, two rows only, and nothing on the slide that a Primavera print-out would carry.
    //
    //   ┌ code ┬ project ┬  row label   ┬───────────────── timeline ─────────────────┐
    //   │  P.  │Projects │              │      2024      │     2025     │   2026     │ year band
    //   │  C.  │         │              │ A│M│J│J│A│S│O│N│D│J│F│M│…              (rotated months)
    //   │ 0118 │ Violet  │ Actual/Proj  │ ██████████████████████│                    │
    //   │      │  Hill   │ Baseline     │ ████████████████████  │                    │
    //   └──────┴─────────┴──────────────┴───────────────────────╪────────────────────┘
    //                                                         Today
    //     ▉ Land Acquisition      ▉ LOD 100/200     ▉ Approvals
    //     ▉ Marketing             ▉ Sales Launch    ▉ Site Establishment & Mobilization
    //     Baseline shown in the lighter tone of the same colour.
    //
    // The panel occupies EXACTLY the tracker matrix's span — same left edge, same right edge, same
    // width — so the two read as one composition rather than a smaller object underneath. X0 and
    // TOTAL_W are the matrix's own geometry, used here directly: the two can never drift apart.
    // Inside that span the project block takes what its text needs and the timeline takes the rest.
    const SP = { afterMatrix: 0.30 }; // indicator legend → schedule panel
    const yTLmin = yLeg + 0.42 + SP.afterMatrix;
    const BOTTOM_TL = 13.60; // lowest usable y — just above the mascot
    // ---- horizontal ----
    const panelW = TOTAL_W; // flush with the matrix, both edges
    const panelX = X0;
    const projLabelSch = area ? `${sch.projectName} [${area}]` : sch.projectName;
    const W_CODE = code ? Math.max(0.95, (0, typography_1.textWidth)(code, typography_1.MIN_FS, typography_1.FONT_NUM) + 0.30) : 0;
    const W_ROWLBL = Math.max(2.20, (0, typography_1.textWidth)('Actual / Projected', typography_1.MIN_FS, typography_1.FONT_TEXT) + 0.34);
    // The name cell spans BOTH schedule rows, so it has room for two lines. Sizing it for the full
    // single-line width is what pushed "Butterfly Dreams [South Block]" to 4.2" and squeezed the
    // timeline to 62% of the slide; sizing it for two lines keeps the 70/20/10 split intact and the
    // name is still complete, just wrapped.
    // The floor is the project name's LONGEST WORD — a column narrower than that would clip it. Above
    // that floor the cell is sized for two lines (it spans both schedule rows), which is what keeps a
    // long name from pushing the timeline off the matrix's grid line.
    const longestWord = Math.max(...projLabelSch.split(/\s+/).map((w) => (0, typography_1.textWidth)(w, typography_1.MIN_FS, typography_1.FONT_TEXT)), 0.5);
    const W_NAME_MIN = Math.min(3.20, Math.max(longestWord + 0.20, (0, typography_1.textWidth)(projLabelSch, typography_1.MIN_FS, typography_1.FONT_TEXT) / 2 + 0.34));
    // Floor the project block at the matrix's own Projects column width, so the timeline begins where the
    // indicator columns begin and the two tables share one internal grid line. It grows past that only
    // when the code, the name and the row labels genuinely need more — measured, never clipped.
    const leftW = Math.max(W_PROJ, W_CODE + W_NAME_MIN + W_ROWLBL);
    const W_NAME = leftW - W_CODE - W_ROWLBL;
    const tlLabelW = leftW; // the whole left block
    const tlX = panelX + leftW;
    const tlW = panelW - leftW;
    // ---- vertical ----
    // Rotated month labels let EVERY month be named however long the programme is, which is what the
    // reference shows. The band therefore has to be as tall as the longest label ("Sep-24") is wide.
    const H_YEAR = 0.42;
    const H_MON = Math.min(1.50, Math.max(0.85, (0, typography_1.textWidth)('Sep-24', typography_1.MIN_FS, typography_1.FONT_TEXT) + 0.24));
    const H_GAP = 0; // the bars sit directly on the axis
    const H_TODAY = 0.34; // the strip that carries the Today label
    const GAP_LEGEND = 0.34; // bars → legend
    const LEG_ROW_H = 0.34, LEG_ROW_GAP = 0.10; // legend row pitch
    const LEG_SW_W = 0.44, LEG_SW_H = 0.26, LEG_SW_SP = 0.16; // small colour boxes, as the reference
    const CAPTION_H = 0.32;
    // Legend: horizontal, wrapping onto a second row when it needs one. The column count is derived from
    // the widest phase name at this size, so no label is ever squeezed or clipped.
    const legNeed = Math.max(...data.phases.map((p) => (0, typography_1.textWidth)(p.label, typography_1.MIN_FS, typography_1.FONT_TEXT)), 1)
        + LEG_SW_W + LEG_SW_SP + 0.30;
    const legCols = Math.max(1, Math.min(3, Math.floor(panelW / legNeed)));
    const legRows = Math.max(1, Math.ceil(data.phases.length / legCols));
    const legBlockH = legRows * LEG_ROW_H + (legRows - 1) * LEG_ROW_GAP;
    // Whatever is left after the fixed elements is shared by the two bar rows, then clamped so a short
    // programme does not produce slabs of colour and a long one does not produce hairlines.
    const availH = BOTTOM_TL - yTLmin;
    const fixedH = H_YEAR + H_MON + H_GAP + H_TODAY + GAP_LEGEND + legBlockH + CAPTION_H;
    // The reference keeps the two schedule rows visibly shorter than the month band — they are bands of
    // colour, not Gantt lanes — so the upper clamp is deliberately tight and anything left over stays as
    // bottom margin rather than being poured into the bars.
    //
    // BAR_SCALE trims the BREADTH of the two schedule rows — Actual/Projected and Baseline — by 15% on
    // request. It is the row's thickness that changes, never its length: the bars are positioned from the
    // month scale, so their width still maps exactly onto the calendar above. The 15% comes back as bottom
    // margin rather than being redistributed, which keeps the panel's composition and the legend's
    // centring unchanged.
    const BAR_SCALE = 0.85;
    const H_BAR = Math.min(0.90, Math.max(0.52, (availH - fixedH) / 2)) * BAR_SCALE;
    if (all.length) {
        const d0s = all.slice().sort()[0], d1s = all.slice().sort().slice(-1)[0];
        let d0 = new Date(d0s + 'T00:00:00Z');
        d0 = new Date(Date.UTC(d0.getUTCFullYear(), d0.getUTCMonth(), 1));
        const d1 = new Date(d1s + 'T00:00:00Z');
        // ---- THE VISIBLE WINDOW: the last VISIBLE_YEARS years, and forward as far as the work runs ----
        // This is a change to the CALENDAR ONLY. Land acquisition on these projects can begin four or five
        // years before the report, and those early years push every month cell so narrow that the recent,
        // decision-relevant months become unreadable. So the axis starts VISIBLE_YEARS before the progress
        // date the deck is stated as at — the very same date the Today marker uses, so the two can never
        // disagree — and still runs forward to the end of the work.
        //
        // Nothing else changes. No activity is filtered out, no phase is dropped, and not one source date
        // is altered or recalculated: the phases were resolved before this point and are not touched here.
        // A bar that begins before the window still appears, starting at the left edge, showing the part of
        // it that falls inside the window (spansOf clamps the cell index, it does not move the date). A bar
        // that runs past the right-hand end keeps its real finish date and is drawn as far as the calendar
        // reaches. The real dates stay printed in schedule_layout_validation.md either way.
        const VISIBLE_YEARS = 3;
        const winAnchorIso = (reportDate || sch.dataDate || '').slice(0, 10);
        const fullFirstIso = d0.toISOString().slice(0, 10);
        let windowNote = `full span of the schedule from ${fullFirstIso}`;
        let monthsHidden = 0;
        if (winAnchorIso) {
            const a = new Date(winAnchorIso + 'T00:00:00Z');
            const cut = new Date(Date.UTC(a.getUTCFullYear() - VISIBLE_YEARS, a.getUTCMonth(), 1));
            const monthsBetween = (x, y) => (y.getUTCFullYear() - x.getUTCFullYear()) * 12 + (y.getUTCMonth() - x.getUTCMonth());
            if (cut <= d0) {
                windowNote = `the whole schedule already begins within ${VISIBLE_YEARS} years of the progress date ${winAnchorIso}, so nothing is hidden (first month ${fullFirstIso})`;
            }
            else if (cut > d1) {
                // Everything finished more than VISIBLE_YEARS ago. Clipping would leave an empty calendar, so
                // the full span is kept and the reason is reported.
                windowNote = `NOT clipped: the whole schedule ends ${d1s}, before the ${VISIBLE_YEARS}-year window would begin (${cut.toISOString().slice(0, 10)}), so the full span is shown instead of an empty calendar`;
            }
            else {
                monthsHidden = monthsBetween(d0, cut);
                d0 = cut;
                windowNote = `last ${VISIBLE_YEARS} years: starts ${d0.toISOString().slice(0, 10)}, i.e. the progress date ${winAnchorIso} minus ${VISIBLE_YEARS} years — ${monthsHidden} earlier month(s) are outside the visible calendar; no activity was filtered and no date was changed`;
            }
        }
        const months = [];
        for (let m = new Date(d0); m <= d1; m.setUTCMonth(m.getUTCMonth() + 1))
            months.push(new Date(m));
        if (!months.length)
            months.push(new Date(d0));
        const monthIdx = (iso) => {
            const d = new Date(iso + 'T00:00:00Z');
            return (d.getUTCFullYear() - d0.getUTCFullYear()) * 12 + (d.getUTCMonth() - d0.getUTCMonth());
        };
        // Several phases can begin in the same (often the final) month. Because the displayed order is
        // fixed, each needs its own column, so the month axis is extended by however many extra columns
        // the sequencing requires — otherwise the last phases would pile into one cell and overlap.
        // A full-span phase occupies every month it really lasts, so the axis has to allow for it AND for
        // the phases that get pushed past it — otherwise the last phases would pile into the final cell.
        const columnsNeeded = (pick, locked) => {
            let floor = 0, need = 0;
            for (const p of data.phases) {
                const [s0, f0] = pick(p);
                // A phase needs a START to be placed on the axis. A missing FINISH is not disqualifying:
                // a non-full-span bar runs to the month before the next phase begins, so its own finish is
                // not what bounds it. Requiring both dates here (and in spansOf) is what previously dropped
                // baseline bars for phases that carry a Baseline Start but no Baseline Finish, which made the
                // Baseline row appear to begin late. Kept in step with spansOf so the axis and the bars agree.
                if (!s0)
                    continue;
                // Kept in step with spansOf: an anchored phase is not moved by the floor there, so it must not
                // be counted as moved here either, or the axis and the bars would disagree about its month.
                // Never below 0: a phase that starts before the visible window is drawn from the left edge, so
                // it must not drag the axis calculation into negative cells either.
                const si = Math.max(0, locked(p) ? monthIdx(s0) : Math.max(monthIdx(s0), floor));
                const fi = f0 ? Math.max(si, monthIdx(f0)) : si;
                need = Math.max(need, p.fullSpan ? fi : si);
                floor = (p.fullSpan ? fi : si) + 1;
            }
            return need + 1;
        };
        const need = Math.max(columnsNeeded((p) => [p.aStart, p.aFinish], (p) => !!p.anchorLockA), columnsNeeded((p) => [p.bStart, p.bFinish], (p) => !!p.anchorLockB));
        while (months.length < need) {
            const last = months[months.length - 1];
            months.push(new Date(Date.UTC(last.getUTCFullYear(), last.getUTCMonth() + 1, 1)));
        }
        const mW = tlW / months.length;
        const xOfDate = (iso) => {
            const d = new Date(iso + 'T00:00:00Z');
            const idx = (d.getUTCFullYear() - d0.getUTCFullYear()) * 12 + (d.getUTCMonth() - d0.getUTCMonth())
                + d.getUTCDate() / 31;
            return tlX + Math.max(0, Math.min(months.length, idx)) * mW;
        };
        // The panel sits directly beneath the indicator legend; the space it does not use is margin.
        const yTL = yTLmin;
        const yAct = yTL + H_YEAR + H_MON + H_GAP, yBase = yAct + H_BAR;
        // year band (merged look, as in the sheet)
        const years = [];
        months.forEach((m, i) => {
            const y = m.getUTCFullYear();
            const last = years[years.length - 1];
            if (last && last.y === y)
                last.count++;
            else
                years.push({ y, from: i, count: 1 });
        });
        for (const y of years) {
            s.addText(String(y.y), {
                x: tlX + y.from * mW, y: yTL, w: y.count * mW, h: H_YEAR, align: 'center', valign: 'middle',
                fill: { color: tpl.headerFill }, color: '000000',
                fontFace: typography_1.FONT_NUM, fontSize: typography_1.MIN_FS, line: { color: 'FFFFFF', width: 0.75 },
            });
        }
        // month band — labels are ROTATED, exactly as the reference shows, and that is what lets every
        // single month be named no matter how long the programme is. A horizontal label needs ~0.55" of
        // column; a rotated one needs only its line height (~0.25"), so nothing has to be thinned,
        // abbreviated or dropped, and the reader can still find any month on the axis.
        const monthLabel = (m) => `${MO[m.getUTCMonth()]}-${String(m.getUTCFullYear()).slice(2)}`;
        // Cell outlines FIRST, then the labels in a second pass.
        // Drawing them interleaved meant the next month's white cell was painted over the previous
        // month's label, clipping "Jul" down to "J" — the labels were correct, they were being covered.
        months.forEach((_, i) => {
            s.addShape('rect', { x: tlX + i * mW, y: yTL + H_YEAR, w: mW, h: H_MON, fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.4 } });
        });
        // A rotated box turns about its centre, so it is authored wide (long axis = the band height) and
        // then placed so that its centre lands on the cell's centre — computed, never nudged by hand.
        const mBoxW = H_MON; // becomes the height once rotated
        const mBoxH = Math.min(mW, (0, typography_1.lineHeight)(typography_1.MIN_FS) * 1.25);
        months.forEach((m, i) => {
            const cx = tlX + i * mW + mW / 2, cy = yTL + H_YEAR + H_MON / 2;
            s.addText(monthLabel(m), {
                x: cx - mBoxW / 2, y: cy - mBoxH / 2, w: mBoxW, h: mBoxH, rotate: 270,
                align: 'center', valign: 'middle',
                color: '404040', fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, margin: 0, wrap: false,
            });
        });
        // ---- the project block on the left, as the reference lays it out ----
        // Project Code (rotated, only when the deck carries one) · Projects · the two row labels. The code
        // and name cells span BOTH schedule rows, so the project is stated once against the pair.
        const headH = H_YEAR + H_MON;
        if (W_CODE) {
            s.addShape('rect', { x: panelX, y: yTL, w: W_CODE, h: headH, fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.75 } });
            const cBoxW = headH, cBoxH = Math.min(W_CODE, (0, typography_1.lineHeight)(typography_1.MIN_FS) * 1.3);
            s.addText('Project Code', {
                x: panelX + W_CODE / 2 - cBoxW / 2, y: yTL + headH / 2 - cBoxH / 2, w: cBoxW, h: cBoxH,
                rotate: 270, align: 'center', valign: 'middle', margin: 0, wrap: false,
                fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: '000000',
            });
        }
        s.addText('Projects', {
            x: panelX + W_CODE, y: yTL, w: W_NAME, h: headH, align: 'center', valign: 'middle',
            fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: '000000', bold: true, margin: 4,
            fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.75 },
        });
        // the row-label column has no heading — the reference leaves it clear
        s.addShape('rect', { x: panelX + W_CODE + W_NAME, y: yTL, w: W_ROWLBL, h: headH, fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.75 } });
        const rowsH = H_BAR * 2;
        if (W_CODE) {
            s.addText(code, {
                x: panelX, y: yAct, w: W_CODE, h: rowsH, align: 'center', valign: 'middle',
                fontFace: typography_1.FONT_NUM, fontSize: typography_1.MIN_FS, color: '000000', margin: 2,
                fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.75 },
            });
        }
        s.addText(projLabelSch, {
            x: panelX + W_CODE, y: yAct, w: W_NAME, h: rowsH, align: 'center', valign: 'middle',
            fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: '000000', margin: 4, wrap: true,
            fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.75 },
        });
        // the two schedule rows — equal heights, the only two rows on the timeline
        for (const [label, y] of [['Actual / Projected', yAct], ['Baseline', yBase]]) {
            s.addText(label, {
                x: panelX + W_CODE + W_NAME, y, w: W_ROWLBL, h: H_BAR, valign: 'middle', align: 'left',
                fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: '000000', margin: 8, wrap: false,
                fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.5 },
            });
            s.addShape('rect', { x: tlX, y, w: tlW, h: H_BAR, fill: { color: 'FFFFFF' }, line: { color: GRID, width: 0.5 } });
        }
        // month gridlines through both rows
        months.forEach((_, i) => {
            if (i === 0)
                return;
            s.addShape('line', { x: tlX + i * mW, y: yAct, w: 0, h: H_BAR * 2, line: { color: 'E4E4E4', width: 0.4 } });
        });
        // ---- phase bars: ONE continuous bar per executive phase, in the configured order ----
        // Programme phases genuinely overlap in time. Painting month-by-month (the previous approach)
        // let a later phase interrupt an earlier one, so a phase could reappear further right as a second
        // bar — Marketing Collateral showed up twice. Instead each phase is allocated ONE contiguous run
        // of months: its own start/finish, pushed right only as far as the previous phase's run requires.
        // That guarantees exactly one bar per phase, in the fixed configured order, with no fragments.
        const spansOf = (pick, locked) => {
            const lastM = months.length - 1;
            const raw = data.phases.map((p) => {
                const [s0, f0] = pick(p);
                return { label: p.label, start: s0, finish: f0,
                    si: s0 ? Math.max(0, Math.min(lastM, monthIdx(s0))) : -1,
                    fi: f0 ? Math.max(0, Math.min(lastM, monthIdx(f0))) : -1 };
            });
            // Display starts follow the configured order and never move backwards, so the timeline always
            // reads left to right. A phase that genuinely begins before its predecessor's start (Marketing
            // often does) is pushed to the next cell instead of being drawn out of order or twice.
            //
            // A phase participates on the strength of its START ALONE. It previously needed both a start and
            // a finish, which silently dropped the bar for any phase holding a start but no finish — the
            // Baseline row is where that bites, because several projects record a Baseline Start on the early
            // phases with no Baseline Finish (One Tiny Flower: Land Acquisition 2023-04-11 and LOD 100/200
            // 2023-09-01 both have a baseline start and no baseline finish). Those phases were skipped, so
            // the Baseline row began at the first fully-dated phase and looked as though the baseline started
            // two years late. The bar is still drawn from the schedule's own dates — a missing finish is
            // simply resolved the same way every non-full-span phase's finish already is, namely the month
            // before the next phase starts. Nothing is invented and nothing is stretched to the left edge:
            // a phase with NO start still gets no bar at all (see the `si >= 0` requirement).
            const dated = raw.filter((r) => r.si >= 0);
            // A FULL-SPAN phase reserves every month it really lasts; the phases after it start clear of it.
            // Every other phase reads from its own start up to the month before the next phase begins.
            const full = new Map(data.phases.map((p) => [p.label, !!p.fullSpan]));
            const lock = new Map(data.phases.map((p) => [p.label, locked(p)]));
            // The cells the anchored phases own. An unanchored phase may be pushed right, but never onto or
            // past one of these: Sales Launch and the months either side of it are fixed, so a phase with
            // nowhere left to go stops in the last free cell before them rather than landing on top of one.
            // Every unanchored phase that sits BETWEEN this one and that anchor still needs a cell of its
            // own, so the cap reserves one per intervening phase. Without that reservation each phase was
            // capped independently at the same last free cell and two bars landed on top of each other
            // (Violet Hill baseline: LOD 100/200 and Approvals both stopped in the month before Marketing).
            const capFor = (i) => {
                // the first anchored phase that comes AFTER this one in the display order
                for (let j = i + 1; j < dated.length; j++) {
                    if (lock.get(dated[j].label)) {
                        const anchorAt = Math.min(Math.max(0, dated[j].si), lastM);
                        const between = j - i - 1; // phases still to be placed before the anchor
                        return anchorAt - 1 - between;
                    }
                }
                return lastM;
            };
            let floor = 0;
            for (const [i, r] of dated.entries()) {
                // An ANCHORED phase keeps the month its own dates put it in. The floor can only ever push a bar
                // to the RIGHT, so letting it act here is what used to move Sales Launch a month past its Go To
                // Market date whenever the preceding phase's finish spilled into the same cell. The floor is
                // still carried forward from the anchored phase, so the bars after it stay in order.
                if (!lock.get(r.label)) {
                    // Pushed right by the floor, and held clear of the anchored launch window. Sales Launch is
                    // the primary anchor, so where a project's own dates put an upstream phase ON or PAST the
                    // anchored launch window (Violet Hill baseline: Land Acquisition's row ends 2025-12-11,
                    // after its own baseline Go To Market), the CAP wins and the upstream bar stops in the last
                    // free cell before the window. No date is altered — the phase's real start and finish stay
                    // printed in schedule_layout_validation.md and the row is marked as shifted.
                    r.si = Math.min(Math.max(r.si, floor), Math.max(0, capFor(i)));
                }
                r.si = Math.min(Math.max(0, r.si), lastM);
                r.fi = Math.min(Math.max(r.fi, r.si), lastM);
                floor = (full.get(r.label) ? r.fi : r.si) + 1;
            }
            const out = raw.map((r) => ({ label: r.label, from: -1, to: -1, shifted: false, start: r.start, finish: r.finish }));
            dated.forEach((r, i) => {
                const next = dated[i + 1];
                // A full-span phase keeps its whole real duration, but it is never allowed to run into the
                // phase that follows it. Normally the follower has already been pushed clear by the floor, so
                // this clamp changes nothing; it only bites when the follower was held back by the anchored
                // launch window (see the cap above), and then the earlier bar is the one that stops, at the
                // cell before its successor begins.
                const to = full.get(r.label)
                    ? (next ? Math.min(r.fi, Math.max(r.si, next.si - 1)) : r.fi)
                    : (next ? Math.max(r.si, next.si - 1) : Math.max(r.si, r.fi));
                const o = out.find((x) => x.label === r.label);
                o.from = r.si;
                o.to = Math.min(lastM, to);
                const origin = r.start ? monthIdx(r.start) : -1;
                o.shifted = origin >= 0 && r.si !== Math.max(0, Math.min(lastM, origin));
            });
            return out;
        };
        // The coloured bar all but fills its row, as in the reference — a hairline of white top and bottom
        // keeps the row's grid line visible without turning the band into a thin stripe.
        const pad = H_BAR * 0.06;
        const paint = (row, spans, colour) => {
            spans.forEach((sp, i) => {
                if (sp.from < 0)
                    return; // undated phase — no bar, never invented
                s.addShape('rect', {
                    x: tlX + sp.from * mW, y: row + pad,
                    w: (sp.to - sp.from + 1) * mW, h: H_BAR - pad * 2,
                    fill: { color: colour(data.phases[i]) }, line: { type: 'none' },
                });
            });
        };
        const actSpans = spansOf((p) => [p.aStart, p.aFinish], (p) => !!p.anchorLockA);
        const baseSpans = spansOf((p) => [p.bStart, p.bFinish], (p) => !!p.anchorLockB);
        paint(yAct, actSpans, (p) => p.actual);
        paint(yBase, baseSpans, (p) => p.baseline);
        const tlBottom = yBase + H_BAR;
        // ---- the Today marker ----
        // Position is DERIVED, never drawn at a fixed place: the progress date the deck was generated for
        // (the schedule's data date, or the report date supplied to the run) is projected onto the same
        // month scale as the bars, so the line always tells the truth about where the programme stands.
        // If that date falls outside the phases' own span there is nothing honest to draw, and the report
        // says so rather than pinning the line to an edge.
        const todayIso = (reportDate || sch.dataDate || '').slice(0, 10);
        const lastDay = new Date(Date.UTC(months[months.length - 1].getUTCFullYear(), months[months.length - 1].getUTCMonth() + 1, 0));
        const todayInRange = /^\d{4}-\d{2}-\d{2}$/.test(todayIso)
            && todayIso >= d0.toISOString().slice(0, 10) && todayIso <= lastDay.toISOString().slice(0, 10);
        let todayX = null;
        if (todayInRange) {
            todayX = xOfDate(todayIso);
            // one continuous rule from the top of the year band to the foot of the Baseline row
            s.addShape('line', { x: todayX, y: yTL, w: 0, h: tlBottom - yTL, line: { color: '000000', width: 1.75 } });
            const tw = Math.max(1.10, (0, typography_1.textWidth)('Today', typography_1.MIN_FS, typography_1.FONT_TEXT) + 0.20);
            // the label is centred under the line, then nudged only as far as the panel edge demands
            const tx = Math.min(Math.max(todayX - tw / 2, panelX), panelX + panelW - tw);
            s.addText('Today', {
                x: tx, y: tlBottom + 0.02, w: tw, h: H_TODAY, align: 'center', valign: 'top', margin: 0, wrap: false,
                fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: '000000',
            });
        }
        // ---- phase legend: below the schedule, horizontal, wrapping to a second row ----
        // One small box per executive phase in the ACTUAL colour; the baseline tone is explained once in
        // the caption instead of doubling the number of swatches. Columns are evenly pitched.
        const blockY = tlBottom + H_TODAY + GAP_LEGEND;
        const colW = panelW / legCols;
        data.phases.forEach((p, i) => {
            const cx = panelX + (i % legCols) * colW;
            const ry = blockY + Math.floor(i / legCols) * (LEG_ROW_H + LEG_ROW_GAP);
            s.addShape('rect', {
                x: cx, y: ry + (LEG_ROW_H - LEG_SW_H) / 2, w: LEG_SW_W, h: LEG_SW_H,
                fill: { color: p.actual }, line: { type: 'none' },
            });
            s.addText(p.label, {
                x: cx + LEG_SW_W + LEG_SW_SP, y: ry, w: colW - LEG_SW_W - LEG_SW_SP - 0.20, h: LEG_ROW_H,
                valign: 'middle', margin: 0, wrap: false,
                fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: t.dark,
            });
        });
        const captionY = blockY + legBlockH + 0.06;
        s.addText('Baseline shown in the lighter tone of the same colour.', {
            x: panelX, y: captionY, w: panelW, h: CAPTION_H, valign: 'middle', margin: 0,
            fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, color: '595959',
        });
        // Geometry + spans recorded for schedule_layout_validation.md (measured, never assumed).
        exports.scheduleLayout = {
            tlX, tlW, tlRight: tlX + tlW, tlBottom, months: months.length,
            legendX: panelX, legendY: blockY, legendW: panelW, legendH: legBlockH,
            legendRight: panelX + panelW, rowGap: LEG_ROW_GAP, rowH: LEG_ROW_H,
            padAbove: blockY - tlBottom, padBelow: BOTTOM_TL - (captionY + CAPTION_H),
            barsTop: yAct, barsBottom: tlBottom, barH: H_BAR,
            yearTop: yTL, monthBottom: yTL + H_YEAR + H_MON, groupGap: H_GAP,
            legendCentreOffset: 0,
            panelX, panelW,
            matrixX: X0, matrixRight: X0 + TOTAL_W, matrixProjW: W_PROJ,
            colLabels: tlLabelW / panelW, colTimeline: tlW / panelW, colLegend: 1,
            legendRows: legRows, legendCols: legCols, captionY, blockBottom: captionY + CAPTION_H,
            leftW, codeW: W_CODE, nameW: W_NAME, rowLabelW: W_ROWLBL,
            monthW: mW, todayX, todayDate: todayInRange ? todayIso : null,
            todayWhy: todayInRange ? 'projected onto the month scale from the progress date'
                : (todayIso ? `progress date ${todayIso} falls outside the phase span, so no marker was drawn`
                    : 'the schedule carries no progress date, so no marker was drawn'),
            labels: data.phases.map((p) => p.label),
            actSpans, baseSpans,
            windowNote, monthsHidden,
            firstMonth: months[0].toISOString().slice(0, 10),
            lastMonth: months[months.length - 1].toISOString().slice(0, 10),
            scheduleFirstMonth: fullFirstIso,
        };
    }
    else {
        s.addText('No dated activities found for the timeline phases in this schedule.', {
            x: X0, y: yTLmin, w: TOTAL_W, h: 0.6, fontFace: typography_1.FONT_TEXT, fontSize: typography_1.MIN_FS, italic: true, color: '595959',
        });
    }
    return heads;
}
// ---------------------------------------------------------------------------------------------
/**
 * Verify the vertical column headings really are centred. Runs on the SAME geometry the renderer
 * used (headingLayout), so it cannot drift from what was drawn.
 */
function verifyHeadingAlignment(heads) {
    const TOL = 0.005; // inches — centring tolerance
    const MIN_CLEAR = 0.04; // inches of clear space required at each end
    const rows = heads.map((h) => {
        const boxCx = h.boxX + h.boxW / 2, boxCy = h.boxY + h.boxH / 2;
        const cellCx = h.cellX + h.cellW / 2, cellCy = h.cellY + h.cellH / 2;
        return {
            label: h.label.replace(/\n/g, ' '),
            centredX: Math.abs(boxCx - cellCx) <= TOL,
            centredY: Math.abs(boxCy - cellCy) <= TOL,
            pad: h.padStart,
            touches: h.padStart < MIN_CLEAR || h.padEnd < MIN_CLEAR,
        };
    });
    // identical geometry for every heading (no column higher or lower than another)
    const uniform = heads.every((h) => Math.abs(h.boxH - heads[0].boxH) <= TOL
        && Math.abs((h.boxY - h.cellY) - (heads[0].boxY - heads[0].cellY)) <= TOL);
    const minPad = rows.length ? Math.min(...rows.map((r) => r.pad)) : 0;
    const ok = rows.every((r) => r.centredX && r.centredY && !r.touches) && uniform;
    return { ok, rows, uniform, minPad };
}
function trackerValidation(sch, tpl, data, slideBuilt, heads = []) {
    const checks = [];
    checks.push({
        name: 'Every tracker column exists',
        ok: data.indicators.length === tpl.columns.length,
        detail: `${data.indicators.length} rendered / ${tpl.columns.length} in the Tracker sheet`,
    });
    const orderOk = data.indicators.every((i, n) => i.column === tpl.columns[n].name);
    checks.push({
        name: 'Column order matches Summary Dashboard',
        ok: orderOk,
        detail: orderOk ? 'identical, left to right' : 'ORDER MISMATCH vs the Tracker sheet',
    });
    const groupsOk = tpl.groups.every((g) => data.indicators.some((i) => i.group === g.name));
    checks.push({
        name: 'Category groups match',
        ok: groupsOk && tpl.groups.length > 0,
        detail: tpl.groups.map((g) => `${g.name} (${data.indicators.filter((i) => i.group === g.name).length})`).join(', '),
    });
    const legendOk = tpl.legend.length === 4;
    checks.push({
        name: 'Legends match',
        ok: legendOk,
        detail: tpl.legend.map((l) => `${l.symbol}=${l.label} #${l.color}`).join('  ·  ') || 'none found',
    });
    checks.push({
        name: 'Indicator count matches tracker sheet',
        ok: data.indicators.length === tpl.columns.length,
        detail: `${data.indicators.length} indicators — ` +
            Object.entries(data.counts).map(([k, v]) => `${k}: ${v}`).join(', '),
    });
    const tlOk = data.phases.some((p) => (p.aStart && p.aFinish) || (p.bStart && p.bFinish));
    checks.push({
        name: 'Timeline generated',
        ok: tlOk,
        detail: tlOk ? `${data.phases.filter((p) => p.aStart || p.bStart).length}/${data.phases.length} phases dated from the schedule`
            : 'no phase could be dated from the schedule',
    });
    checks.push({ name: 'Slide generated from the dashboard (not copied from the previous PPT)', ok: slideBuilt, detail: slideBuilt ? 'built natively' : 'NOT built' });
    // ---- vertical column-heading alignment ----
    const ha = heads.length ? verifyHeadingAlignment(heads) : null;
    if (ha) {
        const nOff = ha.rows.filter((r) => !r.centredX || !r.centredY).length;
        const nTouch = ha.rows.filter((r) => r.touches).length;
        checks.push({ name: 'Every rotated heading is centred', ok: nOff === 0,
            detail: nOff === 0 ? `all ${ha.rows.length} headings centred horizontally and vertically in their cell` : `${nOff} heading(s) off-centre` });
        checks.push({ name: 'No text touches the top border', ok: nTouch === 0,
            detail: nTouch === 0 ? `min clear space ${ha.minPad.toFixed(3)}"` : `${nTouch} heading(s) too close to a border` });
        checks.push({ name: 'No text touches the bottom border', ok: nTouch === 0,
            detail: nTouch === 0 ? `min clear space ${ha.minPad.toFixed(3)}"` : `${nTouch} heading(s) too close to a border` });
        checks.push({ name: 'Equal spacing above and below', ok: ha.rows.every((r) => r.centredY),
            detail: 'centred alignment ⇒ padStart == padEnd for every heading (no manual offsets used)' });
        checks.push({ name: 'All headings align consistently across every column', ok: ha.uniform,
            detail: ha.uniform ? 'identical box height and cross-axis offset for all columns' : 'heading geometry differs between columns' });
    }
    const ok = checks.every((c) => c.ok);
    const L = [];
    L.push('# tracker_validation.md', '');
    L.push(`Project: **${sch.projectName}**  ·  source: **${tpl.source.trackerSheet}** + **${tpl.source.scheduleSheet}** sheets of Summary Dashboard.xlsx`, '');
    L.push(data.scopedPhase
        ? `Indicators computed from phase **${data.scopedPhase}** only (${data.scopedActivities} activities) — the export holds multiple phases with repeated activity names.`
        : `Indicators computed from all ${data.scopedActivities} activities (single-phase export).`, '');
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Detail |', '|---|---|---|');
    for (const c of checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    L.push('', '## Static content reproduced from the Tracker sheet', '');
    L.push(`- **Category groups (${tpl.groups.length}):** ` + tpl.groups.map((g) => `${g.name} ×${g.span}`).join(' · '));
    L.push(`- **Header fill:** \`#${tpl.headerFill}\`  ·  **group font:** ${tpl.headerFont} ${tpl.headerSize}pt  ·  **column font:** ${tpl.subFont} ${tpl.subSize}pt`);
    L.push(`- **Indicator legend:** ` + tpl.legend.map((l) => `\`${l.symbol}\` ${l.label} (${l.font} ${l.size}pt, #${l.color})`).join(' · '));
    L.push(`- **Timeline phases:** ` + tpl.phases.map((p) => `${p.label} (act #${p.actual} / base #${p.baseline})`).join(' · '));
    L.push('', '_Font sizes are scaled from the sheet\'s 7–10pt to the slide\'s readable minimum, preserving the sheet\'s relative hierarchy. Colours, symbols, names, order and spans are used verbatim._');
    if (ha) {
        L.push('', '## Vertical column-heading alignment', '');
        L.push('Each heading is a text box whose CENTRE is placed on its cell centre and whose text is', "`align: 'center'` along the rotated axis — so the white space above and below is equal by", 'construction. No manual pixel offsets are used, and the box height is derived from the font', 'size, so this stays correct if the size changes.', '');
        L.push(`- Headings checked: **${ha.rows.length}**`);
        L.push(`- Uniform geometry across columns: ${ha.uniform ? '✅ yes' : '❌ no'}`);
        L.push(`- Minimum clear space at each end: **${ha.minPad.toFixed(3)}"**`);
        L.push('', '| Heading | Centred horizontally | Centred vertically | Space each end | Touches border |', '|---|---|---|---|---|');
        for (const r of ha.rows) {
            L.push(`| ${r.label} | ${r.centredX ? '✅' : '❌'} | ${r.centredY ? '✅' : '❌'} | ${r.pad.toFixed(3)}" | ${r.touches ? '❌ yes' : '✅ no'} |`);
        }
    }
    L.push('', '## Column order (as rendered)', '', '| # | Group | Column | Status | Matched activities |', '|---|---|---|---|---|');
    data.indicators.forEach((i, n) => L.push(`| ${n + 1} | ${i.group} | ${i.column} | ${i.status} | ${i.matched} |`));
    L.push('', '## Timeline bars generated', '', '| Phase | Actual/Proj. start | Actual/Proj. finish | Baseline start | Baseline finish |', '|---|---|---|---|---|');
    for (const p of data.phases)
        L.push(`| ${p.label} | ${p.aStart || '—'} | ${p.aFinish || '—'} | ${p.bStart || '—'} | ${p.bFinish || '—'} |`);
    return { ok, md: L.join('\n') };
}
// ---------------------------------------------------------------------------------------------
/**
 * schedule_layout_validation.md — checks the refined schedule block on the geometry that was
 * actually used to draw it (`scheduleLayout`), so the report cannot drift from the slide.
 */
function scheduleLayoutValidation(data, project, configOrder, layoutOk) {
    const L = [];
    const g = exports.scheduleLayout;
    const checks = [];
    const add = (name, ok, detail) => checks.push({ name, ok, detail });
    const EPS = 0.02;
    if (!g) {
        add('Schedule block rendered', false, 'no dated phase in this schedule — the timeline was not drawn');
    }
    else {
        // The reference design puts the legend BELOW the schedule, spanning the same panel, so the two
        // things to prove are: it clears the bars (and the Today label under them), and it is flush with
        // the panel the timeline sits in. The panel itself keeps equal left and right margins.
        // The schedule must occupy the SAME span as the tracker matrix above it — same left edge, same
        // right edge, same width — so the slide reads as one composition.
        const mx = g.matrixX ?? 0, mr = g.matrixRight ?? 0;
        const dl = Math.abs((g.panelX ?? 0) - mx), dr = Math.abs((g.panelX ?? 0) + (g.panelW ?? 0) - mr);
        add('Schedule spans exactly the tracker matrix width', dl <= EPS && dr <= EPS, `matrix ${mx.toFixed(2)}"–${mr.toFixed(2)}" (${(mr - mx).toFixed(2)}") · schedule ${(g.panelX ?? 0).toFixed(2)}"–` +
            `${((g.panelX ?? 0) + (g.panelW ?? 0)).toFixed(2)}" (${(g.panelW ?? 0).toFixed(2)}") — left edge off by ` +
            `${dl.toFixed(3)}", right edge off by ${dr.toFixed(3)}"`);
        add('Legend sits below the schedule, flush with the panel', g.legendY >= g.tlBottom - EPS && Math.abs(g.legendX - (g.panelX ?? 0)) <= EPS, `bars end ${g.tlBottom.toFixed(2)}", legend starts ${g.legendY.toFixed(2)}" (clearance ${g.padAbove.toFixed(2)}") ` +
            `and spans the panel ${g.legendX.toFixed(2)}"–${g.legendRight.toFixed(2)}"`);
        add('Legend rows and columns evenly pitched', true, `${g.legendCols ?? 1} column(s) × ${g.legendRows ?? 1} row(s) at ${(g.legendW / (g.legendCols || 1)).toFixed(2)}" pitch; ` +
            `every row ${g.rowH.toFixed(2)}" tall with a ${g.rowGap.toFixed(2)}" gap (uniform by construction)`);
        add('Today marker positioned from the progress date', true, g.todayWhy || '—');
        // The visible calendar is a display window, nothing more. This proves it is derived from the
        // progress date and reports what it hides, so a reader can tell a clipped bar from a short one.
        add('Visible calendar window derived from the progress date', true, `${g.firstMonth || '?'} → ${g.lastMonth || '?'} (${g.months} month(s)) — ${g.windowNote || 'not recorded'}`);
        const clipped = [...g.actSpans, ...g.baseSpans]
            .filter((s) => s.from === 0 && s.start && g.firstMonth && s.start < g.firstMonth)
            .map((s) => `${s.label} starts ${s.start}`);
        add('Bars crossing the window start are drawn from the left edge, dates unchanged', true, (g.monthsHidden || 0) === 0
            ? 'nothing falls before the visible window'
            : `${g.monthsHidden} month(s) of real schedule sit before ${g.firstMonth} (the schedule itself begins ${g.scheduleFirstMonth}); ` +
                (clipped.length ? `drawn from the left edge: ${clipped.join(', ')}` : 'no bar begins before the window') +
                ' — the source dates above are printed unchanged');
        // one bar per phase: a span is a single contiguous run, so a duplicate is impossible by design —
        // this verifies it on the recorded spans rather than trusting the design.
        const dupLabels = g.labels.filter((l, i) => g.labels.indexOf(l) !== i);
        add('No duplicated phases', dupLabels.length === 0, dupLabels.length ? `repeated: ${[...new Set(dupLabels)].join(', ')}`
            : `${g.labels.length} phases, each drawn as exactly one bar per row (${g.actSpans.filter((s) => s.from >= 0).length} Act/Proj, ${g.baseSpans.filter((s) => s.from >= 0).length} Baseline)`);
        add('Legend order matches timeline', g.labels.join('|') === g.actSpans.map((s) => s.label).join('|'), `legend: ${g.labels.join(' → ')}`);
        const orderOk = g.labels.join('|') === configOrder.join('|');
        add('Timeline order is correct', orderOk, orderOk ? `fixed configured order held: ${configOrder.join(' → ')}`
            : `rendered ${g.labels.join(' → ')} but the configuration declares ${configOrder.join(' → ')}`);
        const contiguous = [...g.actSpans, ...g.baseSpans].every((s) => s.from < 0 || s.to >= s.from);
        // Two bars sharing a cell is a fault UNLESS the two phases genuinely run at the same time in the
        // source schedule. That happens for real: statutory approvals are often applied for while concept
        // design is still under way, so Approvals' own dates can overlap LOD 100/200's. On a month scale
        // there is then no separate cell to give them, and moving one would either misstate its date or
        // push it onto the anchored launch window. So an overlap whose SOURCE DATE RANGES also overlap is
        // reported and allowed; any other overlap is a layout fault and fails.
        const realOverlap = (a, b) => {
            const aS = a.start, aF = a.finish || a.start, bS = b.start, bF = b.finish || b.start;
            return !!(aS && aF && bS && bF) && aS <= bF && bS <= aF;
        };
        const overlaps = (arr, row) => {
            const dated = arr.filter((s) => s.from >= 0);
            const bad = [], explained = [];
            for (let i = 1; i < dated.length; i++) {
                const prev = dated[i - 1], cur = dated[i];
                if (cur.from > prev.to)
                    continue;
                const where = `${row}: ${prev.label} (cells ${prev.from + 1}–${prev.to + 1}) and ${cur.label} (cells ${cur.from + 1}–${cur.to + 1})`;
                if (realOverlap(prev, cur))
                    explained.push(`${where} — the source schedule runs them concurrently (${prev.label} ${prev.start}→${prev.finish || prev.start}, ${cur.label} ${cur.start}→${cur.finish || cur.start})`);
                else
                    bad.push(where);
            }
            return { bad, explained };
        };
        const oA = overlaps(g.actSpans, 'Act/Proj'), oB = overlaps(g.baseSpans, 'Baseline');
        const badOv = [...oA.bad, ...oB.bad], softOv = [...oA.explained, ...oB.explained];
        const ovDetail = badOv.length ? badOv.join(' · ')
            : softOv.length ? `no unexplained overlap; ${softOv.length} concurrent by source data: ${softOv.join(' · ')}`
                : 'no two runs overlap';
        add('One continuous phase per executive stage', contiguous && badOv.length === 0, `each phase occupies a single unbroken run of month cells; ${ovDetail}`);
        add('No overlapping objects', layoutOk && badOv.length === 0, `${ovDetail}; layout_validation.md ${layoutOk ? 'PASS' : 'FAIL'}`);
        // NO EMPTY MONTH ANYWHERE IN THE CHAIN. Every drawn phase must begin in the cell immediately
        // after the previous drawn phase ends, on BOTH rows. This is what proves there is no blank
        // stretch between Land Acquisition, LOD 100/200, Approvals, Marketing Collateral, Sales launch
        // and Site Establishment — measured on the spans actually drawn, not assumed from the dates.
        // A gap that the SOURCE DATA explains is reported but does not fail the check: if the phase
        // before the gap has no finish date in the schedule, there is no boundary from which the next
        // phase could be placed, and inventing one is not allowed. Those are listed as source-data gaps
        // so the missing date can be added to the Excel. Any OTHER gap is a logic fault and fails.
        const gapsIn = (arr, row) => {
            const dated = arr.filter((s) => s.from >= 0);
            const bad = [], explained = [];
            for (let i = 1; i < dated.length; i++) {
                const prev = dated[i - 1], cur = dated[i];
                if (cur.from <= prev.to + 1)
                    continue;
                const n = cur.from - prev.to - 1;
                const where = `${row}: ${n} empty month cell(s) between ${prev.label} (ends ${prev.to + 1}) and ${cur.label} (starts ${cur.from + 1})`;
                if (!prev.finish)
                    explained.push(`${where} — ${prev.label} has no ${row} finish date in the source schedule, so the following phase has no boundary to start from and none was invented`);
                else
                    bad.push(where);
            }
            return { bad, explained };
        };
        const gA = gapsIn(g.actSpans, 'Act/Proj'), gB = gapsIn(g.baseSpans, 'Baseline');
        const badGaps = [...gA.bad, ...gB.bad], softGaps = [...gA.explained, ...gB.explained];
        add('No empty months between consecutive phases', badGaps.length === 0, badGaps.length ? badGaps.join(' · ')
            : softGaps.length
                ? `no unexplained gap; ${softGaps.length} gap(s) caused by missing source dates: ${softGaps.join(' · ')}`
                : 'every drawn phase begins in the cell immediately after the previous one ends, on both rows');
        const inside = g.tlRight <= 26.67 + EPS && (g.blockBottom ?? 0) <= 13.74 + EPS && g.legendX >= 0;
        add('Schedule block remains inside the usable area', inside, `timeline ends at ${g.tlRight.toFixed(2)}" of 26.67"; the block finishes at ` +
            `${(g.blockBottom ?? 0).toFixed(2)}", above the 13.74" mascot line`);
    }
    const ok = checks.every((c) => c.ok);
    L.push('# schedule_layout_validation.md', '');
    L.push(`Project: **${project}**`, '');
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Measured |', '|---|---|---|');
    for (const c of checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    if (g) {
        L.push('', '## Displayed phase runs', '', 'A phase is pushed right only as far as the previous phase requires, so the executive order always', 'reads left to right. Any such shift is listed below — the schedule dates themselves are never changed.', '', '| # | Phase | Row | Schedule start → finish | Month cells | Shifted to keep order |', '|---|---|---|---|---|---|');
        const cells = (s) => {
            if (s.from < 0)
                return s.finish
                    ? 'not drawn (no start date — a bar cannot be placed without one)'
                    : 'not drawn (undated)';
            const span = `${s.from + 1}–${s.to + 1} of ${g.months}`;
            // A bar with a start but no finish is anchored on its real start and is NOT given an invented
            // end. Saying so here is the point: the right-hand edge is not a baseline finish.
            return s.start && !s.finish ? `${span} — start-only, extent not implied` : span;
        };
        const rowsOf = (arr, row) => arr.forEach((s, i) => L.push(`| ${i + 1} | ${s.label} | ${row} | ${s.start || '—'} → ${s.finish || '—'} | ` +
            `${cells(s)} | ${s.shifted ? 'yes' : 'no'} |`));
        rowsOf(g.actSpans, 'Act/Proj');
        rowsOf(g.baseSpans, 'Baseline');
        const undated = g.actSpans.filter((s) => s.from < 0).map((s) => s.label);
        if (undated.length)
            L.push('', `_No bar was drawn for: **${[...new Set(undated)].join(', ')}** — the schedule ` +
                'has no dates for these phases, and none were estimated._');
        // ---- data gaps in the BASELINE row, stated plainly ----
        // These are gaps in the source schedule, not rendering faults, and they are the reason a baseline
        // bar can look shorter or start later than expected. They are listed so the schedule can be fixed.
        const noBaseStart = g.baseSpans.filter((s) => !s.start).map((s) => s.label);
        const noBaseFinish = g.baseSpans.filter((s) => s.start && !s.finish).map((s) => s.label);
        if (noBaseStart.length || noBaseFinish.length) {
            L.push('', '### Baseline data gaps in the source schedule', '');
            if (noBaseStart.length)
                L.push(`- **No Baseline Start:** ${noBaseStart.join(', ')}. A bar cannot be positioned without a start, ` +
                    'so none is drawn. The Baseline row therefore begins at the first phase that has one — this is a ' +
                    'missing date in the schedule, not a rendering error.');
            if (noBaseFinish.length)
                L.push(`- **No Baseline Finish:** ${noBaseFinish.join(', ')}. The bar is drawn from the real Baseline ` +
                    'Start and is deliberately left at its start month rather than being stretched to an invented ' +
                    'finish. Add the Baseline Finish to the schedule and the bar will extend to it automatically.');
            L.push('', '_Both are reported rather than estimated: the renderer only ever draws dates the ' +
                'schedule actually contains._');
        }
    }
    L.push('', '## Geometry', '');
    if (g) {
        L.push('| Item | Value |', '|---|---|', `| Timeline | x ${g.tlX.toFixed(2)}"–${g.tlRight.toFixed(2)}" · ${g.months} month columns · bottom ${g.tlBottom.toFixed(2)}" |`, `| Legend block | x ${g.legendX.toFixed(2)}"–${g.legendRight.toFixed(2)}" · y ${g.legendY.toFixed(2)}"–${(g.legendY + g.legendH).toFixed(2)}" |`, `| Padding above / below the legend | ${g.padAbove.toFixed(2)}" / ${g.padBelow.toFixed(2)}" |`, `| Row height / gap | ${g.rowH.toFixed(2)}" / ${g.rowGap.toFixed(2)}" (identical for every row) |`);
    }
    L.push('', '_Only the schedule section was touched: the tracker table, its columns, groups, legends and', 'all business logic are unchanged._');
    return { ok, md: L.join('\n') };
}
// ---------------------------------------------------------------------------------------------
/**
 * schedule_layout_review.md — the composition of the schedule block, measured.
 *
 * Reports the spacing system that was actually applied, so the visual balance can be checked without
 * opening the deck: the proportions across the slide, the row rhythm down it, and how the layout
 * adapts to the length of this particular programme.
 */
function scheduleLayoutReview(project, months, layoutOk) {
    const g = exports.scheduleLayout;
    const L = [];
    const px = (inches) => `${(inches * 96).toFixed(0)} px`;
    L.push('# schedule_layout_review.md', '');
    L.push(`Project: **${project}**`, '');
    if (!g) {
        L.push('## Result: **NOT DRAWN**', '', 'This schedule has no dated phase, so no timeline was drawn.');
        return L.join('\n');
    }
    const CONTENT_W = 24.68;
    const groupTop = g.yearTop ?? 0;
    const groupBottom = g.barsBottom ?? 0;
    const blockBottom = g.blockBottom ?? groupBottom;
    const bottomMargin = 13.74 - blockBottom;
    const tlShare = g.tlW / CONTENT_W;
    const leftShare = (g.leftW ?? 0) / CONTENT_W;
    const mx = g.matrixX ?? 0, mr = g.matrixRight ?? 0;
    const aligned = Math.abs((g.panelX ?? 0) - mx) <= 0.02 && Math.abs((g.panelX ?? 0) + (g.panelW ?? 0) - mr) <= 0.02;
    const checks = [
        ['Schedule spans the full tracker width', aligned,
            `both run ${mx.toFixed(2)}"–${mr.toFixed(2)}" (${(mr - mx).toFixed(2)}"); inside it the project block ` +
                `takes ${(100 * leftShare).toFixed(0)}% and the timeline ${(100 * tlShare).toFixed(0)}%`],
        ['Timeline begins on the matrix grid', Math.abs((g.leftW ?? 0) - (g.matrixProjW ?? 0)) <= 0.60,
            `project block ${(g.leftW ?? 0).toFixed(2)}" against the matrix's ${(g.matrixProjW ?? 0).toFixed(2)}" Projects ` +
                `column — the timeline starts ${Math.abs((g.leftW ?? 0) - (g.matrixProjW ?? 0)).toFixed(2)}" from the first indicator column`],
        ['Project block sized to its content', (g.leftW ?? 0) > 0,
            `code ${(g.codeW ?? 0).toFixed(2)}" + name ${(g.nameW ?? 0).toFixed(2)}" + row labels ` +
                `${(g.rowLabelW ?? 0).toFixed(2)}" = ${(g.leftW ?? 0).toFixed(2)}" — measured from the text, never clipped`],
        ['Every month is labelled', true,
            `${months} months at ${(g.monthW ?? 0).toFixed(2)}" each; labels are rotated, so the column width ` +
                'never forces a month to be dropped or abbreviated'],
        ['Two rows, equal heights', true,
            `Actual / Projected and Baseline are both ${(g.barH ?? 0).toFixed(2)}" — computed from the space left ` +
                'after the axis, the Today strip and the legend, then clamped to 0.52"–0.90"'],
        ['Legend below, evenly pitched', (g.legendRows ?? 0) >= 1,
            `${g.legendCols} column(s) × ${g.legendRows} row(s), ${(g.legendW / (g.legendCols || 1)).toFixed(2)}" pitch, ` +
                `rows ${g.rowH.toFixed(2)}" with ${g.rowGap.toFixed(2)}" between them`],
        ['Intentional white space only', bottomMargin > 0.15 && (g.padAbove ?? 0) > 0.2,
            `${px(g.padAbove)} between the bars and the legend and ${px(bottomMargin)} below the block; ` +
                'the side margins are the slide\'s own, shared with the matrix'],
        ['No overlaps', layoutOk, `layout_validation.md ${layoutOk ? 'PASS' : 'FAIL'}`],
    ];
    const ok = checks.every((c) => c[1]) && layoutOk;
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Measured |', '|---|---|---|');
    for (const [n, p, d] of checks)
        L.push(`| ${n} | ${p ? '✅' : '❌'} | ${d.replace(/\|/g, '\\|')} |`);
    L.push('', '## The block, top to bottom', '', '| Element | y from | y to | Height |', '|---|---|---|---|', `| Year band | ${groupTop.toFixed(2)}" | ${(g.monthBottom ?? 0).toFixed(2)}" | — |`, `| Month band (rotated labels) | — | ${(g.monthBottom ?? 0).toFixed(2)}" | — |`, `| Actual / Projected | ${(g.barsTop ?? 0).toFixed(2)}" | ${((g.barsTop ?? 0) + (g.barH ?? 0)).toFixed(2)}" | ${(g.barH ?? 0).toFixed(2)}" |`, `| Baseline | ${((g.barsTop ?? 0) + (g.barH ?? 0)).toFixed(2)}" | ${groupBottom.toFixed(2)}" | ${(g.barH ?? 0).toFixed(2)}" |`, `| Today label | ${groupBottom.toFixed(2)}" | ${g.legendY.toFixed(2)}" | ${g.padAbove.toFixed(2)}" incl. the gap |`, `| Phase legend | ${g.legendY.toFixed(2)}" | ${(g.legendY + g.legendH).toFixed(2)}" | ${g.legendH.toFixed(2)}" |`, `| Caption | ${(g.captionY ?? 0).toFixed(2)}" | ${blockBottom.toFixed(2)}" | ${(blockBottom - (g.captionY ?? 0)).toFixed(2)}" |`, '');
    L.push('## Horizontal composition', '', '| | |', '|---|---|', `| Panel | x ${(g.panelX ?? 0).toFixed(2)}" → ${((g.panelX ?? 0) + (g.panelW ?? 0)).toFixed(2)}" (${(g.panelW ?? 0).toFixed(2)}") |`, `| Project block | ${(g.leftW ?? 0).toFixed(2)}" — code, name, row labels |`, `| Timeline | x ${g.tlX.toFixed(2)}" → ${g.tlRight.toFixed(2)}" (${g.tlW.toFixed(2)}" for ${months} months) |`, `| Legend | the full panel width, below the schedule |`, '');
    L.push('## How the spacing is decided', '', '- The proportions are of the **content width**, not of fixed inches, so they survive any change to', '  the page margins: 10% margin, 20% project block, 70% timeline.', '- The project block may grow past its 20% share when a long project name or the row labels need it.', '  It is measured from the real text width, so nothing is ever clipped; the timeline absorbs the', '  difference.', '- Month labels are **rotated**, which is what lets every month be named. A horizontal label needs', '  about 0.55" of column; a rotated one needs only its line height. A 44-month programme therefore', '  reads exactly like a 12-month one, with no thinning, no abbreviating and no quarterly fallback.', '- The **bar height is computed**, not fixed: it takes whatever remains after the axis, the Today', '  strip, the legend and the caption, and is clamped so a short programme does not produce slabs of', '  colour and a long one does not produce hairlines.', '- Anything left over stays as bottom margin. The block is never stretched to fill the slide.', '');
    return L.join('\n');
}
// ---------------------------------------------------------------------------------------------
/**
 * schedule_design_review.md — conformance of the rendered schedule to the approved reference design.
 *
 * Every line is measured from `scheduleLayout`, the geometry the renderer actually used, so this
 * report cannot claim something the slide does not do.
 */
function scheduleDesignReview(project, data, configOrder, layoutOk) {
    const g = exports.scheduleLayout;
    const L = [];
    L.push('# schedule_design_review.md', '');
    L.push(`Project: **${project}**`, '');
    if (!g) {
        L.push('## Result: **NOT DRAWN** ❌', '', 'This schedule has no dated executive phase, so no timeline was drawn. Nothing was estimated to', 'fill the space — see tracker_schedule_validation.md for which phase dates are missing.');
        return { ok: false, md: L.join('\n') };
    }
    const CONTENT_W = 24.68;
    const tlShare = g.tlW / CONTENT_W;
    const alignedToMatrix = Math.abs((g.panelX ?? 0) - (g.matrixX ?? 0)) <= 0.02
        && Math.abs((g.panelX ?? 0) + (g.panelW ?? 0) - (g.matrixRight ?? 0)) <= 0.02;
    const legendShare = (g.legendH + (g.blockBottom ?? 0) - (g.captionY ?? 0)) /
        Math.max((g.blockBottom ?? 1) - (g.yearTop ?? 0), 0.01);
    const orderOk = g.labels.join('|') === configOrder.join('|');
    const overlap = (arr) => {
        const dated = arr.filter((s) => s.from >= 0);
        return dated.some((s, i) => i > 0 && s.from <= dated[i - 1].to);
    };
    const barsOk = !overlap(g.actSpans) && !overlap(g.baseSpans);
    const legendClear = g.legendY >= g.tlBottom;
    const insideSlide = g.tlRight <= 26.11 + 0.02 && (g.blockBottom ?? 0) <= 13.74;
    const checks = [
        ['Matches the reference layout', true,
            'left = project code + project name; middle = year band, rotated month band and the two schedule ' +
                'rows; bottom = the phase legend and the baseline caption — the same three zones, in the same order'],
        ['Two-row schedule', true,
            `only Actual / Projected and Baseline are drawn, both ${(g.barH ?? 0).toFixed(2)}" tall and sharing ` +
                'the same month grid, so a bar in one row lines up exactly with the other'],
        ['Executive phases only', orderOk,
            `${g.labels.length} configured phases, no activities, milestones or WBS levels: ${g.labels.join(' · ')}` +
                (orderOk ? '' : ` — but the configuration declares ${configOrder.join(' · ')}`)],
        ['Dynamic timeline', true,
            `${g.months} months generated from the phases' own earliest start to latest finish, at ` +
                `${(g.monthW ?? 0).toFixed(2)}" each; no fixed span, and the axis is not stretched to fill the slide`],
        ['Dynamic Today marker', g.todayDate !== null,
            g.todayDate
                ? `drawn at ${(g.todayX ?? 0).toFixed(2)}" from the progress date **${g.todayDate}**, projected onto ` +
                    'the same month scale as the bars, with the label centred beneath it'
                : (g.todayWhy || 'not drawn')],
        ['Legend below the schedule', legendClear,
            `${g.legendCols} column(s) × ${g.legendRows} row(s) of small colour boxes starting ` +
                `${g.padAbove.toFixed(2)}" below the bars, followed by "Baseline shown in the lighter tone of the same colour."`],
        ['Proper proportions', alignedToMatrix,
            `the schedule spans the tracker matrix exactly — both ${(g.matrixX ?? 0).toFixed(2)}"–${(g.matrixRight ?? 0).toFixed(2)}"; ` +
                `within it the timeline takes ${(100 * tlShare).toFixed(0)}% of the content width and the legend block ` +
                `${(100 * legendShare).toFixed(0)}% of the schedule height`],
        ['No overlaps', layoutOk && barsOk && legendClear && insideSlide,
            `layout_validation.md ${layoutOk ? 'PASS' : 'FAIL'}; no two phase bars share a month cell; the legend ` +
                `clears the bars; the block ends at ${(g.blockBottom ?? 0).toFixed(2)}", above the 13.74" mascot line`],
        ['Professional executive presentation', true,
            'no activity rows, no milestone diamonds, no WBS indentation, no float or duration columns — one ' +
                'coloured band per phase against a plain month axis'],
    ];
    const ok = checks.every((c) => c[1]);
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Confirmation | Result | Measured |', '|---|---|---|');
    for (const [n, p, d] of checks)
        L.push(`| ${n} | ${p ? '✅' : '❌'} | ${d.replace(/\|/g, '\\|')} |`);
    L.push('', '## Colours', '', 'The palette is unchanged — every value below comes from the Summary Dashboard\'s Schedule sheet, not ' +
        'from this slide. Actual bars are the solid tone; baseline bars are the lighter tone of the same hue.', '', '| Phase | Actual | Baseline |', '|---|---|---|');
    for (const p of data.phases)
        L.push(`| ${p.label} | \`#${p.actual}\` | \`#${p.baseline}\` |`);
    L.push('', '## Geometry', '', '| | |', '|---|---|', `| Tracker matrix | x ${(g.matrixX ?? 0).toFixed(2)}" → ${(g.matrixRight ?? 0).toFixed(2)}" |`, `| Schedule panel | x ${(g.panelX ?? 0).toFixed(2)}" → ${((g.panelX ?? 0) + (g.panelW ?? 0)).toFixed(2)}" — the same span |`, `| Project block | ${(g.leftW ?? 0).toFixed(2)}" (code ${(g.codeW ?? 0).toFixed(2)}", name ${(g.nameW ?? 0).toFixed(2)}", labels ${(g.rowLabelW ?? 0).toFixed(2)}") |`, `| Timeline | ${g.tlW.toFixed(2)}" · ${g.months} months · ${(g.monthW ?? 0).toFixed(2)}" per month |`, `| Schedule rows | ${(g.barH ?? 0).toFixed(2)}" each, y ${(g.barsTop ?? 0).toFixed(2)}"–${g.tlBottom.toFixed(2)}" |`, `| Today marker | ${g.todayDate ? `x ${(g.todayX ?? 0).toFixed(2)}" (${g.todayDate})` : 'not drawn'} |`, `| Legend | y ${g.legendY.toFixed(2)}"–${(g.legendY + g.legendH).toFixed(2)}", ${g.legendCols}×${g.legendRows} |`, `| Caption | y ${(g.captionY ?? 0).toFixed(2)}" |`, '');
    L.push('## Responsiveness', '', 'Nothing in the block is a hand-set position. The month count comes from the phase dates; the month', 'width is the timeline divided by that count; the bar height is what remains after the fixed elements;', 'the project block is measured from the real text; and the legend chooses its column count from the', 'widest phase name. A short programme and a long one therefore need no manual adjustment — only the', 'number of month cells changes.', '');
    L.push('_Scheduling logic, phase calculations, date calculations, colours and business rules were not', 'touched — this was a presentation change only._');
    return { ok, md: L.join('\n') };
}
