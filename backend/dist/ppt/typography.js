"use strict";
// typography.ts — the single source of truth for fonts and readable sizes.
//
// HR-approved font roles (exact PowerPoint family names, read from the supplied font package):
//   • Page titles   → Impact              (system font on Windows)
//   • General text  → Helvetica Neue      (HelveticaNeue.ttf / -Bold / -Medium / -Light)
//   • Numbers       → Meta-Normal         (METAN.ttf)  ← dates, %, IDs, delay days, Sl.No, durations
//
// If a face is missing we surface an explicit warning (see checkFonts) rather than letting
// PowerPoint silently substitute something else.
Object.defineProperty(exports, "__esModule", { value: true });
exports.fontFor = exports.lineHeight = exports.HEADER_FS_OVERVIEW = exports.HEADER_FS_LOOKAHEAD = exports.HEADER_FS_APPROVAL = exports.KEYFOCUS_HEAD_FS = exports.HEADER_FS_WIDE = exports.HEADER_FS = exports.LOOKAHEAD_FS = exports.OVERVIEW_FS = exports.KEYFOCUS_FS = exports.BODY_FS = exports.TITLE_FS = exports.MIN_FS_HEADER = exports.MIN_FS = exports.REQUIRED_FONTS = exports.FONT_NUM = exports.FONT_TEXT = exports.FONT_TITLE = void 0;
exports.textWidth = textWidth;
exports.linesNeeded = linesNeeded;
exports.fitCell = fitCell;
exports.isNumeric = isNumeric;
exports.checkFonts = checkFonts;
exports.headingLayout = headingLayout;
exports.scopeToPhase = scopeToPhase;
exports.buildTrackerData = buildTrackerData;
exports.FONT_TITLE = 'Impact';
exports.FONT_TEXT = 'Helvetica Neue';
exports.FONT_NUM = 'Meta-Normal';
/** Every font the generated deck references, with the file that provides it. */
exports.REQUIRED_FONTS = [
    { family: exports.FONT_TITLE, file: '(Windows system font)', role: 'page titles', system: true },
    { family: exports.FONT_TEXT, file: 'HelveticaNeue.ttf', role: 'general text' },
    { family: exports.FONT_NUM, file: 'METAN.ttf', role: 'numbers (dates, %, durations, Sl.No, delay days)' },
];
// ---- Readable size floors (26.67" canvas = 2× a normal slide, so these are ~½ in print) ----
exports.MIN_FS = 18; // absolute floor — never render text smaller than this
exports.MIN_FS_HEADER = 16; // column headers may go slightly tighter than body text
exports.TITLE_FS = 40; // page headings (Impact)
exports.BODY_FS = 28; // Approvals body
// Key Focus content is its own size (was sharing the 28pt Approvals body). Set to 24 on request; the
// header row stays at HEADER_FS_WIDE, so the header still reads as dominant.
exports.KEYFOCUS_FS = 24; // Key Focus Area task/content text
exports.OVERVIEW_FS = 24; // Project Overview body
// 2-Week Look Ahead task text. This is a FLOOR as well as a preferred size: a long task name wraps and
// its row grows, but the type never gets smaller — see slides.lookAheadSlides.
// Body text of the 2 Week Look Ahead. This is a FLOOR as well as a preferred size: a long cell wraps and
// its row grows, but the type never gets smaller — see slides.lookAheadSlides. 24 → 22 → 19: the smaller
// body is what frees the width the Notes column needed.
exports.LOOKAHEAD_FS = 19;
exports.HEADER_FS = 18; // table column headers (dense tables: 8-9 columns)
// Wide tables (few, roomy columns) carry a larger heading so the header row is clearly
// distinguishable from the body text without changing the body size.
exports.HEADER_FS_WIDE = 28; // wide-column headers (matches the body size, bold)
/** Key Focus Area column headers. ONE size for all six columns — Projects, Responsible, Accountable,
 *  Team, Key Focus Area and Delay [Days] — so the header row is uniform. A heading that needs more
 *  room gets a WIDER COLUMN, never a smaller size (see KF_COL_W in slides.ts). */
exports.KEYFOCUS_HEAD_FS = 22;
exports.HEADER_FS_APPROVAL = 24; // Approvals stage headings — larger than the 28pt data? no:
// the Approvals body is 28pt but its stage headings are short two-line labels in narrow columns, so 24
// bold reads as clearly dominant without forcing the header row taller.
exports.HEADER_FS_LOOKAHEAD = 21; // 2 Week Look Ahead column headings, regular weight (body is 19pt)
// Project Overview's column headings. Its own constant rather than a change to HEADER_FS, whose 18pt is
// still what the indicator legends use. 24 → 25 → back to **24** on request. The column widths and the
// header row height are deliberately NOT recomputed from this: they stay on the 25pt budget (see
// HDR_METRIC_FS in overview.ts) so dropping a point moves no column and changes no dimension.
exports.HEADER_FS_OVERVIEW = 24;
/**
 * Width of a string in inches at a given point size, using per-face average advance ratios
 * measured from the supplied fonts. Deliberately slightly pessimistic so we never under-estimate.
 */
function textWidth(s, fs, font = exports.FONT_TEXT) {
    const ratio = font === exports.FONT_NUM ? 0.55 : font === exports.FONT_TITLE ? 0.46 : 0.53;
    // caps and digits are wider than lowercase; weight the estimate by composition
    const caps = (s.match(/[A-Z0-9]/g) || []).length / Math.max(s.length, 1);
    return s.length * fs * (ratio + caps * 0.07) / 72;
}
/**
 * Lines needed to wrap `s` into `w` inches at `fs`.
 *
 * The per-line capacity uses the string's OWN average character width (textWidth already weights caps
 * and digits and is deliberately a little pessimistic). Measuring against "M" — the widest glyph —
 * over-estimated prose by roughly 2×, which made tables paginate long before the page was full.
 */
function linesNeeded(s, w, fs, font = exports.FONT_TEXT) {
    if (!s)
        return 1;
    const avg = textWidth(s, fs, font) / Math.max(String(s).length, 1);
    const per = Math.max(1, Math.floor(w / (avg || 0.1)));
    // wrap on words, not characters
    let lines = 1, len = 0;
    for (const word of String(s).split(/\s+/)) {
        const wl = word.length + (len ? 1 : 0);
        if (len + wl > per) {
            lines++;
            len = word.length;
        }
        else {
            len += wl;
        }
    }
    return lines;
}
/** Height (inches) one line of text occupies at `fs`, including leading. */
const lineHeight = (fs) => (fs * 1.22) / 72;
exports.lineHeight = lineHeight;
/**
 * Fit text to a cell, honouring the rule: **grow the row height first, only shrink the font as a
 * last resort, and never below the readable minimum.**
 */
function fitCell(text, colW, opts, font = exports.FONT_TEXT) {
    const min = opts.min ?? exports.MIN_FS;
    const pad = 0.12; // cell padding allowance
    const w = Math.max(colW - pad, 0.2);
    // 1. try the preferred size, growing row height to fit however many lines it needs
    for (let fs = opts.preferred; fs >= min; fs -= 2) {
        const lines = linesNeeded(text, w, fs, font);
        const need = lines * (0, exports.lineHeight)(fs) + 0.08;
        if (need <= opts.maxRowH) {
            return { fontSize: fs, rowH: Math.max(opts.minRowH, need), lines };
        }
    }
    // 2. floor: keep the minimum readable size and let the row be as tall as it must be
    const lines = linesNeeded(text, w, min, font);
    return { fontSize: min, rowH: Math.max(opts.minRowH, lines * (0, exports.lineHeight)(min) + 0.08), lines };
}
/** True when a cell's value is numeric-ish and must therefore use Meta-Normal. */
function isNumeric(v) {
    const s = String(v ?? '').trim();
    if (!s)
        return false;
    return /^[0-9]+(\.[0-9]+)?%?$/.test(s) // 42, 42%, 3.5
        || /^[0-9]{1,3}d$/.test(s) // 749d
        || /^[0-9]{1,2}\s+[A-Za-z]{3}\s+[0-9]{2,4}$/.test(s) // 20 May 25
        || /^[0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{2,4}$/.test(s) // 20/05/25
        || /^[-–]$/.test(s); // an em-dash placeholder
}
/** Font for a cell value: numbers → Meta-Normal, everything else → Helvetica Neue. */
const fontFor = (v) => (isNumeric(v) ? exports.FONT_NUM : exports.FONT_TEXT);
exports.fontFor = fontFor;
/**
 * Verify the bundled font files are present next to the app. We cannot query the end-user's
 * installed fonts from Node, so we check the shipped package and warn loudly if it is incomplete —
 * the deck would otherwise render in a silently substituted face.
 */
function checkFonts(fontsDir, fs) {
    const missing = [], warnings = [];
    if (!fontsDir || !fs.existsSync(fontsDir)) {
        warnings.push(`FONT WARNING: the "fonts" folder was not found. The deck references ${exports.FONT_TEXT} and ` +
            `${exports.FONT_NUM}; if these are not installed on the machine that opens the file, PowerPoint will ` +
            `substitute a different face and the layout will not match the template. ` +
            `Install the HR-approved fonts from the project's fonts/ folder.`);
        return { ok: false, missing: [exports.FONT_TEXT, exports.FONT_NUM], warnings };
    }
    const files = fs.readdirSync(fontsDir).map((f) => f.toLowerCase());
    for (const f of exports.REQUIRED_FONTS) {
        if (f.system)
            continue;
        const stem = f.file.toLowerCase().replace(/\.(ttf|otf)$/, '');
        if (!files.some((x) => x.includes(stem) || x.includes(stem.replace(/-/g, '')))) {
            missing.push(f.family);
            warnings.push(`FONT WARNING: ${f.family} (${f.file}) is missing — used for ${f.role}. ` +
                `PowerPoint will substitute another face; install it before distributing the deck.`);
        }
    }
    return { ok: missing.length === 0, missing, warnings };
}
function headingLayout(labels, cellXs, cellY, cellW, cellH, fs, font) {
    // one cross-axis height for ALL columns → identical positioning everywhere
    const maxLines = Math.max(1, ...labels.map((l) => l.split('\n').length));
    const boxH = Math.max(cellW, maxLines * (0, exports.lineHeight)(fs) + 0.06);
    return labels.map((label, i) => {
        const lines = label.split('\n');
        const textLen = Math.max(...lines.map((l) => textWidth(l, fs, font)));
        const cx = cellXs[i];
        // centre the (pre-rotation) box on the cell centre
        const boxX = cx + (cellW - cellH) / 2;
        const boxY = cellY + (cellH - boxH) / 2;
        const pad = Math.max(0, (cellH - textLen) / 2); // centring ⇒ identical either side
        return {
            label, lines,
            cellX: cx, cellY, cellW, cellH,
            boxX, boxY, boxW: cellH, boxH,
            textLen, padStart: pad, padEnd: pad,
        };
    });
}
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
function buildTrackerData(sch, tpl, cfg, phaseHint = '', schedCfg = { phases: [] }) {
    const scoped = scopeToPhase(sch, phaseHint);
    const acts0 = scoped.activities;
    const kwCfg = (cfg && cfg.column_keywords) || {};
    const bySymbol = new Map();
    for (const l of tpl.legend)
        bySymbol.set(l.label.replace(/\s+/g, '').toLowerCase(), l);
    const styleFor = (st) => {
        const k = st.replace(/\s+/g, '').toLowerCase();
        const l = bySymbol.get(k);
        return l || { symbol: '', color: '000000', font: exports.FONT_TEXT, size: 9 };
    };
    // Indicator symbol/colour overrides from configuration (e.g. Completed → green ✔ instead of "1").
    // Anything not overridden keeps the dashboard sheet's own symbol, face and colour.
    const overrides = (schedCfg && schedCfg.indicatorOverrides) || {};
    const indicators = tpl.columns.map((c) => {
        const toks = keyTokens(c.name, kwCfg[c.name]);
        const matched = acts0.filter((a) => {
            const an = norm(a.name), ad = norm(a.dept || '');
            return toks.some((t) => t.length >= 2 && (tokenHits(an, t) || (ad && tokenHits(ad, t))));
        });
        let status;
        if (!matched.length)
            status = 'Not Applicable';
        else if (matched.every((a) => a.pct >= 100))
            status = 'Completed';
        else if (matched.some((a) => a.pct > 0))
            status = 'In Progress';
        else
            status = 'Not Started';
        const st = styleFor(status);
        const ov = overrides[status] || {};
        return {
            column: c.name, group: c.group, status,
            symbol: ov.symbol ?? st.symbol,
            color: ov.color ?? st.color,
            font: ov.font ?? st.font,
            size: ov.size ?? st.size,
            matched: matched.length,
        };
    });
    // ---- timeline: the major EXECUTION PHASES, resolved from tracker_schedule_config.json ----
    // Activity names, aliases, keywords, finish milestones and the date rules all live in that file —
    // nothing about a project's wording is hard-coded here.
    const resolved = resolvePhases(sch, acts0, schedCfg);
    const colourOf = new Map(tpl.phases.map((p) => [p.label, p]));
    const phases = resolved.map((r) => {
        const c = colourOf.get(r.label);
        return {
            label: r.label,
            actual: c?.actual || '7C7C7C', baseline: c?.baseline || 'DBDBDB',
            aStart: r.aStart, aFinish: r.aFinish, bStart: r.bStart, bFinish: r.bFinish,
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
