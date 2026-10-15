"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.listDashboardProjects = listDashboardProjects;
exports.readDbDashboard = readDbDashboard;
// dashboard.ts — read the DESIGN & BUILD Summary Dashboard workbook.
//
// WHY THIS EXISTS SEPARATELY FROM parsers/excel.ts
// -----------------------------------------------
// Launch & Design reads sheets called "Red Flag Summary", "Tracker" and "Schedule  ". Design & Build
// uses a DIFFERENT workbook with DIFFERENT sheets — "Red Flag Summary DB", "Tracker DB",
// "Key Focus Area", "Schedule - Design", "Schedule - Build" — a different lead-column set
// ("Sl No. / Project Code / Design and Build Projects"), a different band set, and two month-grid
// timeline sheets that Launch & Design has no equivalent of at all.
//
// Rather than bend the Launch & Design parsers with flags, this module owns the Design & Build
// workbook outright. parsers/excel.ts is not touched, so the Launch & Design path cannot change
// behaviour because of anything done here.
//
// WHAT IS PROVEN ABOUT THIS WORKBOOK (measured, not assumed)
// ---------------------------------------------------------
// Both reference decks embed THIS WORKBOOK as an OLE object behind their Red Flag / Key Focus /
// Tracker slides — the embedded copy carries the same eight sheet names. So the dashboard is not a
// nice-to-have input: it IS the source those three slides were pasted from.
//
//   Red Flag Summary DB   bands on row 4, headers row 8, projects from row 10. Cell FILLS carry the
//                         status: 92D050 good, FFC000 at risk, EE0000 critical. Legend at B27:C29.
//   Tracker DB            same geometry; cell TEXT carries the status glyph: ✓ ✺ ✘ and "-"/NA.
//                         Legend at B24:C28.
//   Key Focus Area        headers row 3; severity BANDS ("Critical", "At Risk") are merged fill rows.
//   Schedule - Design     month grid: 12 month columns per year separated by one spacer column.
//                         Two rows per project — "Actual/Projected" then "Baseline" — hand-painted
//                         with phase colours. Phase legend below the data (Acquisition … LOD 350/500).
//   Schedule - Build      identical shape; construction phase legend (Foundations … Project Closure).
//
// NOTHING here hardcodes a colour, a band name, a column name, a month range or a legend. Every one
// of those is read from the workbook, so when the business edits the dashboard the slides follow.
const exceljs_1 = __importDefault(require("exceljs"));
const tracker_template_1 = require("../parsers/tracker_template");
const match_1 = require("../parsers/match");
const norm = (v) => String(v ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
const textOf = (c) => {
    const v = c?.value;
    if (v == null)
        return '';
    if (typeof v === 'object') {
        if (v.richText)
            return v.richText.map((r) => r.text).join('');
        if (v.text != null)
            return String(v.text);
        if (v.result != null)
            return String(v.result);
        if (v instanceof Date)
            return v.toISOString().slice(0, 10);
    }
    return String(v);
};
/**
 * The same cell AS RUNS, keeping each run's bold flag — the Design & Build twin of cellRuns() in
 * parsers/excel.ts, and for the same reason: textOf() above joins rich text into one string, which is
 * correct for codes, names and dates but silently drops the emphasis the author applied inside a Key
 * Focus note. Wording and order are untouched; only the outer whitespace is trimmed, exactly as the
 * plain string is. A cell with no rich text comes back as one run.
 */
const runsOf = (c) => {
    const v = c?.value;
    const rich = v && typeof v === 'object' && Array.isArray(v.richText) ? v.richText : null;
    const runs = rich
        ? rich.map((r) => ({ text: String(r.text ?? ''), bold: !!(r.font && r.font.bold) }))
        : [{ text: textOf(c), bold: false }];
    if (runs.length) {
        runs[0] = { ...runs[0], text: runs[0].text.replace(/^\s+/, '') };
        const last = runs.length - 1;
        runs[last] = { ...runs[last], text: runs[last].text.replace(/\s+$/, '') };
    }
    return runs.filter((r) => r.text !== '');
};
const dateOf = (c) => {
    const v = c?.value;
    if (v instanceof Date)
        return v.toISOString().slice(0, 10);
    if (v && typeof v === 'object' && v.result instanceof Date)
        return v.result.toISOString().slice(0, 10);
    const m = textOf(c).match(/\d{4}-\d{2}-\d{2}/);
    return m ? m[0] : null;
};
// ---------------------------------------------------------------------------------------------
// Sheet lookup
// ---------------------------------------------------------------------------------------------
/**
 * Find a sheet by normalised name. Candidates are tried in order, so the Design & Build sheet
 * ("redflagsummarydb") is preferred over a Launch & Design one that might share the workbook
 * ("redflagsummary"). An exact hit beats a prefix hit; nothing is guessed beyond that.
 */
function sheet(wb, ...candidates) {
    for (const want of candidates) {
        const exact = wb.worksheets.find((w) => norm(w.name) === want);
        if (exact)
            return exact;
    }
    for (const want of candidates) {
        const pre = wb.worksheets.find((w) => norm(w.name).startsWith(want));
        if (pre)
            return pre;
    }
    return undefined;
}
/**
 * Collapse a row of merged band headers into (label, firstColumn) pairs.
 *
 * ExcelJS repeats a merged value across every column it covers (openpyxl does not), so a run of
 * identical adjacent values is one band. This recovers where each band starts.
 */
function bandRuns(row) {
    const out = [];
    let prev = '';
    row.eachCell({ includeEmpty: false }, (c, col) => {
        const t = textOf(c).trim();
        if (t && t !== prev)
            out.push({ label: t, col });
        prev = t;
    });
    return out;
}
/**
 * The last row of the DATA region — the greatest row where both the code column and the name column
 * carry text.
 *
 * This boundary matters. Every one of these sheets puts its legend BELOW the data, and a legend row
 * reads exactly like a data row to a naive scan: "Critical" sits in the name column with a coloured
 * swatch beside it, just as a project name sits in the name column with coloured status cells beside
 * it. Scanning the whole sheet for legend entries therefore picked up the data cell that literally
 * says "NA" (the OTR row's Termsheet) and invented a legend entry from its neighbour's fill.
 *
 * Legend entries are only read strictly BELOW this row; banners and data only AT or ABOVE it.
 */
function lastDataRow(ws, headerRow, cCode, cName) {
    let last = headerRow;
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const codeTxt = textOf(row.getCell(cCode)).trim();
        // A real project code carries an alphanumeric ("0131", "59-05"). The Tracker legend's own rows
        // put a bare glyph in that column — "✓", "✺", "-" — beside a label in the name column, which
        // otherwise reads as the last data row and hides the whole legend below the scan boundary.
        if (!/[a-z0-9]/i.test(codeTxt))
            continue;
        if (textOf(row.getCell(cName)).trim())
            last = r;
    }
    return last;
}
/**
 * Candidate project rows on the dashboard's project-per-row sheet: [code, name].
 *
 * Read BEFORE the project context exists, because the context may need this list to work the code out
 * when the operator did not type one. Nothing is filtered here — this is the portfolio, deliberately.
 */
async function listDashboardProjects(buf) {
    const wb = new exceljs_1.default.Workbook();
    try {
        await wb.xlsx.load(buf);
    }
    catch {
        return [];
    }
    const ws = sheet(wb, 'redflagsummarydb', 'redflagsummary') || sheet(wb, 'trackerdb', 'tracker');
    if (!ws)
        return [];
    const headerRow = findHeaderRow(ws);
    if (headerRow < 0)
        return [];
    const hdr = ws.getRow(headerRow);
    let cCode = 0, cName = 0;
    hdr.eachCell({ includeEmpty: false }, (c, col) => {
        const n = norm(textOf(c));
        if (n === 'projectcode') {
            cCode = col;
            return;
        }
        if (!cName && n.includes('project') && n !== 'projectcode')
            cName = col;
    });
    if (!cCode || !cName)
        return [];
    const end = lastDataRow(ws, headerRow, cCode, cName);
    const out = [];
    for (let r = headerRow + 1; r <= end; r++) {
        const row = ws.getRow(r);
        const code = textOf(row.getCell(cCode)).trim(), name = textOf(row.getCell(cName)).trim();
        if (code && name)
            out.push({ code, name });
    }
    return out;
}
/** The header row is the one naming the project-code column; bands sit on the last text row above it. */
function findHeaderRow(ws, limit = 14) {
    for (let r = 1; r <= limit; r++) {
        let hit = false;
        ws.getRow(r).eachCell({ includeEmpty: false }, (c) => {
            if (norm(textOf(c)) === 'projectcode')
                hit = true;
        });
        if (hit)
            return r;
    }
    return -1;
}
function findBandRow(ws, headerRow) {
    for (let r = headerRow - 1; r >= 1; r--) {
        let n = 0;
        ws.getRow(r).eachCell({ includeEmpty: false }, (c) => { if (textOf(c).trim())
            n++; });
        if (n >= 2)
            return r;
    }
    return -1;
}
/**
 * The project's own row, decided by the project context and nothing else.
 *
 * A project-per-row sheet must yield AT MOST ONE row. Two hits mean the identity test is too loose
 * for this workbook, and the correct response is to take none — omitting a section is recoverable,
 * putting another project's status on this project's slide is not.
 */
function projectRow(ws, headerRow, cCode, cName, ctx) {
    const end = lastDataRow(ws, headerRow, cCode, cName);
    const hits = [];
    for (let r = headerRow + 1; r <= end; r++) {
        const row = ws.getRow(r);
        const name = textOf(row.getCell(cName)).trim();
        if (!name)
            continue;
        if (!ctx.owns(textOf(row.getCell(cCode)).trim(), name))
            continue;
        hits.push(row);
    }
    // One contiguous block is this project (a lead cell merged down two rows reports twice — see
    // readTimeline). Two separate blocks mean the sheet lists the project twice, which is ambiguous.
    const runs = [];
    for (const r of hits) {
        const last = runs[runs.length - 1];
        if (last && r.number === last[last.length - 1].number + 1)
            last.push(r);
        else
            runs.push([r]);
    }
    return { row: runs.length === 1 ? runs[0][0] : null, hits: hits.length };
}
// ---------------------------------------------------------------------------------------------
// Matrix sheets (Red Flag Summary DB / Tracker DB)
// ---------------------------------------------------------------------------------------------
/**
 * Read one Design & Build matrix sheet.
 *
 * `by` selects how a cell's status is decided:
 *   'fill'   — Red Flag Summary DB, where the colour IS the status
 *   'symbol' — Tracker DB, where a glyph (✓ ✺ ✘) IS the status
 *
 * Both share the same geometry, so the band/column/lead discovery is written once.
 */
async function readMatrix(ws, ctx, by, notes) {
    const headerRow = findHeaderRow(ws);
    if (headerRow < 0) {
        notes.push(`${ws.name}: no row names a "Project Code" column — sheet skipped.`);
        return null;
    }
    const bandRow = findBandRow(ws, headerRow);
    // Lead columns: Sl No. / Project Code / the name column. Everything to the RIGHT of the name
    // column with a non-empty heading is a data column; the sheet's blank spacer columns between
    // bands therefore drop out on their own.
    const hdr = ws.getRow(headerRow);
    let cCode = 0, cName = 0, cSl = 0;
    const headings = [];
    hdr.eachCell({ includeEmpty: false }, (c, col) => {
        const raw = textOf(c).trim();
        if (!raw)
            return;
        const n = norm(raw);
        if (n === 'projectcode') {
            cCode = col;
            return;
        }
        if (n === 'slno' || n === 'sino' || n === 'sno') {
            cSl = col;
            return;
        }
        // The name column is whatever the sheet calls it — "Design and Build Projects" or "Projects".
        if (!cName && (n.includes('project') || n === 'projects')) {
            cName = col;
            return;
        }
        headings.push({ col, name: raw });
    });
    if (!cCode || !cName) {
        notes.push(`${ws.name}: could not locate the code and name columns — sheet skipped.`);
        return null;
    }
    const nameHeading = textOf(hdr.getCell(cName)).trim();
    const cols = headings.filter((h) => h.col > cName);
    if (!cols.length) {
        notes.push(`${ws.name}: no data columns to the right of "${nameHeading}" — sheet skipped.`);
        return null;
    }
    // Bands, with each span counted as the DATA columns it actually covers (not raw column width),
    // so spacer columns never inflate a band.
    const bands = [];
    if (bandRow > 0) {
        const runs = bandRuns(ws.getRow(bandRow));
        runs.forEach((run, i) => {
            const end = runs[i + 1]?.col ?? Number.MAX_SAFE_INTEGER;
            const span = cols.filter((c) => c.col >= run.col && c.col < end).length;
            if (span)
                bands.push([run.label, span]);
        });
    }
    const banded = bands.reduce((n, b) => n + b[1], 0);
    if (banded !== cols.length) {
        notes.push(`${ws.name}: ${cols.length} data column(s) but bands cover ${banded} — the ungrouped ` +
            `columns are rendered without a band heading rather than being dropped.`);
    }
    // Legend. Every row below the data whose label maps to a known status, in the sheet's own order,
    // with the sheet's own swatch colour and (for the Tracker) its own glyph.
    const MAP = {
        good: 'good', atrisk: 'amber', critical: 'red',
        completed: 'good', inprogress: 'amber', notstarted: 'red',
        notapplicable: 'na', na: 'na',
    };
    const legend = [];
    const palette = {};
    const fillLegend = new Map(); // fill hex → status (Red Flag)
    const symLegend = new Map(); // glyph    → status (Tracker)
    const dataEnd = lastDataRow(ws, headerRow, cCode, cName);
    for (let r = dataEnd + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        // Scan the first few columns for "<swatch><label>" or "<glyph><label>".
        for (let col = 1; col <= 8; col++) {
            const rag = MAP[norm(textOf(row.getCell(col)))];
            if (rag === undefined)
                continue;
            const label = textOf(row.getCell(col)).trim();
            let color = '', symbol = '', symbolColor = '';
            for (let k = col - 1; k >= Math.max(1, col - 3); k--) {
                const prev = row.getCell(k);
                const hex = (0, tracker_template_1.fillOf)(prev);
                const glyph = textOf(prev).trim();
                if (!color && hex && hex !== 'FFFFFF')
                    color = hex;
                // A status glyph is a short non-alphanumeric mark (✓ ✺ ✘) or the sheet's "-" for N/A.
                if (!symbol && glyph && glyph.length <= 2 && !/[a-z0-9]/i.test(glyph)) {
                    symbol = glyph;
                    symbolColor = (0, tracker_template_1.resolveColor)(prev.font?.color, '') || '';
                }
            }
            if (!color && !symbol)
                break; // a stray label with no swatch — not a legend row
            if (legend.some((l) => norm(l.label) === norm(label)))
                break;
            legend.push({ label, rag, color, symbol, symbolColor });
            if (color && !palette[rag])
                palette[rag] = color;
            if (color)
                fillLegend.set(color, rag);
            if (symbol)
                symLegend.set(symbol, rag);
            break;
        }
    }
    notes.push(`${ws.name}: data rows ${headerRow + 1}–${dataEnd}; legend read from below row ${dataEnd} ` +
        `(${legend.length} entr${legend.length === 1 ? 'y' : 'ies'}: ${legend.map((l) => l.label).join(', ') || 'none'}).`);
    const found = projectRow(ws, headerRow, cCode, cName, ctx);
    const target = found.row;
    if (!target) {
        notes.push(`${ws.name}: ${found.hits === 0
            ? `no row belongs to ${ctx.name}${ctx.code ? ` (code ${ctx.code})` : ''}`
            : `${found.hits} rows matched ${ctx.name}, which is ambiguous`} — section omitted rather than ` +
            `showing another project's status.`);
        return null;
    }
    const classify = (c) => {
        const txt = textOf(c).trim();
        if (/^na$/i.test(txt))
            return 'na';
        if (txt === '-')
            return ''; // the sheet's explicit "nothing here"
        if (by === 'symbol') {
            if (symLegend.has(txt))
                return symLegend.get(txt);
            // Fall back to the glyphs this workbook uses, so a missing legend row cannot blank the slide.
            if (['✓', '✔', '1'].includes(txt))
                return 'good';
            if (['✺', '*'].includes(txt))
                return 'amber';
            if (['✘', '✗', 'x', 'X'].includes(txt))
                return 'red';
            return '';
        }
        const hex = (0, tracker_template_1.fillOf)(c);
        if (!hex || hex === 'FFFFFF')
            return '';
        if (fillLegend.has(hex))
            return fillLegend.get(hex);
        notes.push(`${ws.name}: cell fill #${hex} in column ${c.col} is not in the sheet's legend — left blank.`);
        return '';
    };
    const leadHeadings = [];
    if (cSl)
        leadHeadings.push(textOf(hdr.getCell(cSl)).trim());
    leadHeadings.push(textOf(hdr.getCell(cCode)).trim(), nameHeading);
    return {
        nameHeading, leadHeadings,
        project: textOf(target.getCell(cName)).trim(),
        code: textOf(target.getCell(cCode)).trim(),
        bands, cols: cols.map((c) => c.name),
        cells: cols.map((c) => classify(target.getCell(c.col))),
        raw: cols.map((c) => textOf(target.getCell(c.col)).trim()),
        legend, palette,
        sheet: ws.name, groupRow: bandRow, headerRow, rowNumber: target.number,
        matchedName: textOf(target.getCell(cName)).trim(),
    };
}
// ---------------------------------------------------------------------------------------------
// Key Focus Area
// ---------------------------------------------------------------------------------------------
/**
 * Read the Key Focus Area rows for one project.
 *
 * Severity is NOT a column — it is a merged banner row ("Critical", "At Risk") that every row below
 * it belongs to until the next banner. The banner's own fill is captured too, so the slide's band
 * colours come from the sheet instead of being hardcoded.
 */
function readKeyFocus(ws, ctx, notes) {
    const empty = { rows: [], bands: [] };
    let headerRow = -1;
    for (let r = 1; r <= Math.min(ws.rowCount, 14); r++) {
        const vals = [];
        ws.getRow(r).eachCell({ includeEmpty: true }, (c) => vals.push(norm(textOf(c))));
        if (vals.includes('responsible') && vals.includes('keyfocusarea')) {
            headerRow = r;
            break;
        }
    }
    if (headerRow < 0) {
        notes.push(`${ws.name}: no header row carries both "Responsible" and "Key Focus Area" — section omitted.`);
        return empty;
    }
    const col = {};
    ws.getRow(headerRow).eachCell({ includeEmpty: true }, (c, i) => {
        const n = norm(textOf(c));
        if (n && !col[n])
            col[n] = i;
    });
    const cCode = col['projectcode'], cProj = col['projects'] || col['project'], cResp = col['responsible'], cAcc = col['accountable'], cFocus = col['keyfocusarea'];
    if (!cFocus) {
        notes.push(`${ws.name}: no "Key Focus Area" column — section omitted.`);
        return empty;
    }
    // Severity banners and data rows live at or above the last row that carries both a code and a
    // project name; the sheet's own legend sits below that and must not be read as a banner.
    const dataEnd = lastDataRow(ws, headerRow, cCode || 1, cProj || 1);
    const rows = [];
    // Rows belonging to OTHER projects, counted so the report can state what was excluded rather than
    // leaving the operator to wonder why the slide is shorter than the sheet.
    let rejected = 0;
    const bands = [];
    let group = '', bandColor = '';
    for (let r = headerRow + 1; r <= dataEnd; r++) {
        const row = ws.getRow(r);
        const vals = [];
        row.eachCell({ includeEmpty: true }, (c) => vals.push(textOf(c).trim()));
        const filled = vals.filter(Boolean);
        const distinct = new Set(filled.map((v) => v.toLowerCase()));
        // A severity banner is a merged row carrying ONE label and nothing else.
        //
        // It is NOT detected by counting text cells. ExcelJS models a merged range as one valued cell
        // plus siblings that are blank-but-filled, so the banner "Critical" — which spans eight columns
        // — presents as a single non-empty string. An earlier version required three or more text cells
        // and so found no banners at all, which silently emptied the whole section. What actually marks
        // a banner is that its cells are PAINTED across the row while only one of them has text.
        if (filled.length === 1 && distinct.size === 1) {
            let painted = 0, fill = '';
            row.eachCell({ includeEmpty: true }, (c) => {
                const hex = (0, tracker_template_1.fillOf)(c);
                if (hex && hex !== 'FFFFFF') {
                    painted++;
                    if (!fill)
                        fill = hex;
                }
            });
            if (painted >= 3) {
                group = filled[0];
                bandColor = fill;
                if (!bands.some((b) => norm(b.label) === norm(group)))
                    bands.push({ label: group, color: bandColor });
            }
            continue;
        }
        const focus = textOf(row.getCell(cFocus)).trim();
        const focusRuns = runsOf(row.getCell(cFocus));
        if (!focus || !group)
            continue;
        const rowCode = cCode ? textOf(row.getCell(cCode)).trim() : '';
        const pn = cProj ? textOf(row.getCell(cProj)).trim() : '';
        // THE PROJECT FILTER, applied before anything else touches the row.
        //
        // Key Focus Area uses the STRICT test, not ctx.owns(). ctx.owns falls back to name matching that
        // accepts a shared word, and the dashboard holds three projects called "Nandi Hills" and two
        // Violet Hill phases — so a shared-word match puts one project's commitments, with named owners,
        // on another project's slide. Here the project CODE decides whenever both sides carry one; failing
        // that, only an exact name or an exactly listed alias is accepted. Every other Design & Build
        // section keeps ctx.owns() unchanged.
        const verdict = (0, match_1.strictProjectMatch)(ctx.name, ctx.code || '', pn, rowCode);
        if (!verdict.ok) {
            rejected++;
            continue;
        }
        rows.push({
            group, bandColor,
            code: cCode ? textOf(row.getCell(cCode)).trim() : '',
            project: pn,
            responsible: cResp ? textOf(row.getCell(cResp)).trim() : '',
            accountable: cAcc ? textOf(row.getCell(cAcc)).trim() : '',
            focus,
            focusRuns,
        });
    }
    // Only bands that actually carry a row for this project are kept — an empty "At Risk" band would
    // be an orphaned heading on the slide.
    const used = bands.filter((b) => rows.some((x) => norm(x.group) === norm(b.label)));
    notes.push(`${ws.name}: ${rows.length} Key Focus row(s) for ${ctx.name} across ${used.length} band(s)` +
        (used.length ? ` (${used.map((b) => b.label).join(', ')})` : '') +
        `; ${rejected} row(s) belonging to other projects were excluded before rendering.`);
    return { rows, bands: used };
}
// ---------------------------------------------------------------------------------------------
// Timeline sheets (Schedule - Design / Schedule - Build)
// ---------------------------------------------------------------------------------------------
/**
 * Read one month-grid timeline sheet for one project.
 *
 * The grid is hand-painted: each project occupies two rows — "Actual/Projected" then "Baseline" —
 * and each painted month cell carries a phase colour. The phase legend sits below the data as
 * [projected swatch][baseline swatch][label].
 *
 * Nothing is inferred from dates here. The bar is exactly the months the business painted, which is
 * the only reading that can reproduce the reference slide.
 */
function readTimeline(ws, ctx, notes) {
    const headerRow = findHeaderRow(ws);
    if (headerRow < 0) {
        notes.push(`${ws.name}: no row names a "Project Code" column — timeline omitted.`);
        return null;
    }
    const hdr = ws.getRow(headerRow);
    let cCode = 0, cName = 0, cBaseStart = 0, cBaseFinish = 0, cStart = 0, cFinish = 0;
    const months = [];
    hdr.eachCell({ includeEmpty: false }, (c, col) => {
        const n = norm(textOf(c));
        if (n === 'projectcode') {
            cCode = col;
            return;
        }
        if (!cName && n === 'projects') {
            cName = col;
            return;
        }
        if (n === 'basestart') {
            cBaseStart = col;
            return;
        }
        if (n === 'basefinish') {
            cBaseFinish = col;
            return;
        }
        if (n === 'start') {
            cStart = col;
            return;
        }
        if (n === 'finish') {
            cFinish = col;
            return;
        }
        // Month columns hold real dates.
        const d = dateOf(c);
        if (d)
            months.push({ col, key: d.slice(0, 7) });
    });
    if (!cCode || !cName || !months.length) {
        notes.push(`${ws.name}: expected a code column, a name column and dated month columns — timeline omitted.`);
        return null;
    }
    // Phase legend: rows below the data of the form [projected fill][baseline fill][label].
    const legend = [];
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        for (let col = 3; col <= 8; col++) {
            const label = textOf(row.getCell(col)).trim();
            if (!label || label.length < 3 || /^\d+$/.test(label))
                continue;
            const projFill = (0, tracker_template_1.fillOf)(row.getCell(col - 2));
            const baseFill = (0, tracker_template_1.fillOf)(row.getCell(col - 1));
            if (!projFill || !baseFill || projFill === 'FFFFFF' || baseFill === 'FFFFFF')
                continue;
            if (legend.some((l) => norm(l.label) === norm(label)))
                continue;
            legend.push({ label, projected: projFill, baseline: baseFill });
            break;
        }
    }
    // The project's pair of rows. The label column ("Actual/Projected" / "Baseline") is whichever
    // column carries those words next to the project's own row.
    const end = lastDataRow(ws, headerRow, cCode, cName);
    const owned = [];
    for (let r = headerRow + 1; r <= end; r++) {
        const row = ws.getRow(r);
        const name = textOf(row.getCell(cName)).trim();
        if (!name)
            continue;
        if (!ctx.owns(textOf(row.getCell(cCode)).trim(), name))
            continue;
        owned.push(r);
    }
    // A project occupies a PAIR of rows here — "Actual/Projected" then "Baseline" — and the code and
    // name cells are merged down the pair. ExcelJS repeats a merged value across every row it covers, so
    // an owned project legitimately reports TWO adjacent hits. Requiring exactly one row rejected both
    // NH Homestay timelines as "ambiguous" and silently dropped its Schedule - Construction slide.
    //
    // So the test is on contiguous RUNS, not on row count: one run is the project's pair, two separate
    // runs would mean the sheet really does list this project twice and is genuinely ambiguous.
    const runs = [];
    for (const r of owned) {
        const last = runs[runs.length - 1];
        if (last && r === last[last.length - 1] + 1)
            last.push(r);
        else
            runs.push([r]);
    }
    if (runs.length !== 1 || runs[0].length > 2) {
        notes.push(`${ws.name}: ${owned.length === 0
            ? `no row belongs to ${ctx.name}${ctx.code ? ` (code ${ctx.code})` : ''}`
            : `${ctx.name} appears in ${runs.length} separate row block(s) (${runs.map((x) => x.join('+')).join(', ')}), ` +
                `which is ambiguous`} — timeline omitted rather than drawing another project's bar.`);
        return null;
    }
    const rowP = runs[0][0], rowB = rowP + 1;
    const matched = ws.getRow(rowP);
    const paint = (r) => months.map(({ col }) => {
        const hex = (0, tracker_template_1.fillOf)(ws.getRow(r).getCell(col));
        return hex && hex !== 'FFFFFF' ? hex : '';
    });
    const projected = paint(rowP), baseline = paint(rowB);
    let from = -1, to = -1;
    months.forEach((_, i) => {
        if (projected[i] || baseline[i]) {
            if (from < 0)
                from = i;
            to = i;
        }
    });
    if (from < 0) {
        notes.push(`${ws.name}: row ${rowP} for ${ctx.name} carries no painted month — this project has ` +
            `no bar on this sheet, so the section is omitted rather than drawn empty.`);
    }
    const pr = ws.getRow(rowP), br = ws.getRow(rowB);
    return {
        sheet: ws.name,
        project: textOf(pr.getCell(cName)).trim(),
        code: textOf(pr.getCell(cCode)).trim(),
        months: months.map((m) => m.key), projected, baseline, legend, from, to,
        // Base dates sit on the Baseline row, live dates on the Actual/Projected row — as the sheet has them.
        baseStart: (cBaseStart ? dateOf(br.getCell(cBaseStart)) : null) || (cBaseStart ? dateOf(pr.getCell(cBaseStart)) : null),
        baseFinish: (cBaseFinish ? dateOf(br.getCell(cBaseFinish)) : null) || (cBaseFinish ? dateOf(pr.getCell(cBaseFinish)) : null),
        start: cStart ? dateOf(pr.getCell(cStart)) : null,
        finish: cFinish ? dateOf(pr.getCell(cFinish)) : null,
        rowProjected: rowP, rowBaseline: rowB,
    };
}
// ---------------------------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------------------------
/**
 * Read everything Design & Build needs from the Summary Dashboard workbook.
 *
 * Every section is independent: a sheet that cannot be read leaves its own field null and records
 * why in `notes`. The caller then omits that section from the deck AND from Contents, rather than
 * rendering a heading over nothing.
 */
async function readDbDashboard(buf, ctx) {
    const notes = [];
    const wb = new exceljs_1.default.Workbook();
    try {
        await wb.xlsx.load(buf);
    }
    catch (e) {
        return { redFlag: null, tracker: null, keyFocus: [], keyFocusBands: [], design: null, build: null,
            notes: [`The Summary Dashboard workbook could not be opened (${e?.message || e}).`] };
    }
    notes.push(`Summary Dashboard sheets: ${wb.worksheets.map((w) => w.name).join(' · ')}`);
    notes.push(`Scoped to ${ctx.name}${ctx.code ? `, code "${ctx.code}"` : ''} (code source: ` +
        `${ctx.codeSource}). Every sheet below is filtered to this project before anything is rendered.`);
    const wsRf = sheet(wb, 'redflagsummarydb', 'redflagsummary');
    const wsTr = sheet(wb, 'trackerdb', 'tracker');
    const wsKf = sheet(wb, 'keyfocusarea');
    const wsSd = sheet(wb, 'scheduledesign');
    const wsSb = sheet(wb, 'schedulebuild', 'scheduleconstruction');
    if (!wsRf)
        notes.push('No "Red Flag Summary DB" sheet — Red Flag Summary omitted.');
    if (!wsTr)
        notes.push('No "Tracker DB" sheet — the Tracker status matrix is omitted.');
    if (!wsKf)
        notes.push('No "Key Focus Area" sheet — Key Focus Area omitted.');
    if (!wsSd)
        notes.push('No "Schedule - Design" sheet — the design timeline is omitted.');
    if (!wsSb)
        notes.push('No "Schedule - Build" sheet — the construction timeline is omitted.');
    const redFlag = wsRf ? await readMatrix(wsRf, ctx, 'fill', notes) : null;
    const tracker = wsTr ? await readMatrix(wsTr, ctx, 'symbol', notes) : null;
    const kf = wsKf ? readKeyFocus(wsKf, ctx, notes) : { rows: [], bands: [] };
    const design = wsSd ? readTimeline(wsSd, ctx, notes) : null;
    const build = wsSb ? readTimeline(wsSb, ctx, notes) : null;
    return { redFlag, tracker, keyFocus: kf.rows, keyFocusBands: kf.bands, design, build, notes };
}
