"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.keyFocusFilter = void 0;
exports.readRedFlag = readRedFlag;
exports.readTracker = readTracker;
exports.readTrackerStatuses = readTrackerStatuses;
exports.readKeyFocus = readKeyFocus;
exports.readApprovals = readApprovals;
const exceljs_1 = __importDefault(require("exceljs"));
const match_1 = require("./match");
const tracker_template_1 = require("./tracker_template");
/** The Key Focus Area project filter's own record of the last read: what it took, what it refused and
 *  why. Surfaced in key_focus_layout_validation.md so a missing row can be explained, not guessed at. */
exports.keyFocusFilter = null;
const norm = (v) => String(v ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
// ExcelJS fills EVERY cell of a merged range with the same value (unlike openpyxl, which only
// fills the top-left). The TE dashboard merges its macro-group headers ("Land Acquisition" etc.)
// across their sub-columns, so we must collapse runs of identical adjacent values back to the
// first column of each run to recover the real group spans.
function mergedGroups(row) {
    const out = [];
    let prev = '';
    row.eachCell({ includeEmpty: false }, (c, col) => {
        const t = cellTextRaw(c).trim();
        if (t && t !== prev)
            out.push([t, col]);
        prev = t;
    });
    return out;
}
// forward-declared alias so mergedGroups can use cellText before its const definition below
const cellTextRaw = (c) => {
    const v = c.value;
    if (v == null)
        return '';
    if (typeof v === 'object') {
        if (v.text)
            return String(v.text);
        if (v.richText)
            return v.richText.map((r) => r.text).join('');
        if (v.result != null)
            return String(v.result);
        if (v instanceof Date)
            return v.toISOString().slice(0, 10);
    }
    return String(v);
};
const cellText = (c) => {
    const v = c.value;
    if (v == null)
        return '';
    if (typeof v === 'object') {
        if (v.text)
            return String(v.text);
        if (v.richText)
            return v.richText.map((r) => r.text).join('');
        if (v.result != null)
            return String(v.result);
        if (v instanceof Date)
            return v.toISOString().slice(0, 10);
    }
    return String(v);
};
const cellRuns = (c) => {
    const v = c.value;
    const plain = cellText(c);
    const rich = v && typeof v === 'object' && Array.isArray(v.richText) ? v.richText : null;
    const runs = rich
        ? rich.map((r) => ({ text: String(r.text ?? ''), bold: !!(r.font && r.font.bold) }))
        : [{ text: plain, bold: false }];
    // trim the outer edges only
    if (runs.length) {
        runs[0] = { ...runs[0], text: runs[0].text.replace(/^\s+/, '') };
        const last = runs.length - 1;
        runs[last] = { ...runs[last], text: runs[last].text.replace(/\s+$/, '') };
    }
    return runs.filter((r) => r.text !== '');
};
const cellDate = (c) => {
    const v = c.value;
    if (v instanceof Date)
        return v.toISOString().slice(0, 10);
    const s = cellText(c);
    const m = s.match(/\d{4}-\d{2}-\d{2}/);
    return m ? m[0] : null;
};
function fillKey(c) {
    const f = c.fill;
    if (!f || f.type !== 'pattern' || f.pattern !== 'solid' || !f.fgColor)
        return null;
    const fg = f.fgColor;
    if (fg.argb)
        return 'argb:' + fg.argb;
    if (fg.theme != null)
        return 'theme:' + fg.theme + ':' + (fg.tint ?? 0);
    return null;
}
// Best-scoring project row in a dashboard sheet (col 2 = code, col 3 = name).
// An exact code hit always wins; otherwise the strongest name match (see match.ts).
function bestRow(ws, project, code) {
    const pc = norm(code);
    let best = null, bestScore = 0;
    ws.eachRow((row) => {
        const name = cellText(row.getCell(3)).trim();
        const bv = norm(cellText(row.getCell(2)));
        if (!name)
            return;
        if (pc && bv && pc === bv) {
            best = row;
            bestScore = 999;
            return;
        }
        const sc = (0, match_1.matchScore)(project, name);
        if (sc > bestScore) {
            best = row;
            bestScore = sc;
        }
    });
    return bestScore > 0 ? best : null;
}
async function load(buf) {
    const wb = new exceljs_1.default.Workbook();
    await wb.xlsx.load(buf);
    return wb;
}
async function readRedFlag(buf, project, code) {
    const wb = await load(buf);
    const ws = wb.getWorksheet('Red Flag Summary');
    if (!ws)
        return null;
    const GROUP_ROW = 4;
    const groups = mergedGroups(ws.getRow(GROUP_ROW));
    // sub-header row = row with the most text cells among 5..11
    let hdr = 8, best = -1;
    for (let r = 5; r <= 11; r++) {
        let n = 0;
        ws.getRow(r).eachCell({ includeEmpty: false }, () => n++);
        if (n > best) {
            best = n;
            hdr = r;
        }
    }
    const subs = [];
    ws.getRow(hdr).eachCell({ includeEmpty: false }, (c, col) => {
        const name = cellText(c).trim();
        if (col >= 4 && name && !['remarks', 'owner'].includes(norm(name)))
            subs.push({ col, name });
    });
    const last = subs[subs.length - 1]?.col ?? 0;
    // Legend: label → status, plus the swatch's own colour, both taken from the sheet. The slide then
    // renders the dashboard's colours rather than any hardcoded palette, so a future change to the
    // workbook flows through without touching the code.
    const legend = {};
    const legendRows = [];
    const palette = {};
    ws.eachRow((row) => {
        const lab = norm(cellText(row.getCell(3)));
        const map = { good: 'good', atrisk: 'amber', critical: 'red', notapplicable: 'na', na: 'na' };
        const rag = map[lab];
        if (!rag)
            return;
        for (let col = 2; col <= 6; col++) {
            const k = fillKey(row.getCell(col));
            if (!k)
                continue;
            legend[k] = rag;
            const hex = (0, tracker_template_1.fillOf)(row.getCell(col));
            if (hex)
                palette[rag] = hex;
            break;
        }
        legendRows.push([cellText(row.getCell(3)).trim(), rag]);
    });
    // project row — pick the BEST-scoring name match (code exact wins outright)
    const target = bestRow(ws, project, code);
    if (!target)
        return null;
    const cells = subs.map(({ col }) => {
        const c = target.getCell(col);
        if (cellText(c).trim().toUpperCase() === 'NA')
            return 'na';
        const k = fillKey(c);
        return (k && legend[k]) || '';
    });
    const macro = [];
    groups.forEach(([g, gc], i) => {
        const gend = groups[i + 1]?.[1] ?? last + 1;
        const span = subs.filter((s) => s.col >= gc && s.col < gend).length;
        if (span)
            macro.push([g, span]);
    });
    const t = target;
    return {
        project: cellText(t.getCell(3)), code: cellText(t.getCell(2)),
        macro, subs: subs.map((s) => s.name), cells,
        palette, legend: legendRows,
        sheet: 'Red Flag Summary', groupRow: GROUP_ROW, headerRow: hdr,
        rowNumber: t.number, matchedName: cellText(t.getCell(3)).trim(),
    };
}
async function readTracker(buf, project, code) {
    const wb = await load(buf);
    const ws = wb.getWorksheet('Tracker');
    if (!ws)
        return null;
    const groups = mergedGroups(ws.getRow(4));
    let hdr = 8, best = -1;
    for (let r = 5; r <= 11; r++) {
        let n = 0;
        ws.getRow(r).eachCell({ includeEmpty: false }, () => n++);
        if (n > best) {
            best = n;
            hdr = r;
        }
    }
    const subs = [];
    ws.getRow(hdr).eachCell({ includeEmpty: false }, (c, col) => { const name = cellText(c).trim(); if (col >= 4 && name && !['remarks', 'owner'].includes(norm(name)))
        subs.push({ col, name }); });
    const last = subs[subs.length - 1]?.col ?? 0;
    const target = bestRow(ws, project, code);
    if (!target)
        return null;
    const cls = (s) => {
        s = s.trim().toUpperCase();
        if (s === 'NA')
            return 'na';
        if (s === '1' || s === '✔' || s === '✓')
            return 'good';
        if (s === '✺' || s === '*')
            return 'amber';
        if (s === '✘' || s === '✗' || s === 'X')
            return 'red';
        return '';
    };
    const cells = subs.map(({ col }) => cls(cellText(target.getCell(col))));
    const macro = [];
    groups.forEach(([g, gc], i) => { const gend = groups[i + 1]?.[1] ?? last + 1; const span = subs.filter((s) => s.col >= gc && s.col < gend).length; if (span)
        macro.push([g, span]); });
    const t = target;
    return { project: cellText(t.getCell(3)), code: cellText(t.getCell(2)), macro, subs: subs.map((s) => s.name), cells };
}
/**
 * The project's OWN status row on the dashboard's Tracker sheet, keyed by SHEET COLUMN NUMBER.
 *
 * Why this exists: the Tracker slide used to take its layout (groups, columns, legend, colours) from
 * this sheet but then RECOMPUTED every indicator from the ASTA schedule by keyword-matching activity
 * names. So the slide could disagree with the very sheet it was drawn from — e.g. Blue in Green's
 * dashboard row says CLU = complete and Land Acquisition = not started, while the keyword pass
 * reported "NA" for both because no schedule activity matched those words. The dashboard row is the
 * agreed status of record, so it is read here and used directly.
 *
 * Keyed by column NUMBER, not by column name, because the sheet legitimately repeats a heading
 * ("Sales Lounge & Experience Centre" appears under both Design and Site Establishment), so a
 * name-keyed map would collide. TrackerTemplate.columns carries the same sheet column numbers, which
 * makes the join exact.
 *
 * Returns null when there is no Tracker sheet or no row for this project — the caller then keeps the
 * schedule-derived statuses rather than showing another project's row.
 */
async function readTrackerStatuses(buf, project, code) {
    const wb = await load(buf);
    const ws = wb.getWorksheet('Tracker');
    if (!ws)
        return null;
    // Same header-row detection as readTracker: the row between 5 and 11 carrying the most cells.
    let hdr = 8, best = -1;
    for (let r = 5; r <= 11; r++) {
        let n = 0;
        ws.getRow(r).eachCell({ includeEmpty: false }, () => n++);
        if (n > best) {
            best = n;
            hdr = r;
        }
    }
    const target = bestRow(ws, project, code);
    if (!target)
        return null;
    const row = target;
    // The sheet's own vocabulary. "1" is what the dashboard uses for complete; the glyphs are the ones
    // its legend defines. Anything else (including an empty cell) is "no opinion" and is left out of the
    // map, so the caller falls back to the schedule for that column only.
    const cls = (s) => {
        s = s.trim().toUpperCase();
        if (s === 'NA')
            return 'na';
        if (s === '1' || s === '✔' || s === '✓')
            return 'good';
        if (s === '✺' || s === '*')
            return 'amber';
        if (s === '✘' || s === '✗' || s === 'X')
            return 'red';
        return '';
    };
    const byCol = new Map();
    ws.getRow(hdr).eachCell({ includeEmpty: false }, (c, col) => {
        const name = cellText(c).trim();
        if (col < 4 || !name || ['remarks', 'owner'].includes(norm(name)))
            return;
        const rag = cls(cellText(row.getCell(col)));
        if (rag)
            byCol.set(col, rag);
    });
    return { project: cellText(row.getCell(3)), code: cellText(row.getCell(2)), byCol };
}
async function readKeyFocus(buf, project, code) {
    const wb = await load(buf);
    const ws = wb.worksheets.find((w) => norm(w.name).startsWith('keyfocusarea') && !norm(w.name).includes('old'));
    if (!ws)
        return null;
    let hdr = -1;
    for (let r = 1; r <= 12; r++) {
        const vals = [];
        ws.getRow(r).eachCell({ includeEmpty: true }, (c) => vals.push(norm(cellText(c))));
        if (vals.includes('responsible') && (vals.includes('keyfocusarea') || vals.includes('delaydays'))) {
            hdr = r;
            break;
        }
    }
    if (hdr < 0)
        return null;
    const col = {};
    ws.getRow(hdr).eachCell({ includeEmpty: true }, (c, i) => { col[norm(cellText(c))] = i; });
    const cCode = col['projectcode'], cProj = col['projects'], cResp = col['responsible'], cAcc = col['accountable'], cFocus = col['keyfocusarea'], cDelay = col['delaydays'];
    const cMajor = cFocus ? cFocus - 1 : undefined;
    let group = null;
    const out = [];
    // Kept so the decision is auditable: which rows were taken, which were refused and why.
    const accepted = [], rejected = [];
    exports.keyFocusFilter = { project, code: code || '', accepted, rejected };
    for (let r = hdr + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const vals = [];
        row.eachCell({ includeEmpty: true }, (c) => vals.push(cellText(c).trim()));
        const filled = vals.filter(Boolean);
        // A section header ("Critical" / "At Risk") is a row whose non-empty cells all carry the SAME
        // label — true both when it is a single cell and when it is a merged range that ExcelJS has
        // duplicated across columns. (The old `filled.length <= 2` test failed on the merged headers.)
        const distinct = new Set(filled.map((v) => v.toLowerCase().trim()));
        if (filled.length && distinct.size === 1) {
            const only = [...distinct][0];
            if (only.includes('critical')) {
                group = 'Critical';
                continue;
            } // incl. "super critical"
            if (only.includes('at risk')) {
                group = 'At Risk';
                continue;
            }
            continue; // other single-label banners (e.g. "New Projects") — skip, don't emit rows
        }
        const codeV = cCode ? cellText(row.getCell(cCode)).trim() : '';
        const projRaw = cProj ? cellText(row.getCell(cProj)).trim() : '';
        const focus = cFocus ? cellText(row.getCell(cFocus)).trim() : '';
        if (!focus)
            continue;
        // ---- STRICT project match: exact name, or an exactly listed alias, and nothing else ----
        // Key Focus Area is the one place where a false positive is worst: it prints another project's
        // commitments, with named owners, on this project's slide. The dashboard currently holds three
        // projects called "Nandi Hills" ([MH] 0100, [Sult] 0130, [Ach] 0107) and two Violet Hill phases,
        // and the old test admitted a row on a shared-token score — so every Nandi Hills row appeared
        // under every Nandi Hills project. It also used OR, which let a fuzzy name hit override a project
        // code that plainly disagreed.
        const verdict = (0, match_1.strictProjectMatch)(project, code || '', projRaw, codeV);
        if (!verdict.ok) {
            rejected.push(`row ${r}: ${verdict.why}`);
            continue;
        }
        if (!group)
            continue;
        accepted.push(`row ${r}: matched by ${verdict.by}`);
        out.push({
            group,
            responsible: cResp ? cellText(row.getCell(cResp)).trim() : '',
            accountable: cAcc ? cellText(row.getCell(cAcc)).trim() : '',
            major: cMajor ? cellText(row.getCell(cMajor)).trim() : '',
            focus,
            // The same text again, as the runs the cell really holds, so any phrase the author bolded in
            // the dashboard is still bold on the slide. `focus` above is unchanged and remains what the
            // project filter, the row-height measurement and the validations work from.
            focusRuns: cFocus ? cellRuns(row.getCell(cFocus)) : undefined,
            delay: cDelay ? cellText(row.getCell(cDelay)).trim() : '',
        });
    }
    return out.length ? out : null;
}
async function readApprovals(buf, project) {
    const wb = await load(buf);
    // pick the sheet whose name best matches the project (handles "OTR - Ph2" vs "OtR Phase 2",
    // "ShreeGaon" vs "Shreegoan", case/spacing drift, etc.)
    let ws, bestScore = 0;
    for (const w of wb.worksheets) {
        if (/^sheet\d*$/i.test(w.name.trim()))
            continue;
        const sc = (0, match_1.matchScore)(project, w.name);
        if (sc > bestScore) {
            bestScore = sc;
            ws = w;
        }
    }
    if (!ws)
        return null;
    let hdr = -1;
    for (let r = 1; r <= 15; r++) {
        const vals = [];
        ws.getRow(r).eachCell({ includeEmpty: true }, (c) => vals.push(norm(cellText(c))));
        if (vals.includes('slno') && vals.includes('approvals')) {
            hdr = r;
            break;
        }
    }
    if (hdr < 0)
        return null;
    const col = {};
    ws.getRow(hdr).eachCell({ includeEmpty: true }, (c, i) => { col[norm(cellText(c))] = i; });
    const find = (...names) => {
        for (const n of names)
            if (col[n])
                return col[n];
        for (const n of names)
            for (const k of Object.keys(col))
                if (k.startsWith(n) || k.includes(n))
                    return col[k];
        return undefined;
    };
    const cSl = find('slno'), cAp = find('approvals'), cDr = find('drawings'), cApp = find('application'), cFee = find('feepayment'), cIss = find('issued'), cBase = find('baselinefinish'), cProj = find('actualprojectedfinish', 'actualprojected');
    const flag = (row, c) => {
        if (!c)
            return false;
        const v = row.getCell(c).value;
        const n = parseFloat(String(v));
        return !isNaN(n) ? n === 1 : ['y', '✔', '✓'].includes(String(v).trim().toLowerCase());
    };
    const rows = [];
    for (let r = hdr + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const ap = cAp ? cellText(row.getCell(cAp)).trim() : '';
        if (!ap || ['name', 'approvals'].includes(ap.toLowerCase()))
            continue;
        const sl = cSl ? parseInt(cellText(row.getCell(cSl))) : rows.length + 1;
        rows.push({
            sl: isNaN(sl) ? rows.length + 1 : sl, approval: ap,
            drawings: flag(row, cDr), application: flag(row, cApp), fee: flag(row, cFee), issued: flag(row, cIss),
            base: cBase ? cellDate(row.getCell(cBase)) : null, proj: cProj ? cellDate(row.getCell(cProj)) : null,
        });
    }
    return rows.length ? rows : null;
}
