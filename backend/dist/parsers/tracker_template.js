"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.applyTint = applyTint;
exports.resolveColor = resolveColor;
exports.fillOf = fillOf;
exports.readTrackerTemplate = readTrackerTemplate;
// tracker_template.ts — reads the STATIC Tracker & Schedule template out of Summary Dashboard.xlsx.
//
// The "Tracker" sheet is the master template for the indicator matrix and the "Schedule  " sheet is
// the master template for the timeline below it. Everything structural — category groups, column
// names, column order, legends, colours, fonts — is read from those sheets, never from the previous
// PowerPoint. Only the indicator values, project information and timeline dates vary per project.
const exceljs_1 = __importDefault(require("exceljs"));
// Excel theme colour scheme order used by ExcelJS/OOXML: 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4..9 accent1..6
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
/** Excel's tint algorithm — negative darkens toward black, positive lightens toward white. */
function applyTint(hex, t) {
    const p = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
    const f = (v) => {
        let x = t < 0 ? v * (1 + t) : t > 0 ? v * (1 - t) + 255 * t : v;
        return Math.max(0, Math.min(255, Math.round(x)));
    };
    return p.map((v) => f(v).toString(16).padStart(2, '0')).join('').toUpperCase();
}
/** Resolve any ExcelJS colour object (argb / theme+tint / indexed) to a plain RRGGBB. */
function resolveColor(c, fallback = '') {
    if (!c)
        return fallback;
    if (typeof c.argb === 'string' && c.argb.length >= 6)
        return c.argb.slice(-6).toUpperCase();
    if (typeof c.theme === 'number') {
        const base = THEME[c.theme] || fallback || 'FFFFFF';
        return applyTint(base, c.tint || 0);
    }
    return fallback;
}
function fillOf(cell, fallback = '') {
    const f = cell?.fill;
    if (f && f.type === 'pattern' && f.pattern === 'solid')
        return resolveColor(f.fgColor, fallback);
    return fallback;
}
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
    }
    return String(v);
};
/**
 * Parse the static template. Returns null when the workbook has no Tracker sheet, so the caller can
 * fall back rather than emit a wrong slide.
 */
async function readTrackerTemplate(buf) {
    const wb = new exceljs_1.default.Workbook();
    await wb.xlsx.load(buf);
    const tr = wb.worksheets.find((w) => w.name.trim().toLowerCase() === 'tracker');
    if (!tr)
        return null;
    const sc = wb.worksheets.find((w) => w.name.trim().toLowerCase().startsWith('schedule'));
    // ---- group headers: the merged bands on the group row ----
    // ExcelJS exposes merges as "E4:K4"; collapse to (name, from, to).
    const GROUP_ROW = 4, HEADER_ROW = 8;
    const groups = [];
    const merges = (tr.model?.merges || []);
    const colOf = (ref) => {
        const m = ref.match(/^([A-Z]+)(\d+)$/);
        if (!m)
            return { c: 0, r: 0 };
        let c = 0;
        for (const ch of m[1])
            c = c * 26 + (ch.charCodeAt(0) - 64);
        return { c, r: +m[2] };
    };
    for (const mg of merges) {
        const [a, b] = mg.split(':');
        if (!b)
            continue;
        const A = colOf(a), B = colOf(b);
        if (A.r !== GROUP_ROW || B.r !== GROUP_ROW)
            continue;
        const name = textOf(tr.getRow(GROUP_ROW).getCell(A.c)).trim();
        if (name)
            groups.push({ name, from: A.c, to: B.c, span: B.c - A.c + 1 });
    }
    groups.sort((x, y) => x.from - y.from);
    // Fallback: single non-merged group cells (a differently-built dashboard)
    if (!groups.length) {
        let prev = '';
        tr.getRow(GROUP_ROW).eachCell({ includeEmpty: false }, (c, i) => {
            const t = textOf(c).trim();
            if (t && t !== prev)
                groups.push({ name: t, from: i, to: i, span: 1 });
            prev = t;
        });
    }
    // ---- sub-column headers, in sheet order, restricted to the group bands ----
    const columns = [];
    for (const g of groups) {
        for (let c = g.from; c <= g.to; c++) {
            const name = textOf(tr.getRow(HEADER_ROW).getCell(c)).trim();
            if (name)
                columns.push({ col: c, name, group: g.name });
        }
    }
    // recompute spans from the columns that actually carry a name
    for (const g of groups)
        g.span = columns.filter((c) => c.group === g.name).length;
    const leadHeaders = [];
    for (let c = 1; c < (groups[0]?.from ?? 4); c++) {
        const t = textOf(tr.getRow(HEADER_ROW).getCell(c)).trim();
        if (t)
            leadHeaders.push(t);
    }
    // ---- indicator legend: rows below the data carrying symbol + label ----
    const legend = [];
    const LABELS = /^(completed|in ?progress|not ?started|not ?applicable)$/i;
    for (let r = 1; r <= tr.rowCount; r++) {
        const symCell = tr.getRow(r).getCell(2), labCell = tr.getRow(r).getCell(3);
        const label = textOf(labCell).trim();
        if (!LABELS.test(label))
            continue;
        const symbol = textOf(symCell).trim();
        const f = symCell.font || {};
        legend.push({
            symbol, label,
            color: resolveColor(f.color, '000000'),
            font: f.name || 'Helvetica Neue',
            size: Number(f.size || 9),
        });
    }
    // ---- timeline template from the Schedule sheet ----
    const phases = [];
    const timelineHeaders = [];
    if (sc) {
        for (let r = 1; r <= sc.rowCount; r++) {
            const label = textOf(sc.getRow(r).getCell(3)).trim();
            if (!label)
                continue;
            const a = fillOf(sc.getRow(r).getCell(1)), b = fillOf(sc.getRow(r).getCell(2));
            // a legend row is one where BOTH swatch cells are filled and there is a label
            if (a && b)
                phases.push({ label, actual: a, baseline: b });
        }
        for (let c = 1; c <= 12; c++) {
            const t = textOf(sc.getRow(6).getCell(c)).trim();
            if (t && !timelineHeaders.includes(t))
                timelineHeaders.push(t);
        }
    }
    const gf = tr.getRow(GROUP_ROW).getCell(groups[0]?.from ?? 5).font || {};
    const sf = tr.getRow(HEADER_ROW).getCell(groups[0]?.from ?? 5).font || {};
    return {
        groups, columns, leadHeaders, legend, phases, timelineHeaders,
        headerFill: fillOf(tr.getRow(GROUP_ROW).getCell(groups[0]?.from ?? 5), 'D9D9D9'),
        headerFont: gf.name || 'Helvetica Neue Medium', headerSize: Number(gf.size || 9),
        subFont: sf.name || 'Helvetica Neue Medium', subSize: Number(sf.size || 8),
        source: { trackerSheet: tr.name, scheduleSheet: sc?.name || '(none)', groupRow: GROUP_ROW, headerRow: HEADER_ROW },
    };
}
