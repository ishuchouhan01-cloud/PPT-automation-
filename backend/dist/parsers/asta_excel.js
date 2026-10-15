"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.lastSheetChoice = void 0;
exports.parseScheduleFromExcel = parseScheduleFromExcel;
exports.wbsPath = wbsPath;
// Parse an ASTA Powerproject *spreadsheet* export (.xlsx/.xlsm, single "Sheet1") into the same
// Schedule model the PPT builders consume. WBS hierarchy is reconstructed from Excel OUTLINE
// LEVELS (the export carries no indentation in the Name text). No Python, no Office automation.
const exceljs_1 = __importDefault(require("exceljs"));
const match_1 = require("./match");
const norm = (s) => String(s ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
const cellText = (c) => {
    const v = c?.value;
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
const cellDate = (c) => {
    const v = c?.value;
    if (v instanceof Date)
        return v.toISOString().slice(0, 10);
    const m = cellText(c).match(/\d{4}-\d{2}-\d{2}/);
    return m ? m[0] : null;
};
/** Which worksheet chosen, and why — printed in the generation log. */
exports.lastSheetChoice = '';
/**
 * WHICH WORKSHEET THE SCHEDULE IS ON.
 *
 * It used to be `worksheets[0]`, on the assumption that an export has one sheet. Several do not.
 *
 * NH Homestay's export carries TWO: a HIDDEN "Sheet1" left over from an earlier run, and the visible
 * "Sheet2" holding the current schedule. `worksheets[0]` is the hidden one, so every date on the deck
 * came from a stale sheet nobody could see — the project finished 05-Nov-27 on the slide and
 * 02-Nov-27 in the file, Design 17-Nov-26 against 07-Oct-27. Both sheets are real exports with the
 * same headers and nearly the same row count, so nothing downstream could have noticed.
 *
 * The rule, in order:
 *   1. VISIBLE sheets first. A hidden sheet is one the author put away; it cannot be the sheet
 *      somebody is reading a date off, which is the whole point of the comparison.
 *   2. Among those, the first that LOOKS LIKE A SCHEDULE — a "Name" column plus at least one of the
 *      date columns — so a notes or lookup tab sitting in front of the data is skipped.
 *   3. If no visible sheet qualifies, the first qualifying hidden one, so a workbook whose only data
 *      sheet happens to be hidden still works instead of failing.
 *   4. Failing everything, `worksheets[0]`, which is the old behaviour and keeps single-sheet exports
 *      working exactly as before.
 *
 * The choice is recorded so the generation log can state which sheet was read.
 */
function pickScheduleSheet(wb) {
    const all = wb.worksheets || [];
    const visible = (w) => (w.state || 'visible') === 'visible';
    const looksLikeSchedule = (w) => {
        const heads = [];
        try {
            w.getRow(1).eachCell({ includeEmpty: false }, (c) => heads.push(norm(cellText(c))));
        }
        catch {
            return false;
        }
        const has = (k) => heads.some((h) => h.includes(k));
        return has('name') && (has('actualprojected') || has('baselinestart') || has('baselinefinish'));
    };
    const describe = (w, why) => {
        const others = all.filter((x) => x !== w)
            .map((x) => `"${x.name}" (${x.state || 'visible'}, ${x.rowCount} rows)`);
        exports.lastSheetChoice = `read "${w.name}" (${w.state || 'visible'}, ${w.rowCount} rows) — ${why}` +
            (others.length ? `; ignored ${others.join(', ')}` : '');
        return w;
    };
    const visibleSchedule = all.find((w) => visible(w) && looksLikeSchedule(w));
    if (visibleSchedule) {
        return describe(visibleSchedule, all.length > 1
            ? 'the first VISIBLE sheet carrying schedule columns' : 'the only sheet');
    }
    const anySchedule = all.find(looksLikeSchedule);
    if (anySchedule)
        return describe(anySchedule, 'no visible sheet carried schedule columns, so the first one that does was used');
    return describe(all[0], 'no sheet carried recognisable schedule columns; the first sheet was used');
}
async function parseScheduleFromExcel(buf) {
    const wb = new exceljs_1.default.Workbook();
    await wb.xlsx.load(buf);
    const ws = pickScheduleSheet(wb);
    // map headers
    const col = {};
    ws.getRow(1).eachCell({ includeEmpty: false }, (c, i) => { col[norm(cellText(c))] = i; });
    /**
     * Locate a column by name: an exact header first, then a header that merely CONTAINS the name.
     *
     * `not` guards the loose second pass. "Duration" and "Duration remaining" are different columns,
     * and an export that carries only the second one would otherwise have it accepted as the first —
     * silently, because the loose pass matches any header containing "duration". That is not a
     * hypothetical: the Straight from the Heart exports carry "Duration remaining" and no "Duration",
     * so every COMPLETED task (nothing left to run, "0d" remaining) was being read as a zero-duration
     * MILESTONE and dropped from the 2 Week Look Ahead. A column is only guessed at when the guess
     * cannot be one of these near-misses.
     */
    const find = (...args) => {
        const keys = args.filter((a) => typeof a === 'string');
        const not = args.find((a) => typeof a === 'object')?.not || [];
        for (const k of keys)
            if (col[k])
                return col[k];
        for (const k of keys) {
            for (const h of Object.keys(col)) {
                if (!h.includes(k))
                    continue;
                if (not.some((n) => h.includes(norm(n))))
                    continue;
                return col[h];
            }
        }
        return undefined;
    };
    // NB: header is "Overall % Complte" — the "%" is stripped by norm() and "Complete" is
    // misspelled in the export, so match on the substring "complt" (covers complte/complete).
    const cName = find('name'), cPct = find('overallcomplt', 'complt', 'percentcomplete'), cBS = find('baselinestart'), cBF = find('baselinefinish'), cAS = find('actualprojectedstart'), cAF = find('actualprojectedfinish'), 
    /**
     * THE TWO DURATION COLUMNS.
     *
     * ASTA can export duration twice: "Duration", counted in the task's working time, and
     * "Duration [Elapsed day]", counted in calendar days. An export may carry either, both, or — in
     * the case of "Duration remaining" — a third column that is neither and must never be mistaken
     * for one, which is what the `not` guards are for.
     *
     * WHICH EXPORTS HAVE WHAT, measured across the files on hand:
     *   • most shapes      "Duration" and "Duration remaining", no elapsed column
     *   • NH Homestay      "Duration [Elapsed day]" and "Duration remaining", NO plain "Duration"
     *
     * That second shape is why `cDur` falls back below. Treating "no plain Duration column" as "no
     * duration at all" gave every row a duration of 0, which made every WBS node a milestone, which
     * emptied the Custom branch picker, blanked both duration columns on the Project Overview, and
     * silently switched off the look-ahead's milestone rule. One missing header should not be able to
     * do that: the elapsed column IS this export's duration, and is read as such.
     */
    cDurPlain = find('duration', { not: ['remaining', 'elapsed'] }), cEDur = find('durationelapsedday', 'elapsedduration', 'durationelapsed', 'elapsedday', { not: ['remaining', 'month'] }), 
    /** The duration everything outside the milestone test uses — the plain column when there is one. */
    cDur = cDurPlain ?? cEDur, cTeam = find('teteam', 'team'), cDept = find('tedepartments', 'departments'), cNote = find('delaydocumentation', 'note', 'notes'), 
    // People. The export's shape has changed between versions, so all of them are accepted:
    //   • current  — three columns: Responsible / Accountable / Head
    //   • reduced  — one column: "Responsible" (or "Responsible Person")
    //   • legacy   — one column: "Task.TE Team"
    // Whatever is missing is reconstructed in ppt/roles.ts; nothing is ever left blank on a slide.
    cResp = find('responsible', 'responsibleperson'), cAcc = find('accountable'), cHead = find('head');
    const raws = [];
    for (let r = 2; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const name = cName ? cellText(row.getCell(cName)).trim() : '';
        if (!name)
            continue;
        const durStr = cDur ? cellText(row.getCell(cDur)) : '';
        const dm = durStr.match(/(\d+(?:\.\d+)?)/);
        const edurStr = cEDur ? cellText(row.getCell(cEDur)) : '';
        const edm = edurStr.match(/(\d+(?:\.\d+)?)/);
        let pct = 0;
        if (cPct) {
            const p = parseFloat(cellText(row.getCell(cPct)));
            if (!isNaN(p))
                pct = p <= 1 ? Math.round(p * 100) : Math.round(p);
        }
        raws.push({
            line: r - 1, level: row.outlineLevel || 0, name, pct,
            bs: cBS ? cellDate(row.getCell(cBS)) : null, bf: cBF ? cellDate(row.getCell(cBF)) : null,
            s: cAS ? cellDate(row.getCell(cAS)) : null, f: cAF ? cellDate(row.getCell(cAF)) : null,
            dur: dm ? Math.round(parseFloat(dm[1])) : 0,
            // The same figure UNROUNDED — "941.67ed" is 941.67 here and 942 above.
            durExact: dm ? parseFloat(dm[1]) : 0,
            // Whether the Duration cell actually yielded a number. "0d" → known and zero (a milestone);
            // an empty or unreadable cell → NOT known, so nothing may conclude "milestone" from it.
            durKnown: !!dm,
            // The same, for the elapsed-duration column when the export carries one.
            edur: edm ? Math.round(parseFloat(edm[1])) : 0,
            edurKnown: !!edm,
            team: cTeam ? cellText(row.getCell(cTeam)).trim() : '',
            dept: cDept ? cellText(row.getCell(cDept)).trim() : '',
            note: cNote ? cellText(row.getCell(cNote)).trim() : '',
            // Responsible falls back to the legacy team column, so an older export still populates it.
            responsible: (cResp ? cellText(row.getCell(cResp)).trim() : '') || (cTeam ? cellText(row.getCell(cTeam)).trim() : ''),
            accountable: cAcc ? cellText(row.getCell(cAcc)).trim() : '',
            head: cHead ? cellText(row.getCell(cHead)).trim() : '',
        });
    }
    // Build WBS tree from outline levels; every row becomes a WBS node.
    const wbs = {};
    const stack = []; // stack[level] = node id at that level
    for (const raw of raws) {
        const id = String(raw.line);
        stack.length = raw.level; // parent is stack[level-1]
        const parent = raw.level > 0 ? stack[raw.level - 1] || null : null;
        wbs[id] = {
            id, name: raw.name, parent, seq: raw.line, depth: raw.level,
            pct: raw.pct, start: raw.s, finish: raw.f, // ASTA's own rolled-up values
            baselineStart: raw.bs, baselineFinish: raw.bf, milestone: raw.dur === 0, duration: raw.dur,
            // The export's own figure, unrounded, for the one column that is supposed to reproduce it.
            durationElapsed: raw.durKnown ? raw.durExact : null,
        };
        stack[raw.level] = id;
    }
    // a row is a "summary" if the next row is deeper (has children)
    const isSummary = (i) => i + 1 < raws.length && raws[i + 1].level > raws[i].level;
    const activities = [];
    raws.forEach((raw, i) => {
        if (isSummary(i))
            return; // leaves only become activities; summaries are WBS nodes
        const id = String(raw.line);
        const parent = wbs[id].parent; // roll leaf up into its containing WBS
        const status = raw.pct >= 100 ? 'Completed' : raw.pct > 0 ? 'In Progress' : 'Not Started';
        activities.push({
            id, name: raw.name, status, pct: raw.pct, wbs: parent,
            start: raw.s, finish: raw.f, baselineStart: raw.bs, baselineFinish: raw.bf,
            actualStart: null, actualFinish: null,
            totalFloat: 999, critical: false, // Excel export has no float → no critical flag
            dept: raw.dept, responsible: raw.responsible,
            accountable: raw.accountable, head: raw.head,
            approval: '', approvalStatus: '', // not in the schedule export
            milestone: raw.dur === 0,
            /**
             * THE 2 WEEK LOOK AHEAD'S MILESTONE TEST — BOTH duration columns must read zero.
             *
             *   Duration                 "0d"   AND
             *   Duration [Elapsed day]   "0d"
             *
             * ASTA keeps the two separately — working time against calendar time — so a row can read zero in
             * one and not the other. Requiring both is the stricter reading: a row is only removed from the
             * look-ahead when the export says it is a moment by every measure it offers.
             *
             * THE TEST IS OVER THE DURATION COLUMNS THE EXPORT ACTUALLY HAS. Every one of them must read
             * zero, so:
             *   • both columns present  → both must be "0"  (the strict case the rule was asked for)
             *   • only one present      → that one must be "0" — including the NH Homestay shape, where
             *                             "Duration [Elapsed day]" IS the duration column and `cEDur` and
             *                             `cDur` are the same column, so it is not demanded twice
             * Requiring a column the export does not have would make the test impossible to satisfy, so
             * NOTHING would ever be recognised as a milestone and every one of them would reappear on the
             * slide — silently, which is the opposite of what this rule is for. The look-ahead report states
             * which columns were checked for each run. Either way the decision is the export's own: the
             * task's NAME is never consulted, and a Duration cell that could not be read still yields false,
             * so a malformed export cannot delete real work.
             *
             * See the field's note in shared/types.ts for why this is kept apart from `milestone` above.
             */
            milestoneSource: raw.durKnown === true && raw.dur === 0
                && (cEDur === undefined || cEDur === cDur || (raw.edurKnown === true && raw.edur === 0)),
            note: raw.note,
        });
    });
    // Top WBS row is the programme name (often "<Project> Acquisition and Launch"); clean it so the
    // cover and all name matching use the real project name.
    const projectName = (0, match_1.cleanName)((raws[0]?.name || 'Project').replace(/\b\w/g, (m) => m.toUpperCase()));
    return {
        projectId: 'excel', projectName, dataDate: null, activities, wbs,
        // Both columns were present and distinct — i.e. the milestone test really did check two of them.
        hasElapsedDuration: cEDur !== undefined && cDurPlain !== undefined && cEDur !== cDurPlain,
    };
}
// Walk the WBS tree from a node id up to the root, returning names root→leaf.
function wbsPath(sch, id) {
    const parts = [];
    const seen = new Set();
    let cur = id;
    while (cur && sch.wbs[cur] && !seen.has(cur)) {
        seen.add(cur);
        parts.push(sch.wbs[cur].name);
        cur = sch.wbs[cur].parent;
    }
    return parts.reverse();
}
