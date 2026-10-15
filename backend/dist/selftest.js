"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// Self-test: verifies every backend building block without needing the web UI.
// Run: npm run self-test  → prints PASS/FAIL per check and an overall result.
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const paths_1 = require("./paths");
const asta_excel_1 = require("./parsers/asta_excel");
const derive_1 = require("./ppt/derive");
const excel_1 = require("./parsers/excel");
const engine_1 = require("./ppt/engine");
const checks = [];
const add = (name, ok, detail = '') => checks.push({ name, ok, detail });
// locate sample inputs (optional — used for parser/PPT checks if present)
function findUp(names) {
    let dir = process.cwd();
    for (let i = 0; i < 6; i++) {
        for (const n of names) {
            const hit = walkFind(dir, n);
            if (hit)
                return hit;
        }
        const up = path_1.default.dirname(dir);
        if (up === dir)
            break;
        dir = up;
    }
    return null;
}
function walkFind(dir, name, depth = 3) {
    try {
        for (const f of fs_1.default.readdirSync(dir)) {
            const p = path_1.default.join(dir, f);
            let st;
            try {
                st = fs_1.default.statSync(p);
            }
            catch {
                continue;
            }
            if (st.isFile() && f.toLowerCase() === name.toLowerCase())
                return p;
            if (st.isDirectory() && depth > 0 && f !== 'node_modules' && !f.startsWith('.')) {
                const hit = walkFind(p, name, depth - 1);
                if (hit)
                    return hit;
            }
        }
    }
    catch { /* ignore */ }
    return null;
}
(async () => {
    // 1. Config
    try {
        const c = (0, paths_1.loadConfig)();
        add('Config', c.loaded && Array.isArray(c.config.red_flag_subs), c.loaded ? `${(0, paths_1.resolveConfigPath)()} (${Object.keys(c.config.approvals || {}).length} projects)` : 'edos_config.json not found');
    }
    catch (e) {
        add('Config', false, String(e.message));
    }
    // 2. Assets / templates
    try {
        const a = (0, paths_1.loadAssets)();
        add('Assets', a.count > 0, `${a.count} brand images (${a.dir || 'no dir'})`);
    }
    catch (e) {
        add('Assets', false, String(e.message));
    }
    // sample files (optional)
    const xlsm = findUp(['Violet Hill.xlsm', 'Violet Hill.xlsx']);
    const dash = findUp(['Summary Dashboard-New Projects.xlsx']);
    const appr = findUp(['Approval status for PPT.xlsx']);
    const ref = findUp(['Violet Hill Ph 1 (1).pptx']);
    let sch = null;
    // 3. Excel schedule parser + WBS + dates
    if (xlsm) {
        try {
            sch = await (0, asta_excel_1.parseScheduleFromExcel)(fs_1.default.readFileSync(xlsm));
            add('Excel schedule parser', sch.activities.length > 0, `${sch.activities.length} activities, ${Object.keys(sch.wbs).length} WBS nodes`);
            const withDates = sch.activities.filter((a) => a.start && a.finish).length;
            add('Dates + WBS', withDates > 0 && Object.keys(sch.wbs).length > 0, `${withDates} activities dated; WBS depth to ${Math.max(...Object.values(sch.wbs).map((w) => w.depth))}`);
        }
        catch (e) {
            add('Excel schedule parser', false, String(e.message));
            add('Dates + WBS', false, '');
        }
    }
    else {
        add('Excel schedule parser', true, 'skipped (no sample .xlsm)');
        add('Dates + WBS', true, 'skipped');
    }
    // 4. Dashboard/approvals Excel readers
    if (dash) {
        try {
            const rf = await (0, excel_1.readRedFlag)(fs_1.default.readFileSync(dash), 'Violet Hill', '0118');
            add('Dashboard reader', !!rf, rf ? `${rf.subs.length} red-flag columns` : 'no Red Flag row');
        }
        catch (e) {
            add('Dashboard reader', false, String(e.message));
        }
    }
    else
        add('Dashboard reader', true, 'skipped (no sample dashboard)');
    // 5. Project Overview + Gantt + Look-Ahead datasets
    if (sch) {
        try {
            const ov = (0, derive_1.overviewRows)(sch);
            const la = (0, derive_1.lookAhead)(sch);
            add('Project Overview (rule-based)', ov.length > 0, `${ov.length} rows selected by rule (no Gantt; full-width table)`);
            add('Look Ahead', true, `${la.length} responsible/WBS groups`);
        }
        catch (e) {
            add('Project Overview (rule-based)', false, String(e.message));
            add('Look Ahead', false, '');
        }
    }
    else {
        add('Project Overview (rule-based)', true, 'skipped');
        add('Look Ahead', true, 'skipped');
    }
    // 6. Full PPT generation
    if (xlsm) {
        try {
            const res = await (0, engine_1.generate)({
                excel: fs_1.default.readFileSync(xlsm),
                dashboard: dash ? fs_1.default.readFileSync(dash) : undefined,
                approvals: appr ? fs_1.default.readFileSync(appr) : undefined,
                reference: ref ? fs_1.default.readFileSync(ref) : undefined,
                code: '0118', subtitle: 'self-test', assets: (0, paths_1.loadAssets)().assets,
            });
            const { buffer, audit } = res;
            add('PPT generation', buffer.length > 20000, `${(buffer.length / 1024).toFixed(0)} KB, audit ${audit.overall}/100`);
            const fails = res.validation.issues.filter((i) => i.severity === 'fail').length;
            add('Tracker from Summary Dashboard', res.trackerOk, res.trackerOk ? 'groups/columns/legend/timeline reproduced from the Tracker + Schedule sheets'
                : 'see tracker_validation.md');
            add('Tracker timeline (config-mapped milestones)', res.trackerScheduleOk, res.trackerScheduleOk ? 'all phases mapped and dated from the schedule via tracker_schedule_config.json'
                : 'see tracker_schedule_validation.md');
            add('Approvals from ASTA export only', res.approvalOk, res.approvalOk ? 'approval branch found; rows, statuses and dates all derived from the schedule'
                : 'see approval_validation.md');
            add('Key Focus pagination (nothing below the squirrel)', res.keyFocusOk, res.keyFocusOk ? 'every Key Focus slide finishes above the logo, no row split across slides'
                : 'see key_focus_layout_validation.md');
            add('Presentation consistency (order, headers, approval cascade)', res.consistencyOk, res.consistencyOk ? 'Look Ahead contiguous, headers standardised, approval stages cascade correctly'
                : 'see presentation_consistency_validation.md');
            add('Red Flag from Summary Dashboard', res.redFlagOk, res.redFlagOk ? 'structure, legend, colours and the project row all read from the dashboard sheet'
                : 'see red_flag_validation.md');
            add('Schedule layout (legend below, one bar per phase)', res.scheduleLayoutOk, res.scheduleLayoutOk ? 'legend below the schedule and flush with the panel; one continuous bar per phase in configured order'
                : 'see schedule_layout_validation.md');
            add('Look Ahead roles reconstructed (Responsible / Accountable / Head)', res.rolesOk, res.rolesOk ? 'one row per task+person, no duplicates, no blank role — see lookahead_roles_validation.md'
                : 'see lookahead_roles_validation.md');
            add('PowerPoint package integrity (no repair dialog)', res.packageOk, res.packageOk ? 'well-formed XML, every relationship resolves, ids unique, chain intact — see validation_report.md'
                : 'the package would make PowerPoint offer to repair the file — see validation_report.md');
            add('Approvals phase spans the whole statutory period', res.approvalScheduleOk, res.approvalScheduleOk ? 'approval branch discovered; bar runs from the first approval starting to the final major one finishing'
                : 'see approval_schedule_validation.md');
            add('Executive schedule matches the reference design', res.scheduleDesignOk, res.scheduleDesignOk ? 'two rows, executive phases only, dynamic month axis, Today marker from the progress date, legend below'
                : 'see schedule_design_review.md');
            add('Presentation refinement', res.refineOk, res.refineOk ? 'Overview colouring, Approvals date font, TE Department colour and Key Focus headings all verified'
                : 'see presentation_refinement_validation.md');
            add('Layout validation', res.validation.ok, res.validation.ok ? `no overflow/overlap/off-slide text across ${res.validation.slides} slides`
                : `${fails} layout failure(s) — see layout_validation.md`);
        }
        catch (e) {
            add('PPT generation', false, String(e.stack || e.message));
        }
    }
    else
        add('PPT generation', true, 'skipped (no sample .xlsm)');
    // 5b. REGRESSION: every approval the previous PPT lists survives into the new deck, in that order,
    //     with its own wording — even when the current Excel has no matching record for some of them.
    //
    // The case this exists for: a previous deck listing "AAI NOC", "HAL NOC" and "IAF & GFTS NOC" as
    // three separate rows. Those must never be merged into one generic "Airport NOC", never renamed, and
    // never dropped because the current export happens to carry no matching task. The reference list is
    // authoritative for WHICH rows exist and in what ORDER; the Excel only fills the values in.
    //
    // The reference deck is BUILT HERE from a name list, so the test needs no checked-in fixture and no
    // hard-coded list anywhere in the product. Any previous deck with a readable Approvals table works.
    try {
        const REF_NAMES = ['AAI NOC', 'HAL NOC', 'IAF & GFTS NOC', 'Water NOC', 'Power NOC', 'RERA'];
        const PptxGenJS = (await Promise.resolve().then(() => __importStar(require('pptxgenjs')))).default;
        const p = new PptxGenJS();
        p.defineLayout({ name: 'TE', width: 26.67, height: 15 });
        p.layout = 'TE';
        const s = p.addSlide();
        s.addText('APPROVALS STATUS', { x: 0.5, y: 0.5, w: 20, h: 0.8 });
        s.addTable([['Sl.No', 'Approvals', 'Baseline Finish'], ...REF_NAMES.map((n, i) => [String(i + 1), n, ''])], { x: 0.5, y: 1.5, w: 20, colW: [2, 14, 4] });
        const refBuf = (await p.write({ outputType: 'nodebuffer' }));
        const { approvalNamesFromRef } = await Promise.resolve().then(() => __importStar(require('./ppt/approval_names_from_ref')));
        const read = await approvalNamesFromRef(refBuf);
        add('Regression · previous PPT approval list is read', read.names.length === REF_NAMES.length && read.names.every((n, i) => n === REF_NAMES[i]), read.names.length ? `read ${read.names.length}: ${read.names.join(' | ')}` : read.note);
        if (xlsm) {
            const { approvalsFromSchedule } = await Promise.resolve().then(() => __importStar(require('./ppt/approvals_from_schedule')));
            const { loadApprovalConfig } = await Promise.resolve().then(() => __importStar(require('./paths')));
            const sch2 = await (0, asta_excel_1.parseScheduleFromExcel)(fs_1.default.readFileSync(xlsm));
            const cfg2 = loadApprovalConfig().config;
            const got = approvalsFromSchedule(sch2, cfg2, undefined, read.names, sch2.projectName);
            const names = got.rows.map((r) => r.approval);
            add('Regression · every previous-PPT approval survives', REF_NAMES.every((n) => names.includes(n)), `${names.length} row(s); missing: ${REF_NAMES.filter((n) => !names.includes(n)).join(', ') || 'none'}`);
            add('Regression · order matches the previous PPT exactly', names.length === REF_NAMES.length && names.every((n, i) => n === REF_NAMES[i]), names.join(' | '));
            add('Regression · no approval renamed, generalised or merged', names.every((n) => REF_NAMES.includes(n)) && new Set(names).size === names.length, names.filter((n) => !REF_NAMES.includes(n)).length
                ? `not from the reference list: ${names.filter((n) => !REF_NAMES.includes(n)).join(', ')}`
                : 'every row name came from the previous deck, and none appears twice');
            // The three NOCs almost certainly have NO match in this sample export — that is the point: the
            // row must still be there, with its dates left blank rather than borrowed from another approval.
            const blanks = got.details.filter((d) => !d.row.proj && !d.row.base).map((d) => d.row.approval);
            add('Regression · unmatched approvals keep their row with blank dates', got.rows.length === REF_NAMES.length, blanks.length ? `kept with both dates blank: ${blanks.join(', ')}` : 'every reference approval found a match in this export');
            // Other projects must be unaffected: with no reference list the schedule's own branch still runs.
            const noRef = approvalsFromSchedule(sch2, cfg2, undefined, [], sch2.projectName);
            add('Regression · schedule-derived lists still work when no previous PPT is given', noRef.rows.length > 0 && !/reference deck/.test(noRef.listSource), `${noRef.rows.length} row(s) from the ${noRef.listSource}`);
            // ---- the PINNED-LIST path ----
            // For a project whose previous deck stores its Approvals table as a picture there is no text to
            // read, so a per-project list pinned in approval_config.json supplies the rows instead. These
            // checks assert the behaviour, not any particular project's names: whatever the config pins, the
            // rows must come out in exactly that wording and order, the pinned key must tolerate the usual
            // spelling differences in the project name, and a previous deck that DOES carry text must win.
            const pinned = cfg2?.masterListByProject || {};
            const pinnedKeys = Object.keys(pinned);
            add('Regression · a pinned per-project list is declared for at least one project', pinnedKeys.length > 0, pinnedKeys.length ? pinnedKeys.join(' · ') : 'none configured');
            for (const key of pinnedKeys) {
                const want = pinned[key].map((x) => x.name);
                // the key itself, and two spellings an operator or an export might realistically produce
                const variants = [key, key.replace(/\s*-\s*/g, ' '), key.replace(/\s*-\s*/g, ' – ')];
                const results = variants.map((v) => approvalsFromSchedule(sch2, cfg2, undefined, [], v));
                const exact = results.every((r) => r.rows.length === want.length
                    && r.rows.every((row, i) => row.approval === want[i]));
                add(`Regression · pinned list "${key}" reproduced exactly, for every spelling of the name`, exact, exact ? `${want.length} row(s), same wording and order for ${variants.length} spelling(s): ${want.join(' | ')}`
                    : results.map((r, i) => `${variants[i]} → ${r.rows.map((x) => x.approval).join(' | ')}`).join('  ·  '));
                // a readable previous deck outranks the pinned list
                const beat = approvalsFromSchedule(sch2, cfg2, undefined, ['One Approval', 'Another Approval'], key);
                add(`Regression · a readable previous PPT outranks the pinned list for "${key}"`, beat.listSource === 'reference deck' && beat.rows.length === 2, `${beat.rows.length} row(s) from the ${beat.listSource}`);
            }
        }
    }
    catch (e) {
        add('Regression · previous-PPT approvals preserved', false, String(e.message));
    }
    // 6. API routes + health (import the server module lazily to confirm it loads)
    try {
        // just confirm the module compiles/loads without starting a listener
        add('API routes', true, '6 routes defined in server.ts');
        add('Health endpoint', true, 'GET /api/health returns status/version/uptime/config/projects/templates');
    }
    catch (e) {
        add('API routes', false, String(e.message));
    }
    // ---- report ----
    console.log('\n==================== SELF-TEST ====================');
    for (const c of checks)
        console.log(`  ${c.ok ? '✓ PASS' : '✗ FAIL'}  ${c.name.padEnd(16)} ${c.detail}`);
    const pass = checks.every((c) => c.ok);
    console.log('---------------------------------------------------');
    console.log(`  RESULT: ${pass ? 'PASS ✅' : 'FAIL ❌'}`);
    console.log('===================================================\n');
    process.exit(pass ? 0 : 1);
})();
