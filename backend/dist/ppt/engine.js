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
exports.lastRun = void 0;
exports.failureReport = failureReport;
exports.generate = generate;
const pptxgenjs_1 = __importDefault(require("pptxgenjs"));
const fs_1 = __importDefault(require("fs"));
const asta_excel_1 = require("../parsers/asta_excel");
const X = __importStar(require("../parsers/excel"));
const S = __importStar(require("./slides"));
const overview_1 = require("./overview");
const derive_1 = require("./derive");
const roles_1 = require("./roles");
const theme_1 = require("./theme");
const postmerge_1 = require("./postmerge");
const package_validate_1 = require("./package_validate");
const audit_1 = require("../audit");
const layout_validate_1 = require("./layout_validate");
const refine_validate_1 = require("./refine_validate");
const red_flag_validate_1 = require("./red_flag_validate");
const consistency_validate_1 = require("./consistency_validate");
const key_focus_validate_1 = require("./key_focus_validate");
const diagnostics_1 = require("./diagnostics");
const match_1 = require("../parsers/match");
const overview_select_1 = require("./overview_select");
const typography_1 = require("./typography");
const paths_1 = require("../paths");
const tracker_template_1 = require("../parsers/tracker_template");
const tracker_1 = require("./tracker");
const schedule_phases_1 = require("./schedule_phases");
const approval_phase_1 = require("./approval_phase");
const approvals_from_schedule_1 = require("./approvals_from_schedule");
const approval_names_from_ref_1 = require("./approval_names_from_ref");
const blocks_1 = require("./blocks");
const paths_2 = require("../paths");
function loadConfig() { return (0, paths_2.loadConfig)().config; }
/**
 * The last run's diagnostics, so a FAILED generation can still produce error_report.md. Set as soon as
 * the tracker exists and kept current as facts are discovered.
 */
exports.lastRun = null;
/** Build error_report.md for a generation that threw, using the last run's recorded stages. */
function failureReport(e) {
    const isStage = e instanceof diagnostics_1.StageError;
    const info = isStage ? e.info : (0, diagnostics_1.describeError)(e);
    const stage = isStage ? e.stage : (exports.lastRun?.tk.stages.slice(-1)[0]?.name || 'Unknown stage');
    const md = exports.lastRun
        ? (0, diagnostics_1.errorReport)({ project: exports.lastRun.project, code: exports.lastRun.code, tracker: exports.lastRun.tk, names: exports.lastRun.names, slides: exports.lastRun.slides, match: match_1.matchScore, fatal: info, fatalStage: stage })
        : ['# error_report.md', '', `## Result: **FAILED** ❌`, '', `Stage: **${stage}**`, '',
            `\`${info.type}: ${info.message}\` — ${info.file}:${info.line} in ${info.fn}`, '',
            '**Suggested fix** — ' + info.hint, '', '```', info.stack, '```'].join('\n');
    return { md, stage, info };
}
async function generate(inp) {
    const log = inp.log || (() => { });
    // Every stage below is tracked. Optional sections never abort the run — they are left blank and the
    // reason is recorded; only a failure that makes the .pptx impossible stops generation, and then the
    // caller gets the stage, file, function, line and stack rather than "Generation failed".
    const tk = new diagnostics_1.StageTracker(log);
    const names = [];
    const refSlides = [];
    exports.lastRun = { tk, names, slides: refSlides, project: '(not detected yet)', code: inp.code || '' };
    let fatal = null;
    const sch = await tk.run('Loading ASTA Excel export', false, async () => {
        const parsed = await (0, asta_excel_1.parseScheduleFromExcel)(inp.excel);
        if (!parsed.activities.length) {
            throw new Error('the export parsed but contains 0 activities — this is not a full ASTA schedule export');
        }
        return parsed;
    });
    // Any further block exports of the SAME project, each parsed on its own. They are never merged:
    // each keeps its own activity tree so no block's data can appear under another's heading.
    const extraSchedules = [];
    for (const [i, buf] of (inp.extraExcels || []).entries()) {
        const parsed = await tk.run(`Loading block export ${i + 2}`, false, async () => {
            const q = await (0, asta_excel_1.parseScheduleFromExcel)(buf);
            if (!q.activities.length)
                throw new Error(`block export ${i + 2} parsed but contains 0 activities`);
            return q;
        });
        extraSchedules.push(parsed);
    }
    const code = inp.code || '';
    exports.lastRun.project = sch.projectName;
    tk.fact('Project name', sch.projectName);
    tk.fact('Project code', code || '(none supplied)');
    tk.fact('Activities parsed', sch.activities.length);
    tk.fact('WBS depth detected', Math.max(0, ...sch.activities.map((a) => (0, asta_excel_1.wbsPath)(sch, a.wbs).length)));
    tk.fact('Schedule data date', sch.dataDate || '(none)');
    tk.fact('Summary Dashboard uploaded', inp.dashboard ? `yes (${(inp.dashboard.length / 1048576).toFixed(1)} MB)` : 'no');
    tk.fact('Previous PPT uploaded', inp.reference ? `yes (${(inp.reference.length / 1048576).toFixed(1)} MB)` : 'no');
    tk.note('Loading ASTA Excel export', `${sch.projectName} · ${sch.activities.length} activities`);
    names.push({ source: 'ASTA Excel export', name: sch.projectName, normalised: (0, match_1.cleanName)(sch.projectName) });
    const cfg = loadConfig();
    const reportDate = inp.reportDate || sch.dataDate || new Date().toISOString().slice(0, 10);
    const theme = await tk.run('Reading theme colours from the previous PPT', true, () => (0, theme_1.extractTheme)(inp.reference), await (0, theme_1.extractTheme)(undefined));
    log(`Theme accent: #${theme.accent}`);
    const assets = inp.assets || {};
    /** A deck on the standard TE canvas. Used for the real deck, and for the throwaway pass that counts
     *  a block-split section's pages before the pages are lettered. */
    const newDeck = () => {
        const d = new pptxgenjs_1.default();
        d.defineLayout({ name: 'TE', width: 26.67, height: 15 });
        d.layout = 'TE';
        return d;
    };
    const pptx = newDeck();
    S.cover(pptx, sch, code, inp.subtitle || '', reportDate, theme, assets);
    // ---- the Contents slide is created here but WRITTEN at the end ----
    // Its position (second, after the cover) has to be fixed now; its content cannot be known until every
    // section has been generated. pptxgenjs serialises on write(), so the slide is created empty and filled
    // once the section list is complete.
    const fillContents = S.contents(pptx, theme, assets);
    // ---- the section register ----
    // Every section that is actually generated calls section(title) and gets its NUMBER back. The number
    // is therefore always the section's real position, and the Contents list is exactly this register — the
    // two cannot disagree, and a section that is skipped costs its number rather than leaving a hole.
    // A continuation slide ("(cont.)") is part of its section and does not register again.
    const sections = [];
    const section = (title) => {
        sections.push({ n: sections.length + 1, title });
        return sections.length;
    };
    /** Take a section back out of the register when it turned out to have nothing to show. Only the
     *  LAST registered section can be withdrawn, which is the only case that arises: a section is
     *  registered immediately before it is rendered. Withdrawing it lets the sections after it move up
     *  rather than leaving a gap in the numbering and a dead line in Contents. */
    const unregisterSection = (title) => {
        const last = sections[sections.length - 1];
        if (last && last.title === title)
            sections.pop();
    };
    // ---- which merge-managed sections will survive? ----
    // Land Acquisition Status and Master Plan are placeholders that the merge step replaces with the
    // reference deck's own slides — and DELETES outright when the reference has no such slide. Deciding
    // that here, before generating, is what stops Contents advertising a section the deck will not contain.
    const refTitles = inp.reference ? await (0, postmerge_1.refSlideTitles)(inp.reference).catch(() => []) : [];
    // ---- BLOCKS: one project, several reporting scopes ----
    // Each export names its own block in its root row; the project's own name is what is left once that
    // qualifier is removed. The previous deck decides the block ORDER and WHICH sections are reported
    // per block — nothing about any project or block is written into the program. With one export, or
    // exports that name no block, `blocks` has a single entry and every branch below is inert.
    const BLOCK_SPLITTABLE = ['Project Overview', 'Approvals Status', '2 Week Look Ahead Plan'];
    const blocks = (0, blocks_1.buildBlocks)([sch, ...extraSchedules], refTitles, BLOCK_SPLITTABLE);
    const multiBlock = blocks.length > 1 && blocks.every((b) => !!b.label);
    let splitSections = new Set();
    if (blocks.length > 1 && !multiBlock) {
        log(`  ! ${blocks.length} exports were supplied but they do not all name a block in their root row ` +
            `(${blocks.map((b) => `"${b.rawName}"`).join(', ')}) — the deck is built from the first export alone.`);
    }
    if (multiBlock) {
        const cp = (0, blocks_1.commonProjectName)(blocks);
        splitSections = (0, blocks_1.blockSplitSections)(refTitles, blocks, BLOCK_SPLITTABLE);
        log(`Blocks: ${blocks.length} export(s) for one project — ${blocks.map((b) => b.label).join(' → ')} ` +
            `(order taken from the previous deck)`);
        if (!cp.agreed) {
            log(`  ! the exports disagree on the project name once the block is removed (${cp.seen.join(' / ')}); ` +
                `using "${cp.name}"`);
        }
        // The project's own name, block qualifier removed. This is what the Summary Dashboard holds — ONE
        // row for the project, not one per block — and Key Focus Area matches it strictly, so an export
        // called "… [North Block]" would otherwise match nothing and that section would come out empty.
        if (sch.projectName !== cp.name) {
            log(`  · project identity for the dashboard and the cover: "${cp.name}" (from "${sch.projectName}")`);
            sch.projectName = cp.name;
            exports.lastRun.project = cp.name;
        }
        log(`  · reported per block, as the previous deck does: ` +
            (splitSections.size ? [...splitSections].join(', ') : 'none — the previous deck splits no section by block'));
        const single = BLOCK_SPLITTABLE.filter((x) => !splitSections.has(x));
        if (single.length)
            log(`  · reported once for the whole project: ${single.join(', ')}`);
    }
    /** Blocks to render a section for: every block when the section is split, else the primary only. */
    const blocksFor = (title) => (multiBlock && splitSections.has(title)) ? blocks : [blocks[0] || { label: '', project: sch.projectName, rawName: sch.projectName, sch }];
    // The reference's CONTENTS slide lists every section title, so a plain text search finds "Master Plan"
    // there and reports the section as present when the deck has no such slide. Contents is excluded and
    // the match is confined to the heading, which is the same discipline mergeStaticSlides applies when it
    // chooses which reference slide to copy — the two must agree or Contents and the deck diverge.
    const refHas = (re) => refTitles.some((t) => !/^\s*contents\b/i.test(t) && re.test(t.slice(0, 80)));
    // With no reference deck nothing is merged, so the placeholder stays and the section exists.
    const willHaveLand = !!inp.land || !inp.reference || refHas(/land acquisition/i);
    const willHaveMasterPlan = !inp.reference || refHas(/master plan/i);
    if (inp.reference) {
        log(`  · sections from the reference deck: Land Acquisition ${willHaveLand ? 'present' : 'ABSENT — omitted from the deck and from Contents'}` +
            ` · Master Plan ${willHaveMasterPlan ? 'present' : 'ABSENT — omitted from the deck and from Contents'}`);
    }
    // ---- Red Flag Summary: generated from the Summary Dashboard, the single source of truth ----
    // Structure (groups, sub-headings, order), legend and colours all come from its "Red Flag Summary"
    // sheet; the one data row is the selected project's, found by name/code. It is NOT copied from the
    // previous PPT, and the ASTA export is not used. If the project cannot be found, the slide is drawn
    // with the dashboard's structure and empty indicators, and the gap is stated in red_flag_validation.md
    // — nothing is derived from the schedule as a substitute.
    const rfRead = await tk.run('Reading Red Flag Summary from the Summary Dashboard', true, async () => {
        if (!inp.dashboard)
            return null;
        const got = await X.readRedFlag(inp.dashboard, sch.projectName, code);
        if (got)
            names.push({ source: 'Summary Dashboard', name: got.project, normalised: (0, match_1.cleanName)(got.project) });
        tk.note('Reading Red Flag Summary from the Summary Dashboard', got ? `row ${got.rowNumber} · ${got.subs.length} columns` : 'project not found on the sheet — indicators left blank');
        tk.fact('Summary Dashboard row', got ? `${got.rowNumber} ("${got.project}")` : 'not found');
        return got;
    }, null);
    const rfv = (0, red_flag_validate_1.redFlagValidation)(rfRead, sch.projectName, code, !!inp.dashboard);
    const redFlagMd = rfv.md, redFlagOk = rfv.ok;
    const rf = rfRead || (0, derive_1.matrixFromAsta)(sch, cfg, code, false);
    S.matrix(pptx, rf, section('Red Flag Summary'), 'Red Flag Summary', 'color', theme, assets);
    if (rfRead)
        log(`  · dashboard row ${rfRead.rowNumber} ("${rfRead.project}"), ` +
            `${rfRead.subs.length} columns in ${rfRead.macro.length} groups`);
    log(`  · red flag validation: ${redFlagOk ? 'PASS' : 'FAIL'}`);
    const kf = await tk.run('Reading Key Focus Area from the Summary Dashboard', true, async () => {
        const got = inp.dashboard ? await X.readKeyFocus(inp.dashboard, sch.projectName, code) : null;
        tk.note('Reading Key Focus Area from the Summary Dashboard', `${got?.length || 0} row(s)`);
        return got;
    }, null);
    S.keyFocus(pptx, sch, code, kf || [], section('Key Focus Area'), theme, assets);
    // Row caps (Critical 24 / At Risk 28) are a formatting correction, not a data change — record
    // every dropped row so layout_validation.md states exactly what was trimmed and why.
    const corrections = [];
    for (const [g, dropped] of Object.entries(S.kfDropped)) {
        if (dropped > 0)
            corrections.push(`Key Focus Area — **${g}**: ${dropped} row(s) beyond the ${S.KF_ROW_CAP[g]}-row limit were omitted, ` +
                `keeping the highest-priority rows. The font was NOT reduced to fit them.`);
    }
    // ---- Tracker & Schedule: built ENTIRELY from the Summary Dashboard template ----
    // The previous PPT is not a data source here; it is only a visual formatting reference.
    log('Tracker & Schedule (generated from the Summary Dashboard Tracker + Schedule sheets)…');
    let trackerTpl = null;
    let trackerData = null;
    let trackerMd = '';
    let trackerOk = false;
    let trackerScheduleMd = '';
    let trackerScheduleOk = false;
    let scheduleLayoutMd = '';
    let scheduleLayoutOk = false;
    let scheduleReviewMd = '# schedule_layout_review.md\n\nNo Summary Dashboard was supplied, so no schedule was drawn.\n';
    let scheduleDesignMd = '# schedule_design_review.md\n\n**FAIL** — no Summary Dashboard was supplied, so the executive schedule was not drawn.\n';
    let scheduleDesignOk = false;
    let approvalScheduleMd = '# approval_schedule_validation.md\n\n**FAIL** — no Summary Dashboard was supplied, so the executive schedule (and its Approvals phase) was not built.\n';
    let approvalScheduleOk = false;
    let packageMd = '# validation_report.md\n\nThe package was never written, so it could not be validated.\n';
    let packageOk = false;
    let rolesMd = '# lookahead_roles_validation.md\n\nThe 2 Week Look Ahead was not generated.\n';
    let rolesOk = false;
    const schedCfgRes = (0, paths_1.loadTrackerScheduleConfig)();
    if (!schedCfgRes.loaded)
        log('  ! tracker_schedule_config.json not found — timeline phase mapping unavailable.');
    trackerTpl = await tk.run('Reading the Tracker template from the Summary Dashboard', true, async () => (inp.dashboard ? await (0, tracker_template_1.readTrackerTemplate)(inp.dashboard) : null), null);
    if (trackerTpl && trackerTpl.columns.length)
        try {
            // The project's OWN status row on the Tracker sheet. The layout came from that sheet already; the
            // statuses used to be recomputed from the schedule by keyword matching, which is how the slide
            // could contradict the dashboard it was drawn from. Matched with the SAME code-first project
            // identification the Red Flag Summary uses, so one project can never show another's row.
            const trackerStatus = await tk.run('Reading the Tracker statuses for this project', true, async () => (inp.dashboard ? await X.readTrackerStatuses(inp.dashboard, sch.projectName, code) : null), null);
            if (trackerStatus) {
                log(`  · Tracker statuses: read ${trackerStatus.byCol.size} indicator(s) from the Summary Dashboard row` +
                    ` "${trackerStatus.code || '—'} ${trackerStatus.project}" — the schedule is used only where that row is blank`);
            }
            else if (inp.dashboard) {
                log('  ! Tracker statuses: no Summary Dashboard Tracker row matched this project — indicators fall back ' +
                    'to the schedule rather than showing another project\'s row');
            }
            trackerData = await tk.run('Building Tracker & Schedule', true, () => (0, tracker_1.buildTrackerData)(sch, trackerTpl, cfg, [inp.area || '', inp.subtitle || ''].join(' '), schedCfgRes.config, trackerStatus), null);
            if (trackerData) {
                const fromDash = trackerData.indicators.filter((i) => i.source === 'dashboard').length;
                log(`  · Tracker indicator sources: ${fromDash} from the dashboard row, ` +
                    `${trackerData.indicators.length - fromDash} derived from the schedule`);
            }
            tk.fact('Tracker indicators', trackerData ? trackerData.indicators.length : 0);
            if (!trackerData)
                throw new Error('__skip_tracker__');
            // reportDate drives the Today marker on the executive timeline — the same date the rest of the
            // deck is stated "as at", so the line can never disagree with the header.
            const heads = (0, tracker_1.trackerSlide)(pptx, sch, trackerTpl, trackerData, code, inp.area || '', section('Tracker & Schedule'), theme, assets, reportDate);
            const v = (0, tracker_1.trackerValidation)(sch, trackerTpl, trackerData, true, heads);
            trackerMd = v.md;
            trackerOk = v.ok;
            // Milestone-mapping / date-provenance report for the timeline (configuration-driven).
            const colours = {};
            for (const ph of trackerTpl.phases)
                colours[ph.label] = { actual: ph.actual, baseline: ph.baseline };
            const sv = (0, schedule_phases_1.schedulePhaseReport)(sch, trackerData.resolved, schedCfgRes.config, trackerData.scopedPhase, colours);
            trackerScheduleMd = sv.md;
            trackerScheduleOk = sv.ok;
            log(`  · schedule phases: ${trackerData.resolved.filter((r) => r.aStart && r.aFinish).length}/${trackerData.resolved.length} dated` +
                `; unresolved milestones: ${trackerData.resolved.reduce((n, r) => n + r.missing.length, 0)}`);
            log(`  · tracker schedule validation: ${sv.ok ? 'PASS' : 'FAIL'}`);
            if (trackerData.scopedPhase)
                log(`  · scoped to phase: ${trackerData.scopedPhase} (${trackerData.scopedActivities} activities)`);
            log(`  · ${trackerData.indicators.length} indicators across ${trackerTpl.groups.length} groups; ` +
                `timeline phases dated: ${trackerData.phases.filter((p) => p.aStart || p.bStart).length}/${trackerData.phases.length}`);
            log(`  · tracker validation: ${v.ok ? 'PASS' : 'FAIL'}`);
            // Schedule-block refinement report. The expected phase order comes from the CONFIG file, so no
            // phase name is hard-coded here either.
            const phaseOrder = (schedCfgRes.config.phases || []).map((ph) => ph.label);
            const slv = (0, tracker_1.scheduleLayoutValidation)(trackerData, sch.projectName, phaseOrder, true);
            scheduleLayoutMd = slv.md;
            scheduleLayoutOk = slv.ok;
            scheduleReviewMd = (0, tracker_1.scheduleLayoutReview)(sch.projectName, tracker_1.scheduleLayout?.months ?? 0, true);
            const sdr = (0, tracker_1.scheduleDesignReview)(sch.projectName, trackerData, phaseOrder, true);
            scheduleDesignMd = sdr.md;
            scheduleDesignOk = sdr.ok;
            log(`  · schedule design review: ${sdr.ok ? 'PASS' : 'FAIL'}`);
            // The Approvals bar spans the whole statutory approval period, derived from the approval branch.
            const apPhase = trackerData.resolved.find((r) => r.approval);
            if (apPhase?.approval) {
                const apCfg = (schedCfgRes.config.phases || []).find((x) => x.id === apPhase.id)?.approvalBranch || {};
                const asv = (0, approval_phase_1.approvalScheduleValidation)(sch.projectName, apPhase.approval, apCfg);
                approvalScheduleMd = asv.md;
                approvalScheduleOk = asv.ok;
                const a = apPhase.approval;
                log(`  · approvals phase: ${a.branch ? a.branch.name : 'branch not found'} — ${a.aStart || '—'} → ${a.aFinish || '—'}` +
                    ` (${a.items.filter((i) => i.used).length} approvals, rule ${a.finishPriority || '-'}): ${asv.ok ? 'PASS' : 'FAIL'}`);
            }
            else {
                approvalScheduleMd = '# approval_schedule_validation.md\n\n**FAIL** — the Approvals phase is not configured for ' +
                    'branch discovery, so the bar could not be derived from the approval branch. Add `approvalBranch` to the ' +
                    'approvals phase in tracker_schedule_config.json.\n';
            }
            log(`  · schedule layout validation: ${slv.ok ? 'PASS' : 'FAIL'}`);
        }
        catch (e) {
            // Any Tracker problem leaves that slide's data blank instead of stopping the deck.
            if (String(e?.message) !== '__skip_tracker__')
                tk.stages.push({ name: 'Rendering Tracker & Schedule', optional: true, ok: false, ms: 0, error: (0, diagnostics_1.describeError)(e) });
            trackerMd = trackerMd || '# tracker_validation.md\n\n**SKIPPED** — see error_report.md.\n';
        }
    else {
        // No dashboard supplied → we cannot reproduce the template, so fall back to the schedule-derived
        // matrix rather than silently emitting a wrong-looking slide.
        log('  ! No Summary Dashboard Tracker sheet available — falling back to the schedule-derived matrix.');
        const tr = (0, derive_1.matrixFromAsta)(sch, cfg, code, true);
        S.matrix(pptx, tr, section('Tracker & Schedule'), 'Tracker & Schedule', 'symbol', theme, assets);
        trackerMd = '# tracker_validation.md\n\n**FAIL** — no Summary Dashboard was supplied, so the ' +
            'Tracker & Schedule slide could not be generated from its template. Upload Summary Dashboard.xlsx.\n';
        trackerScheduleMd = '# tracker_schedule_validation.md\n\n**FAIL** — no Summary Dashboard was supplied, ' +
            'so the timeline phases could not be resolved.\n';
        scheduleLayoutMd = '# schedule_layout_validation.md\n\n**FAIL** — no Summary Dashboard was supplied, so the ' +
            'schedule timeline and its legend were not drawn.\n';
    }
    const ovRows = await tk.run('Generating Project Overview', true, () => {
        // ONE section number for the whole of Project Overview, whether the project reports as one scope
        // or several blocks. When it is split, each block is rendered from ITS OWN schedule through the
        // SAME renderer — same table, same Gantt, same 73:27 split, same palette, same pagination — and
        // the pages letter continuously across the blocks: South 4A 4B 4C, then North 4D 4E 4F.
        //
        // The page count is not known until a block has been laid out, so the blocks are rendered twice:
        // once to count (into a throwaway deck) and once for real. That costs a little time and keeps the
        // lettering right without teaching the renderer about blocks.
        const list = blocksFor('Project Overview');
        const no = section('Project Overview');
        let total = 0;
        if (list.length > 1) {
            const scratch = newDeck();
            for (const b of list)
                total += (0, overview_1.projectOverview)(scratch, b.sch, theme, assets, reportDate, no).pages;
        }
        let rows = 0, from = 0;
        for (const b of list) {
            const r = (0, overview_1.projectOverview)(pptx, b.sch, theme, assets, reportDate, no, list.length > 1
                ? { title: (0, blocks_1.blockTitle)('Project Overview', b.label), letterFrom: from, letterTotal: total }
                : {});
            rows += r.rows;
            from += r.pages;
            if (list.length > 1)
                log(`  · Project Overview [${b.label}]: ${r.rows} row(s) across ${r.pages} page(s)`);
        }
        tk.note('Generating Project Overview', `${rows} row(s) selected by rule across ${from} page(s)` +
            (list.length > 1 ? ` in ${list.length} block(s)` : ''));
        return rows;
    }, 0);
    // Land Acquisition: an uploaded IMAGE replaces the slide outright; an uploaded .pptx is merged
    // later; otherwise the placeholder is swapped for the reference deck's slide by the merge step.
    const landIsPptx = !!inp.land && inp.land.length > 4 && inp.land[0] === 0x50 && inp.land[1] === 0x4B;
    const landIsImage = !!inp.land && !landIsPptx;
    if (landIsImage) {
        log('Land Acquisition Status… (from uploaded image)');
        S.imageSlide(pptx, 'Land Acquisition Status', section('Land Acquisition Status'), inp.land, theme, assets);
    }
    else if (willHaveLand) {
        S.placeholder(pptx, 'Land Acquisition Status', section('Land Acquisition Status'), 'Static slide — copied verbatim from the uploaded reference deck by the merge step (see docs).', theme, assets);
    }
    if (willHaveMasterPlan) {
        S.placeholder(pptx, 'Master Plan', section('Master Plan'), 'Static slide — copied verbatim from the uploaded reference deck if present.', theme, assets);
    }
    // ---- Approvals Status: derived from the ASTA export ALONE (no Approval Status.xlsx) ----
    // Slide format, columns, fonts and colours are untouched — only the data source changed.
    log('Approvals Status (from the ASTA Excel export only)…');
    const apCfgRes = (0, paths_1.loadApprovalConfig)();
    if (!apCfgRes.loaded)
        log('  ! approval_config.json not found — approval matching unavailable.');
    const scopedIds = new Set((trackerData ? (0, tracker_1.scopeToPhase)(sch, [inp.area || '', inp.subtitle || ''].join(' ')).activities : []).map((a) => a.id));
    // WHICH approvals appear, and in what order, comes from the previous/reference deck's own Approvals
    // slide. The Excel supplies the DATES and the ticks. With no reference deck the schedule's approval
    // branch decides the list, exactly as before.
    // ---- one Approvals section, rendered once per block when the previous deck splits it ----
    // Each block reads ITS OWN approval list from the previous deck's own slide for that block, and its
    // dates from ITS OWN export. Nothing crosses: North can neither take South's row list nor its dates.
    const apList = blocksFor('Approvals Status');
    const apNo0 = section('Approvals Status');
    let approvalMd = '', approvalOk = true, apRes = null, apv = null;
    const apDone = [];
    const apReports = [];
    for (const [bi, b] of apList.entries()) {
        const tag = apList.length > 1 ? ` [${b.label}]` : '';
        const refAp = inp.reference
            ? await (0, approval_names_from_ref_1.approvalNamesFromRef)(inp.reference, apList.length > 1 ? b.label : '')
            : { names: [], slide: '', note: 'no reference deck supplied' };
        log(`  · approval list${tag}: ${refAp.note}`);
        if (inp.reference && !refAp.names.length) {
            log('  ! WARNING: a previous PPT was uploaded but NO approval list could be read from it, so the');
            log('    row list falls back to this project\'s own approval branch. The row NAMES and ORDER will');
            log('    therefore be the schedule\'s, not the previous deck\'s. See approval_validation.md.');
        }
        const scoped = new Set((trackerData && bi === 0
            ? (0, tracker_1.scopeToPhase)(b.sch, [inp.area || '', inp.subtitle || ''].join(' ')).activities : []).map((a) => a.id));
        const got = await tk.run(`Reading Approvals from the ASTA export${tag}`, true, () => {
            const r = (0, approvals_from_schedule_1.approvalsFromSchedule)(b.sch, apCfgRes.config, bi === 0 ? scoped : undefined, refAp.names, sch.projectName);
            tk.note(`Reading Approvals from the ASTA export${tag}`, `${r.details.length} approval(s) under "${r.branchesFound.map((x) => x.name).join(', ') || 'no branch found'}"`);
            return r;
        }, { details: [], rows: [], branchesFound: [], branchExcluded: [], missingExpected: [], ok: false });
        const v = (0, approvals_from_schedule_1.approvalValidation)(b.sch, got, trackerData?.scopedPhase || null, { supplied: !!inp.reference, namesRead: refAp.names.length, note: refAp.note });
        apReports.push(apList.length > 1 ? `# Approvals — ${b.label}\n\n${v.md}` : v.md);
        approvalOk = approvalOk && v.ok;
        if (bi === 0) {
            apRes = got;
            apv = v;
        }
        log(`  · approval branch${tag}: ${got.branchesFound.map((x) => x.name).join(', ') || '(none found)'}`);
        log(`  · row list${tag} from the ${got.listSource}: ${got.listNote}`);
        log(`  · ${got.details.length} approvals${tag}; dates from schedule: ` +
            `${got.details.filter((d) => d.row.proj).length}/${got.details.length}`);
        log(`  · approval validation${tag}: ${v.ok ? 'PASS' : 'FAIL'}`);
        if (!got.rows.length) {
            log(`  · Approvals Status${tag}: no approvals — nothing rendered for this block`);
            continue;
        }
        const authoritative = /reference deck|configured list/.test(got.listSource || '');
        log(`  · row order${tag}: ${authoritative
            ? `preserved exactly as the ${got.listSource} lists it`
            : 'sorted by Actual / Projected Finish (the list is schedule-derived, so it has no order of its own)'}`);
        S.approvals(pptx, got.rows, apNo0, theme, assets, authoritative, apList.length > 1
            ? { title: (0, blocks_1.blockTitle)('Approvals Status', b.label), letterFrom: apDone.length, letterTotal: apList.length }
            : {});
        apDone.push({ label: b.label, rows: got.rows.length });
    }
    approvalMd = apReports.join('\n\n---\n\n');
    const ap = apRes ? apRes.rows : [];
    // An Approvals section with nothing in it anywhere is not emitted at all — the same rule Land
    // Acquisition and Master Plan follow. Its number is given back so the sections after it move up.
    if (!apDone.length) {
        unregisterSection('Approvals Status');
        log('  · Approvals Status: no approvals for this project — section omitted from the deck and from Contents');
    }
    const laGroups = await tk.run('Generating the 2 Week Look Ahead', true, () => {
        // One section number; when the previous deck splits it, each block's OWN export is filtered and
        // grouped by the SAME 14-day logic and the pages letter on across the blocks — 8A–8E South, then
        // 8F–8G North. The 2-week filter is never applied to a combined dataset: each block is filtered
        // on its own, so no task can be attributed to the wrong block.
        const list = blocksFor('2 Week Look Ahead Plan');
        const no = section('2 Week Look Ahead Plan');
        const perBlock = list.map((b) => ({ b, g: (0, derive_1.lookAhead)(b.sch, reportDate) }));
        let total = 0;
        if (list.length > 1) {
            const scratch = newDeck();
            for (const x of perBlock)
                total += S.lookAheadSlides(scratch, x.g, no, theme, assets).pages;
        }
        let from = 0;
        for (const x of perBlock) {
            const r = S.lookAheadSlides(pptx, x.g, no, theme, assets, list.length > 1
                ? { title: (0, blocks_1.blockTitle)('2 Week Look Ahead Plan', x.b.label), letterFrom: from, letterTotal: total }
                : {});
            from += r.pages;
            if (list.length > 1) {
                log(`  · 2 Week Look Ahead [${x.b.label}]: ${x.g.reduce((n, y) => n + y.rows.length, 0)} row(s) ` +
                    `in ${x.g.length} group(s) across ${r.pages} page(s)`);
            }
        }
        // The reports below describe the PRIMARY block, as they always have; the per-block figures are in
        // the run log above.
        const g = perBlock[0].g;
        const rows = perBlock.reduce((n, x) => n + x.g.reduce((m, y) => m + y.rows.length, 0), 0);
        tk.note('Generating the 2 Week Look Ahead', `${g.length} group(s), ${rows} row(s)`);
        tk.fact('Look Ahead rows', rows);
        // Role reconstruction and row creation, reported and asserted: one row per task+person, every
        // role filled. Which columns the export actually supplied is recorded, since that is what varies.
        const st = derive_1.lastLookAheadRoles;
        if (st) {
            const cols = {
                responsible: sch.activities.some((a) => a.responsible),
                accountable: sch.activities.some((a) => a.accountable),
                head: sch.activities.some((a) => a.head),
                team: false,
            };
            const rv = (0, roles_1.roleValidation)(sch.projectName, st, (0, derive_1.lookAheadRoleRows)(g), cols);
            rolesMd = rv.md;
            rolesOk = rv.ok;
            log(`  · roles: ${st.rows} row(s) from ${st.activities} activity(ies); ` +
                `${st.splitActivities} multi-person field(s) split, ${st.duplicatesDropped} duplicate(s) collapsed, ` +
                `${st.mirroredAccountable}/${st.mirroredHead} Accountable/Head filled from Responsible: ${rv.ok ? 'PASS' : 'FAIL'}`);
        }
        return g;
    }, []);
    // ---- Contents, written last from the sections that were actually generated ----
    // The list and its numbering are the section register, so the Contents page cannot disagree with the
    // deck: a section that was skipped is absent and the ones after it have already moved up.
    fillContents(sections.map((x) => x.title));
    log(`Contents: ${sections.map((x) => `${x.n}. ${x.title}`).join(' · ')}`);
    S.thankYou(pptx, theme, assets);
    tk.fact('Slides generated before the write step', pptx.slides?.length ?? 0);
    let buffer = await tk.run('Writing the PowerPoint file', false, async () => (await pptx.write({ outputType: 'nodebuffer' })));
    let slideCount = pptx.slides?.length ?? 0;
    tk.note('Writing the PowerPoint file', `${slideCount} slides, ${(buffer.length / 1048576).toFixed(1)} MB`);
    // Copy the static slides (Land Acquisition, Master Plan) verbatim from the reference deck,
    // replacing the placeholders. Master Plan is omitted if the reference has no such slide.
    let landCopied = false;
    let copiedTitles = [];
    if (inp.reference) {
        log('Merging static slides from reference deck…');
        try {
            const specs = [
                // Front page, Red Flag and Tracker come verbatim from the previous deck (exact TE format);
                // the cover date is advanced to the coming Friday.
                { title: 'Cover', position: 'first', transform: 'coverDate' },
                // NOTE: neither Red Flag Summary nor Tracker & Schedule is copied from the previous deck —
                // both are regenerated from the Summary Dashboard, which is the single source of truth for
                // their column structure, group headings, legend colours and the project's indicators.
                // Land is taken from the reference ONLY when no updated Land slide was uploaded.
                ...(inp.land ? [] : [{ title: 'Land Acquisition', placeholderTitle: 'Land Acquisition Status' }]),
                { title: 'Master Plan', placeholderTitle: 'Master Plan' },
                // The closing THANK YOU slide comes from the previous deck VERBATIM — same layout, text,
                // logo, fonts, positions, spacing, images, colours and dimensions. It is matched by POSITION
                // (last slide of each deck) because the two decks word their first run differently: the
                // generated slide leads with "THANK YOU", the previous deck with the contact line, so no one
                // heading string finds both. No transform is applied. The generated placeholder is dropped in
                // the swap, and if the previous deck has no closing slide the generated one simply stays.
                { title: 'Thank You', position: 'last' },
            ];
            const merged = await tk.run('Merging slides from the previous PPT', true, () => (0, postmerge_1.mergeStaticSlides)(buffer, inp.reference, specs, log), { buffer, copied: [] });
            buffer = merged.buffer;
            landCopied = merged.copied.some((c) => /land acquisition/i.test(c));
            copiedTitles = merged.copied;
            tk.note('Merging slides from the previous PPT', merged.copied.length ? `copied: ${merged.copied.join(', ')}` : 'nothing to copy');
        }
        catch (e) {
            log(`  · merge skipped (${e.message}) — placeholders retained`);
        }
    }
    else {
        log('No reference deck supplied — Land Acquisition / Master Plan left as placeholders.');
    }
    // An uploaded Land Acquisition *deck* replaces that slide (image uploads were handled above).
    if (landIsPptx) {
        log('Land Acquisition Status… (from uploaded slide deck)');
        try {
            const m2 = await (0, postmerge_1.mergeStaticSlides)(buffer, inp.land, [
                { title: 'Land Acquisition', placeholderTitle: 'Land Acquisition Status', refFallbackFirst: true },
            ], log, { skipExtras: true, skipTheme: true });
            buffer = m2.buffer;
            if (m2.copied.length) {
                landCopied = true;
                copiedTitles = copiedTitles.concat(m2.copied);
            }
        }
        catch (e) {
            log(`  · Land Acquisition upload could not be merged (${e.message}) — placeholder retained`);
        }
    }
    else if (landIsImage) {
        landCopied = true;
    }
    // ---- package hygiene, then the integrity gate ----
    // Shape ids are made unique BEFORE validation: a duplicate id is one of the defects the validator
    // reports, and this is the last point at which it can still be corrected.
    const hygiene = await tk.run('Tidying the package', true, () => (0, package_validate_1.normaliseShapeIds)(buffer), { buffer, renumbered: 0, parts: [], foldersDropped: 0, mediaDropped: [] });
    buffer = hygiene.buffer;
    const hygieneNote = [
        hygiene.renumbered ? `${hygiene.renumbered} duplicate shape id(s) renumbered across ${hygiene.parts.length} part(s)` : '',
        hygiene.foldersDropped ? `${hygiene.foldersDropped} bare directory entr(y/ies) dropped` : '',
        hygiene.mediaDropped.length ? `${hygiene.mediaDropped.length} orphaned media part(s) removed` : '',
    ].filter(Boolean).join('; ');
    if (hygieneNote) {
        log(`  · ${hygieneNote}`);
        tk.note('Tidying the package', hygieneNote);
    }
    // THE GATE. A .pptx that fails this is not a deck, it is a repair prompt — so generation stops here
    // rather than handing over a file PowerPoint will refuse to open cleanly.
    const pkg = await tk.run('Validating the PowerPoint package', false, () => (0, package_validate_1.validatePackage)(buffer));
    packageMd = (0, package_validate_1.packageReport)(pkg, sch.projectName);
    packageOk = pkg.ok;
    const pkgErrors = pkg.issues.filter((i) => i.severity === 'error');
    log(`  · package integrity: ${pkg.ok ? 'PASS' : 'FAIL'} — ${pkg.parts} parts, ${pkg.slides} slides, ` +
        `${pkgErrors.length} error(s), ${pkg.issues.length - pkgErrors.length} warning(s)`);
    if (!pkg.ok) {
        throw new diagnostics_1.StageError('Validating the PowerPoint package', (0, diagnostics_1.describeError)(new Error('the generated package is not a valid Open XML document, and PowerPoint would offer to repair it. ' +
            `${pkgErrors.length} integrity error(s): ` +
            pkgErrors.slice(0, 5).map((e) => `[${e.area}] ${e.detail}`).join(' · ') +
            (pkgErrors.length > 5 ? ` …and ${pkgErrors.length - 5} more` : '') +
            '. See validation_report.md.')));
    }
    // recount slides from the (possibly merged) package so the audit reflects any omission
    try {
        const JSZip = (await Promise.resolve().then(() => __importStar(require('jszip')))).default;
        const z = await JSZip.loadAsync(buffer);
        const pres = await z.file('ppt/presentation.xml')?.async('string');
        if (pres)
            slideCount = (pres.match(/<p:sldId\b/g) || []).length;
    }
    catch { /* keep pptxgenjs count */ }
    const cp = (t) => copiedTitles.some((c) => c.toLowerCase().includes(t));
    const audit = (0, audit_1.runAudit)(sch, {
        rf: rf || (cp('red flag') ? { cells: ['copied'] } : null),
        kf, tr: trackerData ? { cells: trackerData.indicators.map((i) => i.status) } : null, ap,
        overviewRows: ovRows, lookAheadGroups: laGroups.length, landCopied,
    }, slideCount);
    // ---- font availability (explicit warning, never a silent substitution) ----
    const fontRep = (0, typography_1.checkFonts)((0, paths_1.resolveFontsDir)(), fs_1.default);
    for (const w of fontRep.warnings)
        log('  ! ' + w);
    if (fontRep.ok)
        log(`Fonts OK — ${typography_1.REQUIRED_FONTS.map((f) => f.family).join(', ')}`);
    // ---- automatic layout validation of the finished deck ----
    log('Validating layout (overflow / overlap / off-slide / readability)…');
    const validation = await (0, layout_validate_1.validateLayout)(buffer);
    const layoutMd = (0, layout_validate_1.validationReport)(validation, sch.projectName, fontRep.warnings, corrections);
    log(`Layout validation: ${validation.ok ? 'PASS' : 'FAIL'} — ` +
        `${validation.issues.filter((i) => i.severity === 'fail').length} failure(s), ` +
        `${validation.issues.filter((i) => i.severity === 'warn').length} warning(s)`);
    for (const it of validation.issues.filter((i) => i.severity === 'fail').slice(0, 10)) {
        log(`  ✗ slide ${it.slide}: ${it.kind} — ${it.detail}`);
    }
    // ---- presentation-refinement checks, read back out of the finished package ----
    log('Validating presentation refinement (colouring / typography consistency)…');
    const refine = await (0, refine_validate_1.validateRefinement)(buffer);
    const layoutSummary = `${validation.slides} slides checked, ` +
        `${validation.issues.filter((i) => i.severity === 'fail').length} failure(s)`;
    const refineMd = (0, refine_validate_1.refinementReport)(refine, sch.projectName, validation.ok, layoutSummary);
    log(`Presentation refinement: ${refine.ok ? 'PASS' : 'FAIL'}`);
    for (const c of refine.checks.filter((c) => !c.ok))
        log(`  ✗ ${c.name} — ${c.detail}`);
    // ---- presentation consistency, on the ASSEMBLED deck (order, headers, cascade) ----
    log('Validating presentation consistency (section order / headers / approval cascade)…');
    const apForCheck = (apRes?.details || []).map((d) => ({
        approval: d.row.approval, application: d.row.application, drawings: d.row.drawings,
        fee: d.row.fee, issued: d.row.issued, cascaded: d.cascadeApplied || [],
    }));
    const cons = await (0, consistency_validate_1.validateConsistency)(buffer, apForCheck);
    const consistencyMd = (0, consistency_validate_1.consistencyReport)(cons, sch.projectName, validation.ok, layoutSummary);
    const consistencyOk = cons.ok && validation.ok;
    log(`Presentation consistency: ${consistencyOk ? 'PASS' : 'FAIL'}`);
    for (const c of cons.checks.filter((c) => !c.ok))
        log(`  ✗ ${c.name} — ${c.detail}`);
    const kfCheck = await (0, key_focus_validate_1.validateKeyFocus)(buffer);
    const keyFocusMd = (0, key_focus_validate_1.keyFocusReport)(kfCheck, sch.projectName);
    const keyFocusOk = kfCheck.ok;
    log(`Key Focus layout: ${keyFocusOk ? 'PASS' : 'FAIL'}`);
    for (const c of kfCheck.checks.filter((c) => !c.ok))
        log(`  ✗ ${c.name} — ${c.detail}`);
    const overviewMd = (0, overview_select_1.selectionReport)(sch, (0, overview_select_1.selectOverview)(sch));
    // ---- what the previous deck actually contains (never a silent assumption) ----
    if (inp.reference) {
        await tk.run('Checking the previous PPT for the expected sections', true, async () => {
            const titles = await (0, postmerge_1.refSlideTitles)(inp.reference);
            const want = ['Contents', 'Red Flag', 'Key Focus', 'Tracker', 'Project Overview',
                'Land Acquisition', 'Approval', '2 Week Look Ahead', 'Thank You'];
            const norm = (v) => v.toLowerCase().replace(/[^a-z0-9]/g, '');
            for (const w of want) {
                const hit = titles.find((t) => norm(t).includes(norm(w)));
                refSlides.push({ title: w, found: !!hit, matchedAs: hit });
            }
            // The cover reads "<code> | <project> | <subtitle>" — compare the project part only.
            const cover = titles[0] || '';
            const coverName = (cover.split('|')[1] || cover).trim();
            if (cover)
                names.push({ source: 'Previous PPT (cover)', name: coverName, normalised: (0, match_1.cleanName)(coverName) });
            tk.note('Checking the previous PPT for the expected sections', `${titles.length} slide(s); ${refSlides.filter((x) => x.found).length}/${want.length} expected sections present`);
            tk.fact('Previous PPT slides', titles.length);
            return true;
        }, false);
    }
    const errorMd = (0, diagnostics_1.errorReport)({
        project: sch.projectName, code, tracker: tk, names, slides: refSlides, match: match_1.matchScore,
    });
    const stagesOk = tk.failed.length === 0;
    if (!stagesOk)
        log(`Completed with ${tk.failed.length} skipped section(s) — see error_report.md`);
    const project = {
        code, name: sch.projectName, subtitle: inp.subtitle || '', dataDate: sch.dataDate, activities: sch.activities.length,
    };
    log(`Done — ${slideCount} slides, audit ${audit.overall}/100`);
    return { buffer, audit, project, validation, layoutMd, overviewMd, trackerMd, trackerOk, trackerScheduleMd, trackerScheduleOk, scheduleLayoutMd, scheduleLayoutOk, scheduleReviewMd, scheduleDesignMd, scheduleDesignOk, approvalScheduleMd, approvalScheduleOk, packageMd, packageOk, rolesMd, rolesOk, redFlagMd, redFlagOk, consistencyMd, consistencyOk, keyFocusMd, keyFocusOk, errorMd, stagesOk, approvalMd, approvalOk, refineMd, refineOk: refine.ok, fontWarnings: fontRep.warnings };
}
