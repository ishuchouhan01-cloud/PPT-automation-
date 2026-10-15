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
exports.generateDesignBuild = generateDesignBuild;
exports.dbFailureReport = dbFailureReport;
// engine.ts — the DESIGN & BUILD deck builder.
//
// THE SHAPE OF THIS MODE
// ----------------------
//     PROJECT TYPE → DESIGN & BUILD → the project's own reference PPT + Excel export(s) + dashboard
//                  → learn which sections that project actually has
//                  → map the Excel and dashboard data into them
//                  → generate
//
// It is NOT "run Launch & Design and change a few slides". There is no Approvals Status slide, no
// Launch & Design Red Flag structure, no Land Acquisition / Marketing / Sales phase logic, and no
// hard-coded contents list. The section list is decided per run, from the inputs.
//
// SECTION NUMBERING, AS THE REFERENCE DOES IT
// -------------------------------------------
// Both reference decks number a section that spans several slides with a LETTER suffix and list the
// section ONCE in Contents: "TRACKER & SCHEDULE 3A", "SCHEDULE - CONSTRUCTION 3B", but Contents says
// only "Tracker & Schedule". A single-slide section gets a bare number ("TRACKER & SCHEDULE 3" in
// Meena's deck, which has no construction schedule). That rule is implemented literally here.
//
// WHICH SECTIONS EXIST
// --------------------
//   Red Flag Summary          the dashboard's "Red Flag Summary DB" row for this project
//   Key Focus Area            the dashboard's "Key Focus Area" rows for this project
//   Tracker & Schedule   3A   the "Tracker DB" status matrix + the "Schedule - Design" timeline
//                        3B   the "Schedule - Build" timeline, when that project has one
//   Project Overview     4A   the DESIGN export's tree
//                        4B   the BUILD export's tree, when that export was supplied
//   2 Week Look Ahead    5A…  from whichever export(s) were supplied
//   Site Progress             carried verbatim from the reference deck — photographs are never faked
//
// Any of these is omitted, from the deck AND from Contents, when its input is missing. An omitted
// section is always explained in design_build_validation.md rather than silently dropped.
const pptxgenjs_1 = __importDefault(require("pptxgenjs"));
const asta_excel_1 = require("../parsers/asta_excel");
const S = __importStar(require("../ppt/slides"));
const theme_1 = require("../ppt/theme");
const postmerge_1 = require("../ppt/postmerge");
const package_validate_1 = require("../ppt/package_validate");
const diagnostics_1 = require("../ppt/diagnostics");
const typography_1 = require("../ppt/typography");
const paths_1 = require("../paths");
const dashboard_1 = require("./dashboard");
const project_context_1 = require("./project_context");
const branch_1 = require("./branch");
const matrix_1 = require("./matrix");
// THE LAUNCH & DESIGN RENDERERS, used as-is.
//   projectOverview()  — ppt/overview.ts : the seven-column table, 73:27 Gantt, Today line, legend,
//                        three-level palette, measured pagination and continuation pages.
//   lookAheadSlides()  — ppt/slides.ts   : the look-ahead table, person band joined with " | ",
//                        block pagination, fixed column widths, 19pt body / 21pt headings.
//   lookAhead()        — ppt/derive.ts   : the 14-day window, 100%-complete exclusion, person-wise
//                        grouping, progressed-first ordering, dedup, Unassigned last.
// Design & Build supplies the DATA and the WBS visibility; none of the drawing is duplicated here.
const overview_1 = require("../ppt/overview");
const derive_1 = require("../ppt/derive");
const overview_views_1 = require("./overview_views");
const keyfocus_1 = require("./keyfocus");
const theme_2 = require("./theme");
const validate_1 = require("./validate");
/**
 * Headings of the reference deck's VISIBLE slides.
 *
 * Visibility matters for Site Progress. Meena's previous deck HIDES its Site Progress page, and a
 * hidden slide is one its author chose not to present — registering the section from it would put a
 * line in Contents pointing at a page that never appears in the show. So the section is registered
 * from visible slides only, and the merge is told to skip hidden ones, keeping the two consistent.
 */
async function visibleRefHeadings(ref) {
    const JSZip = require('jszip');
    const z = await JSZip.loadAsync(ref);
    const rels = (await z.file('ppt/_rels/presentation.xml.rels')?.async('string')) || '';
    const pres = (await z.file('ppt/presentation.xml')?.async('string')) || '';
    const map = {};
    for (const m of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g))
        map[m[1]] = m[2];
    const out = [];
    for (const m of pres.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)) {
        const tgt = map[m[1]];
        if (!tgt || !/slides\/slide/.test(tgt))
            continue;
        const part = 'ppt/' + tgt.replace(/^\.\.\//, '').replace(/^ppt\//, '');
        const xml = (await z.file(part)?.async('string')) || '';
        if (/<p:sld\b[^>]*\bshow="0"/.test(xml))
            continue; // hidden — not part of the show
        const runs = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((x) => x[1]).filter((s) => s.trim());
        out.push(runs.slice(0, 3).join(' ').replace(/&amp;/g, '&').trim());
    }
    return out;
}
/**
 * A section register.
 *
 * Numbers are handed out in the order sections are actually created, so an omitted section does not
 * leave a gap and everything after it moves up. Letter suffixes are decided at the END, once the
 * number of slides in each section is known — which is the only point at which "3" versus "3A/3B"
 * can be answered correctly.
 */
class Register {
    constructor() {
        this.list = [];
    }
    open(title) {
        const s = { n: this.list.length + 1, title, labels: [] };
        this.list.push(s);
        return s;
    }
    /** Label for the next slide of a section. Provisional until `finalise()`. */
    label(s, idx, total) {
        return total > 1 ? `${s.n}${String.fromCharCode(65 + idx)}` : String(s.n);
    }
    sections() { return this.list; }
    contents() { return this.list.map((s) => s.title); }
}
const iso = (d) => d.toISOString().slice(0, 10);
/**
 * A required-stage failure, in the shape the API already knows how to report.
 *
 * StageError carries a StructuredError rather than loose strings, so the browser gets the stage, the
 * reason and a plain-language next step instead of "Generation failed".
 */
function stageFail(stage, message, hint) {
    return new diagnostics_1.StageError(stage, {
        message, type: 'DesignBuildInputError',
        file: 'db/engine.ts', fn: 'generateDesignBuild', line: null, column: null,
        stack: '', hint,
    });
}
/**
 * Build a Design & Build deck.
 *
 * Nothing here throws for missing OPTIONAL input; a missing input removes its section and is
 * reported. It throws only when there is no usable input at all, because a deck with a cover and
 * nothing else is not a weekly review.
 */
async function generateDesignBuild(inp) {
    const log = inp.log || (() => { });
    const tk = new diagnostics_1.StageTracker(log);
    const notes = [];
    const omitted = [];
    // ---- inputs ----
    // ONE Excel export view. `designExcel` / `buildExcel` are still accepted so an older client keeps
    // working, but they are the same input — never two schedules to reconcile.
    const excelBuf = inp.excel || inp.designExcel || inp.buildExcel;
    if (!excelBuf) {
        throw stageFail('Receiving the uploaded files', 'Design & Build needs the Excel export view.', 'Upload the Excel export view containing this project\'s Design and/or Construction schedule ' +
            'data. It is the only schedule file required — the Design and Construction branches are read ' +
            'from inside it.');
    }
    const sch = await tk.run('Loading the Excel export view', true, () => (0, asta_excel_1.parseScheduleFromExcel)(excelBuf), null);
    if (!sch) {
        throw stageFail('Loading the Excel export view', 'The Excel export view could not be read.', 'Check it is the full schedule export (the sheet with the activity rows), not a summary tab.');
    }
    // WHICH WORKSHEET WAS READ. Stated every run, because an export can carry more than one and the
    // wrong choice is invisible in the deck: NH Homestay's workbook holds a stale HIDDEN "Sheet1"
    // beside the live "Sheet2", with the same headers and almost the same row count, so every date
    // came out plausible and wrong. Anyone checking a date against the spreadsheet can now see which
    // sheet the deck was built from.
    log(`  · Excel export view: ${asta_excel_1.lastSheetChoice}`);
    const primary = sch;
    const reportDate = inp.reportDate || iso(new Date());
    // -------------------------------------------------------------------------------------------
    // THE PROJECT CONTEXT — resolved ONCE, before any section is built.
    //
    // This is the fix for cross-project contamination. The Summary Dashboard is a portfolio workbook,
    // so every section has to filter it, and the filter used to fall back to name matching whenever the
    // operator left the project-code box empty. The name matcher scores any shared word and four live
    // projects share "Residence", so generating Meena & KrishnaSai with no code produced a Key Focus
    // slide carrying Uma & Kush's, Indiranagar's and Ooty's rows as well — 11 rows instead of 6.
    //
    // The code is now resolved from the DATA when the operator does not supply one: the export's own
    // project row carries it ("0116 - Meena and KrishnaSai Residence"). Every source is then filtered by
    // code alone, and no generator is allowed to query the workbook on its own terms.
    // -------------------------------------------------------------------------------------------
    const dashboardRows = inp.dashboard
        ? await (0, dashboard_1.listDashboardProjects)(inp.dashboard).catch(() => [])
        : [];
    const ctx = (0, project_context_1.resolveProjectContext)({
        operatorCode: inp.code, design: sch, build: null, dashboardRows, reportDate,
    });
    ctx.notes.forEach((n) => notes.push(`Project scope — ${n}`));
    log(`Project context: ${ctx.name}${ctx.code ? ` · code ${ctx.code}` : ' · NO CODE'} ` +
        `(from ${ctx.codeSource}); ${dashboardRows.length} project(s) in the dashboard, filtering to this one.`);
    const projectName = ctx.name;
    const code = ctx.code;
    /** Every section's guard result, checked together before the deck is written. */
    const guards = [];
    // THE ONE EXPORT is inspected for the branches it holds. That is what decides the section list —
    // not which files were uploaded, because there is only ever one.
    const shape = (0, branch_1.readExportShape)(sch, code);
    if (shape)
        shape.notes.forEach((n) => notes.push(`Excel export view — ${n}`));
    const designBranch = shape?.branches.find((b) => b.kind === 'design') || null;
    const buildBranch = shape?.branches.find((b) => b.kind === 'construction') || null;
    log(`Branches in the Excel export view: ` +
        (shape?.branches.length
            ? shape.branches.map((b) => `"${b.name}" (${b.kind}, ${b.children.length} child row(s))`).join('; ')
            : 'none recognised — the Project Overview falls back to the whole project tree'));
    if (!buildBranch) {
        notes.push('The Excel export view contains no "Construction" branch, so the construction-specific ' +
            'Project Overview is not generated. No second file is requested for it: construction data is ' +
            'optional at the data level.');
    }
    // Kept for the validation report, which distinguishes what the export contained.
    const designShape = shape;
    const buildShape = buildBranch ? shape : null;
    // ---- dashboard ----
    let dash = null;
    if (inp.dashboard) {
        dash = await tk.run('Reading the Summary Dashboard', true, () => (0, dashboard_1.readDbDashboard)(inp.dashboard, ctx), null);
        if (dash)
            dash.notes.forEach((n) => notes.push(`Dashboard — ${n}`));
    }
    else {
        notes.push('No Summary Dashboard was uploaded. Red Flag Summary, Key Focus Area and the Tracker ' +
            'status matrix and timelines all come from that workbook, so those sections are omitted.');
    }
    // ---- theme, fonts ----
    const theme = (0, theme_2.dbTheme)(await (0, theme_1.extractTheme)(inp.reference));
    const assets = inp.assets || {};
    const fontRep = (0, typography_1.checkFonts)((0, paths_1.resolveFontsDir)(), require('fs'));
    const pptx = new pptxgenjs_1.default();
    pptx.defineLayout({ name: 'TE', width: 26.67, height: 15 });
    pptx.layout = 'TE';
    const logos = (s) => {
        if (assets.wordmark)
            s.addImage({ data: assets.wordmark, x: 0.57, y: 0.39, w: 3.47, h: 0.5 });
        if (assets.mascot)
            s.addImage({ data: assets.mascot, x: 0.61, y: 13.74, w: 0.74, h: 0.79 });
    };
    // ---- cover + an empty Contents to be filled once the sections are known ----
    S.cover(pptx, primary, code, inp.subtitle || '', reportDate, theme, assets);
    const fillContents = S.contents(pptx, theme, assets);
    const reg = new Register();
    const slideLabels = [];
    // -------------------------------------------------------------------------------------------
    // 1. RED FLAG SUMMARY — the dashboard's band matrix, one project row
    // -------------------------------------------------------------------------------------------
    let redFlagDrawn = false;
    if (dash?.redFlag) {
        const sec = reg.open('Red Flag Summary');
        const s = pptx.addSlide();
        const label = reg.label(sec, 0, 1);
        S.sectionHeader(s, 'Red Flag Summary', label, theme);
        logos(s);
        guards.push((0, project_context_1.guardSection)(ctx, 'Red Flag Summary', [{ code: dash.redFlag.code, name: dash.redFlag.project, what: `row "${dash.redFlag.project}"` }]));
        (0, matrix_1.dbMatrix)(s, dash.redFlag, 'color', theme, { showLegend: true });
        sec.labels.push(label);
        slideLabels.push(label);
        redFlagDrawn = true;
        log(`  · Red Flag Summary: ${dash.redFlag.cols.length} columns in ${dash.redFlag.bands.length} bands`);
    }
    else {
        omitted.push({ section: 'Red Flag Summary', reason: inp.dashboard
                ? 'the dashboard has no readable "Red Flag Summary DB" row for this project'
                : 'no Summary Dashboard was uploaded' });
    }
    // -------------------------------------------------------------------------------------------
    // 2. KEY FOCUS AREA
    // -------------------------------------------------------------------------------------------
    let keyFocusRes = null;
    if (dash && dash.keyFocus.length) {
        const sec = reg.open('Key Focus Area');
        const s = pptx.addSlide();
        const label = reg.label(sec, 0, 1);
        S.sectionHeader(s, 'Key Focus Area', label, theme);
        logos(s);
        guards.push((0, project_context_1.guardSection)(ctx, 'Key Focus Area', dash.keyFocus.map((r, i) => ({
            code: r.code, name: r.project, what: `row ${i + 1} "${r.focus.slice(0, 40)}…" (${r.project})`,
        }))));
        keyFocusRes = (0, keyfocus_1.dbKeyFocus)(s, dash.keyFocus, dash.keyFocusBands, theme);
        sec.labels.push(label);
        slideLabels.push(label);
        log(`  · Key Focus Area: ${keyFocusRes.rendered} row(s) in ${keyFocusRes.bands.length} band(s)`);
    }
    else {
        omitted.push({ section: 'Key Focus Area', reason: inp.dashboard
                ? 'the dashboard lists no Key Focus rows for this project'
                : 'no Summary Dashboard was uploaded' });
    }
    // -------------------------------------------------------------------------------------------
    // 3. TRACKER & SCHEDULE — 3A the status matrix + design timeline, 3B the construction timeline
    //
    // These are ONE section with up to two slides, exactly as the reference numbers them, which is why
    // "Schedule - Construction" does not get its own Contents line. That was checked against both
    // decks: NH Homestay's Contents lists six sections for eleven content slides.
    // -------------------------------------------------------------------------------------------
    if (dash?.tracker) {
        guards.push((0, project_context_1.guardSection)(ctx, 'Tracker & Schedule (status matrix)', [{ code: dash.tracker.code, name: dash.tracker.project, what: `row "${dash.tracker.project}"` }]));
    }
    if (dash?.design) {
        guards.push((0, project_context_1.guardSection)(ctx, 'Tracker & Schedule (design timeline)', [{ code: dash.design.code, name: dash.design.project, what: `row "${dash.design.project}"` }]));
    }
    if (dash?.build && dash.build.from >= 0) {
        guards.push((0, project_context_1.guardSection)(ctx, 'Schedule - Construction', [{ code: dash.build.code, name: dash.build.project, what: `row "${dash.build.project}"` }]));
    }
    const trackerSlides = [];
    if (dash?.tracker || (dash?.design && dash.design.from >= 0)) {
        trackerSlides.push({
            title: 'Tracker & Schedule',
            draw: (s, label) => {
                S.sectionHeader(s, 'Tracker & Schedule', label, theme);
                logos(s);
                let y = matrix_1.DB_TOP;
                if (dash.tracker) {
                    const m = (0, matrix_1.dbMatrix)(s, dash.tracker, 'symbol', theme, { top: y, showLegend: false });
                    y = m.bottom + 0.45;
                }
                if (dash.design && dash.design.from >= 0) {
                    (0, matrix_1.dbTimeline)(s, dash.design, theme, { top: y, showLegend: false });
                    // BOTH legends sit in the bottom-right corner: the phase legend's right edge on the title
                    // bar's right edge, the status legend immediately to its left, and both standing on the
                    // same floor as the squirrel. Neither is placed by a hand-tuned offset — each measures
                    // itself and the phase legend's own width is what sets where the other one ends.
                    (0, matrix_1.dbPhaseLegend)(s, dash.design, theme);
                    if (dash.tracker) {
                        (0, matrix_1.dbMatrixLegend)(s, dash.tracker, 'symbol', theme, { right: matrix_1.DB_RIGHT - (0, matrix_1.dbPhaseLegendSize)(dash.design).w - matrix_1.DB_LEGEND_GAP });
                    }
                }
                else if (dash.tracker) {
                    (0, matrix_1.dbMatrixLegend)(s, dash.tracker, 'symbol', theme);
                }
            },
        });
    }
    if (dash?.build && dash.build.from >= 0) {
        trackerSlides.push({
            title: 'Schedule - Construction',
            draw: (s, label) => {
                S.sectionHeader(s, 'Schedule - Construction', label, theme);
                logos(s);
                (0, matrix_1.dbTimeline)(s, dash.build, theme, { top: matrix_1.DB_TOP, showLegend: false });
                (0, matrix_1.dbPhaseLegend)(s, dash.build, theme);
            },
        });
    }
    if (trackerSlides.length) {
        const sec = reg.open('Tracker & Schedule');
        trackerSlides.forEach((t, i) => {
            const s = pptx.addSlide();
            const label = reg.label(sec, i, trackerSlides.length);
            t.draw(s, label);
            sec.labels.push(label);
            slideLabels.push(label);
            log(`  · ${t.title} → slide ${label}`);
        });
    }
    else {
        omitted.push({ section: 'Tracker & Schedule', reason: inp.dashboard
                ? 'the dashboard has neither a Tracker DB row nor a painted timeline for this project'
                : 'no Summary Dashboard was uploaded' });
    }
    if (!(dash?.build && dash.build.from >= 0)) {
        omitted.push({ section: 'Schedule - Construction (slide 3B)', reason: inp.dashboard
                ? 'the dashboard\'s "Schedule - Build" sheet has no painted months for this project, so it has ' +
                    'no construction bar to draw — this matches the reference deck for such a project, which has ' +
                    'no 3B slide either'
                : 'no Summary Dashboard was uploaded' });
    }
    // -------------------------------------------------------------------------------------------
    // 4. PROJECT OVERVIEW — one slide per supplied export
    // -------------------------------------------------------------------------------------------
    const MAX_OVERVIEW_ROWS = 16;
    // TRIAL: a requested view replaces the row SELECTION only. Everything downstream — the table, the
    // Gantt, the geometry — is the same code either way.
    const view = inp.overviewView && inp.overviewView !== 'default' ? inp.overviewView : 'default';
    const viewStats = [];
    const overviewPlan = [];
    /**
     * Which Project Overview slides this export supports.
     *
     * BOTH branches present → two slides cut out of the ONE tree: "Project Overview" from the Design
     * branch and "Project Overview - Construction" from the Construction branch.
     *
     * ONE branch (or none recognised) → a single slide over the WHOLE project tree, titled from that
     * branch. This is not laziness: Meena's export carries "Master Archives [LOD 500]" as a SIBLING of
     * "Design", and its reference slide shows that row — branch-scoping a single-branch export would
     * silently drop it.
     */
    const overviewTargets = [];
    if (shape) {
        if (designBranch && buildBranch) {
            overviewTargets.push({ title: 'Project Overview', branch: designBranch, label: `"${designBranch.name}" branch` });
            overviewTargets.push({ title: 'Project Overview - Construction', branch: buildBranch, label: `"${buildBranch.name}" branch` });
        }
        else {
            const only = designBranch || buildBranch;
            overviewTargets.push({
                title: only?.kind === 'construction' ? 'Project Overview - Construction' : 'Project Overview',
                branch: null,
                label: only ? `whole project tree (single "${only.name}" branch)` : 'whole project tree',
            });
        }
    }
    for (const tgt of overviewTargets) {
        let rows, note, depth;
        if (view === 'default') {
            const r = (0, branch_1.overviewRowsFor)(sch, shape, MAX_OVERVIEW_ROWS, tgt.branch || undefined);
            rows = r.rows;
            note = r.note;
            depth = r.depthUsed;
        }
        else {
            const v = (0, overview_views_1.resolveOverviewView)(sch, shape, view, inp.overviewCustomIds || [], tgt.branch || undefined);
            viewStats.push({ src: tgt.label, view, v });
            rows = v.rows;
            note = v.note;
            depth = v.maxDepth;
        }
        if (!rows.length) {
            omitted.push({ section: tgt.title, reason: view === 'default'
                    ? `the ${tgt.label} has no rows beneath the project row`
                    : `the "${overview_views_1.VIEW_LABEL[view]}" view selected no rows from the ${tgt.label}` });
            continue;
        }
        // Every row comes from the ONE export whose project row was matched by code, so a foreign row can
        // only appear if that file itself mixes projects. Checked rather than assumed.
        const rootName = sch.wbs[shape.rootId]?.name || '';
        guards.push((0, project_context_1.guardSection)(ctx, `${tgt.title} (Excel export view)`, [
            { code: shape.code, name: rootName, what: `export project row "${rootName}"` },
        ]));
        overviewPlan.push({ title: tgt.title, rows, note, depth, src: tgt.label });
    }
    /**
     * D&B row → the Launch & Design renderer's row type.
     *
     * `major` marks the rows the renderer paints as hierarchy levels 1 and 2: the project row and the
     * branch rows. Everything deeper is a level-3 activity. `duration` is the ELAPSED day count, which
     * is what the renderer prints as "###ed" and divides by 30.4375 for the "##.##em" column — so the
     * two duration columns come out identical to before, from the same figures already in the export.
     */
    const toLdRow = (r) => ({
        name: r.name,
        indent: r.level,
        // The row's true WBS depth, 1-based, so the renderer can colour it by its ACTUAL level and cycle
        // the three colours for anything deeper. DbOverviewRow.level is 0-based (0 = the project row).
        level: r.level + 1,
        major: r.level <= 1,
        summary: r.level <= 1,
        pct: r.pct ?? 0,
        duration: r.elapsedDays ?? 0,
        baseStart: r.baselineStart,
        baseFinish: r.baselineFinish,
        start: r.start,
        finish: r.finish,
        critical: false,
        milestone: false,
        delayDays: 0,
        done: (r.pct ?? 0) >= 100 && !!r.finish,
        category: 'Design & Build WBS view',
    });
    const overviewResults = [];
    if (overviewPlan.length) {
        const sec = reg.open('Project Overview');
        for (const p of overviewPlan) {
            // THE LAUNCH & DESIGN RENDERER. It owns the columns, the 73:27 split, the palette, the Gantt,
            // the Today line, the legend, the row heights, the pagination and the continuation pages. Design
            // & Build passes rows and a title and nothing else — there is no second implementation to drift.
            const r = (0, overview_1.projectOverview)(pptx, sch, theme, assets, reportDate, sec.n, {
                rows: p.rows.map(toLdRow),
                title: p.title,
                // Asta-style nesting: every summary row's colour runs down the left of its own children.
                nestingBars: true,
            });
            for (let i = 0; i < r.pages; i++) {
                const label = r.pages > 1 ? `${sec.n}${String.fromCharCode(65 + i)}` : String(sec.n);
                sec.labels.push(label);
                slideLabels.push(label);
            }
            overviewResults.push({ title: p.title, rows: r.rows, pages: r.pages, note: p.note, depth: p.depth, src: p.src });
            log(`  · ${p.title} → ${r.rows} row(s) across ${r.pages} page(s), depth ${p.depth}, from the ` +
                `${p.src}${view === 'default' ? '' : `, view "${overview_views_1.VIEW_LABEL[view]}"`} — rendered by the Launch & Design engine`);
        }
        // A section that produced several slides takes letter suffixes; the register only knows how many
        // once every entry has been drawn, so the labels are corrected here in one pass.
        if (sec.labels.length > 1) {
            sec.labels.length = 0;
            const total = overviewResults.reduce((n, x) => n + x.pages, 0);
            for (let i = 0; i < total; i++)
                sec.labels.push(`${sec.n}${String.fromCharCode(65 + i)}`);
        }
    }
    else {
        omitted.push({ section: 'Project Overview', reason: 'the Excel export view produced no overview rows' });
    }
    // -------------------------------------------------------------------------------------------
    // 5. 2 WEEK LOOK AHEAD PLAN — paginated, numbered 5A / 5B / …
    //
    // Both exports feed one section: the look-ahead is a fortnight of work, not a branch. Where both
    // are supplied their groups are concatenated with the design work first, matching the reference's
    // ordering of design packages ahead of site activities.
    // -------------------------------------------------------------------------------------------
    let lookPages = 0, lookRows = 0;
    const laParts = [];
    {
        // PROJECT FILTER FIRST, then the fortnight window. The export is one project's file, so the guard
        // is a check that it really is: a file holding two projects would otherwise put the other one's
        // tasks into this fortnight.
        const rootName = shape ? (sch.wbs[shape.rootId]?.name || '') : sch.projectName;
        guards.push((0, project_context_1.guardSection)(ctx, '2 Week Look Ahead (Excel export view)', [
            { code: shape?.code || '', name: rootName, what: `export project row "${rootName}"` },
        ]));
        // THE LAUNCH & DESIGN LOGIC, called directly: the 14-day window on the actual/projected finish,
        // 100%-complete work excluded first, one row per task and person, person-wise grouping in the
        // schedule's own order with progressed work first and Unassigned last, and the WBS sub-band
        // subheadings. Nothing is reimplemented for Design & Build.
        const groups = (0, derive_1.lookAhead)(sch, reportDate);
        laParts.push({
            src: 'Excel export view', considered: sch.activities.length,
            selected: groups.reduce((n, g) => n + g.rows.length, 0), groups: groups.length,
        });
        if (groups.length) {
            const sec = reg.open('2 Week Look Ahead');
            const before = pptx.slides?.length ?? 0;
            // THE LAUNCH & DESIGN RENDERER: the table, the person band joined with " | ", the summary bands,
            // the status icons, the Notes column, the fixed column widths, the block pagination that never
            // orphans a heading, and the continuation pages.
            // The heading is "2 Week Look Ahead" on EVERY page of the section — the shorter wording, and no
            // "(cont.)" on the continuation pages. Launch & Design keeps "2 Week Look Ahead Plan" and its
            // "(cont.)" suffix; both are opt-in arguments, so nothing there moves.
            S.lookAheadSlides(pptx, groups, sec.n, theme, assets, { title: '2 Week Look Ahead', sameHeadingEveryPage: true });
            lookPages = (pptx.slides?.length ?? 0) - before;
            lookRows = groups.reduce((n, g) => n + g.rows.length, 0);
            // Launch & Design keeps ONE section number on every look-ahead page and marks the later ones
            // "(cont.)" in the title, rather than lettering them. Design & Build now does the same.
            for (let i = 0; i < lookPages; i++) {
                sec.labels.push(String(sec.n));
                slideLabels.push(String(sec.n));
            }
            log(`  · 2 Week Look Ahead: ${lookRows} row(s) across ${lookPages} page(s) — rendered by the ` +
                `Launch & Design engine`);
        }
        else {
            omitted.push({ section: '2 Week Look Ahead',
                reason: `no activity in the Excel export view has an actual/projected finish within 14 days of ` +
                    `${reportDate} while still being under 100% complete` });
        }
    }
    // -------------------------------------------------------------------------------------------
    // 6. SITE PROGRESS — registered so Contents lists it, but the SLIDES come from the reference deck.
    //
    // Site photographs cannot be regenerated and must never be invented. The section is registered
    // only when the reference deck actually has such a slide; the pages themselves are copied verbatim
    // by the merge step below.
    // -------------------------------------------------------------------------------------------
    let siteProgressTitles = [];
    if (inp.reference) {
        const titles = await visibleRefHeadings(inp.reference).catch(() => []);
        siteProgressTitles = titles.filter((t) => /site\s*progress/i.test(t.slice(0, 80)) && !/^\s*contents/i.test(t));
        if (siteProgressTitles.length) {
            const sec = reg.open('Site Progress');
            log(`  · Site Progress: ${siteProgressTitles.length} page(s) carried over from the reference deck`);
            // Labels are added after the merge, when the copied slides exist.
            sec.labels.push(...siteProgressTitles.map((_, i) => reg.label(sec, i, siteProgressTitles.length)));
        }
        else {
            omitted.push({ section: 'Site Progress', reason: 'the reference deck has no Site Progress slide' });
        }
    }
    else {
        omitted.push({ section: 'Site Progress',
            reason: 'no reference deck was uploaded, and site photographs are never generated' });
    }
    // -------------------------------------------------------------------------------------------
    // Contents + closing slide
    // -------------------------------------------------------------------------------------------
    const contents = reg.contents();
    fillContents(contents);
    S.thankYou(pptx, theme, assets);
    log(`Contents (${contents.length}): ${contents.join(' · ')}`);
    // -------------------------------------------------------------------------------------------
    // CROSS-PROJECT CONTAMINATION GATE
    //
    // Every section's records were checked against the project context as they were built. Nothing is
    // written until all of those checks pass. Deliberately fatal: a weekly review that quietly
    // attributes another project's commitments to this one is worse than no deck, because somebody acts
    // on it.
    // -------------------------------------------------------------------------------------------
    const contamination = (0, project_context_1.assertNoContamination)(guards);
    if (!contamination.ok) {
        throw stageFail('Checking the deck belongs to one project', 'Records belonging to another project reached this report, so it was not generated.', contamination.report + ` — the project context is ${ctx.name}` +
            (ctx.code ? ` (code ${ctx.code}, resolved from ${ctx.codeSource})` : ' with NO resolved code; ' +
                'enter the project code on the form to make scoping exact'));
    }
    log(`Project scoping: ${guards.length} section check(s) passed — every record belongs to ` +
        `${ctx.name}${ctx.code ? ` (${ctx.code})` : ''}.`);
    // ---- write ----
    let buffer = await tk.run('Writing the PowerPoint file', false, async () => (await pptx.write({ outputType: 'nodebuffer' })));
    // ---- carry the cover and the Site Progress pages over from the reference deck ----
    let copiedTitles = [];
    if (inp.reference) {
        try {
            const merged = await tk.run('Merging slides from the previous PPT', true, () => (0, postmerge_1.mergeStaticSlides)(buffer, inp.reference, [{ title: 'Cover', position: 'first', transform: 'coverDate' }], log, {
                // THE COVER DATE IS THE REPORTING DATE, not the date on the reference deck.
                //
                // The reference supplies the cover's design; it must not supply its date. Both reference
                // covers read "Weekly Update: 14th August 2026", and a regenerated deck was still showing
                // that because those covers glue the day to its label and the month and year into one run,
                // so only the ordinal was ever rewritten. `coverDateInline` handles the embedded form and
                // `coverDate` pins it to the date selected for THIS run.
                coverDate: new Date(reportDate + 'T00:00:00Z'),
                coverDateInline: true,
                // Design & Build generates its own "Schedule - Construction"; without this the reference's
                // copy would be carried over as an unrecognised extra and the deck would hold two.
                stdKeyExtra: (tl) => (tl.includes('schedule') && tl.includes('construction')) ? 'schedulecon' : null,
                // A slide the author hid is not part of the show, so it must not become a page here —
                // Contents is registered from visible slides only and the two have to agree.
                skipHiddenExtras: true,
            }), { buffer, copied: [] });
            buffer = merged.buffer;
            copiedTitles = merged.copied;
        }
        catch (e) {
            log(`  · merge skipped (${e.message}) — the generated cover is kept`);
            notes.push(`Merging from the reference deck failed (${e.message}); the generated cover was kept ` +
                `and no Site Progress pages were copied.`);
        }
    }
    // ---- integrity ----
    const norm = await (0, package_validate_1.normaliseShapeIds)(buffer).catch(() => null);
    if (norm && norm.buffer)
        buffer = norm.buffer;
    const packageResult = await (0, package_validate_1.validatePackage)(buffer);
    const packageMd = (0, package_validate_1.packageReport)(packageResult, projectName);
    if (!packageResult.ok) {
        // Refuse to hand back a deck that would open with the repair dialog. An earlier build shipped a
        // slide whose relationships pointed at embedded workbooks that were never copied; this gate is
        // what caught it, and it stays absolute.
        throw stageFail('Validating the PowerPoint package', 'The generated deck failed its integrity check, so it was not returned.', packageResult.issues.filter((i) => i.severity === 'error')
            .map((i) => `${i.area}: ${i.detail}`).join(' · ') || 'see validation_report.md');
    }
    // ---- report ----
    const vin = {
        project: projectName, code, reportDate,
        contents, slideLabels,
        sections: reg.sections().map((s) => ({ n: s.n, title: s.title, slides: s.labels.length, labels: s.labels })),
        omitted, notes,
        redFlag: dash?.redFlag || null,
        tracker: dash?.tracker || null,
        designTimeline: dash?.design || null,
        buildTimeline: dash?.build || null,
        keyFocus: keyFocusRes,
        keyFocusRows: dash?.keyFocus.length || 0,
        overviews: overviewResults,
        lookAhead: lookPages ? { pages: lookPages, rows: lookRows } : null,
        lookAheadParts: laParts,
        designShape, buildShape,
        copiedTitles, siteProgressTitles,
        packageResult,
        fontWarnings: fontRep.warnings,
        scope: {
            code: ctx.code, codeSource: ctx.codeSource, name: ctx.name, aliases: ctx.aliases,
            portfolioProjects: dashboardRows.length,
            guards: guards.map((g) => ({ section: g.section, checked: g.checked, foreign: g.foreign.length })),
        },
        overviewView: view,
        overviewStats: viewStats.map((x) => ({
            src: x.src, rows: x.v.rows.length, parents: x.v.parents, leaves: x.v.leaves,
            ancestorsAdded: x.v.ancestorsAdded, maxDepth: x.v.maxDepth, note: x.v.note,
        })),
        overviewPages: overviewResults.reduce((n, r) => n + r.pages, 0),
        overviewGanttRows: overviewResults.reduce((n, r) => n + r.rows, 0),
    };
    const { md: validationMd, ok: validationOk } = (0, validate_1.dbValidationReport)(vin);
    const audit = {
        overall: validationOk ? 100 : 80,
        scores: {}, missing: omitted.map((o) => o.section), editabilityIssues: [],
        sections: contents.map((t) => ({ name: t, present: true })),
        slides: slideLabels.length + 2 + (siteProgressTitles.length ? 0 : 0),
    };
    return {
        buffer,
        project: {
            code, name: projectName, subtitle: inp.subtitle || '',
            dataDate: primary.dataDate,
            activities: sch.activities.length,
        },
        audit, contents, slideLabels,
        packageResult, packageMd,
        validationMd, validationOk,
        fontWarnings: fontRep.warnings,
        errorMd: '',
        overview: {
            view,
            stats: viewStats.map((s) => ({
                src: s.src, rows: s.v.rows.length, parents: s.v.parents, leaves: s.v.leaves,
                ancestorsAdded: s.v.ancestorsAdded, maxDepth: s.v.maxDepth, note: s.v.note,
            })),
            pages: overviewResults.reduce((n, r) => n + r.pages, 0),
            // Counted from what the renderer reported it DREW. The table and the Gantt are built from one
            // row list inside projectOverview(), so this is the number that proves they cannot disagree.
            ganttRows: overviewResults.reduce((n, r) => n + r.rows, 0),
        },
    };
}
/** Failure report in the same shape the Launch & Design path returns, so the API stays uniform. */
function dbFailureReport(e, project = '', code = '') {
    const info = e instanceof diagnostics_1.StageError ? e.info : (0, diagnostics_1.describeError)(e);
    const stage = e?.stage || 'Generating the Design & Build deck';
    return {
        stage, info,
        md: (0, diagnostics_1.errorReport)({
            project, code, tracker: new diagnostics_1.StageTracker(), names: [], slides: [],
            fatal: info, fatalStage: stage,
        }),
    };
}
