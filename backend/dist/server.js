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
const express_1 = __importDefault(require("express"));
const cors_1 = __importDefault(require("cors"));
const multer_1 = __importDefault(require("multer"));
const path_1 = __importDefault(require("path"));
const fs_1 = __importDefault(require("fs"));
const engine_1 = require("./ppt/engine");
const engine_2 = require("./db/engine");
const branch_1 = require("./db/branch");
const overview_views_1 = require("./db/overview_views");
const audit_1 = require("./audit");
const asta_excel_1 = require("./parsers/asta_excel");
const storage = __importStar(require("./storage"));
const paths_1 = require("./paths");
const VERSION = '2.56.0-excel';
const START = Date.now();
// Never swallow exceptions — always surface them in the console (and keep the process alive
// so the operator can read the error rather than the window vanishing).
process.on('uncaughtException', (e) => { console.error('\n[FATAL] uncaughtException:\n', e?.stack || e); });
process.on('unhandledRejection', (e) => { console.error('\n[FATAL] unhandledRejection:\n', e?.stack || e); });
// ---- startup diagnostics (Part 2) ----
const cfgPath = (0, paths_1.resolveConfigPath)();
const cfgRes = (0, paths_1.loadConfig)();
const assetsRes = (0, paths_1.loadAssets)();
const projectsLoaded = Object.keys(cfgRes.config?.approvals || {}).length;
storage.init();
const app = (0, express_1.default)();
app.use((0, cors_1.default)());
app.use(express_1.default.json());
// Previous Weekly Review decks are big — One Tiny Flower's is 106 MB — and a file over the limit was
// rejected by multer BEFORE the route ran, so the browser only saw a generic 500. The limit is now
// generous and configurable, and an over-size upload gets a plain-language message naming the file.
const MAX_UPLOAD_MB = Number(process.env.EDOS_MAX_UPLOAD_MB || 400);
const upload = (0, multer_1.default)({ storage: multer_1.default.memoryStorage(), limits: { fileSize: MAX_UPLOAD_MB * 1024 * 1024 } });
const loadAssets = () => (0, paths_1.loadAssets)().assets;
const fields = upload.fields([
    { name: 'excel', maxCount: 1 }, { name: 'dashboard', maxCount: 1 },
    { name: 'reference', maxCount: 1 },
    { name: 'land', maxCount: 1 }, // OPTIONAL updated Land Acquisition slide (.pptx or image)
    // DESIGN & BUILD only. A Design & Build project is planned as two separate schedules and exported
    // separately, so the mode accepts both. Neither field exists in the Launch & Design workflow, and
    // adding them changes nothing for it — multer simply never receives them.
    { name: 'designExcel', maxCount: 1 },
    { name: 'buildExcel', maxCount: 1 },
    // A project reported in BLOCKS supplies one further export per additional block. Absent for every
    // ordinary project, so nothing changes for them.
    { name: 'blockExcels', maxCount: 8 },
]);
const pick = (f, k) => f?.[k]?.[0]?.buffer;
/** Every file uploaded under one field name, in the order the browser sent them. */
const pickAll = (f, k) => (f?.[k] || []).map((x) => x.buffer);
/** Every failure leaves the API in the same shape, so the browser can show the real reason. */
function fail(res, code, stage, reason, details = '', stack = '', extra = {}) {
    console.error(`[FAIL] ${stage}: ${reason}${details ? ' — ' + details : ''}`);
    if (stack)
        console.error(stack);
    return res.status(code).json({ success: false, stage, reason, details, stack, error: reason, ...extra });
}
/** Multer rejects an over-size or unexpected file before the route body runs. */
const uploadFields = (req, res, next) => fields(req, res, (err) => {
    if (!err)
        return next();
    if (err.code === 'LIMIT_FILE_SIZE') {
        return fail(res, 413, 'Receiving the uploaded files', `One of the uploaded files is larger than the ${MAX_UPLOAD_MB} MB limit.`, `Field "${err.field || 'unknown'}" exceeded the limit. Previous Weekly Review decks can be very ` +
            `large; raise the limit by setting EDOS_MAX_UPLOAD_MB before starting the app, or upload a ` +
            `lighter deck.`, String(err.stack || ''));
    }
    return fail(res, 400, 'Receiving the uploaded files', err.message || 'Upload rejected', String(err.code || ''), String(err.stack || ''));
});
app.get('/api/health', (_req, res) => res.json({
    ok: true,
    status: 'OK',
    service: 'EDOS Weekly Review Cloud — ASTA Excel Edition',
    version: VERSION,
    uptime: `${Math.round((Date.now() - START) / 1000)}s`,
    configLoaded: cfgRes.loaded,
    projectsLoaded,
    templatesLoaded: assetsRes.count,
    engine: 'node (pptxgenjs, editable)',
    input: 'ASTA Excel export (.xlsx/.xlsm)',
    modes: ['launch-design', 'design-build'],
}));
app.post('/api/detect', uploadFields, async (req, res) => {
    try {
        const excel = pick(req.files, 'excel');
        if (!excel)
            return res.status(400).json({ error: 'ASTA Excel export required' });
        const sch = await (0, asta_excel_1.parseScheduleFromExcel)(excel);
        if (sch.activities.length === 0) {
            return res.status(422).json({
                error: 'This Excel export contains 0 activities. Please upload the FULL ASTA schedule ' +
                    'spreadsheet export (with thousands of rows), not an empty/partial file.',
                activities: 0,
            });
        }
        res.json({ name: sch.projectName, dataDate: sch.dataDate, activities: sch.activities.length });
    }
    catch (e) {
        res.status(500).json({ error: String(e.message || e) });
    }
});
/**
 * PROJECT TYPE routing.
 *
 * `mode` arrives with the form. Anything other than the Design & Build value falls through to the
 * Launch & Design path below, unchanged — including a request that omits `mode` entirely, which is
 * what every existing client sends. So the Launch & Design workflow cannot be affected by this
 * branch: it is reached by exactly the same requests as before.
 */
const isDesignBuild = (body) => String(body?.mode || '').toLowerCase().replace(/[^a-z]/g, '') === 'designbuild';
app.post('/api/generate', uploadFields, async (req, res) => {
    if (isDesignBuild(req.body))
        return generateDesignBuildRoute(req, res);
    const logLines = [];
    try {
        const files = req.files;
        const excel = pick(files, 'excel');
        if (!excel)
            return fail(res, 400, 'Receiving the uploaded files', 'The ASTA Excel export is required.', 'Upload the schedule export (.xlsx / .xlsm) — it is the only mandatory file.');
        // Parsed here first so an unreadable workbook is reported against the right stage rather than
        // surfacing later as an anonymous failure.
        let preSch;
        try {
            preSch = await (0, asta_excel_1.parseScheduleFromExcel)(excel);
        }
        catch (e) {
            const f = (0, engine_1.failureReport)(e);
            return fail(res, 422, 'Loading ASTA Excel export', `${f.info.type}: ${f.info.message}`, `${f.info.file}:${f.info.line ?? '?'} in ${f.info.fn}. ${f.info.hint}`, f.info.stack, { errorReport: f.md });
        }
        if (preSch.activities.length === 0) {
            return fail(res, 422, 'Loading ASTA Excel export', 'This Excel export contains 0 activities, so the deck would be empty.', 'Upload the FULL ASTA schedule spreadsheet export (Sheet1 with the activity rows), not a summary tab.');
        }
        const dashboard = pick(files, 'dashboard'), reference = pick(files, 'reference'), land = pick(files, 'land');
        storage.saveUpload('excel', files.excel[0].originalname, excel);
        if (dashboard)
            storage.saveUpload('dashboard', files.dashboard[0].originalname, dashboard);
        if (reference)
            storage.saveUpload('reference', files.reference[0].originalname, reference);
        if (land)
            storage.saveUpload('land', files.land[0].originalname, land);
        const code = req.body.code || '', subtitle = req.body.subtitle || '';
        const r = await (0, engine_1.generate)({
            excel, dashboard, reference, land, code, subtitle,
            // One further export per additional reporting block of the same project (empty for all others).
            extraExcels: pickAll(req.files, 'blockExcels'),
            reportDate: req.body.reportDate || undefined, assets: loadAssets(),
            log: (m) => logLines.push(`[${new Date().toISOString()}] ${m}`),
        });
        const buffer = r.buffer;
        const sch = preSch;
        const auditText = (0, audit_1.auditMarkdown)(r.audit, sch);
        const project = { code, name: sch.projectName, subtitle, dataDate: sch.dataDate, activities: sch.activities.length };
        const saved = storage.saveGenerated(project.name, {
            pptx: buffer, audit: auditText, log: logLines.join('\n'),
            'layout_validation.md': r.layoutMd,
            'project_overview_selection.md': r.overviewMd,
            'tracker_validation.md': r.trackerMd,
            'tracker_schedule_validation.md': r.trackerScheduleMd,
            'approval_validation.md': r.approvalMd,
            'presentation_refinement_validation.md': r.refineMd,
            'schedule_layout_validation.md': r.scheduleLayoutMd,
            'schedule_layout_review.md': r.scheduleReviewMd,
            'schedule_design_review.md': r.scheduleDesignMd,
            'approval_schedule_validation.md': r.approvalScheduleMd,
            'validation_report.md': r.packageMd,
            'lookahead_roles_validation.md': r.rolesMd,
            'red_flag_validation.md': r.redFlagMd,
            'presentation_consistency_validation.md': r.consistencyMd,
            'error_report.md': r.errorMd,
            'key_focus_layout_validation.md': r.keyFocusMd,
        });
        const audit = {
            overall: r.audit.overall,
            engine: 'node (pptxgenjs)',
            slides: r.audit.slides || undefined,
            sections: r.audit.sections,
            missing: r.audit.missing, editabilityIssues: r.audit.editabilityIssues, scores: r.audit.scores,
        };
        res.json({
            jobId: saved.id, project, audit,
            downloadUrl: `/api/download/${encodeURIComponent(saved.id)}`,
            auditUrl: `/api/file/${encodeURIComponent(saved.id)}/audit_report.md`,
            layoutValidationUrl: `/api/file/${encodeURIComponent(saved.id)}/layout_validation.md`,
            overviewSelectionUrl: `/api/file/${encodeURIComponent(saved.id)}/project_overview_selection.md`,
            trackerValidationUrl: `/api/file/${encodeURIComponent(saved.id)}/tracker_validation.md`,
            layoutOk: r.validation.ok,
            trackerOk: r.trackerOk,
            trackerScheduleUrl: `/api/file/${encodeURIComponent(saved.id)}/tracker_schedule_validation.md`,
            trackerScheduleOk: r.trackerScheduleOk,
            approvalValidationUrl: `/api/file/${encodeURIComponent(saved.id)}/approval_validation.md`,
            approvalOk: r.approvalOk,
            refineOk: r.refineOk,
            scheduleLayoutOk: r.scheduleLayoutOk,
            redFlagOk: r.redFlagOk,
            consistencyOk: r.consistencyOk,
            stagesOk: r.stagesOk,
            keyFocusOk: r.keyFocusOk,
            errorReportUrl: `/api/file/${encodeURIComponent(saved.id)}/error_report.md`,
            fontWarnings: r.fontWarnings,
        });
    }
    catch (e) {
        // A required stage failed. Report exactly where, why and in which file/function/line, and keep the
        // full forensic report next to the run so it can be read later.
        const f = (0, engine_1.failureReport)(e);
        logLines.push('ERROR ' + f.info.stack);
        try {
            storage.saveGenerated(`FAILED - ${req.body?.name || 'run'}`, {
                pptx: Buffer.alloc(0), audit: '', log: logLines.join('\n'), 'error_report.md': f.md,
            });
        }
        catch { /* reporting must never mask the original failure */ }
        return fail(res, 500, f.stage, `${f.info.type}: ${f.info.message}`, `${f.info.file}:${f.info.line ?? '?'} in ${f.info.fn}. ${f.info.hint}`, f.info.stack, { log: logLines, errorReport: f.md });
    }
});
/**
 * The DESIGN & BUILD generation route.
 *
 * Deliberately its own function rather than flags threaded through the Launch & Design handler: the
 * inputs differ (two schedule exports, a required dashboard, no Land slide), the outputs differ (one
 * validation report rather than fourteen), and the two must not be able to break each other.
 */
async function generateDesignBuildRoute(req, res) {
    const logLines = [];
    const log = (m) => logLines.push(`[${new Date().toISOString()}] ${m}`);
    const files = req.files;
    try {
        // ONE Excel export view is the single source of truth for all schedule data. The older field
        // names are still accepted so a stale client keeps working, but they are the same one input —
        // there is no second schedule file to reconcile.
        const excel = pick(files, 'excel') || pick(files, 'designExcel') || pick(files, 'buildExcel');
        const dashboard = pick(files, 'dashboard');
        const reference = pick(files, 'reference');
        if (!excel) {
            return fail(res, 400, 'Receiving the uploaded files', 'Design & Build needs the Excel export view.', 'Upload the Excel export view containing this project\'s Design and/or Construction schedule ' +
                'data. It is the only schedule file required — the Design and Construction branches are read ' +
                'from inside it.');
        }
        // Keep a copy of every upload beside the run, exactly as the Launch & Design path does, so a
        // deck can be reproduced later from the files it was actually built from.
        const keep = [
            ['excel', excel], ['dashboard', dashboard], ['reference', reference],
        ];
        for (const [k, b] of keep) {
            const f = files?.[k]?.[0]
                || (k === 'excel' ? (files?.designExcel?.[0] || files?.buildExcel?.[0]) : undefined);
            if (b && f)
                storage.saveUpload(k, f.originalname, b);
        }
        const code = req.body.code || '', subtitle = req.body.subtitle || '';
        // TRIAL: Project Overview WBS view. Absent or 'default' → the shipped selector, unchanged.
        const overviewView = String(req.body.overviewView || 'default');
        const overviewCustomIds = String(req.body.overviewCustomIds || '')
            .split(',').map((s) => s.trim()).filter(Boolean);
        const r = await (0, engine_2.generateDesignBuild)({
            excel, dashboard, reference, code, subtitle,
            reportDate: req.body.reportDate || undefined,
            assets: loadAssets(), log,
            overviewView, overviewCustomIds,
        });
        const saved = storage.saveGenerated(r.project.name, {
            pptx: r.buffer, audit: '', log: logLines.join('\n'),
            'design_build_validation.md': r.validationMd,
            'validation_report.md': r.packageMd,
        });
        return res.json({
            jobId: saved.id, mode: 'design-build', project: r.project,
            audit: { overall: r.audit.overall, engine: 'node (pptxgenjs)', sections: r.audit.sections,
                missing: r.audit.missing, slides: r.audit.slides, scores: {}, editabilityIssues: [] },
            contents: r.contents,
            slideLabels: r.slideLabels,
            downloadUrl: `/api/download/${encodeURIComponent(saved.id)}`,
            designBuildValidationUrl: `/api/file/${encodeURIComponent(saved.id)}/design_build_validation.md`,
            validationReportUrl: `/api/file/${encodeURIComponent(saved.id)}/validation_report.md`,
            designBuildOk: r.validationOk,
            packageOk: r.packageResult.ok,
            fontWarnings: r.fontWarnings,
            overviewView: r.overview.view,
            overviewPages: r.overview.pages,
            overviewRows: r.overview.ganttRows,
            overviewStats: r.overview.stats,
        });
    }
    catch (e) {
        const f = (0, engine_2.dbFailureReport)(e, req.body?.name || '', req.body?.code || '');
        logLines.push('ERROR ' + f.info.stack);
        try {
            storage.saveGenerated(`FAILED - Design & Build - ${req.body?.name || 'run'}`, {
                pptx: Buffer.alloc(0), audit: '', log: logLines.join('\n'), 'error_report.md': f.md,
            });
        }
        catch { /* reporting must never mask the original failure */ }
        return fail(res, 500, f.stage, `${f.info.type}: ${f.info.message}`, `${f.info.file}:${f.info.line ?? '?'} in ${f.info.fn}. ${f.info.hint}`, f.info.stack, { log: logLines, errorReport: f.md });
    }
}
/**
 * TRIAL: the WBS tree of an uploaded schedule export, for the Custom view's picker.
 *
 * Returns the project root's subtree with depth measured RELATIVE TO THE PROJECT ROW, because an
 * absolute outline number means different things in different exports — one file nests the project
 * row at level 0, another at level 1. The browser ticks boxes against these ids and posts them back
 * as `overviewCustomIds`.
 */
app.post('/api/wbs-tree', uploadFields, async (req, res) => {
    try {
        const files = req.files;
        const buf = pick(files, 'excel') || pick(files, 'designExcel') || pick(files, 'buildExcel');
        if (!buf)
            return fail(res, 400, 'Reading the WBS hierarchy', 'No Excel export view was uploaded.', 'Upload the Excel export view first.');
        const sch = await (0, asta_excel_1.parseScheduleFromExcel)(buf);
        const shape = (0, branch_1.readExportShape)(sch, req.body.code || '');
        const entries = (0, overview_views_1.wbsEntries)(sch, shape);
        res.json({
            project: sch.projectName,
            root: shape.rootName,
            rows: entries.length,
            suggested: (0, overview_views_1.sampleCustomIds)(sch, shape),
            tree: entries.map((e) => ({
                id: e.id, name: e.name, depth: e.depth, parent: e.parent,
                hasChildren: e.hasChildren, milestone: e.milestone,
            })),
            notes: shape.notes,
        });
    }
    catch (e) {
        const f = (0, engine_2.dbFailureReport)(e);
        return fail(res, 500, 'Reading the WBS hierarchy', `${f.info.type}: ${f.info.message}`, `${f.info.file}:${f.info.line ?? '?'} in ${f.info.fn}. ${f.info.hint}`, f.info.stack);
    }
});
app.get('/api/download/:id', (req, res) => {
    const dir = storage.generatedFile(req.params.id, '');
    if (!fs_1.default.existsSync(dir))
        return res.status(404).send('Not found');
    const pptx = fs_1.default.readdirSync(dir).find((f) => f.endsWith('.pptx'));
    if (!pptx)
        return res.status(404).send('No PPTX');
    res.download(path_1.default.join(dir, pptx));
});
app.get('/api/file/:id/:file', (req, res) => {
    const p = storage.generatedFile(req.params.id, req.params.file);
    if (!fs_1.default.existsSync(p))
        return res.status(404).send('Not found');
    res.type('text/markdown').send(fs_1.default.readFileSync(p, 'utf8'));
});
app.get('/api/recent', (_req, res) => res.json(storage.listGenerated()));
const ROUTES = ['/api/health', '/api/detect', '/api/generate', '/api/wbs-tree', '/api/download/:id', '/api/file/:id/:file', '/api/recent'];
const PORT = Number(process.env.PORT || 4000);
function banner() {
    console.log('====================================================');
    console.log('EDOS Weekly Review Cloud');
    console.log(`Version              ${VERSION}`);
    console.log(`Node Version         ${process.version}`);
    console.log(`Working Directory    ${process.cwd()}`);
    console.log(`Backend Path         ${__dirname}`);
    console.log(`Resolved Config Path ${cfgPath || '(not found)'}`);
    console.log(`Config Exists        ${cfgPath ? 'yes' : 'no'}`);
    console.log(`Config Loaded        ${cfgRes.loaded}`);
    console.log(`Templates Loaded     ${assetsRes.count} (assets dir: ${assetsRes.dir || 'not found'})`);
    console.log(`Projects (approvals) ${projectsLoaded}`);
    console.log(`Assets Loaded        ${assetsRes.count}`);
    console.log(`Port                 ${PORT}`);
    console.log(`Max upload size      ${MAX_UPLOAD_MB} MB per file (EDOS_MAX_UPLOAD_MB)`);
    console.log(`API Routes Registered ${ROUTES.length}: ${ROUTES.join(', ')}`);
    console.log(`Startup Complete     API on http://localhost:${PORT}`);
    console.log('Keep this window open while using the app.');
    console.log('====================================================');
    if (!cfgRes.loaded)
        console.warn('[warn] edos_config.json not found — Tracker/Red-Flag Asta fallback will be empty. (Excel-driven slides still work.)');
    if (assetsRes.count === 0)
        console.warn('[warn] brand assets not found — cover/logos will be omitted (slides still generate).');
}
const server = app.listen(PORT, () => banner());
server.on('error', (e) => {
    console.error('\n[ERROR] Server failed to start');
    console.error(`  Error Type   ${e.code || e.name}`);
    console.error(`  Message      ${e.message}`);
    console.error(`  Port         ${PORT}`);
    if (e.code === 'EADDRINUSE') {
        console.error('  Cause        Port already in use — close the other "EDOS API" window, or set PORT to another value.');
    }
    console.error('  Stack        ' + (e.stack || '').split('\n').slice(0, 4).join('\n               '));
    console.error('  Exit Code    1');
    process.exitCode = 1;
});
