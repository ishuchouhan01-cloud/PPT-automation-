"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// Smoke test / headless generation without the web server.
// Usage: npx tsx src/cli.ts <excel.xlsx|.xlsm> [dashboard.xlsx] [approvals.xlsx] [reference.pptx]
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const engine_1 = require("./ppt/engine");
const audit_1 = require("./audit");
const asta_excel_1 = require("./parsers/asta_excel");
const paths_1 = require("./paths");
const assets = () => (0, paths_1.loadAssets)().assets;
(async () => {
    const [excel, dash, ref] = process.argv.slice(2);
    // Further block exports of the SAME project, one per additional reporting block:
    //   BLOCKS="South.xlsm" node dist/cli.js North.xlsm dashboard.xlsx previous.pptx
    // Separate several with ";". Omit it and the run is single-export, exactly as before.
    const extra = (process.env.BLOCKS || '').split(';').map((x) => x.trim()).filter(Boolean);
    if (!excel) {
        console.error('Usage: tsx src/cli.ts <excel> [dashboard] [reference]');
        process.exit(1);
    }
    const read = (p) => (p && fs_1.default.existsSync(p) ? fs_1.default.readFileSync(p) : undefined);
    const buf = fs_1.default.readFileSync(excel);
    const { buffer, audit, project, validation, layoutMd, overviewMd, trackerMd, trackerOk, trackerScheduleMd, trackerScheduleOk, scheduleLayoutMd, scheduleLayoutOk, scheduleReviewMd, scheduleDesignMd, scheduleDesignOk, approvalScheduleMd, approvalScheduleOk, packageMd, packageOk, rolesMd, rolesOk, redFlagMd, redFlagOk, consistencyMd, consistencyOk, keyFocusMd, keyFocusOk, errorMd, stagesOk, approvalMd, approvalOk, refineMd, refineOk } = await (0, engine_1.generate)({
        excel: buf, dashboard: read(dash), reference: read(ref),
        extraExcels: extra.map((f) => fs_1.default.readFileSync(f)),
        // No default code or subtitle: a hardcoded one (it used to be Violet Hill's "0118") makes the
        // dashboard match THAT project's row for every project. Pass CODE=… / SUBTITLE=… when you need them.
        code: process.env.CODE || '', subtitle: process.env.SUBTITLE || '',
        assets: assets(), log: (m) => console.log('  ' + m),
    });
    const outDir = path_1.default.join(__dirname, '..', '..', 'output');
    fs_1.default.mkdirSync(outDir, { recursive: true });
    const base = `${project.name} - Weekly Review`;
    fs_1.default.writeFileSync(path_1.default.join(outDir, base + '.pptx'), buffer);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'audit_report.md'), (0, audit_1.auditMarkdown)(audit, await (0, asta_excel_1.parseScheduleFromExcel)(buf)));
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'layout_validation.md'), layoutMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'project_overview_selection.md'), overviewMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'tracker_validation.md'), trackerMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'tracker_schedule_validation.md'), trackerScheduleMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'approval_validation.md'), approvalMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'presentation_refinement_validation.md'), refineMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'schedule_layout_validation.md'), scheduleLayoutMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'schedule_layout_review.md'), scheduleReviewMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'schedule_design_review.md'), scheduleDesignMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'approval_schedule_validation.md'), approvalScheduleMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'validation_report.md'), packageMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'lookahead_roles_validation.md'), rolesMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'red_flag_validation.md'), redFlagMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'presentation_consistency_validation.md'), consistencyMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'error_report.md'), errorMd);
    fs_1.default.writeFileSync(path_1.default.join(outDir, 'key_focus_layout_validation.md'), keyFocusMd);
    console.log(`\nSaved: output/${base}.pptx  (${(buffer.length / 1024).toFixed(0)} KB)`);
    console.log(`Audit overall: ${audit.overall}/100`);
    console.log(`Layout validation: ${validation.ok ? 'PASS' : 'FAIL'} ` +
        `(${validation.issues.filter((i) => i.severity === 'fail').length} failures, ` +
        `${validation.issues.filter((i) => i.severity === 'warn').length} warnings)`);
    console.log(`Tracker validation: ${trackerOk ? 'PASS' : 'FAIL'}`);
    console.log(`Tracker schedule validation: ${trackerScheduleOk ? 'PASS' : 'FAIL'}`);
    console.log(`Approval validation: ${approvalOk ? 'PASS' : 'FAIL'}`);
    console.log(`Presentation refinement: ${refineOk ? 'PASS' : 'FAIL'}`);
    console.log(`Schedule layout: ${scheduleLayoutOk ? 'PASS' : 'FAIL'}`);
    console.log(`Red Flag validation: ${redFlagOk ? 'PASS' : 'FAIL'}`);
    console.log(`Presentation consistency: ${consistencyOk ? 'PASS' : 'FAIL'}`);
    console.log(`Key Focus layout: ${keyFocusOk ? 'PASS' : 'FAIL'}`);
    console.log(`Stages: ${stagesOk ? 'all completed' : 'some sections skipped — see error_report.md'}`);
    console.log('Reports: output/layout_validation.md, output/project_overview_selection.md, output/tracker_validation.md, output/tracker_schedule_validation.md, output/approval_validation.md, output/presentation_refinement_validation.md, output/schedule_layout_validation.md, output/red_flag_validation.md, output/presentation_consistency_validation.md, output/error_report.md, output/key_focus_layout_validation.md');
    // A layout failure is a build failure.
    if (!validation.ok)
        process.exit(2);
})();
