"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.storageRoot = void 0;
exports.init = init;
exports.saveUpload = saveUpload;
exports.saveGenerated = saveGenerated;
exports.listGenerated = listGenerated;
exports.generatedFile = generatedFile;
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
// Cloud-ready versioned storage. Local disk today; the same interface can be backed by
// S3 / Azure Blob / GCS later without changing callers. Nothing is overwritten.
const ROOT = process.env.EDOS_STORAGE || path_1.default.join(__dirname, '..', '..', 'storage');
const SUBS = ['uploads/excel', 'uploads/dashboard', 'uploads/reference', 'uploads/land', 'generated'];
function init() {
    for (const s of SUBS)
        fs_1.default.mkdirSync(path_1.default.join(ROOT, s), { recursive: true });
}
const stamp = () => new Date().toISOString().replace(/[:.]/g, '-');
function saveUpload(kind, name, buf) {
    const dir = path_1.default.join(ROOT, 'uploads', kind);
    fs_1.default.mkdirSync(dir, { recursive: true });
    const safe = name.replace(/[^\w.\-() ]/g, '_');
    const p = path_1.default.join(dir, `${stamp()}__${safe}`); // versioned, never overwritten
    fs_1.default.writeFileSync(p, buf);
    return p;
}
function saveGenerated(project, files) {
    const id = `${stamp()}__${project.replace(/[^\w.\-() ]/g, '_')}`;
    const dir = path_1.default.join(ROOT, 'generated', id);
    fs_1.default.mkdirSync(dir, { recursive: true });
    const pptxPath = path_1.default.join(dir, `${project} - Weekly Review.pptx`);
    const auditPath = path_1.default.join(dir, 'audit_report.md');
    const logPath = path_1.default.join(dir, 'generation.log');
    fs_1.default.writeFileSync(pptxPath, files.pptx);
    fs_1.default.writeFileSync(auditPath, files.audit);
    fs_1.default.writeFileSync(logPath, files.log);
    // any extra markdown reports (layout_validation.md, project_overview_selection.md, …)
    for (const [k, v] of Object.entries(files)) {
        if (['pptx', 'audit', 'log'].includes(k) || typeof v !== 'string')
            continue;
        fs_1.default.writeFileSync(path_1.default.join(dir, k), v);
    }
    return { id, pptxPath, auditPath, logPath };
}
function listGenerated(limit = 20) {
    const dir = path_1.default.join(ROOT, 'generated');
    if (!fs_1.default.existsSync(dir))
        return [];
    return fs_1.default.readdirSync(dir)
        .filter((d) => fs_1.default.statSync(path_1.default.join(dir, d)).isDirectory())
        .sort().reverse().slice(0, limit)
        .map((id) => {
        const files = fs_1.default.readdirSync(path_1.default.join(dir, id));
        const pptx = files.find((f) => f.endsWith('.pptx')) || '';
        return { id, project: id.split('__')[1] || id, when: id.split('__')[0], pptx };
    });
}
function generatedFile(id, file) {
    return path_1.default.join(ROOT, 'generated', id, file);
}
const storageRoot = () => ROOT;
exports.storageRoot = storageRoot;
