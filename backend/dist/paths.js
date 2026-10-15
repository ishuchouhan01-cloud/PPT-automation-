"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.resolveConfigPath = resolveConfigPath;
exports.loadConfig = loadConfig;
exports.loadAliases = loadAliases;
exports.loadStrictAliases = loadStrictAliases;
exports.loadTrackerScheduleConfig = loadTrackerScheduleConfig;
exports.loadApprovalConfig = loadApprovalConfig;
exports.resolveFontsDir = resolveFontsDir;
exports.resolveAssetsDir = resolveAssetsDir;
exports.loadAssets = loadAssets;
// Robust resolution of edos_config.json and the assets folder, regardless of whether the
// code runs from src/ (tsx) or dist/ (node), and regardless of the working directory.
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
function firstExisting(candidates) {
    for (const p of candidates) {
        try {
            if (fs_1.default.existsSync(p))
                return p;
        }
        catch { /* ignore */ }
    }
    return null;
}
function resolveConfigPath() {
    return firstExisting([
        path_1.default.join(__dirname, '..', '..', 'edos_config.json'), // dist/ or src/ → project root
        path_1.default.join(__dirname, '..', '..', '..', 'edos_config.json'), // dist/ppt or src/ppt → root
        path_1.default.join(process.cwd(), 'edos_config.json'),
        path_1.default.join(process.cwd(), '..', 'edos_config.json'),
        path_1.default.join(process.cwd(), 'backend', '..', 'edos_config.json'),
    ]);
}
function loadConfig() {
    const p = resolveConfigPath();
    if (p) {
        try {
            return { config: JSON.parse(fs_1.default.readFileSync(p, 'utf8')), loaded: true, pathUsed: p };
        }
        catch (e) {
            console.error('[config] failed to parse', p, e);
        }
    }
    return { config: { red_flag_subs: [], red_flag_macro: [], column_keywords: {}, approvals: {} }, loaded: false, pathUsed: p };
}
// project_aliases.json (optional, user-editable): groups of names that refer to the SAME project
// across the export / dashboard / approvals workbook, e.g. [["Shreegoan","Poinciana","ShreeGaon"]].
function aliasFile() {
    const p = firstExisting([
        path_1.default.join(__dirname, '..', '..', 'project_aliases.json'),
        path_1.default.join(__dirname, '..', '..', '..', 'project_aliases.json'),
        path_1.default.join(process.cwd(), 'project_aliases.json'),
        path_1.default.join(process.cwd(), '..', 'project_aliases.json'),
    ]);
    if (!p)
        return null;
    try {
        return JSON.parse(fs_1.default.readFileSync(p, 'utf8'));
    }
    catch (e) {
        console.error('[aliases] failed to parse', p, e);
        return null;
    }
}
/**
 * Alias groups for the FORGIVING matcher (matchScore) used by the Red Flag row, the Tracker row, the
 * approvals sheet lookup and the config lookups. Unchanged behaviour.
 *
 * project_aliases.json accepts two shapes:
 *   • a plain array of groups            — the original format, used by both matchers
 *   • { "groups": [...], "strictOnly": [...] }
 *       groups     — as before, used by BOTH matchers
 *       strictOnly — used by the STRICT matcher only (Key Focus Area)
 * `strictOnly` exists so that a new mapping can be declared for Key Focus Area without also loosening
 * the forgiving matcher. Adding "Madakuhosahalli" to a shared group would have made matchScore() treat
 * it as a 100% match for every OTHER Nandi Hills project too, because that matcher joins a group on a
 * shared token rather than on exact membership.
 */
function loadAliases() {
    const j = aliasFile();
    if (Array.isArray(j))
        return j.filter((g) => Array.isArray(g));
    if (j && Array.isArray(j.groups))
        return j.groups.filter((g) => Array.isArray(g));
    return [];
}
/** Alias groups visible to the STRICT matcher: the shared groups plus the strict-only ones. */
function loadStrictAliases() {
    const j = aliasFile();
    const shared = loadAliases();
    const strictOnly = (j && !Array.isArray(j) && Array.isArray(j.strictOnly))
        ? j.strictOnly.filter((g) => Array.isArray(g)) : [];
    return [...shared, ...strictOnly];
}
/** Tracker & Schedule milestone mapping (tracker_schedule_config.json). */
function loadTrackerScheduleConfig() {
    const p = firstExisting([
        path_1.default.join(__dirname, '..', '..', 'tracker_schedule_config.json'),
        path_1.default.join(__dirname, '..', '..', '..', 'tracker_schedule_config.json'),
        path_1.default.join(process.cwd(), 'tracker_schedule_config.json'),
        path_1.default.join(process.cwd(), '..', 'tracker_schedule_config.json'),
    ]);
    if (p) {
        try {
            return { config: JSON.parse(fs_1.default.readFileSync(p, 'utf8')), loaded: true, pathUsed: p };
        }
        catch (e) {
            console.error('[tracker config] failed to parse', p, e);
        }
    }
    return { config: { phases: [] }, loaded: false, pathUsed: p };
}
/** Approvals-slide mapping (approval_config.json). */
function loadApprovalConfig() {
    const p = firstExisting([
        path_1.default.join(__dirname, '..', '..', 'approval_config.json'),
        path_1.default.join(__dirname, '..', '..', '..', 'approval_config.json'),
        path_1.default.join(process.cwd(), 'approval_config.json'),
        path_1.default.join(process.cwd(), '..', 'approval_config.json'),
    ]);
    if (p) {
        try {
            return { config: JSON.parse(fs_1.default.readFileSync(p, 'utf8')), loaded: true, pathUsed: p };
        }
        catch (e) {
            console.error('[approval config] failed to parse', p, e);
        }
    }
    return { config: {}, loaded: false, pathUsed: p };
}
/** The bundled HR-approved font package (used only to warn when it is incomplete). */
function resolveFontsDir() {
    return firstExisting([
        path_1.default.join(__dirname, '..', '..', 'fonts'),
        path_1.default.join(__dirname, '..', '..', '..', 'fonts'),
        path_1.default.join(process.cwd(), 'fonts'),
        path_1.default.join(process.cwd(), '..', 'fonts'),
    ]);
}
function resolveAssetsDir() {
    return firstExisting([
        path_1.default.join(__dirname, '..', '..', 'assets'),
        path_1.default.join(__dirname, '..', '..', '..', 'assets'),
        path_1.default.join(process.cwd(), 'assets'),
        path_1.default.join(process.cwd(), '..', 'assets'),
    ]);
}
function loadAssets() {
    const dir = resolveAssetsDir();
    const assets = {};
    if (dir) {
        const map = { cover: 'cover.jpeg', mascot: 'mascot.png', wordmark: 'wordmark.png', thanks: 'thanks_bg.jpeg' };
        for (const [k, f] of Object.entries(map)) {
            const p = path_1.default.join(dir, f);
            try {
                if (fs_1.default.existsSync(p)) {
                    assets[k] = `data:${f.endsWith('png') ? 'image/png' : 'image/jpeg'};base64,${fs_1.default.readFileSync(p).toString('base64')}`;
                }
            }
            catch { /* ignore */ }
        }
    }
    return { assets, count: Object.keys(assets).length, dir };
}
