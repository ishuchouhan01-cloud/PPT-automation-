"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.codeInName = codeInName;
exports.resolveProjectContext = resolveProjectContext;
exports.guardSection = guardSection;
exports.assertNoContamination = assertNoContamination;
const match_1 = require("../parsers/match");
const norm = (v) => String(v ?? '').toLowerCase().replace(/[^a-z0-9]/g, '');
/** Digits-only comparison, so "0131" and "131" are one code. */
const numEq = (a, b) => /^\d+$/.test(a) && /^\d+$/.test(b) && Number(a) === Number(b);
/** A project code embedded in a row name: "0131 | Nandi Hills Homestay", "0116 - Meena and …". */
function codeInName(name) {
    const m = String(name || '').match(/^\s*([0-9]{3,4}(?:-[0-9]{1,2})?)\s*[|\-–:]/);
    return m ? m[1] : '';
}
/**
 * Resolve the project context.
 *
 * Never throws: with nothing to go on it returns a context with no code, which then falls back to
 * unambiguous name matching and says so in `notes`. The caller decides whether that is acceptable —
 * `dbContaminationGuard` will still refuse to ship a deck that mixes projects.
 */
function resolveProjectContext(inp) {
    const notes = [];
    const aliases = [];
    const rootName = (s) => {
        if (!s)
            return '';
        const rows = Object.values(s.wbs).sort((a, b) => a.seq - b.seq);
        // The project row is the topmost row carrying a code; failing that, the first row.
        const coded = rows.find((r) => !!codeInName(r.name));
        return (coded || rows[0])?.name || '';
    };
    const designRoot = rootName(inp.design), buildRoot = rootName(inp.build);
    if (designRoot)
        aliases.push(designRoot);
    if (buildRoot && buildRoot !== designRoot)
        aliases.push(buildRoot);
    // ---- 1. the operator ----
    let code = String(inp.operatorCode || '').trim();
    let codeSource = 'operator';
    if (code) {
        notes.push(`Project code "${code}" taken from the code entered on the form.`);
    }
    else {
        // ---- 2. the export's own project row ----
        const fromDesign = codeInName(designRoot), fromBuild = codeInName(buildRoot);
        if (fromDesign) {
            code = fromDesign;
            codeSource = 'design export';
            notes.push(`No code was entered, so it was read from the Design export's project row ` +
                `("${designRoot}") → "${code}".`);
        }
        else if (fromBuild) {
            code = fromBuild;
            codeSource = 'build export';
            notes.push(`No code was entered, so it was read from the Build export's project row ` +
                `("${buildRoot}") → "${code}".`);
        }
    }
    // ---- 3. an unambiguous dashboard row ----
    const displayFromExport = (0, match_1.cleanName)(designRoot || buildRoot || '') || 'Project';
    const searchName = (inp.design || inp.build)?.projectName || displayFromExport;
    let dashName = '';
    if (inp.dashboardRows?.length) {
        if (code) {
            const hit = inp.dashboardRows.find((r) => {
                const a = norm(code), b = norm(r.code);
                return !!a && !!b && (a === b || numEq(a, b));
            });
            if (hit) {
                dashName = hit.name;
                if (!aliases.includes(hit.name))
                    aliases.push(hit.name);
                notes.push(`Dashboard row for code "${code}" is "${hit.name}".`);
            }
            else {
                notes.push(`No dashboard row carries code "${code}". Dashboard-driven sections will be ` +
                    `omitted rather than filled from another project.`);
            }
        }
        else {
            // Rank by name, and accept ONLY a strict winner. Four projects share the token "Residence", so
            // "best score" alone is not enough — the winner has to beat the runner-up outright.
            const scored = inp.dashboardRows
                .map((r) => ({ r, s: (0, match_1.matchScore)(searchName, r.name) }))
                .filter((x) => x.s > 0)
                .sort((a, b) => b.s - a.s);
            if (scored.length === 1 || (scored.length > 1 && scored[0].s > scored[1].s)) {
                code = scored[0].r.code;
                codeSource = 'dashboard';
                dashName = scored[0].r.name;
                if (!aliases.includes(dashName))
                    aliases.push(dashName);
                notes.push(`No code was available from the form or the export, so it was resolved by name: ` +
                    `"${searchName}" matched dashboard row "${dashName}" (score ${scored[0].s}` +
                    (scored.length > 1 ? `, next best "${scored[1].r.name}" at ${scored[1].s}` : ', no other candidate') +
                    `) → code "${code}".`);
            }
            else if (scored.length > 1) {
                notes.push(`No code was available and the project name "${searchName}" matches ` +
                    `${scored.length} dashboard rows equally well (` +
                    scored.slice(0, 4).map((x) => `"${x.r.name}" ${x.s}`).join(', ') +
                    `). Refusing to guess: dashboard-driven sections are omitted. Enter the project code to ` +
                    `resolve this.`);
            }
            else {
                notes.push(`No code was available and the project name "${searchName}" matches no dashboard row.`);
            }
        }
    }
    if (!code) {
        codeSource = 'none';
        notes.push('WARNING: no project code could be resolved from the form, the export or the ' +
            'dashboard. Filtering falls back to unambiguous name matching, which is weaker. Enter the ' +
            'project code to make scoping exact.');
    }
    const name = (0, match_1.cleanName)(dashName || displayFromExport);
    if (!aliases.includes(name))
        aliases.push(name);
    const strictName = (rowName) => {
        if (!rowName)
            return false;
        // Every alias is tried, because the three sources spell the project differently
        // ("Meena & Krishna Sai Residence" / "0116 - Meena and KrishnaSai Residence" /
        // "Meena&KrishnaSai Residence").
        return aliases.some((a) => (0, match_1.matchScore)(a, rowName) > 0);
    };
    return {
        code, codeSource, name, aliases, reportDate: inp.reportDate, notes,
        owns(rowCode, rowName) {
            const rc = norm(rowCode);
            if (code) {
                const c = norm(code);
                // WITH a code the code DECIDES. A row that carries a different code is rejected outright and
                // its name is never consulted — that is the whole point.
                if (rc)
                    return c === rc || numEq(c, rc);
                // A row with no code of its own falls back to the name, so a sheet that omits the column
                // still works.
                return strictName(String(rowName || ''));
            }
            return strictName(String(rowName || ''));
        },
        describe(rowCode, rowName) {
            const rc = String(rowCode ?? '').trim();
            if (code && rc)
                return `row code "${rc}" ≠ project code "${code}"`;
            if (code && !rc)
                return `row has no code; name "${String(rowName || '').trim()}" does not match ${aliases.map((a) => `"${a}"`).join(' / ')}`;
            return `name "${String(rowName || '').trim()}" does not match ${aliases.map((a) => `"${a}"`).join(' / ')}`;
        },
    };
}
/**
 * Assert that every record about to be rendered belongs to the current project.
 *
 * Called per section, on the rows that will actually be drawn — not on the source data. A section
 * that passes here cannot put another project's row on the slide, because these are the same objects
 * the renderer receives.
 */
function guardSection(ctx, section, records) {
    const foreign = [];
    for (const r of records) {
        if (ctx.owns(r.code, r.name))
            continue;
        foreign.push({ what: r.what, why: ctx.describe(r.code, r.name) });
    }
    return { section, checked: records.length, foreign };
}
/**
 * Fail the run when any section carries a foreign record.
 *
 * Deliberately fatal. A weekly review that quietly attributes another project's commitments to this
 * one is worse than no deck: somebody acts on it. The message names the section, the row and the
 * reason so the cause is obvious.
 */
function assertNoContamination(results) {
    const bad = results.filter((r) => r.foreign.length);
    if (!bad.length) {
        return { ok: true, report: results.map((r) => `${r.section}: ${r.checked} record(s), all belong to this project`).join('\n') };
    }
    const lines = bad.map((r) => `${r.section}: ${r.foreign.length} of ${r.checked} record(s) belong to another project — ` +
        r.foreign.slice(0, 6).map((f) => `${f.what} (${f.why})`).join('; ') +
        (r.foreign.length > 6 ? ` … and ${r.foreign.length - 6} more` : ''));
    return { ok: false, report: lines.join('\n') };
}
