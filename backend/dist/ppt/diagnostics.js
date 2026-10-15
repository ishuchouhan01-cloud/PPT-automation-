"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.StageError = exports.StageTracker = void 0;
exports.describeError = describeError;
exports.errorReport = errorReport;
// diagnostics.ts — stage tracking, structured errors and error_report.md.
//
// Every stage of a generation run is recorded here: when it started, whether it succeeded, and — if it
// threw — the file, function, line, message and full stack. Optional stages never stop the run: their
// section is simply left blank and the reason is stated. Only a failure that makes the .pptx physically
// impossible aborts, and even then the caller receives the exact stage and stack instead of a generic
// "Generation failed".
const path_1 = __importDefault(require("path"));
/** Pull file / function / line out of the first stack frame that belongs to this application. */
function describeError(e) {
    const err = e instanceof Error ? e : new Error(String(e));
    const stack = String(err.stack || '');
    const frames = stack.split('\n').slice(1).map((l) => l.trim());
    const own = frames.find((f) => /(dist|src)[\\/](ppt|parsers)?[\\/]?[\w.-]+\.js/.test(f) && !/node_modules/.test(f))
        || frames.find((f) => !/node_modules/.test(f)) || frames[0] || '';
    const m = own.match(/at\s+(?:async\s+)?([^\s(]+)?\s*\(?([^():]+):(\d+):(\d+)\)?/);
    return {
        message: err.message || String(err),
        type: err.name || err.constructor?.name || 'Error',
        file: m ? path_1.default.basename(m[2]) : '(unknown)',
        fn: m && m[1] ? m[1] : '(top level)',
        line: m ? +m[3] : null,
        column: m ? +m[4] : null,
        stack,
        hint: hintFor(err),
    };
}
/** Best-effort, human-readable suggestion for the most common causes. */
function hintFor(e) {
    const m = (e.message || '').toLowerCase();
    const type = e.name || '';
    if (type === 'MulterError' && m.includes('file too large')) {
        return 'One of the uploaded files is larger than the upload limit. Raise EDOS_MAX_UPLOAD_MB (the ' +
            'server prints the current value at startup) or supply a smaller previous PPT.';
    }
    if (m.includes('end of central directory')) {
        return 'The file is not a valid .xlsx/.xlsm/.pptx package — it may be an old .xls, a renamed file, ' +
            'or a partially downloaded copy. Re-download it and upload again.';
    }
    if (m.includes('enoent'))
        return 'A file or folder the app expected is missing — check Config/ and Assets/ are present.';
    if (m.includes('heap out of memory'))
        return 'The previous PPT is very large. Run the app with more memory ' +
            '(set NODE_OPTIONS=--max-old-space-size=4096) or upload a lighter reference deck.';
    if (m.includes('cannot read properties of undefined'))
        return 'A workbook or deck did not have the shape the ' +
            'parser expected — the reports name the sheet or slide involved.';
    return 'See the stack trace below; the failing file, function and line are named above.';
}
/** Runs the stages of a generation, recording every one of them. */
class StageTracker {
    constructor(log = () => { }) {
        this.log = log;
        this.stages = [];
        this.facts = {};
    }
    fact(k, v) { this.facts[k] = v; }
    /**
     * Run one stage.
     *  • optional: true  → a failure is recorded and `fallback` is returned, generation continues.
     *  • optional: false → a failure is recorded and re-thrown as a StageError, aborting the run.
     */
    async run(name, optional, fn, fallback) {
        const t0 = Date.now();
        this.log(`▶ ${name}…`);
        try {
            const out = await fn();
            this.stages.push({ name, optional, ok: true, ms: Date.now() - t0 });
            return out;
        }
        catch (e) {
            const err = describeError(e);
            this.stages.push({ name, optional, ok: false, ms: Date.now() - t0, error: err });
            this.log(`✗ ${name} FAILED — ${err.type}: ${err.message} (${err.file}:${err.line} in ${err.fn})`);
            if (optional) {
                this.log(`  · this section is optional — it is left blank and the deck continues`);
                return fallback;
            }
            throw new StageError(name, err);
        }
    }
    note(name, note) {
        const s = [...this.stages].reverse().find((x) => x.name === name);
        if (s)
            s.note = note;
    }
    get failed() { return this.stages.filter((s) => !s.ok); }
    get slidesBefore() { return Number(this.facts['Slides generated before the write step'] || 0); }
}
exports.StageTracker = StageTracker;
/** Thrown when a REQUIRED stage fails; carries everything the API needs to report. */
class StageError extends Error {
    constructor(stage, info) {
        super(`${stage}: ${info.message}`);
        this.stage = stage;
        this.info = info;
        this.name = 'StageError';
    }
}
exports.StageError = StageError;
function errorReport(opts) {
    const { tracker: tk } = opts;
    const L = [];
    const okAll = !opts.fatal && tk.failed.length === 0;
    L.push('# error_report.md', '');
    L.push(`Project: **${opts.project || '(not detected)'}**${opts.code ? `  ·  code **${opts.code}**` : ''}`, '');
    L.push(`## Result: ${okAll ? '**COMPLETED — no errors** ✅' : opts.fatal ? '**FAILED** ❌' : '**COMPLETED WITH SKIPPED SECTIONS** ⚠️'}`, '');
    if (opts.fatal) {
        L.push('## The failure', '', '| | |', '|---|---|', `| Stage | **${opts.fatalStage}** |`, `| Exception | \`${opts.fatal.type}: ${opts.fatal.message.replace(/\|/g, '\\|')}\` |`, `| File | \`${opts.fatal.file}\` |`, `| Function | \`${opts.fatal.fn}\` |`, `| Line | ${opts.fatal.line ?? '—'}${opts.fatal.column ? `, column ${opts.fatal.column}` : ''} |`, '', '**Suggested fix** — ' + opts.fatal.hint, '', '```', opts.fatal.stack.split('\n').slice(0, 25).join('\n'), '```', '');
    }
    L.push('## Run facts', '', '| Item | Value |', '|---|---|');
    for (const [k, v] of Object.entries(tk.facts))
        L.push(`| ${k} | ${v} |`);
    L.push('', '## Stages', '', '| # | Stage | Required | Result | Time | Detail |', '|---|---|---|---|---|---|');
    tk.stages.forEach((s, i) => L.push(`| ${i + 1} | ${s.name} | ${s.optional ? 'optional' : '**required**'} | ${s.ok ? '✅' : '❌'} | ${s.ms} ms | ` +
        `${(s.error ? `${s.error.type}: ${s.error.message} — ${s.error.file}:${s.error.line} in ${s.error.fn}` : (s.note || '')).replace(/\|/g, '\\|')} |`));
    const skipped = tk.failed.filter((s) => s.optional);
    if (skipped.length) {
        L.push('', '## Sections left blank (generation continued)', '');
        for (const s of skipped)
            L.push(`- **${s.name}** — ${s.error?.type}: ${s.error?.message}. ${s.error?.hint}`);
    }
    L.push('', '## Project name across the three inputs', '', '| Source | Name as supplied | Compared as |', '|---|---|---|');
    for (const n of opts.names)
        L.push(`| ${n.source} | ${n.name || '—'} | \`${n.normalised || '—'}\` |`);
    // Judged with the application's own matcher, so case, spacing, hyphens, underscores, double spaces,
    // trailing descriptors ("Plotted Development [FKH]") and common abbreviations are all ignored.
    const named = opts.names.filter((n) => n.name);
    // matchScore is the app's own scale: 0 = nothing in common, higher = more shared distinctive tokens,
    // 100 = an explicit alias group. Anything above 0 means the same project.
    const score = opts.match || ((a, b) => (a === b ? 100 : 0));
    const pairs = named.flatMap((a, i) => named.slice(i + 1).map((b) => ({ a, b, s: score(a.name, b.name) })));
    const mismatched = pairs.filter((p) => p.s <= 0);
    L.push('', mismatched.length === 0
        ? `_All ${named.length} supplied names refer to the same project once case, spacing, hyphens, underscores, ` +
            'double spaces and trailing descriptors are ignored' +
            (pairs.length ? ` (weakest pairing still scores ${Math.min(...pairs.map((p) => p.s))} on the app's matcher, where 0 would mean no match)` : '') + '._'
        : '_**The names do not match.** ' + mismatched.map((p) => `"${p.a.name}" (${p.a.source}) vs "${p.b.name}" (${p.b.source}) share no distinctive word`).join('; ') +
            '. All names are listed above exactly as they appear in their file; if these are the same project, add them ' +
            'to `Config/project_aliases.json` as one alias group._');
    if (opts.slides.length) {
        L.push('', '## Slides found in the previous PPT', '', '| Expected section | Present | Matched slide heading |', '|---|---|---|');
        for (const s of opts.slides)
            L.push(`| ${s.title} | ${s.found ? '✅' : '— not in this deck'} | ${s.matchedAs || ''} |`);
        L.push('', '_A section missing from the previous deck is not an error: it is generated from the data instead ' +
            '(or, for Land Acquisition / Master Plan, simply omitted rather than faked)._');
    }
    return L.join('\n');
}
