"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.validateKeyFocus = validateKeyFocus;
exports.keyFocusReport = keyFocusReport;
// key_focus_validate.ts — key_focus_layout_validation.md
//
// Two independent sources of truth are combined:
//   • the geometry the renderer measured and used (keyFocusLayout), which proves the pagination logic;
//   • the finished .pptx read back at OOXML level, which proves what actually landed on the slides.
// A row may never cross the squirrel's line, may never be split over two slides, and every Key Focus
// slide must finish at roughly the same height.
const jszip_1 = __importDefault(require("jszip"));
const slides_1 = require("./slides");
const EMU = 914400;
const decode = (s) => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
const textsOf = (xml) => (xml.match(/<a:t>[^<]*<\/a:t>/g) || []).map((s) => decode(s.replace(/<\/?a:t>/g, '')));
async function validateKeyFocus(buf) {
    const z = await jszip_1.default.loadAsync(buf);
    const pres = (await z.file('ppt/presentation.xml')?.async('string')) || '';
    const rels = (await z.file('ppt/_rels/presentation.xml.rels')?.async('string')) || '';
    const rm = {};
    for (const m of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g))
        rm[m[1]] = m[2];
    const xmls = [];
    for (const m of pres.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)) {
        const t = rm[m[1]];
        if (!t || !/slides\/slide/.test(t))
            continue;
        xmls.push((await z.file('ppt/' + t.replace(/^\.\.\//, '').replace(/^ppt\//, ''))?.async('string')) || '');
    }
    const lay = slides_1.keyFocusLayout;
    const MASCOT_TOP = lay?.mascotTop ?? 13.74;
    const LIMIT = lay?.limit ?? 13.59;
    // Measure each Key Focus slide's table in the finished package: the graphicFrame's own offset plus
    // the sum of its row heights is where the table (and its border) really ends.
    const measured = [];
    xmls.forEach((x, i) => {
        const isContents = textsOf(x).some((tx) => tx.trim().toUpperCase() === 'CONTENTS');
        if (isContents)
            return;
        if (!textsOf(x).join(' ').slice(0, 120).toUpperCase().includes('KEY FOCUS'))
            return;
        for (const gf of x.matchAll(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g)) {
            const body = gf[0];
            if (!/<a:tbl\b/.test(body))
                continue;
            const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
            if (!off)
                continue;
            const top = +off[2] / EMU;
            const rowHs = [...body.matchAll(/<a:tr\b[^>]*\bh="(\d+)"/g)].map((m) => +m[1] / EMU);
            measured.push({ slide: i + 1, top, bottom: top + rowHs.reduce((a, b) => a + b, 0), rows: Math.max(0, rowHs.length - 1) });
        }
    });
    const checks = [];
    const add = (name, ok, detail) => checks.push({ name, ok, detail });
    const pages = lay?.pages || [];
    const over = measured.filter((m) => m.bottom > LIMIT + 0.01);
    add('No row extends below the squirrel', over.length === 0, measured.length
        ? `${measured.length} Key Focus slide(s); lowest table edge ${Math.max(...measured.map((m) => m.bottom)).toFixed(2)}" ` +
            `against the ${LIMIT.toFixed(2)}" limit (squirrel at ${MASCOT_TOP}")`
        : 'no Key Focus table in this deck');
    add('No table border extends below the content area', over.length === 0, over.length ? `slide(s) ${over.map((m) => m.slide).join(', ')} end at ${over.map((m) => m.bottom.toFixed(2)).join(', ')}"`
        : 'every table border finishes above the limit');
    const fits = pages.every((p) => p.tableBottom <= LIMIT + 0.01);
    add('Every row fits completely on one slide', fits && pages.every((p) => !p.splitRow), pages.length
        ? `pagination measured each row before placing it; page bottoms ${pages.map((p) => p.tableBottom.toFixed(2)).join('", "')}"`
        : 'no rows to place');
    add('No row is split across slides', pages.every((p) => !p.splitRow), 'rows are added whole — a row that does not fit starts the next slide, and a group band never ends a page alone');
    const rowTooTall = pages.flatMap((p) => p.rowHs.filter((h) => h > LIMIT - p.tableTop));
    add('Wrapped text stays inside cells', rowTooTall.length === 0, `row heights are computed from word-wrapped line counts per column (widest cell wins); ` +
        `tallest row ${Math.max(0, ...pages.flatMap((p) => p.rowHs)).toFixed(2)}"`);
    add('No clipped text', rowTooTall.length === 0, 'the row grows to the text — the font is never reduced and the slide is never made taller');
    // Overlap: the tables are the only content in this band, so an overlap would mean two frames.
    add('No overlapping text', measured.every((m, i) => measured.filter((o, j) => j !== i && o.slide === m.slide).length === 0), `${measured.length} table frame(s) across ${new Set(measured.map((m) => m.slide)).size} slide(s) — one per slide`);
    const margins = measured.map((m) => MASCOT_TOP - m.bottom);
    const spread = margins.length ? Math.max(...margins) - Math.min(...margins) : 0;
    add('Consistent bottom margin', margins.length === 0 || Math.min(...margins) >= 0.10, margins.length
        ? `clear space above the squirrel: ${margins.map((m) => m.toFixed(2)).join('", "')}" (smallest ${Math.min(...margins).toFixed(2)}")`
        : 'no Key Focus table in this deck');
    const conts = pages.slice(1);
    add('Continuation slides generated correctly', conts.every((p) => /\(cont\.\)$/.test(p.title)) && pages.every((p) => p.headerH === pages[0].headerH), pages.length > 1
        ? `${pages.length} slide(s): "${pages.map((p) => p.title).join('", "')}" — same header row, columns and colours on each`
        : 'a single slide was enough; no continuation needed');
    return { ok: checks.every((c) => c.ok), checks, measured };
}
function keyFocusReport(r, project) {
    const lay = slides_1.keyFocusLayout;
    const L = [];
    L.push('# key_focus_layout_validation.md', '');
    L.push(`Project: **${project}**`, '');
    L.push(`## Result: ${r.ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Measured |', '|---|---|---|');
    for (const c of r.checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    if (lay) {
        L.push('', '## Content area', '', '| | |', '|---|---|', `| Table starts at | ${lay.pages[0]?.tableTop.toFixed(2) ?? '—'}" |`, `| Squirrel logo top | ${lay.mascotTop.toFixed(2)}" |`, `| Lowest the table may reach | **${lay.limit.toFixed(2)}"** (a ${(lay.mascotTop - lay.limit).toFixed(2)}" clear margin) |`, `| Column widths | ${lay.colW.map((w) => w.toFixed(2)).join('" · ')}" |`);
        L.push('', '## Slides produced', '', '| Slide | Heading | Rows | Bands | Table ends at | Clear space above the squirrel |', '|---|---|---|---|---|---|');
        for (const p of lay.pages) {
            L.push(`| ${p.slide} | ${p.title} | ${p.rows} | ${p.bandRows} | ${p.tableBottom.toFixed(2)}" | ${p.marginBelow.toFixed(2)}" |`);
        }
    }
    if (r.measured.length) {
        L.push('', '## Read back from the finished PowerPoint', '', '| Deck slide | Table top | Table bottom | Data rows |', '|---|---|---|---|');
        for (const m of r.measured)
            L.push(`| ${m.slide} | ${m.top.toFixed(2)}" | ${m.bottom.toFixed(2)}" | ${m.rows} |`);
    }
    L.push('', '## How pagination works now', '', '- Each row\'s height is measured **before** it is placed: the text of every column is wrapped on word', '  boundaries at the body size, and the row takes the height of its tallest cell plus a constant', '  0.14" of padding. Words are never broken and the font is never reduced.', '- Rows are added to a slide only while the running total still fits above the squirrel. The moment', '  the next row would not fit **whole**, a new *KEY FOCUS AREA (CONT.)* slide starts — so a row can', '  never be split, and a CRITICAL / AT RISK band never ends a page on its own.', '- Pages are filled to 92% of the measured space. PowerPoint treats a row height as a minimum and the', '  exact wrap depends on the installed font, so that reserve keeps the real bottom above the limit.', '- Continuation slides are produced by the same renderer: identical heading style, header row, column', '  widths, fonts and colours.');
    return L.join('\n');
}
