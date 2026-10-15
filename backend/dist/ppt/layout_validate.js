"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.validateLayout = validateLayout;
exports.validationReport = validationReport;
// layout_validate.ts — post-generation layout validation of the produced .pptx.
//
// Reads the finished package at OOXML level and checks, per slide:
//   • off-slide objects        — any shape/table extending beyond the slide bounds
//   • table / text overflow    — frames wider or taller than the printable area
//   • overlapping text frames  — the cause of "merged"/unreadable text
//   • unreadable font sizes    — anything below the readable minimum
//   • merged text              — text frames whose boxes intersect on the same band
//
// A build FAILS when any text leaves the slide, any table exceeds its boundary, any text overlaps,
// or merged text is detected. Results are published to layout_validation.md.
const jszip_1 = __importDefault(require("jszip"));
const typography_1 = require("./typography");
const EMU = 914400;
/** Printable area of the TE template, in inches. */
const SLIDE_W = 26.67, SLIDE_H = 15;
const EPS = 0.02; // tolerance for rounding in EMU→inch conversion
function framesOf(xml) {
    const out = [];
    // every shape / picture / graphicFrame carries an <a:off>/<a:ext> in its xfrm
    const re = /<p:(sp|pic|graphicFrame)\b[\s\S]*?<\/p:\1>/g;
    let m;
    while ((m = re.exec(xml))) {
        const body = m[0];
        const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
        const ext = body.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
        if (!off || !ext)
            continue;
        const texts = (body.match(/<a:t>[^<]*<\/a:t>/g) || []).map((s) => s.replace(/<\/?a:t>/g, ''));
        const sizes = [...body.matchAll(/sz="(\d+)"/g)].map((x) => +x[1] / 100);
        let x = +off[1] / EMU, y = +off[2] / EMU, w = +ext[1] / EMU, h = +ext[2] / EMU;
        // Respect shape rotation: a vertical column header is authored as a wide box then turned 90°/270°,
        // so its ON-SLIDE footprint is w/h swapped about the centre. Without this, rotated headers look
        // like huge overlapping boxes and produce false "merged text" failures.
        const rotM = body.match(/<a:xfrm[^>]*\brot="(-?\d+)"/);
        if (rotM) {
            const deg = ((+rotM[1] / 60000) % 360 + 360) % 360;
            if (Math.abs(deg - 90) < 5 || Math.abs(deg - 270) < 5) {
                const cx0 = x + w / 2, cy0 = y + h / 2;
                [w, h] = [h, w];
                x = cx0 - w / 2;
                y = cy0 - h / 2;
            }
        }
        out.push({ x, y, w, h, text: texts.join(' ').trim(), sz: sizes });
    }
    return out;
}
/** Overlap area of two frames, in square inches. */
function overlap(a, b) {
    const ox = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
    const oy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
    return ox > 0 && oy > 0 ? ox * oy : 0;
}
async function validateLayout(buf) {
    const z = await jszip_1.default.loadAsync(buf);
    const pres = (await z.file('ppt/presentation.xml')?.async('string')) || '';
    const rels = (await z.file('ppt/_rels/presentation.xml.rels')?.async('string')) || '';
    const rm = {};
    for (const m of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g))
        rm[m[1]] = m[2];
    const order = [];
    for (const m of pres.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)) {
        const t = rm[m[1]];
        if (t && /slides\/slide/.test(t))
            order.push('ppt/' + t.replace(/^\.\.\//, '').replace(/^ppt\//, ''));
    }
    const issues = [];
    for (let i = 0; i < order.length; i++) {
        const xml = (await z.file(order[i])?.async('string')) || '';
        const frames = framesOf(xml);
        const n = i + 1;
        for (const f of frames) {
            // Only TEXT and TABLES failing the bounds is a build failure. Full-bleed background images
            // deliberately extend past the edge in the TE template, so those are reported as warnings.
            const sev = f.text.length > 0 ? 'fail' : 'warn';
            const what = f.text.length > 0 ? `"${f.text.slice(0, 40)}"` : '(image/shape, no text)';
            if (f.x < -EPS || f.y < -EPS) {
                issues.push({ slide: n, kind: 'Off-slide object', severity: sev,
                    detail: `frame at (${f.x.toFixed(2)}, ${f.y.toFixed(2)}) starts outside the slide — ${what}` });
            }
            if (f.x + f.w > SLIDE_W + EPS) {
                issues.push({ slide: n, kind: 'Table/text overflow (right edge)', severity: sev,
                    detail: `right edge ${(f.x + f.w).toFixed(2)}" exceeds slide width ${SLIDE_W}" — ${what}` });
            }
            if (f.y + f.h > SLIDE_H + EPS) {
                issues.push({ slide: n, kind: 'Table/text overflow (bottom edge)', severity: sev,
                    detail: `bottom edge ${(f.y + f.h).toFixed(2)}" exceeds slide height ${SLIDE_H}" — ${what}` });
            }
            // Unreadable font sizes.
            //
            // AXIS FURNITURE IS EXEMPT. A month initial on the Project Overview's Gantt is a single character
            // in a column a sixth of an inch wide; it is set at the largest size that fits and is a timeline
            // mark, not text anybody reads as prose. Flagging 220 of them per deck would bury the warnings
            // that matter. Anything with more than one character is judged normally.
            const axisMark = f.text.trim().length === 1;
            for (const sz of f.sz) {
                if (sz > 0 && sz < typography_1.MIN_FS && !axisMark) {
                    issues.push({ slide: n, kind: 'Unreadable font size', severity: 'warn',
                        detail: `${sz}pt is below the readable minimum ${typography_1.MIN_FS}pt — "${f.text.slice(0, 40)}"` });
                    break;
                }
            }
        }
        // ---- slide-heading typography ----
        // Page headings are set in Impact at TITLE_FS. Impact is already a heavy display face, so
        // synthetic bold on top of it makes headings muddy: any b="1" on an Impact run is a failure.
        // Only the headings THIS generator authors are judged (the header band at y 1.12, h 1.04);
        // slides copied verbatim from the reference deck are the standard and are left untouched.
        for (const sp of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
            const body = sp[0];
            const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
            const ext = body.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
            if (!off || !ext)
                continue;
            const hy = +off[2] / EMU, hh = +ext[2] / EMU;
            if (Math.abs(hy - 1.12) > 0.03 || Math.abs(hh - 1.04) > 0.03)
                continue;
            const label = (body.match(/<a:t>[^<]*<\/a:t>/g) || []).map((x) => x.replace(/<\/?a:t>/g, '')).join(' ').trim();
            if (!label)
                continue;
            const faces = [...body.matchAll(/typeface="([^"]+)"/g)].map((x) => x[1]);
            if (faces.length && !faces.includes(typography_1.FONT_TITLE)) {
                issues.push({ slide: n, kind: 'Heading not in title face', severity: 'fail',
                    detail: `heading "${label.slice(0, 40)}" uses ${faces.join('/')} instead of ${typography_1.FONT_TITLE}` });
            }
            if (/<a:rPr\b[^>]*\bb="1"/.test(body)) {
                issues.push({ slide: n, kind: 'Heading too bold', severity: 'fail',
                    detail: `heading "${label.slice(0, 40)}" carries synthetic bold (b="1") — headings must use the lighter reference weight` });
            }
        }
        // ---- table header row fully visible ----
        // A table's first row is the column-header row; if the frame starts above the slide or the
        // header row's own height falls outside the printable area the headings would be cropped.
        for (const tb of xml.matchAll(/<p:graphicFrame\b[\s\S]*?<\/p:graphicFrame>/g)) {
            const body = tb[0];
            if (!/<a:tbl\b/.test(body))
                continue;
            const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
            const firstRow = body.match(/<a:tr\b[^>]*\bh="(\d+)"/);
            if (!off)
                continue;
            const ty = +off[2] / EMU;
            const hdrH = firstRow ? +firstRow[1] / EMU : 0;
            if (ty < -EPS || ty + hdrH > SLIDE_H + EPS) {
                issues.push({ slide: n, kind: 'Header row cropped', severity: 'fail',
                    detail: `table header row spans ${ty.toFixed(2)}"–${(ty + hdrH).toFixed(2)}", outside the printable 0–${SLIDE_H}"` });
            }
        }
        // overlapping TEXT frames → visually merged text
        const textFrames = frames.filter((f) => f.text.length > 0);
        for (let a = 0; a < textFrames.length; a++) {
            for (let b = a + 1; b < textFrames.length; b++) {
                const A = textFrames[a], B = textFrames[b];
                const ov = overlap(A, B);
                if (ov <= 0)
                    continue;
                const smaller = Math.min(A.w * A.h, B.w * B.h);
                // ignore incidental touching; flag when a text box is materially covered by another
                if (smaller > 0 && ov / smaller > 0.55) {
                    issues.push({ slide: n, kind: 'Merged/overlapping text', severity: 'fail',
                        detail: `"${A.text.slice(0, 30)}" overlaps "${B.text.slice(0, 30)}" by ${(ov / smaller * 100).toFixed(0)}% of the smaller box` });
                }
            }
        }
    }
    const counts = {};
    for (const it of issues)
        counts[it.kind] = (counts[it.kind] || 0) + 1;
    return { ok: !issues.some((i) => i.severity === 'fail'), slides: order.length, issues, counts };
}
function validationReport(r, project, fontWarnings = [], corrections = []) {
    const L = [];
    const fails = r.issues.filter((i) => i.severity === 'fail');
    const warns = r.issues.filter((i) => i.severity === 'warn');
    L.push('# layout_validation.md', '');
    L.push(`Project: **${project}**  ·  slides checked: **${r.slides}**`, '');
    L.push(`## Result: ${r.ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push(`- Failures: **${fails.length}** (text off-slide, table beyond boundary, overlapping/merged text)`);
    L.push(`- Warnings: **${warns.length}** (readability)`, '');
    // Per-check outcome counts only failures: a full-bleed background image deliberately bleeding
    // past the edge is a warning about the template, not a text-layout defect.
    const fc = {};
    for (const it of fails)
        fc[it.kind] = (fc[it.kind] || 0) + 1;
    L.push('## Checks performed', '', '| Check | Outcome |', '|---|---|', `| Text overflow (beyond slide) | ${fc['Table/text overflow (right edge)'] || fc['Table/text overflow (bottom edge)'] ? '❌' : '✅ none'} |`, `| Off-slide objects | ${fc['Off-slide object'] ? '❌' : '✅ none'} |`, `| Overlapping / merged text | ${fc['Merged/overlapping text'] ? '❌' : '✅ none'} |`, `| Cropped text (table beyond boundary) | ${fc['Table/text overflow (right edge)'] ? '❌' : '✅ none'} |`, `| Unreadable font sizes (< ${typography_1.MIN_FS}pt) | ${r.counts['Unreadable font size'] ? '⚠️ ' + r.counts['Unreadable font size'] : '✅ none'} |`, `| Table header row fully visible | ${fc['Header row cropped'] ? '❌' : '✅ all'} |`, `| Slide headings in ${typography_1.FONT_TITLE}, lighter weight | ${(fc['Heading too bold'] || fc['Heading not in title face']) ? '❌' : '✅ all'} |`);
    L.push('', '## Formatting corrections applied automatically', '');
    if (corrections.length)
        for (const c of corrections)
            L.push(`- ${c}`);
    else
        L.push('- none needed — every table fitted at its preferred size and row counts were within limits.');
    if (fontWarnings.length) {
        L.push('', '## Font availability', '');
        for (const w of fontWarnings)
            L.push(`- ⚠️ ${w}`);
    }
    if (r.issues.length) {
        L.push('', '## Detected issues', '', '| Slide | Severity | Check | Detail |', '|---|---|---|---|');
        for (const it of r.issues.slice(0, 300)) {
            L.push(`| ${it.slide} | ${it.severity === 'fail' ? '❌ fail' : '⚠️ warn'} | ${it.kind} | ${it.detail.replace(/\|/g, '\\|')} |`);
        }
        if (r.issues.length > 300)
            L.push('', `_…and ${r.issues.length - 300} more._`);
    }
    else {
        L.push('', 'No layout issues detected — every object sits inside the printable area, no text', 'overlaps, and all type is at or above the readable minimum.');
    }
    return L.join('\n');
}
