"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.validateConsistency = validateConsistency;
exports.consistencyReport = consistencyReport;
// consistency_validate.ts — presentation_consistency_validation.md
//
// Checks that are only meaningful on the ASSEMBLED deck (after the reference slides have been merged
// in), read back from the package at OOXML level:
//
//   • every 2-Week Look Ahead slide sits in one uninterrupted block, with Site Progress (or any other
//     carried-over slide) after it, never between two Look Ahead pages;
//   • the Look Ahead table has no "Line" column and its headings are larger than its body text;
//   • the section-number box and the page title never touch on any generated slide;
//   • every generated header uses the same two colours, the same way round;
//   • the Approvals stage headings are larger than the approval data;
//   • the four approval stages obey Application → Drawings → Fee → Issued.
const jszip_1 = __importDefault(require("jszip"));
const typography_1 = require("./typography");
const EMU = 914400;
const decode = (s) => s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
const textsOf = (xml) => (xml.match(/<a:t>[^<]*<\/a:t>/g) || []).map((s) => decode(s.replace(/<\/?a:t>/g, '')));
function headerBoxes(xml) {
    // the header band the generator authors: y 1.12, h 1.04
    const out = [];
    for (const m of xml.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
        const body = m[0];
        const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
        const ext = body.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
        if (!off || !ext)
            continue;
        const y = +off[2] / EMU, h = +ext[2] / EMU;
        if (Math.abs(y - 1.12) > 0.03 || Math.abs(h - 1.04) > 0.03)
            continue;
        const spPr = body.split('<p:txBody')[0];
        const fill = spPr.match(/<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/);
        out.push({ x: +off[1] / EMU, y, w: +ext[1] / EMU, h, text: textsOf(body).join(' ').trim(), fill: fill ? fill[1].toUpperCase() : null });
    }
    return out;
}
function tablesOf(xml) {
    const tables = [];
    for (const tb of xml.matchAll(/<a:tbl\b[\s\S]*?<\/a:tbl>/g)) {
        const rows = [];
        for (const tr of tb[0].matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g))
            rows.push([...tr[0].matchAll(/<a:tc\b[\s\S]*?<\/a:tc>/g)].map((c) => c[0]));
        tables.push(rows);
    }
    return tables;
}
const sizesOf = (cell) => [...cell.matchAll(/sz="(\d+)"/g)].map((m) => +m[1] / 100);
const cellText = (cell) => textsOf(cell).join(' ').trim();
async function validateConsistency(buf, approvals) {
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
    const checks = [];
    const add = (name, ok, detail) => checks.push({ name, ok, detail });
    // ---------- slide order ----------
    // The Contents slide lists every section title, so it must never be mistaken for a section slide.
    const isContents = (x) => textsOf(x).some((t) => t.trim().toUpperCase() === 'CONTENTS');
    const heading = (x) => (isContents(x) ? 'CONTENTS' : textsOf(x).join(' ').slice(0, 120).toUpperCase());
    const order = xmls.map((x, i) => {
        const h = heading(x);
        if (i === 0)
            return 'Cover';
        if (h.includes('CONTENTS'))
            return 'Contents';
        if (h.includes('RED FLAG'))
            return 'Red Flag Summary';
        if (h.includes('KEY FOCUS'))
            return 'Key Focus Area';
        if (h.includes('TRACKER'))
            return 'Tracker & Schedule';
        if (h.includes('PROJECT OVERVIEW'))
            return 'Project Overview';
        if (h.includes('LOOK AHEAD'))
            return '2 Week Look Ahead';
        if (h.includes('APPROVAL'))
            return 'Approvals Status';
        if (h.includes('LAND ACQUISITION'))
            return 'Land Acquisition';
        if (h.includes('MASTER PLAN'))
            return 'Master Plan';
        if (h.includes('THANK YOU'))
            return 'Thank You';
        return `Other (${textsOf(x).join(' ').trim().slice(0, 28) || 'no text'})`;
    });
    const laIdx = order.map((s, i) => (s === '2 Week Look Ahead' ? i : -1)).filter((i) => i >= 0);
    const laContiguous = laIdx.length === 0 || laIdx[laIdx.length - 1] - laIdx[0] + 1 === laIdx.length;
    add('All Look Ahead slides appear together', laContiguous, laIdx.length ? `${laIdx.length} Look Ahead slide(s) at positions ${laIdx.map((i) => i + 1).join(', ')}` : 'no Look Ahead slide in this deck');
    const between = laIdx.length ? order.slice(laIdx[0], laIdx[laIdx.length - 1] + 1).filter((s) => s !== '2 Week Look Ahead') : [];
    add('No Site Progress slide appears between Look Ahead slides', between.length === 0, between.length ? `interrupted by: ${between.join(', ')}` : 'nothing sits between the first and last Look Ahead slide');
    // ---------- Look Ahead table ----------
    // use the CLASSIFIED order, so a slide that merely mentions a section name is never mistaken for it
    const laXmls = xmls.filter((_, i) => order[i] === '2 Week Look Ahead');
    let laHdr = [], laHdrSizes = [], laBodySizes = [];
    let laHdrBold = 0, laHdrCells = 0, laHdrColours = new Set();
    for (const x of laXmls)
        for (const tbl of tablesOf(x)) {
            const head = tbl[0] || [];
            if (!laHdr.length) {
                laHdr = head.map(cellText);
                laHdrSizes = head.flatMap(sizesOf);
            }
            for (const c of head) {
                laHdrCells++;
                if (/<a:rPr[^>]*\bb="1"/.test(c))
                    laHdrBold++;
                for (const m of c.matchAll(/<a:rPr[\s\S]{0,200}?<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"/g))
                    laHdrColours.add(m[1].toUpperCase());
            }
            for (const r of tbl.slice(1)) {
                if (r.length < 4)
                    continue;
                laBodySizes.push(...r.flatMap(sizesOf));
            }
        }
    add('Line column removed', laXmls.length === 0 || !laHdr.some((h) => /^line$/i.test(h)), laHdr.length ? `columns: ${laHdr.join(' · ')}` : 'no Look Ahead table in this deck');
    // ---- the outer boundary is drawn BY THE CELLS of the last row ----
    // The frame used to be a rectangle sized to the sum of the DECLARED row heights. Those are minimums —
    // a renderer grows a row whose text needs more — so the drawn table could be taller than the rectangle
    // and the final row fell outside the border. In the FILE the two always agreed, which is why measuring
    // the package could not catch it.
    //
    // The outer edges are now cell borders, so the boundary is the table. This checks that directly: the
    // last row of every Look Ahead table must carry the frame colour on its BOTTOM edge, the header row on
    // its TOP, and the outer columns on their LEFT/RIGHT — and no separate frame rectangle may remain.
    // The frame colour is the shade the Look Ahead's own header bar uses — read off the slide rather than
    // assumed, so the check follows a theme change instead of hard-coding one.
    const FRAME = (() => {
        for (const x of laXmls) {
            for (const m of x.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
                const b = m[0];
                if (/<a:t>/.test(b))
                    continue;
                const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(b);
                const fl = /<a:solidFill><a:srgbClr val="([0-9A-Fa-f]{6})"\/><\/a:solidFill>/.exec(b);
                if (off && fl && Math.abs((+off[1]) / 914400 - 1.43) < 0.02)
                    return fl[1].toUpperCase();
            }
        }
        return '80350E';
    })();
    let laLastRowClosed = 0, laLastRowTotal = 0, laStrayFrames = 0;
    const edgeMissing = [];
    laXmls.forEach((x, i) => {
        const tbl = /<a:tbl>[\s\S]*?<\/a:tbl>/.exec(x);
        if (!tbl)
            return;
        const rows = [...tbl[0].matchAll(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g)].map((m) => m[0]);
        if (rows.length < 2)
            return;
        laLastRowTotal += 1;
        const cells = (tr) => [...tr.matchAll(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g)].map((m) => m[0]);
        const hasEdge = (cell, tag) => new RegExp(`<a:ln${tag}[^>]*>[\\s\\S]{0,240}?<a:srgbClr val="${FRAME}"`, 'i').test(cell);
        const lastCells = cells(rows[rows.length - 1]);
        const closed = lastCells.length > 0 && lastCells.every((c) => hasEdge(c, 'B'));
        if (closed)
            laLastRowClosed += 1;
        else
            edgeMissing.push(`page ${i + 1}: bottom edge`);
        const headCells = cells(rows[0]);
        if (!(headCells.length && headCells.every((c) => hasEdge(c, 'T'))))
            edgeMissing.push(`page ${i + 1}: top edge`);
        // A horizontally merged band row emits one styled cell plus `hMerge` fillers that carry no borders of
        // their own, so the right edge of those rows is tested on the last STYLED cell, not the filler.
        const styled = (c) => c.filter((x) => !/hMerge="1"/.test(x));
        const leftOk = rows.every((r) => { const c = styled(cells(r)); return !c.length || hasEdge(c[0], 'L'); });
        const rightOk = rows.every((r) => { const c = styled(cells(r)); return !c.length || hasEdge(c[c.length - 1], 'R'); });
        if (!leftOk)
            edgeMissing.push(`page ${i + 1}: left edge`);
        if (!rightOk)
            edgeMissing.push(`page ${i + 1}: right edge`);
        // A leftover full-width frame RECTANGLE would reintroduce the drift. The header bar is also full
        // width, so only shapes taller than 2" and below the header band count.
        for (const m of x.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)) {
            const b = m[0];
            if (/<a:t>/.test(b))
                continue;
            const off = /<a:off x="(\d+)" y="(\d+)"\/>/.exec(b);
            const ext = /<a:ext cx="(\d+)" cy="(\d+)"\/>/.exec(b);
            if (off && ext && (+ext[1]) / 914400 > 20 && (+ext[2]) / 914400 > 2 && (+off[2]) / 914400 > 1.9)
                laStrayFrames += 1;
        }
    });
    add('Look Ahead boundary is closed by the last row itself', laLastRowTotal === 0 || (laLastRowClosed === laLastRowTotal && edgeMissing.length === 0 && laStrayFrames === 0), laLastRowTotal
        ? `${laLastRowClosed}/${laLastRowTotal} page(s) close on their final row · left/right/top edges ` +
            `${edgeMissing.length ? 'INCOMPLETE: ' + edgeMissing.slice(0, 4).join('; ') : 'complete'} · ` +
            `${laStrayFrames} leftover frame rectangle(s)`
        : 'no Look Ahead table in this deck');
    const maxBody = laBodySizes.length ? Math.max(...laBodySizes) : typography_1.LOOKAHEAD_FS;
    add(`Look Ahead headings are ${typography_1.HEADER_FS_LOOKAHEAD}pt`, laXmls.length === 0 || (laHdrSizes.length > 0 && laHdrSizes.every((v) => v === typography_1.HEADER_FS_LOOKAHEAD)), laHdrSizes.length ? `every heading ${[...new Set(laHdrSizes)].join('/')}pt; body text up to ${maxBody}pt (unchanged)` : 'no Look Ahead table in this deck');
    add('Look Ahead headings are regular weight, white', laXmls.length === 0 || (laHdrBold === 0 && (laHdrColours.size === 0 || laHdrColours.has('FFFFFF'))), laHdrCells ? `${laHdrCells - laHdrBold}/${laHdrCells} heading cells regular weight; colour ${[...laHdrColours].map((c) => '#' + c).join(', ') || 'inherited'}` : 'no Look Ahead table in this deck');
    // ---------- header band ----------
    let overlaps = 0, headerSlides = 0;
    const badgeFills = new Set(), barFills = new Set();
    // Scope: slides this generator AUTHORS. A slide copied verbatim out of the customer's own deck keeps
    // its own header shade — re-colouring it would change their branding, which every brief has forbidden.
    // Such slides are identified exactly as refine_validate identifies them: by referencing a typeface
    // this generator never writes. They are counted separately instead of failing the check.
    // (Before this, merging a reference deck made the check fail on #E97132 — the template's own accent on
    // its own slides. The rule was asserting something the generator does not control.)
    const OURS_FACES = new Set(['Meta-Normal', 'Helvetica Neue', 'Impact']);
    // A slide copied out of the customer's deck marks itself TE-REF-COPY (see postmerge.copyRefSlide);
    // that is authoritative. The font test remains as a fallback for decks produced before the marker.
    const authoredSlide = (x) => !/name="TE-REF-COPY"/.test(x) &&
        [...new Set([...x.matchAll(/typeface="([^"]+)"/g)].map((m) => m[1]))].every((f) => OURS_FACES.has(f));
    const copiedHeaders = new Set();
    let minGap = 99;
    for (const x of xmls) {
        const boxes = headerBoxes(x).filter((b) => b.fill); // the two filled shapes
        if (boxes.length < 2)
            continue;
        headerSlides++;
        const sorted = boxes.slice().sort((a, b) => a.x - b.x);
        const badge = sorted[0], bar = sorted[1];
        if (!authoredSlide(x)) {
            if (badge.fill)
                copiedHeaders.add(badge.fill);
            if (bar.fill)
                copiedHeaders.add(bar.fill);
            continue; // the customer's own slide, in the customer's own colours
        }
        // EVERY authored slide, the 2 Week Look Ahead included, must use the same header scheme. The Look
        // Ahead's own lighter bar was a specified exception for one release; it was reverted, so the exception
        // is gone from this check too and the slide is held to the same rule as the rest of the deck.
        badgeFills.add(badge.fill);
        barFills.add(bar.fill);
        const gap = bar.x - (badge.x + badge.w);
        minGap = Math.min(minGap, gap);
        if (gap < 0)
            overlaps++;
        // the title's own text box must also start after the badge
        for (const tb of headerBoxes(x).filter((b) => b.text && !b.fill)) {
            if (tb.x < badge.x + badge.w - 0.01 && tb.x > badge.x)
                overlaps++;
        }
    }
    add('No Project Overview header overlap', overlaps === 0, `${headerSlides} generated header(s) checked; smallest gap between the number box and the title bar ${minGap === 99 ? 'n/a' : minGap.toFixed(2) + '"'}`);
    // The page number and the page title are one unit: both are set at the title size on every slide.
    // A two-character label ("4A", "4B") used to be reduced to 68%, which made the Project Overview
    // pages look like a smaller heading than the rest of the deck.
    const titleSizes = new Set(), badgeSizes = new Set();
    const badgeLabels = [], copiedBadges = [];
    for (const x of xmls) {
        // Only slides this generator AUTHORS. A slide copied verbatim out of the customer's deck carries their
        // own badge size — Nandi Hills' previous deck sets 3B, 6A, 6B and 6C at 38pt — and re-typesetting it
        // would change their branding. Copied badges are counted separately instead of failing the check, the
        // same discipline the header-colour and Meta-Normal checks already use.
        if (!authoredSlide(x)) {
            for (const m of x.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
                const body = m[0];
                const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
                if (!off || Math.abs(+off[2] / EMU - 1.12) > 0.03)
                    continue;
                const label = textsOf(body).join(' ').trim();
                if (/^[0-9]{1,2}[A-Z]?$/.test(label)) {
                    const sz = [...body.matchAll(/sz="(\d+)"/g)].map((v) => +v[1] / 100);
                    if (sz.length)
                        copiedBadges.push(sz[0]);
                }
            }
            continue;
        }
        for (const m of x.matchAll(/<p:sp\b[\s\S]*?<\/p:sp>/g)) {
            const body = m[0];
            const off = body.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/);
            if (!off || Math.abs(+off[2] / EMU - 1.12) > 0.03)
                continue;
            const label = textsOf(body).join(' ').trim();
            if (!label)
                continue;
            const sz = [...body.matchAll(/sz="(\d+)"/g)].map((v) => +v[1] / 100);
            if (!sz.length)
                continue;
            if (/^[0-9]{1,2}[A-Z]?$/.test(label)) {
                badgeSizes.add(sz[0]);
                badgeLabels.push(label);
            }
            else
                titleSizes.add(sz[0]);
        }
    }
    const allTitleSized = [...badgeSizes, ...titleSizes].every((v) => v === typography_1.TITLE_FS);
    add(`Page numbers and titles are all ${typography_1.TITLE_FS}pt`, allTitleSized && badgeSizes.size > 0, `page numbers ${[...badgeSizes].join('/')}pt (${badgeLabels.join(', ')}) · titles ${[...titleSizes].join('/')}pt` +
        (copiedBadges.length ? ` · ${copiedBadges.length} badge(s) on slides copied verbatim keep the customer's own ${[...new Set(copiedBadges)].join('/')}pt` : ''));
    add('All slide headers use consistent colours', badgeFills.size <= 1 && barFills.size <= 1, `number box ${[...badgeFills].map((c) => '#' + c).join(', ') || 'n/a'} · title bar ${[...barFills].map((c) => '#' + c).join(', ') || 'n/a'}` +
        (badgeFills.size > 1 || barFills.size > 1 ? ' — more than one scheme found' : ' on every generated slide') +
        (copiedHeaders.size ? ` · slides copied verbatim from the reference deck keep their own #${[...copiedHeaders].join(', #')}` : ''));
    // The Look Ahead heading must be indistinguishable from the other sections': same bar, same badge.
    add('2 Week Look Ahead heading matches the other sections', !laHdrCells || (barFills.size <= 1 && badgeFills.size <= 1), `bar ${[...barFills].map((c) => '#' + c).join(', ') || 'n/a'} · badge ${[...badgeFills].map((c) => '#' + c).join(', ') || 'n/a'} — one scheme across every authored slide`);
    // ---------- Approvals ----------
    const apXml = xmls.find((_, i) => order[i] === 'Approvals Status');
    if (apXml) {
        const tbl = tablesOf(apXml)[0] || [];
        const hdrSizes = (tbl[0] || []).flatMap(sizesOf);
        const bodySizes = tbl.slice(1).flatMap((r) => r.flatMap(sizesOf));
        const dataMax = bodySizes.length ? Math.max(...bodySizes.filter((v) => v !== typography_1.HEADER_FS_APPROVAL)) : typography_1.BODY_FS;
        add('Approval headings enlarged', hdrSizes.length > 0 && Math.min(...hdrSizes) === typography_1.HEADER_FS_APPROVAL, `stage headings ${[...new Set(hdrSizes)].join('/')}pt (was ${18}pt) against approval data at up to ${dataMax}pt`);
    }
    else {
        // A project can legitimately have no approvals — a construction-phase deck has none — in which case
        // the section is not emitted at all and there is nothing to check. Not applicable, not a failure.
        add('Approval headings enlarged', true, 'no Approvals Status section in this deck — not applicable');
    }
    const bad = approvals.filter((a) => {
        const f = [a.application, a.drawings, a.fee, a.issued];
        const hi = f.lastIndexOf(true);
        return f.some((v, i) => !v && i < hi);
    });
    const fixed = approvals.filter((a) => (a.cascaded || []).length);
    add('Approval stages satisfy cascading completion logic', bad.length === 0, bad.length
        ? `impossible sequences remain: ${bad.map((a) => a.approval).join(', ')}`
        : `${approvals.length} approval(s) follow Application → Drawings Submission → Fee Payment → Issued` +
            (fixed.length ? `; ${fixed.length} had an earlier stage promoted to Completed (${fixed.map((a) => `${a.approval}: ${(a.cascaded || []).join(', ')}`).join(' · ')})` : ''));
    return { ok: checks.every((c) => c.ok), checks, order };
}
function consistencyReport(r, project, layoutOk, layoutSummary) {
    const L = [];
    const ok = r.ok && layoutOk;
    L.push('# presentation_consistency_validation.md', '');
    L.push(`Project: **${project}**`, '');
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Detail |', '|---|---|---|');
    for (const c of r.checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    L.push(`| No overlapping objects | ${layoutOk ? '✅' : '❌'} | ${layoutSummary} |`);
    L.push(`| No text clipping | ${layoutOk ? '✅' : '❌'} | ${layoutSummary} |`);
    L.push('', '## Slide order as assembled', '', '| # | Slide |', '|---|---|');
    r.order.forEach((s, i) => L.push(`| ${i + 1} | ${s} |`));
    L.push('', '## Rules applied', '', '- **Section contiguity** — every slide of a section stays in one block. A slide carried over from', '  the previous deck (Site Progress and the like) is placed after the *last* slide of the section it', '  follows, and any slide found interrupting a section is moved back behind it.', '- **2 Week Look Ahead** — the *Line* column is gone and its width was shared out across the', `  remaining columns; the column headings are ${typography_1.HEADER_FS_LOOKAHEAD}pt against ${typography_1.LOOKAHEAD_FS}pt body text.`, '- **Header band** — one fixed-width number box and a separate title bar starting 0.08" later, with', '  the title text inset a further 0.22". A two-character label such as "4A" is set slightly smaller', '  so it stays inside its box; the two can never touch.', '- **Header colours** — standardised and swapped: the number box now carries the shade the page', '  title used, and the title bar carries the shade the number box used, on every generated slide.', `- **Approvals** — stage headings at ${typography_1.HEADER_FS_APPROVAL}pt over unchanged row heights, and the four`, '  stages are made consistent with the workflow: every stage up to the highest completed one is shown', '  as complete, so a row can never read ✔ ✔ ✘ ✔. Stages after the highest completed one are left', '  exactly as the schedule reports them, and every promotion is listed above and in', '  `approval_validation.md`.');
    return L.join('\n');
}
