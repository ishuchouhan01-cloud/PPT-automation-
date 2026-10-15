"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.slideHeading = slideHeading;
exports.comingFriday = comingFriday;
exports.updateCoverDate = updateCoverDate;
exports.mergeStaticSlides = mergeStaticSlides;
exports.refSlideTitles = refSlideTitles;
// postmerge.ts — copy STATIC slides (Land Acquisition, Master Plan) verbatim from the uploaded
// reference .pptx into the freshly generated deck, replacing the labelled placeholders.
//
// Pure Node, no Office/Python. Works at the OOXML (zip) level with JSZip. pptxgenjs cannot import
// foreign slides, so we splice the reference slide part + its media into the generated package and
// swap it in at the placeholder's position in the slide order. Layout is retargeted to the
// generated deck's slideLayout1 (content is explicit shapes, so a generic layout is safe).
//
// If the reference lacks a requested slide (e.g. Master Plan), that slide is simply OMITTED
// (its placeholder is removed) — per spec: "Master Plan omitted if absent". Never fabricated.
const jszip_1 = __importDefault(require("jszip"));
const txt = (z, p) => z.file(p)?.async('string');
// concatenated visible text of a slide part (for title matching). Decodes XML entities so that a
// title like "TRACKER &amp; SCHEDULE" normalises to "trackerschedule" (not "trackerampschedule").
/**
 * A slide's HEADING — the first few text runs, which on these decks is the section number followed by the
 * title. Matching a spec against the whole slide text (even truncated) is unsafe: the Red Flag Summary's
 * column headers include "Land Acquisition", so a search for the Land Acquisition slide matched the Red
 * Flag slide and copied it in — the deck ended up with two Red Flag pages and no Land slide, while
 * Contents still advertised one.
 */
function slideHeading(xml) {
    const runs = (xml.match(/<a:t>[^<]*<\/a:t>/g) || [])
        .map((t) => t.replace(/<\/?a:t>/g, '').replace(/&amp;/g, '&').trim())
        .filter(Boolean);
    return runs.slice(0, 3).join(' ');
}
/** Resolve a relationship Target against the part that declares it: "ppt/slides/slide3.xml" + "../media/x.png" → "ppt/media/x.png". */
function resolvePart(fromPart, target) {
    if (target.startsWith('/'))
        return target.slice(1);
    const out = fromPart.split('/').slice(0, -1);
    for (const seg of target.split('/')) {
        if (seg === '.' || seg === '')
            continue;
        if (seg === '..')
            out.pop();
        else
            out.push(seg);
    }
    return out.join('/');
}
function slideText(xml) {
    return (xml.match(/<a:t>[^<]*<\/a:t>/g) || [])
        .map((s) => s.replace(/<\/?a:t>/g, ''))
        .join(' ')
        .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'");
}
// presentation slide order → [{ rId, part }]
async function order(z) {
    const pres = (await txt(z, 'ppt/presentation.xml')) || '';
    const rels = (await txt(z, 'ppt/_rels/presentation.xml.rels')) || '';
    const relMap = {};
    for (const m of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g))
        relMap[m[1]] = m[2].replace(/^\/?/, '');
    const out = [];
    for (const m of pres.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)) {
        const t = relMap[m[1]];
        if (t)
            out.push({ rId: m[1], part: 'ppt/' + t.replace(/^\.\.\//, '').replace(/^ppt\//, '') });
    }
    return out;
}
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
    'September', 'October', 'November', 'December'];
const ordinal = (d) => (d % 10 === 1 && d !== 11) ? 'st' : (d % 10 === 2 && d !== 12) ? 'nd' : (d % 10 === 3 && d !== 13) ? 'rd' : 'th';
/** The coming Friday (or today if today is already Friday). */
function comingFriday(from = new Date()) {
    const d = new Date(from.getFullYear(), from.getMonth(), from.getDate());
    const add = (5 - d.getDay() + 7) % 7; // 5 = Friday
    d.setDate(d.getDate() + add);
    return d;
}
// ---------------------------------------------------------------------------------------------
// Typography of the cover date. The cover is copied from the reference deck verbatim, so its runs
// carry the TEMPLATE's fonts — which set the year and the ordinal in Meta Bold. Two runs are
// re-faced to the HR-approved roles, and nothing else on the slide is touched:
//   • the YEAR      → Meta-Normal, regular weight (it is a number, and numbers are Meta-Normal)
//   • the ORDINAL   → Helvetica Neue, keeping its superscript
// Size, colour, spacing, position and every other run are left exactly as the template has them.
const FACE_YEAR = 'Meta-Normal';
const FACE_ORDINAL = 'Helvetica Neue';
/**
 * The run-properties element of a run, matched as a WHOLE element.
 *
 * This must handle both forms, and the order matters:
 *   <a:rPr .../>                                        — self-closing, no children
 *   <a:rPr ...><a:latin .../><a:solidFill>…</a:rPr>     — with children
 *
 * A single lazy pattern like `/<a:rPr[\s\S]*?(?:\/>|<\/a:rPr>)/` is WRONG: on the second form it stops
 * at the first `/>` it meets, which belongs to the child `<a:latin .../>`, and returns an rPr with no
 * closing tag. Splicing that into a new run produced unbalanced XML and PowerPoint's repair dialog.
 */
const RPR_RE = /<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/;
/**
 * Re-face one run's properties. Works whether `<a:rPr/>` is self-closing or has children, keeps every
 * attribute the template set (size, colour, spacing …) and only replaces the typeface elements.
 */
function reface(rPr, face, opts = {}) {
    const self = rPr ? rPr.match(/^<a:rPr\b([^>]*?)\/>$/) : null;
    const open = !self && rPr ? rPr.match(/^<a:rPr\b([^>]*)>([\s\S]*)<\/a:rPr>$/) : null;
    let attrs = self ? self[1] : open ? open[1] : '';
    let inner = open ? open[2] : '';
    if (opts.unbold)
        attrs = attrs.replace(/\s+b="[^"]*"/g, '');
    if (opts.bold)
        attrs = /\bb="/.test(attrs) ? attrs.replace(/\s+b="[^"]*"/g, ' b="1"') : `${attrs} b="1"`;
    if (opts.superscript && !/\bbaseline="/.test(attrs))
        attrs += ' baseline="30000"';
    inner = inner.replace(/<a:(latin|ea|cs)\b[^>]*\/>/g, '');
    return `<a:rPr${attrs}><a:latin typeface="${face}"/><a:cs typeface="${face}"/>${inner}</a:rPr>`;
}
/** Rewrite the date on a copied cover slide to the coming Friday, keeping the template's run layout.
 *  Handles day / ordinal / month / year in separate runs, or a combined "24th" run. Scoped to the
 *  paragraph that contains "Update" (or a month + year) so the project code/acreage are untouched. */
function updateCoverDate(xml, target = comingFriday(), 
/**
 * OPTIONAL, opt-in, and additive: omitting it reproduces the previous behaviour exactly.
 *
 * `inline` also rewrites dates that are EMBEDDED in a longer run rather than sitting in a run of
 * their own. Some covers split the date into "17" / "th" / "August" / "2026", which the run-exact
 * rules below already handle. Others glue it to its label — the Design & Build reference covers read
 * `"Weekly Update: 14"` · `"th"` · `" August 2026"`, three runs in which only the ordinal is ever a
 * standalone token. On those covers the exact rules changed the "th" and nothing else, so a
 * regenerated deck still carried the old day, month and year.
 *
 * It is opt-in because the Launch & Design reference covers use the split form and are verified
 * byte-for-byte against it; enabling this for them would rewrite runs that currently pass through
 * untouched.
 */
opts = {}) {
    const day = target.getDate(), mon = MONTHS[target.getMonth()], yr = target.getFullYear(), ord = ordinal(day);
    const monthRe = new RegExp(`^(${MONTHS.join('|')})$`, 'i');
    const monthAnyRe = new RegExp(`\\b(${MONTHS.join('|')})\\b`, 'i');
    return xml.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g, (para) => {
        const plain = (para.match(/<a:t>[^<]*<\/a:t>/g) || []).map((t) => t.replace(/<\/?a:t>/g, '')).join(' ');
        const looksLikeDate = /update/i.test(plain) || (/\b20\d\d\b/.test(plain) && monthAnyRe.test(plain));
        if (!looksLikeDate)
            return para;
        // Work at RUN level, not text level, so a run's typeface can be changed with its text.
        return para.replace(/<a:r>([\s\S]*?)<\/a:r>/g, (run, body) => {
            const tm = body.match(/<a:t>([\s\S]*?)<\/a:t>/);
            if (!tm)
                return run;
            const rPr = (body.match(RPR_RE) || [undefined])[0];
            const txt = tm[1];
            const lead = (txt.match(/^\s*/) || [''])[0], trail = (txt.match(/\s*$/) || [''])[0];
            const s = txt.trim();
            const mk = (props, v) => `<a:r>${props}<a:t>${lead}${v}${trail}</a:t></a:r>`;
            // "17th" in ONE run — split it so the number keeps the template's face and the ordinal can be
            // set in Helvetica Neue as its own superscript run.
            if (/^\d{1,2}(st|nd|rd|th)$/i.test(s)) {
                return `<a:r>${rPr || ''}<a:t>${lead}${day}</a:t></a:r>` +
                    `<a:r>${reface(rPr, FACE_ORDINAL, { superscript: true })}<a:t>${ord}${trail}</a:t></a:r>`;
            }
            if (/^\d{1,2}$/.test(s) && +s >= 1 && +s <= 31)
                return mk(rPr || '', String(day)); // day
            if (/^(st|nd|rd|th)$/i.test(s))
                return mk(reface(rPr, FACE_ORDINAL, { superscript: true }), ord);
            if (monthRe.test(s))
                return mk(rPr || '', mon); // month
            // The front page's year is BOLD, in Meta-Normal. (2.16.0 set it regular; that was reversed.)
            if (/^\d{4}$/.test(s) && +s >= 2000)
                return mk(reface(rPr, FACE_YEAR, { bold: true }), String(yr));
            // ---- embedded form, only when asked for ----
            // The date parts sit inside a longer run, e.g. "Weekly Update: 14" or " August 2026". Each part
            // is substituted in place and the run's own typeface is left alone, because the run also carries
            // label text that must not change face.
            if (opts.inline) {
                let v = txt, hit = false;
                // Year first: doing the day first would let a 4-digit year's last two digits look like a day.
                v = v.replace(/\b(20\d\d)\b/g, () => { hit = true; return String(yr); });
                v = v.replace(new RegExp(`\\b(${MONTHS.join('|')})\\b`, 'gi'), () => { hit = true; return mon; });
                // A day number with its ordinal attached, or on its own, but never a digit inside a longer
                // number (a project code such as "0131" must survive untouched).
                v = v.replace(/(^|[^\d])(\d{1,2})(st|nd|rd|th)?(?![\d])/gi, (m0, pre, d, o) => {
                    const n = +d;
                    if (n < 1 || n > 31)
                        return m0;
                    hit = true;
                    return `${pre}${day}${o ? ord : ''}`;
                });
                if (hit && v !== txt)
                    return `<a:r>${rPr || ''}<a:t>${v}</a:t></a:r>`;
            }
            return run;
        });
    });
}
/**
 * Delete a slide part and EVERYTHING that depends on it: its own .rels, its content-type override,
 * and any notesSlide that belongs to it (part, rels and override). Returns the updated [Content_Types].
 *
 * Anything left pointing at a deleted part is a dangling relationship, and PowerPoint treats a
 * dangling relationship as a damaged package.
 */
async function dropSlide(g, ct, part) {
    const relPart = part.replace(/slides\/(slide\d+)\.xml/, 'slides/_rels/$1.xml.rels');
    // the notesSlide(s) this slide owns, read from its own relationships before it is deleted
    const rels = (await txt(g, relPart)) || '';
    const notes = [];
    for (const m of rels.matchAll(/Target="\.\.\/(notesSlides\/[^"]+)"/g))
        notes.push(`ppt/${m[1]}`);
    for (const n of notes) {
        g.remove(n);
        g.remove(n.replace(/notesSlides\/(notesSlide\d+)\.xml/, 'notesSlides/_rels/$1.xml.rels'));
        ct = ct.replace(new RegExp(`<Override PartName="/${n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), '');
    }
    g.remove(part);
    g.remove(relPart);
    ct = ct.replace(new RegExp(`<Override PartName="/${part.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), '');
    return ct;
}
/** Replace placeholder slides with verbatim copies of matching reference slides. */
async function mergeStaticSlides(genBuf, refBuf, specs, log = () => { }, opts = {}) {
    const copied = [];
    const g = await jszip_1.default.loadAsync(genBuf);
    const r = await jszip_1.default.loadAsync(refBuf);
    const gOrder = await order(g);
    const rOrder = await order(r);
    // title -> generated slide text (cache)
    const gText = {};
    for (const s of gOrder)
        gText[s.part] = slideText((await txt(g, s.part)) || '');
    const rText = {};
    for (const s of rOrder)
        rText[s.part] = slideText((await txt(r, s.part)) || '');
    const gHead = {};
    for (const s of gOrder)
        gHead[s.part] = slideHeading((await txt(g, s.part)) || '');
    const rHead = {};
    for (const s of rOrder)
        rHead[s.part] = slideHeading((await txt(r, s.part)) || '');
    const norm = (s) => s.toLowerCase().replace(/[^a-z]/g, '');
    // Match on the slide HEADING only (first ~80 normalised chars) and never the Contents slide —
    // Contents lists every section title, so a whole-text "includes" would false-match it.
    const isContents = (s) => norm(s).startsWith('contents');
    // Match on the slide's HEADING (its first few runs), never on the body. An 80-character slice of the
    // whole text reached into the first table row, which is how "Land Acquisition" found the Red Flag
    // Summary slide.
    const findGen = (ph) => gOrder.find((s) => !isContents(gText[s.part]) && norm(gHead[s.part]).includes(norm(ph)));
    const findRef = (t) => rOrder.find((s) => !isContents(rText[s.part]) && norm(rHead[s.part]).includes(norm(t)));
    // spec → the generated slide it replaces, and the reference slide it copies
    const genFor = (spec) => spec.position === 'first' ? gOrder[0]
        : spec.position === 'last' ? gOrder[gOrder.length - 1]
            : (spec.placeholderTitle ? findGen(spec.placeholderTitle) : undefined);
    const refFor = (spec) => spec.position === 'first' ? rOrder[0]
        : spec.position === 'last' ? rOrder[rOrder.length - 1]
            : (findRef(spec.title) || (spec.refFallbackFirst ? rOrder[0] : undefined));
    let presRels = (await txt(g, 'ppt/_rels/presentation.xml.rels')) || '';
    let pres = (await txt(g, 'ppt/presentation.xml')) || '';
    let ct = (await txt(g, '[Content_Types].xml')) || '';
    const existingSlideNums = Object.keys(g.files)
        .map((f) => f.match(/^ppt\/slides\/slide(\d+)\.xml$/)).filter(Boolean).map((m) => +m[1]);
    let nextSlide = Math.max(0, ...existingSlideNums) + 1;
    let nextMedia = Object.keys(g.files).filter((f) => /^ppt\/media\//.test(f)).length + 1;
    const maxRel = Math.max(0, ...[...presRels.matchAll(/Id="rId(\d+)"/g)].map((m) => +m[1]));
    let nextRel = maxRel + 1;
    let nextSldId = 9000; // ids for freshly-inserted extra sldId elements
    const mediaMap = new Map(); // reference media name → copied name (dedupe)
    // Copy one reference slide (xml + media + rels) into g as a new part; returns its new rId.
    async function copyRefSlide(refPart, transform) {
        const newPart = `ppt/slides/slide${nextSlide}.xml`;
        let slideXml = (await txt(r, refPart)) || '';
        if (transform === 'coverDate') {
            slideXml = updateCoverDate(slideXml, opts.coverDate || comingFriday(), { inline: opts.coverDateInline });
        }
        // PROVENANCE MARKER. A slide copied verbatim out of the customer's deck keeps their typography and
        // their badge sizes, and the validators must not judge it by this generator's rules. Identifying such
        // slides by "does it use a font we never write" was a heuristic that failed here: Nandi Hills' own
        // Site Progress slides use the same three faces, so their 38pt badges were reported as our
        // inconsistency. The slide now says what it is, in its own name attribute, which is a fact rather
        // than an inference. PowerPoint ignores the name; it is metadata.
        slideXml = /<p:cSld\b[^>]*\bname="/.test(slideXml)
            ? slideXml.replace(/(<p:cSld\b[^>]*\bname=")([^"]*)(")/, '$1TE-REF-COPY$3')
            : slideXml.replace(/<p:cSld\b/, '<p:cSld name="TE-REF-COPY"');
        const refRelsPath = refPart.replace(/slides\/(slide\d+)\.xml/, 'slides/_rels/$1.xml.rels');
        let relsXml = (await txt(r, refRelsPath)) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
        // ---- copy EVERY internal part this slide depends on, not just its images ----
        // A slide copied out of the customer's deck can reference far more than media: an embedded Excel
        // workbook (ppt/embeddings/*.xlsx|xlsm) behind a pasted table, an OLE object, a chart. Only
        // ../media/ used to be copied, so those relationships survived while their targets did not — and
        // "rId8 → ../embeddings/Microsoft_Excel_Worksheet5.xlsx, which is not in the package" is precisely
        // the dangling relationship PowerPoint reports as a damaged file. The integrity gate caught it
        // before it shipped; this is the underlying fix.
        //
        // slideLayouts are re-pointed below and notesSlides/comments are dropped, so they are skipped here.
        const CT_BY_EXT = {
            png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
            tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml',
            mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', wav: 'audio/wav', mp3: 'audio/mpeg',
            xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
            xls: 'application/vnd.ms-excel',
            docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            doc: 'application/msword',
            pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            pdf: 'application/pdf', bin: 'application/vnd.openxmlformats-officedocument.oleObject',
            zip: 'application/zip', txt: 'text/plain', xml: 'application/xml',
        };
        const ensureDefault = (ext) => {
            if (new RegExp(`Extension="${ext}"`, 'i').test(ct))
                return;
            const type = CT_BY_EXT[ext] || 'application/octet-stream';
            ct = ct.replace('</Types>', `<Default Extension="${ext}" ContentType="${type}"/></Types>`);
        };
        for (const mm of [...relsXml.matchAll(/<Relationship\b[^>]*\/>/g)]) {
            const tag = mm[0];
            if (/TargetMode="External"/.test(tag))
                continue; // a hyperlink, nothing to copy
            const tgt = (/\bTarget="([^"]+)"/.exec(tag) || [])[1];
            if (!tgt)
                continue;
            if (/slideLayouts\//.test(tgt) || /notesSlides\//.test(tgt) || /comments/.test(tgt))
                continue;
            const srcPart = resolvePart(refPart, tgt); // e.g. ppt/embeddings/…xlsx
            const srcFile = r.file(srcPart);
            if (!srcFile)
                continue; // already missing in the source deck
            const base = srcPart.split('/').pop();
            const ext = (base.includes('.') ? base.split('.').pop() : 'bin').toLowerCase();
            const folder = srcPart.slice(0, srcPart.lastIndexOf('/')); // ppt/media, ppt/embeddings, …
            // One copy per source part, shared by every slide that references it: reference decks reuse the
            // same background image and the same embedded workbook across many slides.
            let newName = mediaMap.get(srcPart);
            if (!newName) {
                newName = folder.endsWith('/media')
                    ? `image_ref${nextMedia++}.${ext}`
                    : `ref${nextMedia++}_${base}`;
                g.file(`${folder}/${newName}`, await srcFile.async('nodebuffer'));
                mediaMap.set(srcPart, newName);
            }
            relsXml = relsXml.split(`${tgt}"`).join(`${tgt.slice(0, tgt.lastIndexOf('/') + 1)}${newName}"`);
            ensureDefault(ext);
        }
        relsXml = relsXml.replace(/Target="\.\.\/slideLayouts\/[^"]+"/g, 'Target="../slideLayouts/slideLayout1.xml"');
        relsXml = relsXml.replace(/<Relationship[^>]*Target="\.\.\/(notesSlides|comments)[^"]*"[^>]*\/>/g, '');
        g.file(newPart, slideXml);
        g.file(`ppt/slides/_rels/slide${nextSlide}.xml.rels`, relsXml);
        ct = ct.replace('</Types>', `<Override PartName="/${newPart}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
        const newRId = `rId${nextRel++}`;
        presRels = presRels.replace('</Relationships>', `<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${nextSlide}.xml"/></Relationships>`);
        nextSlide++;
        return newRId;
    }
    for (const spec of specs) {
        const gp = genFor(spec);
        if (!gp) {
            log(`  · placeholder "${spec.placeholderTitle || spec.title}" not found — skipping`);
            continue;
        }
        const rp = refFor(spec);
        if (!rp) {
            pres = pres.replace(new RegExp(`<p:sldId[^>]*r:id="${gp.rId}"[^>]*/>`), '');
            // The placeholder is gone from the slide list, so its part, rels, notes and overrides must go too.
            ct = await dropSlide(g, ct, gp.part);
            presRels = presRels.replace(new RegExp(`<Relationship Id="${gp.rId}"[^>]*/>`), '');
            log(`  · "${spec.title}" absent in reference — omitted`);
            continue;
        }
        const newRId = await copyRefSlide(rp.part, spec.transform);
        // swap the placeholder's sldId to point at the copied slide (keeps its position)
        pres = pres.replace(new RegExp(`(<p:sldId[^>]*)r:id="${gp.rId}"([^>]*/>)`), `$1r:id="${newRId}"$2`);
        // Remove the now-orphaned placeholder part/rels/content-type/pres-rel — AND anything that pointed
        // at it. pptxgenjs writes a notesSlide for every slide, and each notesSlide holds a relationship
        // back to its own slide. Deleting the slide alone left that notesSlide behind with a dangling
        // reference to a part that no longer exists, which is a missing resource: PowerPoint reports the
        // package as damaged and offers to repair it.
        ct = await dropSlide(g, ct, gp.part);
        presRels = presRels.replace(new RegExp(`<Relationship Id="${gp.rId}"[^>]*/>`), '');
        log(`  · copied "${spec.title}" from reference`);
        copied.push(spec.title);
    }
    // ---- Extras pass: copy any reference slide that is NOT one of the finalized standard sections,
    // placing it at the same relative position (right after the generated slide for the standard
    // slide that preceded it in the reference). e.g. "SITE PHOTOS" pages after the Look-Ahead. ----
    const stdKey = (text, first) => {
        const tl = text.toLowerCase();
        if (first)
            return 'cover';
        if (tl.includes('contents'))
            return 'contents';
        if (tl.includes('red flag'))
            return 'redflag';
        if (tl.includes('key focus'))
            return 'keyfocus';
        if (tl.includes('tracker'))
            return 'tracker';
        if (tl.includes('project overview'))
            return 'overview';
        if (tl.includes('land acquisition'))
            return 'land';
        if (tl.includes('master plan'))
            return 'master';
        if (tl.includes('approval') && tl.includes('status'))
            return 'approvals';
        if (tl.includes('look ahead'))
            return 'lookahead';
        if (/\|/.test(text) && /acres|plotted|residential|towers|villas/.test(tl))
            return 'cover';
        if (/total-environment|discover@/.test(tl))
            return 'thanks';
        // Mode-specific sections get their say only after every standard one has declined, so a caller
        // that supplies no hook behaves exactly as before.
        if (opts.stdKeyExtra) {
            const k = opts.stdKeyExtra(tl);
            if (k)
                return k;
        }
        return null; // → an "extra" slide
    };
    // current generated order (from the live pres/presRels strings) + text + anchor per standard key
    const orderOf = (presStr, relsStr) => {
        const rm = {};
        for (const m of relsStr.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g))
            rm[m[1]] = m[2];
        const out = [];
        for (const m of presStr.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)) {
            const t = rm[m[1]];
            if (t && /slides\/slide/.test(t))
                out.push({ rId: m[1], part: 'ppt/' + t.replace(/^\.\.\//, '').replace(/^ppt\//, '') });
        }
        return out;
    };
    const cur = opts.skipExtras ? [] : orderOf(pres, presRels);
    const anchorByKey = {};
    for (let i = 0; i < cur.length; i++) {
        const k = stdKey(slideText((await txt(g, cur[i].part)) || '').slice(0, 120), i === 0);
        if (k)
            anchorByKey[k] = cur[i].rId; // last generated slide of each section
    }
    let anchorRId = cur[0]?.rId;
    let extraCount = 0;
    for (let i = 0; !opts.skipExtras && i < rOrder.length; i++) {
        const k = stdKey(rText[rOrder[i].part].slice(0, 120), i === 0);
        if (k) {
            if (anchorByKey[k])
                anchorRId = anchorByKey[k];
            continue;
        }
        if (!anchorRId)
            continue;
        if (opts.skipHiddenExtras) {
            const sx = (await txt(r, rOrder[i].part)) || '';
            if (/<p:sld\b[^>]*\bshow="0"/.test(sx)) {
                log(`  · skipped ${rOrder[i].part} — hidden in the previous deck`);
                continue;
            }
        }
        const newRId = await copyRefSlide(rOrder[i].part);
        pres = pres.replace(new RegExp(`(<p:sldId[^>]*r:id="${anchorRId}"[^>]*/>)`), `$1<p:sldId id="${nextSldId++}" r:id="${newRId}"/>`);
        anchorRId = newRId; // chain consecutive extras in order
        extraCount++;
    }
    if (extraCount)
        log(`  · copied ${extraCount} extra slide(s) from reference at their original positions`);
    // ---- section contiguity ----
    // A generated section can span several slides (the 2-Week Look Ahead often runs to three). An extra
    // slide copied from the previous deck must never land in the middle of one: the whole section has to
    // read as one block, with the next section starting only after its last slide. Any interrupting
    // slide is moved to just after the final slide of the section it split.
    {
        const ord = orderOf(pres, presRels);
        const keys = [];
        for (let i = 0; i < ord.length; i++)
            keys.push(stdKey(slideText((await txt(g, ord[i].part)) || '').slice(0, 120), i === 0));
        const lastOf = {};
        keys.forEach((k, i) => { if (k)
            lastOf[k] = i; });
        const moved = [];
        for (let i = 0; i < ord.length; i++) {
            const k = keys[i];
            // find a section that starts before i and ends after i, while this slide belongs to another
            const interrupted = Object.keys(lastOf).find((sec) => {
                const first = keys.indexOf(sec);
                return first >= 0 && first < i && lastOf[sec] > i && k !== sec;
            });
            if (!interrupted)
                continue;
            const tag = `<p:sldId id="[^"]*" r:id="${ord[i].rId}"/>`;
            const m = pres.match(new RegExp(tag));
            if (!m)
                continue;
            pres = pres.replace(m[0], '');
            pres = pres.replace(new RegExp(`(<p:sldId[^>]*r:id="${ord[lastOf[interrupted]].rId}"[^>]*/>)`), `$1${m[0]}`);
            moved.push(`${ord[i].part.split('/').pop()} (was inside the ${interrupted} section)`);
        }
        if (moved.length)
            log(`  · moved ${moved.length} slide(s) so each section stays contiguous: ${moved.join(', ')}`);
    }
    // Copied slides colour their bars/fills via THEME references (schemeClr accent2 …). Those resolve
    // against the deck's theme, so unless we adopt the reference's theme the copied slides render in
    // pptxgenjs's default palette (wrong orange). Overwrite the generated theme with the reference's
    // so copied slides show TE's exact colours — and the generated slides already use the same hexes.
    if (copied.length && !opts.skipTheme) {
        const refTheme = await txt(r, 'ppt/theme/theme1.xml');
        if (refTheme && g.file('ppt/theme/theme1.xml'))
            g.file('ppt/theme/theme1.xml', refTheme);
    }
    g.file('ppt/presentation.xml', pres);
    g.file('ppt/_rels/presentation.xml.rels', presRels);
    g.file('[Content_Types].xml', ct);
    // Bare directory entries and orphaned media are cleared afterwards, by normaliseShapeIds in
    // package_validate.ts — it rebuilds the archive from the parts alone, which is the only reliable way
    // (JSZip's createFolders option belongs on file(), and remove() on a folder is recursive).
    const out = (await g.generateAsync({ type: 'nodebuffer' }));
    return { buffer: out, copied };
}
/** Headings of every slide in a deck, in presentation order — used by the diagnostics report. */
async function refSlideTitles(buf) {
    const z = await jszip_1.default.loadAsync(buf);
    const pres = (await z.file('ppt/presentation.xml')?.async('string')) || '';
    const rels = (await z.file('ppt/_rels/presentation.xml.rels')?.async('string')) || '';
    const rm = {};
    for (const m of rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g))
        rm[m[1]] = m[2];
    const out = [];
    for (const m of pres.matchAll(/<p:sldId[^>]*r:id="([^"]+)"/g)) {
        const t = rm[m[1]];
        if (!t || !/slides\/slide/.test(t))
            continue;
        const xml = (await z.file('ppt/' + t.replace(/^\.\.\//, '').replace(/^ppt\//, ''))?.async('string')) || '';
        // The HEADING, so a caller asking "does the reference have a Master Plan slide?" cannot be answered
        // by the Contents slide's agenda or by a table's column header.
        out.push(slideHeading(xml).trim());
    }
    return out;
}
