"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.xmlWellFormed = xmlWellFormed;
exports.validatePackage = validatePackage;
exports.packageReport = packageReport;
exports.normaliseShapeIds = normaliseShapeIds;
// package_validate.ts — OOXML package integrity. Runs on EVERY generated deck, and the build fails
// if it finds an error.
//
// WHY THIS EXISTS
// ---------------
// A .pptx is an OPC package: a zip whose parts are tied together by [Content_Types].xml and by
// relationship files. PowerPoint validates those ties before it renders anything, and if one is
// broken it does not try to cope — it shows "PowerPoint found a problem with content" and offers to
// repair. Everything this module checks is a condition that produces that dialog:
//
//   • a part that is not well-formed XML
//   • a relationship whose Target is not in the package        (missing resource)
//   • an r:id used in a slide that its .rels does not define   (broken reference)
//   • a part with no content type
//   • duplicate relationship ids inside one .rels
//   • duplicate slide ids, or ids outside the legal range
//   • a slide part that no sldId points at, or a sldId pointing at nothing
//   • a slide with no layout, a layout with no master, a master with no theme
//   • duplicate <p:cNvPr id> within one slide
//   • an empty media part
//
// It reads the finished buffer — not the objects that produced it — so it cannot be fooled by an
// intention that did not make it into the file.
const jszip_1 = __importDefault(require("jszip"));
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
/** posix path join + normalise, for resolving a relationship Target against its source part. */
function resolve(fromPart, target) {
    if (target.startsWith('/'))
        return target.slice(1);
    const base = fromPart.split('/').slice(0, -1);
    const out = [...base];
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
const relsPathFor = (part) => {
    const i = part.lastIndexOf('/');
    return `${part.slice(0, i)}/_rels/${part.slice(i + 1)}.rels`;
};
function parseRels(xml) {
    const out = [];
    for (const m of xml.matchAll(/<Relationship\b[^>]*\/>|<Relationship\b[^>]*>[\s\S]*?<\/Relationship>/g)) {
        const tag = m[0];
        const id = (tag.match(/\bId="([^"]*)"/) || [])[1] || '';
        const type = (tag.match(/\bType="([^"]*)"/) || [])[1] || '';
        const target = (tag.match(/\bTarget="([^"]*)"/) || [])[1] || '';
        out.push({ id, type, target, external: /TargetMode="External"/.test(tag) });
    }
    return out;
}
/**
 * Well-formedness without an XML parser dependency: a tag-balance walk. It catches exactly the class
 * of damage that string surgery produces — an element left unclosed, or a stray close tag — which is
 * what put an unterminated <a:rPr> into a cover slide and triggered the repair dialog.
 */
function xmlWellFormed(xml) {
    const stack = [];
    const re = /<\/?([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>|<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<!DOCTYPE[^>]*>/g;
    let m;
    while ((m = re.exec(xml))) {
        if (!m[1])
            continue; // declaration / comment / CDATA
        const whole = m[0];
        if (whole.startsWith('</')) {
            const top = stack.pop();
            if (!top)
                return { ok: false, detail: `unexpected closing tag </${m[1]}> at offset ${m.index}` };
            if (top.name !== m[1]) {
                return { ok: false, detail: `mismatched tag: <${top.name}> opened at offset ${top.at} but </${m[1]}> found at ${m.index}` };
            }
        }
        else if (!m[3]) {
            stack.push({ name: m[1], at: m.index });
        }
    }
    if (stack.length) {
        const t = stack[stack.length - 1];
        return { ok: false, detail: `unclosed <${t.name}> opened at offset ${t.at}` };
    }
    return { ok: true, detail: 'balanced' };
}
async function validatePackage(buf) {
    const z = await jszip_1.default.loadAsync(buf);
    const issues = [];
    const err = (area, detail) => issues.push({ severity: 'error', area, detail });
    const warn = (area, detail) => issues.push({ severity: 'warning', area, detail });
    const names = Object.keys(z.files).filter((n) => !z.files[n].dir);
    const dirs = Object.keys(z.files).filter((n) => z.files[n].dir);
    const read = async (n) => (z.file(n) ? await z.file(n).async('string') : null);
    // ---------- 1. zip ----------
    const dup = names.filter((n, i) => names.indexOf(n) !== i);
    if (dup.length)
        err('ZIP', `duplicate entries: ${[...new Set(dup)].join(', ')}`);
    if (dirs.length)
        warn('ZIP', `${dirs.length} bare directory entr(y/ies): ${dirs.join(', ')} — an OPC package describes parts, not folders`);
    // ---------- 2. XML well-formedness ----------
    const xmlOf = new Map();
    let malformed = 0;
    for (const n of names) {
        if (!/\.(xml|rels)$/i.test(n))
            continue;
        const x = (await read(n)) || '';
        xmlOf.set(n, x);
        const w = xmlWellFormed(x);
        if (!w.ok) {
            err('XML', `${n} — ${w.detail}`);
            malformed++;
        }
    }
    // ---------- 3. content types ----------
    const ctXml = xmlOf.get('[Content_Types].xml') || '';
    if (!ctXml)
        err('ContentTypes', '[Content_Types].xml is missing');
    const defaults = new Set([...ctXml.matchAll(/<Default\b[^>]*Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase()));
    const overrides = new Set([...ctXml.matchAll(/<Override\b[^>]*PartName="([^"]+)"/g)].map((m) => m[1].replace(/^\//, '')));
    const dupDefault = [...ctXml.matchAll(/<Default\b[^>]*Extension="([^"]+)"/g)].map((m) => m[1].toLowerCase())
        .filter((x, i, a) => a.indexOf(x) !== i);
    if (dupDefault.length)
        err('ContentTypes', `duplicate Default extensions: ${[...new Set(dupDefault)].join(', ')}`);
    const dupOverride = [...overrides].filter((x, i, a) => a.indexOf(x) !== i);
    if (dupOverride.length)
        err('ContentTypes', `duplicate Override part names: ${dupOverride.join(', ')}`);
    for (const n of names) {
        if (n === '[Content_Types].xml')
            continue;
        if (overrides.has(n))
            continue;
        const ext = n.includes('.') ? n.split('.').pop().toLowerCase() : '';
        if (!defaults.has(ext))
            err('ContentTypes', `no content type declared for ${n}`);
    }
    // ---------- 4. relationships ----------
    const relsOf = new Map();
    for (const n of names) {
        if (!n.endsWith('.rels'))
            continue;
        const rels = parseRels(xmlOf.get(n) || '');
        relsOf.set(n, rels);
        const ids = rels.map((r) => r.id);
        const d = ids.filter((x, i) => ids.indexOf(x) !== i);
        if (d.length)
            err('Relationships', `${n} declares ${[...new Set(d)].join(', ')} more than once`);
        // The part a .rels belongs to: "ppt/slides/_rels/slide1.xml.rels" → "ppt/slides/slide1.xml".
        // The package-level "_rels/.rels" belongs to the package root, so its targets resolve from "".
        const srcPart = n === '_rels/.rels' ? 'x' : n.replace('/_rels/', '/').replace(/\.rels$/, '');
        for (const r of rels) {
            if (r.external)
                continue;
            if (!r.target) {
                err('Relationships', `${n}: ${r.id} has no Target`);
                continue;
            }
            const t = resolve(srcPart, r.target);
            if (!names.includes(t))
                err('Relationships', `${n}: ${r.id} → "${r.target}" points at ${t}, which is not in the package`);
        }
    }
    const relsFor = (part) => relsOf.get(relsPathFor(part)) || [];
    // ---------- 5. presentation.xml ----------
    const pres = xmlOf.get('ppt/presentation.xml') || '';
    if (!pres)
        err('Presentation', 'ppt/presentation.xml is missing');
    const presRels = relsFor('ppt/presentation.xml');
    const presRelById = new Map(presRels.map((r) => [r.id, r]));
    const sldIds = [];
    for (const m of pres.matchAll(/<p:sldId\b[^>]*\/>/g)) {
        const id = (m[0].match(/\bid="(\d+)"/) || [])[1] || '';
        const rId = (m[0].match(/r:id="([^"]+)"/) || [])[1] || '';
        sldIds.push({ id, rId });
    }
    const slideParts = [];
    for (const s of sldIds) {
        const n = Number(s.id);
        if (!Number.isFinite(n) || n < 256 || n > 2147483647)
            err('Presentation', `slide id ${s.id} is outside the legal 256–2147483647 range`);
        const r = presRelById.get(s.rId);
        if (!r) {
            err('Presentation', `sldId ${s.id} references ${s.rId}, which presentation.xml.rels does not define`);
            continue;
        }
        const part = resolve('ppt/presentation.xml', r.target);
        slideParts.push(part);
        if (!names.includes(part))
            err('Presentation', `sldId ${s.id} → ${part} is not in the package`);
    }
    const dupSlideId = sldIds.map((s) => s.id).filter((x, i, a) => a.indexOf(x) !== i);
    if (dupSlideId.length)
        err('Presentation', `duplicate slide ids: ${[...new Set(dupSlideId)].join(', ')}`);
    const dupSldRid = sldIds.map((s) => s.rId).filter((x, i, a) => a.indexOf(x) !== i);
    if (dupSldRid.length)
        err('Presentation', `duplicate sldId r:ids: ${[...new Set(dupSldRid)].join(', ')}`);
    const pkgSlides = names.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n));
    const orphanSlides = pkgSlides.filter((p) => !slideParts.includes(p));
    if (orphanSlides.length)
        err('Presentation', `slide part(s) in the package that no sldId points at: ${orphanSlides.join(', ')}`);
    const masterIds = [...pres.matchAll(/<p:sldMasterId\b[^>]*\/>/g)].map((m) => (m[0].match(/\bid="(\d+)"/) || [])[1]);
    const dupMaster = masterIds.filter((x, i, a) => a.indexOf(x) !== i);
    if (dupMaster.length)
        err('Presentation', `duplicate slide master ids: ${dupMaster.join(', ')}`);
    for (const m of pres.matchAll(/<p:sldMasterId\b[^>]*r:id="([^"]+)"[^>]*\/>/g)) {
        if (!presRelById.has(m[1]))
            err('Presentation', `sldMasterId references ${m[1]}, which is not in presentation.xml.rels`);
    }
    // ---------- 6. every slide / layout / master / notes ----------
    let dupShapeIds = 0;
    const parts = names.filter((n) => /^ppt\/(slides|slideLayouts|slideMasters|notesSlides|notesMasters)\/[^/]+\.xml$/.test(n));
    for (const part of parts) {
        const xml = xmlOf.get(part) || '';
        const rels = relsFor(part);
        const byId = new Set(rels.map((r) => r.id));
        // every r:id the part uses must be declared by its own .rels
        for (const m of xml.matchAll(/r:(?:id|embed|link|pict|dm|lo|qs|cs)="([^"]+)"/g)) {
            if (!byId.has(m[1]))
                err('Relationships', `${part} uses ${m[1]} but its .rels does not define it`);
        }
        if (part.startsWith('ppt/slides/')) {
            if (!rels.some((r) => r.type.endsWith('/slideLayout')))
                err('Chain', `${part} has no slideLayout relationship`);
            const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
            const d = ids.filter((x, i) => ids.indexOf(x) !== i);
            if (d.length) {
                warn('ShapeIds', `${part} reuses shape id(s) ${[...new Set(d)].join(', ')} — ids must be unique within a slide`);
                dupShapeIds++;
            }
        }
        if (part.startsWith('ppt/slideLayouts/') && !rels.some((r) => r.type.endsWith('/slideMaster'))) {
            err('Chain', `${part} has no slideMaster relationship`);
        }
        if (part.startsWith('ppt/slideMasters/')) {
            if (!rels.some((r) => r.type.endsWith('/theme')))
                err('Chain', `${part} has no theme relationship`);
            const layoutIds = [...xml.matchAll(/<p:sldLayoutId\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
            const d = layoutIds.filter((x, i) => layoutIds.indexOf(x) !== i);
            if (d.length)
                err('Chain', `${part} has duplicate layout ids ${[...new Set(d)].join(', ')}`);
        }
    }
    // ---------- 7. media ----------
    const media = names.filter((n) => n.startsWith('ppt/media/'));
    const referenced = new Set();
    for (const [relPath, rels] of relsOf) {
        const srcPart = relPath === '_rels/.rels' ? 'x' : relPath.replace('/_rels/', '/').replace(/\.rels$/, '');
        for (const r of rels) {
            if (r.external || !r.target)
                continue;
            const t = resolve(srcPart, r.target);
            if (t.startsWith('ppt/media/'))
                referenced.add(t);
        }
    }
    for (const m of media) {
        const size = (z.file(m)?._data?.uncompressedSize) ?? (await z.file(m).async('nodebuffer')).length;
        if (!size)
            err('Media', `${m} is zero bytes`);
        if (!referenced.has(m))
            warn('Media', `${m} is in the package but nothing references it`);
    }
    const errors = issues.filter((i) => i.severity === 'error');
    const checks = [
        { name: 'ZIP structure', ok: !issues.some((i) => i.area === 'ZIP' && i.severity === 'error'), detail: `${names.length} parts, no duplicate entries` },
        { name: 'Every XML part is well-formed', ok: malformed === 0, detail: malformed ? `${malformed} malformed part(s)` : `${xmlOf.size} XML/rels parts parsed and tag-balanced` },
        { name: 'Every part has a content type', ok: !issues.some((i) => i.area === 'ContentTypes' && i.severity === 'error'), detail: `${defaults.size} Default extension(s), ${overrides.size} Override(s)` },
        { name: 'Every relationship resolves', ok: !issues.some((i) => i.area === 'Relationships' && i.severity === 'error'), detail: `${[...relsOf.values()].reduce((a, b) => a + b.length, 0)} relationship(s) across ${relsOf.size} .rels file(s)` },
        { name: 'Slide ids and references are sound', ok: !issues.some((i) => i.area === 'Presentation'), detail: `${sldIds.length} sldId(s), ${pkgSlides.length} slide part(s), all matched` },
        { name: 'Slide → layout → master → theme chain intact', ok: !issues.some((i) => i.area === 'Chain'), detail: `${parts.length} slide/layout/master part(s) checked` },
        { name: 'Media present and referenced', ok: !issues.some((i) => i.area === 'Media' && i.severity === 'error'), detail: `${media.length} media part(s), ${referenced.size} referenced` },
        { name: 'Shape ids unique within each slide', ok: dupShapeIds === 0, detail: dupShapeIds ? `${dupShapeIds} slide(s) reuse a shape id` : 'no slide reuses a shape id' },
    ];
    return { ok: errors.length === 0, parts: names.length, slides: pkgSlides.length, media: media.length, issues, checks };
}
function packageReport(r, project) {
    const L = [];
    const errors = r.issues.filter((i) => i.severity === 'error');
    const warns = r.issues.filter((i) => i.severity === 'warning');
    L.push('# validation_report.md', '');
    L.push(`Project: **${project}**`, '');
    L.push('OOXML package integrity, read from the finished .pptx. Every condition below is one that makes', 'PowerPoint show *"PowerPoint found a problem with content"* and offer to repair the file.', '');
    L.push(`## Result: ${r.ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push(`${r.parts} parts · ${r.slides} slides · ${r.media} media`, '');
    L.push('| Check | Result | Measured |', '|---|---|---|');
    for (const c of r.checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    L.push('', '## Errors', '');
    if (errors.length)
        for (const e of errors)
            L.push(`- ❌ **${e.area}** — ${e.detail}`);
    else
        L.push('None. The package is internally consistent.');
    L.push('', '## Warnings', '');
    if (warns.length)
        for (const w of warns.slice(0, 40))
            L.push(`- ⚠️ **${w.area}** — ${w.detail}`);
    else
        L.push('None.');
    if (warns.length > 40)
        L.push('', `_…and ${warns.length - 40} more._`);
    L.push('', '## What is checked', '', '| Area | Failure it prevents |', '|---|---|', '| ZIP | a corrupt archive, duplicate entries, bare directory entries |', '| XML | a part that is not well-formed — the classic result of editing OOXML as a string |', '| Content types | a part PowerPoint cannot type, so it refuses to load it |', '| Relationships | a Target that is not in the package, an r:id a part uses but does not declare, duplicate ids |', '| Presentation | duplicate or out-of-range slide ids, a sldId pointing nowhere, a slide part nothing points at |', '| Chain | slide without a layout, layout without a master, master without a theme |', '| Media | zero-byte images, orphaned media |', '| Shape ids | a slide that reuses a <p:cNvPr id> |', '');
    L.push('_This runs on every generated deck. An error fails the build; the deck is not delivered._', '');
    return L.join('\n');
}
// ---------------------------------------------------------------------------------------------
/**
 * Package hygiene, applied to the finished buffer before it is validated. Three defects, each of which
 * the validator would otherwise report:
 *
 *   1. DUPLICATE SHAPE IDS. pptxgenjs numbers shapes and tables from separate counters, so a slide
 *      carrying both ends up with two objects sharing an id (Key Focus: "Text 3" and "Table 0" were
 *      both id 5). ECMA-376 requires the id to be unique within the part. Only the attribute value is
 *      rewritten — no element is added, removed or reordered — so the slide renders identically.
 *   2. BARE DIRECTORY ENTRIES. An OPC package describes parts, not folders. `createFolders: false`
 *      stops JSZip adding new ones; entries already in the object have to be removed explicitly.
 *   3. ORPHANED MEDIA. When a placeholder slide is replaced, the images only it used are left behind,
 *      referenced by nothing. Every .rels in the package is scanned first, so a picture still used by
 *      any slide, layout, master, notes slide or theme is kept.
 *
 * There is no pptxgenjs API for any of this, which is why it happens here — and it is the smallest
 * intervention that fixes each one.
 */
async function normaliseShapeIds(buf) {
    const z = await jszip_1.default.loadAsync(buf);
    const touched = [];
    let renumbered = 0;
    // ---- 1. shape ids ----
    const rewritten = new Map();
    for (const name of Object.keys(z.files)) {
        if (z.files[name].dir)
            continue;
        if (!/^ppt\/(slides|slideLayouts|slideMasters|notesSlides)\/[^/]+\.xml$/.test(name))
            continue;
        const xml = await z.file(name).async('string');
        const used = new Set();
        let next = 0;
        for (const m of xml.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g))
            next = Math.max(next, Number(m[1]));
        let changed = 0;
        const out = xml.replace(/(<p:cNvPr\b[^>]*\bid=")(\d+)(")/g, (whole, pre, id, post) => {
            const n = Number(id);
            if (!used.has(n)) {
                used.add(n);
                return whole;
            }
            next += 1;
            used.add(next);
            changed += 1;
            return `${pre}${next}${post}`;
        });
        if (changed) {
            rewritten.set(name, out);
            touched.push(name);
            renumbered += changed;
        }
    }
    // ---- 2. what is referenced, so orphaned media can be identified ----
    const referenced = new Set();
    for (const name of Object.keys(z.files)) {
        if (z.files[name].dir || !name.endsWith('.rels'))
            continue;
        const srcPart = name === '_rels/.rels' ? 'x' : name.replace('/_rels/', '/').replace(/\.rels$/, '');
        for (const r of parseRels(await z.file(name).async('string'))) {
            if (r.external || !r.target)
                continue;
            referenced.add(resolve(srcPart, r.target));
        }
    }
    const folders = Object.keys(z.files).filter((n) => z.files[n].dir);
    const files = Object.keys(z.files).filter((n) => !z.files[n].dir);
    const mediaDropped = files.filter((n) => n.startsWith('ppt/media/') && !referenced.has(n));
    const drop = new Set(mediaDropped);
    // ---- 3. rebuild the archive from the parts we are keeping ----
    // A fresh JSZip is used deliberately. `z.remove(folder)` would have been the obvious way to drop the
    // bare directory entries, but JSZip removes a folder RECURSIVELY — it took the whole package with it.
    // Copying the parts forward cannot do that.
    //
    // `createFolders` belongs on file(), NOT on generateAsync() — passing it to generateAsync is silently
    // ignored, which is why an earlier attempt still emitted "ppt/media/". With it set per part, JSZip
    // never creates the parent folder entries in the first place.
    const out = new jszip_1.default();
    for (const name of files) {
        if (drop.has(name))
            continue;
        if (name === '[Content_Types].xml' && mediaDropped.length) {
            let ct = await z.file(name).async('string');
            for (const m of mediaDropped) {
                ct = ct.replace(new RegExp(`<Override PartName="/${m.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"[^>]*/>`), '');
            }
            out.file(name, ct, { createFolders: false });
            continue;
        }
        const r = rewritten.get(name);
        out.file(name, r !== undefined ? r : await z.file(name).async('nodebuffer'), { createFolders: false });
    }
    const buffer = (await out.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    return { buffer, renumbered, parts: touched, foldersDropped: folders.length, mediaDropped };
}
