"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.approvalNamesFromRef = approvalNamesFromRef;
// approval_names_from_ref.ts — read the APPROVAL LIST off the previous/reference deck.
//
// WHY
// ---
// The reference PPT is the master source for WHICH approvals appear on the slide, and in WHAT ORDER.
// The Excel is the master source for the DATES. Deriving the list from the Excel instead produced a
// slide whose rows changed shape whenever the schedule was re-cut — an approval the business still
// tracks would vanish because its branch had been renamed, and sub-activities would appear as rows.
//
// WHAT THIS READS
// ---------------
// The reference deck's own Approvals Status slide: its table, its "Approvals" column, top to bottom.
// Nothing else is taken from it — not the ticks, not the dates. Those are recomputed from the Excel.
//
// WHAT IT REFUSES TO DO
// --------------------
// It never invents a name and never reorders. If the slide, the table or the column cannot be found it
// returns nothing and says why, and the caller falls back to the schedule-derived list rather than
// silently shipping an empty table.
const jszip_1 = __importDefault(require("jszip"));
const texts = (xml) => [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decode(m[1]));
const decode = (s) => s
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
/** Cell texts of one <a:tr>, in column order, with a gridSpan counted as the columns it covers. */
function rowCells(trXml) {
    const out = [];
    for (const m of trXml.matchAll(/<a:tc\b([^>]*)>([\s\S]*?)<\/a:tc>/g)) {
        const attrs = m[1], body = m[2];
        if (/hMerge="1"|vMerge="1"/.test(attrs)) {
            out.push('');
            continue;
        }
        out.push(texts(body).join('').replace(/\s+/g, ' ').trim());
        const span = /gridSpan="(\d+)"/.exec(attrs);
        if (span)
            for (let k = 1; k < +span[1]; k++)
                out.push('');
    }
    return out;
}
const flat = (s) => s.toLowerCase().replace(/[^a-z0-9]+/g, '');
/**
 * Pull the approval list out of a reference deck.
 *
 * The slide is found by its TITLE text ("Approvals Status") rather than by index, because the slide
 * number moves between projects. Inside it, the column is found by its HEADING ("Approvals"), not by
 * position, so an extra column on either side cannot shift the read.
 */
async function approvalNamesFromRef(ref, 
/**
 * BLOCK-SPLIT DECKS. A project reported in blocks has one Approvals slide per block —
 * "APPROVALS STATUS [ South Block ]" and "APPROVALS STATUS [ North Block ]" — and each block must
 * read ITS OWN list. Pass the block's label and only a slide naming it is accepted, so North can
 * never be given South's approvals. Omit it and the first readable Approvals slide wins, exactly
 * as before.
 */
blockLabel = '') {
    const fail = (note) => ({ names: [], slide: '', note });
    let z;
    try {
        z = await jszip_1.default.loadAsync(ref);
    }
    catch (e) {
        return fail(`the reference deck could not be opened (${e})`);
    }
    const slideNames = Object.keys(z.files)
        .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
        .sort((a, b) => (+(a.match(/\d+/)[0])) - (+(b.match(/\d+/)[0])));
    if (!slideNames.length)
        return fail('the reference deck contains no slides');
    // Every candidate is tried before giving up. The CONTENTS slide lists "7. Approvals Status" as a line
    // of its agenda, so a title-only match found it first and stopped there — it has no table, and the
    // reader reported the list as unreadable while the real slide sat further on. A candidate now has to
    // carry a table with an "Approvals" column to be accepted, and the scan continues until one does.
    const skipped = [];
    for (const name of slideNames) {
        const xml = await z.file(name).async('string');
        const all = texts(xml).join(' ').toUpperCase().replace(/\s+/g, ' ');
        if (!/APPROVAL[S]?\s+STATUS/.test(all))
            continue;
        if (/\bCONTENTS\b/.test(all)) {
            skipped.push(`${name} (the Contents slide)`);
            continue;
        }
        if (blockLabel && !all.replace(/[^A-Z0-9]+/g, '').includes(blockLabel.toUpperCase().replace(/[^A-Z0-9]+/g, ''))) {
            skipped.push(`${name} (does not name the ${blockLabel})`);
            continue;
        }
        const tbl = /<a:tbl>[\s\S]*?<\/a:tbl>/.exec(xml);
        if (!tbl) {
            skipped.push(`${name} (no table)`);
            continue;
        }
        const rows = [...tbl[0].matchAll(/<a:tr\b[^>]*>[\s\S]*?<\/a:tr>/g)].map((m) => rowCells(m[0]));
        if (rows.length < 2) {
            skipped.push(`${name} (table has ${rows.length} row(s))`);
            continue;
        }
        // The heading row is the first row that names the column; the column is wherever that heading sits.
        let hdrIdx = -1, colIdx = -1;
        for (let r = 0; r < Math.min(rows.length, 3) && colIdx < 0; r++) {
            const c = rows[r].findIndex((v) => flat(v) === 'approvals' || flat(v) === 'approval');
            if (c >= 0) {
                hdrIdx = r;
                colIdx = c;
            }
        }
        if (colIdx < 0) {
            skipped.push(`${name} (no "Approvals" column heading)`);
            continue;
        }
        const names = [];
        for (let r = hdrIdx + 1; r < rows.length; r++) {
            const v = (rows[r][colIdx] || '').trim();
            if (!v)
                continue; // a spacer row
            if (flat(v) === 'approvals' || flat(v) === 'approval')
                continue; // a repeated heading
            if (names.some((x) => flat(x) === flat(v)))
                continue; // the same approval twice
            names.push(v);
        }
        if (!names.length) {
            skipped.push(`${name} (the Approvals column is empty)`);
            continue;
        }
        return { names, slide: name, note: `${names.length} approval(s) read from ${name}, in the deck's own order` };
    }
    return fail(`no readable Approvals table${blockLabel ? ` for the ${blockLabel}` : ''} in the reference deck` +
        (skipped.length ? ` — candidates examined and rejected: ${skipped.join('; ')}` : ' (no slide is titled "Approvals Status")'));
}
