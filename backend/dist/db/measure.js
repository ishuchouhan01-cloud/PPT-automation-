"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.dbWidth = dbWidth;
exports.dbLines = dbLines;
exports.dbRowH = dbRowH;
exports.dbRowHMax = dbRowHMax;
// measure.ts — text measurement calibrated against what actually renders.
//
// WHY THIS EXISTS
// ---------------
// Row heights are computed before anything is drawn, so they depend on an estimate of how wide a
// string will be. The shared estimator in ppt/typography.ts is deliberately pessimistic, which is
// the right default — under-estimating clips text, and clipped text on a slide is unacceptable.
//
// But at the Look Ahead's size it over-reserves by enough to flip the line count. Measured from a
// generated deck rendered to PDF, reading the real glyph boxes out of the Name column at 19 pt:
//
//     chars   rendered width   implied ratio
//        52        5.889 in        0.429
//        50        5.728 in        0.434
//        47        5.670 in        0.457
//        46        6.135 in        0.505
//        42        5.599 in        0.505
//        42        5.013 in        0.452
//        41        5.473 in        0.506
//                        mean      0.470
//
// The shared estimator assumes ≈0.537 for mixed-case text, so it reserves about 14% more width than
// Helvetica Neue actually needs. Thirteen per cent does not sound like much, but a heading or task
// name sitting near the column boundary gets a second line it never uses — and every such row grows
// by a whole line. On Meena's look-ahead that inflated the section from four pages to eight, each one
// visibly airy.
//
// So this module measures with the ratio that was OBSERVED, and keeps a safety margin on top rather
// than inside the ratio, so the margin is explicit and adjustable instead of buried in a constant.
//
// It is used ONLY by the Design & Build renderers. ppt/typography.ts is untouched, so Launch & Design
// keeps the calibration its eleven shipped projects were verified against.
const typography_1 = require("../ppt/typography");
/**
 * Observed average glyph width as a fraction of font size, per face.
 *
 * `text` is the measured 0.470 for Helvetica Neue. `num` is set a little wider because Meta-Normal's
 * digits are monospaced and wider than its lowercase letters. `title` keeps the shared estimate,
 * since Impact is only ever used for headings that are never wrapped.
 */
const RATIO = {
    [typography_1.FONT_TEXT]: 0.470,
    [typography_1.FONT_NUM]: 0.520,
    [typography_1.FONT_TITLE]: 0.460,
};
/**
 * A margin applied to the measured width, so a string that would land exactly on the boundary still
 * gets a line of its own rather than being trusted to fit to the pixel. 6% was chosen as the widest
 * deviation seen in the samples above (0.506 vs the 0.470 mean is 7.7%, so this is deliberately
 * a little tighter than the worst case and is checked by rendering).
 */
const SAFETY = 1.06;
/** Estimated rendered width of `s`, in inches. Capitals and digits are weighted, as they are wider. */
function dbWidth(s, fs, font = typography_1.FONT_TEXT) {
    const base = RATIO[font] ?? RATIO[typography_1.FONT_TEXT];
    const caps = (String(s).match(/[A-Z0-9]/g) || []).length / Math.max(String(s).length, 1);
    return String(s).length * fs * (base + caps * 0.055) / 72 * SAFETY;
}
/**
 * Lines `s` needs when wrapped into `w` inches at `fs`.
 *
 * Wrapping is on WORDS, because that is what a renderer does — measuring by character count alone
 * would split a long word and under-count. A word longer than the column takes a line to itself.
 *
 * HARD LINE BREAKS ARE COUNTED. The text is split on its own newlines first and each paragraph is
 * wrapped separately, with an empty one still taking a line. Splitting on /\s+/ alone treated a
 * newline as an ordinary space, so a dashboard cell written as two paragraphs with a blank line
 * between them measured as one continuous paragraph and came out one to three lines short — the row
 * was built to the wrong height and the renderer, which does honour the breaks, overflowed it.
 * Meena's "LOD 100: Schematic discussions…" Key Focus cell is that case.
 */
function dbLines(s, w, fs, font = typography_1.FONT_TEXT) {
    const text = String(s ?? '');
    if (!text.trim())
        return 1;
    const usable = Math.max(w, 0.2);
    const wrapOne = (para) => {
        if (!para.trim())
            return 1; // a blank line still occupies one
        let lines = 1, cur = '';
        for (const word of para.split(/\s+/).filter(Boolean)) {
            const trial = cur ? `${cur} ${word}` : word;
            if (dbWidth(trial, fs, font) <= usable) {
                cur = trial;
                continue;
            }
            if (!cur) {
                lines++;
                cur = '';
                continue;
            } // single word wider than the column
            lines++;
            cur = word;
        }
        return lines;
    };
    return text.split(/\r\n|\r|\n|\u000b/).reduce((n, para) => n + wrapOne(para), 0);
}
/**
 * Height a cell needs for `s` at a FIXED font size.
 *
 * The font is never reduced here. On these slides the type size is part of the design — the Look
 * Ahead is specified at 19 pt and the Project Overview at 22 pt — so the row grows and the size
 * stays put. Callers that genuinely cannot grow paginate instead.
 */
function dbRowH(s, w, fs, font = typography_1.FONT_TEXT, pad = 0.10) {
    return dbLines(s, w - 0.14, fs, font) * (0, typography_1.lineHeight)(fs) + pad;
}
/** Tallest requirement across several cells of the same row. */
function dbRowHMax(cells, fs, min, pad = 0.10) {
    return Math.max(min, ...cells.map((c) => dbRowH(c.text, c.w, fs, c.font || typography_1.FONT_TEXT, pad)));
}
