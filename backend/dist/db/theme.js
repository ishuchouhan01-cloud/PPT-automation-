"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DB_FS = exports.DB_GRID = exports.LOOKAHEAD_STATUS = exports.LOOKAHEAD_DB = exports.OVERVIEW_BAR = exports.OVERVIEW_LEAF_TEXT = exports.OVERVIEW_LEVEL_TEXT = exports.OVERVIEW_LEAF_FILL = exports.OVERVIEW_LEVEL_FILL = exports.FONT_NUM = exports.FONT_TEXT = exports.FONT_TITLE = void 0;
exports.dbTheme = dbTheme;
const typography_1 = require("../ppt/typography");
Object.defineProperty(exports, "FONT_TITLE", { enumerable: true, get: function () { return typography_1.FONT_TITLE; } });
Object.defineProperty(exports, "FONT_TEXT", { enumerable: true, get: function () { return typography_1.FONT_TEXT; } });
Object.defineProperty(exports, "FONT_NUM", { enumerable: true, get: function () { return typography_1.FONT_NUM; } });
/**
 * PROJECT OVERVIEW HIERARCHY FILLS, by row level (0 = the project row).
 *
 * The two references disagreed here and the difference was resolved deliberately:
 *
 *   Meena         project row #D99370 → "Design" #BFD8D8 → its children #9F9F5F
 *   NH Homestay   project row khaki   → "Construction" #D99370 → its children #BFD8D8 → leaves white
 *
 * Meena's is the same three-colour ramp the Launch & Design Project Overview already uses, so it was
 * chosen as the single Design & Build system. NH's four-tier variant, with its extra khaki, is
 * therefore not reproduced — recorded here so the difference is not mistaken for an oversight.
 *
 * Levels deeper than the ramp render WHITE with dark text, which is what NH's reference does for its
 * individual activity rows.
 */
exports.OVERVIEW_LEVEL_FILL = ['D99370', 'BFD8D8', '9F9F5F'];
exports.OVERVIEW_LEAF_FILL = 'FFFFFF';
/**
 * Row text colour per level. The reference sets the deeper summary levels in a tinted tone and the
 * leaf activity rows in near-black; measured off both decks.
 */
exports.OVERVIEW_LEVEL_TEXT = ['1A1A1A', 'B4531C', '2B2B18'];
exports.OVERVIEW_LEAF_TEXT = '1A1A1A';
/**
 * PROJECT OVERVIEW GANTT — the reference legend reads "Baseline · Projection · Actual Progress",
 * drawn as three stacked bars per row. These are the reference's own colours, sampled from both
 * decks, and they differ from the Launch & Design tracker's greens/blues, which is why they are
 * defined here rather than reused.
 */
exports.OVERVIEW_BAR = {
    baseline: 'FFFF00', // solid yellow, black outline
    projection: '00B050', // bright green
    actual: '9DC3E6', // light blue
    outline: '000000',
    grid: 'BFBFBF',
};
/** 2 WEEK LOOK AHEAD band fills — measured from both reference decks' look-ahead slides. */
exports.LOOKAHEAD_DB = {
    /** One row per person in the hierarchy; the reference stacks them rather than joining with " | ". */
    person: 'F8CBAD',
    /** Summary-activity / package rows. */
    summary: 'BFD8D8',
    bandText: '1A1A1A',
    /** Column header row: solid mid-orange with white bold text. */
    headerFill: 'C55A11',
    headerText: 'FFFFFF',
    grid: 'BFBFBF',
};
/**
 * Status glyphs for the Look Ahead "Status" column.
 *
 * The reference uses a small boxed icon per row — a red cross for not started and an amber warning
 * for in progress. Rendered here as a filled rounded box with a glyph, because the reference's icons
 * are raster images that cannot be regenerated and must not be faked as screenshots.
 */
exports.LOOKAHEAD_STATUS = {
    notStarted: { glyph: '✕', fill: 'C00000', text: 'FFFFFF' },
    inProgress: { glyph: '!', fill: 'E8B33A', text: '3D2C00' },
    completed: { glyph: '✓', fill: '2E741C', text: 'FFFFFF' },
};
/** Grid line used by every Design & Build table. Measured from the reference tables. */
exports.DB_GRID = 'BFBFBF';
/**
 * Type sizes on the 26.67 × 15 in canvas.
 *
 * The canvas is exactly twice a normal 13.3 in slide, so a size here reads at about HALF its number
 * on screen — 28 here is ~14 pt to a reader. The dashboard's own sheets are set in Helvetica Neue at
 * 8–11 pt, which is why these land where they do.
 */
// Several of these were raised when the Project Code and project-name columns came off the Red Flag,
// Tracker, Schedule and Key Focus tables. Those two columns held 5.2"–6.2" of a 24.68" band, and
// handing that width to the remaining columns without also raising the type would have left the
// slides reading at the old size in wider cells — more white space rather than more legible. The
// sizes below are the largest at which every reference project still fits its page; see
// design_build_validation.md, which measures the fit per run rather than assuming it.
exports.DB_FS = {
    /** Band headings on the matrices ("Acquisition", "Design", …). */
    band: 24,
    /** Rotated column headings on the matrices. */
    matrixCol: 21,
    /** The single project row on a matrix. */
    matrixRow: 26,
    /** Project Overview column headings and body. */
    overviewHeader: 22,
    overviewBody: 22,
    /** Timeline month labels (rotated) and year band. */
    month: 21,
    year: 24,
    /** Look Ahead body and column headings. */
    lookahead: 19,
    lookaheadHeader: 21,
    /** Key Focus Area headings and body. */
    keyFocusHeader: 30,
    keyFocus: 28,
    /** Legend text. */
    legend: 20,
};
/**
 * The Design & Build slide chrome is the deck's standard chrome — the same Impact title on the same
 * two accent shades that every EDOS section header uses, which is what both reference decks show.
 * Reusing it is not inheriting a Launch & Design assumption; it is the shared brand header, and the
 * references confirm Design & Build uses it unchanged.
 */
function dbTheme(base) {
    return { ...base, fontBody: typography_1.FONT_TEXT, fontHeader: typography_1.FONT_TITLE, fontTable: typography_1.FONT_NUM };
}
