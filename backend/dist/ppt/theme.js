"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.TABLE_FS = exports.LOOKAHEAD = exports.STATUS = exports.DEFAULT_THEME = void 0;
exports.lumMod = lumMod;
exports.extractTheme = extractTheme;
// Apply an OOXML lumMod (+ optional lumOff) to a hex colour — the same transform PowerPoint uses,
// so a shade computed here matches a copied slide's `<a:schemeClr><a:lumMod/></a:schemeClr>`.
function lumMod(hex, mod, off = 0) {
    const r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2, d = max - min;
    let h = 0;
    const s = d === 0 ? 0 : d / (1 - Math.abs(2 * l - 1));
    if (d !== 0) {
        if (max === r)
            h = ((g - b) / d) % 6;
        else if (max === g)
            h = (b - r) / d + 2;
        else
            h = (r - g) / d + 4;
        h *= 60;
        if (h < 0)
            h += 360;
    }
    const L = Math.max(0, Math.min(1, l * mod + off));
    const c = (1 - Math.abs(2 * L - 1)) * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = L - c / 2;
    let rp = 0, gp = 0, bp = 0;
    if (h < 60)
        [rp, gp, bp] = [c, x, 0];
    else if (h < 120)
        [rp, gp, bp] = [x, c, 0];
    else if (h < 180)
        [rp, gp, bp] = [0, c, x];
    else if (h < 240)
        [rp, gp, bp] = [0, x, c];
    else if (h < 300)
        [rp, gp, bp] = [x, 0, c];
    else
        [rp, gp, bp] = [c, 0, x];
    const to = (v) => Math.round((v + m) * 255).toString(16).padStart(2, '0').toUpperCase();
    return to(rp) + to(gp) + to(bp);
}
// Total Environment default palette (accent2 = E97132), matched to the reference decks.
// The template renders header BARS at accent2·lumMod75 (C04F15) and the small icon box at
// accent2·lumMod50 (80350E) — copied slides use those exact shades, so the generated slides do too.
exports.DEFAULT_THEME = {
    accent: 'E97132',
    accentBar: 'C04F15', // accent2 · lumMod 75%  — section/table header bars
    accentIcon: '80350E', // accent2 · lumMod 50%  — the numbered icon box
    headerText: 'FFFFFF',
    dark: '222222',
    ragGood: 'A9D08E',
    ragAmber: 'FFAA33',
    ragRed: 'EE4B2B',
    fontBody: 'Arial',
    fontHeader: 'Impact',
    fontTable: 'Meta-Normal', // TE brand body face used on the data slides
};
/**
 * STATUS COLOUR TOKENS — one definition, used everywhere a status is shown.
 *
 * The Red Flag Summary is the origin of these two: its amber and red cells are `ragAmber` / `ragRed`.
 * Any other place that means "At Risk" or "Critical" — the Key Focus section bands, any bar — reads the
 * same token, so the two can never drift into similar-but-different shades. Nothing hardcodes a second
 * amber or a second red.
 */
exports.STATUS = {
    atRisk: (t) => t.ragAmber, // Red Flag Summary amber — #FFAA33
    critical: (t) => t.ragRed, // Red Flag Summary red   — #EE4B2B
    good: (t) => t.ragGood, // Red Flag Summary green  — #A9D08E
};
/**
 * 2 WEEK LOOK AHEAD PALETTE — one definition for the slide's three levels of hierarchy.
 *
 * The slide reads top-down as title → person → summary activity → task, and each level has exactly one
 * fill. They live here, next to STATUS, so nothing in slides.ts hardcodes a second shade of any of them
 * and a future change is one edit. These are the only fills the Look Ahead sets; task rows, borders,
 * status indicators and Notes are untouched by them.
 *
 * The two band text colours travel with the fills: both new backgrounds are LIGHT, so band text is set
 * near-black. Person bands were previously white-on-dark; white on #F7C7AC would have been unreadable,
 * and an unreadable band is not what "apply this background" can mean.
 */
exports.LOOKAHEAD = {
    // NOTE: there is no titleBar entry. The 2 Week Look Ahead uses the deck's standard section-header
    // colours, like every other section — it briefly had its own lighter bar (#C65911) and that was
    // reverted on request. Only the two band fills below belong to this slide.
    person: 'F7C7AC', // Responsible / Accountable / Head name rows
    summary: 'BFD8D8', // summary activity / subheading rows
    bandText: '1A1A1A', // text on either band — both fills are light
};
/** Standard data-table type size on the TE 26.67" canvas (≈14pt on a normal 13.3" slide). */
exports.TABLE_FS = 28;
// Best-effort theme extraction from a reference .pptx (theme1.xml accent2).
// Falls back to DEFAULT_THEME. Kept dependency-free; if a zip reader is present it is used.
async function extractTheme(referenceBuf) {
    if (!referenceBuf)
        return { ...exports.DEFAULT_THEME };
    try {
        // ExcelJS bundles jszip; reuse it to peek at the pptx theme without a new dependency.
        const JSZip = require('jszip');
        const zip = await JSZip.loadAsync(referenceBuf);
        const themeFile = Object.keys(zip.files).find((f) => /ppt\/theme\/theme1\.xml$/.test(f));
        if (themeFile) {
            const xml = await zip.files[themeFile].async('string');
            const m = xml.match(/<a:accent2>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"/);
            if (m) {
                const accent = m[1].toUpperCase();
                return { ...exports.DEFAULT_THEME, accent, accentBar: lumMod(accent, 0.75), accentIcon: lumMod(accent, 0.5) };
            }
        }
    }
    catch {
        /* fall back to defaults */
    }
    return { ...exports.DEFAULT_THEME };
}
