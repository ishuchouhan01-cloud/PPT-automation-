"use strict";
// Robust project-name matching shared by the dashboard, approvals and config lookups.
//
// ASTA exports, the Summary Dashboard, the Approval workbook and the reference decks all spell the
// same project slightly differently: "ItQE Ph 2c & 3a [...]" vs "ITQE Phase 2c & 3a", "OtR Phase 2"
// vs "OTR - Ph2", "Shreegoan Acquisition And Launch" vs "ShreeGaon". A plain substring test misses
// all of these, so we match on shared "core tokens" with a prefix/near fallback for spelling drift.
Object.defineProperty(exports, "__esModule", { value: true });
exports.cleanName = cleanName;
exports.coreTokens = coreTokens;
exports.matchScore = matchScore;
exports.projectMatches = projectMatches;
exports.normName = normName;
exports.codesEqual = codesEqual;
exports.sameAliasGroup = sameAliasGroup;
exports.strictProjectMatch = strictProjectMatch;
// Alias groups: names that mean the SAME project across the export / dashboard / approvals.
// Loaded once from project_aliases.json (see paths.loadAliases). A name belongs to a group if it
// shares a core token with any member (so "Shreegoan Acquisition and Launch" still resolves).
let ALIAS_GROUPS = null;
function aliasGroups() {
    if (ALIAS_GROUPS)
        return ALIAS_GROUPS;
    try {
        ALIAS_GROUPS = require('../paths').loadAliases();
    }
    catch {
        ALIAS_GROUPS = [];
    }
    return ALIAS_GROUPS;
}
// Which alias group (index) does a name fall into? -1 if none.
function aliasIndex(name) {
    const tokens = coreTokens(name);
    if (!tokens.length)
        return -1;
    const groups = aliasGroups();
    for (let i = 0; i < groups.length; i++) {
        for (const member of groups[i]) {
            const mt = coreTokens(member);
            if (mt.some((t) => tokens.includes(t)))
                return i;
            // spelling drift within an alias member (Shreegoan/Shreegaon)
            if (mt[0] && tokens[0] && mt[0].length >= 6 && tokens[0].length >= 6 && mt[0].slice(0, 6) === tokens[0].slice(0, 6))
                return i;
        }
    }
    return -1;
}
const STOP = new Set([
    'phase', 'ph', 'ph1', 'ph2', 'ph3', 'acquisition', 'launch', 'residential', 'block', 'south',
    'north', 'east', 'west', 'villas', 'villa', 'and', 'the', 'project', 'plotted', 'development',
    'acres', 'acre', 'tower', 'towers', 'homes', 'in', 'of', 'for', 'new', 'projects', 'nitco',
]);
/** Light display cleanup: drop the generic programme suffixes but keep meaningful brackets. */
function cleanName(raw) {
    return String(raw || '')
        .replace(/\s*[-–]\s*residential\s*$/i, '')
        .replace(/\s+acquisition\s+and\s+launch\s*$/i, '')
        .trim() || 'Project';
}
/** Distinctive alphabetic tokens (>=3 chars, stopwords removed). Brackets/numbers ignored. */
function coreTokens(name) {
    return String(name || '')
        .toLowerCase()
        .replace(/\[[^\]]*\]/g, ' ') // drop bracketed qualifiers
        .split(/[^a-z]+/) // alpha tokens only
        .filter((t) => t.length >= 3 && !STOP.has(t));
}
/** 0 = no match; higher = stronger. Shared full token beats a prefix/near match. */
function matchScore(a, b) {
    // explicit alias groups win outright (e.g. Shreegoan == Poinciana)
    const ai = aliasIndex(a);
    if (ai >= 0 && ai === aliasIndex(b))
        return 100;
    const ta = coreTokens(a), tb = coreTokens(b);
    if (!ta.length || !tb.length)
        return 0;
    const shared = ta.filter((t) => tb.includes(t));
    if (shared.length)
        return 2 * shared.length + (ta.length === tb.length ? 1 : 0);
    // spelling drift (Shreegoan/Shreegaon): first 6 chars of the primary tokens agree
    const pa = ta[0], pb = tb[0];
    if (pa.length >= 6 && pb.length >= 6 && pa.slice(0, 6) === pb.slice(0, 6))
        return 1;
    return 0;
}
function projectMatches(a, b) {
    return matchScore(a, b) > 0;
}
// ============================================================================
// STRICT MATCHING — used where a wrong match is unacceptable (Key Focus Area)
// ============================================================================
//
// matchScore() above is deliberately forgiving, because it has to reconcile the same project spelled
// four different ways across four different files. That tolerance is wrong for Key Focus Area, where
// a false positive silently prints another project's commitments on this project's slide.
//
// The Summary Dashboard currently carries THREE projects called Nandi Hills:
//     0100  Nandi Hills [MH]
//     0130  Nandi Hills [Sult]
//     0107  Nandi Hills [Ach]
// coreTokens() strips bracketed qualifiers, so all three reduce to ["nandi","hills"] and score 5
// against each other — a strong "match". Two Violet Hill phases share code 0118 the same way.
//
// The strict rules below are therefore: an EXACT name, or an EXACTLY listed alias, and nothing else.
// No shared tokens, no prefixes, no "contains", no "starts with", no spelling-drift allowance.
/** Punctuation-and-case-insensitive identity. "Nandi Hills [MH]" == "nandi hills (MH)" == "NandiHillsMH".
 *  It does NOT make two different names equal — only two spellings of the same name. */
function normName(s) {
    return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}
/** Project codes compared exactly, tolerating leading zeros only ("0100" == "100").
 *  The dashboard writes some codes padded ("0113") and some not ("106", "117"). */
function codesEqual(a, b) {
    const x = normName(a), y = normName(b);
    if (!x || !y)
        return false;
    if (x === y)
        return true;
    if (/^\d+$/.test(x) && /^\d+$/.test(y))
        return Number(x) === Number(y);
    return false;
}
/** True when both names appear in the SAME group of project_aliases.json, compared EXACTLY.
 *  Unlike aliasIndex() above, a name does not join a group by sharing a token with a member —
 *  it has to BE a member. That is what stops "Nandi Hills [Sult]" joining a Nandi Hills group. */
let STRICT_GROUPS = null;
function strictAliasGroups() {
    if (STRICT_GROUPS)
        return STRICT_GROUPS;
    try {
        STRICT_GROUPS = require('../paths').loadStrictAliases();
    }
    catch {
        STRICT_GROUPS = [];
    }
    return STRICT_GROUPS;
}
function sameAliasGroup(a, b) {
    const x = normName(a), y = normName(b);
    if (!x || !y)
        return false;
    for (const group of strictAliasGroups()) {
        const members = group.map(normName);
        if (members.includes(x) && members.includes(y))
            return true;
    }
    return false;
}
/**
 * The strict project test, in the required order of precedence:
 *   1. If BOTH sides carry a project code, THE CODE DECIDES — equal codes accept, different codes
 *      reject. The name is not consulted at all, in either direction. (Previously a fuzzy name hit
 *      could admit a row whose code plainly disagreed.)
 *   2. Otherwise an EXACT project name (ignoring case and punctuation only).
 *   3. Otherwise an EXACT alias listed in project_aliases.json for this project.
 *   4. Otherwise REJECT, with the reason.
 */
function strictProjectMatch(selName, selCode, rowName, rowCode) {
    if (selCode && rowCode) {
        return codesEqual(selCode, rowCode)
            ? { ok: true, by: 'code' }
            : { ok: false, why: `project code "${rowCode}" is not "${selCode}"` };
    }
    if (normName(selName) && normName(selName) === normName(rowName))
        return { ok: true, by: 'name' };
    if (sameAliasGroup(selName, rowName))
        return { ok: true, by: 'alias' };
    return {
        ok: false,
        why: rowCode && !selCode
            ? `name "${rowName}" is not "${selName}" and is not a listed alias (no project code was given, ` +
                `so the code "${rowCode}" on the row could not be used to decide)`
            : `name "${rowName}" is neither "${selName}" nor a listed alias of it`,
    };
}
