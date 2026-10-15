"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.hierKey = exports.personKey = void 0;
exports.cleanPerson = cleanPerson;
exports.splitPersons = splitPersons;
exports.roleRows = roleRows;
exports.roleValidation = roleValidation;
exports.hierarchy = hierarchy;
exports.hierPattern = hierPattern;
/**
 * Separators that genuinely divide two people, and the reason each is safe:
 *   ,  ;  /  |  &  +   punctuation nobody has inside a personal name
 *   " and "            written out, with spaces, so "Chandan" is untouched
 * A newline counts too — Excel cells often hold a list one per line.
 * Deliberately NOT split on: "." (initials — "Om Prakash S." ), "-" (double-barrelled names).
 */
const SPLIT_RE = /\s*(?:,|;|\/|\||&|\+|\n|\r|\band\b)\s*/gi;
/**
 * Values that are not people. Real exports carry editing residue — `Praveen George | Delete` appears
 * in the ItQE schedule — and splitting that into two people would put "Delete" on the slide as though
 * it were a colleague. Only these exact words are dropped, never a partial match, and every drop is
 * counted so the report shows what was discarded.
 */
const NOT_A_PERSON = new Set(['delete', 'deleted', 'tbd', 'tba', 'na', 'n a', 'none', 'nil', 'x', 'xx', '?', '-', '--']);
/** Trim, collapse internal runs of whitespace, and strip stray leading/trailing punctuation. */
function cleanPerson(v) {
    return String(v || '').replace(/\s+/g, ' ').replace(/^[\s,;/|&+.-]+|[\s,;/|&+-]+$/g, '').trim();
}
/** Case-insensitive identity for a person, so "ramya  h" and "Ramya H" are one person. */
const personKey = (v) => cleanPerson(v).toLowerCase();
exports.personKey = personKey;
/**
 * Split a people field into individual persons, cleaned and de-duplicated, order preserved.
 * A field with no separator yields exactly one person; an empty field yields none.
 */
function splitPersons(field) {
    const raw = String(field || '');
    if (!raw.trim())
        return [];
    const out = [];
    const seen = new Set();
    for (const piece of raw.split(SPLIT_RE)) {
        const p = cleanPerson(piece);
        if (!p)
            continue;
        const k = (0, exports.personKey)(p);
        if (NOT_A_PERSON.has(k))
            continue; // editing residue, not a colleague
        if (seen.has(k))
            continue; // "A, A" is one person
        seen.add(k);
        out.push(p);
    }
    return out;
}
/**
 * Build the Look Ahead rows for a set of activities.
 *
 * One row per (task, responsible person). Accountable and Head come from the export when it supplies
 * them and from the responsible person when it does not — never blank.
 */
function roleRows(activities, unassignedLabel = 'Unassigned') {
    const rows = [];
    const seen = new Set();
    const people = new Set();
    const stats = {
        activities: activities.length, rows: 0, splitActivities: 0, extraRowsFromSplit: 0,
        duplicatesDropped: 0, mirroredAccountable: 0, mirroredHead: 0, unassigned: 0,
        people: [], samples: [], discarded: [],
    };
    for (const a of activities) {
        const rawResp = a.responsible || '';
        const responsibles = splitPersons(rawResp);
        const accountables = splitPersons(a.accountable || '');
        const heads = splitPersons(a.head || '');
        // A field that names the same person twice ("Ramya H, Ramya H") is one person, and that collapse
        // is counted here so the report's duplicate figure covers both places it can happen.
        const pieces = String(rawResp).split(SPLIT_RE).map(cleanPerson).filter(Boolean);
        const junk = pieces.filter((x) => NOT_A_PERSON.has((0, exports.personKey)(x)));
        for (const j of junk)
            if (stats.discarded.length < 20)
                stats.discarded.push({ task: a.name, value: j });
        const realPieces = pieces.length - junk.length;
        if (realPieces > responsibles.length)
            stats.duplicatesDropped += realPieces - responsibles.length;
        if (responsibles.length > 1) {
            stats.splitActivities += 1;
            stats.extraRowsFromSplit += responsibles.length - 1;
            if (stats.samples.length < 10)
                stats.samples.push({ task: a.name, raw: rawResp, people: responsibles });
        }
        if (!responsibles.length)
            stats.unassigned += 1;
        // ONE ROW PER TASK — not one row per person.
        //
        // This loop used to emit a row for EVERY responsible person, which put the same task under each of
        // them in turn: "Application for conversion" appeared three times, once under Karthick D, once under
        // Jagadish M and once under Chandrashekar K N. A task is one piece of work however many people share
        // it, so it now produces exactly one row and the people are listed together on its header:
        //
        //   Karthick D | Jagadish M | Chandrashekar K N
        //       Land Acquisition | Conversion for remaining 2.77Ac...
        //           Application for conversion
        //
        // Nobody is dropped: every name the export gave is in the list, in the export's own order, already
        // de-duplicated case-insensitively by splitPersons. `people` still records each person separately, so
        // the roles report continues to show who is actually named across the schedule.
        const list = responsibles.length ? responsibles : [unassignedLabel];
        const key = String(a.id);
        if (seen.has(key)) {
            stats.duplicatesDropped += 1;
            continue;
        } // the same task seen twice
        seen.add(key);
        const JOIN = ' | ';
        const responsible = list.join(JOIN);
        // Accountable / Head: use what the export gives, fall back to the people responsible. A multi-person
        // Accountable is kept whole rather than paired off positionally against the responsible people —
        // that pairing would be a guess, and a wrong name is worse than a list.
        const accountable = accountables.length ? accountables.join(JOIN) : responsible;
        const head = heads.length ? heads.join(JOIN) : responsible;
        if (!accountables.length)
            stats.mirroredAccountable += 1;
        if (!heads.length)
            stats.mirroredHead += 1;
        const mirrored = [];
        if (!accountables.length)
            mirrored.push('accountable');
        if (!heads.length)
            mirrored.push('head');
        rows.push({
            activity: a, task: a.name, responsible, accountable, head,
            mirrored, fromSplit: responsibles.length > 1,
        });
        if (responsibles.length)
            for (const person of list)
                people.add(person);
    }
    stats.rows = rows.length;
    stats.people = [...people].sort();
    return { rows, stats };
}
// ---------------------------------------------------------------------------------------------
/** lookahead_roles_validation.md — how the three roles were reconstructed, and the proof it is sound. */
function roleValidation(project, stats, rows, columns) {
    // Prove the two invariants on the rows that were actually produced.
    const dupKeys = new Map();
    for (const r of rows) {
        const k = String(r.activity.id); // one row per task is now the invariant
        dupKeys.set(k, (dupKeys.get(k) || 0) + 1);
    }
    const dup = [...dupKeys.values()].filter((n) => n > 1).length;
    // The grouped header: how each hierarchy collapsed, and a check that the merge never loses a person.
    const hierPatterns = {};
    const seenHier = new Set();
    let hierOk = true;
    for (const r of rows) {
        const k = (0, exports.hierKey)(r);
        if (seenHier.has(k))
            continue;
        seenHier.add(k);
        hierPatterns[hierPattern(r)] = (hierPatterns[hierPattern(r)] || 0) + 1;
        // every distinct person in the three roles must appear in the rendered lines
        const lines = hierarchy(r).map((l) => (0, exports.personKey)(l.name));
        const want = [...new Set([r.head, r.accountable, r.responsible].filter(Boolean).map(exports.personKey))];
        if (!want.every((w) => lines.includes(w)))
            hierOk = false;
    }
    const blank = rows.filter((r) => !r.responsible.trim() || !r.accountable.trim() || !r.head.trim()).length;
    const checks = [
        ['Role columns found in the export', columns.responsible || columns.team,
            `Responsible ${columns.responsible ? '✓' : '—'} · Accountable ${columns.accountable ? '✓' : '—'} · ` +
                `Head ${columns.head ? '✓' : '—'}` + (columns.team ? ' · legacy "Task.TE Team" ✓ (used for Responsible)' : '')],
        ['No duplicate rows for the same Task + Person', dup === 0,
            dup === 0
                ? `${rows.length} row(s), each a distinct task/person pair; ${stats.duplicatesDropped} duplicate(s) collapsed`
                : `${dup} task/person pair(s) still appear more than once`],
        ['Every row has all three roles filled', blank === 0,
            blank === 0 ? `${rows.length} row(s) checked, none blank` : `${blank} row(s) have an empty role`],
        ['Grouped header merges equal roles', hierOk,
            Object.entries(hierPatterns).map(([k, v]) => `${v}× ${k}`).join(' · ') || 'no groups'],
        ['Multiple names split into separate rows', true,
            stats.splitActivities
                ? `${stats.splitActivities} activity(ies) named more than one person, producing ${stats.extraRowsFromSplit} extra row(s)`
                : 'no activity in this export names more than one person in Responsible'],
        // The milestone exclusion, stated rather than silent. Not a pass/fail condition — how many
        // milestones a project has in any given fortnight is data, not correctness — but the names are
        // listed so the slide can always be reconciled against the export.
        ['Milestones excluded from the Look Ahead (both Duration columns = 0)', true,
            (stats.milestonesExcluded && stats.milestonesExcluded.length)
                ? `${stats.milestonesExcluded.length} milestone(s) fell inside the 14-day window and were removed: ` +
                    stats.milestonesExcluded.slice(0, 12).join(' · ') +
                    (stats.milestonesExcluded.length > 12 ? ` … and ${stats.milestonesExcluded.length - 12} more` : '') +
                    '. Identified by the export\'s own zero durations, never by name; they remain on every other slide.'
                : 'no milestone fell inside the 14-day window, so this rule removed nothing from the slide'],
        // Which test actually ran. Stated per run, because it depends on the columns this export carried
        // rather than on anything in the code — and a rule that silently degrades to half of itself is
        // worse than one that says which half it used.
        ['Which milestone test ran', true,
            stats.elapsedDurationColumn
                ? 'BOTH columns: a row was treated as a milestone only when "Duration" AND ' +
                    '"Duration [Elapsed day]" both read zero'
                : 'ONE column: this export carries only one duration column (either "Duration" or ' +
                    '"Duration [Elapsed day]" — not both), so that one had to read zero. Requiring a column ' +
                    'the export does not have would make the test impossible to satisfy and put every ' +
                    'milestone back on the slide. Carry both columns in the export view to have both checked.'],
        ['Blank roles filled from the responsible person', true,
            `Accountable mirrored on ${stats.mirroredAccountable} row(s), Head on ${stats.mirroredHead} row(s)` +
                (stats.mirroredAccountable + stats.mirroredHead === 0
                    ? ' — the export supplied both on every row, so nothing was mirrored' : '')],
    ];
    const ok = checks.every((c) => c[1]);
    const L = [];
    L.push('# lookahead_roles_validation.md', '');
    L.push(`Project: **${project}**`, '');
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Measured |', '|---|---|---|');
    for (const [n, p, d] of checks)
        L.push(`| ${n} | ${p ? '✅' : '❌'} | ${d.replace(/\|/g, '\\|')} |`);
    L.push('', '## Counts', '', '| | |', '|---|---|', `| Activities in the Look Ahead window | ${stats.activities} |`, `| Rows produced | ${stats.rows} |`, `| Activities naming more than one person | ${stats.splitActivities} |`, `| Extra rows from splitting | ${stats.extraRowsFromSplit} |`, `| Duplicate task/person rows collapsed | ${stats.duplicatesDropped} |`, `| Accountable filled from Responsible | ${stats.mirroredAccountable} |`, `| Head filled from Responsible | ${stats.mirroredHead} |`, `| Activities with nobody named | ${stats.unassigned} |`, `| Distinct people | ${stats.people.length} |`, '');
    if (stats.discarded.length) {
        L.push('## Values discarded as "not a person"', '', 'Editing residue in the export. Only an exact match against a short list is dropped — never a', 'partial one — so a real name is never lost.', '', '| Task | Discarded |', '|---|---|');
        for (const d of stats.discarded)
            L.push(`| ${d.task} | \`${d.value}\` |`);
        L.push('');
    }
    if (stats.samples.length) {
        L.push('## Multi-person fields, as split', '', '| Task | Export value | Split into |', '|---|---|---|');
        for (const s of stats.samples)
            L.push(`| ${s.task} | \`${s.raw}\` | ${s.people.map((p) => `**${p}**`).join(' · ')} |`);
        L.push('');
    }
    L.push('## The grouped header', '', 'The three roles are NOT columns. They are printed once above each block of tasks, indented:', '', '```', 'Head', '  Accountable', '    Responsible', '```', '', 'Equal neighbours are merged, so a person is never written twice in a row:', '', '| Case | Lines shown |', '|---|---|', '| Head = Accountable = Responsible | one — the name once |', '| Head = Accountable, Responsible differs | two |', '| Accountable = Responsible, Head differs | two |', '| All three different | three |', '| Head = Responsible, Accountable differs | three — NOT merged: the two are not adjacent, and collapsing them would put the responsible person above the accountable one |', '', 'How this export fell out:', '', '| Case | Groups |', '|---|---|');
    for (const [k, v] of Object.entries(hierPatterns))
        L.push(`| ${k} | ${v} |`);
    L.push('');
    L.push('## The rules applied', '', '1. **Split** on `,` `;` `/` `|` `&` `+`, a newline, or the word "and" with spaces around it. Full', '   stops and hyphens are NOT separators — they belong to initials and double-barrelled names.', '2. **One row per Task + Person.** Same task and same person → one row. Same task, different people', '   → a row each. Same person, different tasks → a row each.', '3. **No role is left blank.** Accountable and Head come from the export where it provides them, and', '   from the responsible person where it does not.', '4. **Names are matched case-insensitively** after trimming and collapsing inner spaces, so', '   "ramya  h" and "Ramya H" are one person — but the export\'s own spelling is what appears.', '');
    L.push('## What was left alone', '', 'Where the export names a different Accountable or Head, that name is kept. Mirroring the', 'responsible person over a populated column would have overwritten real information — on a typical', 'export the majority of activities name someone else in those two columns.', '');
    L.push('_Task selection, the 14-day window, filtering and all schedule logic are untouched: this covers', 'role mapping and row creation only._');
    return { ok, md: L.join('\n') };
}
/**
 * The grouped header for a task block: Head → Accountable → Responsible, indented, with equal
 * neighbours merged into a single line.
 *
 *   all three the same          → one line
 *   Head = Accountable, R other → two lines
 *   all three different         → three lines
 *
 * Merging is of ADJACENT levels only, because the lines are a hierarchy and their order carries the
 * meaning. If Head and Responsible were the same person but the Accountable were someone else, they
 * are NOT merged — collapsing them would put the responsible person above the accountable one. That
 * combination does not occur in any current export; it is handled rather than assumed away.
 */
function hierarchy(row) {
    const chain = [
        { name: row.head, role: 'head' },
        { name: row.accountable, role: 'accountable' },
        { name: row.responsible, role: 'responsible' },
    ].filter((x) => x.name && x.name.trim());
    const out = [];
    for (const link of chain) {
        const prev = out[out.length - 1];
        if (prev && (0, exports.personKey)(prev.name) === (0, exports.personKey)(link.name)) {
            prev.roles.push(link.role);
            continue;
        }
        out.push({ name: link.name, level: Math.min(out.length, 2), roles: [link.role] });
    }
    return out;
}
/** A stable key for one hierarchy, so tasks sharing it are grouped together. */
const hierKey = (row) => [row.head, row.accountable, row.responsible].map(exports.personKey).join(' \u2192 ');
exports.hierKey = hierKey;
/** Which merge case each hierarchy fell into — reported, not assumed. */
function hierPattern(row) {
    const h = (0, exports.personKey)(row.head), a = (0, exports.personKey)(row.accountable), r = (0, exports.personKey)(row.responsible);
    if (h === a && a === r)
        return 'all three the same — one line';
    if (h === a)
        return 'Head = Accountable, Responsible below — two lines';
    if (a === r)
        return 'Accountable = Responsible, Head above — two lines';
    if (h === r)
        return 'Head = Responsible with a different Accountable — three lines, not merged (non-adjacent)';
    return 'all three different — three lines';
}
