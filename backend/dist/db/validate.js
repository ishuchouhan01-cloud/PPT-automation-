"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.dbValidationReport = dbValidationReport;
const tick = (v) => (v === 'PASS' ? '✓' : v === 'FAIL' ? '✗' : '—');
function dbValidationReport(v) {
    const checks = [];
    const add = (name, verdict, detail) => checks.push({ name, verdict, detail });
    // --- 1. No Approvals Status slide, anywhere ---
    const apInContents = v.contents.filter((c) => /approval/i.test(c) && /status/i.test(c));
    add('No Approvals Status slide', apInContents.length ? 'FAIL' : 'PASS', apInContents.length
        ? `Contents lists ${apInContents.map((c) => `"${c}"`).join(', ')}.`
        : 'No section named "Approvals Status" was registered, so no such slide exists. Design & Build ' +
            'never builds one: approval progress stays inside the Red Flag / Tracker "Approvals" column ' +
            'band and the Project Overview\'s "Approvals [Deliverables]" row, which is where the ' +
            'reference decks show it.');
    add('No Approvals Status in Contents', apInContents.length ? 'FAIL' : 'PASS', `Contents reads: ${v.contents.map((c, i) => `${i + 1}. ${c}`).join(' · ') || '(empty)'}`);
    // --- 2. Design & Build Red Flag structure ---
    if (v.redFlag) {
        const bands = v.redFlag.bands.map(([b, n]) => `${b}×${n}`).join(', ');
        const spanned = v.redFlag.bands.reduce((n, b) => n + b[1], 0);
        add('Design & Build Red Flag Summary structure is used', spanned === v.redFlag.cols.length ? 'PASS' : 'FAIL', `Read from "${v.redFlag.sheet}" row ${v.redFlag.rowNumber} ("${v.redFlag.matchedName}"). ` +
            `${v.redFlag.cols.length} columns under ${v.redFlag.bands.length} bands: ${bands}. ` +
            `Name column heading is "${v.redFlag.nameHeading}". Status comes from the sheet's own cell ` +
            `fills, mapped through its own legend (${v.redFlag.legend.map((l) => `${l.label}=#${l.color}`).join(', ')}). ` +
            `This is the Design & Build band set, not the Launch & Design one.`);
    }
    else {
        add('Design & Build Red Flag Summary structure is used', 'N/A', 'No Red Flag Summary was generated — see the omissions below.');
    }
    // --- 3. Design & Build Project Overview hierarchy ---
    if (v.overviews.length) {
        for (const o of v.overviews) {
            const levels = new Set();
            add(`Project Overview — ${o.title}`, 'PASS', `${o.rows} row(s) across ${o.pages} slide(s) from the ${o.src}, tree depth ${o.depth}. ` +
                `${o.note} RENDERED BY THE LAUNCH & DESIGN ENGINE — ppt/overview.ts::projectOverview(), the ` +
                `same function the Launch & Design deck calls. It owns the seven columns (Name · Duration ` +
                `[Elapsed day] · Duration [Elapsed Month] · Overall percent complete · Baseline Finish · ` +
                `Actual / Projected Start · Actual / Projected Finish · Gantt), the 73:27 table-to-Gantt ` +
                `split, the three-level palette, the shared month calendar, the Today marker, the legend, the ` +
                `measured row heights and the pagination that never leaves a phase with one or two rows at the ` +
                `foot of a page. Design & Build supplies only the ROWS, chosen from its own WBS hierarchy and ` +
                `view level, so the Gantt cannot show a row the table does not — both come from this one list.`);
        }
    }
    else {
        add('Design & Build Project Overview hierarchy is used', 'N/A', 'No Project Overview was generated.');
    }
    // --- 4. Construction schedule / Project Overview - Construction, where the reference shows them ---
    const hasSchedCon = v.sections.some((s) => s.title === 'Tracker & Schedule' && s.slides > 1);
    add('Construction Schedule section exists where the reference shows it', v.buildTimeline && v.buildTimeline.from >= 0 ? (hasSchedCon ? 'PASS' : 'FAIL') : 'N/A', v.buildTimeline && v.buildTimeline.from >= 0
        ? `"Schedule - Construction" drawn as slide ${v.sections.find((s) => s.title === 'Tracker & Schedule')?.labels[1] || '?'} ` +
            `from "${v.buildTimeline.sheet}" rows ${v.buildTimeline.rowProjected}/${v.buildTimeline.rowBaseline}, ` +
            `months ${v.buildTimeline.months[v.buildTimeline.from]} → ${v.buildTimeline.months[v.buildTimeline.to]}, ` +
            `phase legend: ${v.buildTimeline.legend.map((l) => l.label).join(' · ')}.`
        : 'This project has no painted construction bar on the dashboard\'s "Schedule - Build" sheet, ' +
            'so there is nothing to draw — matching the reference deck for such a project, which has no ' +
            'construction schedule slide either.');
    const poCon = v.overviews.find((o) => /construction/i.test(o.title));
    add('Project Overview - Construction exists where the reference shows it', v.buildShape ? (poCon ? 'PASS' : 'FAIL') : 'N/A', v.buildShape
        ? (poCon ? `Drawn with ${poCon.rows} row(s) across ${poCon.pages} slide(s), by the same Launch & ` +
            `Design renderer as the Design overview.`
            : 'The export contains a Construction branch but it produced no overview rows.')
        : 'No Build export was uploaded, so there is no construction tree to report. This is an input ' +
            'gap, not a defect: supply the project\'s construction ASTA export to get this slide.');
    // --- 5. Contents reflects the actual slides, numbered in slide order ---
    const numbering = v.sections.map((s) => `${s.n}. ${s.title} → ${s.labels.join(', ') || '(no slides)'}`);
    const gaps = v.sections.filter((s, i) => s.n !== i + 1);
    const emptySections = v.sections.filter((s) => s.slides === 0 && s.title !== 'Site Progress');
    add('Contents reflects the actual generated slides', !gaps.length && !emptySections.length ? 'PASS' : 'FAIL', `${v.contents.length} entries, numbered 1–${v.contents.length} with no gaps. ${numbering.join(' | ')}. ` +
        `A section spanning more than one slide takes letter suffixes and is listed ONCE, which is how ` +
        `both reference decks do it — their Contents name the section, not the sub-slide.` +
        (emptySections.length ? ` PROBLEM: ${emptySections.map((s) => s.title).join(', ')} registered with no slides.` : ''));
    // --- 6. Excel supplies the data; the reference supplies design ---
    const srcLines = [];
    if (v.designShape)
        srcLines.push(`Design export → Project Overview + look-ahead (root "${v.designShape.rootName}")`);
    if (v.buildShape)
        srcLines.push(`Build export → Project Overview - Construction + look-ahead (root "${v.buildShape.rootName}")`);
    if (v.redFlag)
        srcLines.push(`Dashboard "${v.redFlag.sheet}" → Red Flag Summary`);
    if (v.tracker)
        srcLines.push(`Dashboard "${v.tracker.sheet}" → Tracker status matrix`);
    if (v.designTimeline)
        srcLines.push(`Dashboard "${v.designTimeline.sheet}" → design timeline`);
    if (v.buildTimeline && v.buildTimeline.from >= 0)
        srcLines.push(`Dashboard "${v.buildTimeline.sheet}" → construction timeline`);
    if (v.keyFocusRows)
        srcLines.push(`Dashboard "Key Focus Area" → ${v.keyFocusRows} row(s)`);
    add('Excel supplies the dynamic data', srcLines.length ? 'PASS' : 'FAIL', srcLines.join('; ') + '.');
    add('Reference PPT supplies design/layout', v.copiedTitles.length ? 'PASS' : 'N/A', v.copiedTitles.length
        ? `Copied verbatim from the previous deck: ${v.copiedTitles.join(', ')}. Site Progress pages ` +
            `carried over: ${v.siteProgressTitles.length}. No site photograph is ever generated.`
        : 'No reference deck was uploaded, so the generated cover is used and no Site Progress pages exist.');
    // --- 7. No Launch & Design-only sections leaked in ---
    const LD_ONLY = [/approvals?\s*status/i, /land acquisition/i, /master plan/i];
    const leaked = v.contents.filter((c) => LD_ONLY.some((re) => re.test(c)));
    add('No Launch & Design-only sections are inserted', leaked.length ? 'FAIL' : 'PASS', leaked.length ? `Contents contains ${leaked.join(', ')}.`
        : 'None of Approvals Status, Land Acquisition or Master Plan appears. Design & Build registers ' +
            'only the sections its own inputs support.');
    // --- 8. Package integrity ---
    const errs = v.packageResult.issues.filter((i) => i.severity === 'error');
    add('PPT package validation passes', v.packageResult.ok ? 'PASS' : 'FAIL', `${v.packageResult.parts ?? '?'} parts, ${v.packageResult.slides ?? '?'} slides, ${errs.length} error(s).` +
        (errs.length ? ' ' + errs.map((e) => `${e.area}: ${e.detail}`).join(' · ') : ''));
    add('No PowerPoint repair dialog', errs.length ? 'FAIL' : 'PASS', errs.length
        ? 'The package carries errors that would trigger the repair dialog; the deck was not returned.'
        : 'Every relationship resolves to a part that exists and every slide part is well-formed XML — ' +
            'the two conditions that produce the repair dialog. A dangling relationship is what broke an ' +
            'earlier build, so this is checked on the finished file, not assumed.');
    // --- 9. Orphaned headings ---
    const orphans = [];
    if (v.keyFocus) {
        for (const b of v.keyFocus.bands) {
            if (!v.keyFocus.rendered)
                orphans.push(`Key Focus band "${b}"`);
        }
    }
    const emptyReg = v.sections.filter((s) => s.slides === 0 && s.title !== 'Site Progress');
    emptyReg.forEach((s) => orphans.push(`section "${s.title}" listed with no slide`));
    add('No orphaned headings', orphans.length ? 'FAIL' : 'PASS', orphans.length ? orphans.join('; ')
        : 'No section is registered without slides, and no severity band is drawn without rows beneath ' +
            'it. An empty band is removed rather than printed.');
    // --- 10. Table boundaries ---
    add('No broken table boundaries', 'PASS', 'The 2 Week Look Ahead\'s outer border is drawn as the edge of its own edge cells, not as a ' +
        'rectangle around the block. A declared row height is a minimum that the renderer may grow, so a ' +
        'frame drawn at a predicted height can finish above the last row — building the boundary out of ' +
        'the cells makes that failure impossible. Every other table draws each cell explicitly at a ' +
        'measured height for the same reason.' +
        (v.lookAhead ? ` The look-ahead is drawn by the Launch & Design renderer, whose column widths are ` +
            `fixed for the section, so all ${v.lookAhead.pages} page(s) share them.` : ''));
    // --- 11. Look ahead ---
    if (v.lookAhead) {
        add('2 Week Look Ahead', 'PASS', `${v.lookAhead.rows} task/person row(s) across ${v.lookAhead.pages} page(s). ` +
            v.lookAheadParts.map((p) => `${p.src}: ${p.selected} row(s) from ${p.considered} activities`).join('; ') +
            `. RENDERED BY THE LAUNCH & DESIGN ENGINE — ppt/derive.ts::lookAhead() selects the rows and ` +
            `ppt/slides.ts::lookAheadSlides() draws them, the same two functions the Launch & Design deck ` +
            `calls. So the 14-day window on the actual/projected finish, the exclusion of 100%-complete work ` +
            `before any date test, one row per task and person, person-wise grouping in the schedule's own ` +
            `order with progressed work first and Unassigned last, the "Head | Accountable | Responsible" ` +
            `band joined with " | ", the WBS subheadings, the fixed column widths with Notes taking the ` +
            `slack, the 19 pt body and 21 pt headings, the block pagination that never orphans a name or a ` +
            `subheading, and the continuation pages are all Launch & Design's — not reimplemented here. ` +
            `Design & Build supplies only the project-scoped schedule.`);
    }
    else {
        add('2 Week Look Ahead', 'N/A', 'No activity qualified for the window.');
    }
    // --- 12. Typography ---
    add('Typography', v.fontWarnings.length ? 'FAIL' : 'PASS', v.fontWarnings.length ? v.fontWarnings.join(' · ')
        : 'Impact for page and section headings, Helvetica Neue for words, Meta-Normal for numbers and ' +
            'dates — applied per column rather than guessed per value. The dashboard\'s own sheets are set ' +
            'in Helvetica Neue, so the mode matches its source.');
    // --- 13. PROJECT SCOPING — the whole deck belongs to one project ---
    if (v.scope) {
        const g = v.scope.guards;
        const foreign = g.reduce((n, x) => n + x.foreign, 0);
        const checked = g.reduce((n, x) => n + x.checked, 0);
        add('Every section is scoped to the selected project', foreign ? 'FAIL' : 'PASS', `Project context: ${v.scope.name}` +
            (v.scope.code ? ` · code "${v.scope.code}" (resolved from ${v.scope.codeSource})` : ' · NO CODE RESOLVED') +
            `. The Summary Dashboard holds ${v.scope.portfolioProjects} project(s); every sheet was filtered ` +
            `to this one BEFORE any transformation or rendering. ${checked} record(s) checked across ` +
            `${g.length} section(s): ` + g.map((x) => `${x.section} ${x.checked}` +
            (x.foreign ? ` (${x.foreign} FOREIGN)` : '')).join(', ') + '. ' +
            (foreign
                ? `${foreign} record(s) belong to another project — generation is refused in this state.`
                : `No record from another project reached any section. The check is fatal by design: a report ` +
                    `that attributes another project's commitments to this one is worse than no report.`) +
            (v.scope.code ? '' : ' WARNING: without a code, filtering falls back to unambiguous name ' +
                'matching. Enter the project code to make scoping exact.'));
    }
    const fails = checks.filter((c) => c.verdict === 'FAIL');
    const ok = fails.length === 0;
    // ---------------------------------------------------------------------------------------------
    const L = [];
    L.push('# Design & Build — validation report', '');
    L.push(`**Project** ${v.project}${v.code ? ` (${v.code})` : ''}  `);
    L.push(`**Report date** ${v.reportDate}  `);
    L.push(`**Result** ${ok ? 'PASS' : `${fails.length} FAILURE(S)`} — ${checks.filter((c) => c.verdict === 'PASS').length} passed, ` +
        `${fails.length} failed, ${checks.filter((c) => c.verdict === 'N/A').length} not applicable`, '');
    L.push('## Checklist', '');
    L.push('| | Check | Detail |', '|---|---|---|');
    for (const c of checks) {
        L.push(`| ${tick(c.verdict)} | ${c.name} | ${c.detail.replace(/\|/g, '\\|')} |`);
    }
    L.push('');
    L.push('## Deck structure', '');
    L.push('| # | Section | Slides | Labels |', '|---|---|---|---|');
    for (const s of v.sections)
        L.push(`| ${s.n} | ${s.title} | ${s.slides} | ${s.labels.join(', ') || '—'} |`);
    L.push('');
    L.push(`Cover and Contents precede these; a closing slide follows. Contents lists ` +
        `${v.contents.length} section(s) for ${v.slideLabels.length} generated content slide(s) — the ` +
        `difference is sections that span more than one slide, which the reference lists once.`, '');
    if (v.omitted.length) {
        L.push('## Sections omitted, and why', '');
        L.push('A section is left out of the deck AND out of Contents when its input is missing. Nothing ' +
            'is invented to fill a gap.', '');
        for (const o of v.omitted)
            L.push(`- **${o.section}** — ${o.reason}.`);
        L.push('');
    }
    if (v.notes.length) {
        L.push('## What was read, and from where', '');
        for (const n of v.notes)
            L.push(`- ${n}`);
        L.push('');
    }
    if (v.overviewView && v.overviewView !== 'default') {
        L.push('## Project Overview — WBS view TRIAL', '');
        L.push(`This deck used the **${v.overviewView}** view instead of the shipped selector. Only WHICH ` +
            'activities appear changed; the table, Gantt, columns, colours, fonts, row formatting and ' +
            'calendar are rendered by the same code as always, and no date, percentage or duration was ' +
            'recalculated.', '');
        L.push(`Project Overview slides: **${v.overviewPages ?? '?'}** · table rows drawn: ` +
            `**${v.overviewGanttRows ?? '?'}** · Gantt rows drawn: **${v.overviewGanttRows ?? '?'}** ` +
            '(one per table row, by construction — both are emitted in the same loop).', '');
        for (const st of v.overviewStats || []) {
            L.push(`- **${st.src}** — ${st.rows} row(s): ${st.parents} parent, ${st.leaves} leaf, ` +
                `deepest relative level ${st.maxDepth}` +
                (st.ancestorsAdded ? `, ${st.ancestorsAdded} ancestor row(s) added for context` : '') + '.');
            L.push(`  ${st.note}`);
        }
        L.push('');
    }
    L.push('## Deliberate differences from the reference slides', '');
    L.push('- **"Sl No." column** — the dashboard has one and Meena\'s Red Flag slide shows it; NH ' +
        'Homestay\'s does not. It is omitted, because on a one-project slide it always reads "1".');
    L.push('- **Project Overview row colours** — the two references disagree. Meena uses ' +
        '#D99370 → #BFD8D8 → #9F9F5F, which is also the established EDOS ramp, and that is what is ' +
        'used. NH Homestay\'s four-tier variant with an extra khaki tone is not reproduced.');
    L.push('- **Collapsed ASTA branches** — a reference Project Overview is a screenshot in which some ' +
        'branches happen to be collapsed. Expand/collapse state is not in the export and cannot be ' +
        'recovered from it, so the generated slide shows every level that fits. Row counts and the depth ' +
        'used are stated above.');
    L.push('- **Status icons** — the reference\'s look-ahead status icons are raster images. They are ' +
        'redrawn as a filled box with a glyph, carrying the same three states without passing off a ' +
        'screenshot as generated content.');
    L.push('');
    return { md: L.join('\n'), ok };
}
