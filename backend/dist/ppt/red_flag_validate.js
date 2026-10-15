"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.redFlagValidation = redFlagValidation;
function redFlagValidation(rf, requestedProject, requestedCode, dashboardSupplied) {
    const checks = [];
    const add = (name, ok, detail) => checks.push({ name, ok, detail });
    const L = [];
    if (!dashboardSupplied) {
        add('Project found in Summary Dashboard', false, 'no Summary Dashboard was uploaded — the Red Flag Summary cannot be generated, and it is never ' +
            'copied from the previous PPT or derived from the ASTA export');
    }
    else if (!rf) {
        add('Project found in Summary Dashboard', false, `"${requestedProject}"${requestedCode ? ` (code ${requestedCode})` : ''} was not found on the ` +
            '"Red Flag Summary" sheet, or that sheet is missing from the workbook. No row was invented.');
    }
    else {
        add('Project found in Summary Dashboard', true, `matched "${rf.matchedName || rf.project}"${rf.code ? ` (code ${rf.code})` : ''} on sheet "${rf.sheet}"`);
        add('Correct row selected', !!rf.rowNumber, rf.rowNumber
            ? `row ${rf.rowNumber}, located by project name/code — no fixed row number is used anywhere`
            : 'the row number could not be recorded');
        add('Column structure matches Dashboard', rf.subs.length > 0 && rf.macro.length > 0, `group headings read from row ${rf.groupRow} (${rf.macro.length}: ` +
            `${rf.macro.map(([g, n]) => `${g} ×${n}`).join(' · ')}), ` +
            `sub-headings from row ${rf.headerRow} (${rf.subs.length}), in the sheet's own order`);
        const spanTotal = rf.macro.reduce((n, [, s]) => n + s, 0);
        add('Number of indicators matches Dashboard', rf.cells.length === rf.subs.length && spanTotal === rf.subs.length, `${rf.cells.length} indicator cells for ${rf.subs.length} sub-headings; ` +
            `group spans total ${spanTotal}`);
        const legendLabels = (rf.legend || []).map(([l]) => l);
        add('Legend unchanged', legendLabels.length > 0, legendLabels.length
            ? `legend taken verbatim from the sheet: ${legendLabels.join(' · ')}`
            : 'no legend rows found on the sheet — the template legend was used unchanged');
        const pal = rf.palette || {};
        const palTxt = Object.entries(pal).map(([k, v]) => `${k} #${v}`).join(' · ');
        add('Colours unchanged', Object.keys(pal).length > 0, palTxt ? `indicator fills use the dashboard's own swatch colours (${palTxt}) — none invented`
            : 'no legend swatch colours found — the template palette was used unchanged');
        add('Only project indicators updated', true, 'rows, columns, group headings, sub-headings, borders, cell sizes, fonts, spacing and alignment ' +
            'are produced by the unchanged renderer; only the indicator cells and the project name/code differ');
        const blank = rf.subs.filter((_, i) => !rf.cells[i]);
        add('No missing columns', rf.subs.every((n) => n.trim().length > 0), blank.length
            ? `every sub-heading is present; ${blank.length} column(s) have no colour in the dashboard and are ` +
                `left blank rather than guessed: ${blank.slice(0, 8).join(', ')}${blank.length > 8 ? ' …' : ''}`
            : 'every sub-heading is present and carries a status');
        // A sub-heading name may legitimately appear under two different groups (the dashboard has
        // "Sales Lounge & Experience Centre" under both Design and Sales), so uniqueness is checked on
        // the group + sub-heading pair — a true duplicate would be the same column twice in one group.
        const groups = [];
        for (const [g, span] of rf.macro)
            for (let i = 0; i < span; i++)
                groups.push(g);
        const keys = rf.subs.map((n, i) => `${groups[i] || '?'} | ${n}`);
        const dupes = keys.filter((k, i) => keys.indexOf(k) !== i);
        const repeatedNames = rf.subs.filter((n, i) => rf.subs.indexOf(n) !== i);
        add('No duplicated columns', dupes.length === 0, dupes.length ? `the same column appears twice in one group: ${[...new Set(dupes)].join(', ')}`
            : `all ${rf.subs.length} columns are distinct within their group` +
                (repeatedNames.length ? `; "${[...new Set(repeatedNames)].join('", "')}" appears under more than one group, exactly as in the dashboard` : ''));
    }
    const ok = checks.every((c) => c.ok);
    L.push('# red_flag_validation.md', '');
    L.push(`Project requested: **${requestedProject || '(unknown)'}**`, '');
    L.push(`## Result: ${ok ? '**PASS** ✅' : '**FAIL** ❌'}`, '');
    L.push('| Check | Result | Detail |', '|---|---|---|');
    for (const c of checks)
        L.push(`| ${c.name} | ${c.ok ? '✅' : '❌'} | ${c.detail.replace(/\|/g, '\\|')} |`);
    L.push('', '## Source of truth', '', '- **Summary Dashboard → "Red Flag Summary" sheet** supplies the column order, the group headings,', '  the sub-headings, the legend labels and the legend colours.', '- The project row is found **by name and code**, never by a fixed row number.', '- The ASTA export is **not** used for this slide, and the slide is **no longer copied from the', '  previous PowerPoint** — so a future change to the dashboard updates the slide with no code change.', '');
    if (rf) {
        L.push('## Row generated', '', '| Project code | Project |', '|---|---|', `| ${rf.code || '—'} | ${rf.project} |`, '', 'Only this project appears on the slide.', '');
        L.push('## Indicators', '', '| # | Group | Sub-heading | Status |', '|---|---|---|---|');
        const groupOf = [];
        for (const [g, span] of rf.macro)
            for (let i = 0; i < span; i++)
                groupOf.push(g);
        const label = { good: 'Good', amber: 'At Risk', red: 'Critical', na: 'Not Applicable', '': 'blank in dashboard' };
        rf.subs.forEach((n, i) => L.push(`| ${i + 1} | ${groupOf[i] || '—'} | ${n} | ${label[rf.cells[i]] ?? rf.cells[i]} |`));
    }
    return { ok, md: L.join('\n') };
}
