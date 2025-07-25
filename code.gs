/**
* Counts the number of cells matching the background of the cell with the formula.
* @param {range} countRange Range to be evaluated
* @param {range} checkboxCell Toggle this checkbox to refresh count
* @customfunction
*/
function COUNTCELLS(countRange, checkboxCell) {
    let activeRange = SpreadsheetApp.getActiveRange();
    let activeSheet = activeRange.getSheet();
    let color = activeRange.getBackground();
    let formula = activeRange.getFormula();
    let match = formula.match(/^\s*=\s*COUNTCELLS\s*\(\s*([$A-Z]+\d+:[$A-Z]+\d+)(?:\s*,.*)?\s*\)$/i);
    let rangeA1Notation = match ? match[1] : null;
    let range = activeSheet.getRange(rangeA1Notation);
    let bg = range.getBackgrounds();
    let count = 0;

    for (i = 0; i < bg.length; i++)
        for (j = 0; j < bg[0].length; j++)
            if (bg[i][j] == color)
                count = count + 1;
    return count;
}
