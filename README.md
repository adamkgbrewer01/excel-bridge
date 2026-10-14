# Excel Bridge

An Excel add-in that brings three Power BI and Tableau features to any Excel table:

- charts that filter each other when you click a bar
- drill-down from year to quarter to month, and through hierarchies such as region to country
- small multiples, one chart per region or product on a shared scale

It can also apply the same filters to the table on the sheet, so totals rows and SUBTOTAL formulas follow your clicks.

Try it in a browser with made-up sales data: https://adamkgbrewer01.github.io/excel-bridge/dev/dev.html

## Install in Excel for Mac

1. Download `manifest.xml`.
2. Put it in `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (create the folder if needed).
3. Restart Excel, then choose Home > Dashboard.

## Files

- `taskpane.html`, `styles.css`, `js/` the add-in itself (no build step)
- `js/engine.js` field detection, hierarchies, cross-filtering and drill logic, with tests in `tests/` (`node --test tests/`)
- `js/excel-bridge.js` the only file that calls the Excel JavaScript API
- `dev/` the browser preview, using `dev/mock-bridge.js` in place of Excel
- `tools/` scripts that make the sample data, icons and manifest
