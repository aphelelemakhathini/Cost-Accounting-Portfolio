# Aphelele: Cost Accounting Automation Portfolio

I'm a Cost Accountant Controller in a pump and valve manufacturer. My month runs on SAP exports, and these are the Excel VBA tools I built to turn those exports into reports that production, management and Industrial Engineering act on.

> **All data here is fictional.** "Acme Pumps", its plants (1100 / 1200 / 1300), customers, work centres, materials and every figure were made up by the `generate_data.py` script in each project. The macros are my own work, with employer-specific names, codes and settings replaced.

| # | Project | Before → After |
|---|---------|----------------|
| 01 | [Production variances](01-production-variances/) | Order cost report + operations export → flagged orders, department tabs, favourable variances, and a live overview with a plant slicer |
| 02 | [Warranty & goodwill claims](02-warranty-claims/) | Grouped claims export → one-line-per-order summary → posted into the year-to-date register |
| 03 | [Costed BOM](03-bom-costing/) | CK11N cost estimate text export → grouped BOM with live formulas that reconciles to SAP |
| 04 | [Sales margin (CMII)](04-sales-margin/) | Invoiced sales export → CMII table, loss and margin analysis, target progress |
| 05 | [Cost element breakdown](05-cost-element-breakdown/) | Actual costs report → grouped cost elements with reconciling lines |
| 06 | [Cost Controlling Dashboard](06-cost-controlling-dashboard/) | All the daily SAP downloads → one interactive dashboard (demo with screenshots) |

## Time saved
Compilation that used to be done by hand now runs in minutes, leaving more time for analysis.

| Task | Before | After |
|---|---|---|
| Production variance compilation (3 plants) | A whole day of compiling | 10 minutes |
| Monthly claims compilation | 4 hours | 5 minutes |
| Sales report compilation | 2 hours per market area | 10 minutes |
| Sales order cost breakdown for a requester | 30 minutes | 3 minutes |
| SAP downloads for the dashboard | Run one by one | One 5-minute SAP script |

## Techniques used across the projects
- **Parsing messy ERP exports:** outline blanks stored as empty text, numbers as text in SAP's format (`1.234,56`, `1 234,56`, trailing minus), title lines, encodings, variable-depth hierarchies
- **Rules management can follow:** ±45 % labour / ±20 % material bands, 2-hour / 1-hour / 50 % operation thresholds, partial-invoicing outliers left out of margin averages
- **Workbooks that stay correct:** formulas rebuilt by code, no hard-coded column letters, re-running a month replaces it instead of doubling it, typed remarks kept when sheets are rebuilt
- **Live Excel:** LET / FILTER / SORTBY dynamic arrays, slicers driving named ranges, pivot tables, ListObjects with total rows, conditional formatting, charts built from code
- **Regional settings handled:** formulas and conditional formats written so they work on comma-decimal PCs

## Running a project
1. Open the demo file from the project's `data/` folder in Excel.
2. Press **Alt+F11**, choose **File → Import File…**, and import the `.bas` module(s) from `vba/` **into that same workbook**.
3. Press **Alt+F8** and run the macro named in the project README.

The README for each project lists the results you should get on the demo data.

**Certifications:** FMVA · BIDA · FPAP · Google Data Analytics · BCom Accounting (UJ)
