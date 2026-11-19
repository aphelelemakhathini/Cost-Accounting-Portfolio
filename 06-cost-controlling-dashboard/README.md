# 06: Cost Controlling Dashboard

## The problem
My daily and month-end work sits in a dozen separate SAP downloads and workbooks: sales CMII, order cost reports, operation confirmations, the TB, the WIP report, FI vs CO-PA differences, and the claims register. There was no single place to see where the month stands, what needs attention, and what I still owe people.

## What I built
A single-page dashboard (HTML + JavaScript, with SheetJS to read Excel and text files in the browser) that loads the SAP downloads directly and works out which file is which:

- **Overview:** headline tiles, sales by market area, duties due, and **auto-drafted commentary** written from the numbers
- **Sales & CMII:** by market area, distribution channel and segment, with losses, partial invoicing and memos flagged
- **Labour / Material Variances:** by department and work centre, ±45 % and ±R500 rules, zero-standard orders
- **Work Centres:** efficiency, cost variance, and machining vs setup hours by month and by week
- **Cost Centres:** plan vs actual, converted from EUR to ZAR at an adjustable rate
- **Claims & Goodwill:** monthly cost, % of turnover, year-on-year, by plant
- **CO Checks:** invoices hanging between FI and CO-PA, and the TB WIP vs WIP report difference
- **Action Plan, Duties and Email Queries:** month-end duties scheduled by working day with due-date pop-ups, and outstanding email queries with responded ticks and update notes
- Filters by month / YTD / trailing 12 months, plant, order type, market area, and a global order or material search

The real version keeps its history in a database and checks my Outlook inbox every weekday morning.

## Demo
`index.html` is the demo build: open it in any browser. **Every figure, customer, plant and work centre is fictional (Acme Pumps).** Screenshots are in `screenshots/`.
