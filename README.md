# Westline Collections Portfolio

Four projects that show how a B2B collections desk runs end to end, built around **Westline Auto Parts**, a fictional four-yard auto parts distributor in California that sells on credit to body shops and garages.

**Live site:** open `index.html` through GitHub Pages (Settings → Pages).

| Project | What it is | Where |
| --- | --- | --- |
| Collections Desk dashboard | Morning view of cash against target, aging, risk by yard, a ranked work queue, team scorecard and cross-team requests | [`dashboard/`](dashboard/) |
| AR Operating Workbook | Google Sheets workbook behind the dashboard: priority scoring, next-action rules, work queue and KPIs, all formula-driven | Google Sheets (link on the site) |
| Apps Script automation | Adds a menu to the workbook: builds the call-disposition Google Form, rolls the as-of date each morning, saves a daily snapshot, emails the morning summary, flags requests past service level | [`automation/Westline_AR_Automation.gs`](automation/Westline_AR_Automation.gs) |
| Month-End AR Reconciliation SOP | Six-day close: ledger tie-out, six-way classification of every open invoice, RACI, controls, approval limits and a worked September 2026 close | [`documents/`](documents/) |
| Chargeback and Refund Root-Cause Map | 175 chargebacks and 430 escalations traced to the failed hand-off between teams, with countermeasures and results | [`documents/`](documents/) |
| Collections Team Playbook | 19-slide training deck: 30-day ramp, call framework, scripts, objections, promise rules, escalations, fraud red flags, QA scorecard, coaching | [`documents/`](documents/) |

## Headline results (synthetic data)

- October collections on pace for 166% of the $400K monthly target
- Balances older than 60 days down from 27.1% to 5.4%
- Card chargebacks down from 12.6 to 3.8 a month
- Dials per specialist up from 49 to 85 a day, with more calls reaching the payer

## Data

Westline is fictional. Every account, invoice, payment, call, request, chargeback and escalation comes from one simulated dataset covering about 700 credit accounts over 19 months, so the numbers agree across all projects. No real customer data is used.
