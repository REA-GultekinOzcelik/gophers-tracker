# Gophers Tracker + CRM

Standalone web app (GitHub Pages) for Gültekin's daily system. Sign-in with Microsoft; all data lives in OneDrive Excel files.

- `index.html` — Tracker: daily checklist, tasks (Planner sync via Power Automate), weekly 4-1-1, CRM due list.
- `crm.html` — CRM: contact log (table `CRM` in Gophers_CRM_2026.xlsx), today's calls, contracts, rentals, opportunities, new leads from the Rehber folder (.vcf/.txt exports).

Both pages share one sign-in (redirect URI = app root; the CRM page bounces back via `state=crm`).
