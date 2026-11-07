# PilotScout — Microsoft 365 Copilot Seat Savings Audit

Find the Copilot seats you're paying $360/year for and nobody's using — then recover the spend.
PilotScout runs **read-only inside your own tenant**. It has no server and collects nothing; the report is written to your machine.

---

## This repository: the free Lite build

This is **PilotScout Lite**, free and open source under the MIT License. A run prints a console summary and writes `never-used-seats.csv`, `dormant-seats.csv` and `disabled-but-licensed.csv`; it shows no dollar totals and builds no HTML report. Anything below that mentions the HTML report, dollar totals, Solo or Pro belongs to the paid editions at [pilotscout.dev](https://pilotscout.dev); that code is not in this script. The `-Lite` switch changes nothing in this build: it always runs as Lite.

## What you get

| File | Purpose |
|------|---------|
| `PilotScout-Audit.ps1` | The audit engine (PowerShell + Microsoft Graph) |
| `config/PriceTable.csv` | Editable Copilot price list — set your real seat price here |
| `test/mock-tenant.json` | Demo data to try the tool with no tenant access |

Licensed MIT — Solo and Pro are under the EULA.

Outputs per run: a polished **HTML savings report** + **CSV exports** for each finding (Lite: console summary + never-used / dormant / disabled CSVs, no HTML).

---

## Requirements

- **PowerShell 7+** (Windows, macOS, or Linux) — or Windows PowerShell 5.1
- **Microsoft Graph PowerShell SDK**:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- An account that can consent to **read-only** Graph scopes and read usage reports
  (**Reports Reader** or Global Reader is enough)
- **No Microsoft Entra ID P1 required** — Copilot activity comes from the usage report, not sign-in logs

---

## Quick start (try it with demo data — no tenant needed)

```powershell
.\PilotScout-Audit.ps1 -MockDataPath .\test\mock-tenant.json
```
Open the generated `PilotScout-Report-<timestamp>\PilotScout-Report.html`.

## Run against your tenant

```powershell
# 1) (first time) install the SDK
Install-Module Microsoft.Graph -Scope CurrentUser

# 2) run the audit (you'll be prompted to sign in & consent to read-only scopes)
.\PilotScout-Audit.ps1 -InactiveDays 60
```

You'll be asked to consent to: `Organization.Read.All`, `User.Read.All`, `Directory.Read.All`, `Reports.Read.All`. All read-only.

---

## Options

| Parameter | Default | Description |
|-----------|---------|-------------|
| `-InactiveDays` | `60` | A seat with no Copilot activity in this many days is flagged dormant |
| `-UsagePeriod` | `D180` | Usage report window: `D7` `D30` `D90` `D180` `ALL` |
| `-OutputPath` | auto | Folder for the report + CSVs |
| `-PriceTablePath` | `.\config\PriceTable.csv` | Your price list |
| `-CompanyName` | — | **(Pro)** White-label the report header |
| `-CurrencySymbol` / `-CurrencyLabel` | `$` / `USD` | Display currency |
| `-Lite` | off | Free mode: CSV + summary only (no $ totals, no HTML) |
| `-MockDataPath` | — | Use demo JSON instead of connecting to Graph |

**Example (consultant, white-labelled, 90-day threshold):**
```powershell
.\PilotScout-Audit.ps1 -InactiveDays 90 -CompanyName "Contoso IT Partners"
```

---

## Set your price (important)

Copilot pricing varies by agreement and bundle. Open `config/PriceTable.csv` and check `MonthlyPriceUSD` (ships with the 2026 list price of 30 USD for `Microsoft_365_Copilot` and `M365_Copilot`, and a row for `Microsoft_365_Copilot_EDU` with the price left blank: no Microsoft page states it).

PilotScout finds your Copilot seats by **service plan**, not by name: every subscribed SKU that includes `M365_COPILOT_APPS` is audited — the add-on under any part number, Copilot Business, and bundles such as Microsoft 365 E7 (several Copilot SKUs are each audited and added up). Copilot Chat on its own (`Bing_Chat_Enterprise`, included in E3/E5/Business plans) is free and never counted. The price table only sets the price: a Copilot SKU without a row, or with a blank price, is priced at 30 USD and flagged **price assumed** in the console and the report. Set a price of 0 to leave a SKU out of the dollar totals. For Copilot Business (21 USD list) or a bundle, run `Get-MgSubscribedSku | Select SkuPartNumber` and add that part number with its Copilot-attributable price.

---

## What it checks

1. **Unassigned purchased seats** — Copilot bought but assigned to no one (cleanest win)
2. **Never-used seats** — assigned, account active, zero Copilot activity in the window
3. **Dormant seats** — no Copilot activity in > N days
4. **Disabled but licensed** — blocked accounts still holding Copilot
5. **Chat-only users** *(advisory)* — active only in Copilot Chat; free-tier candidates
6. **App adoption depth** *(advisory)* — active users per app (Word, Excel, PowerPoint, Outlook, Teams, Chat, OneNote, Loop)

Headline savings = items 1–4 (conservative). Advisory items are listed but not added to the total.

---

## Security & privacy

- **Read-only.** PilotScout never writes or changes anything in your tenant.
- **No exfiltration.** No network calls except to Microsoft Graph. The report is generated locally.
- Review the script before running — it's plain PowerShell, intentionally readable.

---

## Troubleshooting

- *"Copilot usage report unavailable"* → your account lacks `Reports.Read.All` or a reports-capable role; PilotScout still reports unassigned + disabled seats.
- *Names show as concealed / findings are counts only* → your tenant hides names in reports. Microsoft 365 admin center → **Settings → Org settings → Reports** → untick *Display concealed user, group, and site names*, wait for the next report refresh (~24-48 h), re-run.
- *"No Microsoft 365 Copilot seat SKU found"* → no subscribed SKU in the tenant includes the `M365_COPILOT_APPS` service plan, so there are no paid Copilot seats to audit (Copilot Chat alone is free). The message lists the part numbers it saw.
- *"price assumed for …"* → that Copilot SKU has no row, or a blank price, in `config/PriceTable.csv`; it is priced at 30 USD until you set one.
- *Module not found* → `Install-Module Microsoft.Graph -Scope CurrentUser`
- *Throttling on large tenants* → re-run; Graph paging is handled, transient 429s resolve on retry.

---

PilotScout · pilotscout.dev · Part of the Scout suite (SeatScout · StorageScout) · Independent tool, not affiliated with or endorsed by Microsoft. "Microsoft 365 Copilot" is a trademark of Microsoft Corporation.

## Author

Built and maintained by Uğur Sayar, a SharePoint and Microsoft 365 architect in Istanbul ([LinkedIn](https://www.linkedin.com/in/ugursayar)). Bug reports and pull requests are welcome through GitHub issues.
