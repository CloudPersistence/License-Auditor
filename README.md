CloudPersistence License Auditor — Microsoft 365 license waste audit
A read-only PowerShell tool that connects to Microsoft Graph and produces a
local HTML report showing where a tenant is wasting money on M365 licenses.
It detects three high-confidence sources of waste:
Shelfware — paid licenses that are purchased but not assigned.
Disabled users with licenses — departed accounts still consuming paid seats.
Inactive users — licensed users with no sign-in for N+ days (requires Entra ID P1/P2).
No tenant data ever leaves the machine it runs on. Nothing is written back to
the tenant — the tool only reads.
---
Requirements
PowerShell 7+ (recommended) or Windows PowerShell 5.1.
Microsoft Graph PowerShell modules:
```powershell
  Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
  ```
An account that can read the tenant. `Global Reader` is enough to read
the data. Note: the first time the "Microsoft Graph Command Line Tools" app is
used in a tenant, an account that can grant admin consent (Global
Administrator, Privileged Role Administrator, or Cloud Application
Administrator) may be required once. After that, Global Reader works.
The scopes requested are all read-only:
`Organization.Read.All`, `User.Read.All`, `Directory.Read.All`, `AuditLog.Read.All`.
---
How to run
```powershell
# Interactive (default): asks currency + price source + a price per tenant SKU
.\Invoke-CPLicenseAuditor.ps1

# Different inactivity threshold and output folder
.\Invoke-CPLicenseAuditor.ps1 -InactiveDays 60 -OutputFolder "C:\Reports"

# No prompts: fill missing prices from the built-in USD hints (rough)
.\Invoke-CPLicenseAuditor.ps1 -UseListPrices

# No prompts: use only what's already in the price file (rest shown as n/a)
.\Invoke-CPLicenseAuditor.ps1 -NonInteractive

# Per-client price file + an explicit source label
.\Invoke-CPLicenseAuditor.ps1 -PriceFile "C:\Clients\EHL\prices.json" -PriceSource "Invoice 2026-05"
```
On the first run the tool asks for the currency, a price source label (free
text, shown in the report footer), and a monthly price for each SKU the tenant
actually has. Answers are saved to `CPLicenseAuditor-prices.json` so later runs are
prompt-free. Prices accept either comma or dot as the decimal separator.
---
Where do the prices come from?
Microsoft Graph tells you how many licenses exist and how many are assigned,
but not how much they cost — price is contractual data that is not exposed
through Graph. You provide the price. Where you find the real figure depends on
how the tenant buys its licenses.
Step 0 — identify the billing channel.
In the Microsoft 365 admin center go to Billing → Billing accounts and check
the account type:
Account type	What it means	Where the price lives
MOSA (Microsoft Online Subscription Agreement)	Bought directly, self-service	Admin center (see below)
MCA (Microsoft Customer Agreement), direct	Bought direct from Microsoft / via rep	Admin center (see below)
MCA via partner / CSP	A partner resells and bills the customer	Partner invoice (not the admin center)
If the tenant buys directly (MOSA / MCA direct):
`Billing → Your products` → price per subscription / seat.
`Billing → Bills & payments` (Invoices) → the actual invoiced price. This
is the most accurate source — real amount, real currency, discounts applied.
If the tenant is on CSP (partner/reseller):
The real price is not shown in the customer's admin center. In CSP the
partner sets the price and bills the customer.
Get the figure from the partner's invoice. If you (the operator) are the
CSP partner, the prices are in your Partner Center price sheet for the
customer's market — you don't need the customer's admin center at all.
Two normalisation rules so the numbers are consistent:
Monthly vs annual. A monthly term is ~20% more expensive per month than an
annual commitment. Enter all prices on the same basis (monthly is recommended).
Exclude VAT. Microsoft commercial pricing is quoted excluding tax (except
AU/BR). Use the net price, not the tax-inclusive line on an invoice.
Record what you used in the price source label (e.g. `Invoice 2026-05`,
`CSP price sheet DE (EUR)`, `Manual estimate`). It appears in the report footer
so anyone reading the report knows how solid the savings figures are.
---
Output
A self-contained HTML file (`CPLicenseAuditor-Report-<timestamp>.html`) written to the
output folder and opened in the default browser. It contains:
Summary cards: estimated monthly spend, shelfware cost, disabled-user cost,
total recoverable per month.
License utilisation by plan.
Shelfware table (paid but unassigned).
Licenses on disabled accounts.
Inactive licensed users (or an N/A notice if the tenant has no Entra ID P1).
---
Known limitations
Inactive users need Entra ID P1/P2. Without it, Graph rejects the
`signInActivity` query entirely; the tool detects this, retries without it,
and marks that section N/A.
Prices are only as good as what you enter. Any SKU with no price is shown
as `n/a` rather than guessed.
Feature-level downgrades (e.g. E5 → E3) are not in this version. That
analysis depends on per-user usage reports, which are de-identified by default
in the tenant (`Reports` setting "Display concealed user, group, and site
names"). It is planned as a later iteration.
---
Powered by cloudpersistence.com
