# CloudPersistence License Auditor

A read-only PowerShell tool that audits Microsoft 365 license waste and generates a local HTML report.

**What it finds:**

- **Shelfware** — paid licenses that are purchased but never assigned to anyone
- **Disabled users with licenses** — departed employees still consuming paid seats
- **Inactive users** — licensed users with no sign-in for 90+ days *(requires Entra ID P1/P2)*

No data leaves your machine. Nothing is written back to the tenant.

---

## Quick start

```powershell
# Install Graph modules (once)
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser

# Run the audit
.\Invoke-CPLicenseAuditor.ps1
```

The script opens a browser for interactive sign-in, asks for currency and prices on first run, then generates an HTML report.

---

## Requirements

| Requirement | Details |
|---|---|
| PowerShell | 7+ recommended, or Windows PowerShell 5.1 |
| Modules | `Microsoft.Graph.Authentication`, `.Users`, `.Identity.DirectoryManagement` |
| Role | **Global Reader** (read-only) |
| First-time consent | May require admin consent once per tenant for "Microsoft Graph Command Line Tools" |

**Scopes requested** (all read-only): `Organization.Read.All`, `User.Read.All`, `Directory.Read.All`, `AuditLog.Read.All`

---

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-InactiveDays` | `90` | Days without sign-in before flagging a user |
| `-OutputFolder` | Current directory | Where to save the HTML report |
| `-PriceFile` | `CPLicenseAuditor-prices.json` | Path to the JSON price file |
| `-PriceSource` | *(prompted)* | Label for the report footer, e.g. `"Invoice 2026-05"` |
| `-UseListPrices` | Off | Skip prompts, use built-in USD list-price estimates |
| `-NonInteractive` | Off | Skip prompts, use only what's in the price file |
| `-NoOpen` | Off | Don't auto-open the report in the browser |

**Examples:**

```powershell
# Interactive (default)
.\Invoke-CPLicenseAuditor.ps1

# Quick run with built-in USD estimates (no prompts)
.\Invoke-CPLicenseAuditor.ps1 -UseListPrices

# Per-client price file + custom source label
.\Invoke-CPLicenseAuditor.ps1 -PriceFile "C:\Clients\ClientA\prices.json" -PriceSource "EA invoice Q2-2026"

# Different inactivity threshold
.\Invoke-CPLicenseAuditor.ps1 -InactiveDays 60
```

---

## Where do the prices come from?

Microsoft Graph returns *how many* licenses exist and *how many* are assigned, but **not how much they cost**. You provide the price. On first run the script prompts you for each SKU it finds, with a rough USD hint as reference.

Where to find the real price depends on how the tenant buys licenses:

### Step 1 — Identify the billing channel

Go to **Admin center → Billing → Billing accounts** and check the account type.

| Account type | Meaning |
|---|---|
| **MOSA** | Bought directly from Microsoft (self-service) |
| **MCA direct** | Bought direct from Microsoft or via a rep |
| **MCA via partner / CSP** | A partner resells and bills the customer |

### Step 2 — Find the price

**Direct purchase (MOSA / MCA direct):**
- `Billing → Your products` → price per subscription
- `Billing → Bills & payments` → invoiced price *(most accurate — includes discounts)*

**CSP (partner/reseller):**
- The price is on the **partner's invoice**, not in the customer's admin center
- If *you* are the CSP partner, check your Partner Center price sheet for the customer's market

### Normalisation rules

- **Monthly vs annual:** monthly terms are ~20% more expensive per month. Enter all prices on the same basis (monthly recommended).
- **Exclude VAT:** Microsoft commercial pricing excludes tax (except AU/BR). Use the net price.

---

## Multi-client workflow

Keep a folder per client with its own price file:

```
C:\Clients\
├── ClientA\
│   └── prices.json      # CHF, negotiated EA prices
├── ClientB\
│   └── prices.json      # EUR, CSP partner prices
└── ClientC\
    └── prices.json      # USD, direct purchase
```

```powershell
.\Invoke-CPLicenseAuditor.ps1 `
    -PriceFile "C:\Clients\ClientA\prices.json" `
    -PriceSource "EA invoice Q2-2026" `
    -OutputFolder "C:\Clients\ClientA"
```

---

## Known limitations

- **Inactive users require Entra ID P1/P2.** Without it, Graph rejects the `signInActivity` query. The tool detects this, retries without it, and marks the section as N/A.
- **Prices are only as good as what you enter.** Unknown SKUs show `n/a` instead of guessing.
- **Feature-level downgrades** (e.g. E5 → E3 based on actual feature usage) are planned for a future version.

---

## License

MIT

---

Powered by [cloudpersistence.com](https://www.cloudpersistence.com)
