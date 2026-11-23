SharePoint Regional Settings Manager
Planned Windows Forms application for PowerShell 7 and PnP.PowerShell. It will manage SharePoint Online regional settings at tenant scale, targeting English (New Zealand), LocaleId 5129.
Project scope
1. Verify PnP.PowerShell at startup. Offer installation if missing and exit cleanly if it remains unavailable.
2. Sign in interactively with Connect-PnPOnline and discover tenant sites.
3. Export current regional settings before changes.
4. Apply LocaleId 5129 to sites with a different locale.
5. Restore previous LocaleId values from an export.
6. Show live progress and each site's before and after state.
The application is intended for tenants with hundreds or thousands of sites. This repository currently contains the project brief.
