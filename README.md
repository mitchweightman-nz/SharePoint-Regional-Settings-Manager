# SharePoint Regional Settings Manager

A Windows PowerShell tool for discovering, exporting, updating and restoring SharePoint Online site collection locales. It applies English (New Zealand), locale ID `5129`, to the root webs of discovered site collections.

## Requirements

- Windows with PowerShell 7.4 or later.
- PnP.PowerShell 2.12.0 or later, compatible with your version of PowerShell. The tool offers to install it if a compatible version is missing.
- The SharePoint tenant admin URL and an Entra public client application ID. Configure the app for interactive sign-in with the `http://localhost` redirect URI under **Mobile and desktop applications**, and consent to the appropriate SharePoint delegated permissions (normally `AllSites.FullControl`). The signed-in account needs SharePoint admin permissions for discovery and permission to manage regional settings on the sites it updates.

## Run

Open PowerShell in the repository folder and run the following command. It sets `RemoteSigned` for the app's PowerShell session only. Organisation Group Policy takes precedence.

```powershell
pwsh -NoProfile -STA -ExecutionPolicy RemoteSigned -File .\SharePointRegionalSettingsManager.ps1
```

If you downloaded the script and Windows marked it as coming from the internet, review the file and run the following command before launching it.

```powershell
Unblock-File -LiteralPath .\SharePointRegionalSettingsManager.ps1
```

[PowerShell execution policy reference](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/about/about_execution_policies)

Enter the tenant admin URL and client ID in the app, or prefill them with `-TenantAdminUrl` and `-ClientId`. For a Multi-Geo tenant, pass the additional admin URLs with `-AdditionalAdminUrls`.

1. Select **Connect**, then **Discover Sites**.
2. Select **Export Settings** to save the current locales before applying changes.
3. Select **Apply NZ Locale** to update discovered sites. The tool checks each site's current locale, skips sites already set to `5129` and verifies changes in SharePoint.
4. To restore previous locales, select **Rollback** and load an export or recovery CSV. The app shows the recovery CSV path in the live log. Each original locale is written to that file before the corresponding site is updated.

Use **Grid State CSV** to export the full inventory and its current status.
