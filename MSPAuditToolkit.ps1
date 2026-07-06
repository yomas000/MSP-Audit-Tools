#Requires -Version 5.1
<#
.SYNOPSIS
    MSP Audit Toolkit - WinForms GUI wrapping tenant app registration and the
    device/license audit into a single packaged tool (for IExpress bundling).

.DESCRIPTION
    Three tabs:
      "Tenant Setup"   - reads Tenants.csv, registers an app + grants Graph
                         permissions in each tenant (one-time, interactive sign-in
                         with MFA per tenant), writes TenantAuditApps.csv.
      "Run Audit"      - reads TenantAuditApps.csv, authenticates silently
                         (app-only) to each tenant, pulls device + license counts,
                         writes dated CSV reports and shows results in-app.
      "Software Audit" - reads AutomateConfig.csv, pulls installed software per
                         device from ConnectWise Automate for billing/licensing
                         audits, writes a raw per-device CSV and a per-title
                         summary CSV (count of devices running each title).

    Must run in STA mode (WinForms requirement):
        powershell.exe -STA -ExecutionPolicy Bypass -File MSPAuditToolkit.ps1

    Logic here mirrors New-TenantAuditAppRegistration.ps1 and
    Invoke-TenantDeviceLicenseAudit.ps1 - see those scripts for line-by-line
    comments on what each Graph call does.

    NOTE ON CONNECTWISE AUTOMATE: the endpoint paths and field names in the
    Software Audit tab reflect Automate's standard documented REST API shape.
    Automate versions/configurations can vary. If a call fails or field names
    don't match your instance, check your server's own API docs at:
        https://<your-automate-server>/cwa/api/v1/apidocs
    and adjust the Get-AutomateToken / Get-AutomateComputers / Get-AutomateSoftware
    functions below accordingly - the log will show the raw error/response.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Microsoft.Graph PowerShell module is not installed on this machine.`n`nRun this once, then relaunch:`nInstall-Module Microsoft.Graph -Scope CurrentUser",
        "Missing Dependency", "OK", "Warning") | Out-Null
    return
}

# ----------------------------- SHARED CONFIG -----------------------------

$WorkDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TenantsCsv = Join-Path $WorkDir "Tenants.csv"
$AppCredsCsv = Join-Path $WorkDir "TenantAuditApps.csv"
$AutomateConfigCsv = Join-Path $WorkDir "AutomateConfig.csv"
$AppDisplayName = "MSP-Audit-Tool"
$GraphAppId = "00000003-0000-0000-c000-000000000000"

$RequiredPermissions = @(
    @{ Name = "Directory.Read.All"; Id = "7ab1d382-f21e-4acd-a863-ba3e13f7da61" }
    @{ Name = "DeviceManagementManagedDevices.Read.All"; Id = "2f51be20-0bb4-4fed-bf7b-db946066c75e" }
    @{ Name = "Organization.Read.All"; Id = "498476ce-e0fe-48b0-b801-37ba7e2685c6" }
)

$SkuFriendlyNames = @{
    "SPE_E3" = "Microsoft 365 E3"; "SPE_E5" = "Microsoft 365 E5"
    "ENTERPRISEPACK" = "Office 365 E3"; "ENTERPRISEPREMIUM" = "Office 365 E5"
    "O365_BUSINESS_PREMIUM" = "Microsoft 365 Business Standard"; "SPB" = "Microsoft 365 Business Premium"
    "EXCHANGESTANDARD" = "Exchange Online (Plan 1)"; "EXCHANGEENTERPRISE" = "Exchange Online (Plan 2)"
    "AAD_PREMIUM" = "Entra ID P1"; "AAD_PREMIUM_P2" = "Entra ID P2"
    "EMS" = "Enterprise Mobility + Security E3"; "EMSPREMIUM" = "Enterprise Mobility + Security E5"
}

# ----------------------------- MAIN FORM -----------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "MSP Audit Toolkit"
$form.Size = New-Object System.Drawing.Size(760, 620)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = $form.Size

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = "Fill"
$form.Controls.Add($tabs)

$tabSetup = New-Object System.Windows.Forms.TabPage
$tabSetup.Text = "Tenant Setup"
$tabAudit = New-Object System.Windows.Forms.TabPage
$tabAudit.Text = "Run Audit"
$tabSoftware = New-Object System.Windows.Forms.TabPage
$tabSoftware.Text = "Software Audit"
$tabs.TabPages.Add($tabSetup)
$tabs.TabPages.Add($tabAudit)
$tabs.TabPages.Add($tabSoftware)

# Helper: thread-safe append to a log TextBox from a background runspace
function New-LogWriter($textbox) {
    return {
        param($msg)
        if ($textbox.InvokeRequired) {
            $textbox.Invoke([Action[string]] { param($m)
                    $textbox.AppendText("$m`r`n")
                    $textbox.SelectionStart = $textbox.Text.Length
                    $textbox.ScrollToCaret()
                }, $msg)
        }
        else {
            $textbox.AppendText("$msg`r`n")
        }
    }.GetNewClosure()
}

# =====================================================================
# TAB 1: TENANT SETUP
# =====================================================================

$lblSetupInfo = New-Object System.Windows.Forms.Label
$lblSetupInfo.Text = "Reads Tenants.csv (Name,TenantId) from this folder. Registers an app + grants Graph permissions in each new tenant. You will get one browser sign-in prompt (with MFA) per tenant."
$lblSetupInfo.Location = New-Object System.Drawing.Point(15, 15)
$lblSetupInfo.Size = New-Object System.Drawing.Size(700, 40)
$tabSetup.Controls.Add($lblSetupInfo)

$btnLoadTenants = New-Object System.Windows.Forms.Button
$btnLoadTenants.Text = "Load Tenants.csv"
$btnLoadTenants.Location = New-Object System.Drawing.Point(15, 60)
$btnLoadTenants.Size = New-Object System.Drawing.Size(140, 30)
$tabSetup.Controls.Add($btnLoadTenants)

$gridTenants = New-Object System.Windows.Forms.DataGridView
$gridTenants.Location = New-Object System.Drawing.Point(15, 100)
$gridTenants.Size = New-Object System.Drawing.Size(700, 150)
$gridTenants.ReadOnly = $true
$gridTenants.AutoSizeColumnsMode = "Fill"
$gridTenants.AllowUserToAddRows = $false
$tabSetup.Controls.Add($gridTenants)

$btnRunSetup = New-Object System.Windows.Forms.Button
$btnRunSetup.Text = "Run Setup"
$btnRunSetup.Location = New-Object System.Drawing.Point(15, 260)
$btnRunSetup.Size = New-Object System.Drawing.Size(140, 30)
$btnRunSetup.Enabled = $false
$tabSetup.Controls.Add($btnRunSetup)

$progressSetup = New-Object System.Windows.Forms.ProgressBar
$progressSetup.Location = New-Object System.Drawing.Point(165, 265)
$progressSetup.Size = New-Object System.Drawing.Size(550, 20)
$tabSetup.Controls.Add($progressSetup)

$logSetup = New-Object System.Windows.Forms.TextBox
$logSetup.Multiline = $true
$logSetup.ScrollBars = "Vertical"
$logSetup.ReadOnly = $true
$logSetup.Font = New-Object System.Drawing.Font("Consolas", 9)
$logSetup.Location = New-Object System.Drawing.Point(15, 300)
$logSetup.Size = New-Object System.Drawing.Size(700, 230)
$tabSetup.Controls.Add($logSetup)

$script:TenantsData = @()

$btnLoadTenants.Add_Click({
        if (-not (Test-Path $TenantsCsv)) {
            "Name,TenantId`nCustomer A,customerA.onmicrosoft.com" | Out-File -FilePath $TenantsCsv -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("No Tenants.csv found - a template was created at:`n$TenantsCsv`n`nEdit it with real tenant names/domains, then click Load again.`n`nTenantId = tenant's .onmicrosoft.com domain or GUID, NOT a user email.", "Template Created", "OK", "Information") | Out-Null
            return
        }

        $data = Import-Csv -Path $TenantsCsv | Where-Object { $_.Name -and $_.TenantId }
        $badRows = $data | Where-Object { $_.TenantId -match "^[^@]+@" }
        if ($badRows) {
            $names = ($badRows | ForEach-Object { $_.Name }) -join ", "
            [System.Windows.Forms.MessageBox]::Show("These rows have an email/UPN instead of a tenant domain or GUID: $names`n`nFix Tenants.csv and reload.", "Invalid TenantId", "OK", "Error") | Out-Null
            return
        }

        $script:TenantsData = @($data)
        $gridTenants.DataSource = [System.Collections.ArrayList]@($data)
        $btnRunSetup.Enabled = (@($data).Count -gt 0)
    })

$btnRunSetup.Add_Click({
        $btnRunSetup.Enabled = $false
        $btnLoadTenants.Enabled = $false
        $logSetup.Clear()
        $progressSetup.Minimum = 0
        $progressSetup.Maximum = $script:TenantsData.Count
        $progressSetup.Value = 0

        $writeLog = New-LogWriter $logSetup

        $alreadyDone = @()
        if (Test-Path $AppCredsCsv) { $alreadyDone = (Import-Csv -Path $AppCredsCsv).TenantId }
        $toProcess = $script:TenantsData | Where-Object { $alreadyDone -notcontains $_.TenantId }

        if ($toProcess.Count -eq 0) {
            & $writeLog "All tenants already have an app registered in TenantAuditApps.csv. Nothing to do."
            $btnRunSetup.Enabled = $true
            $btnLoadTenants.Enabled = $true
            return
        }

        $results = @()

        foreach ($tenant in $toProcess) {
            & $writeLog "==== $($tenant.Name) [$($tenant.TenantId)] ===="
            & $writeLog "A browser window will open for interactive admin sign-in (MFA)..."
            [System.Windows.Forms.Application]::DoEvents()

            try {
                Connect-MgGraph -TenantId $tenant.TenantId -Scopes "Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All" -NoWelcome -ErrorAction Stop
                $resolvedTenantId = (Get-MgContext).TenantId

                $existingApp = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ErrorAction SilentlyContinue
                if ($existingApp) {
                    & $writeLog "App already exists, reusing it."
                    $app = $existingApp
                }
                else {
                    $app = New-MgApplication -DisplayName $AppDisplayName -SignInAudience "AzureADMyOrg" -ErrorAction Stop
                    & $writeLog "Created app registration: $($app.AppId)"
                }

                $sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -ErrorAction SilentlyContinue
                if (-not $sp) {
                    $sp = New-MgServicePrincipal -AppId $app.AppId -ErrorAction Stop
                    & $writeLog "Created service principal."
                }

                $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'" -ErrorAction Stop
                foreach ($perm in $RequiredPermissions) {
                    $already = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All |
                    Where-Object { $_.AppRoleId -eq $perm.Id -and $_.ResourceId -eq $graphSp.Id }
                    if (-not $already) {
                        New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $perm.Id -ErrorAction Stop | Out-Null
                        & $writeLog "  Granted: $($perm.Name)"
                    }
                    else {
                        & $writeLog "  Already granted: $($perm.Name)"
                    }
                }

                $secretParams = @{ PasswordCredential = @{ DisplayName = "AuditToolSecret"; EndDateTime = (Get-Date).AddMonths(24) } }
                $secret = Add-MgApplicationPassword -ApplicationId $app.Id -BodyParameter $secretParams -ErrorAction Stop
                & $writeLog "Client secret generated."

                $results += [PSCustomObject]@{
                    TenantName = $tenant.Name; TenantId = $resolvedTenantId
                    ClientId = $app.AppId; ClientSecret = $secret.SecretText
                    SecretExpiry = $secretParams.PasswordCredential.EndDateTime
                }
                & $writeLog "SUCCESS`r`n"

            }
            catch {
                & $writeLog "ERROR: $_`r`n"
            }
            finally {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }

            $progressSetup.Invoke([Action] { $progressSetup.Value++ })
            [System.Windows.Forms.Application]::DoEvents()
        }

        if ($results.Count -gt 0) {
            if (Test-Path $AppCredsCsv) {
                $results | Export-Csv -Path $AppCredsCsv -NoTypeInformation -Append
            }
            else {
                $results | Export-Csv -Path $AppCredsCsv -NoTypeInformation
            }
            & $writeLog "Done. $($results.Count) tenant(s) added to TenantAuditApps.csv"
            & $writeLog "SECURITY: move that file into ITGlue/your vault and delete the local copy."
        }

        $btnRunSetup.Enabled = $true
        $btnLoadTenants.Enabled = $true
    })

# =====================================================================
# TAB 2: RUN AUDIT
# =====================================================================

$lblAuditInfo = New-Object System.Windows.Forms.Label
$lblAuditInfo.Text = "Reads TenantAuditApps.csv from this folder. Authenticates silently (no login) to every tenant and pulls device + license counts."
$lblAuditInfo.Location = New-Object System.Drawing.Point(15, 15)
$lblAuditInfo.Size = New-Object System.Drawing.Size(700, 30)
$tabAudit.Controls.Add($lblAuditInfo)

$btnRunAudit = New-Object System.Windows.Forms.Button
$btnRunAudit.Text = "Run Audit"
$btnRunAudit.Location = New-Object System.Drawing.Point(15, 55)
$btnRunAudit.Size = New-Object System.Drawing.Size(140, 30)
$tabAudit.Controls.Add($btnRunAudit)

$btnOpenReports = New-Object System.Windows.Forms.Button
$btnOpenReports.Text = "Open Reports Folder"
$btnOpenReports.Location = New-Object System.Drawing.Point(165, 55)
$btnOpenReports.Size = New-Object System.Drawing.Size(160, 30)
$tabAudit.Controls.Add($btnOpenReports)

$progressAudit = New-Object System.Windows.Forms.ProgressBar
$progressAudit.Location = New-Object System.Drawing.Point(15, 95)
$progressAudit.Size = New-Object System.Drawing.Size(700, 20)
$tabAudit.Controls.Add($progressAudit)

$gridDevices = New-Object System.Windows.Forms.DataGridView
$gridDevices.Location = New-Object System.Drawing.Point(15, 125)
$gridDevices.Size = New-Object System.Drawing.Size(700, 150)
$gridDevices.ReadOnly = $true
$gridDevices.AutoSizeColumnsMode = "Fill"
$gridDevices.AllowUserToAddRows = $false
$tabAudit.Controls.Add($gridDevices)

$logAudit = New-Object System.Windows.Forms.TextBox
$logAudit.Multiline = $true
$logAudit.ScrollBars = "Vertical"
$logAudit.ReadOnly = $true
$logAudit.Font = New-Object System.Drawing.Font("Consolas", 9)
$logAudit.Location = New-Object System.Drawing.Point(15, 285)
$logAudit.Size = New-Object System.Drawing.Size(700, 245)
$tabAudit.Controls.Add($logAudit)

$btnOpenReports.Add_Click({ Start-Process explorer.exe $WorkDir })

$btnRunAudit.Add_Click({
        if (-not (Test-Path $AppCredsCsv)) {
            [System.Windows.Forms.MessageBox]::Show("TenantAuditApps.csv not found. Run Tenant Setup first.", "Missing File", "OK", "Warning") | Out-Null
            return
        }

        $btnRunAudit.Enabled = $false
        $logAudit.Clear()
        $writeLog = New-LogWriter $logAudit

        $creds = @(Import-Csv -Path $AppCredsCsv | Where-Object { $_.TenantId -and $_.ClientId -and $_.ClientSecret })
        $progressAudit.Minimum = 0
        $progressAudit.Maximum = [Math]::Max($creds.Count, 1)
        $progressAudit.Value = 0

        $deviceResults = @()
        $licenseResults = @()
        $auditDate = Get-Date -Format "yyyy-MM-dd HH:mm"

        foreach ($tenant in $creds) {
            & $writeLog "==== $($tenant.TenantName) ===="
            [System.Windows.Forms.Application]::DoEvents()

            if ($tenant.SecretExpiry) {
                $daysLeft = ([datetime]$tenant.SecretExpiry - (Get-Date)).Days
                if ($daysLeft -lt 30) { & $writeLog "WARNING: client secret expires in $daysLeft day(s)." }
            }

            try {
                $secureSecret = ConvertTo-SecureString $tenant.ClientSecret -AsPlainText -Force
                $credential = New-Object System.Management.Automation.PSCredential($tenant.ClientId, $secureSecret)
                Connect-MgGraph -TenantId $tenant.TenantId -ClientSecretCredential $credential -NoWelcome -ErrorAction Stop

                $entraDevices = Get-MgDevice -All
                $entraStale90 = ($entraDevices | Where-Object { $_.ApproximateLastSignInDateTime -and $_.ApproximateLastSignInDateTime -lt (Get-Date).AddDays(-90) }).Count

                $intuneDevices = Get-MgDeviceManagementManagedDevice -All
                $compliant = ($intuneDevices | Where-Object { $_.ComplianceState -eq "compliant" }).Count
                $nonCompliant = ($intuneDevices | Where-Object { $_.ComplianceState -eq "noncompliant" }).Count

                & $writeLog "Entra ID devices: $($entraDevices.Count) (stale 90d+: $entraStale90)"
                & $writeLog "Intune managed: $($intuneDevices.Count) (compliant: $compliant, non-compliant: $nonCompliant)"

                $deviceResults += [PSCustomObject]@{
                    TenantName = $tenant.TenantName; TenantId = $tenant.TenantId
                    EntraIdDevicesTotal = $entraDevices.Count; EntraIdDevicesStale90d = $entraStale90
                    IntuneManagedTotal = $intuneDevices.Count; IntuneCompliant = $compliant; IntuneNonCompliant = $nonCompliant
                    AuditDate = $auditDate
                }

                $skus = Get-MgSubscribedSku -All
                foreach ($sku in $skus) {
                    $friendly = if ($SkuFriendlyNames.ContainsKey($sku.SkuPartNumber)) { $SkuFriendlyNames[$sku.SkuPartNumber] } else { $sku.SkuPartNumber }
                    $purchased = $sku.PrepaidUnits.Enabled
                    $consumed = $sku.ConsumedUnits
                    & $writeLog "License [$friendly]: $consumed / $purchased used"
                    $licenseResults += [PSCustomObject]@{
                        TenantName = $tenant.TenantName; TenantId = $tenant.TenantId
                        SkuPartNumber = $sku.SkuPartNumber; FriendlyName = $friendly
                        Purchased = $purchased; Consumed = $consumed; Available = $purchased - $consumed
                        AuditDate = $auditDate
                    }
                }
                & $writeLog "SUCCESS`r`n"

            }
            catch {
                & $writeLog "ERROR: $_`r`n"
            }
            finally {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            }

            $progressAudit.Value++
            [System.Windows.Forms.Application]::DoEvents()
        }

        $deviceCsv = Join-Path $WorkDir "DeviceAudit_$(Get-Date -Format 'yyyy-MM-dd').csv"
        $licenseCsv = Join-Path $WorkDir "LicenseAudit_$(Get-Date -Format 'yyyy-MM-dd').csv"
        $deviceResults | Export-Csv -Path $deviceCsv -NoTypeInformation
        $licenseResults | Export-Csv -Path $licenseCsv -NoTypeInformation

        $gridDevices.DataSource = [System.Collections.ArrayList]@($deviceResults)

        $overAllocated = $licenseResults | Where-Object { $_.Available -lt 0 }
        if ($overAllocated) {
            $msg = ($overAllocated | ForEach-Object { "$($_.TenantName): $($_.FriendlyName) over by $(-$_.Available)" }) -join "`n"
            [System.Windows.Forms.MessageBox]::Show("Over-allocated licenses detected:`n`n$msg", "License Warning", "OK", "Warning") | Out-Null
        }

        & $writeLog "Reports saved:`n$deviceCsv`n$licenseCsv"
        $btnRunAudit.Enabled = $true
    })

# =====================================================================
# TAB 3: SOFTWARE AUDIT (ConnectWise Automate)
# =====================================================================

# --- ConnectWise Automate API helper functions ---
# Verify these against your own instance's API docs if calls fail - see the
# note at the top of this file.

function Get-AutomateToken {
    param($ServerUrl, $ClientId, $Username, $Password)
    $body = @{ UserName = $Username; Password = $Password } | ConvertTo-Json
    $headers = @{ ClientID = $ClientId; "Content-Type" = "application/json" }
    $resp = Invoke-RestMethod -Uri "$ServerUrl/cwa/api/v1/apitoken" -Method Post -Body $body -Headers $headers -ErrorAction Stop
    return $resp.AccessToken
}

function Get-AutomateComputers {
    param($ServerUrl, $ClientId, $Token)
    $headers = @{ Authorization = "Bearer $Token"; ClientID = $ClientId }
    $all = @()
    $page = 1
    $pageSize = 500
    do {
        $resp = Invoke-RestMethod -Uri "$ServerUrl/cwa/api/v1/computers?pageSize=$pageSize&page=$page" -Headers $headers -Method Get -ErrorAction Stop
        $batch = @($resp)
        $all += $batch
        $page++
    } while ($batch.Count -eq $pageSize)
    return $all
}

function Get-AutomateSoftware {
    param($ServerUrl, $ClientId, $Token, $ComputerId)
    $headers = @{ Authorization = "Bearer $Token"; ClientID = $ClientId }
    Invoke-RestMethod -Uri "$ServerUrl/cwa/api/v1/computers/$ComputerId/software" -Headers $headers -Method Get -ErrorAction Stop
}

$lblSoftwareInfo = New-Object System.Windows.Forms.Label
$lblSoftwareInfo.Text = "Reads AutomateConfig.csv (ServerUrl,ClientId,Username,Password) from this folder. Pulls installed software per device from ConnectWise Automate so you can see what's installed where for billing/license reconciliation."
$lblSoftwareInfo.Location = New-Object System.Drawing.Point(15, 15)
$lblSoftwareInfo.Size = New-Object System.Drawing.Size(700, 40)
$tabSoftware.Controls.Add($lblSoftwareInfo)

$btnLoadAutomateConfig = New-Object System.Windows.Forms.Button
$btnLoadAutomateConfig.Text = "Load AutomateConfig.csv"
$btnLoadAutomateConfig.Location = New-Object System.Drawing.Point(15, 60)
$btnLoadAutomateConfig.Size = New-Object System.Drawing.Size(170, 30)
$tabSoftware.Controls.Add($btnLoadAutomateConfig)

$gridAutomateConfig = New-Object System.Windows.Forms.DataGridView
$gridAutomateConfig.Location = New-Object System.Drawing.Point(15, 100)
$gridAutomateConfig.Size = New-Object System.Drawing.Size(700, 70)
$gridAutomateConfig.ReadOnly = $true
$gridAutomateConfig.AutoSizeColumnsMode = "Fill"
$gridAutomateConfig.AllowUserToAddRows = $false
$tabSoftware.Controls.Add($gridAutomateConfig)

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = "Billable software filter (comma-separated, optional - leave blank to see everything):"
$lblFilter.Location = New-Object System.Drawing.Point(15, 178)
$lblFilter.Size = New-Object System.Drawing.Size(700, 18)
$tabSoftware.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(15, 198)
$txtFilter.Size = New-Object System.Drawing.Size(700, 22)
$txtFilter.PlaceholderText = "e.g. Adobe Acrobat, AutoCAD, Visio"
$tabSoftware.Controls.Add($txtFilter)

$btnRunSoftwareAudit = New-Object System.Windows.Forms.Button
$btnRunSoftwareAudit.Text = "Run Software Audit"
$btnRunSoftwareAudit.Location = New-Object System.Drawing.Point(15, 230)
$btnRunSoftwareAudit.Size = New-Object System.Drawing.Size(160, 30)
$btnRunSoftwareAudit.Enabled = $false
$tabSoftware.Controls.Add($btnRunSoftwareAudit)

$btnOpenSoftwareReports = New-Object System.Windows.Forms.Button
$btnOpenSoftwareReports.Text = "Open Reports Folder"
$btnOpenSoftwareReports.Location = New-Object System.Drawing.Point(185, 230)
$btnOpenSoftwareReports.Size = New-Object System.Drawing.Size(160, 30)
$tabSoftware.Controls.Add($btnOpenSoftwareReports)
$btnOpenSoftwareReports.Add_Click({ Start-Process explorer.exe $WorkDir })

$progressSoftware = New-Object System.Windows.Forms.ProgressBar
$progressSoftware.Location = New-Object System.Drawing.Point(355, 235)
$progressSoftware.Size = New-Object System.Drawing.Size(360, 20)
$tabSoftware.Controls.Add($progressSoftware)

$gridSoftwareSummary = New-Object System.Windows.Forms.DataGridView
$gridSoftwareSummary.Location = New-Object System.Drawing.Point(15, 270)
$gridSoftwareSummary.Size = New-Object System.Drawing.Size(700, 140)
$gridSoftwareSummary.ReadOnly = $true
$gridSoftwareSummary.AutoSizeColumnsMode = "Fill"
$gridSoftwareSummary.AllowUserToAddRows = $false
$tabSoftware.Controls.Add($gridSoftwareSummary)

$logSoftware = New-Object System.Windows.Forms.TextBox
$logSoftware.Multiline = $true
$logSoftware.ScrollBars = "Vertical"
$logSoftware.ReadOnly = $true
$logSoftware.Font = New-Object System.Drawing.Font("Consolas", 9)
$logSoftware.Location = New-Object System.Drawing.Point(15, 420)
$logSoftware.Size = New-Object System.Drawing.Size(700, 110)
$tabSoftware.Controls.Add($logSoftware)

$script:AutomateConfig = @()

$btnLoadAutomateConfig.Add_Click({
        if (-not (Test-Path $AutomateConfigCsv)) {
            "ServerUrl,ClientId,Username,Password`nhttps://automate.yourdomain.com,00000000-0000-0000-0000-000000000000,apiuser,changeme" |
            Out-File -FilePath $AutomateConfigCsv -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show(
                "No AutomateConfig.csv found - a template was created at:`n$AutomateConfigCsv`n`nEdit it with your real Automate server URL and credentials, then click Load again.`n`nClientId is the API ClientID GUID issued by ConnectWise for API access (not your Automate login) - see your Automate admin console or ConnectWise developer portal.",
                "Template Created", "OK", "Information") | Out-Null
            return
        }

        $data = @(Import-Csv -Path $AutomateConfigCsv | Where-Object { $_.ServerUrl -and $_.ClientId -and $_.Username -and $_.Password })
        if ($data.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("AutomateConfig.csv has no complete rows (needs ServerUrl, ClientId, Username, Password).", "Invalid Config", "OK", "Error") | Out-Null
            return
        }

        $script:AutomateConfig = $data

        # Mask the password before displaying it in the grid
        $displayData = $data | ForEach-Object {
            [PSCustomObject]@{ ServerUrl = $_.ServerUrl; ClientId = $_.ClientId; Username = $_.Username; Password = "********" }
        }
        $gridAutomateConfig.DataSource = [System.Collections.ArrayList]@($displayData)
        $btnRunSoftwareAudit.Enabled = $true
    })

$btnRunSoftwareAudit.Add_Click({
        $btnRunSoftwareAudit.Enabled = $false
        $logSoftware.Clear()
        $writeLog = New-LogWriter $logSoftware

        $filterList = @()
        if ($txtFilter.Text.Trim()) {
            $filterList = $txtFilter.Text.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        }

        $rawResults = @()

        foreach ($instance in $script:AutomateConfig) {
            & $writeLog "==== Connecting to $($instance.ServerUrl) ===="
            [System.Windows.Forms.Application]::DoEvents()

            try {
                $token = Get-AutomateToken -ServerUrl $instance.ServerUrl -ClientId $instance.ClientId -Username $instance.Username -Password $instance.Password
                & $writeLog "Authenticated. Pulling device list..."

                $computers = Get-AutomateComputers -ServerUrl $instance.ServerUrl -ClientId $instance.ClientId -Token $token
                & $writeLog "Found $($computers.Count) device(s). Pulling installed software per device (this can take a while for large fleets)..."

                $progressSoftware.Invoke([Action] { $progressSoftware.Minimum = 0; $progressSoftware.Maximum = [Math]::Max($computers.Count, 1); $progressSoftware.Value = 0 })

                foreach ($pc in $computers) {
                    try {
                        $software = Get-AutomateSoftware -ServerUrl $instance.ServerUrl -ClientId $instance.ClientId -Token $token -ComputerId $pc.Id
                        foreach ($sw in @($software)) {
                            $rawResults += [PSCustomObject]@{
                                ClientName   = $pc.ClientName
                                LocationName = $pc.LocationName
                                DeviceName   = $pc.ComputerName
                                SoftwareName = $sw.Name
                                Version      = $sw.Version
                                Publisher    = $sw.Publisher
                                InstallDate  = $sw.InstallDate
                            }
                        }
                    }
                    catch {
                        & $writeLog "  WARNING: could not pull software for $($pc.ComputerName): $_"
                    }
                    $progressSoftware.Invoke([Action] { $progressSoftware.Value++ })
                    [System.Windows.Forms.Application]::DoEvents()
                }

                & $writeLog "Done with $($instance.ServerUrl)."

            }
            catch {
                & $writeLog "ERROR connecting to $($instance.ServerUrl): $_"
            }
        }

        if ($rawResults.Count -eq 0) {
            & $writeLog "No software data collected - check the errors above."
            $btnRunSoftwareAudit.Enabled = $true
            return
        }

        # Apply the optional billable-software filter to the raw results before summarizing
        $filteredResults = $rawResults
        if ($filterList.Count -gt 0) {
            $filteredResults = $rawResults | Where-Object {
                $swName = $_.SoftwareName
                ($filterList | Where-Object { $swName -like "*$_*" }).Count -gt 0
            }
            & $writeLog "Filter applied: $($filteredResults.Count) of $($rawResults.Count) rows matched [$($filterList -join ', ')]"
        }

        # Per-title summary: how many distinct devices run each software title (for billing counts)
        $summary = $filteredResults | Group-Object SoftwareName | ForEach-Object {
            $devices = $_.Group | Select-Object -ExpandProperty DeviceName -Unique
            [PSCustomObject]@{
                SoftwareName = $_.Name
                DeviceCount  = $devices.Count
                Devices      = ($devices -join "; ")
            }
        } | Sort-Object DeviceCount -Descending

        $dateStamp = Get-Date -Format "yyyy-MM-dd"
        $rawCsv = Join-Path $WorkDir "SoftwareAudit_Raw_$dateStamp.csv"
        $summaryCsv = Join-Path $WorkDir "SoftwareAudit_Summary_$dateStamp.csv"
        $filteredResults | Export-Csv -Path $rawCsv -NoTypeInformation
        $summary | Export-Csv -Path $summaryCsv -NoTypeInformation

        $gridSoftwareSummary.DataSource = [System.Collections.ArrayList]@($summary | Select-Object SoftwareName, DeviceCount)

        & $writeLog "Reports saved:`n$rawCsv`n$summaryCsv"
        $btnRunSoftwareAudit.Enabled = $true
    })

[System.Windows.Forms.Application]::EnableVisualStyles()
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()