# ──────────────────────────────────────────────────────────────
#  Generate a colour‑coded licence inventory report (Excel)
# ──────────────────────────────────────────────────────────────

# 1️⃣  Connect to Azure AD
Connect-AzureAD 

# 2️⃣  Get all subscribed SKUs
$licenses = Get-AzureADSubscribedSku

# 3️⃣  Friendly‑name dictionary
$friendlyNames = @{
    "ENTERPRISEPACK"                                     = "Office 365 E3"
    "STANDARDPACK"                                       = "Office 365 E1"
    "EMS"                                                = "Enterprise Mobility + Security E3"
    "EMSSECURITY"                                        = "Enterprise Mobility + Security E5"
    "SPE_E3"                                             = "Microsoft 365 E3"
    "SPE_E5"                                             = "Microsoft 365 E5"
    "INTUNE_A_VL"                                        = "Microsoft Intune Volume Licensing"
    "INTUNE_A_D"                                         = "Microsoft Intune Plan 1"
    "MICROSOFT_BUSINESS"                                 = "Microsoft 365 Business Standard"
    "BUSINESS_PREMIUM"                                   = "Microsoft 365 Business Premium"
    "MICROSOFT_365_COPILOT"                              = "Microsoft 365 Copilot"
    "PROJECTPROFESSIONAL"                                = "Project Plan 3"
    "VISIOONLINE_PLAN2"                                  = "Visio Plan 2"
    "MCOEV"                                              = "Microsoft 365 A1 for devices"
    "EOP_ENTERPRISE"                                     = "Exchange Online Protection (EOP 1)"
    "EXCHANGEARCHIVE"                                    = "Exchange Online Archiving (EOP 2)"
    "AAD_PREMIUM"                                        = "Microsoft Entra ID P1"
    "AAD_PREMIUM_P2"                                     = "Microsoft Entra ID P2"
    "Dynamics_365_for_Operations_Sandbox_Tier1_SKU"      = "Dynamics 365 for Operations Sandbox Tier 1"
    "VISIOCLIENT"                                        = "Visio Professional"
    "STREAM"                                             = "Microsoft Stream"
    "VIRTUAL_AGENT_USL"                                  = "Power Virtual Agents"
    "FLOW_PER_USER"                                      = "Power Automate per user"
    "Microsoft_365_E3_Extra_Features"                    = "Microsoft 365 E3 Extra Features"
    "POWER_BI_PRO"                                       = "Power BI Pro"
    "SPZA_IW"                                            = "Microsoft 365 F3"
    "DYN365_TEAM_MEMBERS"                                = "Dynamics 365 Team Members"
    "WINDOWS_STORE"                                      = "Windows Store"
    "FLOW_FREE"                                          = "Power Automate Free"
    "DYN365_SCM_ATTACH"                                  = "Dynamics 365 SCM Attach"
    "Dynamics_365_for_Operations_Sandbox_Tier2_SKU"      = "Dynamics 365 for Operations Sandbox Tier 2"
    "PROJECT_P1"                                         = "Project Plan 1"
    "Power_Virtual_Agents"                               = "Power Virtual Agents"
    "PROJECTPREMIUM"                                     = "Project Plan 3"
    "Win10_VDA_E3"                                       = "Windows 10 Enterprise E3 (VDA)"
    "PHONESYSTEM_VIRTUALUSER"                            = "Phone System – Virtual User"
    "CCIBOTS_PRIVPREV_VIRAL"                             = "Conversational AI for Bots (Preview)"
    "POWERAPPS_VIRAL"                                    = "Power Apps for Developers (Viral)"
    "EXCHANGESTANDARD"                                   = "Exchange Online Plan 1"
    "Microsoft_Teams_Premium"                            = "Microsoft Teams Premium"
    "M365_DISC0VER_RESPOND"                              = "Microsoft 365 Discover and Respond"
    "POWER_BI_STANDARD"                                  = "Power BI (Standard/Free)"
    "PBI_PREMIUM_PER_USER"                               = "Power BI Premium Per User"
    "POWERAPPS_PER_USER"                                 = "Power Apps per user"
    "PBI_PREMIUM_P2_ADDON"                               = "Power BI Premium Per User Add‑on"
    "Dynamics_365_for_Operations"                        = "Dynamics 365 for Operations"
    "VISIOONLINE_PLAN1"                                  = "Visio Plan 1"
    "EXCHANGEENTERPRISE"                                 = "Exchange Online Plan 2"
    "PROJECT_MADEIRA_PREVIEW_IW_SKU"                     = "Dynamics 365 Business Central Preview"
    "Microsoft_Teams_Rooms_Pro"                          = "Microsoft Teams Rooms Pro"
    "DEFENDER_ENDPOINT_P1"                               = "Microsoft Defender for Endpoint P1"
    "CRM_HYBRIDCONNECTOR"                                = "Dynamics CRM Hybrid Connector"
    "RIGHTSMANAGEMENT_ADHOC"                             = "Azure Rights Management"
    "Microsoft_Teams_Audio_Conferencing_select_dial_out" = "Microsoft Teams Audio Conferencing (Dial‑out Select)"
    "POWERAPPS_DEV"                                      = "Power Apps for Developers"
    "O365_DLP"                                           = "Microsoft 365 Data Loss Prevention (DLP)"
    "M365_INFO_PROTECTION_GOVERNANCE"                    = "Microsoft 365 Information Protection and Governance"
    "DYN365_FINANCE"                                     = "Dynamics 365 Finance"
    "M365_INSIDER_RISK_MANAGEMENT"                       = "Microsoft 365 Insider Risk Management"
    "DYN365_BUSCENTRAL_ESSENTIAL"                        = "Dynamics 365 Business Central Essential"
    "DYN365_HUMAN_RESOURCES_SELF_SERVE"                  = "Dynamics 365 Human Resources Self Serve"
    "DYN365_HUMAN_RESOURCES_ATTACH"                      = "Dynamics 365 Human Resources Attach"
    "DYN365_SCM"                                         = "Dynamics 365 Supply Chain Management (SCM)"
    "DYN365_BUSCENTRAL_PREMIUM"                          = "Dynamics 365 Business Central Premium"
    "DYN365_HUMAN_RESOURCES"                             = "Dynamics 365 Human Resources"
    "DYN365_FINANCE_ATTACH"                              = "Dynamics 365 Finance Attach"
    "DYN365_PROJECT_OPERATIONS"                          = "Dynamics 365 Project Operations"
    "DYN365_BUSCENTRAL_TEAM_MEMBER"                      = "Dynamics 365 Business Central Team Member"
}

# 4️⃣  Build report array
$licenseReport = foreach ($license in $licenses) {
    $sku = $license.SkuPartNumber
    $friendlyName = $friendlyNames[$sku]
    if (-not $friendlyName) { $friendlyName = "Unknown" }

    [PSCustomObject]@{
        "Product Name"      = $sku
        "SkuId"             = $license.SkuId
        "FriendlyName"      = $friendlyName
        "AssignedQuantity"  = $license.ConsumedUnits
        "AvailableQuantity" = $license.PrepaidUnits.Enabled - $license.ConsumedUnits
    }
}

# 5️⃣  Paths & filenames  (ruta actualizada)
$currentDate = Get-Date -Format "dd.MM.yyyy"
$reportDir   = "C:\ INFORM YOUR PATH"
$reportPath  = "$reportDir\LicenseReport_$currentDate.xlsx"
$tempCsvPath = "$env:TEMP\LicenseReport.csv"

# 6️⃣  Export to CSV
$licenseReport | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

# 7️⃣  Open in Excel and format
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($tempCsvPath)
$ws = $workbook.Sheets.Item(1)
$usedRng = $ws.UsedRange

# ── colour palette (decimal BGR) ──
$green        =  5296274   # Office 365 E3
$lightBlue    = 15773696   # Office 365 E1
$orange       =   39423    # Microsoft 365 Copilot
$yellow       =   65535    # Microsoft Intune
$lightPurple  = 15782096   # Microsoft 365 E3   (#D0D0F0)
$dlpColor     = 13158640   # Microsoft 365 DLP  (#F0C8C8)
$infoGovColor = 15790280   # Info Protection & Governance (#C8F0F0)
$headerColor  = 15921906   # light grey

# 8️⃣  Style header
$header = $ws.Rows.Item(1)
$header.Font.Bold = $true
$header.Interior.Color = $headerColor

# 9️⃣  Hide gridlines
$ws.Application.ActiveWindow.DisplayGridlines = $false

# 🔟  Colour‑code FriendlyName (column C)
for ($i = 2; $i -le $usedRng.Rows.Count; $i++) {
    $cell = $ws.Cells.Item($i, 3)
    if (-not $cell) { continue }

    switch ($cell.Text) {
        "Microsoft 365 E3"                               { $cell.Interior.Color = $lightPurple }
        "Office 365 E3"                                  { $cell.Interior.Color = $green }
        "Office 365 E1"                                  { $cell.Interior.Color = $lightBlue }
        "Microsoft 365 Copilot"                          { $cell.Interior.Color = $orange }
        "Microsoft Intune Plan 1"                        { $cell.Interior.Color = $yellow }
        "Microsoft 365 Data Loss Prevention (DLP)"       { $cell.Interior.Color = $dlpColor }
        "Microsoft 365 Information Protection and Governance" { $cell.Interior.Color = $infoGovColor }
    }
}

# 1️⃣1️⃣  Auto‑filter & auto‑fit
$header.AutoFilter()       | Out-Null
$ws.Cells.EntireColumn.AutoFit() | Out-Null

# 1️⃣2️⃣  Legend
$rowStart = $usedRng.Rows.Count + 2
$ws.Cells.Item($rowStart,1).Value2 = "Color Legend:"
$ws.Cells.Item($rowStart,1).Font.Bold = $true

$legend = @{
    "Microsoft 365 E3"                               = $lightPurple
    "Office 365 E3"                                  = $green
    "Office 365 E1"                                  = $lightBlue
    "Microsoft 365 Copilot"                          = $orange
    "Microsoft Intune"                               = $yellow
    "Microsoft 365 Data Loss Prevention (DLP)"       = $dlpColor
    "Microsoft 365 Information Protection and Governance" = $infoGovColor
}

$row = $rowStart + 1
foreach ($entry in $legend.GetEnumerator()) {
    $c = $ws.Cells.Item($row,1)
    $c.Value2 = $entry.Key
    $c.Interior.Color = $entry.Value
    $c.Font.Bold = $true
    $row++
}

# 1️⃣3️⃣  Save, close, clean up
$workbook.SaveAs($reportPath,51)   # 51 = xlOpenXMLWorkbook
$workbook.Close($false)
$excel.Quit()

[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Write-Host "`n✅ Report generated successfully at: $reportPath"
