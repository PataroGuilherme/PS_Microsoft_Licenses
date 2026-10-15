# ---------------------------------------------------------------------------
# LICENSE SUMMARY BY COUNTRY  – Country per Country
# ---------------------------------------------------------------------------

Import-Module AzureAD
Import-Module ImportExcel

# ---------- Conexión Azure AD ----------
try { Get-AzureADTenantDetail -ErrorAction Stop | Out-Null }
catch { Write-Host 'Connecting to Azure AD…'; Connect-AzureAD }

# ---------- Usuarios activos con UsageLocation ----------
Write-Host 'Getting active users…'
$allUsers = Get-AzureADUser -All $true |
            Where-Object { $_.UsageLocation -and $_.AccountEnabled }

# Agrupar usuarios por país para mostrar progreso
$usersByCountry = $allUsers | Group-Object UsageLocation | Sort-Object Name
$totalCountries = $usersByCountry.Count
$countryIndex   = 0

$pivotTable  = @{}
$allLicenses = @()

foreach ($countryGroup in $usersByCountry) {
    $countryIndex++
    $country = $countryGroup.Name

    Write-Progress -Activity 'Creating the license summary...' `
                   -Status  "Processing the country $country ($countryIndex / $totalCountries)" `
                   -PercentComplete ([math]::Round(($countryIndex / $totalCountries)*100,0))
    Write-Host "Processing the country: $country"

    foreach ($user in $countryGroup.Group) {
        try {
            $licenseDetails = Get-AzureADUserLicenseDetail -ObjectId $user.ObjectId
            foreach ($lic in $licenseDetails) {
                $sku        = $lic.SkuPartNumber
                $allLicenses += $sku

                $company = if ($user.CompanyName) { $user.CompanyName.Trim() } else { 'N/A' }
                $grupo   = switch ($true) {
                    { $country -eq 'FR' -and ($company -eq 'Comapny International' -or $company -eq 'SPI') } { 'FR - International + SPI' }
                    { $country -eq 'FR' -and $company -eq 'Company X' }                                      { 'FR - Company X' }
                    { $country -eq 'ES' -and $company -eq 'Company Y' }                                      { 'ES - Company Y' }
                    default { $country }
                }

                if (-not $pivotTable.ContainsKey($grupo)) { $pivotTable[$grupo] = @{} }
                if ($pivotTable[$grupo].ContainsKey($sku)) { $pivotTable[$grupo][$sku]++ }
                else                                      { $pivotTable[$grupo][$sku]  = 1 }
            }
        } catch {
            Write-Warning "Didnt found licenses for $($user.UserPrincipalName): $($_.Exception.Message)"
        }
    }
}

Write-Progress -Activity 'License Summary' -Completed

# ---------- Construir tabla‑resumen ----------
$uniqueLicenses = $allLicenses | Sort-Object -Unique
$pivotOutput    = @()

foreach ($grupo in $pivotTable.Keys) {
    $row   = [ordered]@{ COUNTRY = $grupo }
    $total = 0
    foreach ($sku in $uniqueLicenses) {
        $val = if ($pivotTable[$grupo].ContainsKey($sku)) { $pivotTable[$grupo][$sku] } else { 0 }
        $row[$sku] = $val
        $total += $val
    }
    $row['TOTAL'] = $total
    $pivotOutput += [pscustomobject]$row
}

$pivotOutput = $pivotOutput | Sort-Object COUNTRY

# ---------- Exportar a Excel ----------
$outputPath = 'C:\ Inform Your Path'
if (-not (Test-Path $outputPath)) { New-Item -ItemType Directory -Path $outputPath | Out-Null }

$fecha     = Get-Date -Format 'yyyy-MM-dd'
$pivotPath = Join-Path $outputPath "License_Summary_By_Country_$fecha.xlsx"

$pivotExcel = $pivotOutput | Export-Excel -Path $pivotPath -AutoSize `
    -Title 'License Summary by Country' -WorksheetName 'Summary' `
    -TableName 'PivotTable' -AutoFilter -BoldTopRow -PassThru

$ws = $pivotExcel.Workbook.Worksheets['Summary']

# Encabezado y columna COUNTRY en negrita
$ws.Cells['1:1'].Style.Font.Bold = $true
for ($r = 2; $r -le $ws.Dimension.End.Row; $r++) {
    $ws.Cells[$r,1].Style.Font.Bold = $true
}

# Fila TOTAL en negrita y sumas
$firstDataRow = 3
$lastRow      = $ws.Dimension.End.Row + 1
$lastCol      = $ws.Dimension.End.Column

$ws.Cells["A$lastRow"].Value           = 'TOTAL'
$ws.Cells["A$lastRow"].Style.Font.Bold = $true

for ($c = 2; $c -le $lastCol; $c++) {
    $colLetter = [OfficeOpenXml.ExcelAddress]::GetColumnLetter($c)
    # ---------- CORRECCIÓN: usar subexpresiones $() ----------
    $formula   = "=SUM($($colLetter)$firstDataRow:$($colLetter)$($lastRow-1))"
    $ws.Cells[$lastRow, $c].Formula         = $formula
    $ws.Cells[$lastRow, $c].Style.Font.Bold = $true
}

# Quitar negrita de la fila 3 (B3‑AF3)
for ($c = 2; $c -le 32; $c++) { $ws.Cells[3, $c].Style.Font.Bold = $false }

Close-ExcelPackage $pivotExcel

Write-Host "`n✅ File generated:"
Write-Host "- Summary: $pivotPath"
Write-Host "`n✔️ Process finished."
