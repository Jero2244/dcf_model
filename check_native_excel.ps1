param([string]$OutputDirectory)
$ErrorActionPreference = 'Stop'
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $PSScriptRoot 'outputs'
}
$dcfOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path
$dcfLatest = Get-Content -LiteralPath (Join-Path $dcfOutput 'latest_workbook.json') -Raw | ConvertFrom-Json
$dcfSpec = Get-Content -LiteralPath (Join-Path $dcfOutput 'excel_workbook_definition.json') -Raw | ConvertFrom-Json
$dcfOriginal = (Resolve-Path -LiteralPath $dcfLatest.workbook).Path
$dcfHash = (Get-FileHash -LiteralPath $dcfOriginal -Algorithm SHA256).Hash
$dcfTestDirectory = Join-Path $dcfOutput 'native_excel_check'
New-Item -ItemType Directory -Force -Path $dcfTestDirectory | Out-Null
$dcfCopy = Join-Path $dcfTestDirectory ('native_check_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx')
Copy-Item -LiteralPath $dcfOriginal -Destination $dcfCopy
$dcfEvidence = [ordered]@{
    workbook = $dcfOriginal
    checked_at = [DateTime]::UtcNow.ToString('o')
    engine = 'Microsoft Excel desktop through its native COM object model'
    native_version = $null
    native_build = $null
    formula_comparisons = 0
    comparisons_failed = @()
    behavioral_checks = @()
    formula_errors = @()
    clipped_numeric_cells = @()
    original_preserved = $false
    passed = $false
}
$dcfExcel = $null
$dcfBook = $null
function Assert-DcfClose([object]$Actual, [double]$Expected, [string]$Label) {
    if ($null -eq $Actual -or $Actual -is [string] -or $Actual -is [bool] -or
        [Math]::Abs([double]$Actual - $Expected) -gt (1e-7 * [Math]::Max(1,[Math]::Abs($Expected)))) {
        throw "$Label expected $Expected, received $Actual"
    }
}
function Wait-DcfCalculation {
    $dcfDeadline = [DateTime]::UtcNow.AddSeconds(30)
    while ($dcfExcel.CalculationState -ne 0) {
        if ([DateTime]::UtcNow -gt $dcfDeadline) { throw 'Excel calculation timed out' }
        Start-Sleep -Milliseconds 100
    }
}
try {
    # A private Excel instance and disposable workbook copy isolate these tests
    # from the user's own sessions, settings and delivered workbook.
    $dcfExcel = New-Object -ComObject Excel.Application
    $dcfExcel.Visible = $false
    $dcfExcel.DisplayAlerts = $false
    $dcfEvidence.native_version = $dcfExcel.Version
    $dcfEvidence.native_build = $dcfExcel.Build
    $dcfBook = $dcfExcel.Workbooks.Open($dcfCopy, 0, $false)
    $dcfExcel.Calculation = -4105 # xlCalculationAutomatic
    $dcfExcel.CalculateFullRebuild()
    Wait-DcfCalculation
    foreach ($dcfCheck in $dcfSpec.comparisons) {
        $dcfCell = $dcfBook.Worksheets.Item($dcfCheck.sheet).Range($dcfCheck.cell)
        $dcfActual = $dcfCell.Value2
        try { Assert-DcfClose $dcfActual $dcfCheck.expected "$($dcfCheck.sheet)!$($dcfCheck.cell)" }
        catch { $dcfEvidence.comparisons_failed += $_.Exception.Message }
        if ($dcfCell.Text -match '^#+$') { $dcfEvidence.clipped_numeric_cells += "$($dcfCheck.sheet)!$($dcfCheck.cell)" }
        $dcfEvidence.formula_comparisons++
    }
    if ($dcfEvidence.comparisons_failed.Count) { throw 'Native Excel calculations differ from Python; see evidence file' }
    $dcfInputs = $dcfBook.Worksheets.Item('Assumptions')
    $dcfValuation = $dcfBook.Worksheets.Item('Valuation')
    $dcfScenarios = $dcfBook.Worksheets.Item('Scenarios')
    $dcfForecast = $dcfBook.Worksheets.Item('Forecast')
    $dcfSelected = $dcfInputs.Range('E6').Value2
    $dcfBaseline = $dcfValuation.Range('E23').Value2
    $dcfCaseColumns = @{ Base='E'; Bear='F'; Bull='G' }
    foreach ($dcfCase in @('Base','Bear','Bull')) {
        $dcfInputs.Range('E6').Value2 = $dcfCase
        Wait-DcfCalculation
        Assert-DcfClose $dcfValuation.Range('E23').Value2 $dcfScenarios.Range($dcfCaseColumns[$dcfCase]+'8').Value2 "Case $dcfCase"
        $dcfEvidence.behavioral_checks += "Automatic recalculation: $dcfCase selector"
    }
    $dcfInputs.Range('E6').Value2 = $dcfSelected
    $dcfCaseRows = @{ Base=52; Bear=94; Bull=136 }
    $dcfStart = $dcfCaseRows[$dcfSelected]
    $dcfPrice = $dcfInputs.Range('I'+($dcfStart+3))
    $dcfOriginalPrice = $dcfPrice.Value2
    $dcfPriorCashflow = $dcfForecast.Range('H33').Value2
    $dcfPrice.Value2 = [double]($dcfOriginalPrice+10)
    Wait-DcfCalculation
    if ($dcfValuation.Range('E23').Value2 -le $dcfBaseline) { throw 'Oil-price change did not increase value automatically' }
    Assert-DcfClose $dcfForecast.Range('H33').Value2 $dcfPriorCashflow 'Earlier-year cash flow'
    $dcfEvidence.behavioral_checks += "Later-year oil price changed ADR value from $dcfBaseline to $($dcfValuation.Range('E23').Value2), preserving earlier years"
    $dcfPrice.Value2 = $dcfOriginalPrice

    $dcfWacc = $dcfInputs.Range('E'+($dcfStart+28))
    $dcfOriginalWacc = $dcfWacc.Value2
    $dcfWacc.Value2 = $dcfInputs.Range('E'+($dcfStart+29)).Value2
    Wait-DcfCalculation
    if ($dcfValuation.Range('E23').Value2 -ne 'Review inputs') { throw 'WACC = growth did not block the native Excel valuation' }
    $dcfEvidence.behavioral_checks += 'WACC equal to terminal growth blocks valuation'
    if ($null -eq $dcfOriginalWacc) { $dcfWacc.ClearContents() } else { $dcfWacc.Value2 = $dcfOriginalWacc }

    $dcfPrice = $dcfInputs.Range('E'+($dcfStart+3))
    $dcfOriginalPrice = $dcfPrice.Value2
    $dcfPrice.ClearContents()
    Wait-DcfCalculation
    if ($dcfValuation.Range('E23').Value2 -ne 'Review inputs') { throw 'Missing annual price did not block native Excel valuation' }
    $dcfPrice.Value2 = $dcfOriginalPrice
    $dcfEvidence.behavioral_checks += 'Missing annual price blocks valuation'

    $dcfCapex = $dcfInputs.Range('E'+($dcfStart+17))
    $dcfOriginalCapex = $dcfCapex.Value2
    $dcfCapex.Value2 = [double]0
    Wait-DcfCalculation
    if ($dcfValuation.Range('E29').Value2 -ne 'Capex below production requirement') { throw 'Capex shortfall warning is missing' }
    if ($null -eq $dcfOriginalCapex) { $dcfCapex.ClearContents() } else { $dcfCapex.Value2 = $dcfOriginalCapex }
    $dcfEvidence.behavioral_checks += 'Zero capex override displays investment shortfall'

    Wait-DcfCalculation
    Assert-DcfClose $dcfValuation.Range('E23').Value2 $dcfBaseline 'Restored baseline'
    foreach ($dcfSheet in $dcfBook.Worksheets) {
        $dcfErrors = $null
        try { $dcfErrors = $dcfSheet.UsedRange.SpecialCells(-4123,16) }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2146827284) { throw }
        }
        if ($null -ne $dcfErrors) { $dcfEvidence.formula_errors += "$($dcfSheet.Name)!$($dcfErrors.Address())" }
    }
    if ($dcfEvidence.formula_errors.Count) { throw 'Native Excel contains formula errors' }
    if ($dcfEvidence.clipped_numeric_cells.Count) { throw 'Native Excel contains clipped numeric output' }
    $dcfBook.Save()
    $dcfBook.Close($false)
    $dcfBook = $dcfExcel.Workbooks.Open($dcfCopy,0,$true)
    Wait-DcfCalculation
    Assert-DcfClose $dcfBook.Worksheets.Item('Valuation').Range('E23').Value2 $dcfBaseline 'Saved and reopened native workbook'
    $dcfEvidence.behavioral_checks += 'Restored assumptions survive native Excel save and reopen'
    $dcfEvidence.original_preserved = (Get-FileHash -LiteralPath $dcfOriginal -Algorithm SHA256).Hash -eq $dcfHash
    if (-not $dcfEvidence.original_preserved) { throw 'Delivered workbook changed during the test' }
    $dcfEvidence.passed = $true
    $dcfEvidence | ConvertTo-Json -Depth 10 | Write-Output
}
catch {
    $dcfEvidence.failure = $_.Exception.Message
    throw
}
finally {
    $dcfEvidence | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $dcfOutput 'native_excel_validation.json') -Encoding utf8
    if ($null -ne $dcfBook) { $dcfBook.Close($false) }
    if ($null -ne $dcfExcel) { $dcfExcel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($dcfExcel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
