[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path,

    [string] $ManifestPath,

    [switch] $RequireExcel,

    [switch] $RefreshPivots,

    [ValidateRange(1, 600)]
    [int] $RefreshTimeoutSeconds = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([AllowNull()][object] $Value)

    if ($null -ne $Value -and [System.Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) | Out-Null } catch { }
    }
}

function Get-ExpectedValue {
    param([object] $Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Management.Automation.PSCustomObject] -and $Value.PSObject.Properties['value']) {
        return $Value.value
    }
    return $Value
}

function Assert-CellValue {
    param(
        [Parameter(Mandatory)] [object] $Worksheet,
        [Parameter(Mandatory)] [string] $Address,
        [AllowNull()] [object] $Expected,
        [Parameter(Mandatory)] [string] $Context
    )

    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        $actual = $range.Value2
        $expectedValue = Get-ExpectedValue $Expected

        if ($null -eq $expectedValue) {
            if ($null -ne $actual -and $actual -ne '') {
                throw "$Context $Address expected an empty cell, got '$actual'."
            }
            return
        }

        if ($expectedValue -is [bool]) {
            if ([bool]$actual -ne [bool]$expectedValue) {
                throw "$Context $Address expected '$expectedValue', got '$actual'."
            }
            return
        }

        if ($expectedValue -is [double] -or $expectedValue -is [int] -or $expectedValue -is [long] -or
            $expectedValue -is [decimal]) {
            $actualNumber = [double]$actual
            $expectedNumber = [double]$expectedValue
            if ([Math]::Abs($actualNumber - $expectedNumber) -gt 0.000001) {
                throw "$Context $Address expected '$expectedNumber', got '$actualNumber'."
            }
            return
        }

        if ([string]$actual -ne [string]$expectedValue) {
            throw "$Context $Address expected '$expectedValue', got '$actual'."
        }
    }
    finally {
        Release-ComObject $range
    }
}

function Get-Worksheet {
    param([object] $Workbook, [string] $SheetName)

    try {
        return $Workbook.Worksheets.Item($SheetName)
    }
    catch {
        throw "Worksheet '$SheetName' was not found."
    }
}

function Wait-ExcelReady {
    param(
        [Parameter(Mandatory)] [object] $Excel,
        [Parameter(Mandatory)] [int] $TimeoutSeconds
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        try {
            if ([bool]$Excel.Ready) { return }
        }
        catch { }
        Start-Sleep -Milliseconds 200
    }
    throw "Excel did not become ready within $TimeoutSeconds second(s)."
}

function Get-ExcelFileFormat {
    param(
        [Parameter(Mandatory)] [object] $Workbook,
        [Parameter(Mandatory)] [int] $TimeoutSeconds
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        try {
            $format = [int][double]$Workbook.FileFormat
            if ($format -ne 0) { return $format }
        }
        catch { }
        Start-Sleep -Milliseconds 200
    }
    throw "Excel did not expose a valid FileFormat within $TimeoutSeconds second(s)."
}

function Assert-WorkbookExpectation {
    param(
        [Parameter(Mandatory)] [object] $Workbook,
        [AllowNull()] [object] $Expectation,
        [Parameter(Mandatory)] [string] $FilePath,
        [switch] $CheckPivotRefresh,
        [switch] $RefreshTables
    )

    $context = [IO.Path]::GetFileName($FilePath)
    $expectedSheets = @()
    if ($null -ne $Expectation -and $Expectation.PSObject.Properties['sheets']) {
        $expectedSheets = @($Expectation.sheets)
    }

    if ($expectedSheets.Count -gt 0) {
        $actualNames = @()
        for ($sheetIndex = 1; $sheetIndex -le [int]$Workbook.Worksheets.Count; $sheetIndex++) {
            $worksheet = $null
            try {
                $worksheet = $Workbook.Worksheets.Item($sheetIndex)
                $actualNames += [string]$worksheet.Name
            }
            finally { Release-ComObject $worksheet }
        }
        if (($actualNames -join "`0") -ne ($expectedSheets -join "`0")) {
            throw "$context expected sheets [$($expectedSheets -join ', ')], got [$($actualNames -join ', ')]."
        }
    }

    if ($null -ne $Expectation -and $Expectation.PSObject.Properties['cells']) {
        foreach ($cell in @($Expectation.cells)) {
            $worksheet = $null
            try {
                $worksheet = Get-Worksheet $Workbook ([string]$cell.sheet)
                Assert-CellValue $worksheet ([string]$cell.address) $cell.value $context
            }
            finally { Release-ComObject $worksheet }
        }
    }

    $pivotExpectations = @()
    if ($null -ne $Expectation -and $Expectation.PSObject.Properties['pivots']) {
        $pivotExpectations = @($Expectation.pivots)
    }

    $pivotCount = 0
    for ($sheetIndex = 1; $sheetIndex -le [int]$Workbook.Worksheets.Count; $sheetIndex++) {
        $worksheet = $null
        $pivotTables = $null
        try {
            $worksheet = $Workbook.Worksheets.Item($sheetIndex)
            $pivotTables = $worksheet.PivotTables()
            for ($pivotIndex = 1; $pivotIndex -le [int]$pivotTables.Count; $pivotIndex++) {
                $pivot = $null
                $cache = $null
                try {
                    $pivot = $pivotTables.Item($pivotIndex)
                    $cache = $pivot.PivotCache()
                    $pivotCount++

                    if ($CheckPivotRefresh) {
                        $refreshOnOpen = $false
                        try { $refreshOnOpen = [bool]$cache.RefreshOnFileOpen } catch { }
                        if (-not $refreshOnOpen) {
                            throw "$context pivot '$($pivot.Name)' does not have RefreshOnFileOpen enabled."
                        }
                    }

                    if ($RefreshTables) {
                        $pivot.RefreshTable() | Out-Null
                    }
                }
                finally {
                    Release-ComObject $cache
                    Release-ComObject $pivot
                }
            }
        }
        finally {
            Release-ComObject $pivotTables
            Release-ComObject $worksheet
        }
    }

    if ($CheckPivotRefresh -and $pivotExpectations.Count -gt 0 -and $pivotCount -eq 0) {
        throw "$context contains no pivot tables."
    }

    foreach ($pivotExpectation in $pivotExpectations) {
        $pivotSheet = [string]$pivotExpectation.sheet
        $pivotName = [string]$pivotExpectation.name
        $worksheet = $null
        $pivotTables = $null
        $pivot = $null
        try {
            $worksheet = Get-Worksheet $Workbook $pivotSheet
            $pivotTables = $worksheet.PivotTables()
            $pivot = $pivotTables.Item($pivotName)
            if ($null -eq $pivot) { throw "Pivot '$pivotName' was not found." }

            if ($pivotExpectation.PSObject.Properties['cells']) {
                foreach ($cell in @($pivotExpectation.cells)) {
                    Assert-CellValue $worksheet ([string]$cell.address) $cell.value $context
                }
            }

            if (($RefreshTables -or $CheckPivotRefresh) -and $pivotExpectation.PSObject.Properties['refreshCells']) {
                foreach ($cell in @($pivotExpectation.refreshCells)) {
                    Assert-CellValue $worksheet ([string]$cell.address) $cell.value $context
                }
            }
        }
        finally {
            Release-ComObject $pivot
            Release-ComObject $pivotTables
            Release-ComObject $worksheet
        }
    }
}

function Close-Workbook {
    param([AllowNull()][object] $Workbook)
    if ($null -ne $Workbook) {
        try { $Workbook.Close($false) } catch { }
        Release-ComObject $Workbook
    }
}

$inputPaths = [System.Collections.Generic.List[string]]::new()
foreach ($providedPath in @($Path)) {
    if (-not [string]::IsNullOrWhiteSpace([string]$providedPath)) {
        $inputPaths.Add([string]$providedPath)
    }
}

if ([string]::IsNullOrWhiteSpace($ManifestPath) -eq $false) {
    $manifestFile = [IO.Path]::GetFullPath($ManifestPath)
    if (-not [IO.File]::Exists($manifestFile)) { throw "Manifest not found: $manifestFile" }
    $manifest = Get-Content -LiteralPath $manifestFile -Raw | ConvertFrom-Json
    if ($inputPaths.Count -eq 0 -and $manifest.PSObject.Properties['files']) {
        foreach ($manifestFileEntry in @($manifest.files)) {
            $inputPaths.Add([string]$manifestFileEntry.path)
        }
    }
}
else {
    $manifest = $null
}

if ($inputPaths.Count -eq 0) { throw 'At least one XLSX/XLSB path or a manifest is required.' }

$excelType = $null
try { $excelType = [Type]::GetTypeFromProgID('Excel.Application') } catch { }
if ($null -eq $excelType) {
    if ($RequireExcel) {
        Write-Error 'Microsoft Excel COM is not installed or is not registered.'
        exit 2
    }
    Write-Warning 'Microsoft Excel COM is not installed; validation skipped.'
    exit 0
}

$excelApp = $null
$workbooks = $null
$failures = [System.Collections.Generic.List[string]]::new()

try {
    $excelApp = [Activator]::CreateInstance($excelType)
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.ScreenUpdating = $false
    $excelApp.EnableEvents = $false
    try { $excelApp.AskToUpdateLinks = $false } catch { }
    $workbooks = $excelApp.Workbooks

    foreach ($inputPath in $inputPaths) {
        $fullPath = [IO.Path]::GetFullPath($inputPath)
        $extension = [IO.Path]::GetExtension($fullPath).ToLowerInvariant()
        if ($extension -notin @('.xlsx', '.xlsb')) {
            $failures.Add("${inputPath}: only .xlsx and .xlsb are supported")
            continue
        }
        if (-not [IO.File]::Exists($fullPath)) {
            $failures.Add("${inputPath}: file not found")
            continue
        }

        $expectation = $null
        if ($null -ne $manifest -and $manifest.PSObject.Properties['files']) {
            foreach ($candidate in @($manifest.files)) {
                $candidatePath = [IO.Path]::GetFullPath([string]$candidate.path)
                if ([StringComparer]::OrdinalIgnoreCase.Equals($candidatePath, $fullPath)) {
                    $expectation = $candidate
                    break
                }
            }
        }

        $workbook = $null
        $refreshCopy = $null
        try {
            $workbook = $workbooks.Open($fullPath, 0, $true)
            Wait-ExcelReady $excelApp $RefreshTimeoutSeconds
            $expectedFileFormat = if ($extension -eq '.xlsx') { 51 } else { 50 }
            $actualFileFormat = Get-ExcelFileFormat $workbook $RefreshTimeoutSeconds
            if (-not ($actualFileFormat -eq [int]$expectedFileFormat)) {
                throw "unexpected Excel FileFormat $actualFileFormat, expected $expectedFileFormat"
            }
            Assert-WorkbookExpectation $workbook $expectation $fullPath
            Close-Workbook $workbook
            $workbook = $null

            if ($RefreshPivots -or ($null -ne $expectation -and $expectation.PSObject.Properties['pivots'])) {
                $refreshCopy = Join-Path ([IO.Path]::GetTempPath()) ("SpreadSheetTasks_ExcelCom_" + [Guid]::NewGuid().ToString('N') + $extension)
                Copy-Item -LiteralPath $fullPath -Destination $refreshCopy
                $workbook = $workbooks.Open($refreshCopy, 0, $false)
                Wait-ExcelReady $excelApp $RefreshTimeoutSeconds
                Assert-WorkbookExpectation $workbook $expectation $refreshCopy -CheckPivotRefresh -RefreshTables
                Wait-ExcelReady $excelApp $RefreshTimeoutSeconds
                $workbook.Save()
                Close-Workbook $workbook
                $workbook = $null

                $workbook = $workbooks.Open($refreshCopy, 0, $true)
                Wait-ExcelReady $excelApp $RefreshTimeoutSeconds
                Assert-WorkbookExpectation $workbook $expectation $refreshCopy -CheckPivotRefresh
                Close-Workbook $workbook
                $workbook = $null
            }

            Write-Host ("Excel COM OK: " + [IO.Path]::GetFileName($fullPath))
        }
        catch {
            $failures.Add("$([IO.Path]::GetFileName($fullPath)): $($_.Exception.Message)")
        }
        finally {
            Close-Workbook $workbook
            if ($null -ne $refreshCopy -and [IO.File]::Exists($refreshCopy)) {
                try { Remove-Item -LiteralPath $refreshCopy -Force } catch { }
            }
        }
    }
}
finally {
    if ($null -ne $workbooks) { Release-ComObject $workbooks }
    if ($null -ne $excelApp) {
        try { $excelApp.Quit() } catch { }
        Release-ComObject $excelApp
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($failures.Count -gt 0) {
    Write-Error ("Excel COM validation failed:`n" + ($failures -join "`n"))
    exit 1
}

Write-Host ("Excel COM validation passed: " + $inputPaths.Count + " file(s)")
exit 0
