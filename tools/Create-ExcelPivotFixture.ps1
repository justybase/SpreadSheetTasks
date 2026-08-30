[CmdletBinding()]
param(
    [string] $OutputDirectory = (Join-Path (Get-Location) 'test_output'),
    [switch] $RequireExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Release-ComObject {
    param([AllowNull()][object] $Value)
    if ($null -ne $Value -and [System.Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) | Out-Null } catch { }
    }
}

function New-DataMatrix {
    $data = New-Object 'object[,]' 4, 2
    $data[0, 0] = 'Category'
    $data[0, 1] = 'Amount'
    $data[1, 0] = 'A'
    $data[1, 1] = 10
    $data[2, 0] = 'B'
    $data[2, 1] = 20
    $data[3, 0] = 'A'
    $data[3, 1] = 5
    return $data
}

function New-PivotWorkbook {
    param(
        [Parameter(Mandatory)] [object] $Excel,
        [Parameter(Mandatory)] [string] $Destination,
        [Parameter(Mandatory)] [int] $FileFormat
    )

    $workbook = $null
    $dataSheet = $null
    $reportSheet = $null
    $sourceRange = $null
    $pivotCaches = $null
    $pivotCache = $null
    $pivotTable = $null
    $categoryField = $null
    $amountField = $null
    try {
        $workbook = $Excel.Workbooks.Add()
        $dataSheet = $workbook.Worksheets.Item(1)
        $dataSheet.Name = 'Data'
        $reportSheet = $workbook.Worksheets.Add()
        $reportSheet.Name = 'Report'

        $sourceRange = $dataSheet.Range('A1').Resize(4, 2)
        $sourceRange.Value2 = New-DataMatrix
        $sourceAddress = 'Data!' + $sourceRange.Address($true, $true, 1)

        $pivotCaches = $workbook.PivotCaches()
        $pivotCache = $pivotCaches.Create(1, $sourceAddress)
        $pivotTable = $pivotCache.CreatePivotTable($reportSheet.Range('A3'), 'SalesPivot')
        $categoryField = $pivotTable.PivotFields('Category')
        $categoryField.Orientation = 1
        $categoryField.Position = 1
        $amountField = $pivotTable.AddDataField($pivotTable.PivotFields('Amount'), 'Sum of Amount', -4157)
        $pivotTable.RefreshTable()

        if ([IO.File]::Exists($Destination)) {
            Remove-Item -LiteralPath $Destination -Force
        }
        $workbook.SaveAs($Destination, $FileFormat)
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch { }
        }
        Release-ComObject $amountField
        Release-ComObject $categoryField
        Release-ComObject $pivotTable
        Release-ComObject $pivotCache
        Release-ComObject $pivotCaches
        Release-ComObject $sourceRange
        Release-ComObject $reportSheet
        Release-ComObject $dataSheet
        Release-ComObject $workbook
    }
}

$outputDirectory = [IO.Path]::GetFullPath($OutputDirectory)
[IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
$excelType = $null
try { $excelType = [Type]::GetTypeFromProgID('Excel.Application') } catch { }
if ($null -eq $excelType) {
    if ($RequireExcel) {
        Write-Error 'Microsoft Excel COM is not installed or is not registered.'
        exit 2
    }
    Write-Warning 'Microsoft Excel COM is not installed; fixture generation skipped.'
    exit 0
}

$excelApp = $null
try {
    $excelApp = [Activator]::CreateInstance($excelType)
    $excelApp.Visible = $false
    $excelApp.DisplayAlerts = $false
    $excelApp.ScreenUpdating = $false
    $excelApp.EnableEvents = $false
    try { $excelApp.AskToUpdateLinks = $false } catch { }

    $xlsxPath = Join-Path $outputDirectory 'updater_pivot_source.xlsx'
    $xlsbPath = Join-Path $outputDirectory 'updater_pivot_source.xlsb'
    New-PivotWorkbook $excelApp $xlsxPath 51
    New-PivotWorkbook $excelApp $xlsbPath 50

    $metadata = [ordered]@{
        xlsx = $xlsxPath
        xlsb = $xlsbPath
        dataSheet = 'Data'
        reportSheet = 'Report'
        pivotName = 'SalesPivot'
        originalRows = 3
    }
    $metadataPath = Join-Path $outputDirectory 'updater_pivot_fixture.json'
    $metadata | ConvertTo-Json | Set-Content -LiteralPath $metadataPath -Encoding UTF8
    Write-Host "Created Excel pivot fixtures in $outputDirectory"
}
finally {
    if ($null -ne $excelApp) {
        try { $excelApp.Quit() } catch { }
        Release-ComObject $excelApp
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
