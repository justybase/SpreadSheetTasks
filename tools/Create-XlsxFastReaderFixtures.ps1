# Generates an Excel-authored XLSX corpus for the fast XLSX reader tests.
# Run from the repository root:  pwsh -File tools/Create-XlsxFastReaderFixtures.ps1
#
# The script writes test_output/fast-reader-fixtures/*.xlsx and manifest.json.
# The manifest records the values that were put into each cell, so the test suite can
# verify the reader's actual values (not only fast-vs-XmlReader parity).

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $repo 'test_output\fast-reader-fixtures'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Get-ChildItem -Path $outDir -Filter '*.xlsx' -ErrorAction SilentlyContinue | Remove-Item -Force

$excel = $null
$manifestFiles = [System.Collections.Generic.List[object]]::new()

function New-Workbook {
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    return $wb
}

function Save-Workbook($wb, $name, $sheets) {
    $path = Join-Path $outDir $name
    $wb.SaveAs($path, 51) # xlOpenXMLWorkbook
    $wb.Close($false)
    $manifestFiles.Add([ordered]@{ name = $name; sheets = $sheets })
}

function New-Cells {
    # The leading comma stops PowerShell from enumerating the empty list on return.
    return ,([System.Collections.Generic.List[object]]::new())
}

function Add-Cell($cells, $row, $col, $kind, $value) {
    $cells.Add([ordered]@{ row = $row; col = $col; kind = $kind; value = $value })
}

function Set-CellValue($cell, $kind, $value) {
    switch ($kind) {
        'date' {
            $cell.NumberFormat = 'yyyy-mm-dd hh:mm:ss'
            $cell.Value2 = [datetime]$value
        }
        'number' {
            # PowerShell's COM binder handles Int32 poorly for Range.Value2; use a double.
            $cell.Value2 = [double]$value
        }
        'bool' {
            $cell.Value2 = [bool]$value
        }
        default {
            $cell.Value2 = [string]$value
        }
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # ── basic mixed types ───────────────────────────────────────────────────
    $wb = New-Workbook
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Data'
    $cells = New-Cells

    $basic = @(
        @{ r = 1; c = 1; kind = 'string'; value = 'Alpha' },
        @{ r = 1; c = 2; kind = 'number'; value = 42 },
        @{ r = 1; c = 3; kind = 'number'; value = -7 },
        @{ r = 1; c = 4; kind = 'number'; value = 1.5 },
        @{ r = 1; c = 5; kind = 'bool'; value = $true },
        @{ r = 1; c = 6; kind = 'date'; value = '2024-05-06T07:08:09' },
        @{ r = 1; c = 7; kind = 'string'; value = 'A&B <tag> "q" ''a''' },
        @{ r = 1; c = 8; kind = 'string'; value = 'Zażółć gęślą jaźń 😀' },
        @{ r = 2; c = 1; kind = 'string'; value = '  spaced  ' },
        @{ r = 2; c = 2; kind = 'number'; value = 0.0000001 },
        @{ r = 2; c = 3; kind = 'number'; value = -1234567890 },
        @{ r = 2; c = 4; kind = 'number'; value = 1250000000000000 },
        @{ r = 2; c = 5; kind = 'bool'; value = $false },
        @{ r = 2; c = 6; kind = 'date'; value = '1900-01-01T00:00:00' },
        @{ r = 2; c = 7; kind = 'string'; value = "line`nbreak" },
        @{ r = 2; c = 8; kind = 'string'; value = '日本語テキスト' },
        @{ r = 3; c = 1; kind = 'string'; value = ('x'.PadRight(5000, 'x')) },
        @{ r = 3; c = 2; kind = 'number'; value = 123456789 },
        @{ r = 3; c = 3; kind = 'number'; value = 0 },
        @{ r = 3; c = 4; kind = 'number'; value = -0.5 },
        @{ r = 3; c = 5; kind = 'bool'; value = $true },
        @{ r = 3; c = 6; kind = 'date'; value = '2023-12-31T23:59:59' },
        @{ r = 3; c = 7; kind = 'string'; value = 'tab	inside' },
        @{ r = 3; c = 8; kind = 'string'; value = 'plain' }
    )

    foreach ($item in $basic) {
        $cell = $ws.Cells.Item($item.r, $item.c)
        try {
            Set-CellValue $cell $item.kind $item.value
        }
        catch {
            throw "com-basic r=$($item.r) c=$($item.c) kind=$($item.kind) valueType=$($item.value.GetType().FullName): $_"
        }
        Add-Cell $cells $item.r $item.c $item.kind $item.value
    }

    Save-Workbook $wb 'com-basic.xlsx' @([ordered]@{ name = 'Data'; cells = $cells })

    # ── sparse cells and gaps ───────────────────────────────────────────────
    $wb = New-Workbook
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Sparse'
    $cells = New-Cells
    $sparse = @(
        @{ r = 1; c = 1; kind = 'number'; value = 1 },
        @{ r = 1; c = 3; kind = 'number'; value = 3 },
        @{ r = 5; c = 2; kind = 'string'; value = 'five' },
        @{ r = 10; c = 5; kind = 'number'; value = 10.5 },
        @{ r = 20; c = 10; kind = 'date'; value = '2024-02-29T12:00:00' },
        @{ r = 21; c = 1; kind = 'bool'; value = $true }
    )
    foreach ($item in $sparse) {
        $cell = $ws.Cells.Item($item.r, $item.c)
        Set-CellValue $cell $item.kind $item.value
        Add-Cell $cells $item.r $item.c $item.kind $item.value
    }
    Save-Workbook $wb 'com-sparse.xlsx' @([ordered]@{ name = 'Sparse'; cells = $cells })

    # ── wide sheet (columns past Z) ─────────────────────────────────────────
    $wb = New-Workbook
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Wide'
    $cells = New-Cells
    for ($r = 1; $r -le 10; $r++) {
        for ($c = 1; $c -le 80; $c++) {
            $cell = $ws.Cells.Item($r, $c)
            if ($c % 7 -eq 0) {
                $value = "text-$r-$c"
                Set-CellValue $cell 'string' $value
                Add-Cell $cells $r $c 'string' $value
            } else {
                $value = ($r * 1000) + $c
                Set-CellValue $cell 'number' $value
                Add-Cell $cells $r $c 'number' $value
            }
        }
    }
    Save-Workbook $wb 'com-wide.xlsx' @([ordered]@{ name = 'Wide'; cells = $cells })

    # ── formulas, cached values and errors ──────────────────────────────────
    $wb = New-Workbook
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Formulas'
    $cells = New-Cells
    $ws.Cells.Item(1, 1).Formula = '=1+1'
    $ws.Cells.Item(1, 2).Formula = '="text"'
    $ws.Cells.Item(1, 3).Formula = '=TRUE()'
    $ws.Cells.Item(1, 4).Formula = '=1/0'
    $ws.Cells.Item(1, 5).Formula = '=NA()'
    $ws.Cells.Item(1, 6).Formula = '=A1*2'
    $wb.Application.Calculate()
    Add-Cell $cells 1 1 'number' 2
    Add-Cell $cells 1 2 'string' 'text'
    Add-Cell $cells 1 3 'bool' $true
    Add-Cell $cells 1 4 'error' $null
    Add-Cell $cells 1 5 'error' $null
    Add-Cell $cells 1 6 'number' 4
    Save-Workbook $wb 'com-formulas.xlsx' @([ordered]@{ name = 'Formulas'; cells = $cells })

    # ── styles: dates, percentages, currency ────────────────────────────────
    $wb = New-Workbook
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Styles'
    $cells = New-Cells
    $ws.Cells.Item(1, 1).NumberFormat = 'yyyy-mm-dd'
    $ws.Cells.Item(1, 1).Value2 = [datetime]'2024-01-02'
    Add-Cell $cells 1 1 'date' '2024-01-02T00:00:00'
    $ws.Cells.Item(1, 2).NumberFormat = '0.00%'
    $ws.Cells.Item(1, 2).Value2 = 0.1234
    Add-Cell $cells 1 2 'number' 0.1234
    $ws.Cells.Item(1, 3).NumberFormat = '#,##0.00'
    $ws.Cells.Item(1, 3).Value2 = 1234567.89
    Add-Cell $cells 1 3 'number' 1234567.89
    $ws.Cells.Item(2, 1).NumberFormat = 'yyyy-mm-dd hh:mm:ss'
    $ws.Cells.Item(2, 1).Value2 = [datetime]'2025-06-07T08:09:10'
    Add-Cell $cells 2 1 'date' '2025-06-07T08:09:10'
    $ws.Cells.Item(2, 2).Value2 = [double]9007199254740992
    Add-Cell $cells 2 2 'number' 9007199254740992
    Save-Workbook $wb 'com-styles.xlsx' @([ordered]@{ name = 'Styles'; cells = $cells })

    # ── multiple sheets, one hidden ─────────────────────────────────────────
    $wb = New-Workbook
    $first = $wb.Worksheets.Item(1)
    $first.Name = 'First'
    $second = $wb.Worksheets.Add()
    $second.Name = 'Second'
    $third = $wb.Worksheets.Add()
    $third.Name = 'Hidden'
    $third.Visible = 0 # xlSheetHidden
    $cellsFirst = New-Cells
    $cellsSecond = New-Cells
    $cellsHidden = New-Cells
    $first.Cells.Item(1, 1).Value2 = 'first'
    Add-Cell $cellsFirst 1 1 'string' 'first'
    $second.Cells.Item(1, 1).Value2 = [double]2
    Add-Cell $cellsSecond 1 1 'number' 2
    $third.Cells.Item(1, 1).Value2 = [double]3
    Add-Cell $cellsHidden 1 1 'number' 3
    Save-Workbook $wb 'com-multisheet.xlsx' @(
        [ordered]@{ name = 'First'; cells = $cellsFirst },
        [ordered]@{ name = 'Second'; cells = $cellsSecond },
        [ordered]@{ name = 'Hidden'; cells = $cellsHidden }
    )

    # ── large sheet: parity and throughput only (no per-cell manifest) ──────
    $rows = 20000
    $cols = 12
    $wb = New-Workbook
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'Large'
    $values = New-Object 'object[,]' $rows, $cols
    for ($r = 0; $r -lt $rows; $r++) {
        for ($c = 0; $c -lt $cols; $c++) {
            if ($c % 4 -eq 0) { $values[$r, $c] = "row $r col $c" }
            elseif ($c % 4 -eq 1) { $values[$r, $c] = $r * 3 + $c }
            elseif ($c % 4 -eq 2) { $values[$r, $c] = ($r + $c) / 7.0 }
            else { $values[$r, $c] = ($r % 2) -eq 0 }
        }
    }
    $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rows, $cols)).Value2 = $values
    $path = Join-Path $outDir 'com-large.xlsx'
    $wb.SaveAs($path, 51)
    $wb.Close($false)
    $manifestFiles.Add([ordered]@{ name = 'com-large.xlsx'; sheets = @() })

    $manifest = [ordered]@{ generated = (Get-Date).ToString('o'); files = $manifestFiles }
    $manifest | ConvertTo-Json -Depth 12 | Set-Content -Path (Join-Path $outDir 'manifest.json') -Encoding UTF8
    Write-Output "Generated $($manifestFiles.Count) workbooks in $outDir"
}
finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
