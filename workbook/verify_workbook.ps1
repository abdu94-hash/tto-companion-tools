<#
.SYNOPSIS
  Opens the companion workbook read-only in Excel, recalculates it fully, and reports:
  the Verification sheet (25 numeric CALC-SPEC vectors), the Examples sheet, and a scan of every
  sheet's used range for spreadsheet error values (#VALUE!, #REF!, #NAME?, #DIV/0!, #N/A, #NUM!).
  Optionally exports Read_Me, Examples and Verification to PDF.

  powershell -NoProfile -ExecutionPolicy Bypass -File verify_workbook.ps1 [-Path workbook.xlsx] [-PdfDir folder]
  Exit code 0 when all numeric vectors pass, all example checks match and no error values are found.
#>
param([string]$Path = '', [string]$PdfDir = '')
$ErrorActionPreference = 'Stop'
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Path) { $Path = Join-Path $Here 'TTO_Companion_Workbook_v1.0.xlsx' }
$xl = $null; $wb = $null; $bad = 0
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($Path, 0, $true)
    $xl.CalculateFull()

    # ---- error scan
    $errTotal = 0; $fTotal = 0
    foreach ($ws in $wb.Worksheets) {
        $nf = 0; $ne = 0
        $ws.Unprotect()   # read-only session; SpecialCells needs an unprotected sheet
        try { $nf = $ws.UsedRange.SpecialCells(-4123).Count } catch {}
        try { $e = $ws.UsedRange.SpecialCells(-4123, 16); $ne = $e.Count; Write-Host ("  errors on " + $ws.Name + ": " + $e.Address(0, 0)) } catch {}
        try { $e2 = $ws.UsedRange.SpecialCells(2, 16); $ne += $e2.Count } catch {}
        $fTotal += $nf; $errTotal += $ne
        Write-Host ("{0,-24} formulas {1,6}   error cells {2}" -f $ws.Name, $nf, $ne)
    }
    Write-Host ("Formula cells: $fTotal; error-value cells: $errTotal")
    if ($errTotal -gt 0) { $bad++ }

    # ---- Verification
    $v = $wb.Worksheets.Item('Verification')
    Write-Host ''
    Write-Host ('Verification overall: ' + $v.Range('C4').Text)
    $r = 8; $pass = 0; $num = 0
    while ($v.Cells.Item($r, 1).Text -match '^V\d\d$') {
        $id = $v.Cells.Item($r, 1).Text; $type = $v.Cells.Item($r, 4).Text; $res = $v.Cells.Item($r, 8).Text
        Write-Host ("  {0} {1,-22} {2,-10} checks {3,2}/{4,-2} {5}" -f $id, $v.Cells.Item($r, 2).Text, $type, $v.Cells.Item($r, 7).Text, $v.Cells.Item($r, 6).Text, $res)
        if ($type -eq 'numeric') { $num++; if ($res -eq 'PASS') { $pass++ } }
        $r++
    }
    Write-Host ("Numeric vectors passing: $pass of $num")
    if ($pass -ne 25 -or $num -ne 25) { $bad++ }
    # failing checks (read the check table in one call)
    $vals = $v.Range('A1:H400').Value2
    $nChecks = 0; $nPass = 0
    for ($r = 1; $r -le 400; $r++) {
        $p = $vals[$r, 8]
        if ($p -is [bool]) {
            $nChecks++; if ($p) { $nPass++ } else { Write-Host ("  FAILED CHECK row $r : " + $vals[$r, 2] + ' ' + $vals[$r, 3] + ' expected [' + $vals[$r, 6] + '] got [' + $vals[$r, 7] + ']') }
        }
    }
    Write-Host ("Individual checks passing: $nPass of $nChecks")

    # ---- Examples
    $e = $wb.Worksheets.Item('Examples')
    Write-Host ''
    Write-Host ('Examples: ' + $e.Range('C4').Text)
    $r = 7; $m = 0; $n = 0
    while ($e.Cells.Item($r, 1).Text -match '^\d+$') {
        $n++; if ($e.Cells.Item($r, 12).Text -eq 'TRUE') { $m++ } else { Write-Host ('  MISMATCH: ' + $e.Cells.Item($r, 2).Text + ' ' + $e.Cells.Item($r, 4).Text + ' chapter [' + $e.Cells.Item($r, 6).Text + '] workbook [' + $e.Cells.Item($r, 10).Text + ']') }
        $r++
    }
    Write-Host ("Example checks matching: $m of $n")
    if ($m -ne $n) { $bad++ }

    if ($PdfDir) {
        foreach ($s in @('Read_Me', 'Examples', 'Verification')) {
            $pdf = Join-Path $PdfDir ("wb_" + $s + ".pdf")
            $wb.Worksheets.Item($s).ExportAsFixedFormat(0, $pdf)
            Write-Host "PDF: $pdf"
        }
    }
    $wb.Close($false); $wb = $null
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {}; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ($bad -gt 0) { Write-Host 'RESULT: FAIL'; exit 1 } else { Write-Host 'RESULT: PASS'; exit 0 }
