<#
.SYNOPSIS
  Opens the companion workbook read-only in Excel, recalculates it fully, and reports:
  the Verification sheet (36 numeric CALC-SPEC 1.1 vectors, V01-V38), the Examples sheet (fictional cases EX-C01,
  EX-C04 and EX-C28 against their home chapters and the calculator's example cases), a scan of every sheet's used
  range for spreadsheet error values (#VALUE!, #REF!, #NAME?, #DIV/0!, #N/A, #NUM!), the absence of a VBA project,
  and the Summary and Export sheets for the chosen example case. Optionally exports Read_Me, Examples, Verification
  and Summary to PDF.

  powershell -NoProfile -ExecutionPolicy Bypass -File verify_workbook.ps1 [-Path workbook.xlsx] [-PdfDir folder] [-ListChecks]
  Exit code 0 when all 36 numeric vectors pass, all example checks match, no error values are found and the file
  has no VBA project.
#>
param([string]$Path = '', [string]$PdfDir = '', [switch]$ListChecks)
$ErrorActionPreference = 'Stop'
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Path) { $Path = Join-Path $Here 'TTO_Companion_Workbook_v1.1.xlsx' }
$ExpectedNumeric = 36
$xl = $null; $wb = $null; $bad = 0
# COM message filter: retry calls that a busy Excel rejects (RPC_E_CALL_REJECTED) instead of failing or misreading
if (-not ('TtoMessageFilter' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}
public class TtoMessageFilter : IOleMessageFilter {
    [DllImport("Ole32.dll")] private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new TtoMessageFilter(), out old); }
    public static void Revoke() { IOleMessageFilter old; CoRegisterMessageFilter(null, out old); }
    int IOleMessageFilter.HandleInComingCall(int t, IntPtr h, int tick, IntPtr info) { return 0; }
    int IOleMessageFilter.RetryRejectedCall(IntPtr h, int tick, int type) { if (tick < 120000) { return 250; } return -1; }
    int IOleMessageFilter.MessagePending(IntPtr h, int tick, int type) { return 2; }
}
'@
}
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') { [TtoMessageFilter]::Register() }

# ---- no VBA project: an .xlsx package must not contain xl/vbaProject.bin
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
try {
    $vba = @($zip.Entries | Where-Object { $_.FullName -match 'vbaProject|\.bin$' -and $_.FullName -notmatch 'printerSettings' })
    $ext = @($zip.Entries | Where-Object { $_.FullName -match 'externalLink|connections\.xml' })
} finally { $zip.Dispose() }
Write-Host ("Package: " + [System.IO.Path]::GetFileName($Path) + "; VBA parts: " + $vba.Count + "; external links or connections: " + $ext.Count)
if ($vba.Count -gt 0 -or $ext.Count -gt 0) { $bad++ }

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($Path, 0, $true)
    $xl.CalculateFull()
    Write-Host ("HasVBProject: " + $wb.HasVBProject)
    if ($wb.HasVBProject) { $bad++ }

    # ---- error scan
    $errTotal = 0; $fTotal = 0
    foreach ($ws in $wb.Worksheets) {
        $nf = 0; $ne = 0
        $ws.Unprotect()   # read-only session
        # every cell of the used range (hidden helper columns included): error values come back as Int32 codes
        $ur = $ws.UsedRange
        $nr = $ur.Rows.Count; $nc = $ur.Columns.Count; $first = $null
        # read row by row (a whole-range Value2 call occasionally comes back as a single value over COM)
        for ($i = 1; $i -le $nr; $i++) {
            $row = $ur.Rows.Item($i)
            $vals = $null; $fms = $null
            for ($try = 0; $try -lt 3 -and -not ($vals -is [array] -or $nc -eq 1); $try++) { $vals = $row.Value2 }
            $fms = $row.Formula
            for ($j = 1; $j -le $nc; $j++) {
                if ($nc -eq 1) { $v = $vals; $f = $fms } else { $v = $vals[1, $j]; $f = $fms[1, $j] }
                if ($f -is [string] -and $f.StartsWith('=')) { $nf++ }
                if ($v -is [int]) { $ne++; if (-not $first) { $first = $ur.Cells.Item($i, $j).Address(0, 0) } }
            }
            if ($nc -gt 1 -and -not ($vals -is [array])) { Write-Host ("  could not read row $i of " + $ws.Name); $bad++ }
        }
        if ($ne -gt 0) { Write-Host ("  errors on " + $ws.Name + ", first at " + $first) }
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
        Write-Host ("  {0} {1,-30} {2,-10} checks {3,2}/{4,-2} {5}" -f $id, $v.Cells.Item($r, 2).Text, $type, $v.Cells.Item($r, 7).Text, $v.Cells.Item($r, 6).Text, $res)
        if ($type -eq 'numeric') { $num++; if ($res -eq 'PASS') { $pass++ } }
        $r++
    }
    Write-Host ("Numeric vectors passing: $pass of $num (expected $ExpectedNumeric)")
    if ($pass -ne $ExpectedNumeric -or $num -ne $ExpectedNumeric) { $bad++ }
    $vals = $v.Range('A1:H600').Value2
    $nChecks = 0; $nPass = 0
    for ($r = 1; $r -le 600; $r++) {
        $p = $vals[$r, 8]
        if ($p -is [bool]) {
            $nChecks++; if ($p) { $nPass++ } else { Write-Host ("  FAILED CHECK row $r : " + $vals[$r, 2] + ' ' + $vals[$r, 3] + ' expected [' + $vals[$r, 6] + '] got [' + $vals[$r, 7] + ']') }
            if ($ListChecks) { Write-Host ("    {0} {1,-44} expected [{2}] got [{3}]" -f $vals[$r, 2], $vals[$r, 3], $vals[$r, 6], $vals[$r, 7]) }
            if (([string]$vals[$r, 7]) -eq '(row not found)') { Write-Host ("  ROW NOT FOUND row $r : " + $vals[$r, 5]); $bad++ }
        }
    }
    Write-Host ("Individual checks passing: $nPass of $nChecks")

    # ---- Examples
    $e = $wb.Worksheets.Item('Examples')
    Write-Host ''
    Write-Host ('Examples: ' + $e.Range('C4').Text)
    $r = 7; $m = 0; $n = 0
    while ($e.Cells.Item($r, 1).Text -match '^\d+$') {
        $n++
        if ($e.Cells.Item($r, 12).Text -eq 'TRUE') { $m++ } else { Write-Host ('  MISMATCH: ' + $e.Cells.Item($r, 2).Text + ' ' + $e.Cells.Item($r, 4).Text + ' chapter [' + $e.Cells.Item($r, 6).Text + '] workbook [' + $e.Cells.Item($r, 10).Text + ']') }
        if ($ListChecks) { Write-Host ("    {0} {1,-60} expected [{2}] workbook [{3}]" -f $e.Cells.Item($r, 2).Text, $e.Cells.Item($r, 4).Text, $e.Cells.Item($r, 8).Text, $e.Cells.Item($r, 10).Text) }
        if ($e.Cells.Item($r, 10).Text -eq '(row not found)') { $bad++ }
        $r++
    }
    Write-Host ("Example checks matching: $m of $n")
    if ($m -ne $n -or $n -eq 0) { $bad++ }

    # ---- Summary and Export (as saved: case EX-C01)
    $s = $wb.Worksheets.Item('Summary'); $x = $wb.Worksheets.Item('Export')
    Write-Host ''
    # printed size on US Letter landscape (the file's page setup), independent of the local default printer
    $sps = $s.PageSetup; $ar = $s.Range($sps.PrintArea)
    $pw = 11 * 72 - $sps.LeftMargin - $sps.RightMargin; $ph = 8.5 * 72 - $sps.TopMargin - $sps.BottomMargin
    $sc = [math]::Min(1.0, $pw / $ar.Width); $pg = [math]::Ceiling($ar.Height * $sc / $ph)
    if ($sps.FitToPagesTall -is [int] -and $pg -gt $sps.FitToPagesTall) { $sc = [math]::Min($sc, $sps.FitToPagesTall * $ph / $ar.Height); $pg = $sps.FitToPagesTall }
    Write-Host ('Summary: ' + $s.Range('G4').Text + ' - ' + $s.Range('H4').Text + '; print area ' + $sps.PrintArea + ' (' + [int]$ar.Width + ' x ' + [int]$ar.Height + ' pt); fit ' + $sps.FitToPagesWide + ' page wide x ' + $sps.FitToPagesTall + ' tall; on US Letter landscape: ' + $pg + ' page(s) at ' + [int]($sc * 100) + '%')
    if ($pg -gt 2) { $bad++ }
    Write-Host ('Export: ' + $x.Range('H4').Text + ' - ' + $x.Range('I4').Text + '; header ' + $x.Range('A7').Text + ' ... ' + $x.Range('AD7').Text)
    $hdr = @(); for ($j = 1; $j -le 30; $j++) { $hdr += $x.Cells.Item(7, $j).Text }
    $want = 'Case_ID,Case_Version,Result_ID,Label,Drug_Event_Link,Exposure_Episode_ID,Event_ID,End_Point,Interval_Value,Interval_Minimum,Interval_Maximum,Unit,Displayed_Value,Calculation_Mode,Anchor_Type,Anchor_Record_ID,Anchor_Description,Sequence_Status,Treatment_Day,Reported_Timing,Constraint_Source,Results,Parameter_Set,Assumptions,Calculation_Version,QC_Flags,Flag_Messages,Limitations,Export_Timestamp,Tool_Version'
    $hdrOk = (($hdr -join ',') -eq $want)
    Write-Host ('Export columns in CALC-SPEC 1.1 Section 8 order: ' + $hdrOk)
    if (-not $hdrOk) { $bad++ }

    if ($PdfDir) {
        foreach ($sn in @('Read_Me', 'Examples', 'Verification', 'Summary')) {
            $pdf = Join-Path $PdfDir ("wb11_" + $sn + ".pdf")
            $wb.Worksheets.Item($sn).ExportAsFixedFormat(0, $pdf)
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
