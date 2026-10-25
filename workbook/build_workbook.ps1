<#
.SYNOPSIS
  Builds TTO_Companion_Workbook_v1.1.xlsx, the companion Excel workbook described in Chapter 23 of
  "Time to Onset in Drug Safety" (Dr. Hafez Selim), through Excel COM automation.

.DESCRIPTION
  TTO Companion Tools v1.1 (CALC-SPEC 1.1). Formulas only: no macros, no VBA, no external links.
  Rules C1-C7 and Q1-Q11, flags and messages follow CALC-SPEC 1.1 (Stage 1 document C), including the
  agreed rules of its Section 9, and match the calculator engine (docs/tto_calc_engine.js).
  The fictional teaching cases (EX-C01, EX-C04, EX-C28) and the verification vectors (VEC-V01 ... VEC-V38)
  are written as reserved, locked rows on the production sheets, so every vector is computed by the
  same column formulas a user's own case uses. The Verification and Examples sheets read those rows.
  Version 1.1 adds a printable one-case Summary sheet and a CSV-ready Export sheet (CALC-SPEC 1.1 Section 8).

  Rebuild (Windows, Microsoft Excel installed):
    powershell -NoProfile -ExecutionPolicy Bypass -File build_workbook.ps1
  Then check:
    powershell -NoProfile -ExecutionPolicy Bypass -File verify_workbook.ps1

  Educational tool. Not validated for regulated production use. Not a dosing calculator.
  Never assigns causality. License: CC BY-NC 4.0.
#>
param([string]$OutPath = '')

$ErrorActionPreference = 'Stop'
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
# COM message filter: retry calls that a busy Excel rejects (RPC_E_CALL_REJECTED) instead of failing the build
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
if (-not $OutPath) { $OutPath = Join-Path $Here 'TTO_Companion_Workbook_v1.1.xlsx' }
$VersionLabel = 'TTO Companion Tools v1.1 (CALC-SPEC 1.1)'
$ToolVersion = 'TTO Companion Tools v1.1'
$SpecVersion = 'CALC-SPEC 1.1'
$VecFile = 'tto_calc_test_vectors_v1.1.json'
# non-ASCII characters (the source stays ASCII-only)
$U_MINUS = [string][char]0x2212; $U_TIMES = [string][char]0x00D7; $U_ARROW = [string][char]0x2192
$U_RSQ = [string][char]0x2019; $U_NDASH = [string][char]0x2013; $U_SQ2 = [string][char]0x00B2; $U_MU = [string][char]0x00B5
$U_SUP = @{ '-' = [string][char]0x207B; '0' = [string][char]0x2070; '1' = [string][char]0x00B9; '2' = [string][char]0x00B2; '3' = [string][char]0x00B3;
            '4' = [string][char]0x2074; '5' = [string][char]0x2075; '6' = [string][char]0x2076; '7' = [string][char]0x2077; '8' = [string][char]0x2078; '9' = [string][char]0x2079 }
$LIMIT_TOOL = 'Educational tool, not validated for regulated production use. It is not a dosing calculator and it does not assess causality: no result is a causality category, a probability, or a statement that an event is related to a drug. Results are arithmetic on the facts entered, under the stated rules and assumptions, and must be read with their flags and mandatory text.'
$LIMIT_C = 'The interval describes the sequence and spacing of recorded times; it is not evidence of causation.'
$LIMIT_C4 = ' The true interval lies between the bounds; no single date is imputed.'
$LIMIT_C7 = ' A recovery interval is not, by itself, evidence of positive dechallenge.'
$RepoUrl = 'https://github.com/abdu94-hash/tto-companion-tools'
$Missing = [System.Reflection.Missing]::Value

# ------------------------------------------------------------------ small helpers
function ColL([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char](65 + $m) + $s; $n = [int][math]::Floor(($n - 1) / 26) }; $s }
function RGBc([int]$r, [int]$g, [int]$b) { $r + 256 * $g + 65536 * $b }
$C_Title = RGBc 31 56 100
$C_InHdr = RGBc 31 78 121
$C_FHdr  = RGBc 89 89 89
$C_In    = RGBc 255 242 204
$C_F     = RGBc 242 242 242
$C_Res   = RGBc 221 235 247
$C_Sep   = RGBc 180 198 231
$C_Warn  = RGBc 252 228 214
$C_Pass  = RGBc 198 239 206
$C_Fail  = RGBc 255 199 206

# Token expansion: {Name} -> column letter of Name in the current table + row number.
$script:TokMap = $null; $script:TokRow = 0
$script:TokEval = [System.Text.RegularExpressions.MatchEvaluator] {
    param($m)
    $k = $m.Groups[1].Value
    if (-not $script:TokMap.ContainsKey($k)) { throw "Unknown token {$k}" }
    [string]$script:TokMap[$k] + [string]$script:TokRow
}
function Expand([string]$tpl, $map, [int]$row) {
    $script:TokMap = $map; $script:TokRow = $row
    [regex]::Replace($tpl, '\{([A-Za-z][A-Za-z0-9_]*)\}', $script:TokEval)
}

# Formula fragments (strings), used inside templates
function MSGX([string]$code, [string]$needs) { 'SUBSTITUTE(IFERROR(INDEX(nMSG_Text,MATCH(' + $code + ',nMSG_Code,0)),""),"[input]",' + $needs + ')' }
# CALC-SPEC 1.1 rule G9: every E-code once, in Section 5 order. $raw is a text containing codes in any order.
$ECODES = @('E_INVALID_DATE', 'E_REVERSED_BOUNDS', 'E_DUPLICATE_ID', 'E_MISSING_LINK', 'E_MISSING_INPUT', 'E_INVALID_UNIT', 'E_NONPOSITIVE_PARAMETER', 'E_OUT_OF_RANGE', 'E_APPLICABILITY_UNCONFIRMED', 'E_CREATININE_STABILITY_UNCONFIRMED', 'E_OVERLAPPING_SEGMENTS')
$ECODES_INT = @('E_INVALID_DATE', 'E_REVERSED_BOUNDS', 'E_DUPLICATE_ID', 'E_MISSING_LINK', 'E_MISSING_INPUT', 'E_INVALID_UNIT')
function ORDE([string]$raw, $codes = $ECODES) {
    'MID(' + (($codes | ForEach-Object { 'IF(ISNUMBER(SEARCH("' + $_ + '",' + $raw + ')),"; ' + $_ + '","")' }) -join '&') + ',3,999)'
}
# " | CODE: message" for every E-code present in $err ([input] filled from $needs for E_MISSING_INPUT)
function EMSGS([string]$err, [string]$needs, $codes = $ECODES) {
    ($codes | ForEach-Object {
        $nd = '""'; if ($_ -eq 'E_MISSING_INPUT') { $nd = 'IF(' + $needs + '="","an input",' + $needs + ')' }
        'IF(ISNUMBER(SEARCH("' + $_ + '",' + $err + '))," | ' + $_ + ': "&' + (MSGX ('"' + $_ + '"') $nd) + ',"")' }) -join '&'
}
function WMSG([string]$cond, [string]$code) { 'IF(' + $cond + '," | ' + $code + ': "&' + (MSGX ('"' + $code + '"') '""') + ',"")' }
function FIRSTCODE([string]$x) { 'IFERROR(LEFT(' + $x + ',FIND(";",' + $x + ')-1),' + $x + ')' }
function MINUSF([string]$x) { 'SUBSTITUTE(' + $x + ',"-","' + $U_MINUS + '")' }
# integer expression -> superscript digits (powers of ten, CALC-SPEC 1.1 Section 7)
function SUPX([string]$x) {
    $s = 'TEXT(' + $x + ',"0")'
    foreach ($k in @('-', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9')) { $s = 'SUBSTITUTE(' + $s + ',"' + $k + '","' + $U_SUP[$k] + '")' }
    $s
}
function OFFMIN([string]$x) { '(IF(LEFT(' + $x + ',1)="-",-1,1)*(VALUE(MID(' + $x + ',2,2))*60+VALUE(MID(' + $x + ',5,2))))' }
function DTXT([string]$x) { 'DAY(' + $x + ')&" "&INDEX(nLIST_Months,MONTH(' + $x + '))&" "&YEAR(' + $x + ')' }
function ISOD([string]$x) { 'YEAR(' + $x + ')&"-"&RIGHT("0"&MONTH(' + $x + '),2)&"-"&RIGHT("0"&DAY(' + $x + '),2)' }
function HHMM([string]$m) { 'RIGHT("0"&INT(MOD(' + $m + ',1440)/60),2)&":"&RIGHT("0"&MOD(' + $m + ',60),2)' }
function SEQ([string]$x) { 'IF(' + $x + '>0,"AFTER",IF(' + $x + '<0,"BEFORE","SIMULTANEOUS"))' }
# Percent to three significant figures as text, without the % sign (CALC-SPEC 1.1 Section 7, as the calculator's
# percent3): below 0.0001 (percent) three significant figures in powers of ten, for example 3.41 x 10^-20.
function P3([string]$x) {
    $e = 'INT(LOG10(ABS(' + $x + ')))'
    $m = 'ROUND(' + $x + '/10^' + $e + ',2)'
    $sci = 'FIXED(IF(ABS(' + $m + ')>=10,' + $m + '/10,' + $m + '),2,TRUE)&" ' + $U_TIMES + ' 10"&' + (SUPX ('(' + $e + '+IF(ABS(' + $m + ')>=10,1,0))'))
    $r = 'ROUND(' + $x + ',MAX(0,2-' + $e + '))'
    $norm = 'FIXED(' + $r + ',MAX(0,2-INT(LOG10(ABS(' + $r + ')))),FALSE)'
    'IF(' + $x + '=0,"0",IF(ABS(' + $x + ')<0.0001,' + $sci + ',' + $norm + '))'
}
# three significant figures, never scientific (the calculator's sig3, used for k)
function SIG3([string]$x) { 'IF(' + $x + '=0,"0",FIXED(' + $x + ',MAX(0,2-INT(LOG10(ABS(' + $x + ')))),TRUE))' }
function MergeH($hs) { $r = [ordered]@{}; foreach ($h in $hs) { foreach ($k in $h.Keys) { $r[$k] = $h[$k] } }; $r }

# ------------------------------------------------------------------ column-spec helpers
# k: 'in' (input, unlocked in user rows) | 'f' (formula, locked)
function CIn($n, $h, $w = 12, $v = $null, $fmt = $null, $g = $null) { @{ n = $n; h = $h; k = 'in'; w = $w; v = $v; fmt = $fmt; g = $g } }
function CF($n, $h, $f, $w = 12, $fmt = $null, $g = $null) { @{ n = $n; h = $h; k = 'f'; w = $w; f = $f; fmt = $fmt; g = $g } }

# Timepoint block (CALC-SPEC Section 2): 10 inputs + 7 derived columns. '@' is replaced by the prefix.
# $lab: the onset block of Events also accepts precision LAB (CALC-SPEC 1.1 C4 laboratory-defined onset): the
# bounds are the date of the last normal test and the date of the first abnormal test (two Milestone_IDs).
function TPBlock([string]$p, [string]$label, [bool]$lab = $false) {
    $labPart = ''
    if ($lab) { $labPart = 'IF({@_Prec}="LAB",{@_LabErr},' }
    $tErr = '=IF(COUNTA({@_Prec}:{@_Upper})=0,"",IF({@_Prec}="","E_MISSING_INPUT",' + $labPart +
      'IF({@_Prec}="TIMESTAMP",IF(OR({@_Date}="",{@_Time}="",{@_Off}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({@_Date})),NOT(ISNUMBER({@_Time}))),"E_INVALID_DATE",IF(OR({@_Date}<>INT({@_Date}),{@_Time}<0,{@_Time}>=1),"E_INVALID_DATE",IF(ISNA(MATCH({@_Off},nLIST_Offset,0)),"E_INVALID_DATE","")))),' +
      'IF({@_Prec}="DATE",IF({@_Date}="","E_MISSING_INPUT",IF(NOT(ISNUMBER({@_Date})),"E_INVALID_DATE",IF({@_Date}<>INT({@_Date}),"E_INVALID_DATE",""))),' +
      'IF({@_Prec}="MONTH",IF(OR({@_Year}="",{@_Month}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({@_Year})),NOT(ISNUMBER({@_Month}))),"E_INVALID_DATE",IF(OR({@_Year}<1900,{@_Year}>9999,{@_Year}<>INT({@_Year}),{@_Month}<1,{@_Month}>12,{@_Month}<>INT({@_Month})),"E_INVALID_DATE",""))),' +
      'IF({@_Prec}="YEAR",IF({@_Year}="","E_MISSING_INPUT",IF(NOT(ISNUMBER({@_Year})),"E_INVALID_DATE",IF(OR({@_Year}<1900,{@_Year}>9999,{@_Year}<>INT({@_Year})),"E_INVALID_DATE",""))),' +
      'IF({@_Prec}="RANGE",IF(OR({@_Lower}="",{@_Upper}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({@_Lower})),NOT(ISNUMBER({@_Upper}))),"E_INVALID_DATE",IF(OR({@_Lower}<>INT({@_Lower}),{@_Upper}<>INT({@_Upper})),"E_INVALID_DATE",IF({@_Lower}>{@_Upper},"E_REVERSED_BOUNDS","")))),' +
      '"E_MISSING_INPUT")))))))' + $(if ($lab) { ')' } else { '' })
    $tHT = '=IF(AND({@_Err}="",{@_Prec}="TIMESTAMP"),1,0)'
    $tOK = '=IF(AND({@_HT}=1,{@_Off}<>"unknown"),1,0)'
    $labLo = ''; $labHi = ''
    if ($lab) { $labLo = 'IF({@_Prec}="LAB",INDEX(nMS_M_Lo,{@_LN_Row}),'; $labHi = 'IF({@_Prec}="LAB",INDEX(nMS_M_Hi,{@_LA_Row}),' }
    $cl = $(if ($lab) { ')' } else { '' })
    $tLo = '=IF(OR({@_Err}<>"",{@_Prec}=""),"",' + $labLo + 'IF(OR({@_Prec}="TIMESTAMP",{@_Prec}="DATE"),{@_Date},IF({@_Prec}="MONTH",DATE({@_Year},{@_Month},1),IF({@_Prec}="YEAR",DATE({@_Year},1,1),{@_Lower}))))' + $cl
    $tHi = '=IF(OR({@_Err}<>"",{@_Prec}=""),"",' + $labHi + 'IF(OR({@_Prec}="TIMESTAMP",{@_Prec}="DATE"),{@_Date},IF({@_Prec}="MONTH",EOMONTH(DATE({@_Year},{@_Month},1),0),IF({@_Prec}="YEAR",DATE({@_Year},12,31),{@_Upper}))))' + $cl
    $tLoc = '=IF({@_HT}=1,{@_Date}*1440+ROUND({@_Time}*1440,0),"")'
    $tUTC = '=IF({@_OK}=1,{@_Loc}-' + (OFFMIN '{@_Off}') + ',"")'
    $cols = @(
        (CIn "${p}_Text" "${label}: original text (as reported)" 26 $null '@' "${label} (timepoint)"),
        (CIn "${p}_Prec" $(if ($lab) { "${label}: precision (LAB = laboratory-defined onset)" } else { "${label}: precision" }) 11 $(if ($lab) { 'list:nLIST_OnsetPrecision' } else { 'list:nLIST_Precision' })),
        (CIn "${p}_Date" "${label}: date (TIMESTAMP, DATE)" 15 'date' 'd mmmm yyyy'),
        (CIn "${p}_Time" "${label}: local time hh:mm (TIMESTAMP only)" 9 'time' 'hh:mm'),
        (CIn "${p}_Off" "${label}: UTC offset or unknown (TIMESTAMP)" 9 'list:nLIST_Offset' '@'),
        (CIn "${p}_Year" "${label}: year (MONTH, YEAR)" 7 'year' '0'),
        (CIn "${p}_Month" "${label}: month 1-12 (MONTH)" 7 'month' '0'),
        (CIn "${p}_Lower" "${label}: earliest date (RANGE)" 15 'date' 'd mmmm yyyy'),
        (CIn "${p}_Upper" "${label}: latest date (RANGE)" 15 'date' 'd mmmm yyyy')
    )
    if ($lab) {
        $cols += @(
            (CIn "${p}_LabNormal" "${label} (LAB): Milestone_ID of the last normal test (lower bound)" 12 $null '@'),
            (CIn "${p}_LabAbnormal" "${label} (LAB): Milestone_ID of the first abnormal test (upper bound)" 12 $null '@'))
    }
    $cols += @(
        (CIn "${p}_Src" "${label}: source (document, record)" 22 $null '@'),
        (CF "${p}_Err" "${label}: check (E-code)" $tErr 12),
        (CF "${p}_HT" "${label}: has time (1/0)" $tHT 6 '0'),
        (CF "${p}_OK" "${label}: offset known (1/0)" $tOK 6 '0'),
        (CF "${p}_Lo" "${label}: lower bound (date)" $tLo 14 'd mmmm yyyy'),
        (CF "${p}_Hi" "${label}: upper bound (date)" $tHi 14 'd mmmm yyyy'),
        (CF "${p}_Loc" "${label}: local clock (minutes since 1900)" $tLoc 11 '0'),
        (CF "${p}_UTC" "${label}: UTC (minutes since 1900)" $tUTC 11 '0')
    )
    if ($lab) {
        $rowF = { param($ref) '=IF({@_' + $ref + '}="",0,IFERROR(MATCH({Case_ID}&"|"&{Case_Version}&"|"&{@_' + $ref + '},nMS_Key,0),0))' }
        $codeF = { param($ref, $row) '=IF({@_Prec}<>"LAB","",IF({@_' + $ref + '}="","E_MISSING_INPUT",IF({@_' + $row + '}=0,"E_MISSING_INPUT",IF(INDEX(nMS_M_Lo,{@_' + $row + '})="",IF(INDEX(nMS_M_Err,{@_' + $row + '})="","E_MISSING_INPUT",INDEX(nMS_M_Err,{@_' + $row + '})),""))))' }
        $rev = 'IF(INDEX(nMS_M_HT,{@_LN_Row})*INDEX(nMS_M_HT,{@_LA_Row})=1,IF(INDEX(nMS_M_OK,{@_LN_Row})*INDEX(nMS_M_OK,{@_LA_Row})=1,INDEX(nMS_M_UTC,{@_LN_Row})>INDEX(nMS_M_UTC,{@_LA_Row}),INDEX(nMS_M_Loc,{@_LN_Row})>INDEX(nMS_M_Loc,{@_LA_Row})),INDEX(nMS_M_Lo,{@_LN_Row})>INDEX(nMS_M_Hi,{@_LA_Row}))'
        $uni = (@('E_INVALID_DATE', 'E_REVERSED_BOUNDS', 'E_MISSING_INPUT', 'E_INVALID_UNIT') | ForEach-Object { 'IF(OR({@_LN_Code}="' + $_ + '",{@_LA_Code}="' + $_ + '"),"; ' + $_ + '","")' }) -join '&'
        $cols += @(
            (CF "${p}_LN_Row" "${label} (LAB): last normal test row" (& $rowF 'LabNormal') 5 '0'),
            (CF "${p}_LA_Row" "${label} (LAB): first abnormal test row" (& $rowF 'LabAbnormal') 5 '0'),
            (CF "${p}_LN_Code" "${label} (LAB): last normal test check" (& $codeF 'LabNormal' 'LN_Row') 10),
            (CF "${p}_LA_Code" "${label} (LAB): first abnormal test check" (& $codeF 'LabAbnormal' 'LA_Row') 10),
            (CF "${p}_LabErr" "${label} (LAB): laboratory-defined onset check (CALC-SPEC 1.1 C4)" ('=IF({@_Prec}<>"LAB","",IF(AND({@_LN_Code}="",{@_LA_Code}=""),IF(' + $rev + ',"E_REVERSED_BOUNDS",""),MID(' + $uni + ',3,99)))') 12)
        )
    }
    foreach ($c in $cols) { if ($c.f) { $c.f = $c.f.Replace('{@_', '{' + $p + '_') } }
    return $cols
}

# ------------------------------------------------------------------ Excel session
$xl = $null; $wb = $null
$script:Infos = [ordered]@{}

# Numbers go through .Formula as invariant text (avoids a PowerShell COM setter type-caching problem with Value2).
function SetV($rng, $v) {
    if ($null -eq $v) { return }
    if ($v -is [string]) { $rng.Value2 = $v }
    else { $rng.Formula = ([double]$v).ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
}
function Write-Val($ws, [string]$addr, $v, $spec) {
    if ($null -eq $v -or ($v -is [string] -and $v -eq '')) { return }
    if ($v -is [string]) {
        if ($v.StartsWith('TEXT:')) { $ws.Range($addr).Value2 = "'" + $v.Substring(5); return }
        if ($spec.v -eq 'date' -and $v -match '^\d{4}-\d{2}-\d{2}$') {
            SetV $ws.Range($addr) ([datetime]::ParseExact($v, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)).ToOADate(); return
        }
        if ($spec.v -eq 'time' -and $v -match '^(\d{2}):(\d{2})$') {
            SetV $ws.Range($addr) (([int]$matches[1] * 60 + [int]$matches[2]) / 1440.0); return
        }
    }
    SetV $ws.Range($addr) $v
}

function Add-Val($rng, $v, [string]$help = '') {
    if (-not $v -and -not $help) { return }
    $val = $rng.Validation
    $val.Delete()
    if (-not $v) { $val.Add(0) }
    elseif ($v -match '^list:(.+)$') { $val.Add(3, 1, 1, '=' + $matches[1]) }
    elseif ($v -match '^listwarn:(.+)$') { $val.Add(3, 3, 1, '=' + $matches[1]) }
    elseif ($v -eq 'date') { $val.Add(4, 1, 1, '1', '2958465') }
    elseif ($v -eq 'time') { $val.Add(5, 1, 1, '0', '0.9993055') }
    elseif ($v -eq 'num') { $val.Add(2, 1, 1, '-1000000000000', '1000000000000') }
    elseif ($v -eq 'year') { $val.Add(1, 1, 1, '1900', '9999') }
    elseif ($v -eq 'month') { $val.Add(1, 1, 1, '1', '12') }
    elseif ($v -eq 'posint') { $val.Add(1, 1, 7, '1') }
    else { throw "Unknown validation $v" }
    $val.IgnoreBlank = $true
    if ($help) { $val.InputMessage = $help.Substring(0, [math]::Min(250, $help.Length)) }
}

# Lays out one table (headers, formats, reserved values). Formulas are written later (Write-Formulas).
function Layout-Table {
    param($ws, [int]$Start, [string]$Title, [string]$Note, $Cols, [int]$NUser, $Examples, $Vectors,
          [string]$Tid, [string]$ExLabel = '', [string]$VecLabel = '', [hashtable]$Help = @{})
    $map = @{}; $spec = @{}
    for ($i = 0; $i -lt $Cols.Count; $i++) { $map[$Cols[$i].n] = ColL ($i + 1); $spec[$Cols[$i].n] = $Cols[$i] }
    $lastCol = ColL $Cols.Count
    $headRow = $Start + 3; $first = $Start + 4
    $blocks = @(); $seps = @(); $r = $first
    $blocks += @{ s = $r; e = $r + $NUser - 1; kind = 'user'; rows = $null }; $r += $NUser
    if ($Examples -and $Examples.Count -gt 0) {
        $seps += @{ row = $r; label = $ExLabel }; $r++
        $blocks += @{ s = $r; e = $r + $Examples.Count - 1; kind = 'example'; rows = $Examples }; $r += $Examples.Count
    }
    if ($Vectors -and $Vectors.Count -gt 0) {
        $seps += @{ row = $r; label = $VecLabel }; $r++
        $blocks += @{ s = $r; e = $r + $Vectors.Count - 1; kind = 'vector'; rows = $Vectors }; $r += $Vectors.Count
    }
    $last = $r - 1

    # titles and headers
    $ws.Range("A$Start").Value2 = $Title
    $ws.Range("A$Start").Font.Bold = $true; $ws.Range("A$Start").Font.Size = 13; $ws.Range("A$Start").Font.Color = $C_Title
    if ($Note) { $ws.Range("A$($Start+1)").Value2 = $Note; $ws.Range("A$($Start+1)").Font.Italic = $true }
    foreach ($c in $Cols) {
        $L = $map[$c.n]
        if ($c.g) { $ws.Range("$L$($Start+2)").Value2 = $c.g; $ws.Range("$L$($Start+2)").Font.Bold = $true }
        $hc = $ws.Range("$L$headRow")
        $hc.Value2 = $c.h
        $hc.Interior.Color = $(if ($c.k -eq 'in') { $C_InHdr } else { $C_FHdr })
        $hc.Font.Color = RGBc 255 255 255; $hc.Font.Bold = $true; $hc.WrapText = $true; $hc.VerticalAlignment = -4160
        $rng = $ws.Range("$L$($first):$L$($last)")
        if ($c.fmt) { $rng.NumberFormat = $c.fmt }
        if ($c.k -eq 'in') {
            $rng.Interior.Color = $C_Res; $rng.Locked = $true
            $u = $ws.Range("$L$($blocks[0].s):$L$($blocks[0].e)")
            $u.Interior.Color = $C_In; $u.Locked = $false
            $hp = ''; if ($Help.ContainsKey($c.n)) { $hp = $Help[$c.n] }
            Add-Val $u $c.v $hp
        } else {
            $rng.Interior.Color = $C_F; $rng.Locked = $true
        }
    }
    $ws.Rows($headRow).RowHeight = [double]66
    foreach ($s in $seps) {
        $sr = $ws.Range("A$($s.row):$lastCol$($s.row)")
        $sr.Interior.Color = $C_Sep; $sr.Locked = $true; $sr.NumberFormat = '@'
        $ws.Range("A$($s.row)").Value2 = $s.label; $ws.Range("A$($s.row)").Font.Bold = $true; $ws.Range("A$($s.row)").Font.Italic = $true
    }
    # reserved values
    foreach ($b in $blocks) {
        if (-not $b.rows) { continue }
        for ($i = 0; $i -lt $b.rows.Count; $i++) {
            $row = $b.s + $i; $h = $b.rows[$i]
            foreach ($k in $h.Keys) {
                if (-not $map.ContainsKey($k)) { throw "Table $Tid ($($ws.Name)) has no column '$k'" }
                if ($spec[$k].k -ne 'in') { throw "Column '$k' in $Tid is not an input" }
                Write-Val $ws ($map[$k] + $row) $h[$k] $spec[$k]
            }
        }
    }
    $info = [pscustomobject]@{ Ws = $ws; Sheet = $ws.Name; Tid = $Tid; Map = $map; Spec = $spec; Cols = $Cols; Blocks = $blocks;
                               Seps = $seps; First = $first; Last = $last; HeadRow = $headRow; LastCol = $lastCol }
    $script:Infos[$Tid] = $info
    return $info
}

function Def-Names($info, [string]$prefix, $only = $null) {
    foreach ($c in $info.Cols) {
        if ($only -and ($only -notcontains $c.n)) { continue }
        $L = $info.Map[$c.n]
        [void]$wb.Names.Add($prefix + $c.n, "='" + $info.Sheet + "'!`$" + $L + "`$" + $info.First + ":`$" + $L + "`$" + $info.Last)
    }
}

function Write-Formulas($info) {
    foreach ($c in $info.Cols) {
        if ($c.k -ne 'f') { continue }
        $L = $info.Map[$c.n]
        foreach ($b in $info.Blocks) {
            $f = Expand $c.f $info.Map $b.s
            try { $info.Ws.Range("$L$($b.s):$L$($b.e)").Formula = $f }
            catch { throw "Formula rejected in $($info.Tid).$($c.n): $f  :: $($_.Exception.Message)" }
        }
    }
}

function Set-Widths($ws, $tables) {
    $w = @{}
    foreach ($t in $tables) { for ($i = 0; $i -lt $t.Cols.Count; $i++) { $cw = [double]$t.Cols[$i].w; if (-not $w.ContainsKey($i) -or $w[$i] -lt $cw) { $w[$i] = $cw } } }
    foreach ($i in $w.Keys) { $ws.Columns($i + 1).ColumnWidth = [double]$w[$i] }
}

# ================================================================== CONTENT DEFINITIONS
$Lists = [ordered]@{
    Precision     = @('TIMESTAMP', 'DATE', 'MONTH', 'YEAR', 'RANGE')
    OnsetPrecision = @('TIMESTAMP', 'DATE', 'MONTH', 'YEAR', 'RANGE', 'LAB')
    Offset        = @('unknown', '-12:00', '-11:00', '-10:00', '-09:30', '-09:00', '-08:00', '-07:00', '-06:00', '-05:00', '-04:00', '-03:30', '-03:00', '-02:00', '-01:00', '+00:00', '+01:00', '+02:00', '+03:00', '+03:30', '+04:00', '+04:30', '+05:00', '+05:30', '+05:45', '+06:00', '+06:30', '+07:00', '+08:00', '+08:45', '+09:00', '+09:30', '+10:00', '+10:30', '+11:00', '+12:00', '+12:45', '+13:00', '+14:00')
    AnchorType    = @('FIRST_ADMIN', 'LAST_DOSE_BEFORE_ONSET', 'DOSE_CHANGE', 'RESTART', 'CONCOMITANT_CHANGE', 'CESSATION', 'OTHER')
    Mode          = @('C1', 'C2', 'C4')
    TD            = @('inclusive', 'not requested')
    ReportedUnit  = @('minutes', 'hours', 'days', 'calendar days', 'weeks', 'months', 'years')
    TimeUnit      = @('minutes', 'hours', 'days')
    DaysOnly      = @('days')
    CLUnit        = @('L/hour', 'mL/minute')
    Status        = @('label fact', 'study observation', 'case measurement', 'model assumption', 'hypothetical example')
    YesNo         = @('Yes', 'No')
    Sex           = @('male', 'female', 'not recorded')
    WeightType    = @('actual', 'ideal', 'adjusted', 'not recorded')
    SCrUnit       = @('mg/dL', ([string][char]0x00B5 + 'mol/L'), 'umol/L')
    AdminKind     = @('DOSE', 'DOSE_CHANGE', 'INFUSION_SEGMENT', 'OTHER')
    RateUnit      = @('mg/hour', 'mL/hour')
    DoseUnit      = @('mg', 'micrograms', 'g', 'mL', 'IU', 'units', 'other (see regimen text)')
    Role          = @('suspect', 'concomitant', 'interacting', 'prophylaxis', 'treatment of the event', 'other')
    MilestoneType = @('BASELINE', 'TEST', 'ORGAN_FUNCTION', 'CONCENTRATION', 'INTERVENTION', 'RECOVERY', 'DIAGNOSIS', 'EXPOSURE_CHANGE', 'OTHER')
    MethodB       = @('FIRST', 'LAST', 'STATED')
    Q6Use         = @('replacement or recovery', 'approach to a new level')
    Fit           = @('compatible', 'atypical but plausible', 'inconsistent with the stated hypothesis', 'indeterminate')
    TimingType    = @('typical onset or latency', 'risk window', 'maximum or distribution reported', 'monitoring or precaution window', 'definitional threshold', 'not a timing statement')
    DocType       = @('USPI', 'SmPC', 'primary study', 'other')
    MissingReason = @('NOT_REPORTED', 'REQUESTED', 'NOT_APPLICABLE', 'UNKNOWN')
    HalfLifeType  = @('terminal elimination', 'effective', 'distribution', 'not applicable')
    MechStatus    = @('established', 'hypothetical')
    Action        = @('withdrawal', 'dose reduction', 're-exposure', 'completed before onset', 'none')
    Priority      = @('high', 'medium', 'low')
    QueryStatus   = @('open', 'requested', 'answered', 'closed')
    ItemType      = @('alternative cause', 'follow-up query')
    Months        = @('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')
}

# CALC-SPEC 1.1 Section 5, verbatim (typographic apostrophe and dash built with [char])
$Flags = @(
    @('E_INVALID_DATE', 'Error', 'A date does not exist or cannot be read', 'This date does not exist. Check the source, or record it as a partial date.'),
    @('E_REVERSED_BOUNDS', 'Error', 'A lower bound is after or above its upper bound', 'The earliest value is after the latest value. Check the order of the bounds.'),
    @('E_DUPLICATE_ID', 'Error', 'Two records share an identifier that must be unique', 'This identifier is already used in this case. Each episode needs its own ID.'),
    @('E_MISSING_LINK', 'Error', 'An interval is requested without a Drug_Event_Link', 'Link this drug episode to this event before calculating an interval.'),
    @('E_MISSING_INPUT', 'Error', 'A required input is blank, including a unit, a gate detail or a Parameter_Status, or the recorded facts do not settle a selection (G10)', 'This result needs [input]. It is left blank rather than estimated.'),
    @('E_INVALID_UNIT', 'Error', 'A unit not listed in Section 6', 'This unit is not supported here. Use one of the listed units.'),
    @('E_NONPOSITIVE_PARAMETER', 'Error', 'A half-life, clearance, volume, interval, rate, dose or BSA is zero or negative', 'This value must be greater than zero.'),
    @('E_OUT_OF_RANGE', 'Error', ('Target fraction outside 0' + $U_NDASH + '1, negative time, or age outside 18' + $U_NDASH + '120'), 'This value is outside the range the equation accepts.'),
    @('E_APPLICABILITY_UNCONFIRMED', 'Error', 'The common gate is not confirmed with all its details, or the Q6 switch is off', ('Confirm the model, the half-life used, the reference time and each input' + $U_RSQ + 's status before a result is shown.')),
    @('E_CREATININE_STABILITY_UNCONFIRMED', 'Error', 'Q11 without stability confirmation', ('Cockcroft' + $U_NDASH + 'Gault assumes stable creatinine. Confirm stability before an estimate is shown.')),
    @('E_OVERLAPPING_SEGMENTS', 'Error', 'Infusion segments overlap', 'Two infusion segments overlap. Check the recorded start and stop times.'),
    @('W_TZ_UNKNOWN', 'Warning', 'A UTC offset is unknown', 'Offset unknown: this interval uses local clock times and could be wrong across a clock change or time zone.'),
    @('W_SAME_DATE_ORDER_UNKNOWN', 'Warning', 'Same date and no times', 'Same calendar date: the order within the day is unknown.'),
    @('W_SEQUENCE_UNCERTAIN', 'Warning', 'Bounds contain zero', 'The available dates allow onset before or after this anchor. Sequence is not established.'),
    @('W_ADMIN_SAME_DATE_AS_ONSET', 'Warning', 'The selected last dose shares the onset date without times', 'This dose was on the onset date; whether it came before onset is unknown.'),
    @('W_REPORTED_TIMING_CONFLICT', 'Warning', 'Reported timing does not overlap the computed interval', 'The reported timing conflicts with the recorded dates. Both are kept; query the source.'),
    @('W_INTERVENTION_BEFORE_RECOVERY', 'Warning', 'A recorded intervention can lie between cessation and recovery, ends included', 'Treatment was given before recovery, so recovery cannot be attributed to withdrawal alone.'),
    @('W_BIOLOGIC_MODEL_CAUTION', 'Warning', 'Every Q7 result, and Q8 for a recorded monoclonal antibody', 'Antibody disposition can be nonlinear or time-varying; this simple model may not apply.'),
    @('W_SEX_NOT_RECORDED', 'Warning', 'Q11 without recorded sex', 'Sex not recorded: both estimates are shown.'),
    @('W_PARTIAL_BSA', 'Warning', 'Q10 with a missing BSA', 'BSA missing for some administrations: the value shown is a lower bound.'),
    @('I_TREATMENT_DAY_UNDEFINED', 'Info', 'C2 is negative', 'No treatment-day number exists before the first dose.'),
    @('I_APPROXIMATE_NOT_COMPARED', 'Info', 'Reported timing has no bounds, or units that C6 does not compare', 'Approximate timing is kept as reported and not converted.'),
    @('I_SCENARIO_NOT_CI', 'Info', 'Every Q8 result', 'Scenario range from assumed parameters, not a confidence interval.'),
    @('I_TEACHING_DEMONSTRATION', 'Info', 'Every Q6 result', 'Teaching demonstration with hypothetical values.')
)

# ---------------- record helpers for reserved rows
function TS($p, $date, $time, $off, $text, $src) { @{ "${p}_Prec" = 'TIMESTAMP'; "${p}_Date" = $date; "${p}_Time" = $time; "${p}_Off" = $off; "${p}_Text" = $text; "${p}_Src" = $src } }
function DT($p, $date, $text, $src) { @{ "${p}_Prec" = 'DATE'; "${p}_Date" = $date; "${p}_Text" = $text; "${p}_Src" = $src } }
function RG($p, $lo, $hi, $text, $src) { @{ "${p}_Prec" = 'RANGE'; "${p}_Lower" = $lo; "${p}_Upper" = $hi; "${p}_Text" = $text; "${p}_Src" = $src } }
function CV($case, $ver) { @{ Case_ID = $case; Case_Version = $ver } }
function IVR($case, $ver, $id, $link, $atype, $adesc, $arec, $end, $mode, $extra = @{}) {
    MergeH @((CV $case $ver), @{ Interval_ID = $id; Link_ID = $link; Anchor_Type = $atype; Anchor_Description = $adesc; Anchor_Record_ID = $arec; End_Record_ID = $end; Calculation_Mode = $mode }, $extra)
}
function GY($type, $pop, $src, $ref) { @{ G_Model = 'Yes'; G_HL = 'Yes'; G_HL_Type = $type; G_HL_Pop = $pop; G_HL_Src = $src; G_Ref = 'Yes'; G_RefText = $ref; G_Status = 'Yes' } }
function PV($n, $v, $st, $src) { @{ $n = $v; "${n}_Status" = $st; "${n}_Source" = $src } }

$EX_LABEL = 'Reserved, locked rows: fictional teaching cases EX-C01 (case version 2), EX-C04 (version 2) and EX-C28 (version 2), read by the Examples sheet. Copy rows to the yellow area with a new Case_ID to experiment.'
$VEC_LABEL = 'Reserved, locked rows: CALC-SPEC 1.1 verification vectors (VEC-V01 ... VEC-V38), read by the Verification sheet. They use the same column formulas as your own rows.'

# ---------------- Cases
$CasesEx = @(
    @{ Case_ID = 'EX-C01'; Case_Version = 2; Received_Date = '2026-01-14'; Setting = 'Hospital infusion (fictional)'; Fictional = 'Yes'; Summary = 'Urticaria, wheeze and hypotension 15 minutes after the start of a first Drug A infusion (Chapter 3, Section 3.7).'; Version_Note = 'Version 2 (follow-up 14 January 2026) adds Drug B, the antiseptic, pump rate, first exposure, diagnosis and tryptase sample; version 1 facts unchanged.' },
    @{ Case_ID = 'EX-C04'; Case_Version = 2; Received_Date = '2026-07-20'; Setting = 'Primary care and emergency department (fictional)'; Fictional = 'Yes'; Summary = 'Nausea and symptomatic bradycardia on digoxin after renal deterioration (Chapter 11, Section 11.8; Chapter 23, Section 23.11).'; Version_Note = 'Version 2 (GP letter received 20 July 2026) adds dose, last dose time, ibuprofen, creatinine series, digoxin level with sample time and IV fluids.' },
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Received_Date = '2026-06-15'; Setting = 'Hematology and hepatology (fictional)'; Fictional = 'Yes'; Summary = 'HBV reactivation months after completing rituximab (Chapter 20, Section 20.4; Chapter 23, Section 23.12).'; Version_Note = 'Version 2 (received 15 June 2026) adds cycle dates, serology, prophylaxis, HBV DNA monitoring, prednisolone course and B-cell/IgG results.' }
)
$CasesVec = @()
foreach ($v in @('V01', 'V02', 'V03', 'V04', 'V05', 'V06', 'V07', 'V08', 'V09', 'V10', 'V11', 'V12', 'V13', 'V15', 'V16', 'V17', 'V18', 'V19', 'V20', 'V21', 'V22', 'V23', 'V24', 'V25', 'V26', 'V27', 'V28', 'V29', 'V30', 'V31', 'V32', 'V33', 'V34', 'V35', 'V36', 'V37', 'V38')) {
    $CasesVec += @{ Case_ID = "VEC-$v"; Case_Version = 1; Fictional = 'Yes'; Summary = "CALC-SPEC 1.1 verification vector $v ($VecFile)"; Version_Note = 'Vector inputs as in the JSON file' }
    if ($v -eq 'V12') { $CasesVec += @{ Case_ID = 'VEC-V12'; Case_Version = 2; Fictional = 'Yes'; Summary = 'V12 follow-up: onset moves from 10 to 8 May 2026'; Version_Note = 'Version 1 row is kept; version 2 is added (V12)' } }
}

# ---------------- Exposures
$ExpEx = @(
    (MergeH @((CV 'EX-C01' 2), @{ Drug_ID = 'DRUG-A'; Exposure_Episode_ID = 'EX-A'; Product = 'Drug A (placeholder)'; Role = 'suspect'; Route = 'intravenous infusion'; Regimen_Text = 'First exposure to Drug A (version 2)' },
        (TS 'S' '2026-01-10' '09:00' '-05:00' '09:00 Drug A infusion commenced' 'Nursing chart; medication administration record'),
        (TS 'T' '2026-01-10' '09:16' '-05:00' '09:16 infusion stopped, doctor called' 'Nursing chart'))),
    (MergeH @((CV 'EX-C01' 2), @{ Drug_ID = 'DRUG-B'; Exposure_Episode_ID = 'EX-B'; Product = 'Drug B (placeholder)'; Role = 'concomitant'; Route = 'intravenous infusion, same cannula'; Regimen_Text = 'Previous courses, dates unknown (version 2)' },
        (TS 'S' '2026-01-10' '08:20' '-05:00' 'Drug B IV given earlier this morning via same cannula' 'Follow-up, version 2'),
        (TS 'T' '2026-01-10' '08:50' '-05:00' 'Drug B infusion ended 08:50' 'Follow-up, version 2'))),
    (MergeH @((CV 'EX-C01' 2), @{ Drug_ID = 'ANTISEPTIC'; Exposure_Episode_ID = 'EX-AS'; Product = 'Skin antiseptic (concomitant material)'; Role = 'concomitant'; Route = 'topical, at cannulation' },
        (TS 'S' '2026-01-10' '08:15' '-05:00' 'Antiseptic applied at cannulation' 'Follow-up, version 2'))),
    (MergeH @((CV 'EX-C04' 2), @{ Drug_ID = 'DIGOXIN'; Exposure_Episode_ID = 'EX-DIG'; Product = 'Digoxin tablets'; Role = 'suspect'; Route = 'oral'; Dose = 125; Dose_Unit = 'micrograms'; Regimen_Text = '125 micrograms each morning since 1 January 2026, dose never changed' },
        (DT 'S' '2026-01-01' 'digoxin 125 micrograms each morning since 1 January 2026' 'GP letter (version 2)'),
        (DT 'T' '2026-07-05' 'Digoxin and ibuprofen withheld (5 July)' 'Emergency department note'))),
    (MergeH @((CV 'EX-C04' 2), @{ Drug_ID = 'IBUPROFEN'; Exposure_Episode_ID = 'EX-IBU'; Product = 'Ibuprofen'; Role = 'interacting'; Route = 'oral'; Dose = 400; Dose_Unit = 'mg'; Regimen_Text = '400 mg three times a day for back pain' },
        (DT 'S' '2026-06-26' 'Bought ibuprofen ... on 26 June' 'GP letter (version 2)'),
        (DT 'T' '2026-07-04' 'last taken 4 July' 'GP letter (version 2)'))),
    (MergeH @((CV 'EX-C04' 2), @{ Drug_ID = 'BETA-BLOCKER'; Exposure_Episode_ID = 'EX-BB'; Product = 'Beta-blocker (product not named)'; Role = 'concomitant'; Route = 'oral'; Regimen_Text = 'Long term; dose unchanged since 2024'; Missing_Reason = 'NOT_REPORTED'; Notes = 'Start date not reported; additive AV-conduction effect (pharmacodynamic)' })),
    (MergeH @((CV 'EX-C28' 2), @{ Drug_ID = 'RITUXIMAB'; Exposure_Episode_ID = 'EX-RTX'; Product = 'Rituximab (NHL)'; Role = 'suspect'; Route = 'intravenous'; Regimen_Text = 'R-Regimen X x 6 cycles every 21 days; doses as recorded in the source' },
        (DT 'S' '2024-11-18' 'R-Regimen X x 6, 18/11/2024 to 03/03/2025' 'Hematology summary'),
        (DT 'T' '2025-03-03' 'last infusion 03/03/2025' 'Hematology summary'))),
    (MergeH @((CV 'EX-C28' 2), @{ Drug_ID = 'REGIMEN-X'; Exposure_Episode_ID = 'EX-RX'; Product = 'Regimen X (placeholder; includes a glucocorticoid)'; Role = 'concomitant'; Route = 'combination chemotherapy' },
        (DT 'S' '2024-11-18' 'same cycles as rituximab' 'Hematology summary'),
        (DT 'T' '2025-03-03' 'same cycles as rituximab' 'Hematology summary'))),
    (MergeH @((CV 'EX-C28' 2), @{ Drug_ID = 'ANTIVIRAL-A'; Exposure_Episode_ID = 'EX-AVA'; Product = 'Antiviral A (placeholder), prophylaxis'; Role = 'prophylaxis'; Route = 'oral' },
        (DT 'S' '2024-11-11' 'Antiviral A prophylaxis 11/11/2024 to 03/09/2025' 'Hematology summary'),
        (DT 'T' '2025-09-03' 'prophylaxis to 03/09/2025' 'Hematology summary'))),
    (MergeH @((CV 'EX-C28' 2), @{ Drug_ID = 'PREDNISOLONE'; Exposure_Episode_ID = 'EX-PRED'; Product = 'Prednisolone, 10-day oral course'; Role = 'concomitant'; Route = 'oral'; Missing_Reason = 'NOT_REPORTED'; Notes = 'Dose not reported' },
        (DT 'S' '2026-03-02' '10-day oral prednisolone course 02-11/03/2026' 'Hepatology letter'),
        (DT 'T' '2026-03-11' '10-day oral prednisolone course 02-11/03/2026' 'Hepatology letter'))),
    (MergeH @((CV 'EX-C28' 2), @{ Drug_ID = 'ANTIVIRAL-A'; Exposure_Episode_ID = 'EX-AVA2'; Product = 'Antiviral A, treatment of the event'; Role = 'treatment of the event'; Route = 'oral' },
        (DT 'S' '2026-05-27' 'Antiviral A restarted 27/05/2026' 'Hepatology letter')))
)
$V = $VecFile
function VX($case, $ver, $ep, $extra) { MergeH @((CV $case $ver), @{ Drug_ID = 'D1'; Exposure_Episode_ID = $ep; Product = 'Vector drug' }, $extra) }
$ExpVec = @(
    (VX 'VEC-V01' 1 'EX1' (TS 'S' '2026-01-10' '09:00' '-05:00' '2026-01-10T09:00-05:00' "$V V01")),
    (VX 'VEC-V02' 1 'EX1' (DT 'S' '2026-01-01' '2026-01-01' "$V V02")),
    (VX 'VEC-V03' 1 'EX1' (DT 'S' '2026-01-05' '2026-01-05' "$V V03")),
    (VX 'VEC-V04' 1 'EX1' (DT 'S' '2024-02-28' '2024-02-28' "$V V04")),
    (VX 'VEC-V05' 1 'EX1' (RG 'S' '2026-01-01' '2026-01-10' 'start 1-10 January 2026' "$V V05")),
    (VX 'VEC-V06' 1 'EX1' (RG 'S' '2026-01-01' '2026-01-10' 'start 1-10 January 2026' "$V V06")),
    (VX 'VEC-V07' 1 'EX1' (TS 'S' '2026-03-01' '08:00' '-05:00' 'first dose' "$V V07")),
    (VX 'VEC-V08' 1 'EX1' (MergeH @((DT 'S' '2026-01-01' 'first administration' "$V V08"), (DT 'T' '2026-01-20' 'stop' "$V V08"), (DT 'R' '2026-02-10' 'restart' "$V V08")))),
    (VX 'VEC-V09' 1 'EX1' (TS 'S' '2026-11-01' '01:30' '-04:00' '2026-11-01T01:30-04:00' "$V V09")),
    (VX 'VEC-V10' 1 'EX1' (TS 'S' '2026-03-08' '01:30' 'unknown' '2026-03-08T01:30 (offset unknown)' "$V V10")),
    (VX 'VEC-V11' 1 'EX1' (TS 'S' '2026-04-01' '08:00' '-05:00' 'dose 08:00' "$V V11")),
    (VX 'VEC-V12' 1 'EX1' (DT 'S' '2026-05-01' '2026-05-01' "$V V12")),
    (VX 'VEC-V12' 2 'EX1' (DT 'S' '2026-05-01' '2026-05-01' "$V V12")),
    (VX 'VEC-V13' 1 'EX1' (DT 'S' '2026-01-01' 'first of two episodes both called EX1' "$V V13")),
    (VX 'VEC-V13' 1 'EX1' (DT 'S' '2026-01-02' 'second episode reusing the ID EX1' "$V V13")),
    (VX 'VEC-V13' 1 'EX2' (RG 'S' '2026-01-10' '2026-01-01' 'bounds entered in reverse order' "$V V13")),
    (VX 'VEC-V13' 1 'EX3' (DT 'S' '2026-01-15' '2026-01-15' "$V V13")),
    (VX 'VEC-V15' 1 'EX1' @{ Notes = 'No start date: V15 tests reported timing only' }),
    (VX 'VEC-V25' 1 'EX1' (TS 'S' '2026-01-10' '09:00' '-05:00' 'infusion start' "$V V25")),
    (VX 'VEC-V26' 1 'EX1' @{ Notes = 'Administration dates not part of vector V26' }),
    (VX 'VEC-V28' 1 'EX1' (TS 'T' '2026-01-10' '09:16' '-05:00' 'cessation 2026-01-10T09:16-05:00' "$V V28")),
    (VX 'VEC-V29' 1 'EX1' (DT 'T' '2026-07-05' 'cessation 2026-07-05' "$V V29")),
    (VX 'VEC-V30' 1 'EX1' (DT 'T' '2026-07-05' 'cessation 2026-07-05' "$V V30")),
    (MergeH @((CV 'VEC-V30' 1), @{ Drug_ID = 'D2'; Exposure_Episode_ID = 'EXO1'; Product = 'Other drug 1'; Role = 'treatment of the event'; Notes = 'TX-START in the vector: a treatment exposure, not recorded as an intervention' }, (DT 'S' '2026-07-06' 'treatment started 2026-07-06' "$V V30"))),
    (VX 'VEC-V31' 1 'EX1' (DT 'S' '2026-03-01' '2026-03-01' "$V V31")),
    (VX 'VEC-V33' 1 'EX1' (TS 'S' '2026-03-02' '08:00' '-05:00' '2026-03-02T08:00-05:00' "$V V33 (first administration)")),
    (VX 'VEC-V34' 1 'EX1' (DT 'S' '2026-03-01' '2026-03-01' "$V V34 (first administration)")),
    (VX 'VEC-V35' 1 'EX1' (DT 'S' '2026-01-01' '2026-01-01' "$V V35")),
    (VX 'VEC-V38' 1 'EX1' @{ Notes = 'Administrations listed out of date order (V38)' })
)

# ---------------- Administrations
$AdmEx = @(
    (MergeH @((CV 'EX-C01' 2), @{ Admin_ID = 'ADM-A1'; Exposure_Episode_ID = 'EX-A'; Kind = 'INFUSION_SEGMENT'; Rate = 50; Rate_Unit = 'mL/hour'; Notes = 'Pump rate 50 mL/hour from 09:00, unchanged until the stop; drug concentration in the bag not reported'; Missing_Reason = 'REQUESTED' },
        (TS 'A' '2026-01-10' '09:00' '-05:00' 'Drug A IV, started 09:00' 'Medication administration record'),
        @{ End_Date = '2026-01-10'; End_Time = '09:16'; End_Off = '-05:00'; End_Src = 'stopped 09:16 (reaction)' })),
    (MergeH @((CV 'EX-C04' 2), @{ Admin_ID = 'ADM-DIG-LAST'; Exposure_Episode_ID = 'EX-DIG'; Kind = 'DOSE'; Dose = 125; Dose_Unit = 'micrograms' },
        (TS 'A' '2026-07-04' '08:00' 'unknown' 'Took last digoxin yesterday morning, about 08:00' 'Emergency department note, 5 July 2026'))),
    (MergeH @((CV 'EX-C28' 2), @{ Admin_ID = 'ADM-RTX-1'; Exposure_Episode_ID = 'EX-RTX'; Kind = 'DOSE'; Notes = 'Dose as recorded in the source (not itemized)' },
        (DT 'A' '2024-11-18' 'first infusion 18/11/2024' 'Hematology summary'))),
    (MergeH @((CV 'EX-C28' 2), @{ Admin_ID = 'ADM-RTX-6'; Exposure_Episode_ID = 'EX-RTX'; Kind = 'DOSE'; Notes = 'Intermediate cycle dates not itemized in the source' },
        (DT 'A' '2025-03-03' 'last infusion 03/03/2025' 'Hematology summary')))
)
$AdmVec = @(
    (MergeH @((CV 'VEC-V07' 1), @{ Admin_ID = 'A1'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE' }, (TS 'A' '2026-03-01' '08:00' '-05:00' '2026-03-01T08:00-05:00' "$V V07"))),
    (MergeH @((CV 'VEC-V07' 1), @{ Admin_ID = 'A2'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE' }, (TS 'A' '2026-03-02' '08:00' '-05:00' '2026-03-02T08:00-05:00' "$V V07"))),
    (MergeH @((CV 'VEC-V07' 1), @{ Admin_ID = 'A3'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE' }, (TS 'A' '2026-03-03' '08:00' '-05:00' '2026-03-03T08:00-05:00' "$V V07"))),
    (MergeH @((CV 'VEC-V25' 1), @{ Admin_ID = 'S1'; Exposure_Episode_ID = 'EX1'; Kind = 'INFUSION_SEGMENT'; Rate = 50; Rate_Unit = 'mg/hour'; End_Date = '2026-01-10'; End_Time = '09:30'; End_Off = '-05:00'; End_Src = "$V V25" }, (TS 'A' '2026-01-10' '09:00' '-05:00' 'segment 1 start' "$V V25"))),
    (MergeH @((CV 'VEC-V25' 1), @{ Admin_ID = 'S2'; Exposure_Episode_ID = 'EX1'; Kind = 'INFUSION_SEGMENT'; Rate = 100; Rate_Unit = 'mg/hour'; End_Date = '2026-01-10'; End_Time = '13:00'; End_Off = '-05:00'; End_Src = "$V V25 (recorded planned stop)" }, (TS 'A' '2026-01-10' '09:30' '-05:00' 'segment 2 start' "$V V25"))),
    (MergeH @((CV 'VEC-V26' 1), @{ Admin_ID = 'D1'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE'; Dose = 100; Dose_Unit = 'mg'; BSA_m2 = 1.80 })),
    (MergeH @((CV 'VEC-V26' 1), @{ Admin_ID = 'D2'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE'; Dose = 100; Dose_Unit = 'mg'; BSA_m2 = 1.80 })),
    (MergeH @((CV 'VEC-V26' 1), @{ Admin_ID = 'D3'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE'; Dose = 90; Dose_Unit = 'mg'; BSA_m2 = 1.70 })),
    (MergeH @((CV 'VEC-V33' 1), @{ Admin_ID = 'AD2'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE' }, (TS 'A' '2026-03-02' '09:00' '+01:00' '2026-03-02T09:00+01:00' "$V V33"))),
    (MergeH @((CV 'VEC-V34' 1), @{ Admin_ID = 'AD2'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE' }, (RG 'A' '2026-03-01' '2026-03-10' 'another dose in early March' "$V V34"))),
    (MergeH @((CV 'VEC-V38' 1), @{ Admin_ID = 'D1'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE'; Dose = 90; Dose_Unit = 'mg'; BSA_m2 = 1.70 }, (DT 'A' '2026-02-01' '2026-02-01' "$V V38"))),
    (MergeH @((CV 'VEC-V38' 1), @{ Admin_ID = 'D2'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE'; Dose = 100; Dose_Unit = 'mg'; BSA_m2 = 1.80 }, (DT 'A' '2026-01-01' '2026-01-01' "$V V38"))),
    (MergeH @((CV 'VEC-V38' 1), @{ Admin_ID = 'D3'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE'; Dose = 100; Dose_Unit = 'mg'; BSA_m2 = 1.80 }, (DT 'A' '2026-01-15' '2026-01-15' "$V V38")))
)

# ---------------- Events
$EvtEx = @(
    (MergeH @((CV 'EX-C01' 2), @{ Event_ID = 'E1'; Episode_ID = 'EV-1'; Event_Term = 'Urticaria, wheeze and hypotension (acute infusion-associated reaction)'; Baseline_Status = '08:50 observations stable; skin clear' },
        (TS 'O' '2026-01-10' '09:15' '-05:00' '09:15 pt reports itching; urticaria chest and arms; audible wheeze; BP 84/52' 'Nursing observation chart'))),
    (MergeH @((CV 'EX-C04' 2), @{ Event_ID = 'E1'; Episode_ID = 'EV-1'; Event_Term = 'Nausea and symptomatic bradycardia'; Baseline_Status = 'Six months of tolerated treatment' },
        (DT 'O' '2026-07-05' 'Nausea since this morning, dizzy, pulse 38' 'Emergency department note'))),
    (MergeH @((CV 'EX-C28' 2), @{ Event_ID = 'E1'; Episode_ID = 'EV-HBV'; Event_Term = 'Hepatitis B virus reactivation'; Baseline_Status = '11/11/2024 HBsAg negative, anti-HBc positive, HBV DNA not detected' },
        @{ O_Prec = 'LAB'; O_LabNormal = 'M-HBV3'; O_LabAbnormal = 'M-HBV4'; O_Text = 'Between the last undetectable HBV DNA (9 February 2026) and the first positive result (12 May 2026)'; O_Src = 'Laboratory-defined onset (CALC-SPEC 1.1 C4): test records M-HBV3 (last normal) and M-HBV4 (first abnormal)' },
        (DT 'D' '2026-05-12' '12/05/2026 HBV DNA 1,850 IU/mL' 'Hepatology letter')))
)
function VE($case, $ver, $ep, $extra) { MergeH @((CV $case $ver), @{ Event_ID = 'E1'; Episode_ID = $ep; Event_Term = 'Vector event' }, $extra) }
$EvtVec = @(
    (VE 'VEC-V01' 1 'EV1' (TS 'O' '2026-01-10' '09:15' '-05:00' '2026-01-10T09:15-05:00' "$V V01")),
    (VE 'VEC-V02' 1 'EV1' (DT 'O' '2026-01-06' '2026-01-06' "$V V02")),
    (VE 'VEC-V03' 1 'EV1' (DT 'O' '2026-01-05' '2026-01-05' "$V V03")),
    (VE 'VEC-V04' 1 'EV1' (DT 'O' '2024-03-01' '2024-03-01' "$V V04")),
    (VE 'VEC-V05' 1 'EV1' (DT 'O' '2026-01-15' '2026-01-15' "$V V05")),
    (VE 'VEC-V06' 1 'EV1' (DT 'O' '2026-01-05' '2026-01-05' "$V V06")),
    (VE 'VEC-V07' 1 'EV1' (TS 'O' '2026-03-02' '20:00' '-05:00' '2026-03-02T20:00-05:00' "$V V07")),
    (VE 'VEC-V08' 1 'EV1' (DT 'O' '2026-02-14' '2026-02-14' "$V V08")),
    (VE 'VEC-V09' 1 'EV1' (TS 'O' '2026-11-01' '01:30' '-05:00' '2026-11-01T01:30-05:00' "$V V09")),
    (VE 'VEC-V10' 1 'EV1' (TS 'O' '2026-03-08' '03:30' 'unknown' '2026-03-08T03:30 (offset unknown)' "$V V10")),
    (VE 'VEC-V11' 1 'EV1' (TS 'O' '2026-04-01' '14:00' '-05:00' 'onset 14:00' "$V V11")),
    (VE 'VEC-V12' 1 'EV1' (DT 'O' '2026-05-10' 'onset 10 May (version 1)' "$V V12")),
    (VE 'VEC-V12' 2 'EV1' (DT 'O' '2026-05-08' 'onset 8 May (follow-up, version 2)' "$V V12")),
    (VE 'VEC-V13' 1 'EV1' (DT 'O' 'TEXT:2026-02-30' '30 February 2026' "$V V13")),
    (VE 'VEC-V13' 1 'EV2' (DT 'O' '2026-01-20' '2026-01-20' "$V V13")),
    (VE 'VEC-V15' 1 'EV1' @{ Notes = 'No onset date: V15 tests reported timing only' }),
    (VE 'VEC-V25' 1 'EV1' (TS 'O' '2026-01-10' '09:50' '-05:00' 'symptoms begin 09:50' "$V V25")),
    (VE 'VEC-V28' 1 'EV1' @{ Notes = 'C7 vector: the onset is not part of the vector' }),
    (VE 'VEC-V29' 1 'EV1' @{ Notes = 'C7 vector: the onset is not part of the vector' }),
    (VE 'VEC-V30' 1 'EV1' @{ Notes = 'C7 vector: the onset is not part of the vector' }),
    (VE 'VEC-V31' 1 'EV1' @{ O_Prec = 'LAB'; O_LabNormal = 'MS-NORMAL'; O_LabAbnormal = 'MS-ABNORMAL'; O_Text = 'between the tests'; O_Src = "$V V31 (laboratory-defined onset)" }),
    (VE 'VEC-V33' 1 'EV1' (TS 'O' '2026-03-02' '20:00' '-05:00' '2026-03-02T20:00-05:00' "$V V33")),
    (VE 'VEC-V34' 1 'EV1' (DT 'O' '2026-03-12' '2026-03-12' "$V V34")),
    (VE 'VEC-V35' 1 'EV1' (DT 'O' 'TEXT:2026-02-30' '2026-02-30' "$V V35"))
)

# ---------------- Milestones
function MSR($case, $id, $type, $desc, $tp, $extra = @{}) { MergeH @((CV $case 2), @{ Milestone_ID = $id; Type = $type; Description = $desc }, $tp, $extra) }
$MsEx = @(
    (MSR 'EX-C01' 'M-BASE' 'BASELINE' 'Pre-infusion observations: BP 124/78, HR 76, SpO2 98%, skin clear' (TS 'M' '2026-01-10' '08:50' '-05:00' '08:50 pre-infusion obs' 'Nursing chart')),
    (MSR 'EX-C01' 'M-EMERG' 'INTERVENTION' 'Emergency medicines administered (agents not given here)' (TS 'M' '2026-01-10' '09:18' '-05:00' 'Emergency medicines administered from 09:18' 'Medication administration record')),
    (MSR 'EX-C01' 'M-IMPROVE' 'RECOVERY' 'Wheeze settling, BP 108/70' (TS 'M' '2026-01-10' '09:35' '-05:00' '09:35 wheeze settling, BP 108/70' 'Nursing chart')),
    (MSR 'EX-C01' 'M-DIAG' 'DIAGNOSIS' 'Physician note: acute reaction during first Drug A infusion' (TS 'M' '2026-01-10' '09:50' '-05:00' 'Physician note written at 09:50' 'Physician note (version 2)')),
    (MSR 'EX-C01' 'M-TRYP' 'TEST' 'Tryptase sample taken; result pending' (TS 'M' '2026-01-10' '10:05' '-05:00' 'tryptase sample 10:05' 'Follow-up, version 2') @{ Missing_Reason = 'REQUESTED' }),
    (MSR 'EX-C04' 'M-CR1' 'ORGAN_FUNCTION' 'Serum creatinine (baseline)' (DT 'M' '2025-11-14' 'Creatinine 1.0 mg/dL on 14 November 2025' 'GP letter') @{ Value = 1.0; Unit = 'mg/dL' }),
    (MSR 'EX-C04' 'M-CR2' 'ORGAN_FUNCTION' 'Serum creatinine (baseline)' (DT 'M' '2026-03-12' 'Creatinine 1.0 mg/dL on 12 March 2026' 'GP letter') @{ Value = 1.0; Unit = 'mg/dL' }),
    (MSR 'EX-C04' 'M-CR3' 'ORGAN_FUNCTION' 'First documented renal deterioration (detection, not the start of the decline)' (DT 'M' '2026-07-01' 'Seen 1 July ... creatinine 1.7' 'GP letter') @{ Value = 1.7; Unit = 'mg/dL' }),
    (MSR 'EX-C04' 'M-CR4' 'ORGAN_FUNCTION' 'Serum creatinine at onset (still rising)' (DT 'M' '2026-07-05' 'Creatinine 2.4 mg/dL' 'Emergency department note') @{ Value = 2.4; Unit = 'mg/dL' }),
    (MSR 'EX-C04' 'M-DIGC' 'CONCENTRATION' 'Serum digoxin' (TS 'M' '2026-07-05' '11:40' 'unknown' 'Serum digoxin 2.9 ng/mL (sample 11:40; laboratory reference 0.5-2.0)' 'Emergency department note') @{ Value = 2.9; Unit = 'ng/mL'; Reference_Range = '0.5-2.0 ng/mL' }),
    (MSR 'EX-C04' 'M-K' 'TEST' 'Potassium' (DT 'M' '2026-07-05' 'potassium 4.8 mmol/L' 'Emergency department note') @{ Value = 4.8; Unit = 'mmol/L' }),
    (MSR 'EX-C04' 'M-FLUIDS' 'INTERVENTION' 'IV fluids, cardiac monitoring (treatments 5-8 July beyond fluids not reported)' (DT 'M' '2026-07-05' 'IV fluids, cardiac monitoring' 'Emergency department note')),
    (MSR 'EX-C04' 'M-IMPROVE' 'RECOVERY' 'Nausea resolved, ventricular rate 64; creatinine 1.9' (DT 'M' '2026-07-08' 'Nausea resolved, ventricular rate 64' 'Ward note, 8 July 2026')),
    (MSR 'EX-C28' 'M-HBV0' 'BASELINE' 'HBsAg negative, anti-HBc positive, HBV DNA not detected' (DT 'M' '2024-11-11' 'Baseline 11/11/2024' 'Hematology summary')),
    (MSR 'EX-C28' 'M-HBV1' 'TEST' 'HBV DNA not detected' (DT 'M' '2025-09-01' 'not detected 01/09/2025' 'Hematology summary')),
    (MSR 'EX-C28' 'M-HBV2' 'TEST' 'HBV DNA not detected' (DT 'M' '2025-12-01' 'not detected 01/12/2025' 'Hematology summary')),
    (MSR 'EX-C28' 'M-HBV3' 'TEST' 'HBV DNA not detected (last undetectable)' (DT 'M' '2026-02-09' 'not detected 09/02/2026' 'Hematology summary')),
    (MSR 'EX-C28' 'M-HBV4' 'TEST' 'HBV DNA 1,850 IU/mL (first positive), ALT 28 U/L' (DT 'M' '2026-05-12' '12/05/2026 HBV DNA 1,850 IU/mL' 'Hepatology letter') @{ Value = 1850; Unit = 'IU/mL' }),
    (MSR 'EX-C28' 'M-BCELL' 'TEST' 'CD19 B cells and IgG below reference range' (DT 'M' '2026-05-12' 'CD19 B cells and IgG below reference range on 12/05/2026' 'Hepatology letter') @{ Missing_Reason = 'NOT_REPORTED'; Notes = 'Values not reported' }),
    (MSR 'EX-C28' 'M-HEP' 'TEST' 'Biochemical hepatitis: HBV DNA 64,000 IU/mL, HBsAg positive, ALT 142 U/L (ULN 40)' (DT 'M' '2026-05-26' '26/05/2026 HBV DNA 64,000 IU/mL, HBsAg positive, ALT 142 U/L' 'Hepatology letter') @{ Value = 142; Unit = 'U/L (ALT)'; Reference_Range = 'ULN 40 U/L' }),
    (MSR 'EX-C28' 'M-AV' 'INTERVENTION' 'Antiviral A restarted' (DT 'M' '2026-05-27' 'Antiviral A restarted 27/05/2026' 'Hepatology letter'))
)

function MSV($case, $id, $type, $desc, $tp) { MergeH @((CV $case 1), @{ Milestone_ID = $id; Type = $type; Description = $desc }, $tp) }
$MsVec = @(
    (MSV 'VEC-V28' 'MS-REC' 'RECOVERY' 'Recovery' (TS 'M' '2026-01-10' '09:35' '-05:00' '2026-01-10T09:35-05:00' "$V V28")),
    (MSV 'VEC-V28' 'MS-RX' 'INTERVENTION' 'emergency medicines from 09:18' (TS 'M' '2026-01-10' '09:18' '-05:00' '2026-01-10T09:18-05:00' "$V V28")),
    (MSV 'VEC-V29' 'MS-REC' 'RECOVERY' 'Recovery' (DT 'M' '2026-07-08' '2026-07-08' "$V V29")),
    (MSV 'VEC-V29' 'MS-FLU' 'INTERVENTION' 'IV fluids and cardiac monitoring' (DT 'M' '2026-07-05' '2026-07-05' "$V V29")),
    (MSV 'VEC-V30' 'MS-REC' 'RECOVERY' 'Recovery' (DT 'M' '2026-07-08' '2026-07-08' "$V V30")),
    (MSV 'VEC-V30' 'MS-PRE' 'INTERVENTION' 'intervention before cessation' (DT 'M' '2026-07-03' '2026-07-03' "$V V30")),
    (MSV 'VEC-V30' 'MS-POST' 'INTERVENTION' 'intervention after recovery' (DT 'M' '2026-07-09' '2026-07-09' "$V V30")),
    (MSV 'VEC-V31' 'MS-NORMAL' 'TEST' 'Last normal test' (DT 'M' '2026-03-15' '2026-03-15' "$V V31")),
    (MSV 'VEC-V31' 'MS-ABNORMAL' 'TEST' 'First abnormal test' (DT 'M' '2026-03-29' '2026-03-29' "$V V31"))
)

# ---------------- Links
function LK($case, $ver, $id, $ex, $ev, $why) { MergeH @((CV $case $ver), @{ Link_ID = $id; Exposure_Episode_ID = $ex; Event_Episode_ID = $ev; Reason = $why; Selected_By = 'reviewer' }) }
$LnkEx = @(
    (LK 'EX-C01' 2 'L1' 'EX-A' 'EV-1' 'Suspect drug named in the report'),
    (LK 'EX-C01' 2 'L2' 'EX-B' 'EV-1' 'Competing exposure through the same cannula (version 2)'),
    (LK 'EX-C01' 2 'L3' 'EX-AS' 'EV-1' 'Competing exposure: skin antiseptic (version 2)'),
    (LK 'EX-C04' 2 'L1' 'EX-DIG' 'EV-1' 'Suspect drug'),
    (LK 'EX-C04' 2 'L2' 'EX-IBU' 'EV-1' 'Interacting co-medication (NSAID may impair digoxin excretion)'),
    (LK 'EX-C28' 2 'L1' 'EX-RTX' 'EV-HBV' 'Suspect drug'),
    (LK 'EX-C28' 2 'L2' 'EX-AVA' 'EV-HBV' 'End of antiviral prophylaxis (concomitant change)'),
    (LK 'EX-C28' 2 'L3' 'EX-PRED' 'EV-HBV' 'Competing immunosuppressive exposure'),
    (LK 'EX-C28' 2 'L4' 'EX-RX' 'EV-HBV' 'Concomitant chemotherapy with a glucocorticoid; same dates as rituximab')
)
$LnkVec = @()
foreach ($v in @('V01', 'V02', 'V03', 'V04', 'V05', 'V06', 'V07', 'V08', 'V09', 'V10', 'V11', 'V15', 'V28', 'V29', 'V30', 'V31', 'V33', 'V34')) { $LnkVec += (LK "VEC-$v" 1 'L1' 'EX1' 'EV1' "Vector $v pair") }
$LnkVec += (LK 'VEC-V12' 1 'L1' 'EX1' 'EV1' 'Vector V12 pair (version 1)')
$LnkVec += (LK 'VEC-V12' 2 'L1' 'EX1' 'EV1' 'Vector V12 pair (version 2)')
$LnkVec += (LK 'VEC-V13' 1 'L1' 'EX2' 'EV2' 'Vector V13 reversed-bounds pair')
$LnkVec += (LK 'VEC-V13' 1 'L2' 'EX1' 'EV2' 'Vector V13 duplicate-ID pair')
$LnkVec += (LK 'VEC-V13' 1 'L3' 'EX3' 'EV1' 'Vector V13 invalid-date pair')

# ---------------- Intervals
$IntEx = @(
    (IVR 'EX-C01' 2 'I01' 'L1' 'FIRST_ADMIN' '' 'EX-A/START' 'EV-1/ONSET' 'C1'),
    (IVR 'EX-C01' 2 'I02' 'L1' 'FIRST_ADMIN' '' 'EX-A/START' 'EV-1/ONSET' 'C2' @{ Treatment_Day_Convention = 'inclusive'; Reviewer_Note = 'Date-level result; C1 decides the order because both times are known' }),
    (IVR 'EX-C01' 2 'I03' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' 'EV-1/ONSET' 'C1' @{ Reviewer_Note = 'The infusion running at onset (its start, EX-A/START); its 09:16 end belongs to the chronology, not the anchor' }),
    (IVR 'EX-C01' 2 'I04' 'L2' 'CONCOMITANT_CHANGE' 'Drug B infusion start' 'EX-B/START' 'EV-1/ONSET' 'C1'),
    (IVR 'EX-C01' 2 'I05' 'L2' 'OTHER' 'End of Drug B infusion' 'EX-B/STOP' 'EV-1/ONSET' 'C1'),
    (IVR 'EX-C01' 2 'I06' 'L3' 'OTHER' 'Skin antiseptic at cannulation (topical exposure)' 'EX-AS/START' 'EV-1/ONSET' 'C1'),
    (IVR 'EX-C01' 2 'I07' 'L1' 'OTHER' 'Baseline observations (last documented normal status)' 'M-BASE' 'EV-1/ONSET' 'C1'),
    (IVR 'EX-C01' 2 'I08' 'L1' 'CESSATION' '' 'EX-A/STOP' 'EV-1/ONSET' 'C1' @{ Reviewer_Note = 'Exposure continued 1 minute after onset' }),
    (IVR 'EX-C01' 2 'I09' 'L1' 'OTHER' 'Start of emergency treatment' 'M-EMERG' 'EV-1/ONSET' 'C1'),
    (IVR 'EX-C01' 2 'I10' 'L1' 'CESSATION' '' 'EX-A/STOP' 'M-IMPROVE' 'C1' @{ Reviewer_Note = 'Recovery interval (C7); Chapter 17' }),
    (IVR 'EX-C01' 2 'I11' 'L1' 'OTHER' 'Start of emergency treatment' 'M-EMERG' 'M-IMPROVE' 'C1'),
    (IVR 'EX-C01' 2 'I12' 'L1' 'OTHER' 'Onset of the reaction' 'EV-1/ONSET' 'M-IMPROVE' 'C1'),
    (IVR 'EX-C01' 2 'I13' 'L1' 'FIRST_ADMIN' '' 'EX-A/START' 'EX-A/STOP' 'C1' @{ Reviewer_Note = 'Duration of the Drug A infusion' }),
    (IVR 'EX-C01' 2 'I14' 'L1' 'OTHER' 'Onset of the reaction' 'EV-1/ONSET' 'M-TRYP' 'C1' @{ Reviewer_Note = 'Tryptase sampling time relative to onset' }),
    (IVR 'EX-C04' 2 'I01' 'L1' 'FIRST_ADMIN' '' 'EX-DIG/START' 'EV-1/ONSET' 'C2'),
    (IVR 'EX-C04' 2 'I02' 'L1' 'OTHER' 'First documented renal deterioration (detection, not the start of the decline)' 'M-CR3' 'EV-1/ONSET' 'C2'),
    (IVR 'EX-C04' 2 'I03' 'L2' 'CONCOMITANT_CHANGE' 'Start of ibuprofen' 'EX-IBU/START' 'EV-1/ONSET' 'C2'),
    (IVR 'EX-C04' 2 'I04' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' 'EV-1/ONSET' 'C2'),
    (IVR 'EX-C04' 2 'I05' 'L1' 'OTHER' 'Last dose before the concentration sample' 'ADM-DIG-LAST' 'M-DIGC' 'C1' @{ Reviewer_Note = 'Sample time after the last dose (sampling advice LANOXIN US 2.5)' }),
    (IVR 'EX-C04' 2 'I06' 'L1' 'CESSATION' '' 'EX-DIG/STOP' 'M-IMPROVE' 'C2' @{ Reviewer_Note = 'Recovery interval (C7); IV fluids recorded, ibuprofen stopped the same day' }),
    (IVR 'EX-C28' 2 'I01' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' 'EV-HBV/DETECTION' 'C2' @{ Reported_Text = 'fourteen months after the last dose'; Reviewer_Note = 'Reporter comment kept as reported; months are not converted' }),
    (IVR 'EX-C28' 2 'I02' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' 'EV-HBV/ONSET' 'C4'),
    (IVR 'EX-C28' 2 'I03' 'L1' 'FIRST_ADMIN' '' 'EX-RTX/START' 'EV-HBV/DETECTION' 'C2'),
    (IVR 'EX-C28' 2 'I04' 'L1' 'FIRST_ADMIN' '' 'EX-RTX/START' 'EV-HBV/ONSET' 'C4'),
    (IVR 'EX-C28' 2 'I05' 'L2' 'CONCOMITANT_CHANGE' 'End of Antiviral A prophylaxis' 'EX-AVA/STOP' 'EV-HBV/DETECTION' 'C2'),
    (IVR 'EX-C28' 2 'I06' 'L2' 'CONCOMITANT_CHANGE' 'End of Antiviral A prophylaxis' 'EX-AVA/STOP' 'EV-HBV/ONSET' 'C4'),
    (IVR 'EX-C28' 2 'I07' 'L3' 'CONCOMITANT_CHANGE' 'Start of prednisolone course' 'EX-PRED/START' 'EV-HBV/DETECTION' 'C2'),
    (IVR 'EX-C28' 2 'I08' 'L3' 'CONCOMITANT_CHANGE' 'Start of prednisolone course' 'EX-PRED/START' 'EV-HBV/ONSET' 'C4'),
    (IVR 'EX-C28' 2 'I09' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' 'M-HEP' 'C2' @{ Reviewer_Note = 'Biochemical hepatitis milestone' }),
    (IVR 'EX-C28' 2 'I10' 'L1' 'OTHER' 'Last undetectable HBV DNA' 'M-HBV3' 'EV-HBV/DETECTION' 'C2' @{ Reviewer_Note = 'Width of the detection window' }),
    (IVR 'EX-C28' 2 'I11' 'L1' 'OTHER' 'Last rituximab infusion' 'ADM-RTX-6' 'EX-AVA/STOP' 'C2' @{ Reviewer_Note = 'Prophylaxis stopped after the last infusion' })
)
$IntVec = @(
    (IVR 'VEC-V01' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C1'),
    (IVR 'VEC-V02' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C2' @{ Treatment_Day_Convention = 'inclusive' }),
    (IVR 'VEC-V03' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C2'),
    (IVR 'VEC-V04' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C2'),
    (IVR 'VEC-V05' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C4'),
    (IVR 'VEC-V06' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C4'),
    (IVR 'VEC-V07' 1 'I1' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' 'EV1/ONSET' 'C1'),
    (IVR 'VEC-V08' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C2'),
    (IVR 'VEC-V08' 1 'I2' 'L1' 'RESTART' '' 'EX1/RESTART' 'EV1/ONSET' 'C2'),
    (IVR 'VEC-V09' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C1'),
    (IVR 'VEC-V10' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C1'),
    (IVR 'VEC-V11' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C1' @{ Reported_Text = 'two hours after the dose'; Reported_Lower = 120; Reported_Upper = 120; Reported_Unit = 'minutes'; Reported_Bounds_Justification = 'Bounds stated in the vector' }),
    (IVR 'VEC-V12' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C2'),
    (IVR 'VEC-V12' 2 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C2'),
    (IVR 'VEC-V13' 1 'I-DATE' 'L3' 'FIRST_ADMIN' '' 'EX3/START' 'EV1/ONSET' 'C2'),
    (IVR 'VEC-V13' 1 'I-BOUNDS' 'L1' 'FIRST_ADMIN' '' 'EX2/START' 'EV2/ONSET' 'C4'),
    (IVR 'VEC-V13' 1 'I-IDS' 'L2' 'FIRST_ADMIN' '' 'EX1/START' 'EV2/ONSET' 'C2'),
    (IVR 'VEC-V13' 1 'I-LINK' 'L9' 'FIRST_ADMIN' '' 'EX3/START' 'EV2/ONSET' 'C2' @{ Reviewer_Note = 'Pair D1 x E1 has no Drug_Event_Links row' }),
    (IVR 'VEC-V15' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C2' @{ Reported_Text = 'about one month after starting' }),
    (IVR 'VEC-V28' 1 'I1' 'L1' 'CESSATION' '' 'EX1/STOP' 'MS-REC' ''),
    (IVR 'VEC-V29' 1 'I1' 'L1' 'CESSATION' '' 'EX1/STOP' 'MS-REC' ''),
    (IVR 'VEC-V30' 1 'I1' 'L1' 'CESSATION' '' 'EX1/STOP' 'MS-REC' '' @{ Reviewer_Note = 'Labeled no-flag variant of V29' }),
    (IVR 'VEC-V31' 1 'I1' 'L1' 'FIRST_ADMIN' '' '' '' '' @{ Reviewer_Note = 'Anchor and end point left blank: the episode START and the linked event onset (laboratory-defined)' }),
    (IVR 'VEC-V33' 1 'I1' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' '' ''),
    (IVR 'VEC-V34' 1 'I1' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' '' ''),
    (IVR 'VEC-V35' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' '' @{ Reviewer_Note = 'No Drug_Event_Links row for this case version (V35)' })
)

# ================================================================== INTERVALS COLUMNS
# CALC-SPEC 1.1: C1-C7 with the agreed rules of Section 9 (finest rule, C5 selection, C7 window, every E-code listed).
function LkA([string]$n, [string]$tp, [string]$def = '""') { CF "A_$n" "Anchor: $n" ('=IF({A_Row}=0,' + $def + ',INDEX(nTP_' + $tp + ',{A_Row}))') 8 }
function LkO([string]$n, [string]$tp, [string]$def = '""') { CF "O_$n" "End point: $n" ('=IF({O_Row}=0,' + $def + ',INDEX(nTP_' + $tp + ',{O_Row}))') 8 }

# ---- C5 (LAST_DOSE_BEFORE_ONSET) array fragments over Timepoints; candidates = the episode's START, RESTART and
#      administration (DOSE, DOSE_CHANGE, OTHER) records (nTP_CandKey), compared with the end point O_*.
$C5BT = 'nTP_HT*{O_HT}'
$C5CMP = '(nTP_OK*{O_OK}*(nTP_UTC_n-N({O_UTC}))+(1-nTP_OK*{O_OK})*(nTP_Loc_n-N({O_Loc})))'
$C5BEF = '(' + $C5BT + '*(' + $C5CMP + '<=0)+(1-' + $C5BT + ')*(nTP_Hi_n<N({O_Lo})))'
$C5AFT = '(' + $C5BT + '*(' + $C5CMP + '>0)+(1-' + $C5BT + ')*(nTP_Lo_n>N({O_Hi})))'
$C5SAME = '((1-' + $C5BT + ')*(nTP_Lo_n=nTP_Hi_n)*(nTP_Lo_n=N({O_Lo}))*(N({O_Lo})=N({O_Hi})))'
$C5MV = '(nTP_CandKey={C5_CK})*(nTP_Err="")'
$C5EL = '(nTP_CandKey={C5_CK})*nTP_Exact*(' + $C5BEF + '+' + $C5SAME + ')'
$C5PB = $C5MV + '*(1-nTP_Exact)*(' + $C5BEF + '+' + $C5SAME + ')'
$C5UNC = $C5MV + '*(1-' + $C5BEF + '-' + $C5AFT + '-' + $C5SAME + ')'
function KTH([string]$rowN, [string]$mask, [int]$k) { 'IFERROR(AGGREGATE(15,6,' + $rowN + '/(' + $mask + '),' + $k + '),0)' }
# up to three record IDs (with labels) matching a Timepoints mask, joined with "; "
function TPLIST([string]$mask, [string]$count, [bool]$withLabel = $true) {
    $parts = @()
    for ($k = 1; $k -le 3; $k++) {
        $r = KTH 'nTP_RowN' $mask $k
        $item = 'INDEX(nTP_Record_ID,' + $r + ')'
        if ($withLabel) { $item = $item + '&IF(INDEX(nTP_Label,' + $r + ')="",""," ("&INDEX(nTP_Label,' + $r + ')&")")' }
        $parts += ('IF(' + $count + '>=' + $k + ',"; "&' + $item + ',"")')
    }
    'MID(' + ($parts -join '&') + '&IF(' + $count + '>3,"; and "&(' + $count + '-3)&" more",""),3,999)'
}
# up to three IDs from helper columns {P_R1}..{P_R3} (Timepoints rows), joined with "; "
function TPL([string]$pre, [string]$count, [bool]$withLabel = $true) {
    $parts = @()
    for ($k = 1; $k -le 3; $k++) {
        $r = '{' + $pre + '_R' + $k + '}'
        $item = 'INDEX(nTP_Record_ID,' + $r + ')'
        if ($withLabel) { $item = $item + '&IF(INDEX(nTP_Label,' + $r + ')="",""," ("&INDEX(nTP_Label,' + $r + ')&")")' }
        $parts += ('IF(' + $r + '>0,"; "&' + $item + ',"")')
    }
    '=IF(' + $count + '=0,"",MID(' + ($parts -join '&') + '&IF(' + $count + '>3,"; and "&(' + $count + '-3)&" more",""),3,999))'
}
function TPRows([string]$pre, [string]$mask, [string]$count) {
    $a = @()
    for ($k = 1; $k -le 3; $k++) { $a += (CF ($pre + '_R' + $k) ('work: ' + $pre + ' row ' + $k) ('=IF(' + $count + '<' + $k + ',0,' + (KTH 'nTP_RowN' $mask $k) + ')') 5 '0') }
    return $a
}
# C7 recorded interventions (CALC-SPEC 1.1 C7): only INTERVENTION milestones of the case version; nothing inferred
$C7IM = '(nTP_Case_Key={CaseKey})*(nTP_Record_Type="MILESTONE:INTERVENTION")*(nTP_Err="")'
$C7ALL = 'nTP_HT*{A_HT}*{O_HT}'
$C7WIN = '(' + $C7ALL + '*(nTP_OK*{A_OK}*(nTP_UTC_n>=N({A_UTC}))+(1-nTP_OK*{A_OK})*(nTP_Loc_n>=N({A_Loc})))*(nTP_OK*{O_OK}*(nTP_UTC_n<=N({O_UTC}))+(1-nTP_OK*{O_OK})*(nTP_Loc_n<=N({O_Loc})))+(1-' + $C7ALL + ')*(nTP_Hi_n>=N({A_Lo}))*(nTP_Lo_n<=N({O_Hi})))'
$C7TX = '(nTP_Case_Key={CaseKey})*(nTP_Role="treatment of the event")*((nTP_Record_Type="EXPOSURE_START")+(nTP_Record_Type="ADMINISTRATION"))*(nTP_Err="")'
$C7UND = '(nMS_CaseKey={CaseKey})*(nMS_Type="INTERVENTION")*nMS_Undated'

function PRECLBL([string]$prec, [string]$ok) { 'IF(' + $prec + '="TIMESTAMP",IF(' + $ok + '=1,"timestamp with UTC offset","timestamp, offset unknown"),IF(' + $prec + '="DATE","day",IF(' + $prec + '="MONTH","month",IF(' + $prec + '="YEAR","year",IF(' + $prec + '="LAB","laboratory-defined range","range")))))' }
function C4SGN([string]$v) { 'IF(AND({V_Min4}<0,{V_Max4}>0,' + $v + '>0),"+","")&' + (MINUSF ($v + '&""')) }

$IntFlagsW = @(
    @('{f_TZ}=1', 'W_TZ_UNKNOWN'), @('{f_SameDate}=1', 'W_SAME_DATE_ORDER_UNKNOWN'), @('{f_SeqUnc}=1', 'W_SEQUENCE_UNCERTAIN'),
    @('{f_AdminSame}=1', 'W_ADMIN_SAME_DATE_AS_ONSET'), @('{f_Conflict}=1', 'W_REPORTED_TIMING_CONFLICT'), @('{f_Interv}=1', 'W_INTERVENTION_BEFORE_RECOVERY'),
    @('{f_TDundef}=1', 'I_TREATMENT_DAY_UNDEFINED'), @('{f_Approx}=1', 'I_APPROXIMATE_NOT_COMPARED'))
$IntQC = '=IF({RowEmpty}=1,"",IF({K_Err}<>"",{K_Err},MID(' + (($IntFlagsW | ForEach-Object { 'IF(' + $_[0] + ',"; ' + $_[1] + '","")' }) -join '&') + ',3,999)))'
$IntMsg = '=IF({RowEmpty}=1,"",MID(' + (EMSGS '{K_Err}' '{K_Needs}' $ECODES_INT) + '&' + (($IntFlagsW | ForEach-Object { WMSG $_[0] $_[1] }) -join '&') + ',4,9999))'

$IntAssume = '=IF({K_Ok}=0,"",MID(' +
    'IF({f_C7}=1," | Recovery interval (C7) computed with "&{Mode}&".","")' +
    '&IF({AutoC5}=1," | Anchor selected under C5 (CALC-SPEC 1.1): latest administration with an exact date or time at or before "&IF({End_Is_Onset}=1,"onset","the end point")&" ("&{Anchor_Record_ID_Used}&"), ordered in UTC when offsets are known; later administrations are never selected.","")' +
    '&IF(AND({AutoC5}=1,{C5_NPB}>0)," | Partial-date administrations wholly before the selected one were not candidates: "&{PB_List}&".","")' +
    '&IF({O_Lab_Note}<>""," | "&{O_Lab_Note},"")' +
    '&IF(AND({Mode}="C2",{V_Days}=0,{V_Min}<>"")," | Same calendar date; both times are known, so C1 decides the order ("&{Sequence_Status}&").","")' +
    '&IF(AND({Mode}="C2",{A_HT}*{O_HT}=1,{f_C7}=0)," | Calendar-date difference from the recorded local dates; it is never described as elapsed days.","")' +
    '&IF(AND({Mode}="C4",{f_C7}=0)," | Bounds are treated as independent; no single date is imputed from them.","")' +
    '&IF({Treatment_Day_Used}=1," | Treatment day under the inclusive convention (C3); never written to E2B fields automatically.","")' +
    '&IF(AND({Mode}="C4",{f_C7}=0,{K_HasC}=0)," | No constraint applied.","")' +
    '&IF(AND({Mode}="C4",{K_HasC}=1)," | Constraint applied to the anchor bounds: "&IF(ISNUMBER({Constraint_Not_Before}),"not before "&' + (DTXT '{Constraint_Not_Before}') + ',"")&IF(AND(ISNUMBER({Constraint_Not_Before}),ISNUMBER({Constraint_Not_After})),", ","")&IF(ISNUMBER({Constraint_Not_After}),"not after "&' + (DTXT '{Constraint_Not_After}') + ',"")&". Source: "&IF({Constraint_Source}="","not recorded",{Constraint_Source})&IF({K_CErr}="",". Constrained result shown beside the unconstrained one.",". Constrained result blocked ("&{K_CErr}&")."),"")' +
    '&IF({f_C7}=1,IF({IntervCount}>0," | Intervention recorded between cessation and recovery (ends included): "&{IV_List}&"."," | No intervention is recorded between cessation and recovery; absence of a record is not evidence that none was given."),"")' +
    '&IF(AND({f_C7}=1,{IntvUndated}>0)," | Intervention recorded without a date, so whether it came before recovery is unknown: "&INDEX(nMS_Milestone_ID,MATCH(1,INDEX(' + $C7UND + ',0),0))&IF({IntvUndated}>1," and "&({IntvUndated}-1)&" more","")&".","")' +
    '&IF(AND({f_C7}=1,{TxNotCounted}>0)," | Not counted as interventions, because they are recorded as exposures rather than as intervention milestones (nothing is inferred): "&{TX_List}&". If treatment was given, record it as an intervention milestone.","")' +
    ',4,9999))'

$IntParamSet = '=IF({K_Ok}=0,"","Anchor ("&{Anchor_Type}&") "&{Anchor_Record_ID_Used}&": "&{Anchor_ISO}&" ["&' + (PRECLBL '{A_Prec}' '{A_OK}') + '&"; case measurement; original: "&CHAR(34)&{A_Text}&CHAR(34)&IF({A_Src}<>"","; source: "&{A_Src},"")&"] | End point "&{End_Record_ID_Used}&": "&{End_ISO}&" ["&' + (PRECLBL '{O_Prec}' '{O_OK}') + '&"; case measurement; original: "&CHAR(34)&{O_Text}&CHAR(34)&IF({O_Src}<>"","; source: "&{O_Src},"")&"]")'

$IntCols = @(
    (CIn 'Interval_ID' 'Interval_ID (unique in the case version)' 10 $null '@' 'INPUTS (yellow): one row per link and anchor'),
    (CIn 'Case_ID' 'Case_ID' 10 $null '@'),
    (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
    (CIn 'Link_ID' 'Link_ID (Drug_Event_Links)' 8 $null '@'),
    (CIn 'Anchor_Type' 'Anchor_Type (C5)' 16 'list:nLIST_AnchorType'),
    (CIn 'Anchor_Description' 'Anchor description (required for OTHER)' 22 $null '@'),
    (CIn 'Anchor_Record_ID' 'Anchor_Record_ID (see Timepoints). Blank: LAST_DOSE_BEFORE_ONSET is selected by C5; FIRST_ADMIN uses the episode START' 16 $null '@'),
    (CIn 'End_Record_ID' 'End-point record ID (blank = onset of the linked event; or EV-1/DETECTION, a RECOVERY milestone ...)' 16 $null '@'),
    (CIn 'Calculation_Mode' 'Rule requested: blank = the finest the records support (C1 timestamps, C2 dates, C4 bounds); a coarser rule may be chosen' 9 'list:nLIST_Mode'),
    (CIn 'Treatment_Day_Convention' 'Treatment day (C3): inclusive or not requested' 11 'list:nLIST_TD'),
    (CIn 'Reported_Text' 'Reported relative timing: text as reported (C6)' 22 $null '@'),
    (CIn 'Reported_Lower' 'Reported lower bound (only if stated)' 8 'num'),
    (CIn 'Reported_Upper' 'Reported upper bound (only if stated)' 8 'num'),
    (CIn 'Reported_Unit' 'Reported unit' 10 'list:nLIST_ReportedUnit'),
    (CIn 'Reported_Bounds_Justification' 'Source or justification for reported bounds' 18 $null '@'),
    (CIn 'Constraint_Not_Before' 'C4 constraint: anchor not before (date)' 13 'date' 'd mmmm yyyy'),
    (CIn 'Constraint_Not_After' 'C4 constraint: anchor not after (date)' 13 'date' 'd mmmm yyyy'),
    (CIn 'Constraint_Source' 'C4 constraint source (required)' 16 $null '@'),
    (CIn 'Reviewer_Note' 'Reviewer note' 22 $null '@'),

    (CF 'Result_Display' 'Result (display per CALC-SPEC Section 7; the E-codes and no number when blocked)' ('=IF({RowEmpty}=1,"",IF({K_Ok}=0,{K_Err}&": no result is shown",{Displayed_Value}&IF({Mode}="C1",IF(ABS({V_Min})>=60," ("&FIXED({V_Min}/60,1,FALSE)&" hours)",""),"")&IF(AND({Mode}="C2",{V_Days}=0,{V_Min}<>"")," (order within the day from C1: "&{Sequence_Status}&")","")&IF({Treatment_Day}<>"","; treatment Day "&{Treatment_Day},"")&IF({Constrained_Minimum}<>"","; with constraint "&' + (MINUSF '{Constrained_Minimum}&""') + '&" to "&' + (MINUSF '{Constrained_Maximum}&""') + '&" calendar days",IF({K_CErr}<>"","; constraint: "&{K_CErr},""))))') 34 $null 'RESULT (grey formulas, locked)'),
    (CF 'Displayed_Value' 'Displayed_Value (as the calculator; Not calculable when blocked)' ('=IF({RowEmpty}=1,"",IF({K_Ok}=0,"Not calculable",IF({Mode}="C1",' + (MINUSF 'FIXED({V_Min},0,FALSE)') + '&" "&IF(ABS({V_Min})=1,SUBSTITUTE({Unit},"minutes","minute"),{Unit}),IF({Mode}="C2",' + (MINUSF '{V_Days}&""') + '&" calendar "&IF(ABS({V_Days})=1,"day","days"),' + (C4SGN '{V_Min4}') + '&" to "&' + (C4SGN '{V_Max4}') + '&" calendar days"))))') 22),
    (CF 'Interval_Value' 'Interval_Value (unrounded; C1 minutes or C2 calendar days)' '=IF({K_Ok}=0,"",IF({Mode}="C1",{V_Min},IF({Mode}="C2",{V_Days},"")))' 10),
    (CF 'Interval_Minimum' 'Interval_Minimum (C4, calendar days)' '=IF(AND({K_Ok}=1,{Mode}="C4"),{V_Min4},"")' 9),
    (CF 'Interval_Maximum' 'Interval_Maximum (C4, calendar days)' '=IF(AND({K_Ok}=1,{Mode}="C4"),{V_Max4},"")' 9),
    (CF 'Unit' 'Unit' '=IF({K_Ok}=0,"",IF({Mode}="C1",IF({K_TZU}=1,"minutes (local clock)","minutes"),"calendar days"))' 12),
    (CF 'Calculation_Mode_Used' 'Calculation_Mode (rule applied)' ('=IF({RowEmpty}=1,"",IF({K_Ok}=0,IF({f_C7}=1,"C7","C1' + $U_NDASH + 'C6"),IF({f_C7}=1,"C7 ("&{Mode}&")",{Mode}&IF({Treatment_Day_Used}=1,"; C3","")&IF({C6_On}=1,"; C6",""))))') 9),
    (CF 'Sequence_Status' 'Sequence_Status' ('=IF({K_Ok}=0,"",IF({Mode}="C1",' + (SEQ '{V_Min}') + ',IF({Mode}="C2",IF({V_Days}>0,"AFTER",IF({V_Days}<0,"BEFORE",IF({V_Min}<>"",' + (SEQ '{V_Min}') + ',"SAME_DATE_ORDER_UNKNOWN"))),IF({V_Min4}>0,"AFTER",IF({V_Max4}<0,"BEFORE",IF(AND({V_Min4}=0,{V_Max4}=0),"SAME_DATE_ORDER_UNKNOWN","SEQUENCE_UNCERTAIN"))))))') 16),
    (CF 'Treatment_Day' 'Treatment day number (C3, inclusive)' '=IF(AND({Treatment_Day_Used}=1,{V_Days}>=0),{V_Days}+1,"")' 8),
    (CF 'QC_Flags' 'QC_Flags (E-, W-, I-codes; rule G9 order)' $IntQC 22),
    (CF 'Messages' 'Flag_Messages (CALC-SPEC Section 5, verbatim)' $IntMsg 40),
    (CF 'Constrained_Minimum' 'C4 with constraint: minimum' '=IF(OR({K_CLo}="",{K_CErr}<>""),"",{O_Lo}-{K_CHi})' 9),
    (CF 'Constrained_Maximum' 'C4 with constraint: maximum' '=IF(OR({K_CLo}="",{K_CErr}<>""),"",{O_Hi}-{K_CLo})' 9),
    (CF 'Constraint_Check' 'Constrained result: E-codes and messages (shown beside the unconstrained result)' ('=IF({K_CErr}="","",MID(' + (EMSGS '{K_CErr}' '"the source of the constraint"' @('E_REVERSED_BOUNDS', 'E_MISSING_INPUT')) + ',4,999))') 24),
    (CF 'Reported_Timing_Result' 'Reported timing (C6): kept beside the computed value' '=IF({C6_On}=0,"",IF({C6_Code}="OVERLAP","Overlaps the computed interval. ",IF({C6_Code}="W_REPORTED_TIMING_CONFLICT","Conflicts with the computed interval; both kept. ",IF({C6_Code}="I_APPROXIMATE_NOT_COMPARED","Not compared. ",IF({C6_Code}="","Not compared (no computed interval). ",{C6_Code}&": not compared. "))))&"Reported: "&CHAR(34)&{Reported_Text}&CHAR(34)&IF(AND(ISNUMBER({Reported_Lower}),ISNUMBER({Reported_Upper}))," ["&{Reported_Lower}&" to "&{Reported_Upper}&" "&{Reported_Unit}&"]",""))' 30),
    (CF 'Label' 'Label' ('=IF({RowEmpty}=1,"",IF({K_Ok}=0,"Interval "&{Interval_ID},{Anchor_Record_ID_Used}&" ' + $U_ARROW + ' "&{End_Record_ID_Used}&IF({f_C7}=1," (recovery)","")))') 24),
    (CF 'Anchor_Record_ID_Used' 'Anchor_Record_ID used' '=IF({AutoC5}=1,{C5_Selected},IF(AND({Anchor_Type}="FIRST_ADMIN",{Anchor_Record_ID}=""),IF({Exposure_Episode_ID}="","",{Exposure_Episode_ID}&"/START"),{Anchor_Record_ID}&""))' 13),
    (CF 'End_Record_ID_Used' 'End-point record used' '=IF({End_Record_ID}<>"",{End_Record_ID}&"",IF({Event_Episode_ID}<>"",{Event_Episode_ID}&"/ONSET",""))' 13),
    (CF 'Anchor_ISO' 'Anchor (ISO 8601)' '=IF({A_Row}=0,"",INDEX(nTP_ISO,{A_Row}))' 20),
    (CF 'Anchor_Display' 'Anchor (as recorded)' '=IF({A_Row}=0,"",INDEX(nTP_Display,{A_Row}))' 28),
    (CF 'End_ISO' 'End point (ISO 8601)' '=IF({O_Row}=0,"",INDEX(nTP_ISO,{O_Row}))' 20),
    (CF 'End_Display' 'End point (as recorded)' '=IF({O_Row}=0,"",INDEX(nTP_Display,{O_Row}))' 28),
    (CF 'Exposure_Episode_ID' 'Exposure_Episode_ID (from link)' '=IF({LinkRow}=0,"",INDEX(nLNK_Exposure_Episode_ID,{LinkRow})&"")' 10),
    (CF 'Event_Episode_ID' 'Event Episode_ID (from link)' '=IF({LinkRow}=0,"",INDEX(nLNK_Event_Episode_ID,{LinkRow})&"")' 10),
    (CF 'Parameter_Set' 'Parameter_Set (anchor and end point: value, precision, status, original text, source)' $IntParamSet 40),
    (CF 'Assumptions' 'Assumptions (separated by " | ")' $IntAssume 40),
    (CF 'Limitations' 'Limitations (CALC-SPEC 1.1 Section 3)' ('=IF({K_Ok}=0,"","' + $LIMIT_C + '"&IF({Mode}="C4","' + $LIMIT_C4 + '","")&IF({f_C7}=1,"' + $LIMIT_C7 + '",""))') 30),
    (CF 'Interventions_Counted' 'C7: interventions counted (Milestone_IDs)' ('=IF(OR({K_Ok}=0,{f_C7}=0,{IntervCount}=0),"",{IV_IDs})') 16),
    (CF 'Calculation_Version' 'Calculation_Version' ('=IF({RowEmpty}=1,"","' + $SpecVersion + '")') 11),

    (CF 'RowEmpty' 'work: row empty' '=IF(AND({Interval_ID}="",{Link_ID}="",{Anchor_Type}="",{End_Record_ID}=""),1,0)' 6 '0' 'WORKING COLUMNS (locked; shown for inspection)'),
    (CF 'CaseKey' 'work: case key' '={Case_ID}&"|"&{Case_Version}' 10),
    (CF 'Key' 'work: interval key' '=IF({Interval_ID}="","",{CaseKey}&"|"&{Interval_ID})' 14),
    (CF 'Dup' 'work: key count' '=IF({Key}="",0,COUNTIF(nINT_Key,{Key}))' 5 '0'),
    (CF 'LinkRow' 'work: link row' '=IF({Link_ID}="",0,IFERROR(MATCH({CaseKey}&"|"&{Link_ID},nLNK_Key,0),0))' 5 '0'),
    (CF 'Lnk_ExpOK' 'work: linked exposure exists' '=IF({LinkRow}=0,1,INDEX(nLNK_Exp_Found,{LinkRow}))' 5 '0'),
    (CF 'Lnk_EvtOK' 'work: linked event exists' '=IF({LinkRow}=0,1,INDEX(nLNK_Evt_Found,{LinkRow}))' 5 '0'),
    (CF 'ExpDup' 'work: exposure episode ID count' '=IF({Exposure_Episode_ID}="",0,COUNTIF(nEXP_Key,{CaseKey}&"|"&{Exposure_Episode_ID}))' 5 '0'),
    (CF 'EvtDup' 'work: event episode ID count' '=IF({Event_Episode_ID}="",0,COUNTIF(nEVT_Key,{CaseKey}&"|"&{Event_Episode_ID}))' 5 '0'),
    (CF 'TypeOK' 'work: anchor type listed' '=IF(ISNA(MATCH({Anchor_Type},nLIST_AnchorType,0)),0,1)' 5 '0'),
    (CF 'AutoC5' 'work: C5 selection requested' '=IF(AND({TypeOK}=1,{Anchor_Type}="LAST_DOSE_BEFORE_ONSET",{Anchor_Record_ID}=""),1,0)' 5 '0'),
    (CF 'End_Is_Onset' 'work: end point is the onset' '=IF(RIGHT({End_Record_ID_Used},6)="/ONSET",1,0)' 5 '0'),
    (CF 'O_Row' 'work: end-point row in Timepoints' '=IF(OR({RowEmpty}=1,{End_Record_ID_Used}=""),0,IFERROR(MATCH({CaseKey}&"|"&{End_Record_ID_Used},nTP_Key,0),0))' 6 '0'),
    (LkO 'Dup' 'Dup' '0'), (LkO 'Type' 'Record_Type'), (LkO 'Prec' 'Precision'), (LkO 'Err' 'Err'),
    (LkO 'HT' 'HT' '0'), (LkO 'OK' 'OK' '0'), (LkO 'Loc' 'Loc'), (LkO 'UTC' 'UTC'), (LkO 'Lo' 'Lo'), (LkO 'Hi' 'Hi'),
    (LkO 'Text' 'Original_Text'), (LkO 'Src' 'Source'), (LkO 'Lab_Note' 'Lab_Note'),
    (CF 'O_Res' 'work: end point resolved' '=IF(AND({O_Row}>0,{O_Err}=""),1,0)' 5 '0'),
    (CF 'C5_CK' 'work: C5 candidate key' '={CaseKey}&"|"&{Exposure_Episode_ID}' 12),
    (CF 'C5_NAll' 'work: C5 candidates' '=IF({AutoC5}=0,0,COUNTIF(nTP_CandKey,{C5_CK}))' 5 '0'),
    (CF 'C5_NValid' 'work: C5 candidates resolved' '=IF({AutoC5}=0,0,COUNTIFS(nTP_CandKey,{C5_CK},nTP_Err,""))' 5 '0'),
    (CF 'C5_ErrTxt' 'work: C5 candidate date errors' ('=IF({C5_NAll}={C5_NValid},"",' + ((@('E_INVALID_DATE', 'E_REVERSED_BOUNDS', 'E_MISSING_INPUT', 'E_INVALID_UNIT') | ForEach-Object { 'IF(COUNTIFS(nTP_CandKey,{C5_CK},nTP_Err,"' + $_ + '")>0,"; ' + $_ + '","")' }) -join '&') + ')') 12),
    (CF 'C5_Go' 'work: C5 selection runs' '=IF(AND({AutoC5}=1,{C5_NValid}>0,{O_Res}=1),1,0)' 5 '0'),
    (CF 'C5_NUnc' 'work: C5 positions uncertain' ('=IF({C5_Go}=0,0,SUMPRODUCT(' + $C5UNC + '))') 5 '0'),
    (CF 'C5_NEl' 'work: C5 eligible' ('=IF({C5_Go}=0,0,SUMPRODUCT(' + $C5EL + '))') 5 '0'),
    (CF 'C5_NPB' 'work: C5 partial-date candidates not after the end point' ('=IF({C5_Go}=0,0,SUMPRODUCT(' + $C5PB + '))') 5 '0'),
    (CF 'C5_MaxLo' 'work: latest eligible date' ('=IF({C5_NEl}=0,0,SUMPRODUCT(MAX(' + $C5EL + '*nTP_Lo_n)))') 8 '0'),
    (CF 'C5_CntMaxLo' 'work: eligible on that date' ('=IF({C5_NEl}=0,0,SUMPRODUCT(' + $C5EL + '*(nTP_Lo_n={C5_MaxLo})))') 5 '0'),
    (CF 'C5_DAtMax' 'work: date-only eligible on that date' ('=IF({C5_NEl}=0,0,SUMPRODUCT(' + $C5EL + '*(1-nTP_HT)*(nTP_Lo_n={C5_MaxLo})))') 5 '0'),
    (CF 'C5_MaxLoD' 'work: latest date-only eligible' ('=IF({C5_NEl}=0,0,SUMPRODUCT(MAX(' + $C5EL + '*(1-nTP_HT)*nTP_Lo_n)))') 8 '0'),
    (CF 'C5_MaxUTC' 'work: latest eligible UTC timestamp' ('=IF({C5_NEl}=0,0,SUMPRODUCT(MAX(' + $C5EL + '*nTP_HT*nTP_OK*nTP_UTC_n)))') 10 '0'),
    (CF 'C5_CntUTC' 'work: eligible at that UTC instant' ('=IF({C5_MaxUTC}=0,0,SUMPRODUCT(' + $C5EL + '*nTP_HT*nTP_OK*(nTP_UTC_n={C5_MaxUTC})))') 5 '0'),
    (CF 'C5_LocAtUTC' 'work: its local clock' ('=IF({C5_CntUTC}<>1,0,SUMPRODUCT(' + $C5EL + '*nTP_HT*nTP_OK*(nTP_UTC_n={C5_MaxUTC})*nTP_Loc_n))') 10 '0'),
    (CF 'C5_LoAtUTC' 'work: its date' ('=IF({C5_CntUTC}<>1,0,SUMPRODUCT(' + $C5EL + '*nTP_HT*nTP_OK*(nTP_UTC_n={C5_MaxUTC})*nTP_Lo_n))') 8 '0'),
    (CF 'C5_MaxLocN' 'work: latest eligible local timestamp (offset unknown)' ('=IF({C5_NEl}=0,0,SUMPRODUCT(MAX(' + $C5EL + '*nTP_HT*(1-nTP_OK)*nTP_Loc_n)))') 10 '0'),
    (CF 'C5_CntLocN' 'work: eligible at that local time' ('=IF({C5_MaxLocN}=0,0,SUMPRODUCT(' + $C5EL + '*nTP_HT*(1-nTP_OK)*(nTP_Loc_n={C5_MaxLocN})))') 5 '0'),
    (CF 'C5_LoAtLocN' 'work: its date' ('=IF({C5_CntLocN}<>1,0,SUMPRODUCT(' + $C5EL + '*nTP_HT*(1-nTP_OK)*(nTP_Loc_n={C5_MaxLocN})*nTP_Lo_n))') 8 '0'),
    (CF 'C5_MaxLocOK' 'work: latest local clock of offset-known timestamps' ('=IF({C5_NEl}=0,0,SUMPRODUCT(MAX(' + $C5EL + '*nTP_HT*nTP_OK*nTP_Loc_n)))') 10 '0'),
    (CF 'C5_Case' 'work: C5 latest candidate (1 date, 2 UTC timestamp, 3 local timestamp, 0 not determined)' '=IF({C5_NEl}=0,0,IF({C5_DAtMax}>0,IF({C5_CntMaxLo}=1,1,0),IF(AND({C5_MaxUTC}>0,{C5_CntUTC}=1,{C5_LocAtUTC}>{C5_MaxLocN},{C5_LoAtUTC}>{C5_MaxLoD}),2,IF(AND({C5_MaxLocN}>0,{C5_CntLocN}=1,{C5_MaxLocN}>{C5_MaxLocOK},{C5_LoAtLocN}>{C5_MaxLoD}),3,0))))' 5 '0'),
    (CF 'C5_Row' 'work: C5 row in Timepoints' ('=IF({C5_Case}=0,0,IF({C5_Case}=1,MATCH(1,INDEX(' + $C5EL + '*(1-nTP_HT)*(nTP_Lo_n={C5_MaxLo}),0),0),IF({C5_Case}=2,MATCH(1,INDEX(' + $C5EL + '*nTP_HT*nTP_OK*(nTP_UTC_n={C5_MaxUTC}),0),0),MATCH(1,INDEX(' + $C5EL + '*nTP_HT*(1-nTP_OK)*(nTP_Loc_n={C5_MaxLocN}),0),0))))') 5 '0'),
    (CF 'C5_BestDay' 'work: selected date' '=IF({C5_Row}=0,0,INDEX(nTP_Lo_n,{C5_Row}))' 8 '0'),
    (CF 'C5_PBlock' 'work: partial-date candidates that could be the last dose' ('=IF({C5_Row}=0,0,SUMPRODUCT(' + $C5PB + '*(nTP_Hi_n>={C5_BestDay})))') 5 '0'),
    (CF 'C5_Fail' 'work: C5 outcome' '=IF({AutoC5}=0,"",IF({C5_NValid}=0,"none",IF({C5_Go}=0,"",IF({C5_NUnc}>0,"uncertain",IF({C5_NEl}=0,"noneligible",IF({C5_Case}=0,"tie",IF({C5_PBlock}>0,"partial","")))))))' 9),
    (CF 'C5_Selected' 'work: C5 selected record' '=IF(AND({AutoC5}=1,{C5_Go}=1,{C5_Fail}="",{C5_Row}>0),INDEX(nTP_Record_ID,{C5_Row})&"","")' 10),
    (CF 'C5_Needs' 'work: C5 [input]' ('=IF(OR({C5_Fail}="",{AutoC5}=0),"",IF({C5_Fail}="none","an administration record (last dose before onset)",IF({C5_Fail}="uncertain","the anchor record: "&INDEX(nTP_Record_ID,MATCH(1,INDEX(' + $C5UNC + ',0),0))&IF({C5_NUnc}>1," and "&({C5_NUnc}-1)&" more","")&" may fall before or after the end point, so the last dose cannot be identified from the recorded dates",IF({C5_Fail}="noneligible","the anchor record: no administration with an exact date is recorded at or before the end point",IF({C5_Fail}="tie","the anchor record: the order of the latest administrations cannot be told from the recorded dates and times","the anchor record: partial-date administration "&INDEX(nTP_Record_ID,MATCH(1,INDEX(' + $C5PB + '*(nTP_Hi_n>={C5_BestDay}),0),0))&" could be the last dose before the end point")))))') 16),
    (CF 'A_Row' 'work: anchor row in Timepoints' '=IF(OR({RowEmpty}=1,{Anchor_Record_ID_Used}=""),0,IFERROR(MATCH({CaseKey}&"|"&{Anchor_Record_ID_Used},nTP_Key,0),0))' 6 '0'),
    (LkA 'Dup' 'Dup' '0'), (LkA 'Type' 'Record_Type'), (LkA 'Prec' 'Precision'), (LkA 'Err' 'Err'),
    (LkA 'HT' 'HT' '0'), (LkA 'OK' 'OK' '0'), (LkA 'Loc' 'Loc'), (LkA 'UTC' 'UTC'), (LkA 'Lo' 'Lo'), (LkA 'Hi' 'Hi'),
    (LkA 'Text' 'Original_Text'), (LkA 'Src' 'Source'),
    (CF 'A_Res' 'work: anchor resolved' '=IF(AND({A_Row}>0,{A_Err}=""),1,0)' 5 '0'),
    (CF 'Anc_Err' 'work: anchor E-codes' '=IF(OR({RowEmpty}=1,{TypeOK}=0),"",IF({AutoC5}=1,IF({LinkRow}=0,"",{C5_ErrTxt}&IF({C5_Fail}="","","; E_MISSING_INPUT")),IF({A_Row}=0,IF(AND({LinkRow}=0,{Anchor_Record_ID}="",{Anchor_Type}="FIRST_ADMIN"),"","; E_MISSING_INPUT"),IF({A_Dup}>1,"; E_DUPLICATE_ID","")&IF({A_Err}<>"","; "&{A_Err},""))))' 12),
    (CF 'Anc_Needs' 'work: anchor [input]' '=IF(OR({RowEmpty}=1,{TypeOK}=0),"",IF({AutoC5}=1,IF({LinkRow}=0,"",{C5_Needs}),IF({A_Row}=0,IF({Anchor_Record_ID_Used}="",IF({Anchor_Type}="FIRST_ADMIN",IF({LinkRow}=0,"","the first administration (the Start of the linked exposure episode)"),"the anchor record"),"an existing, dated anchor record ("&{Anchor_Record_ID_Used}&")"),IF(ISNUMBER(SEARCH("E_MISSING_INPUT",{A_Err})),"the date fields required by the precision of anchor record "&{Anchor_Record_ID_Used},""))))' 14),
    (CF 'End_Err' 'work: end-point E-codes' '=IF({RowEmpty}=1,"",IF({O_Row}=0,IF(AND({LinkRow}=0,{End_Record_ID}=""),"","; E_MISSING_INPUT"),IF({O_Dup}>1,"; E_DUPLICATE_ID","")&IF({O_Err}<>"","; "&{O_Err},"")))' 12),
    (CF 'End_Needs' 'work: end-point [input]' '=IF({RowEmpty}=1,"",IF({O_Row}=0,IF(AND({LinkRow}=0,{End_Record_ID}=""),"",IF({End_Record_ID_Used}="","an end-point record (End_Record_ID, or a link to an event episode)","an existing, dated end-point record ("&{End_Record_ID_Used}&")")),IF(ISNUMBER(SEARCH("E_MISSING_INPUT",{O_Err})),IF({O_Prec}="LAB","the dated last normal and first abnormal tests of the laboratory-defined onset ("&{End_Record_ID_Used}&"); no other date is used in their place","the date fields required by the precision of end-point record "&{End_Record_ID_Used}),"")))' 14),
    (CF 'Mode_Finest' 'work: finest rule the records support' '=IF(OR({A_Res}=0,{O_Res}=0),"",IF({A_HT}*{O_HT}=1,"C1",IF(AND(OR({A_Prec}="TIMESTAMP",{A_Prec}="DATE"),OR({O_Prec}="TIMESTAMP",{O_Prec}="DATE")),"C2","C4")))' 6),
    (CF 'Mode_Bad' 'work: requested rule finer than the records' '=IF(OR({Mode_Finest}="",{Calculation_Mode}=""),0,IF(ISNA(MATCH({Calculation_Mode},nLIST_Mode,0)),1,IF(AND({Calculation_Mode}="C1",{Mode_Finest}<>"C1"),1,IF(AND({Calculation_Mode}="C2",{Mode_Finest}="C4"),1,0))))' 5 '0'),
    (CF 'Mode' 'work: rule applied (C1, C2 or C4)' '=IF(OR({Mode_Finest}="",{Mode_Bad}=1),"",IF({Calculation_Mode}="",{Mode_Finest},{Calculation_Mode}&""))' 6),
    (CF 'f_C7' 'C7 recovery interval (CESSATION to a RECOVERY milestone)' '=IF(AND({Anchor_Type}="CESSATION",{O_Type}="MILESTONE:RECOVERY"),1,0)' 4 '0'),
    (CF 'IntvErr' 'work: C7 intervention date errors' ('=IF({f_C7}=0,"",' + ((@('E_INVALID_DATE', 'E_REVERSED_BOUNDS', 'E_MISSING_INPUT', 'E_INVALID_UNIT') | ForEach-Object { 'IF(COUNTIFS(nTP_Case_Key,{CaseKey},nTP_Record_Type,"MILESTONE:INTERVENTION",nTP_Err,"' + $_ + '")>0,"; ' + $_ + '","")' }) -join '&') + ')') 12),
    (CF 'IntvUndated' 'work: C7 interventions without a date' ('=IF({f_C7}=0,0,SUMPRODUCT(' + $C7UND + '))') 5 '0'),
    (CF 'ErrPre' 'work: E-codes found (any order)' '=IF({RowEmpty}=1,"",IF({Interval_ID}="","; E_MISSING_INPUT","")&IF({Dup}>1,"; E_DUPLICATE_ID","")&IF({LinkRow}=0,"; E_MISSING_LINK",IF(OR({Lnk_ExpOK}=0,{Lnk_EvtOK}=0),"; E_MISSING_INPUT",""))&IF(OR({ExpDup}>1,{EvtDup}>1),"; E_DUPLICATE_ID","")&IF({TypeOK}=0,"; E_MISSING_INPUT",IF(AND({Anchor_Type}="OTHER",TRIM({Anchor_Description}&"")=""),"; E_MISSING_INPUT",""))&{End_Err}&{Anc_Err}&{IntvErr}&IF({Mode_Bad}=1,"; E_MISSING_INPUT",""))' 14),
    (CF 'PreErr' 'work: E-codes before calculation (G9 order)' ('=IF({RowEmpty}=1,"",' + (ORDE '{ErrPre}' $ECODES_INT) + ')') 14),
    (CF 'K_TZU' 'work: offset unknown (1/0)' '=IF({A_OK}*{O_OK}=1,0,1)' 5 '0'),
    (CF 'V_Min' 'work: C1 minutes' '=IF(AND({PreErr}="",{A_Res}=1,{O_Res}=1,{A_HT}*{O_HT}=1),IF({K_TZU}=0,{O_UTC}-{A_UTC},{O_Loc}-{A_Loc}),"")' 8),
    (CF 'V_Days' 'work: C2 calendar days' '=IF(AND({PreErr}="",{Mode}="C2"),{O_Lo}-{A_Lo},"")' 7),
    (CF 'V_Min4' 'work: C4 minimum' '=IF(AND({PreErr}="",{Mode}="C4"),{O_Lo}-{A_Hi},"")' 7),
    (CF 'V_Max4' 'work: C4 maximum' '=IF(AND({PreErr}="",{Mode}="C4"),{O_Hi}-{A_Lo},"")' 7),
    (CF 'C6_On' 'work: reported timing given (C6)' '=IF(AND({f_C7}=0,OR(TRIM({Reported_Text}&"")<>"",{Reported_Lower}<>"")),1,0)' 4 '0'),
    (CF 'C6_Factor' 'work: C6 unit factor (0 = not comparable)' '=IF({Mode}="C1",IF({Reported_Unit}="minutes",1,IF({Reported_Unit}="hours",60,IF({Reported_Unit}="days",1440,0))),IF(OR({Mode}="C2",{Mode}="C4"),IF({Reported_Unit}="calendar days",1,0),0))' 6 '0'),
    (CF 'C6_Code' 'work: C6 result code' '=IF({C6_On}=0,"",IF(OR(NOT(ISNUMBER({Reported_Lower})),NOT(ISNUMBER({Reported_Upper}))),"I_APPROXIMATE_NOT_COMPARED",IF(AND({Reported_Unit}<>"",ISNA(MATCH({Reported_Unit},nLIST_ReportedUnit,0))),"E_INVALID_UNIT",IF({Reported_Unit}="","E_MISSING_INPUT",IF({Reported_Lower}>{Reported_Upper},"E_REVERSED_BOUNDS",IF(OR({Reported_Unit}="weeks",{Reported_Unit}="months",{Reported_Unit}="years"),"I_APPROXIMATE_NOT_COMPARED",IF({PreErr}<>"","",IF({C6_Factor}=0,"I_APPROXIMATE_NOT_COMPARED",IF(AND({Reported_Lower}*{C6_Factor}<=IF({Mode}="C4",{V_Max4},IF({Mode}="C1",{V_Min},{V_Days})),{Reported_Upper}*{C6_Factor}>=IF({Mode}="C4",{V_Min4},IF({Mode}="C1",{V_Min},{V_Days}))),"OVERLAP","W_REPORTED_TIMING_CONFLICT")))))))))' 14),
    (CF 'K_Err' 'work: blocking E-codes (G9 order)' '=IF({RowEmpty}=1,"",IF({PreErr}<>"",{PreErr},IF(LEFT({C6_Code},2)="E_",{C6_Code},"")))' 14),
    (CF 'K_Needs' 'work: [input] for E_MISSING_INPUT' '=IF(ISERROR(SEARCH("E_MISSING_INPUT",{K_Err})),"",MID(IF({Interval_ID}="",", an Interval_ID","")&IF(AND({LinkRow}>0,{Lnk_ExpOK}=0),", an existing drug exposure episode for link "&{Link_ID},"")&IF(AND({LinkRow}>0,{Lnk_EvtOK}=0),", an existing event episode for link "&{Link_ID},"")&IF({TypeOK}=0,", an anchor type",IF(AND({Anchor_Type}="OTHER",TRIM({Anchor_Description}&"")=""),", a description of the OTHER anchor",""))&IF({End_Needs}<>"",", "&{End_Needs},"")&IF({Anc_Needs}<>"",", "&{Anc_Needs},"")&IF(ISNUMBER(SEARCH("E_MISSING_INPUT",{IntvErr})),", a complete date for every recorded intervention of the case version","")&IF({Mode_Bad}=1,IF({Calculation_Mode}="C1",", a time of day and UTC offset (or unknown) on both records: C1 needs two timestamps; leave the rule blank or choose C2 or C4",IF({Calculation_Mode}="C2",", an exact date on both records: C2 needs two exact dates; leave the rule blank or choose C4",", a listed rule (C1, C2 or C4), or blank for the finest")),"")&IF(AND({PreErr}="",{C6_Code}="E_MISSING_INPUT"),", the unit of the reported timing",""),3,999))' 14),
    (CF 'K_Ok' 'work: result shown (1/0)' '=IF(AND({RowEmpty}=0,{K_Err}=""),1,0)' 4 '0'),
    (CF 'Treatment_Day_Used' 'work: C3 applied' '=IF(AND({K_Ok}=1,{Mode}="C2",{Treatment_Day_Convention}="inclusive",{f_C7}=0),1,0)' 4 '0'),
    (CF 'K_HasC' 'work: constraint entered' '=IF(AND({K_Ok}=1,{Mode}="C4",OR(ISNUMBER({Constraint_Not_Before}),ISNUMBER({Constraint_Not_After}))),1,0)' 5 '0'),
    (CF 'K_CLo' 'work: constrained anchor lower' '=IF({K_HasC}=0,"",IF(ISNUMBER({Constraint_Not_Before}),MAX({A_Lo},{Constraint_Not_Before}),{A_Lo}))' 10 'd mmmm yyyy'),
    (CF 'K_CHi' 'work: constrained anchor upper' '=IF({K_HasC}=0,"",IF(ISNUMBER({Constraint_Not_After}),MIN({A_Hi},{Constraint_Not_After}),{A_Hi}))' 10 'd mmmm yyyy'),
    (CF 'K_CErr' 'work: constrained result E-codes' '=IF({K_CLo}="","",MID(IF({K_CLo}>{K_CHi},"; E_REVERSED_BOUNDS","")&IF(TRIM({Constraint_Source}&"")="","; E_MISSING_INPUT",""),3,99))' 10),
    (CF 'IntervCount' 'work: C7 interventions counted' ('=IF(OR({f_C7}=0,{PreErr}<>""),0,SUMPRODUCT(' + $C7IM + '*' + $C7WIN + '))') 6 '0'),
    (CF 'TxNotCounted' 'work: C7 treatment exposures in the window (pointed out, never counted)' ('=IF(OR({f_C7}=0,{PreErr}<>""),0,SUMPRODUCT(' + $C7TX + '*' + $C7WIN + '))') 6 '0')) + @(TPRows 'IV' ($C7IM + '*' + $C7WIN) '{IntervCount}') + @(TPRows 'TX' ($C7TX + '*' + $C7WIN) '{TxNotCounted}') + @(TPRows 'PB' $C5PB '{C5_NPB}') + @(
    (CF 'IV_List' 'work: interventions counted (with labels)' (TPL 'IV' '{IntervCount}') 16),
    (CF 'IV_IDs' 'work: interventions counted (IDs)' (TPL 'IV' '{IntervCount}' $false) 12),
    (CF 'TX_List' 'work: treatment exposures pointed out' (TPL 'TX' '{TxNotCounted}') 16),
    (CF 'PB_List' 'work: earlier partial-date administrations' (TPL 'PB' '{C5_NPB}' $false) 12),
    (CF 'f_TZ' 'flag W_TZ_UNKNOWN' '=IF(AND({K_Ok}=1,{K_TZU}=1,{V_Min}<>"",OR({Mode}="C1",AND({Mode}="C2",{V_Days}=0))),1,0)' 4 '0'),
    (CF 'f_SameDate' 'flag W_SAME_DATE_ORDER_UNKNOWN' '=IF(AND({K_Ok}=1,OR(AND({Mode}="C2",{V_Days}=0,{V_Min}=""),AND({Mode}="C4",{V_Min4}=0,{V_Max4}=0))),1,0)' 4 '0'),
    (CF 'f_SeqUnc' 'flag W_SEQUENCE_UNCERTAIN' '=IF({Sequence_Status}="SEQUENCE_UNCERTAIN",1,0)' 4 '0'),
    (CF 'f_AdminSame' 'flag W_ADMIN_SAME_DATE_AS_ONSET' '=IF(AND({K_Ok}=1,{AutoC5}=1,{A_HT}*{O_HT}=0,{A_Lo}={O_Lo},{O_Lo}={O_Hi}),1,0)' 4 '0'),
    (CF 'f_Conflict' 'flag W_REPORTED_TIMING_CONFLICT' '=IF(AND({K_Ok}=1,{C6_Code}="W_REPORTED_TIMING_CONFLICT"),1,0)' 4 '0'),
    (CF 'f_Interv' 'flag W_INTERVENTION_BEFORE_RECOVERY' '=IF(AND({K_Ok}=1,{f_C7}=1,{IntervCount}>0),1,0)' 4 '0'),
    (CF 'f_TDundef' 'flag I_TREATMENT_DAY_UNDEFINED' '=IF(AND({Treatment_Day_Used}=1,{V_Days}<0),1,0)' 4 '0'),
    (CF 'f_Approx' 'flag I_APPROXIMATE_NOT_COMPARED' '=IF(AND({K_Ok}=1,{C6_Code}="I_APPROXIMATE_NOT_COMPARED"),1,0)' 4 '0')
)

# ================================================================== RECORD SHEET COLUMNS
function RecCheckMsg($idCol, $idLabel) {
    '=IF({Record_Check}="","",{Record_Check}&": "&' + (MSGX '{Record_Check}' ('IF({' + $idCol + '}="","' + $idLabel + '","the date fields required by the selected precision")')) + ')'
}
$CaseCols = @(
    (CIn 'Case_ID' 'Case_ID' 11 $null '@' 'INPUTS: one row per case version (follow-up adds a version; earlier rows stay)'),
    (CIn 'Case_Version' 'Case_Version' 8 'posint' '0'),
    (CIn 'Received_Date' 'Received date' 15 'date' 'd mmmm yyyy'),
    (CIn 'Setting' 'Setting' 24 $null '@'),
    (CIn 'Reviewer' 'Reviewer' 14 $null '@'),
    (CIn 'Review_Date' 'Review date' 15 'date' 'd mmmm yyyy'),
    (CIn 'Fictional' 'Fictional or de-identified teaching case?' 10 'list:nLIST_YesNo'),
    (CIn 'Summary' 'Summary' 50 $null '@'),
    (CIn 'Version_Note' 'What this version adds' 50 $null '@'),
    (CF 'Key' 'Case key' '=IF({Case_ID}="","",{Case_ID}&"|"&{Case_Version})' 12 $null 'CHECKS'),
    (CF 'Dup' 'Key count' '=IF({Key}="",0,COUNTIF(nCASE_Key,{Key}))' 6 '0'),
    (CF 'Record_Check' 'Record check' '=IF({Key}="","",IF({Case_Version}="","E_MISSING_INPUT",IF({Dup}>1,"E_DUPLICATE_ID","")))' 14),
    (CF 'Check_Message' 'Check message' ('=IF({Record_Check}="","",{Record_Check}&": "&' + (MSGX '{Record_Check}' '"a Case_Version"') + ')') 40)
)
$ExpCols = @(
    (CIn 'Case_ID' 'Case_ID' 10 $null '@' 'INPUTS: one row per exposure episode'),
    (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
    (CIn 'Drug_ID' 'Drug_ID' 11 $null '@'),
    (CIn 'Exposure_Episode_ID' 'Exposure_Episode_ID (timepoints: ID/START, ID/STOP, ID/RESTART)' 11 $null '@'),
    (CIn 'Product' 'Product' 24 $null '@'),
    (CIn 'Role' 'Role' 12 'list:nLIST_Role'),
    (CIn 'Route' 'Route' 14 $null '@'),
    (CIn 'Dose' 'Dose (as reported)' 8 'num'),
    (CIn 'Dose_Unit' 'Dose unit' 10 'list:nLIST_DoseUnit'),
    (CIn 'Regimen_Text' 'Regimen (as reported)' 26 $null '@'),
    (CIn 'Missing_Reason' 'Missing-data reason (example codes)' 12 'listwarn:nLIST_MissingReason'),
    (CIn 'Notes' 'Notes' 22 $null '@')
) + @(TPBlock 'S' 'Start') + @(TPBlock 'T' 'Stop') + @(TPBlock 'R' 'Restart') + @(
    (CF 'Key' 'Exposure key' '=IF({Exposure_Episode_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Exposure_Episode_ID})' 14 $null 'CHECKS'),
    (CF 'Dup' 'Key count' '=IF({Key}="",0,COUNTIF(nEXP_Key,{Key}))' 5 '0'),
    (CF 'Record_Check' 'Record check (E-code)' '=IF({Exposure_Episode_ID}="",IF(COUNTA({Drug_ID},{Product},{S_Prec},{S_Date})>0,"E_MISSING_INPUT",""),IF({Dup}>1,"E_DUPLICATE_ID",IF({S_Err}<>"",{S_Err},IF({T_Err}<>"",{T_Err},{R_Err}))))' 14),
    (CF 'Check_Message' 'Check message' (RecCheckMsg 'Exposure_Episode_ID' 'an Exposure_Episode_ID') 40)
)
$SegErr = '=IF(COUNTA({End_Date}:{End_Off})=0,"",IF(OR({End_Date}="",{End_Time}="",{End_Off}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({End_Date})),NOT(ISNUMBER({End_Time}))),"E_INVALID_DATE",IF(OR({End_Date}<>INT({End_Date}),{End_Time}<0,{End_Time}>=1),"E_INVALID_DATE",IF(ISNA(MATCH({End_Off},nLIST_Offset,0)),"E_INVALID_DATE","")))))'
$AdmCols = @(
    (CIn 'Case_ID' 'Case_ID' 10 $null '@' 'INPUTS: one row per dose, administration or infusion rate segment'),
    (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
    (CIn 'Admin_ID' 'Admin_ID (record ID)' 12 $null '@'),
    (CIn 'Exposure_Episode_ID' 'Exposure_Episode_ID' 11 $null '@'),
    (CIn 'Kind' 'Kind (DOSE, DOSE_CHANGE, INFUSION_SEGMENT, OTHER); INFUSION_SEGMENT rows are Q9 rate segments, not C5 candidates' 13 'list:nLIST_AdminKind'),
    (CIn 'Dose' 'Dose (Q10 needs mg)' 8 'num'),
    (CIn 'Dose_Unit' 'Dose unit' 9 'list:nLIST_DoseUnit'),
    (CIn 'Rate' 'Infusion rate (Q9 needs mg/hour)' 8 'num'),
    (CIn 'Rate_Unit' 'Rate unit' 9 'list:nLIST_RateUnit'),
    (CIn 'BSA_m2' ('BSA at this administration (m' + [char]0x00B2 + ')') 8 'num'),
    (CIn 'Missing_Reason' 'Missing-data reason (example codes)' 12 'listwarn:nLIST_MissingReason'),
    (CIn 'Notes' 'Notes' 24 $null '@')
) + @(TPBlock 'A' 'Administration or segment start') + @(
    (CIn 'End_Date' 'Segment end: date' 15 'date' 'd mmmm yyyy' 'Infusion segment end (INFUSION_SEGMENT only)'),
    (CIn 'End_Time' 'Segment end: local time hh:mm' 9 'time' 'hh:mm'),
    (CIn 'End_Off' 'Segment end: UTC offset or unknown' 9 'list:nLIST_Offset' '@'),
    (CIn 'End_Src' 'Segment end: source' 18 $null '@'),
    (CF 'SE_Err' 'Segment end: check' $SegErr 10),
    (CF 'SE_Loc' 'Segment end: local (minutes)' '=IF(AND({SE_Err}="",ISNUMBER({End_Date}),ISNUMBER({End_Time})),{End_Date}*1440+ROUND({End_Time}*1440,0),"")' 10 '0'),
    (CF 'SE_OK' 'Segment end: offset known' '=IF(AND({SE_Loc}<>"",{End_Off}<>"unknown"),1,0)' 5 '0'),
    (CF 'SE_UTC' 'Segment end: UTC (minutes)' ('=IF({SE_OK}=1,{SE_Loc}-' + (OFFMIN '{End_Off}') + ',"")') 10 '0'),
    (CF 'Key' 'Administration key' '=IF({Admin_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Admin_ID})' 14 $null 'WORKING AND CHECK COLUMNS'),
    (CF 'Dup' 'Key count' '=IF({Key}="",0,COUNTIF(nADM_Key,{Key}))' 5 '0'),
    (CF 'SelKey' 'Episode key' '=IF({Admin_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Exposure_Episode_ID})' 14),
    (CF 'Valid' 'Exact date or timestamp, no error' '=IF(AND({Admin_ID}<>"",{A_Err}="",{A_Lo}<>"",OR({A_Prec}="TIMESTAMP",{A_Prec}="DATE")),1,0)' 5 '0'),
    (CF 'A_Dated' 'Date recorded (any precision)' '=IF(AND({Admin_ID}<>"",COUNTA({A_Prec}:{A_Upper})>0),1,0)' 5 '0'),
    (CF 'Loc_n' 'Local minutes (0 if none)' '=IF({A_HT}=1,{A_Loc},0)' 9 '0'),
    (CF 'UTC_n' 'UTC minutes (0 if none)' '=IF({A_OK}=1,{A_UTC},0)' 9 '0'),
    (CF 'Lo_n' 'Date (0 if none)' '=IF({Valid}=1,{A_Lo},0)' 7 '0'),
    (CF 'RowN' 'Row order (as entered)' '=ROW()' 6 '0'),
    (CF 'Rate_n' 'Rate (0 if none)' '=IF(ISNUMBER({Rate}),{Rate},0)' 6),
    (CF 'Dose_n' 'Dose (0 if none)' '=IF(ISNUMBER({Dose}),{Dose},0)' 6),
    (CF 'Dose_pos_n' 'Positive dose (0 otherwise)' '=IF(AND(ISNUMBER({Dose}),{Dose}>0),{Dose},0)' 6),
    (CF 'HasDose' 'Has dose' '=IF(AND({Admin_ID}<>"",ISNUMBER({Dose})),1,0)' 5 '0'),
    (CF 'BSA_n' 'BSA (0 if none)' '=IF(ISNUMBER({BSA_m2}),{BSA_m2},0)' 6),
    (CF 'BSA_key' 'BSA key (-1 blank, -2 not a number)' '=IF(ISNUMBER({BSA_m2}),{BSA_m2},IF({BSA_m2}="",-1,-2))' 6),
    (CF 'HasBSA' 'Has BSA' '=IF(AND({Admin_ID}<>"",ISNUMBER({BSA_m2})),1,0)' 5 '0'),
    (CF 'BSA_missing' 'BSA blank (W_PARTIAL_BSA)' '=IF(AND({Admin_ID}<>"",{BSA_m2}=""),1,0)' 5 '0'),
    (CF 'Known' 'Dose and positive BSA recorded (Method A)' '=IF(AND({Admin_ID}<>"",ISNUMBER({Dose}),ISNUMBER({BSA_m2})),IF({BSA_m2}>0,1,0),0)' 5 '0'),
    (CF 'Dose_per_BSA' 'Dose / BSA, mg/m2 (0 unless Known)' '=IF({Known}=1,{Dose_n}/{BSA_m2},0)' 8),
    (CF 'Q10_Raw' 'Q10 checks of this administration (E-codes)' '=IF(OR({Admin_ID}="",{Kind}="INFUSION_SEGMENT"),"",IF({Dose_Unit}="","; E_MISSING_INPUT",IF({Dose_Unit}<>"mg","; E_INVALID_UNIT",""))&IF(NOT(ISNUMBER({Dose})),"; E_MISSING_INPUT",IF({Dose}<=0,"; E_NONPOSITIVE_PARAMETER",""))&IF(AND({BSA_m2}<>"",NOT(ISNUMBER({BSA_m2}))),"; E_MISSING_INPUT",IF(AND(ISNUMBER({BSA_m2}),{BSA_m2}<=0),"; E_NONPOSITIVE_PARAMETER","")))' 14),
    (CF 'Q10_DateErr' 'Q10 Method B date check (E-code)' '=IF(OR({Admin_ID}="",{Kind}="INFUSION_SEGMENT",{A_Dated}=0,{Valid}=1),"",IF({A_Err}<>"",{A_Err},"E_MISSING_INPUT"))' 10),
    (CF 'SegValid' 'Segment: start and end timestamps valid' '=IF(AND({Kind}="INFUSION_SEGMENT",{A_Err}="",{A_HT}=1,{SE_Err}="",{SE_Loc}<>""),1,0)' 5 '0'),
    (CF 'SegOK' 'Segment: both offsets known' '=IF(AND({SegValid}=1,{A_OK}=1,{SE_OK}=1),1,0)' 5 '0'),
    (CF 'SegS_UTC_n' 'Segment start UTC' '=IF({SegOK}=1,{A_UTC},0)' 9 '0'),
    (CF 'SegE_UTC_n' 'Segment end UTC' '=IF({SegOK}=1,{SE_UTC},0)' 9 '0'),
    (CF 'SegS_Loc_n' 'Segment start local' '=IF({SegValid}=1,{A_Loc},0)' 9 '0'),
    (CF 'SegE_Loc_n' 'Segment end local' '=IF({SegValid}=1,{SE_Loc},0)' 9 '0'),
    (CF 'Rev_Loc' 'Segment reversed (local clock)' '=IF(AND({SegValid}=1,{SegE_Loc_n}<{SegS_Loc_n}),1,0)' 4 '0'),
    (CF 'Rev_UTC' 'Segment reversed (UTC)' '=IF(AND({SegOK}=1,{SegE_UTC_n}<{SegS_UTC_n}),1,0)' 4 '0'),
    (CF 'Ovl_Loc' 'Overlapping segments, local clock (count)' '=IF({SegValid}=0,0,COUNTIFS(nADM_SelKey,{SelKey},nADM_SegValid,1,nADM_SegS_Loc_n,"<"&{SegE_Loc_n},nADM_SegE_Loc_n,">"&{SegS_Loc_n})-IF({SegE_Loc_n}>{SegS_Loc_n},1,0))' 5 '0'),
    (CF 'Ovl_UTC' 'Overlapping segments, UTC (count)' '=IF({SegOK}=0,0,COUNTIFS(nADM_SelKey,{SelKey},nADM_SegOK,1,nADM_SegS_UTC_n,"<"&{SegE_UTC_n},nADM_SegE_UTC_n,">"&{SegS_UTC_n})-IF({SegE_UTC_n}>{SegS_UTC_n},1,0))' 5 '0'),
    (CF 'Seg_Raw' 'Segment checks (Q9 E-codes)' '=IF({Kind}<>"INFUSION_SEGMENT","",IF({A_Err}<>"","; "&{A_Err},IF({A_HT}=0,"; E_MISSING_INPUT",""))&IF({SE_Err}<>"","; "&{SE_Err},IF({SE_Loc}="","; E_MISSING_INPUT",""))&IF({Rate_Unit}="","; E_MISSING_INPUT",IF({Rate_Unit}<>"mg/hour","; E_INVALID_UNIT",""))&IF(NOT(ISNUMBER({Rate})),"; E_MISSING_INPUT",IF({Rate}<=0,"; E_NONPOSITIVE_PARAMETER","")))' 14),
    (CF 'Ep_Found' 'Exposure episode exists' '=IF({Admin_ID}="",1,IF(COUNTIF(nEXP_Key,{SelKey})>0,1,0))' 5 '0'),
    (CF 'Record_Check' 'Record check (first E-code)' ('=IF({Admin_ID}="",IF(COUNTA({Exposure_Episode_ID},{Dose},{Rate},{A_Prec},{A_Date})>0,"E_MISSING_INPUT",""),IF({Dup}>1,"E_DUPLICATE_ID",IF({Ep_Found}=0,"E_MISSING_INPUT",IF({A_Err}<>"",{A_Err},IF({Seg_Raw}<>"",' + (FIRSTCODE 'MID({Seg_Raw},3,999)') + ',IF(AND(ISNUMBER({Dose}),{Dose}<=0),"E_NONPOSITIVE_PARAMETER",IF(AND(ISNUMBER({BSA_m2}),{BSA_m2}<=0),"E_NONPOSITIVE_PARAMETER","")))))))') 14),
    (CF 'Check_Message' 'Check message' ('=IF({Record_Check}="","",{Record_Check}&": "&' + (MSGX '{Record_Check}' 'IF({Admin_ID}="","an Admin_ID",IF({Ep_Found}=0,"an Exposures row for episode "&{Exposure_Episode_ID},IF({Seg_Raw}<>"","segment start and end timestamps and a rate in mg/hour","the date fields required by the selected precision")))') + ')') 40)
)
$EvtCols = @(
    (CIn 'Case_ID' 'Case_ID' 10 $null '@' 'INPUTS: one row per event episode'),
    (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
    (CIn 'Event_ID' 'Event_ID' 8 $null '@'),
    (CIn 'Episode_ID' 'Episode_ID (timepoints: ID/ONSET, ID/DETECTION)' 11 $null '@'),
    (CIn 'Event_Term' 'Event term' 28 $null '@'),
    (CIn 'Baseline_Status' 'Baseline status' 24 $null '@'),
    (CIn 'Missing_Reason' 'Missing-data reason (example codes)' 12 'listwarn:nLIST_MissingReason'),
    (CIn 'Notes' 'Notes' 22 $null '@')
) + @(TPBlock 'O' 'Onset' $true) + @(TPBlock 'D' 'First detection') + @(
    (CF 'Key' 'Event key' '=IF({Episode_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Episode_ID})' 14 $null 'CHECKS'),
    (CF 'Dup' 'Key count' '=IF({Key}="",0,COUNTIF(nEVT_Key,{Key}))' 5 '0'),
    (CF 'Record_Check' 'Record check (first E-code)' ('=IF({Episode_ID}="",IF(COUNTA({Event_Term},{O_Prec},{O_Date})>0,"E_MISSING_INPUT",""),IF({Dup}>1,"E_DUPLICATE_ID",IF({O_Err}<>"",' + (FIRSTCODE '{O_Err}') + ',{D_Err})))') 14),
    (CF 'Check_Message' 'Check message' (RecCheckMsg 'Episode_ID' 'an Episode_ID') 40)
)
$MsCols = @(
    (CIn 'Case_ID' 'Case_ID' 10 $null '@' 'INPUTS: one row per test, organ-function value, intervention, recovery or other milestone'),
    (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
    (CIn 'Milestone_ID' 'Milestone_ID (record ID)' 11 $null '@'),
    (CIn 'Type' 'Type (RECOVERY ends C7; INTERVENTION raises W_INTERVENTION_BEFORE_RECOVERY)' 14 'list:nLIST_MilestoneType'),
    (CIn 'Related_Episode_ID' 'Related episode (optional)' 10 $null '@'),
    (CIn 'Description' 'Description' 30 $null '@'),
    (CIn 'Value' 'Value' 8),
    (CIn 'Unit' 'Unit' 9 $null '@'),
    (CIn 'Reference_Range' 'Reference range' 12 $null '@'),
    (CIn 'Missing_Reason' 'Missing-data reason (example codes)' 12 'listwarn:nLIST_MissingReason'),
    (CIn 'Notes' 'Notes' 20 $null '@')
) + @(TPBlock 'M' 'Milestone date/time') + @(
    (CF 'Key' 'Milestone key' '=IF({Milestone_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Milestone_ID})' 14 $null 'CHECKS'),
    (CF 'CaseKey' 'Case key' '=IF({Milestone_ID}="","",{Case_ID}&"|"&{Case_Version})' 10),
    (CF 'Undated' 'Recorded without any date (1/0)' '=IF(AND({Milestone_ID}<>"",COUNTA({M_Prec}:{M_Upper})=0),1,0)' 5 '0'),
    (CF 'Dup' 'Key count' '=IF({Key}="",0,COUNTIF(nMS_Key,{Key}))' 5 '0'),
    (CF 'Record_Check' 'Record check (E-code)' '=IF({Milestone_ID}="",IF(COUNTA({Description},{M_Prec},{M_Date})>0,"E_MISSING_INPUT",""),IF({Dup}>1,"E_DUPLICATE_ID",IF({Type}="","E_MISSING_INPUT",{M_Err})))' 14),
    (CF 'Check_Message' 'Check message' ('=IF({Record_Check}="","",{Record_Check}&": "&' + (MSGX '{Record_Check}' 'IF({Milestone_ID}="","a Milestone_ID",IF({Type}="","a milestone Type","the date fields required by the selected precision"))') + ')') 40)
)
$LnkCols = @(
    (CIn 'Case_ID' 'Case_ID' 10 $null '@' 'INPUTS: one row per drug-event pair the reviewer selects (no automatic pairing)'),
    (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
    (CIn 'Link_ID' 'Link_ID' 8 $null '@'),
    (CIn 'Exposure_Episode_ID' 'Exposure_Episode_ID' 12 $null '@'),
    (CIn 'Event_Episode_ID' 'Event Episode_ID' 12 $null '@'),
    (CIn 'Reason' 'Reason for the pairing (required)' 40 $null '@'),
    (CIn 'Selected_By' 'Selected by' 12 $null '@'),
    (CF 'Key' 'Link key' '=IF({Link_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Link_ID})' 14 $null 'CHECKS'),
    (CF 'Dup' 'Key count' '=IF({Key}="",0,COUNTIF(nLNK_Key,{Key}))' 5 '0'),
    (CF 'Exp_Found' 'Exposure episode exists' '=IF({Link_ID}="",1,IF(COUNTIF(nEXP_Key,{Case_ID}&"|"&{Case_Version}&"|"&{Exposure_Episode_ID})>0,1,0))' 6 '0'),
    (CF 'Evt_Found' 'Event episode exists' '=IF({Link_ID}="",1,IF(COUNTIF(nEVT_Key,{Case_ID}&"|"&{Case_Version}&"|"&{Event_Episode_ID})>0,1,0))' 6 '0'),
    (CF 'Record_Check' 'Record check (E-code)' '=IF({Link_ID}="",IF(COUNTA({Exposure_Episode_ID},{Event_Episode_ID})>0,"E_MISSING_INPUT",""),IF({Dup}>1,"E_DUPLICATE_ID",IF(OR({Exp_Found}=0,{Evt_Found}=0,{Reason}=""),"E_MISSING_INPUT","")))' 14),
    (CF 'Check_Message' 'Check message' ('=IF({Record_Check}="","",{Record_Check}&": "&' + (MSGX '{Record_Check}' 'IF({Link_ID}="","a Link_ID",IF({Exp_Found}=0,"an Exposures row for "&{Exposure_Episode_ID},IF({Evt_Found}=0,"an Events row for "&{Event_Episode_ID},"a reason for the pairing")))') + ')') 40)
)

# ================================================================== TECHNICAL TABLES (Q1-Q11)
# CALC-SPEC 1.1 Section 4 and rule G9, as the calculator's evaluateTechnical: every E-code is collected and listed in
# Section 5 order; the common gate (Q1-Q5, Q7, Q8) needs four confirmations and every detail and Parameter_Status.
function QHead() {
    @((CIn 'Calc_ID' 'Calc_ID (Result_ID)' 14 $null '@' 'INPUTS'), (CIn 'Case_ID' 'Case_ID' 10 $null '@'), (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
      (CIn 'Question' 'Question this calculation answers (Label)' 30 $null '@'), (CIn 'Stated_Assumptions' 'Stated assumptions (optional; copied to Assumptions)' 22 $null '@'))
}
function QGate([string]$what) {
    @((CIn 'G_Model' 'Gate 1: the simple model applies (one compartment, linear first-order elimination, constant parameters over the period)? Yes/No' 11 'list:nLIST_YesNo' $null 'COMMON GATE (CALC-SPEC 1.1 Section 4: four confirmations, each with its details)'),
      (CIn 'G_HL' "Gate 2: the $what used is confirmed? Yes/No" 10 'list:nLIST_YesNo'),
      (CIn 'G_HL_Type' "Gate 2: type of the $what used (for example terminal elimination)" 16 $null '@'),
      (CIn 'G_HL_Pop' "Gate 2: population of the $what used" 18 $null '@'),
      (CIn 'G_HL_Src' "Gate 2: source of the $what used (document, version, section)" 20 $null '@'),
      (CIn 'G_Ref' 'Gate 3: the reference time is confirmed? Yes/No' 10 'list:nLIST_YesNo'),
      (CIn 'G_RefText' 'Gate 3: reference time, in words (what the result is a fraction of, and when)' 22 $null '@'),
      (CIn 'G_Status' 'Gate 4: each input''s Parameter_Status is recorded? Yes/No' 10 'list:nLIST_YesNo'))
}
function Par($n, $h, $w = 9) { @((CIn $n $h $w 'num'), (CIn "${n}_Status" "$n status (Parameter_Status)" 13 'list:nLIST_Status'), (CIn "${n}_Source" "$n source (document, version, section)" 18 $null '@')) }
function PSET($n, $u) { '"' + $n + ' = "&{' + $n + '}&" "&' + $u + '&" ["&IF({' + $n + '_Status}="","status not recorded",{' + $n + '_Status})&IF({' + $n + '_Source}<>"","; "&{' + $n + '_Source},"")&"]"' }
# check helpers: each check is @(condition, E-code, [input] text for E_MISSING_INPUT)
function NumChk([string]$n, [string]$label, [bool]$pos = $true) {
    $a = @(@(('{' + $n + '}=""'), 'E_MISSING_INPUT', $label), @(('AND({' + $n + '}<>"",NOT(ISNUMBER({' + $n + '})))'), 'E_MISSING_INPUT', ($label + ' as a number')))
    if ($pos) { $a += , @(('AND(ISNUMBER({' + $n + '}),{' + $n + '}<=0)'), 'E_NONPOSITIVE_PARAMETER', '') }
    return $a
}
function OptChk([string]$n, [string]$label) { @(@(('AND({' + $n + '}<>"",NOT(ISNUMBER({' + $n + '})))'), 'E_MISSING_INPUT', ($label + ' as a number')), @(('AND(ISNUMBER({' + $n + '}),{' + $n + '}<=0)'), 'E_NONPOSITIVE_PARAMETER', '')) }
function TimeChk() { @(@('{Time_Unit}=""', 'E_MISSING_INPUT', 'a time unit'), @('AND({Time_Unit}<>"",ISNA(MATCH({Time_Unit},nLIST_TimeUnit,0)))', 'E_INVALID_UNIT', '')) }
function TNonNeg([string]$n = 't') { , @(('AND(ISNUMBER({' + $n + '}),{' + $n + '}<0)'), 'E_OUT_OF_RANGE', '') }
function CalcIdChk() { , @('{Calc_ID}=""', 'E_MISSING_INPUT', 'a Calc_ID') }
# $params: @(@(name, label, optional))
function GateOKF($params) {
    $c = @('{G_Model}="Yes"', '{G_HL}="Yes"', '{G_Ref}="Yes"', '{G_Status}="Yes"', 'TRIM({G_HL_Type}&"")<>""', 'TRIM({G_HL_Pop}&"")<>""', 'TRIM({G_HL_Src}&"")<>""', 'TRIM({G_RefText}&"")<>""')
    foreach ($p in $params) { if ($p[2]) { $c += ('OR({' + $p[0] + '}="",{' + $p[0] + '_Status}<>"")') } else { $c += ('{' + $p[0] + '_Status}<>""') } }
    '=IF(AND(' + ($c -join ',') + '),1,0)'
}
function GateChk([string]$what, $params) {
    $a = @(@('TRIM({G_HL_Type}&"")=""', 'E_MISSING_INPUT', "the type of the $what used"), @('TRIM({G_HL_Pop}&"")=""', 'E_MISSING_INPUT', "the population of the $what used"),
           @('TRIM({G_HL_Src}&"")=""', 'E_MISSING_INPUT', "the source of the $what used"), @('TRIM({G_RefText}&"")=""', 'E_MISSING_INPUT', 'the reference time'))
    foreach ($p in $params) {
        if ($p[2]) { $a += , @(('AND({' + $p[0] + '}<>"",{' + $p[0] + '_Status}="")'), 'E_MISSING_INPUT', ('the Parameter_Status of ' + $p[1].ToLower())) }
        else { $a += , @(('{' + $p[0] + '_Status}=""'), 'E_MISSING_INPUT', ('the Parameter_Status of ' + $p[1].ToLower())) }
    }
    $a += , @('{GateOK}=0', 'E_APPLICABILITY_UNCONFIRMED', '')
    return $a
}
function GateAssume([string]$what) {
    'IF({GateOK}=1," | Gate confirmed: the simple model applies (one compartment, linear first-order elimination, constant parameters over the period). | Gate confirmed: the ' + $what + ' used ' + $U_NDASH + ' type: "&{G_HL_Type}&"; population: "&{G_HL_Pop}&"; source: "&{G_HL_Src}&". | Gate confirmed: reference time ' + $U_NDASH + ' "&{G_RefText}&". | Gate confirmed: each input' + $U_RSQ + 's Parameter_Status is recorded."," | Common gate not confirmed"&IF(AND({G_Model}="Yes",{G_HL}="Yes",{G_Ref}="Yes",{G_Status}="Yes"),": a gate detail or a Parameter_Status is blank.",": "&MID(IF({G_Model}<>"Yes",", model","")&IF({G_HL}<>"Yes",", ' + $what + ' used","")&IF({G_Ref}<>"Yes",", reference time","")&IF({G_Status}<>"Yes",", Parameter_Status",""),3,99)&" not confirmed."))'
}

# Builds the column list of one technical table.
#  $Items: @(@(nameFormula, valueToken, displayFormula, unitFormula, conditionFormula-or-'')) as the calculator's Results
#  $Flags: @(@(condition, code)) W- and I-codes raised with a result, in Section 5 order
function QTable {
    param([string]$Rule, [string]$Title, $Inputs, [string]$ResultDisplay, $Results, $Items, $Checks, $Flags, [string]$Mand, [string]$Words, [string]$PSetF,
          [string]$EmptyF, [string]$What = '', $GateParams = $null, [string]$AssumePre = '', [string]$AssumePost = '', [string]$NoteF = '', $ExtraWork = @())
    $gate = [bool]$What
    $cols = @(QHead)
    if ($gate) { $cols += @(QGate $What) }
    $cols += $Inputs
    $cols += , (CF 'Result_Display' 'Result (display per CALC-SPEC 1.1 Section 7; the E-codes and no number when blocked)' $ResultDisplay 36 $null 'RESULT')
    $dv = @(); $rt = @()
    foreach ($it in $Items) {
        $cond = $(if ($it[4]) { $it[4] } else { 'TRUE' })
        $dv += ('IF(' + $cond + ',"; "&' + $it[0] + '&": "&' + $it[2] + ',"")')
        $rt += ('IF(' + $cond + ',"; "&' + $it[0] + '&" = "&' + $it[1] + '&" ("&' + $it[2] + '&")","")')
    }
    $cols += , (CF 'Displayed_Value' 'Displayed_Value (as the calculator; Not calculable when blocked)' ('=IF({RowEmpty}=1,"",IF({Ok}=0,"Not calculable",MID(' + ($dv -join '&') + ',3,9999)))') 30)
    $cols += $Results
    $fl = '""'
    if ($Flags -and $Flags.Count -gt 0) { $fl = 'MID(' + (($Flags | ForEach-Object { 'IF(' + $_[0] + ',"; ' + $_[1] + '","")' }) -join '&') + ',3,999)' }
    $wm = ''
    if ($Flags -and $Flags.Count -gt 0) { $wm = '&IF({Ok}=1,' + (($Flags | ForEach-Object { WMSG $_[0] $_[1] }) -join '&') + ',"")' }
    $ga = ''; if ($gate) { $ga = '&' + (GateAssume $What) }
    $pre = $(if ($AssumePre) { '&' + $AssumePre } else { '' }); $post = $(if ($AssumePost) { '&' + $AssumePost } else { '' })
    $note = $(if ($NoteF) { '&IF({Ok}=1,' + $NoteF + ',"")' } else { '' })
    $cols += @(
        (CF 'QC_Flags' 'QC_Flags (E-, W-, I-codes; rule G9 order)' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err},' + $fl + '))') 22),
        (CF 'Messages' 'Flag_Messages (verbatim)' ('=IF({RowEmpty}=1,"",MID(' + (EMSGS '{Err}' '{Needs}') + $wm + ',4,9999))') 40),
        (CF 'Parameter_Set' 'Parameter_Set (value, unit, status, source)' ('=IF({RowEmpty}=1,"",' + $PSetF + ')') 40),
        (CF 'Assumptions' 'Assumptions (separated by " | ")' ('=IF({RowEmpty}=1,"",MID(""' + $pre + $ga + $post + '&IF(TRIM({Stated_Assumptions}&"")<>""," | "&{Stated_Assumptions},"")' + $note + ',4,9999))') 40),
        (CF 'Results_Text' 'Results (name = unrounded value (display))' ('=IF({Ok}=0,"",MID(' + ($rt -join '&') + ',3,9999))') 36),
        (CF 'Unit_Out' 'Unit (first result)' ('=IF({Ok}=0,"",' + $Items[0][3] + ')') 12),
        (CF 'Label_Out' 'Label' ('=IF({RowEmpty}=1,"",IF(TRIM({Question}&"")<>"",{Question},"' + $Rule + ' ' + $Title + '"))') 20),
        (CF 'Mandatory_Text' 'Limitations: mandatory text (CALC-SPEC 1.1 Section 4)' ('=IF({RowEmpty}=1,"","' + $Mand + '")') 40),
        (CF 'Formula_in_words' 'Formula in words' ('=IF({RowEmpty}=1,"","' + $Words + '")') 30),
        (CF 'Calculation_Version' 'Calculation_Version' ('=IF({RowEmpty}=1,"","' + $SpecVersion + '")') 11)
    )
    $raw = ($Checks | ForEach-Object { 'IF(' + $_[0] + ',"; ' + $_[1] + '","")' }) -join '&'
    $nd = ($Checks | Where-Object { $_[1] -eq 'E_MISSING_INPUT' -and $_[2] } | ForEach-Object { 'IF(' + $_[0] + ',", ' + $_[2] + '","")' }) -join '&'
    if (-not $nd) { $nd = '""' }
    $cols += @(
        (CF 'RowEmpty' 'work: row empty' $EmptyF 5 '0' 'WORKING'),
        (CF 'CaseKey' 'work: case key' '=IF({RowEmpty}=1,"",{Case_ID}&"|"&{Case_Version})' 12),
        (CF 'Key' 'work: key' '=IF({Calc_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Calc_ID})' 14)
    )
    if ($gate) { $cols += , (CF 'GateOK' 'work: common gate confirmed (1/0)' (GateOKF $GateParams) 5 '0') }
    $cols += $ExtraWork
    $cols += @(
        (CF 'ErrRaw' 'work: E-codes found (any order)' ('=IF({RowEmpty}=1,"",' + $raw + ')') 14),
        (CF 'NeedsRaw' 'work: [input] texts' ('=IF({RowEmpty}=1,"",' + $nd + ')') 14),
        (CF 'Err' 'work: blocking E-codes (G9 order)' ('=IF({RowEmpty}=1,"",' + (ORDE '{ErrRaw}') + ')') 14),
        (CF 'Needs' 'work: [input]' '=MID({NeedsRaw},3,999)' 14),
        (CF 'Ok' 'work: result shown (1/0)' '=IF(AND({RowEmpty}=0,{Err}=""),1,0)' 5 '0')
    )
    return $cols
}
function TD([string]$x) { 'IF({Time_Unit}="minutes",' + (MINUSF ('FIXED(' + $x + ',0,FALSE)')) + ',FIXED(' + $x + ',1,FALSE))' }
function PCT([string]$x) { (P3 ('100*' + $x)) + '&"%"' }

# ---- Q1
$Q1P = @(@('t_half', 't_half', $false), @('t', 't', $false))
$Q1Cols = QTable -Rule 'Q1' -Title 'Elimination and fractional persistence' -What 'half-life' -GateParams $Q1P `
    -Inputs (@((CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS (one time unit)')) + @(Par 't_half' 't_half (in Time_Unit)') + @(Par 't' 't, time since reference (in Time_Unit)')) `
    -ResultDisplay ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown",' + (P3 '{Percent_Remaining}') + '&"% of the reference remains after "&{t}&" "&{Time_Unit}&"; k = "&' + (SIG3 '{k_per_unit}') + '&" per "&LEFT({Time_Unit},LEN({Time_Unit})-1)))') `
    -Results @((CF 'Fraction_Remaining' 'Fraction remaining (unrounded)' '=IF({Ok}=1,2^(-{t}/{t_half}),"")' 12),
               (CF 'Percent_Remaining' 'Percent remaining (unrounded)' '=IF({Ok}=1,100*{Fraction_Remaining},"")' 12),
               (CF 'k_per_unit' 'k (per Time_Unit, unrounded)' '=IF({Ok}=1,LN(2)/{t_half},"")' 12),
               (CF 'k_per_hour' 'k (per hour, unrounded)' '=IF({Ok}=1,LN(2)/({t_half}*IF({Time_Unit}="minutes",1/60,IF({Time_Unit}="hours",1,24))),"")' 12)) `
    -Items @(@('"Fraction remaining"', '{Fraction_Remaining}', (PCT '{Fraction_Remaining}'), '"fraction (shown as percent)"', ''),
             @('"Elimination rate constant k"', '{k_per_unit}', ((SIG3 '{k_per_unit}') + '&" per "&LEFT({Time_Unit},LEN({Time_Unit})-1)'), '"per "&LEFT({Time_Unit},LEN({Time_Unit})-1)', '')) `
    -Checks (@(CalcIdChk) + @(TimeChk) + @(NumChk 't_half' 'the half-life') + @(NumChk 't' 'the time' $false) + @(TNonNeg) + @(GateChk 'half-life' $Q1P)) `
    -Flags @() -Mand 'Fraction of the chosen reference amount or concentration under the stated model. It is not a causal or pharmacological threshold.' `
    -Words 'k = ln 2 / t_half; fraction remaining = 2^(-t / t_half) = exp(-k x t)' -PSetF ((PSET 't_half' '{Time_Unit}') + '&"; "&' + (PSET 't' '{Time_Unit}')) `
    -EmptyF '=IF(AND({Calc_ID}="",{t_half}="",{t}=""),1,0)'

# ---- Q2
$Q2P = @(@('t_half', 't_half', $false), @('f_target', 'the target fraction', $false))
$Q2Cols = QTable -Rule 'Q2' -Title 'Time to approach steady state' -What 'half-life' -GateParams $Q2P `
    -Inputs (@((CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 't_half' 't_half (in Time_Unit)') + @(Par 'f_target' 'Target fraction f of steady state (0 < f < 1)')) `
    -ResultDisplay ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown",' + (TD '{Time_to_f}') + '&" "&{Time_Unit}&" to reach "&' + (P3 '100*{f_target}') + '&"% of steady state ("&FIXED({Half_Lives},2,TRUE)&" half-lives)"))') `
    -Results @((CF 'Half_Lives' 'Half-lives needed n (unrounded)' '=IF({Ok}=1,-LOG(1-{f_target},2),"")' 12),
               (CF 'Time_to_f' 'Time to reach f (in Time_Unit, unrounded)' '=IF({Ok}=1,{Half_Lives}*{t_half},"")' 12)) `
    -Items @(@(('"Time to reach "&' + (PCT '{f_target}')), '{Time_to_f}', ((TD '{Time_to_f}') + '&" "&{Time_Unit}'), '{Time_Unit}', ''),
             @('"Half-lives needed n"', '{Half_Lives}', 'FIXED({Half_Lives},2,TRUE)&" half-lives"', '"half-lives"', '')) `
    -Checks (@(CalcIdChk) + @(TimeChk) + @(NumChk 't_half' 'the half-life') + @(NumChk 'f_target' 'the target fraction' $false) + @(, @('AND(ISNUMBER({f_target}),OR({f_target}<=0,{f_target}>=1))', 'E_OUT_OF_RANGE', '')) + @(GateChk 'half-life' $Q2P)) `
    -Flags @() -Mand 'Assumes constant input from zero. Intermittent dosing compares corresponding points in the dosing cycle. Toxicity need not wait for steady state.' `
    -Words 'n = -log2(1 - f); time = n x t_half' -PSetF ((PSET 't_half' '{Time_Unit}') + '&"; "&' + (PSET 'f_target' '"(fraction)"')) `
    -EmptyF '=IF(AND({Calc_ID}="",{t_half}="",{f_target}=""),1,0)'

# ---- Q3
$Q3P = @(@('t_half', 't_half', $false), @('tau', 'the dosing interval', $false))
$Q3Cols = QTable -Rule 'Q3' -Title 'Accumulation during regular dosing' -What 'half-life' -GateParams $Q3P `
    -Inputs (@((CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 't_half' 't_half (in Time_Unit)') + @(Par 'tau' 'Dosing interval tau (in Time_Unit)')) `
    -ResultDisplay '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","Accumulation factor R = "&FIXED({Accumulation_Factor},2,TRUE)&" at corresponding points in the dosing cycle"))' `
    -Results @((CF 'Accumulation_Factor' 'Accumulation factor R (unrounded)' '=IF({Ok}=1,1/(1-2^(-{tau}/{t_half})),"")' 12)) `
    -Items @(, @('"Accumulation factor R"', '{Accumulation_Factor}', 'FIXED({Accumulation_Factor},2,TRUE)', '"ratio"', '')) `
    -Checks (@(CalcIdChk) + @(TimeChk) + @(NumChk 't_half' 'the half-life') + @(NumChk 'tau' 'the dosing interval') + @(GateChk 'half-life' $Q3P)) `
    -Flags @() -Mand 'Model-based repeated-dose comparison. It is not a toxicity multiplier and does not give the time of clinical onset.' `
    -Words 'R = 1 / (1 - 2^(-tau / t_half))' -PSetF ((PSET 't_half' '{Time_Unit}') + '&"; "&' + (PSET 'tau' '{Time_Unit}')) `
    -EmptyF '=IF(AND({Calc_ID}="",{t_half}="",{tau}=""),1,0)'

# ---- Q4
$CLH = 'IF({CL_Unit}="mL/minute",0.06,1)'
$Q4P = @(@('CL_before', 'the clearance before', $false), @('CL_after', 'the clearance after', $false), @('V_L', 'the apparent volume', $true))
$Q4Cols = QTable -Rule 'Q4' -Title 'A change in clearance' -What 'parameters' -GateParams $Q4P `
    -Inputs (@((CIn 'CL_Unit' 'Clearance unit (L/hour or mL/minute; same for both)' 10 'list:nLIST_CLUnit' $null 'INPUTS')) + @(Par 'CL_before' 'Clearance before') + @(Par 'CL_after' 'Clearance after') + @(Par 'V_L' 'Apparent volume V (L), optional')) `
    -ResultDisplay '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","Half-life ratio "&FIXED({HalfLife_Ratio},2,TRUE)&"; average steady-state concentration ratio "&FIXED({Css_Ratio},2,TRUE)&IF({t_half_before_h}<>""," (t_half "&FIXED({t_half_before_h},1,TRUE)&" hours before, "&FIXED({t_half_after_h},1,TRUE)&" hours after)","")))' `
    -Results @((CF 'HalfLife_Ratio' 'Half-life ratio after/before (unrounded)' '=IF({Ok}=1,{CL_before}/{CL_after},"")' 12),
               (CF 'Css_Ratio' 'Average Css ratio after/before (unrounded)' '=IF({Ok}=1,{CL_before}/{CL_after},"")' 12),
               (CF 't_half_before_h' 't_half before, hours (needs V)' ('=IF(OR({Ok}=0,{V_L}=""),"",LN(2)*{V_L}/({CL_before}*' + $CLH + '))') 12),
               (CF 't_half_after_h' 't_half after, hours (needs V)' ('=IF(OR({Ok}=0,{V_L}=""),"",LN(2)*{V_L}/({CL_after}*' + $CLH + '))') 12)) `
    -Items @(@('"Half-life ratio (after / before)"', '{HalfLife_Ratio}', 'FIXED({HalfLife_Ratio},2,TRUE)', '"ratio"', ''),
             @('"Average steady-state concentration ratio"', '{Css_Ratio}', 'FIXED({Css_Ratio},2,TRUE)', '"ratio"', ''),
             @('"Half-life before"', '{t_half_before_h}', 'FIXED({t_half_before_h},1,FALSE)&" hours"', '"hours"', '{t_half_before_h}<>""'),
             @('"Half-life after"', '{t_half_after_h}', 'FIXED({t_half_after_h},1,FALSE)&" hours"', '"hours"', '{t_half_after_h}<>""')) `
    -Checks (@(CalcIdChk) + @(@('{CL_Unit}=""', 'E_MISSING_INPUT', 'the clearance unit'), @('AND({CL_Unit}<>"",ISNA(MATCH({CL_Unit},nLIST_CLUnit,0)))', 'E_INVALID_UNIT', '')) + @(NumChk 'CL_before' 'the clearance before') + @(NumChk 'CL_after' 'the clearance after') + @(OptChk 'V_L' 'the volume') + @(GateChk 'parameters' $Q4P)) `
    -Flags @() -Mand ('Describes the eventual average steady state. It does not date the patient' + $U_RSQ + 's onset; the transition depends on prior concentrations and when clearance changed.') `
    -Words 'ratio = CL_before / CL_after (V constant); with V: t_half = ln 2 x V / CL for each period' -PSetF ((PSET 'CL_before' '{CL_Unit}') + '&"; "&' + (PSET 'CL_after' '{CL_Unit}') + '&IF({V_L}<>"","; "&' + (PSET 'V_L' '"L"') + ',"")') `
    -EmptyF '=IF(AND({Calc_ID}="",{CL_before}="",{CL_after}=""),1,0)'

# ---- Q5
$Q5P = @(@('Parent_t_half', 'the parent half-life', $false), @('Metabolite_t_half', 'the metabolite half-life', $false), @('t', 't', $false))
$Q5Cols = QTable -Rule 'Q5' -Title 'Parent and metabolite persistence' -What 'half-life' -GateParams $Q5P `
    -Inputs (@((CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 'Parent_t_half' 'Parent t_half') + @(Par 'Metabolite_t_half' 'Metabolite t_half') + @(Par 't' 't, time since reference')) `
    -ResultDisplay ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","Parent "&' + (P3 '{Parent_Percent}') + '&"% and metabolite "&' + (P3 '{Metabolite_Percent}') + '&"% remain, each of its own starting value"))') `
    -Results @((CF 'Parent_Percent' 'Parent percent remaining (unrounded)' '=IF({Ok}=1,100*2^(-{t}/{Parent_t_half}),"")' 12),
               (CF 'Metabolite_Percent' 'Metabolite percent remaining (unrounded)' '=IF({Ok}=1,100*2^(-{t}/{Metabolite_t_half}),"")' 12)) `
    -Items @(@('"Parent remaining (own reference)"', '{Parent_Percent}', ((P3 '{Parent_Percent}') + '&"%"'), '"percent"', ''),
             @('"Metabolite remaining (own reference)"', '{Metabolite_Percent}', ((P3 '{Metabolite_Percent}') + '&"%"'), '"percent"', '')) `
    -Checks (@(CalcIdChk) + @(TimeChk) + @(NumChk 'Parent_t_half' 'the parent half-life') + @(NumChk 'Metabolite_t_half' 'the metabolite half-life') + @(NumChk 't' 'the time' $false) + @(TNonNeg) + @(GateChk 'half-life' $Q5P)) `
    -Flags @() -Mand 'Assumes no further formation or input. Continued metabolite formation, potency, distribution and nonlinear kinetics can make this calculation invalid.' `
    -Words 'Two separate Q1 calculations, each normalized to its own starting value; their sum is never calculated' -PSetF ((PSET 'Parent_t_half' '{Time_Unit}') + '&"; "&' + (PSET 'Metabolite_t_half' '{Time_Unit}') + '&"; "&' + (PSET 't' '{Time_Unit}')) `
    -EmptyF '=IF(AND({Calc_ID}="",{Parent_t_half}="",{t}=""),1,0)'

# ---- Q7
$Q7P = @(@('t_half', 't_half', $false), @('t', 't', $false))
$Q7Cols = QTable -Rule 'Q7' -Title 'Long persistence of a monoclonal antibody' -What 'half-life' -GateParams $Q7P `
    -Inputs (@((CIn 'Time_Unit' 'Time unit (days only for Q7)' 9 'list:nLIST_DaysOnly' $null 'INPUTS')) + @(Par 't_half' 't_half (days)') + @(Par 't' 't (days since reference)')) `
    -ResultDisplay ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown",' + (P3 '{Percent_Remaining}') + '&"% of the post-infusion concentration remains after "&{t}&" days (W_BIOLOGIC_MODEL_CAUTION)"))') `
    -Results @((CF 'Fraction_Remaining' 'Fraction remaining (unrounded)' '=IF({Ok}=1,2^(-{t}/{t_half}),"")' 12),
               (CF 'Percent_Remaining' 'Percent remaining (unrounded)' '=IF({Ok}=1,100*{Fraction_Remaining},"")' 12),
               (CF 'k_per_day' 'k (per day, unrounded)' '=IF({Ok}=1,LN(2)/{t_half},"")' 12)) `
    -Items @(@('"Fraction remaining"', '{Fraction_Remaining}', (PCT '{Fraction_Remaining}'), '"fraction (shown as percent)"', ''),
             @('"Elimination rate constant k"', '{k_per_day}', ((SIG3 '{k_per_day}') + '&" per day"'), '"per day"', '')) `
    -Checks (@(CalcIdChk) + @(@('{Time_Unit}=""', 'E_MISSING_INPUT', 'a time unit'), @('AND({Time_Unit}<>"",{Time_Unit}<>"days")', 'E_INVALID_UNIT', '')) + @(NumChk 't_half' 'the half-life') + @(NumChk 't' 'the time' $false) + @(TNonNeg) + @(GateChk 'half-life' $Q7P)) `
    -Flags @(, @('TRUE', 'W_BIOLOGIC_MODEL_CAUTION')) -Mand 'Antibody disposition can depend on dose, time, target burden and immunogenicity. Exposure persistence is not receptor occupancy, immune activity or biological persistence.' `
    -Words 'As Q1: fraction remaining = 2^(-t / t_half), t and t_half in days' -PSetF ((PSET 't_half' '"days"') + '&"; "&' + (PSET 't' '"days"')) `
    -EmptyF '=IF(AND({Calc_ID}="",{t_half}="",{t}=""),1,0)'

# ---- Q8
$Q8P = @(@('t_half_lower', 'the lower half-life', $false), @('t_half_upper', 'the upper half-life', $false), @('t', 't', $false))
$Q8Cols = QTable -Rule 'Q8' -Title 'Parameter uncertainty and scenarios' -What 'half-life' -GateParams $Q8P `
    -Inputs (@((CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS'), (CIn 'Antibody' 'Molecule recorded as a monoclonal antibody? Yes/No (Yes adds W_BIOLOGIC_MODEL_CAUTION)' 10 'list:nLIST_YesNo')) + @(Par 't_half_lower' 'Lower t_half') + @(Par 't_half_upper' 'Upper t_half') + @(Par 't' 't, time since reference')) `
    -ResultDisplay ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","From "&' + (P3 '{Percent_at_Lower}') + '&"% (t_half "&{t_half_lower}&" "&{Time_Unit}&") to "&' + (P3 '{Percent_at_Upper}') + '&"% (t_half "&{t_half_upper}&" "&{Time_Unit}&") remaining; scenario range, not a confidence interval"))') `
    -Results @((CF 'Fraction_at_Lower' 'Fraction at lower t_half (unrounded)' '=IF({Ok}=1,2^(-{t}/{t_half_lower}),"")' 12),
               (CF 'Fraction_at_Upper' 'Fraction at upper t_half (unrounded)' '=IF({Ok}=1,2^(-{t}/{t_half_upper}),"")' 12),
               (CF 'Percent_at_Lower' 'Percent at lower t_half (unrounded)' '=IF({Ok}=1,100*{Fraction_at_Lower},"")' 12),
               (CF 'Percent_at_Upper' 'Percent at upper t_half (unrounded)' '=IF({Ok}=1,100*{Fraction_at_Upper},"")' 12)) `
    -Items @(@('"Fraction remaining at the lower half-life"', '{Fraction_at_Lower}', (PCT '{Fraction_at_Lower}'), '"fraction (shown as percent)"', ''),
             @('"Fraction remaining at the upper half-life"', '{Fraction_at_Upper}', (PCT '{Fraction_at_Upper}'), '"fraction (shown as percent)"', '')) `
    -Checks (@(CalcIdChk) + @(TimeChk) + @(NumChk 't_half_lower' 'the lower half-life') + @(NumChk 't_half_upper' 'the upper half-life') + @(, @('AND(ISNUMBER({t_half_lower}),ISNUMBER({t_half_upper}),{t_half_lower}>0,{t_half_upper}>0,{t_half_lower}>{t_half_upper})', 'E_REVERSED_BOUNDS', '')) + @(NumChk 't' 'the time' $false) + @(TNonNeg) + @(GateChk 'half-life' $Q8P)) `
    -Flags @(@('{Antibody}="Yes"', 'W_BIOLOGIC_MODEL_CAUTION'), @('TRUE', 'I_SCENARIO_NOT_CI')) -Mand 'Assumption-based scenario range, not a confidence interval.' `
    -Words 'Q1 evaluated at each half-life bound; no midpoint or average is calculated' -PSetF ((PSET 't_half_lower' '{Time_Unit}') + '&"; "&' + (PSET 't_half_upper' '{Time_Unit}') + '&"; "&' + (PSET 't' '{Time_Unit}') + '&"; monoclonal antibody: "&IF({Antibody}="Yes","yes","no")') `
    -AssumePost 'IF({Antibody}="Yes"," | The molecule is a monoclonal antibody, so the Q7 caution applies to each bound.","")' `
    -EmptyF '=IF(AND({Calc_ID}="",{t_half_lower}="",{t}=""),1,0)'

# ---- Q9: amount infused before onset, from Administrations rate segments
$MINE = '((nADM_SegE_@_n+{O_@})-ABS(nADM_SegE_@_n-{O_@}))/2'
function Q9Sum([string]$b) {
    $mi = $MINE.Replace('@', $b)
    'SUMPRODUCT((nADM_SelKey={SelKey})*nADM_SegValid*nADM_Rate_n*(((' + $mi + ')-nADM_SegS_' + $b + '_n)+ABS((' + $mi + ')-nADM_SegS_' + $b + '_n))/2)/60'
}
$Q9SEG = '(nADM_SelKey={SelKey})*(nADM_Kind="INFUSION_SEGMENT")'
$Q9Cols = QTable -Rule 'Q9' -Title 'Dose received before onset during an infusion' `
    -Inputs (@((CIn 'Exposure_Episode_ID' 'Exposure_Episode_ID (rate segments on Administrations)' 12 $null '@' 'INPUTS (no common gate; segment checks)'),
               (CIn 'Onset_Record_ID' 'Onset record ID (a timestamp, e.g. EV-1/ONSET)' 14 $null '@')) + @(Par 'Planned_Dose_mg' 'Planned dose (mg), optional')) `
    -ResultDisplay ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown"&IF(ISNUMBER(SEARCH("E_INVALID_UNIT",{Err}))," (infusion rates must be in mg/hour; mL/hour needs a concentration and is not supported in ' + $SpecVersion + ')",""),FIXED({Amount_mg},1,FALSE)&" mg infused before onset"&IF({Percent_of_Planned}<>"","; "&FIXED({Percent_of_Planned},1,TRUE)&"% of the planned "&{Planned_Dose_mg}&" mg","")&" ("&{Segments_Used}&" segments; "&{Time_Basis}&")"))') `
    -Results @((CF 'Amount_mg' 'Amount infused before onset, mg (unrounded)' '=IF({Ok}=1,IF({AllKnown}=1,{Sum_UTC},{Sum_Loc}),"")' 12),
               (CF 'Percent_of_Planned' 'Percent of planned dose (unrounded)' '=IF(AND({Ok}=1,ISNUMBER({Planned_Dose_mg})),100*{Amount_mg}/{Planned_Dose_mg},"")' 12),
               (CF 'Segments_Used' 'Segments used' '=IF({RowEmpty}=1,"",{N_Seg})' 8),
               (CF 'Time_Basis' 'Time basis' '=IF({Ok}=0,"",IF({AllKnown}=1,"UTC","local clock"))' 10)) `
    -Items @(@('"Amount infused before onset"', '{Amount_mg}', 'FIXED({Amount_mg},1,FALSE)&" mg"', '"mg"', ''),
             @('"Percent of the planned dose"', '{Percent_of_Planned}', 'FIXED({Percent_of_Planned},1,TRUE)&"%"', '"percent"', '{Percent_of_Planned}<>""')) `
    -Checks (@(CalcIdChk) + @(@('{Exposure_Episode_ID}=""', 'E_MISSING_INPUT', 'an Exposure_Episode_ID'), @('{Onset_Record_ID}=""', 'E_MISSING_INPUT', 'the onset time'),
             @('AND({Onset_Record_ID}<>"",{O_Row}=0)', 'E_MISSING_INPUT', 'the onset record "&{Onset_Record_ID}&"'), @('{O_Err}<>""', '"&{O_Err}&"', ''),
             @('AND({O_Row}>0,{O_Err}="",{O_HT}=0)', 'E_MISSING_INPUT', 'the onset time (a timestamp)'),
             @('AND({Exposure_Episode_ID}<>"",{N_Seg}=0)', 'E_MISSING_INPUT', 'at least one infusion segment (INFUSION_SEGMENT rows on Administrations for "&{Exposure_Episode_ID}&")'),
             @('{SegCodes}<>""', '"&MID({SegCodes},3,999)&"', ''),
             @('ISNUMBER(SEARCH("E_MISSING_INPUT",{SegCodes}))', 'E_MISSING_INPUT', 'segment start and end timestamps and a rate in mg/hour for every segment'),
             @('AND({Planned_Dose_mg}<>"",NOT(ISNUMBER({Planned_Dose_mg})))', 'E_MISSING_INPUT', 'the planned dose as a number'),
             @('AND(ISNUMBER({Planned_Dose_mg}),{Planned_Dose_mg}<=0)', 'E_NONPOSITIVE_PARAMETER', ''),
             @('{RevAny}>0', 'E_REVERSED_BOUNDS', ''), @('{OvlAny}>0', 'E_OVERLAPPING_SEGMENTS', ''))) `
    -Flags @(, @('{AllKnown}=0', 'W_TZ_UNKNOWN')) -Mand 'Amount infused, not systemic or effect-site exposure. Assumes the recorded times without interruption and ignores line dead volume.' `
    -Words 'Amount = sum over segments of rate (mg/hour) x hours of that segment before onset' `
    -PSetF '"Rate segments of exposure episode "&{Exposure_Episode_ID}&" (Administrations, "&{N_Seg}&" segments) [case measurement]; onset "&{Onset_Record_ID}&" [case measurement]"&IF({Planned_Dose_mg}<>"","; planned dose = "&{Planned_Dose_mg}&" mg ["&IF({Planned_Dose_mg_Status}="","status not recorded",{Planned_Dose_mg_Status})&IF({Planned_Dose_mg_Source}<>"","; "&{Planned_Dose_mg_Source},"")&"]","")' `
    -EmptyF '=IF(AND({Calc_ID}="",{Exposure_Episode_ID}="",{Onset_Record_ID}=""),1,0)' `
    -ExtraWork @(
        (CF 'SelKey' 'work: episode key' '={Case_ID}&"|"&{Case_Version}&"|"&{Exposure_Episode_ID}' 14),
        (CF 'O_Row' 'work: onset row' '=IF({Onset_Record_ID}="",0,IFERROR(MATCH({Case_ID}&"|"&{Case_Version}&"|"&{Onset_Record_ID},nTP_Key,0),0))' 6 '0'),
        (CF 'O_Err' 'work: onset check' '=IF({O_Row}=0,"",INDEX(nTP_Err,{O_Row})&"")' 10),
        (CF 'O_HT' 'work: onset has time' '=IF({O_Row}=0,0,INDEX(nTP_HT,{O_Row}))' 5 '0'),
        (CF 'O_OK' 'work: onset offset known' '=IF({O_Row}=0,0,INDEX(nTP_OK,{O_Row}))' 5 '0'),
        (CF 'O_Loc' 'work: onset local minutes' '=IF({O_Row}=0,0,INDEX(nTP_Loc_n,{O_Row}))' 10 '0'),
        (CF 'O_UTC' 'work: onset UTC minutes' '=IF({O_Row}=0,0,INDEX(nTP_UTC_n,{O_Row}))' 10 '0'),
        (CF 'N_Seg' 'work: segments in episode' '=COUNTIFS(nADM_SelKey,{SelKey},nADM_Kind,"INFUSION_SEGMENT")' 5 '0'),
        (CF 'AllKnown' 'work: every offset known (UTC basis)' '=IF(AND({O_OK}=1,COUNTIFS(nADM_SelKey,{SelKey},nADM_SegValid,1,nADM_SegOK,0)=0),1,0)' 5 '0'),
        (CF 'SegCodes' 'work: segment E-codes' ('=' + ((@('E_INVALID_DATE', 'E_MISSING_INPUT', 'E_INVALID_UNIT', 'E_NONPOSITIVE_PARAMETER') | ForEach-Object { 'IF(SUMPRODUCT(' + $Q9SEG + '*ISNUMBER(SEARCH("' + $_ + '",nADM_Seg_Raw)))>0,"; ' + $_ + '","")' }) -join '&')) 14),
        (CF 'RevAny' 'work: reversed segments' '=IF({AllKnown}=1,SUMIFS(nADM_Rev_UTC,nADM_SelKey,{SelKey}),SUMIFS(nADM_Rev_Loc,nADM_SelKey,{SelKey}))' 5 '0'),
        (CF 'OvlAny' 'work: overlapping segments' '=IF({AllKnown}=1,SUMIFS(nADM_Ovl_UTC,nADM_SelKey,{SelKey}),SUMIFS(nADM_Ovl_Loc,nADM_SelKey,{SelKey}))' 5 '0'),
        (CF 'Sum_UTC' 'work: amount (UTC basis)' ('=IF({O_Row}=0,0,' + (Q9Sum 'UTC') + ')') 10),
        (CF 'Sum_Loc' 'work: amount (local basis)' ('=IF({O_Row}=0,0,' + (Q9Sum 'Loc') + ')') 10))

# ---- Q10: Method A by default; Method B on request, first and last by administration date (CALC-SPEC 1.1 agreed rule)
$Q10A = '(nADM_SelKey={SelKey})*(nADM_Kind<>"INFUSION_SEGMENT")'
$Q10Cols = QTable -Rule 'Q10' -Title 'Cumulative dose normalized to body surface area' `
    -Inputs @((CIn 'Exposure_Episode_ID' 'Exposure_Episode_ID (doses, BSA and dates on Administrations)' 12 $null '@' 'INPUTS (Method A is the default; BSA is recorded, never calculated)'),
              (CIn 'Method_B' 'Method B on request: FIRST or LAST administration by date, or STATED BSA (blank = not requested)' 12 'list:nLIST_MethodB'),
              (CIn 'Method_B_Stated_BSA' ('Method B stated BSA (m' + $U_SQ2 + ')') 9 'num'),
              (CIn 'Method_B_BSA_Source' 'Method B BSA source' 18 $null '@')) `
    -ResultDisplay ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","Method A (sum of dose / BSA at each administration): "&FIXED({Method_A_mg_m2},1,FALSE)&" mg/m' + $U_SQ2 + '"&IF({N_Missing}>0," (lower bound from "&{N_With_BSA}&" of "&{N_Doses}&" administrations)","")&IF({Method_B}="","","; Method B (total "&{Total_Dose_mg}&" mg / "&{Method_B_BSA_m2}&" m' + $U_SQ2 + ', "&{Method_B}&" BSA): "&FIXED({Method_B_mg_m2},1,FALSE)&" mg/m' + $U_SQ2 + '")))') `
    -Results @((CF 'Method_A_mg_m2' 'Method A, mg/m2 (unrounded; lower bound if W_PARTIAL_BSA)' ('=IF({Ok}=1,SUMPRODUCT(' + $Q10A + '*nADM_Dose_per_BSA),"")') 12),
               (CF 'Total_Dose_mg' 'Total dose, mg' ('=IF({Ok}=1,SUMPRODUCT(' + $Q10A + '*nADM_Dose_pos_n),"")') 10),
               (CF 'N_Doses' 'Administrations' '=IF({RowEmpty}=1,"",{N_Adm})' 8),
               (CF 'N_With_BSA' 'Administrations with a dose and BSA' ('=IF({RowEmpty}=1,"",SUMPRODUCT(' + $Q10A + '*nADM_Known))') 8),
               (CF 'Method_B_BSA_m2' 'Method B BSA used (m2)' '=IF(OR({Ok}=0,{Method_B}=""),"",{B_BSA})' 10),
               (CF 'Method_B_mg_m2' 'Method B, mg/m2 (unrounded)' '=IF(OR({Ok}=0,{Method_B}=""),"",{Total_Dose_mg}/{B_BSA})' 12)) `
    -Items @(@('"Cumulative dose, Method A (sum of dose / BSA at each administration)"', '{Method_A_mg_m2}', ('FIXED({Method_A_mg_m2},1,FALSE)&" mg/m' + $U_SQ2 + '"'), ('"mg/m' + $U_SQ2 + '"'), '{N_Missing}=0'),
             @('"Cumulative dose, Method A, lower bound from the known administrations"', '{Method_A_mg_m2}', ('FIXED({Method_A_mg_m2},1,FALSE)&" mg/m' + $U_SQ2 + ' (lower bound)"'), ('"mg/m' + $U_SQ2 + '"'), '{N_Missing}>0'),
             @(('IF({Method_B}="STATED","Method B (total dose / stated BSA "&{B_BSA}&" m' + $U_SQ2 + ')","Method B (total dose / "&LOWER({Method_B})&" BSA)")'), '{Method_B_mg_m2}', ('FIXED({Method_B_mg_m2},1,FALSE)&" mg/m' + $U_SQ2 + '"'), ('"mg/m' + $U_SQ2 + '"'), '{Method_B}<>""')) `
    -Checks (@(CalcIdChk) + @(@('{Exposure_Episode_ID}=""', 'E_MISSING_INPUT', 'an Exposure_Episode_ID'),
             @('AND({Exposure_Episode_ID}<>"",{N_Adm}=0)', 'E_MISSING_INPUT', 'at least one administration'),
             @('{AdmCodes}<>""', '"&MID({AdmCodes},3,999)&"', ''),
             @('ISNUMBER(SEARCH("E_MISSING_INPUT",{AdmCodes}))', 'E_MISSING_INPUT', 'the dose, dose unit and any BSA of every administration of "&{Exposure_Episode_ID}&" as numbers'),
             @('AND({N_Adm}>0,{N_Known}=0,{AdmCodes}="")', 'E_MISSING_INPUT', 'the BSA recorded at one or more administrations'),
             @('AND({Method_B}<>"",ISNA(MATCH({Method_B},nLIST_MethodB,0)))', 'E_MISSING_INPUT', 'a Method B choice (FIRST, LAST or STATED)'),
             @('AND({FL}=1,{N_Dated}>0,{N_Dated}<{N_Adm})', 'E_MISSING_INPUT', 'the date of every administration (Method B by first or last administration)'),
             @('{DateCodes}<>""', '"&MID({DateCodes},3,999)&"', ''),
             @('{TieFirst}=1', 'E_MISSING_INPUT', 'one BSA for the first administration: administrations share its date with different BSAs'),
             @('{TieLast}=1', 'E_MISSING_INPUT', 'one BSA for the last administration: administrations share its date with different BSAs'),
             @('AND({FL}=1,{N_Adm}>0,NOT(ISNUMBER({B_BSA})))', 'E_MISSING_INPUT', 'the "&LOWER({Method_B})&" BSA for Method B'),
             @('AND({Method_B}="STATED",NOT(ISNUMBER({Method_B_Stated_BSA})))', 'E_MISSING_INPUT', 'the stated BSA for Method B'),
             @('AND({Method_B}="STATED",ISNUMBER({Method_B_Stated_BSA}),{Method_B_Stated_BSA}<=0)', 'E_NONPOSITIVE_PARAMETER', ''))) `
    -Flags @(, @('{N_Missing}>0', 'W_PARTIAL_BSA')) -Mand 'Describes cumulative exposure. It is not a toxicity threshold.' `
    -Words 'Method A (default): sum of dose_i / BSA_i; Method B (on request): total dose / one BSA (first or last administration by date, or stated)' `
    -PSetF ('"Doses (mg), dates and BSA (m' + $U_SQ2 + ') recorded at each administration of "&{Exposure_Episode_ID}&" on Administrations ("&{N_Adm}&" administrations) [case measurement]"&IF({Method_B}<>"","; Method B BSA: "&{Method_B}&IF({Method_B_BSA_Source}<>""," ("&{Method_B_BSA_Source}&")",""),"")') `
    -AssumePre '" | BSA is the value recorded at each administration; the tool never calculates BSA (CALC-SPEC 1.1, Q10)."' `
    -NoteF ('IF({FL}=1,IF({ByDate}=1," | Method B takes the first and last administrations by date ("&' + (DTXT '{DMin}') + '&" and "&' + (DTXT '{DMax}') + '&").",IF({N_Dated}=0," | No administration date is recorded, so Method B takes the first and last administrations in the order entered.","")),"")') `
    -EmptyF '=IF(AND({Calc_ID}="",{Exposure_Episode_ID}=""),1,0)' `
    -ExtraWork @(
        (CF 'SelKey' 'work: episode key' '={Case_ID}&"|"&{Case_Version}&"|"&{Exposure_Episode_ID}' 14),
        (CF 'N_Adm' 'work: administrations' ('=SUMPRODUCT(' + $Q10A + ')') 5 '0'),
        (CF 'N_Known' 'work: with dose and BSA' ('=SUMPRODUCT(' + $Q10A + '*nADM_Known)') 5 '0'),
        (CF 'N_Missing' 'work: BSA blank' ('=SUMPRODUCT(' + $Q10A + '*nADM_BSA_missing)') 5 '0'),
        (CF 'AdmCodes' 'work: administration E-codes' ('=' + ((@('E_MISSING_INPUT', 'E_INVALID_UNIT', 'E_NONPOSITIVE_PARAMETER') | ForEach-Object { 'IF(SUMPRODUCT(' + $Q10A + '*ISNUMBER(SEARCH("' + $_ + '",nADM_Q10_Raw)))>0,"; ' + $_ + '","")' }) -join '&')) 14),
        (CF 'FL' 'work: Method B first or last' '=IF(OR({Method_B}="FIRST",{Method_B}="LAST"),1,0)' 4 '0'),
        (CF 'N_Dated' 'work: administrations with a date' ('=SUMPRODUCT(' + $Q10A + '*nADM_A_Dated)') 5 '0'),
        (CF 'DateCodes' 'work: Method B date E-codes' ('=IF(OR({FL}=0,{N_Adm}=0,{N_Dated}<{N_Adm}),"",' + ((@('E_INVALID_DATE', 'E_REVERSED_BOUNDS', 'E_MISSING_INPUT', 'E_INVALID_UNIT') | ForEach-Object { 'IF(SUMPRODUCT(' + $Q10A + '*(nADM_Q10_DateErr="' + $_ + '"))>0,"; ' + $_ + '","")' }) -join '&') + ')') 12),
        (CF 'ByDate' 'work: Method B by date' ('=IF(AND({FL}=1,{N_Adm}>0,{N_Dated}={N_Adm},SUMPRODUCT(' + $Q10A + '*nADM_Valid)={N_Adm}),1,0)') 4 '0'),
        (CF 'DMin' 'work: earliest administration date' ('=IF({ByDate}=0,0,1E+12-SUMPRODUCT(MAX(' + $Q10A + '*(1E+12-nADM_Lo_n))))') 8 'd mmmm yyyy'),
        (CF 'DMax' 'work: latest administration date' ('=IF({ByDate}=0,0,SUMPRODUCT(MAX(' + $Q10A + '*nADM_Lo_n)))') 8 'd mmmm yyyy'),
        (CF 'FirstRow' 'work: first administration row' ('=IF({N_Adm}=0,0,IF({ByDate}=1,MATCH(1,INDEX(' + $Q10A + '*(nADM_Lo_n={DMin}),0),0),MATCH(1,INDEX(' + $Q10A + ',0),0)))') 5 '0'),
        (CF 'LastRow' 'work: last administration row' ('=IF({N_Adm}=0,0,IF({ByDate}=1,MATCH(1,INDEX(' + $Q10A + '*(nADM_Lo_n={DMax}),0),0),MATCH(SUMPRODUCT(MAX(' + $Q10A + '*nADM_RowN)),nADM_RowN,0)))') 5 '0'),
        (CF 'TieFirst' 'work: first date shared with different BSAs' ('=IF(OR({ByDate}=0,{FirstRow}=0),0,IF(SUMPRODUCT(' + $Q10A + '*(nADM_Lo_n={DMin})*(nADM_BSA_key<>INDEX(nADM_BSA_key,{FirstRow})))>0,1,0))') 4 '0'),
        (CF 'TieLast' 'work: last date shared with different BSAs' ('=IF(OR({ByDate}=0,{LastRow}=0),0,IF(SUMPRODUCT(' + $Q10A + '*(nADM_Lo_n={DMax})*(nADM_BSA_key<>INDEX(nADM_BSA_key,{LastRow})))>0,1,0))') 4 '0'),
        (CF 'B_BSA' 'work: Method B BSA' '=IF({Method_B}="STATED",IF(ISNUMBER({Method_B_Stated_BSA}),{Method_B_Stated_BSA},""),IF(AND({FL}=1,{N_Adm}>0),IF({Method_B}="FIRST",IF({FirstRow}=0,"",INDEX(nADM_BSA_m2,{FirstRow})),IF({LastRow}=0,"",INDEX(nADM_BSA_m2,{LastRow}))),""))' 8))

# ---- Q11
$SCRU = ($U_MU + 'mol/L')
$Q11Cols = QTable -Rule 'Q11' -Title ('Estimated renal function (Cockcroft' + $U_NDASH + 'Gault)') `
    -Inputs @((CIn 'Stability_Confirmed' 'Gate: serum creatinine stable at this time? Yes/No' 10 'list:nLIST_YesNo' $null 'GATE (Q11)'),
              (CIn 'Stability_Basis' 'Basis for stability (values and dates)' 26 $null '@'),
              (CIn 'Age_years' 'Age (years, 18-120)' 7 'num' $null 'INPUTS'),
              (CIn 'Age_Status' 'Age status' 13 'list:nLIST_Status'),
              (CIn 'Weight_kg' 'Weight (kg)' 7 'num'),
              (CIn 'Weight_Type' 'Weight type (required)' 10 'list:nLIST_WeightType'),
              (CIn 'Weight_Status' 'Weight status' 13 'list:nLIST_Status'),
              (CIn 'SCr' 'Serum creatinine' 8 'num'),
              (CIn 'SCr_Unit' ('Creatinine unit (mg/dL or ' + $SCRU + ')') 9 'list:nLIST_SCrUnit'),
              (CIn 'SCr_Status' 'Creatinine status' 13 'list:nLIST_Status'),
              (CIn 'SCr_Source' 'Creatinine source (record, date)' 18 $null '@'),
              (CIn 'Sex' 'Sex as recorded' 10 'list:nLIST_Sex')) `
    -ResultDisplay '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown",IF({Sex}="not recorded","Female "&FIXED({CrCl_Female_If_Sex_Not_Recorded},0,TRUE)&" mL/minute; male "&FIXED({CrCl_Male_If_Sex_Not_Recorded},0,TRUE)&" mL/minute (sex not recorded; both shown)",FIXED({CrCl_mL_min},0,TRUE)&" mL/minute estimated creatinine clearance ("&{Sex}&")")))' `
    -Results @((CF 'SCr_mg_dL' 'Creatinine, mg/dL (unrounded)' '=IF({Ok}=1,IF({SCr_Unit}="mg/dL",{SCr},{SCr}/88.4),"")' 10),
               (CF 'CrCl_mL_min' 'CrCl for recorded sex, mL/minute (unrounded)' '=IF(AND({Ok}=1,{Sex}<>"not recorded"),(140-{Age_years})*{Weight_kg}/(72*{SCr_mg_dL})*IF({Sex}="female",0.85,1),"")' 12),
               (CF 'CrCl_Female_If_Sex_Not_Recorded' 'If sex not recorded: female estimate (unrounded)' '=IF(AND({Ok}=1,{Sex}="not recorded"),0.85*(140-{Age_years})*{Weight_kg}/(72*{SCr_mg_dL}),"")' 12),
               (CF 'CrCl_Male_If_Sex_Not_Recorded' 'If sex not recorded: male estimate (unrounded)' '=IF(AND({Ok}=1,{Sex}="not recorded"),(140-{Age_years})*{Weight_kg}/(72*{SCr_mg_dL}),"")' 12)) `
    -Items @(@('"Estimated creatinine clearance"', '{CrCl_mL_min}', 'FIXED({CrCl_mL_min},0,TRUE)&" mL/minute"', '"mL/minute"', '{Sex}<>"not recorded"'),
             @('"Estimated CrCl with the female factor (0.85)"', '{CrCl_Female_If_Sex_Not_Recorded}', 'FIXED({CrCl_Female_If_Sex_Not_Recorded},0,TRUE)&" mL/minute"', '"mL/minute"', '{Sex}="not recorded"'),
             @('"Estimated CrCl without the female factor"', '{CrCl_Male_If_Sex_Not_Recorded}', 'FIXED({CrCl_Male_If_Sex_Not_Recorded},0,TRUE)&" mL/minute"', '"mL/minute"', '{Sex}="not recorded"'),
             @('"Serum creatinine in mg/dL"', '{SCr_mg_dL}', 'FIXED({SCr_mg_dL},2,TRUE)&" mg/dL"', '"mg/dL"', '{SCr_Unit}<>"mg/dL"')) `
    -Checks (@(CalcIdChk) + @(, @('{Stability_Confirmed}<>"Yes"', 'E_CREATININE_STABILITY_UNCONFIRMED', '')) + @(NumChk 'Weight_kg' 'the weight') + @(NumChk 'SCr' 'the serum creatinine') + @(NumChk 'Age_years' 'the age' $false) +
             @(@('AND(ISNUMBER({Age_years}),OR({Age_years}<18,{Age_years}>120))', 'E_OUT_OF_RANGE', ''), @('{SCr_Unit}=""', 'E_MISSING_INPUT', 'the creatinine unit'),
               @(('AND({SCr_Unit}<>"",{SCr_Unit}<>"mg/dL",{SCr_Unit}<>"' + $SCRU + '",{SCr_Unit}<>"umol/L")'), 'E_INVALID_UNIT', ''),
               @('AND({Sex}<>"male",{Sex}<>"female",{Sex}<>"not recorded")', 'E_MISSING_INPUT', 'sex as recorded (male, female or not recorded)'),
               @('TRIM({Weight_Type}&"")=""', 'E_MISSING_INPUT', 'the weight type (for example actual body weight)'))) `
    -Flags @(, @('{Sex}="not recorded"', 'W_SEX_NOT_RECORDED')) -Mand 'Assumes stable creatinine. It is not drug clearance and not a dosing recommendation.' `
    -Words ('Cockcroft' + $U_NDASH + 'Gault: CrCl = (140 - age) x weight / (72 x SCr), x 0.85 when female sex is recorded; ' + $SCRU + ' / 88.4 = mg/dL') `
    -PSetF '"age = "&{Age_years}&" years ["&IF({Age_Status}="","status not recorded",{Age_Status})&"]; weight = "&{Weight_kg}&" kg, "&{Weight_Type}&" ["&IF({Weight_Status}="","status not recorded",{Weight_Status})&"]; serum creatinine = "&{SCr}&" "&{SCr_Unit}&" ["&IF({SCr_Status}="","status not recorded",{SCr_Status})&IF({SCr_Source}<>"","; "&{SCr_Source},"")&"]; sex: "&{Sex}' `
    -AssumePre 'IF({Stability_Confirmed}="Yes"," | Stable creatinine confirmed by the reviewer"&IF(TRIM({Stability_Basis}&"")<>"",": "&{Stability_Basis},".")," | Creatinine stability not confirmed.")' `
    -EmptyF '=IF(AND({Calc_ID}="",{Age_years}="",{SCr}=""),1,0)'

# ---- Q6 (PD_Profile, teaching demonstration): two named uses, one formula (CALC-SPEC 1.1)
$Q6Cols = QTable -Rule 'Q6' -Title 'Biological replacement, recovery or approach to a new level (teaching demonstration)' `
    -Inputs (@((CIn 'Teaching_Switch' 'Teaching-demonstration switch on? Yes/No' 10 'list:nLIST_YesNo' $null 'GATE (Q6)'),
               (CIn 'Use' 'Use (required): replacement or recovery, or approach to a new level' 16 'list:nLIST_Q6Use'),
               (CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 'Replacement_Half_Time' 'Half-time of the turnover (in Time_Unit)') + @(Par 't' 't (in Time_Unit)')) `
    -ResultDisplay ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown"&IF(ISNUMBER(SEARCH("E_APPLICABILITY_UNCONFIRMED",{Err}))," (Q6 is shown only with the teaching-demonstration switch on)",""),' + (P3 '{Percent_Q6}') + '&IF({Use}="approach to a new level","% of the change to the new level completed after ","% replaced after ")&{t}&" "&{Time_Unit}&" (teaching demonstration with hypothetical values)"))') `
    -Results @((CF 'Output_Label' 'Output (by use)' '=IF({Ok}=0,"",IF({Use}="approach to a new level","Fraction of the change to the new level completed","Fraction replaced"))' 22),
               (CF 'Fraction_Q6' 'Fraction replaced, or of the change completed (unrounded)' '=IF({Ok}=1,1-2^(-{t}/{Replacement_Half_Time}),"")' 12),
               (CF 'Percent_Q6' 'Percent (unrounded)' '=IF({Ok}=1,100*{Fraction_Q6},"")' 12)) `
    -Items @(, @('{Output_Label}', '{Fraction_Q6}', (PCT '{Fraction_Q6}'), '"fraction (shown as percent)"', '')) `
    -Checks (@(CalcIdChk) + @(TimeChk) + @(NumChk 'Replacement_Half_Time' 'the replacement half-time') + @(NumChk 't' 'the time' $false) + @(TNonNeg) +
             @(@('{Teaching_Switch}<>"Yes"', 'E_APPLICABILITY_UNCONFIRMED', ''), @('ISNA(MATCH({Use},nLIST_Q6Use,0))', 'E_MISSING_INPUT', 'the use of Q6 (replacement or recovery, or approach to a new level)'))) `
    -Flags @(, @('TRUE', 'I_TEACHING_DEMONSTRATION')) -Mand 'Hypothetical turnover illustration. It is not a measured recovery curve for platelets or any other clinical target.' `
    -Words 'fraction = 1 - 2^(-t / half-time): the fraction replaced, or the fraction of the change to a new level completed' `
    -PSetF ((PSET 'Replacement_Half_Time' '{Time_Unit}') + '&"; "&' + (PSET 't' '{Time_Unit}') + '&"; use: "&IF({Use}="","not recorded",{Use})') `
    -AssumePost 'IF({Teaching_Switch}="Yes"," | Teaching-demonstration switch on."," | Teaching-demonstration switch off.")' `
    -EmptyF '=IF(AND({Calc_ID}="",{Replacement_Half_Time}="",{t}=""),1,0)'

# ---------------- reserved rows for technical tables
$SV = 'hypothetical example'
function QR($tid, $case, $ver, $q, $extra) { MergeH @(@{ Calc_ID = $tid; Case_ID = $case; Case_Version = $ver; Question = $q }, $extra) }
$GV = GY 'terminal elimination (hypothetical)' 'hypothetical population (vector file)' $VecFile 'Reference amount at time zero'
$Q1Vec = @(
    (QR 'Q1-V16' 'VEC-V16' 1 'Vector V16' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half' 24 $SV "$V V16"), (PV 't' 72 $SV "$V V16")))),
    (QR 'Q1-V24a' 'VEC-V24' 1 'Vector V24, case 1: half-life 0' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half' 0 $SV "$V V24"), (PV 't' 72 $SV "$V V24")))),
    (QR 'Q1-V24b' 'VEC-V24' 1 'Vector V24, case 2: unit fortnights' (MergeH @($GV, @{ Time_Unit = 'fortnights' }, (PV 't_half' 24 $SV "$V V24"), (PV 't' 72 $SV "$V V24")))),
    (QR 'Q1-V24c' 'VEC-V24' 1 'Vector V24, case 3: gate unconfirmed' (MergeH @($GV, @{ G_Model = 'No'; Time_Unit = 'hours' }, (PV 't_half' 24 $SV "$V V24"), (PV 't' 72 $SV "$V V24")))),
    (QR 'Q1-V35a' 'VEC-V35' 1 'Vector V35, case 1: half-life 0, unit fortnights, gate unconfirmed (every E-code listed)' (MergeH @($GV, @{ G_Model = 'No'; Time_Unit = 'fortnights' }, (PV 't_half' 0 $SV "$V V35"), (PV 't' 72 $SV "$V V35")))),
    (QR 'Q1-V36a' 'VEC-V36' 1 'Vector V36, case 1: half-life type blank and the Parameter_Status of t missing' (MergeH @((GY '' 'adults' 'Label 12.3' 'end of the last infusion'), @{ Time_Unit = 'hours' }, (PV 't_half' 24 'label fact' "$V V36"), (PV 't' 72 '' "$V V36")))),
    (QR 'Q1-V36b' 'VEC-V36' 1 'Vector V36, case 2: gate complete' (MergeH @((GY 'terminal half-life' 'adults' 'Label 12.3' 'end of the last infusion'), @{ Time_Unit = 'hours' }, (PV 't_half' 24 'label fact' "$V V36"), (PV 't' 72 'case measurement' "$V V36")))),
    (QR 'Q1-V36c' 'VEC-V36' 1 'Vector V36, case 3: model not confirmed' (MergeH @((GY 'terminal half-life' 'adults' 'Label 12.3' 'end of the last infusion'), @{ G_Model = 'No'; Time_Unit = 'hours' }, (PV 't_half' 24 'label fact' "$V V36"), (PV 't' 72 'case measurement' "$V V36"))))
)
$Q2Vec = @((QR 'Q2-V17' 'VEC-V17' 1 'Vector V17' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half' 24 $SV "$V V17"), (PV 'f_target' 0.9 $SV "$V V17")))))
$Q3Vec = @((QR 'Q3-V18' 'VEC-V18' 1 'Vector V18' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half' 24 $SV "$V V18"), (PV 'tau' 12 $SV "$V V18")))))
$Q4Vec = @((QR 'Q4-V19' 'VEC-V19' 1 'Vector V19' (MergeH @((GY 'total clearance (hypothetical)' 'hypothetical population (vector file)' $VecFile 'Eventual average steady state'), @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 5 $SV "$V V19"), (PV 'CL_after' 2.5 $SV "$V V19")))))
$C04Q4G = GY 'relative total clearance (baseline = 1, entered on the L/hour scale; only the ratio is used)' 'fe 0.5-0.7 from healthy volunteers given IV digoxin (not the patient)' 'LANOXIN US 12.3; Chapter 11, Worked exercise 11.2' 'Eventual average steady state after the renal change'
$Q4Ex = @(
    (QR 'Q4-C04-half-fe0.5' 'EX-C04' 2 'Renal elimination halved, fe 0.5: CL_after/CL_before = 0.5 + 0.5 x 0.5 = 0.75 (Chapter 11, Worked exercise 11.2)' (MergeH @($C04Q4G, @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 1 'model assumption' 'Chapter 11, Worked exercise 11.2'), (PV 'CL_after' 0.75 'model assumption' 'Chapter 11, Worked exercise 11.2')))),
    (QR 'Q4-C04-half-fe0.7' 'EX-C04' 2 'Renal elimination halved, fe 0.7: 0.3 + 0.7 x 0.5 = 0.65' (MergeH @($C04Q4G, @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 1 'model assumption' 'Chapter 11, Worked exercise 11.2'), (PV 'CL_after' 0.65 'model assumption' 'Chapter 11, Worked exercise 11.2')))),
    (QR 'Q4-C04-quarter-fe0.5' 'EX-C04' 2 'Renal elimination at a quarter, fe 0.5: 0.625' (MergeH @($C04Q4G, @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 1 'model assumption' 'Chapter 11, Worked exercise 11.2'), (PV 'CL_after' 0.625 'model assumption' 'Chapter 11, Worked exercise 11.2')))),
    (QR 'Q4-C04-quarter-fe0.7' 'EX-C04' 2 'Renal elimination at a quarter, fe 0.7: 0.475' (MergeH @($C04Q4G, @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 1 'model assumption' 'Chapter 11, Worked exercise 11.2'), (PV 'CL_after' 0.475 'model assumption' 'Chapter 11, Worked exercise 11.2'))))
)
$Q5Vec = @((QR 'Q5-V20' 'VEC-V20' 1 'Vector V20' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 'Parent_t_half' 6 $SV "$V V20"), (PV 'Metabolite_t_half' 72 $SV "$V V20"), (PV 't' 48 $SV "$V V20")))))
$Q7Vec = @((QR 'Q7-V22' 'VEC-V22' 1 'Vector V22' (MergeH @($GV, @{ Time_Unit = 'days' }, (PV 't_half' 21 $SV "$V V22"), (PV 't' 42 $SV "$V V22")))))
$C28G = GY 'terminal half-life; NHL median 22 days, range 6.1-52 days (label fact for a population, not this patient)' 'patients with non-Hodgkin lymphoma' 'RITUXAN US 12.3; MabThera EU 5.2' 'Post-infusion concentration after the last infusion (3 March 2025); accumulation over earlier cycles ignored'
$Q7Ex = @(
    (QR 'Q7-C28-435' 'EX-C28' 2 'Exposure persistence at the first positive test, 435 days after the last infusion (median half-life)' (MergeH @($C28G, @{ Time_Unit = 'days' }, (PV 't_half' 22 'label fact' 'RITUXAN US 12.3 (NHL, median)'), (PV 't' 435 'case measurement' 'Intervals I01')))),
    (QR 'Q7-C28-343' 'EX-C28' 2 'At the earliest possible onset, 343 days (median half-life)' (MergeH @($C28G, @{ Time_Unit = 'days' }, (PV 't_half' 22 'label fact' 'RITUXAN US 12.3 (NHL, median)'), (PV 't' 343 'case measurement' 'Intervals I02 (minimum)'))))
)
$Q8Vec = @(
    (QR 'Q8-V23' 'VEC-V23' 1 'Vector V23' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half_lower' 18 $SV "$V V23"), (PV 't_half_upper' 36 $SV "$V V23"), (PV 't' 72 $SV "$V V23")))),
    (QR 'Q8-V37' 'VEC-V37' 1 'Vector V37: a recorded monoclonal antibody' (MergeH @($GV, @{ Time_Unit = 'days'; Antibody = 'Yes' }, (PV 't_half_lower' 21 $SV "$V V37"), (PV 't_half_upper' 42 $SV "$V V37"), (PV 't' 84 $SV "$V V37"))))
)
$C04G1 = GY 'terminal elimination half-life 1.5-2 days' 'healthy volunteers with normal renal function (not the patient''s population)' 'LANOXIN US 12.3' 'Concentration when digoxin was withheld on 5 July 2026; no further doses; t = 3 days to the improvement on 8 July'
$C04G2 = GY 'terminal elimination half-life 3.5-5 days' 'anuric patients (not the patient''s population)' 'LANOXIN US 12.3' 'Concentration when digoxin was withheld on 5 July 2026; no further doses; t = 3 days to the improvement on 8 July'
$C28G8 = GY 'terminal half-life; labeled median and longest value of the range 6.1-52 days' 'non-Hodgkin lymphoma' 'RITUXAN US 12.3; MabThera EU 5.2' 'concentration after the last infusion, 3 March 2025; accumulation over earlier cycles ignored'
$Q8Ex = @(
    (QR 'Q8-C04-normal' 'EX-C04' 2 'Fraction remaining at the documented improvement, normal renal function half-lives (Chapter 23, Section 23.11)' (MergeH @($C04G1, @{ Time_Unit = 'days'; Antibody = 'No' }, (PV 't_half_lower' 1.5 'label fact' 'LANOXIN US 12.3'), (PV 't_half_upper' 2 'label fact' 'LANOXIN US 12.3'), (PV 't' 3 'case measurement' 'Intervals I06 (5 to 8 July)')))),
    (QR 'Q8-C04-anuric' 'EX-C04' 2 'Same, anuric-patient half-lives' (MergeH @($C04G2, @{ Time_Unit = 'days'; Antibody = 'No' }, (PV 't_half_lower' 3.5 'label fact' 'LANOXIN US 12.3'), (PV 't_half_upper' 5 'label fact' 'LANOXIN US 12.3'), (PV 't' 3 'case measurement' 'Intervals I06 (5 to 8 July)')))),
    (QR 'Q8-C28-435' 'EX-C28' 2 'Labeled NHL half-life range at the first positive test (435 days)' (MergeH @($C28G, @{ Time_Unit = 'days'; Antibody = 'Yes' }, (PV 't_half_lower' 6.1 'label fact' 'RITUXAN US 12.3 (NHL range)'), (PV 't_half_upper' 52 'label fact' 'RITUXAN US 12.3 (NHL range)'), (PV 't' 435 'case measurement' 'Intervals I01')))),
    (QR 'Q8-C28-343' 'EX-C28' 2 'Labeled NHL half-life range at the earliest possible onset (343 days)' (MergeH @($C28G, @{ Time_Unit = 'days'; Antibody = 'Yes' }, (PV 't_half_lower' 6.1 'label fact' 'RITUXAN US 12.3 (NHL range)'), (PV 't_half_upper' 52 'label fact' 'RITUXAN US 12.3 (NHL range)'), (PV 't' 343 'case measurement' 'Intervals I02 (minimum)')))),
    (QR 'Q8-C28-435m' 'EX-C28' 2 'Rituximab remaining at the first positive test (435 days), median and longest labeled half-lives (calculator example T1)' (MergeH @($C28G8, @{ Time_Unit = 'days'; Antibody = 'Yes' }, (PV 't_half_lower' 22 'label fact' 'RITUXAN US 12.3 (NHL median)'), (PV 't_half_upper' 52 'label fact' 'RITUXAN US 12.3 (NHL longest)'), (PV 't' 435 'case measurement' 'Intervals I01')))),
    (QR 'Q8-C28-343m' 'EX-C28' 2 'Rituximab remaining at the earliest onset bound (343 days), median and longest labeled half-lives (calculator example T2)' (MergeH @($C28G8, @{ Time_Unit = 'days'; Antibody = 'Yes' }, (PV 't_half_lower' 22 'label fact' 'RITUXAN US 12.3 (NHL median)'), (PV 't_half_upper' 52 'label fact' 'RITUXAN US 12.3 (NHL longest)'), (PV 't' 343 'case measurement' 'Intervals I02 (minimum)'))))
)
$Q9Ex = @((QR 'Q9-C01' 'EX-C01' 2 'Amount of Drug A infused before onset (pump rate recorded in mL/hour only)' @{ Exposure_Episode_ID = 'EX-A'; Onset_Record_ID = 'EV-1/ONSET' }))
$Q9Vec = @((QR 'Q9-V25' 'VEC-V25' 1 'Vector V25' (MergeH @(@{ Exposure_Episode_ID = 'EX1'; Onset_Record_ID = 'EV1/ONSET' }, (PV 'Planned_Dose_mg' 500 $SV "$V V25")))))
$Q10Vec = @(
    (QR 'Q10-V26-first' 'VEC-V26' 1 'Vector V26, Method B with the first BSA (no dates: order entered)' @{ Exposure_Episode_ID = 'EX1'; Method_B = 'FIRST'; Method_B_BSA_Source = 'BSA at the first administration' }),
    (QR 'Q10-V26-last' 'VEC-V26' 1 'Vector V26, Method B with the last BSA (no dates: order entered)' @{ Exposure_Episode_ID = 'EX1'; Method_B = 'LAST'; Method_B_BSA_Source = 'BSA at the last administration' }),
    (QR 'Q10-V38-first' 'VEC-V38' 1 'Vector V38, Method B with the first BSA by date (listed out of date order)' @{ Exposure_Episode_ID = 'EX1'; Method_B = 'FIRST'; Method_B_BSA_Source = 'BSA at the earliest administration' }),
    (QR 'Q10-V38-last' 'VEC-V38' 1 'Vector V38, Method B with the last BSA by date (listed out of date order)' @{ Exposure_Episode_ID = 'EX1'; Method_B = 'LAST'; Method_B_BSA_Source = 'BSA at the latest administration' })
)
function Q11R($id, $case, $ver, $q, $stable, $basis, $age, $wt, $scr, $unit, $sex, $st = $SV, $src = '') {
    QR $id $case $ver $q @{ Stability_Confirmed = $stable; Stability_Basis = $basis; Age_years = $age; Age_Status = $st; Weight_kg = $wt; Weight_Type = 'actual'; Weight_Status = $st; SCr = $scr; SCr_Unit = $unit; SCr_Status = $st; SCr_Source = $src; Sex = $sex }
}
$Q11Ex = @(
    (Q11R 'Q11-C04-baseline' 'EX-C04' 2 'Estimated renal function at baseline (March 2026)' 'Yes' '1.0 mg/dL on 14 November 2025 and 12 March 2026 (two equal values four months apart)' 81 58 1.0 'mg/dL' 'female' 'case measurement' 'M-CR2, 12 March 2026'),
    (Q11R 'Q11-C04-1Jul' 'EX-C04' 2 'Estimated renal function on 1 July 2026' 'No' 'Creatinine rising during a documented deterioration' 81 58 1.7 'mg/dL' 'female' 'case measurement' 'M-CR3, 1 July 2026'),
    (Q11R 'Q11-C04-5Jul' 'EX-C04' 2 'Estimated renal function on 5 July 2026' 'No' 'Creatinine still rising (2.4 mg/dL)' 81 58 2.4 'mg/dL' 'female' 'case measurement' 'M-CR4, 5 July 2026')
)
$Q11Vec = @(
    (Q11R 'Q11-V27a' 'VEC-V27' 1 'Vector V27, case 1' 'Yes' 'Vector input' 70 70 1.0 'mg/dL' 'male'),
    (Q11R 'Q11-V27b' 'VEC-V27' 1 'Vector V27, case 2' 'Yes' 'Vector input' 70 70 2.0 'mg/dL' 'male'),
    (Q11R 'Q11-V27c' 'VEC-V27' 1 'Vector V27, case 3' 'Yes' 'Vector input' 70 70 1.0 'mg/dL' 'female'),
    (Q11R 'Q11-V27d' 'VEC-V27' 1 'Vector V27, case 4 (sex not recorded)' 'Yes' 'Vector input' 70 70 88.4 $SCRU 'not recorded'),
    (Q11R 'Q11-V27e' 'VEC-V27' 1 'Vector V27, case 5 (stability unconfirmed)' 'No' 'Vector input: stability not confirmed' 70 70 1.0 'mg/dL' 'male')
)
$Q6Vec = @(
    (QR 'Q6-V21' 'VEC-V21' 1 'Vector V21' (MergeH @(@{ Teaching_Switch = 'Yes'; Use = 'replacement or recovery'; Time_Unit = 'days' }, (PV 'Replacement_Half_Time' 3 $SV "$V V21"), (PV 't' 6 $SV "$V V21")))),
    (QR 'Q6-V32' 'VEC-V32' 1 'Vector V32: approach to a new level (protein C, hypothetical half-time 8 hours, at 24 hours)' (MergeH @(@{ Teaching_Switch = 'Yes'; Use = 'approach to a new level'; Time_Unit = 'hours' }, (PV 'Replacement_Half_Time' 8 $SV "$V V32"), (PV 't' 24 $SV "$V V32"))))
)

# ================================================================== REVIEWER SHEETS
function TxtCols($defs) { $a = @(); foreach ($d in $defs) { $a += (CIn $d[0] $d[1] $d[2] $d[3] $(if ($d[3] -match 'date') { 'd mmmm yyyy' } elseif ($d[3]) { $null } else { '@' })) }; return $a }
$PDSCols = @(TxtCols @(
    @('Statement_ID', 'Statement_ID', 12, $null), @('Case_ID', 'Case_ID (optional)', 10, $null), @('Product', 'Product', 14, $null),
    @('Product_Document_ID', 'Product_Document_ID (e.g. LANOXIN US)', 16, $null), @('Document_Type', 'Document type', 11, 'list:nLIST_DocType'),
    @('Jurisdiction', 'Jurisdiction', 10, $null), @('Product_Version', 'Version (as printed)', 22, $null), @('Revision_Date', 'Revision date (as printed)', 14, $null),
    @('Retrieval_Date', 'Retrieval date', 14, 'date'), @('Section', 'Section', 8, $null), @('Statement', 'Statement (extract, own words or quoted)', 50, $null),
    @('Timing_Statement_Type', 'Kind of timing statement (Chapter 23 proposal; only typical onset and maximum reported describe latency)', 22, 'list:nLIST_TimingType'),
    @('Definitional_Window', 'Window built into an event definition? (never temporal evidence)', 11, 'list:nLIST_YesNo'),
    @('Label_Discrepancy', 'Label_Discrepancy (US vs EU)', 26, $null), @('Rationale', 'Rationale', 26, $null))) + @(
    (CF 'Provenance_Check' 'Provenance check' '=IF({Statement}="","",IF(OR({Product_Document_ID}="",{Product_Version}="",{Section}=""),"Incomplete: document, version and section are needed","Document, version and section recorded"))' 30 $null 'CHECK'))
$PDSEx = @(
    @{ Statement_ID = 'PD-DIG-1'; Case_ID = 'EX-C04'; Product = 'Digoxin'; Product_Document_ID = 'LANOXIN US'; Document_Type = 'USPI'; Jurisdiction = 'US'; Product_Version = 'NDA 020405, supplement 15'; Revision_Date = '02/2019'; Retrieval_Date = '2026-10-20'; Section = '12.3'; Statement = 'First-order elimination; 50-70% of an IV dose excreted unchanged in urine (healthy volunteers); renal excretion proportional to creatinine clearance; half-life 1.5-2 days with normal renal function, 3.5-5 days in anuric patients.'; Timing_Statement_Type = 'not a timing statement'; Definitional_Window = 'No'; Rationale = 'Label facts (PK)' },
    @{ Statement_ID = 'PD-DIG-2'; Case_ID = 'EX-C04'; Product = 'Digoxin'; Product_Document_ID = 'LANOXIN US'; Document_Type = 'USPI'; Jurisdiction = 'US'; Product_Version = 'NDA 020405, supplement 15'; Revision_Date = '02/2019'; Retrieval_Date = '2026-10-20'; Section = '12.3; 8.6'; Statement = 'Renal impairment: prolonged half-life and greater exposure; a longer time to reach an initial or new steady state.'; Timing_Statement_Type = 'not a timing statement'; Definitional_Window = 'No'; Rationale = 'Label facts (PK)' },
    @{ Statement_ID = 'PD-DIG-3'; Case_ID = 'EX-C04'; Product = 'Digoxin'; Product_Document_ID = 'LANOXIN US'; Document_Type = 'USPI'; Jurisdiction = 'US'; Product_Version = 'NDA 020405, supplement 15'; Revision_Date = '02/2019'; Retrieval_Date = '2026-10-20'; Section = '5.3'; Statement = 'Toxicity usually associated with levels above 2 ng/mL, although symptoms may occur at lower levels; impaired renal function, low body weight, advanced age and electrolyte disturbances may predispose.'; Timing_Statement_Type = 'not a timing statement'; Definitional_Window = 'No'; Rationale = 'Label facts (event pattern)' },
    @{ Statement_ID = 'PD-DIG-4'; Case_ID = 'EX-C04'; Product = 'Digoxin'; Product_Document_ID = 'LANOXIN US'; Document_Type = 'USPI'; Jurisdiction = 'US'; Product_Version = 'NDA 020405, supplement 15'; Revision_Date = '02/2019'; Retrieval_Date = '2026-10-20'; Section = '2.5'; Statement = 'Sample just before the next scheduled dose or at least 6 hours after the last dose.'; Timing_Statement_Type = 'monitoring or precaution window'; Definitional_Window = 'No'; Rationale = 'Sampling advice; not latency' },
    @{ Statement_ID = 'PD-RTX-1'; Case_ID = 'EX-C28'; Product = 'Rituximab'; Product_Document_ID = 'RITUXAN US'; Document_Type = 'USPI'; Jurisdiction = 'US'; Product_Version = 'BLA 103705, supplement 5467'; Revision_Date = '12/2021'; Retrieval_Date = '2026-10-19'; Section = '5.3'; Statement = 'HBV reactivation has been reported up to 24 months after completion of therapy.'; Timing_Statement_Type = 'maximum or distribution reported'; Definitional_Window = 'No'; Label_Discrepancy = 'Timing in one region only; do not read EU silence as no timing'; Rationale = 'Reported maximum, not a limit' },
    @{ Statement_ID = 'PD-RTX-2'; Case_ID = 'EX-C28'; Product = 'Rituximab'; Product_Document_ID = 'MabThera EU'; Document_Type = 'SmPC'; Jurisdiction = 'EU'; Product_Version = 'EPAR product information, EMEA/H/C/000165'; Revision_Date = 'Updated 6 August 2026'; Retrieval_Date = '2026-10-19'; Section = '4.4'; Statement = 'HBV screening and specialist management of seropositive patients; no timing stated for reactivation.'; Timing_Statement_Type = 'not a timing statement'; Definitional_Window = 'No'; Label_Discrepancy = 'EU gives no timing'; Rationale = 'Management instruction' },
    @{ Statement_ID = 'PD-RTX-3'; Case_ID = 'EX-C28'; Product = 'Rituximab'; Product_Document_ID = 'RITUXAN US'; Document_Type = 'USPI'; Jurisdiction = 'US'; Product_Version = 'BLA 103705, supplement 5467'; Revision_Date = '12/2021'; Retrieval_Date = '2026-10-19'; Section = '12.3'; Statement = 'NHL: terminal half-life median 22 days (range 6.1-52 days); rituximab detectable in serum 3-6 months after the last dose.'; Timing_Statement_Type = 'not a timing statement'; Definitional_Window = 'No'; Label_Discrepancy = 'None (EU 5.2 gives the same numbers)'; Rationale = 'Exposure persistence' },
    @{ Statement_ID = 'PD-RTX-4'; Case_ID = 'EX-C28'; Product = 'Rituximab'; Product_Document_ID = 'RITUXAN US'; Document_Type = 'USPI'; Jurisdiction = 'US'; Product_Version = 'BLA 103705, supplement 5467'; Revision_Date = '12/2021'; Retrieval_Date = '2026-10-19'; Section = '12.2'; Statement = 'NHL: B-cell recovery begins about 6 months after treatment; median levels normal by 12 months.'; Timing_Statement_Type = 'not a timing statement'; Definitional_Window = 'No'; Label_Discrepancy = 'EU 5.1 adds a longer recovery in some settings'; Rationale = 'Pharmacodynamic recovery' },
    @{ Statement_ID = 'PD-RTX-5'; Case_ID = 'EX-C28'; Product = 'Rituximab'; Product_Document_ID = 'MabThera EU'; Document_Type = 'SmPC'; Jurisdiction = 'EU'; Product_Version = 'EPAR product information, EMEA/H/C/000165'; Revision_Date = 'Updated 6 August 2026'; Retrieval_Date = '2026-10-19'; Section = '5.1'; Statement = 'B-cell recovery up to a median of 23 months after induction in some patients.'; Timing_Statement_Type = 'not a timing statement'; Definitional_Window = 'No'; Label_Discrepancy = 'Longer than the US statement'; Rationale = 'Pharmacodynamic recovery' },
    @{ Statement_ID = 'PD-RTX-6'; Product = 'Rituximab'; Product_Document_ID = 'RITUXAN US'; Document_Type = 'USPI'; Jurisdiction = 'US'; Product_Version = 'BLA 103705, supplement 5467'; Revision_Date = '12/2021'; Retrieval_Date = '2026-10-19'; Section = '6.1'; Statement = 'CLL trials defined late-onset neutropenia as starting 42 days or more after the last dose.'; Timing_Statement_Type = 'definitional threshold'; Definitional_Window = 'Yes'; Rationale = 'A window built into an event definition is never temporal evidence' }
)
$PRCols = @(TxtCols @(
    @('Parameter_ID', 'Parameter_ID', 12, $null), @('Product', 'Product', 12, $null), @('Parameter_Name', 'Parameter name', 22, $null),
    @('Value', 'Value', 8, 'num'), @('Range_Lower', 'Range lower', 8, 'num'), @('Range_Upper', 'Range upper', 8, 'num'), @('Unit', 'Unit', 9, $null),
    @('Statistic', 'Statistic (median, mean, range)', 12, $null), @('HalfLife_Type', 'Half-life type', 14, 'list:nLIST_HalfLifeType'),
    @('Population', 'Population', 26, $null), @('Route', 'Route', 10, $null), @('Regimen', 'Regimen', 14, $null),
    @('Parameter_Status', 'Parameter_Status', 14, 'list:nLIST_Status'), @('Source_Document_ID', 'Source document', 14, $null),
    @('Section', 'Section', 8, $null), @('Version', 'Version', 18, $null), @('Retrieval_Date', 'Retrieval date', 14, 'date'),
    @('Missing_Reason', 'Missing-data reason', 12, 'listwarn:nLIST_MissingReason'))) + @(
    (CF 'Ready_Check' 'Ready for use?' '=IF({Parameter_Name}="","",IF({Parameter_Status}="","Not ready: Parameter_Status missing",IF({Population}="","Not ready: population missing",IF(AND(ISNUMBER(SEARCH("half",{Parameter_Name})),{HalfLife_Type}=""),"Not ready: half-life type missing",IF(OR({Source_Document_ID}="",{Section}=""),"Not ready: source document and section missing","Ready: status, population and source recorded")))))' 34 $null 'CHECK'))
$PREx = @(
    @{ Parameter_ID = 'PR-DIG-1'; Product = 'Digoxin'; Parameter_Name = 'Terminal elimination half-life'; Range_Lower = 1.5; Range_Upper = 2; Unit = 'days'; Statistic = 'range'; HalfLife_Type = 'terminal elimination'; Population = 'Healthy volunteers with normal renal function'; Parameter_Status = 'label fact'; Source_Document_ID = 'LANOXIN US'; Section = '12.3'; Version = 'Revised 02/2019'; Retrieval_Date = '2026-10-20' },
    @{ Parameter_ID = 'PR-DIG-2'; Product = 'Digoxin'; Parameter_Name = 'Terminal elimination half-life'; Range_Lower = 3.5; Range_Upper = 5; Unit = 'days'; Statistic = 'range'; HalfLife_Type = 'terminal elimination'; Population = 'Anuric patients'; Parameter_Status = 'label fact'; Source_Document_ID = 'LANOXIN US'; Section = '12.3'; Version = 'Revised 02/2019'; Retrieval_Date = '2026-10-20' },
    @{ Parameter_ID = 'PR-DIG-3'; Product = 'Digoxin'; Parameter_Name = 'Fraction of an IV dose excreted unchanged in urine'; Range_Lower = 0.5; Range_Upper = 0.7; Unit = 'fraction'; Statistic = 'range'; HalfLife_Type = 'not applicable'; Population = 'Healthy volunteers, intravenous dose'; Route = 'intravenous'; Parameter_Status = 'label fact'; Source_Document_ID = 'LANOXIN US'; Section = '12.3'; Version = 'Revised 02/2019'; Retrieval_Date = '2026-10-20' },
    @{ Parameter_ID = 'PR-RTX-1'; Product = 'Rituximab'; Parameter_Name = 'Terminal half-life (NHL)'; Value = 22; Range_Lower = 6.1; Range_Upper = 52; Unit = 'days'; Statistic = 'median (range)'; HalfLife_Type = 'terminal elimination'; Population = 'Patients with NHL'; Route = 'intravenous'; Parameter_Status = 'label fact'; Source_Document_ID = 'RITUXAN US'; Section = '12.3'; Version = 'Revised 12/2021'; Retrieval_Date = '2026-10-19' },
    @{ Parameter_ID = 'PR-RTX-2'; Product = 'Rituximab'; Parameter_Name = 'Terminal half-life (CLL)'; Value = 32; Range_Lower = 14; Range_Upper = 62; Unit = 'days'; Statistic = 'median (US) / mean (EU), range'; HalfLife_Type = 'terminal elimination'; Population = 'Patients with CLL'; Route = 'intravenous'; Parameter_Status = 'label fact'; Source_Document_ID = 'RITUXAN US'; Section = '12.3'; Version = 'Revised 12/2021'; Retrieval_Date = '2026-10-19' }
)
$PDPCols = @(TxtCols @(
    @('Case_ID', 'Case_ID', 10, $null), @('Case_Version', 'Case_Version', 7, 'posint'), @('Link_ID', 'Link_ID', 7, $null),
    @('Proposed_Mechanism', 'Proposed mechanism of harm (MOA_Hypothesis)', 30, $null), @('Mechanism_Status', 'Established or hypothetical?', 11, 'list:nLIST_MechStatus'),
    @('Delay_Steps', 'Biological steps that could impose delay', 28, $null), @('PD_Statements', 'What the documents say about PD onset, duration and recovery (with sources)', 34, $null),
    @('Measured_Markers', 'Measured markers in the case, with dates', 28, $null), @('Missing_Information', 'What is missing', 24, $null),
    @('Expected_PD_Course', 'Expected_PD_Course (reviewer, in words)', 30, $null)))
$PDPEx = @(
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Link_ID = 'L1'; Proposed_Mechanism = 'Loss of immune control of HBV during B-cell depletion (Stage 1 HBV profile)'; Mechanism_Status = 'hypothetical'; Delay_Steps = 'B-cell depletion and slow recovery; withdrawal of antiviral prophylaxis; viral replication to a detectable level'; PD_Statements = 'B cells depleted within 3 weeks; recovery begins about 6 months, median normal by 12 months in NHL [RITUXAN US 12.2]; up to a median of 23 months in some [MabThera EU 5.1]'; Measured_Markers = 'CD19 B cells and IgG below reference range on 12 May 2026 (M-BCELL)'; Missing_Information = 'B-cell and IgG values; earlier counts'; Expected_PD_Course = 'Risk follows immune suppression rather than drug concentration and can persist through B-cell depletion and recovery' }
)
$TPRCols = @(TxtCols @(
    @('Case_ID', 'Case_ID', 10, $null), @('Case_Version', 'Case_Version', 7, 'posint'), @('Link_ID', 'Link_ID', 7, $null),
    @('MOA_Hypothesis', 'MOA_Hypothesis', 26, $null), @('Expected_Exposure_Course', 'Expected_Exposure_Course', 28, $null), @('Expected_PD_Course', 'Expected_PD_Course', 26, $null),
    @('Clinical_Latency_Evidence', 'Clinical_Latency_Evidence (with sources)', 28, $null), @('Observed_Course', 'Observed_Course (read from Intervals, never retyped)', 28, $null),
    @('Interval_IDs_Cited', 'Interval_IDs cited', 12, $null), @('Temporal_Fit_Rationale', 'Temporal_Fit_Rationale', 34, $null)))
$TPREx = @(
    @{ Case_ID = 'EX-C04'; Case_Version = 2; Link_ID = 'L1'; MOA_Hypothesis = 'Concentration-related toxicity after reduced renal clearance at an unchanged dose'; Expected_Exposure_Course = 'Exposure rises over days at a rate set by the lengthening half-life after clearance falls [LANOXIN US 12.3; 8.6]'; Expected_PD_Course = 'Nausea and bradycardia follow concentration; beta-blocker adds AV-conduction effect [LANOXIN US 5.3; 7.3]'; Clinical_Latency_Evidence = 'No fixed latency; toxicity can appear during long, previously tolerated treatment when clearance falls'; Observed_Course = 'See Intervals I01-I06'; Interval_IDs_Cited = 'I01; I02; I03; I05; I06'; Temporal_Fit_Rationale = 'Onset 4 calendar days after the first documented renal deterioration and 9 after ibuprofen started, with a well-timed level above 2 ng/mL' },
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Link_ID = 'L1'; MOA_Hypothesis = 'Loss of immune control during B-cell depletion (hypothesis)'; Expected_Exposure_Course = 'Exposure negligible within months after the last infusion [RITUXAN US 12.3]'; Expected_PD_Course = 'B-cell depletion persists; recovery from about 6 months, may take a year or more'; Clinical_Latency_Evidence = 'Reactivation reported up to 24 months after completion [RITUXAN US 5.3]; EU gives no timing'; Observed_Course = 'See Intervals I01-I11'; Interval_IDs_Cited = 'I01; I02; I05; I06; I08'; Temporal_Fit_Rationale = 'Onset bounded 343 to 435 calendar days after the last infusion, within the reported period, with B cells and IgG still low at detection' }
)
$BIOCols = @(TxtCols @(
    @('Case_ID', 'Case_ID', 10, $null), @('Case_Version', 'Case_Version', 7, 'posint'), @('Link_ID', 'Link_ID', 7, $null),
    @('Biological_Fit', 'Biological_Fit (reviewer judgment; never set by a formula)', 22, 'list:nLIST_Fit'), @('Rationale', 'Rationale (required)', 44, $null),
    @('Alternative_Reading', 'Alternative defensible reading', 40, $null), @('Reviewer', 'Reviewer', 12, $null), @('Review_Date', 'Review date', 14, 'date'))) + @(
    (CF 'Check' 'Check' '=IF({Biological_Fit}="","",IF({Rationale}="","A rationale is required for every Biological_Fit entry","Rationale recorded; this is not a relatedness category"))' 34 $null 'CHECK'))
$BIOEx = @(
    @{ Case_ID = 'EX-C01'; Case_Version = 2; Link_ID = 'L1'; Biological_Fit = 'compatible'; Rationale = 'A 15-minute interval is compatible with an acute infusion-associated reaction; the timing does not identify which exposure (Drug A, Drug B and the antiseptic all preceded onset by an hour or less).'; Alternative_Reading = 'More weight to Drug B: previous courses, exposure ended 25 minutes before onset through the same cannula.' },
    @{ Case_ID = 'EX-C04'; Case_Version = 2; Link_ID = 'L1'; Biological_Fit = 'compatible'; Rationale = 'Onset followed documented renal deterioration during unchanged dosing, with a well-timed digoxin level above 2 ng/mL; prior tolerance does not exclude increased exposure after a clearance change (Chapter 11).'; Alternative_Reading = 'On case version 1 alone: compatible, with increased exposure plausible rather than supported, until the level and creatinine series are known.' },
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Link_ID = 'L1'; Biological_Fit = 'compatible'; Rationale = 'Exposure persistence is negligible on every labeled half-life; anti-HBc positive at baseline; B cells and IgG still low at detection; detection falls within the period for which the US label reports cases.'; Alternative_Reading = 'Treat the end of prophylaxis as the most informative anchor: reactivation after withdrawal of prophylaxis during persistent B-cell depletion.' }
)
$DCRCols = @(TxtCols @(
    @('Case_ID', 'Case_ID', 10, $null), @('Case_Version', 'Case_Version', 7, 'posint'), @('Link_ID', 'Link_ID', 7, $null),
    @('Action', 'Action', 14, 'list:nLIST_Action'), @('Action_Record_ID', 'Action record (timepoint ID)', 14, $null), @('Response_Record_ID', 'Response milestone ID', 14, $null),
    @('Recovery_Interval_ID', 'Recovery interval (Intervals, C7)', 10, $null), @('Interventions_Recorded', 'Interventions between action and response', 30, $null),
    @('Interventions_Status', 'Interventions status (recorded / none recorded, unknown)', 18, $null), @('Confounding', 'Confounding', 30, $null), @('Dechallenge_Rationale', 'Dechallenge_Rationale', 36, $null)))
$DCREx = @(
    @{ Case_ID = 'EX-C01'; Case_Version = 2; Link_ID = 'L1'; Action = 'withdrawal'; Action_Record_ID = 'EX-A/STOP'; Response_Record_ID = 'M-IMPROVE'; Recovery_Interval_ID = 'I10'; Interventions_Recorded = 'Emergency medicines from 09:18 (M-EMERG)'; Interventions_Status = 'recorded'; Confounding = 'Improvement followed both cessation and emergency treatment'; Dechallenge_Rationale = 'Recovery cannot be attributed to withdrawal alone (W_INTERVENTION_BEFORE_RECOVERY); Chapter 17' },
    @{ Case_ID = 'EX-C04'; Case_Version = 2; Link_ID = 'L1'; Action = 'withdrawal'; Action_Record_ID = 'EX-DIG/STOP'; Response_Record_ID = 'M-IMPROVE'; Recovery_Interval_ID = 'I06'; Interventions_Recorded = 'IV fluids and cardiac monitoring (M-FLUIDS); ibuprofen withheld the same day'; Interventions_Status = 'recorded; other treatments 5-8 July not reported'; Confounding = 'Renal function recovering; ibuprofen stopped'; Dechallenge_Rationale = 'Not clearly positive: improvement coincided with fluids, the end of ibuprofen and renal recovery' },
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Link_ID = 'L1'; Action = 'completed before onset'; Action_Record_ID = 'EX-RTX/STOP'; Interventions_Status = 'not applicable'; Dechallenge_Rationale = 'Therapy was completed before onset; no dechallenge applies' }
)
$ALTCols = @(TxtCols @(
    @('Case_ID', 'Case_ID', 10, $null), @('Case_Version', 'Case_Version', 7, 'posint'), @('Item_Type', 'Item type', 14, 'list:nLIST_ItemType'),
    @('Description', 'Description (competing cause or query)', 40, $null), @('Timing_Evidence', 'Timing evidence', 26, $null),
    @('Alternative_Cause_Assessment', 'Alternative_Cause_Assessment', 30, $null), @('Query_Recipient', 'Query recipient', 14, $null),
    @('Followup_Priority', 'Followup_Priority', 10, 'list:nLIST_Priority'), @('Status', 'Status', 10, 'list:nLIST_QueryStatus'))) + @(
    (CF 'Export_Safe_Description' 'Export-safe description (CSV formula-injection escape)' '=IF({Description}="","",IF(OR(LEFT({Description},1)="=",LEFT({Description},1)="+",LEFT({Description},1)="-",LEFT({Description},1)="@",LEFT({Description},1)=CHAR(9),LEFT({Description},1)=CHAR(13)),"''"&{Description},{Description}))' 40 $null 'EXPORT'))
$ALTEx = @(
    @{ Case_ID = 'EX-C01'; Case_Version = 2; Item_Type = 'alternative cause'; Description = 'Reaction to Drug B (same cannula; previous courses)'; Timing_Evidence = 'Ended 25 minutes before onset (I05)'; Alternative_Cause_Assessment = 'Timing cannot separate it from Drug A' },
    @{ Case_ID = 'EX-C01'; Case_Version = 2; Item_Type = 'follow-up query'; Description = 'Tryptase result; emergency medicines with doses and times; Drug A concentration in the bag; earlier Drug B courses'; Query_Recipient = 'Reporting hospital'; Followup_Priority = 'high'; Status = 'open' },
    @{ Case_ID = 'EX-C04'; Case_Version = 2; Item_Type = 'alternative cause'; Description = 'Bradycardia from the unchanged beta-blocker, made more likely by dehydration'; Timing_Evidence = 'Dose unchanged since 2024'; Alternative_Cause_Assessment = 'Explains susceptibility better than timing' },
    @{ Case_ID = 'EX-C04'; Case_Version = 2; Item_Type = 'follow-up query'; Description = 'Creatinine values between March and July; magnesium, calcium, thyroid function; earlier ECGs; full treatment record 5-8 July'; Query_Recipient = 'GP and hospital'; Followup_Priority = 'high'; Status = 'open' },
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Item_Type = 'alternative cause'; Description = 'Prednisolone course 2-11 March 2026'; Timing_Evidence = 'Onset bounds -21 to 71 calendar days (I08)'; Alternative_Cause_Assessment = 'Sequence uncertain; possible contributor' },
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Item_Type = 'follow-up query'; Description = 'Planned length of prophylaxis and why it stopped; prednisolone dose; lymphoma relapse or further treatment'; Query_Recipient = 'Hematology'; Followup_Priority = 'high'; Status = 'open' }
)
$AQCCols = @(TxtCols @(
    @('Case_ID', 'Case_ID', 10, $null), @('Case_Version', 'Case_Version', 7, 'posint'), @('Narrative', 'Narrative (reviewer)', 50, $null),
    @('Formal_Method', 'Formal method used (e.g. WHO-UMC)', 14, $null), @('Method_Result_Entered_By_Assessor', 'Method result (entered by the assessor; never calculated)', 18, $null),
    @('Reviewer', 'Reviewer', 12, $null), @('Review_Date', 'Review date', 14, 'date'), @('Unresolved_Issues', 'Unresolved issues', 34, $null),
    @('Version_Printed', 'Case version printed in the summary', 10, $null))) + @(
    (CF 'Interval_Rows' 'Interval rows for this case version' '=IF({Case_ID}="","",COUNTIF(nINT_CaseKey,{Case_ID}&"|"&{Case_Version}))' 10 '0' 'CHECKS'),
    (CF 'Interval_Rows_Blocked' 'Interval rows with an E-code (query each)' '=IF({Case_ID}="","",COUNTIFS(nINT_CaseKey,{Case_ID}&"|"&{Case_Version},nINT_K_Err,"E_*"))' 10 '0'),
    (CF 'Interval_Rows_With_W' 'Interval rows with a W-code (print with result)' '=IF({Case_ID}="","",COUNTIFS(nINT_CaseKey,{Case_ID}&"|"&{Case_Version},nINT_QC_Flags,"*W_*"))' 10 '0'),
    (CF 'Export_Safe_Narrative' 'Export-safe narrative (CSV escape)' '=IF({Narrative}="","",IF(OR(LEFT({Narrative},1)="=",LEFT({Narrative},1)="+",LEFT({Narrative},1)="-",LEFT({Narrative},1)="@",LEFT({Narrative},1)=CHAR(9),LEFT({Narrative},1)=CHAR(13)),"''"&{Narrative},{Narrative}))' 40))
$AQCEx = @(
    @{ Case_ID = 'EX-C01'; Case_Version = 2; Narrative = 'Urticaria, wheeze and hypotension began at 09:15, 15 minutes after the start of the first Drug A infusion and 25 minutes after the end of a Drug B infusion through the same cannula (C1; UTC-05:00). Improvement followed both cessation and emergency treatment.'; Formal_Method = 'not entered (example)'; Unresolved_Issues = 'Tryptase result; emergency medicines; Drug A concentration; Drug B history'; Version_Printed = '2' },
    @{ Case_ID = 'EX-C04'; Case_Version = 2; Narrative = 'Nausea and symptomatic bradycardia began on 5 July 2026, 185 calendar days after digoxin was started at an unchanged dose, 4 calendar days after renal deterioration was first documented and 9 after ibuprofen was started.'; Formal_Method = 'not entered (example)'; Unresolved_Issues = 'Creatinine series; electrolytes; ECG history; treatments 5-8 July'; Version_Printed = '2' },
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Narrative = 'HBV DNA was first detected on 12 May 2026, 435 calendar days after the last rituximab infusion; virological reactivation began between 343 and 435 days after the last infusion (C4).'; Formal_Method = 'not entered (example)'; Unresolved_Issues = 'Prophylaxis plan; prednisolone dose; relapse or further treatment'; Version_Printed = '2' }
)

# ================================================================== EXAMPLES AND VERIFICATION CHECKS
# Each check: @(vectorOrCase, description, table, column, key, kind, expected, digits, stated)
$ExChecks = @(
    @('EX-C01', 'Ch 3, 3.7 (Table 8)', 'Drug A start to onset, FIRST_ADMIN (C1)', 'INT', 'Interval_Value', 'EX-C01|2|I01', 'num', 15, 0, '15 minutes'),
    @('EX-C01', 'Ch 3, 3.7', 'Sequence of I01', 'INT', 'Sequence_Status', 'EX-C01|2|I01', 'text', 'AFTER', $null, 'AFTER'),
    @('EX-C01', 'Ch 3, 3.7 (Table 8)', 'Same pair, calendar-date difference (C2)', 'INT', 'Interval_Value', 'EX-C01|2|I02', 'num', 0, 0, '0 calendar days'),
    @('EX-C01', 'Ch 3, 3.7 (Table 8)', 'C2 order decided by C1', 'INT', 'Sequence_Status', 'EX-C01|2|I02', 'text', 'AFTER', $null, 'order decided by C1 (AFTER)'),
    @('EX-C01', 'Ch 3, 3.7 (Calculations)', 'Treatment Day under the inclusive convention', 'INT', 'Treatment_Day', 'EX-C01|2|I02', 'num', 1, 0, 'Treatment Day 1'),
    @('EX-C01', 'Ch 3, 3.7 (Question 3)', 'LAST_DOSE_BEFORE_ONSET selects the 09:00 infusion start (C5 candidates: start, administration, dose-change, restart)', 'INT', 'Anchor_Record_ID_Used', 'EX-C01|2|I03', 'text', 'EX-A/START', $null, 'infusion started 09:00'),
    @('EX-C01', 'Ch 3, 3.7 (Question 2)', 'LAST_DOSE_BEFORE_ONSET to onset (C1)', 'INT', 'Interval_Value', 'EX-C01|2|I03', 'num', 15, 0, '15 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Table 5)', 'Drug B infusion start to onset', 'INT', 'Interval_Value', 'EX-C01|2|I04', 'num', 55, 0, '55 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Table 5)', 'Drug B infusion end to onset', 'INT', 'Interval_Value', 'EX-C01|2|I05', 'num', 25, 0, '25 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Table 5)', 'Skin antiseptic to onset', 'INT', 'Interval_Value', 'EX-C01|2|I06', 'num', 60, 0, '60 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Table 5)', 'Baseline observations to onset', 'INT', 'Interval_Value', 'EX-C01|2|I07', 'num', 25, 0, '25 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Table 5)', 'Drug A stopped (CESSATION) to onset', 'INT', 'Interval_Value', 'EX-C01|2|I08', 'num', -1, 0, '-1 minute'),
    @('EX-C01', 'Ch 3, 3.7 (Table 5)', 'Sequence: the stop came after onset', 'INT', 'Sequence_Status', 'EX-C01|2|I08', 'text', 'BEFORE', $null, 'exposure continued 1 minute after onset'),
    @('EX-C01', 'Ch 3, 3.7 (Table 5)', 'Emergency treatment to onset', 'INT', 'Interval_Value', 'EX-C01|2|I09', 'num', -3, 0, '-3 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Table 8)', 'Recovery: stop to improvement (C7 with C1)', 'INT', 'Interval_Value', 'EX-C01|2|I10', 'num', 19, 0, '19 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Table 8)', 'Recovery row raises W_INTERVENTION_BEFORE_RECOVERY', 'INT', 'QC_Flags', 'EX-C01|2|I10', 'contains', 'W_INTERVENTION_BEFORE_RECOVERY', $null, 'W_INTERVENTION_BEFORE_RECOVERY'),
    @('EX-C01', 'Ch 3, 3.7', 'Recovery row labeled C7', 'INT', 'Calculation_Mode_Used', 'EX-C01|2|I10', 'text', 'C7 (C1)', $null, 'C1 (C7)'),
    @('EX-C01', 'Ch 3, 3.7 (Calculations)', 'Emergency treatment to improvement', 'INT', 'Interval_Value', 'EX-C01|2|I11', 'num', 17, 0, '17 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Calculations)', 'Onset to improvement', 'INT', 'Interval_Value', 'EX-C01|2|I12', 'num', 20, 0, '20 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Calculations)', 'Drug A ran 09:00 to 09:16', 'INT', 'Interval_Value', 'EX-C01|2|I13', 'num', 16, 0, '16 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Biological reasoning)', 'Tryptase sample after onset', 'INT', 'Interval_Value', 'EX-C01|2|I14', 'num', 50, 0, '50 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Calculations)', 'Amount infused (Q9): rate only in mL/hour, not calculable', 'Q9', 'QC_Flags', 'EX-C01|2|Q9-C01', 'text', 'E_INVALID_UNIT', $null, 'not calculable from the available information'),
    @('EX-C04', 'Ch 11, 11.8 (Q1)', 'Digoxin start to onset, FIRST_ADMIN (C2)', 'INT', 'Interval_Value', 'EX-C04|2|I01', 'num', 185, 0, '185 calendar days'),
    @('EX-C04', 'Ch 11, 11.8 (Q1)', 'First documented renal deterioration to onset', 'INT', 'Interval_Value', 'EX-C04|2|I02', 'num', 4, 0, '4 calendar days'),
    @('EX-C04', 'Ch 11, 11.8 (Q1)', 'Ibuprofen start (CONCOMITANT_CHANGE) to onset', 'INT', 'Interval_Value', 'EX-C04|2|I03', 'num', 9, 0, '9 calendar days'),
    @('EX-C04', 'Ch 23, 23.11 (version 2)', 'LAST_DOSE_BEFORE_ONSET selects the 4 July dose', 'INT', 'Anchor_Record_ID_Used', 'EX-C04|2|I04', 'text', 'ADM-DIG-LAST', $null, 'last dose about 08:00 on 4 July'),
    @('EX-C04', 'Ch 23, 23.11 (version 2)', 'Last dose before onset to onset (C2)', 'INT', 'Interval_Value', 'EX-C04|2|I04', 'num', 1, 0, '1 calendar day'),
    @('EX-C04', 'Ch 11, 11.8 (Calculations)', 'Last dose to digoxin sample (C1, local clock)', 'INT', 'Interval_Value', 'EX-C04|2|I05', 'num', 1660, 0, '1,660 minutes'),
    @('EX-C04', 'Ch 11, 11.8 (Calculations)', 'Unit of the sample interval', 'INT', 'Unit', 'EX-C04|2|I05', 'text', 'minutes (local clock)', $null, 'minutes, local clock'),
    @('EX-C04', 'Ch 11, 11.8 (Calculations)', 'Sample interval raises W_TZ_UNKNOWN', 'INT', 'QC_Flags', 'EX-C04|2|I05', 'contains', 'W_TZ_UNKNOWN', $null, 'W_TZ_UNKNOWN'),
    @('EX-C04', 'Ch 11, 11.8 (Q5)', 'Withholding to improvement (C7 with C2)', 'INT', 'Interval_Value', 'EX-C04|2|I06', 'num', 3, 0, '3 calendar days'),
    @('EX-C04', 'Ch 11, 11.8 (Q5)', 'Recovery row raises W_INTERVENTION_BEFORE_RECOVERY', 'INT', 'QC_Flags', 'EX-C04|2|I06', 'contains', 'W_INTERVENTION_BEFORE_RECOVERY', $null, 'W_INTERVENTION_BEFORE_RECOVERY'),
    @('EX-C04', 'Ch 11, 11.8 (Calculations)', 'Q11 at baseline, displayed', 'Q11', 'CrCl_mL_min', 'EX-C04|2|Q11-C04-baseline', 'num', 40, 0, '40 mL/minute'),
    @('EX-C04', 'Ch 11, 11.8 (Calculations)', 'Q11 at baseline, unrounded', 'Q11', 'CrCl_mL_min', 'EX-C04|2|Q11-C04-baseline', 'num', 40.399, 3, 'unrounded 40.399'),
    @('EX-C04', 'Ch 11, 11.8 (Calculations)', 'Q11 on 1 July is blocked', 'Q11', 'QC_Flags', 'EX-C04|2|Q11-C04-1Jul', 'text', 'E_CREATININE_STABILITY_UNCONFIRMED', $null, 'E_CREATININE_STABILITY_UNCONFIRMED'),
    @('EX-C04', 'Ch 11, 11.8 (Calculations)', 'Q11 on 5 July is blocked', 'Q11', 'QC_Flags', 'EX-C04|2|Q11-C04-5Jul', 'text', 'E_CREATININE_STABILITY_UNCONFIRMED', $null, 'E_CREATININE_STABILITY_UNCONFIRMED'),
    @('EX-C04', 'Ch 23, 23.11 (Q8)', 'Normal renal function, t_half 1.5 days: percent remaining', 'Q8', 'Percent_at_Lower', 'EX-C04|2|Q8-C04-normal', 'num', 25.0, 1, '25.0%'),
    @('EX-C04', 'Ch 23, 23.11 (Q8)', 'Normal renal function, t_half 2 days', 'Q8', 'Percent_at_Upper', 'EX-C04|2|Q8-C04-normal', 'num', 35.4, 1, '35.4%'),
    @('EX-C04', 'Ch 23, 23.11 (Q8)', 'Anuric, t_half 3.5 days', 'Q8', 'Percent_at_Lower', 'EX-C04|2|Q8-C04-anuric', 'num', 55.2, 1, '55.2%'),
    @('EX-C04', 'Ch 23, 23.11 (Q8)', 'Anuric, t_half 5 days', 'Q8', 'Percent_at_Upper', 'EX-C04|2|Q8-C04-anuric', 'num', 66.0, 1, '66.0%'),
    @('EX-C04', 'Ch 23, 23.11 (Q8)', 'Q8 rows carry I_SCENARIO_NOT_CI', 'Q8', 'QC_Flags', 'EX-C04|2|Q8-C04-normal', 'text', 'I_SCENARIO_NOT_CI', $null, 'I_SCENARIO_NOT_CI'),
    @('EX-C04', 'Ch 11, Worked exercise 11.2 (Q4)', 'Renal elimination halved, fe 0.5: ratio', 'Q4', 'Css_Ratio', 'EX-C04|2|Q4-C04-half-fe0.5', 'num', 1.33, 2, '1.33'),
    @('EX-C04', 'Ch 11, Worked exercise 11.2 (Q4)', 'Renal elimination halved, fe 0.7: ratio', 'Q4', 'Css_Ratio', 'EX-C04|2|Q4-C04-half-fe0.7', 'num', 1.54, 2, '1.54'),
    @('EX-C04', 'Ch 11, Worked exercise 11.2 (Q4)', 'Renal elimination at a quarter, fe 0.5: ratio', 'Q4', 'Css_Ratio', 'EX-C04|2|Q4-C04-quarter-fe0.5', 'num', 1.60, 2, '1.60'),
    @('EX-C04', 'Ch 11, Worked exercise 11.2 (Q4)', 'Renal elimination at a quarter, fe 0.7: ratio', 'Q4', 'Css_Ratio', 'EX-C04|2|Q4-C04-quarter-fe0.7', 'num', 2.11, 2, '2.11'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Last infusion to first positive test (C2)', 'INT', 'Interval_Value', 'EX-C28|2|I01', 'num', 435, 0, '435 calendar days'),
    @('EX-C28', 'Ch 20, 20.4 (Question 3)', 'Reporter''s fourteen months kept, not converted', 'INT', 'QC_Flags', 'EX-C28|2|I01', 'contains', 'I_APPROXIMATE_NOT_COMPARED', $null, 'I_APPROXIMATE_NOT_COMPARED'),
    @('EX-C28', 'Ch 20, 20.4 (Question 1-2)', 'LAST_DOSE_BEFORE_ONSET selects 3 March 2025', 'INT', 'Anchor_Record_ID_Used', 'EX-C28|2|I02', 'text', 'ADM-RTX-6', $null, 'last infusion 3 March 2025'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Last infusion to onset bounds: minimum (C4)', 'INT', 'Interval_Minimum', 'EX-C28|2|I02', 'num', 343, 0, '343'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Last infusion to onset bounds: maximum (C4)', 'INT', 'Interval_Maximum', 'EX-C28|2|I02', 'num', 435, 0, '435'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Sequence of the bounded interval', 'INT', 'Sequence_Status', 'EX-C28|2|I02', 'text', 'AFTER', $null, 'AFTER'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'First infusion to first positive test', 'INT', 'Interval_Value', 'EX-C28|2|I03', 'num', 540, 0, '540 calendar days'),
    @('EX-C28', 'Ch 23, 23.12', 'First infusion to onset bounds: minimum', 'INT', 'Interval_Minimum', 'EX-C28|2|I04', 'num', 448, 0, '448'),
    @('EX-C28', 'Ch 23, 23.12', 'First infusion to onset bounds: maximum', 'INT', 'Interval_Maximum', 'EX-C28|2|I04', 'num', 540, 0, '540'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'End of prophylaxis to first positive test', 'INT', 'Interval_Value', 'EX-C28|2|I05', 'num', 251, 0, '251 days'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'End of prophylaxis to onset bounds: minimum', 'INT', 'Interval_Minimum', 'EX-C28|2|I06', 'num', 159, 0, '159'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'End of prophylaxis to onset bounds: maximum', 'INT', 'Interval_Maximum', 'EX-C28|2|I06', 'num', 251, 0, '251'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Prednisolone start to first positive test', 'INT', 'Interval_Value', 'EX-C28|2|I07', 'num', 71, 0, '71 days'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Prednisolone start to onset bounds: minimum', 'INT', 'Interval_Minimum', 'EX-C28|2|I08', 'num', -21, 0, '-21'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Prednisolone start to onset bounds: maximum', 'INT', 'Interval_Maximum', 'EX-C28|2|I08', 'num', 71, 0, '71'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Prednisolone row: SEQUENCE_UNCERTAIN', 'INT', 'Sequence_Status', 'EX-C28|2|I08', 'text', 'SEQUENCE_UNCERTAIN', $null, 'SEQUENCE_UNCERTAIN'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Prednisolone row raises W_SEQUENCE_UNCERTAIN', 'INT', 'QC_Flags', 'EX-C28|2|I08', 'contains', 'W_SEQUENCE_UNCERTAIN', $null, 'W_SEQUENCE_UNCERTAIN'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Last infusion to biochemical hepatitis (26 May)', 'INT', 'Interval_Value', 'EX-C28|2|I09', 'num', 449, 0, '449 days'),
    @('EX-C28', 'Ch 20, 20.4 (Question 1)', 'Tests 92 calendar days apart', 'INT', 'Interval_Value', 'EX-C28|2|I10', 'num', 92, 0, '92 calendar days'),
    @('EX-C28', 'Ch 20, 20.4 (Table 9)', 'Prophylaxis stopped after the last infusion', 'INT', 'Interval_Value', 'EX-C28|2|I11', 'num', 184, 0, '184 days'),
    @('EX-C28', 'Ch 23, 23.12 (Q7)', 'Median t_half 22 days at 435 days: percent remaining', 'Q7', 'Percent_Remaining', 'EX-C28|2|Q7-C28-435', 'num', 0.000112, 6, '0.000112%'),
    @('EX-C28', 'Ch 23, 23.12 (Q7)', 'Median t_half 22 days at 343 days', 'Q7', 'Percent_Remaining', 'EX-C28|2|Q7-C28-343', 'num', 0.00203, 5, '0.00203%'),
    @('EX-C28', 'Ch 23, 23.12 (Q7)', 'Q7 rows carry W_BIOLOGIC_MODEL_CAUTION', 'Q7', 'QC_Flags', 'EX-C28|2|Q7-C28-435', 'text', 'W_BIOLOGIC_MODEL_CAUTION', $null, 'W_BIOLOGIC_MODEL_CAUTION'),
    @('EX-C28', 'Ch 23, 23.12 (Q8)', 'Longest labeled t_half 52 days at 435 days', 'Q8', 'Percent_at_Upper', 'EX-C28|2|Q8-C28-435', 'num', 0.303, 3, '0.303%'),
    @('EX-C28', 'Ch 23, 23.12 (Q8)', 'Shortest labeled t_half 6.1 days at 435 days', 'Q8', 'Percent_at_Lower', 'EX-C28|2|Q8-C28-435', 'lt', 0.0001, $null, 'below 0.0001%'),
    @('EX-C28', 'Ch 20, 20.4 (Question 5)', 'Longest labeled t_half 52 days at 343 days', 'Q8', 'Percent_at_Upper', 'EX-C28|2|Q8-C28-343', 'num', 1.03, 2, '1.03%'),
    @('EX-C01', 'Ch 3, 3.7 (Table 8)', 'Displayed value of I01 (as the calculator)', 'INT', 'Displayed_Value', 'EX-C01|2|I01', 'text', '15 minutes', $null, '15 minutes'),
    @('EX-C01', 'Ch 3, 3.7 (Table 5)', 'Displayed value of I08 (minus sign, singular unit)', 'INT', 'Displayed_Value', 'EX-C01|2|I08', 'text', ($U_MINUS + '1 minute'), $null, '-1 minute'),
    @('EX-C01', 'Ch 3, 3.7 (Table 8); CALC-SPEC 1.1 C7', 'C7 names the recorded intervention it counts', 'INT', 'Interventions_Counted', 'EX-C01|2|I10', 'text', 'M-EMERG', $null, 'emergency medicines from 09:18'),
    @('EX-C04', 'Ch 11, 11.8 (Calculations)', 'Displayed value of I05 (as the calculator)', 'INT', 'Displayed_Value', 'EX-C04|2|I05', 'text', '1,660 minutes (local clock)', $null, '1,660 minutes, local clock'),
    @('EX-C04', 'Ch 11, 11.8 (Q5); CALC-SPEC 1.1 C7', 'C7 counts the IV fluids recorded on the cessation date (ends included)', 'INT', 'Interventions_Counted', 'EX-C04|2|I06', 'text', 'M-FLUIDS', $null, 'IV fluids, 5 July'),
    @('EX-C04', 'Ch 11, 11.8 (Q5)', 'Recovery row labeled C7 (C2)', 'INT', 'Calculation_Mode_Used', 'EX-C04|2|I06', 'text', 'C7 (C2)', $null, 'C7 at date level'),
    @('EX-C04', 'Ch 23, 23.11 (Q8)', 'Q8 normal renal function: displayed 25.0%', 'Q8', 'Displayed_Value', 'EX-C04|2|Q8-C04-normal', 'contains', '25.0%', $null, '25.0%'),
    @('EX-C28', 'Ch 20, 20.4; CALC-SPEC 1.1 C4', 'Laboratory-defined onset: last normal test (9 February) to first abnormal test (12 May 2026)', 'INT', 'End_ISO', 'EX-C28|2|I02', 'text', '2026-02-09/2026-05-12', $null, '9 February to 12 May 2026'),
    @('EX-C28', 'CALC-SPEC 1.1 C4', 'Assumptions state that the onset is laboratory-defined', 'INT', 'Assumptions', 'EX-C28|2|I02', 'contains', 'laboratory-defined', $null, 'laboratory-defined onset'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Displayed value of I02 (as the calculator)', 'INT', 'Displayed_Value', 'EX-C28|2|I02', 'text', '343 to 435 calendar days', $null, '343 to 435 calendar days'),
    @('EX-C28', 'Ch 20, 20.4 (Question 2)', 'Displayed value of I08 (bounds across zero are signed)', 'INT', 'Displayed_Value', 'EX-C28|2|I08', 'text', ($U_MINUS + '21 to +71 calendar days'), $null, '-21 to +71 calendar days'),
    @('EX-C28', 'Ch 23, 23.12 (Q8); CALC-SPEC 1.1 Q8', 'Q8 for a recorded monoclonal antibody carries the Q7 caution', 'Q8', 'QC_Flags', 'EX-C28|2|Q8-C28-435', 'text', 'W_BIOLOGIC_MODEL_CAUTION; I_SCENARIO_NOT_CI', $null, 'W_BIOLOGIC_MODEL_CAUTION; I_SCENARIO_NOT_CI'),
    @('EX-C28', 'Ch 23, 23.12 (Q8); CALC-SPEC 1.1 Section 7', 'Shortest labeled t_half 6.1 days at 435 days, shown in powers of ten', 'Q8', 'Displayed_Value', 'EX-C28|2|Q8-C28-435', 'contains', ('3.41 ' + $U_TIMES + ' 10' + $U_SUP['-'] + $U_SUP['2'] + $U_SUP['0'] + '%'), $null, '3.41 x 10^-20 %'),
    @('EX-C28', 'Ch 23, 23.12 (calculator example T1)', 'Q8 median t_half 22 days at 435 days', 'Q8', 'Percent_at_Lower', 'EX-C28|2|Q8-C28-435m', 'num', 0.000112, 6, '0.000112%'),
    @('EX-C28', 'Ch 23, 23.12 (calculator example T1)', 'Q8 median t_half 22 days at 435 days: displayed', 'Q8', 'Displayed_Value', 'EX-C28|2|Q8-C28-435m', 'contains', '0.000112%', $null, '0.000112%'),
    @('EX-C28', 'Ch 23, 23.12 (calculator example T1)', 'Q8 longest t_half 52 days at 435 days', 'Q8', 'Percent_at_Upper', 'EX-C28|2|Q8-C28-435m', 'num', 0.303, 3, '0.303%'),
    @('EX-C28', 'Ch 23, 23.12 (calculator example T1)', 'Q8 antibody flags', 'Q8', 'QC_Flags', 'EX-C28|2|Q8-C28-435m', 'text', 'W_BIOLOGIC_MODEL_CAUTION; I_SCENARIO_NOT_CI', $null, 'W_BIOLOGIC_MODEL_CAUTION; I_SCENARIO_NOT_CI'),
    @('EX-C28', 'Ch 23, 23.12 (calculator example T2)', 'Q8 median t_half 22 days at 343 days', 'Q8', 'Percent_at_Lower', 'EX-C28|2|Q8-C28-343m', 'num', 0.00203, 5, '0.00203%'),
    @('EX-C28', 'Ch 23, 23.12 (calculator example T2)', 'Q8 longest t_half 52 days at 343 days', 'Q8', 'Percent_at_Upper', 'EX-C28|2|Q8-C28-343m', 'num', 1.03, 2, '1.03%')
)
$VerVectors = @(
    @('V01', 'C1', 'numeric', 'Intervals', '15 minutes; AFTER; no flags'),
    @('V02', 'C2+C3', 'numeric', 'Intervals', '5 calendar days; treatment Day 6; AFTER'),
    @('V03', 'C2', 'numeric', 'Intervals', '0 calendar days; SAME_DATE_ORDER_UNKNOWN; W_SAME_DATE_ORDER_UNKNOWN'),
    @('V04', 'C2', 'numeric', 'Intervals', '2 calendar days; AFTER'),
    @('V05', 'C4', 'numeric', 'Intervals', '5 to 14 calendar days; AFTER'),
    @('V06', 'C4', 'numeric', 'Intervals', '-5 to +4 calendar days; SEQUENCE_UNCERTAIN; W_SEQUENCE_UNCERTAIN'),
    @('V07', 'C5', 'numeric', 'Administrations, Intervals', 'Anchor 2026-03-02T08:00-05:00; 720 minutes; AFTER'),
    @('V08', 'C5', 'numeric', 'Exposures, Intervals', 'FIRST_ADMIN 44 and RESTART 4 calendar days, both kept'),
    @('V09', 'C1', 'numeric', 'Intervals', '60 minutes; AFTER'),
    @('V10', 'C1', 'numeric', 'Intervals', '120 minutes (local clock); W_TZ_UNKNOWN'),
    @('V11', 'C6', 'numeric', 'Intervals', '360 minutes; reported timing kept; W_REPORTED_TIMING_CONFLICT'),
    @('V12', 'Versioning', 'behavioral', 'Cases, Intervals', 'Version 1 (9 days) preserved; version 2 gives 7 days'),
    @('V13', 'Validation', 'numeric', 'Exposures, Events, Drug_Event_Links, Intervals', 'E_INVALID_DATE, E_REVERSED_BOUNDS, E_DUPLICATE_ID, E_MISSING_LINK'),
    @('V14', 'Export', 'behavioral', 'Export (outside the workbook)', 'Bounds, unit, precision, links, status and flags preserved on export and re-import of V06'),
    @('V15', 'C6', 'numeric', 'Intervals', 'No value; I_APPROXIMATE_NOT_COMPARED'),
    @('V16', 'Q1', 'numeric', 'PK_Exercises (Q1)', '0.125; k = 0.028881133 per hour'),
    @('V17', 'Q2', 'numeric', 'PK_Exercises (Q2)', '79.726274 hours; 3.321928 half-lives'),
    @('V18', 'Q3', 'numeric', 'PK_Exercises (Q3)', '3.414214'),
    @('V19', 'Q4', 'numeric', 'PK_Exercises (Q4)', 'Half-life and average Css ratios both 2'),
    @('V20', 'Q5', 'numeric', 'PK_Exercises (Q5)', '0.390625% and 62.996052%; no sum'),
    @('V21', 'Q6', 'numeric', 'PD_Profile (Q6)', '0.75 replaced; I_TEACHING_DEMONSTRATION'),
    @('V22', 'Q7', 'numeric', 'PK_Exercises (Q7)', '0.25; W_BIOLOGIC_MODEL_CAUTION'),
    @('V23', 'Q8', 'numeric', 'PK_Exercises (Q8)', '0.0625 to 0.25; I_SCENARIO_NOT_CI'),
    @('V24', 'Technical validation', 'numeric', 'PK_Exercises (Q1)', 'E_NONPOSITIVE_PARAMETER, E_INVALID_UNIT, E_APPLICABILITY_UNCONFIRMED'),
    @('V25', 'Q9', 'numeric', 'Administrations, PK_Exercises (Q9)', '58.333333 mg; 11.666667% of 500 mg'),
    @('V26', 'Q10', 'numeric', 'Administrations, PK_Exercises (Q10)', '164.052288 (Method A); 161.111111 / 170.588235 (Method B)'),
    @('V27', 'Q11', 'numeric', 'PK_Exercises (Q11)', '68.055556, 34.027778, female 57.847222; sex missing shows both; unconfirmed stability blocks'),
    @('V28', 'C7', 'numeric', 'Exposures, Milestones, Intervals', '19 minutes; AFTER; MS-RX counted; W_INTERVENTION_BEFORE_RECOVERY'),
    @('V29', 'C7', 'numeric', 'Exposures, Milestones, Intervals', '3 calendar days; AFTER; MS-FLU on the cessation date counted (ends included); W_INTERVENTION_BEFORE_RECOVERY'),
    @('V30', 'C7', 'numeric', 'Exposures, Milestones, Intervals', 'No-flag variant: 3 calendar days; AFTER; nothing counted (TX-START is an exposure, not an intervention)'),
    @('V31', 'C4 (laboratory-defined onset)', 'numeric', 'Events (LAB), Milestones, Intervals', 'Onset 2026-03-15 to 2026-03-29; 14 to 28 calendar days; AFTER; no flags'),
    @('V32', 'Q6', 'numeric', 'PD_Profile (Q6)', '0.875 of the change to a new level completed; I_TEACHING_DEMONSTRATION'),
    @('V33', 'C5', 'numeric', 'Exposures, Administrations, Intervals', 'Anchor 2026-03-02T08:00-05:00 (UTC order); 720 minutes; AFTER; no flags'),
    @('V34', 'C5', 'numeric', 'Exposures, Administrations, Intervals', 'A partial-date administration could be the last dose: no anchor, E_MISSING_INPUT'),
    @('V35', 'Error listing', 'numeric', 'PK_Exercises (Q1), Intervals', 'E_INVALID_UNIT; E_NONPOSITIVE_PARAMETER; E_APPLICABILITY_UNCONFIRMED, and E_INVALID_DATE; E_MISSING_LINK'),
    @('V36', 'Gate', 'numeric', 'PK_Exercises (Q1)', 'Blank detail: E_MISSING_INPUT; E_APPLICABILITY_UNCONFIRMED. Complete: 0.125. Model unconfirmed: E_APPLICABILITY_UNCONFIRMED'),
    @('V37', 'Q8', 'numeric', 'PK_Exercises (Q8)', '0.0625 to 0.25; W_BIOLOGIC_MODEL_CAUTION; I_SCENARIO_NOT_CI'),
    @('V38', 'Q10', 'numeric', 'Administrations, PK_Exercises (Q10)', '164.052288 (Method A); 161.111111 / 170.588235 (Method B by date)')
)
# @(vector, what, table, column, key, kind, expected)
$VerChecks = @(
    @('V01', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V01|1|I1', 'num', 15), @('V01', 'Unit', 'INT', 'Unit', 'VEC-V01|1|I1', 'text', 'minutes'), @('V01', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V01|1|I1', 'text', 'AFTER'), @('V01', 'Flags', 'INT', 'QC_Flags', 'VEC-V01|1|I1', 'flags', ''),
    @('V02', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V02|1|I1', 'num', 5), @('V02', 'Unit', 'INT', 'Unit', 'VEC-V02|1|I1', 'text', 'calendar days'), @('V02', 'Treatment day', 'INT', 'Treatment_Day', 'VEC-V02|1|I1', 'num', 6), @('V02', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V02|1|I1', 'text', 'AFTER'), @('V02', 'Flags', 'INT', 'QC_Flags', 'VEC-V02|1|I1', 'flags', ''),
    @('V03', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V03|1|I1', 'num', 0), @('V03', 'Unit', 'INT', 'Unit', 'VEC-V03|1|I1', 'text', 'calendar days'), @('V03', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V03|1|I1', 'text', 'SAME_DATE_ORDER_UNKNOWN'), @('V03', 'Flags', 'INT', 'QC_Flags', 'VEC-V03|1|I1', 'flags', 'W_SAME_DATE_ORDER_UNKNOWN'),
    @('V04', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V04|1|I1', 'num', 2), @('V04', 'Unit', 'INT', 'Unit', 'VEC-V04|1|I1', 'text', 'calendar days'), @('V04', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V04|1|I1', 'text', 'AFTER'), @('V04', 'Flags', 'INT', 'QC_Flags', 'VEC-V04|1|I1', 'flags', ''),
    @('V05', 'Interval_Minimum', 'INT', 'Interval_Minimum', 'VEC-V05|1|I1', 'num', 5), @('V05', 'Interval_Maximum', 'INT', 'Interval_Maximum', 'VEC-V05|1|I1', 'num', 14), @('V05', 'Unit', 'INT', 'Unit', 'VEC-V05|1|I1', 'text', 'calendar days'), @('V05', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V05|1|I1', 'text', 'AFTER'), @('V05', 'Flags', 'INT', 'QC_Flags', 'VEC-V05|1|I1', 'flags', ''),
    @('V06', 'Interval_Minimum', 'INT', 'Interval_Minimum', 'VEC-V06|1|I1', 'num', -5), @('V06', 'Interval_Maximum', 'INT', 'Interval_Maximum', 'VEC-V06|1|I1', 'num', 4), @('V06', 'Unit', 'INT', 'Unit', 'VEC-V06|1|I1', 'text', 'calendar days'), @('V06', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V06|1|I1', 'text', 'SEQUENCE_UNCERTAIN'), @('V06', 'Flags', 'INT', 'QC_Flags', 'VEC-V06|1|I1', 'flags', 'W_SEQUENCE_UNCERTAIN'),
    @('V07', 'Selected anchor (ISO 8601)', 'INT', 'Anchor_ISO', 'VEC-V07|1|I1', 'text', '2026-03-02T08:00-05:00'), @('V07', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V07|1|I1', 'num', 720), @('V07', 'Unit', 'INT', 'Unit', 'VEC-V07|1|I1', 'text', 'minutes'), @('V07', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V07|1|I1', 'text', 'AFTER'), @('V07', 'Flags', 'INT', 'QC_Flags', 'VEC-V07|1|I1', 'flags', ''),
    @('V08', 'FIRST_ADMIN interval', 'INT', 'Interval_Value', 'VEC-V08|1|I1', 'num', 44), @('V08', 'RESTART interval', 'INT', 'Interval_Value', 'VEC-V08|1|I2', 'num', 4), @('V08', 'Unit (FIRST_ADMIN)', 'INT', 'Unit', 'VEC-V08|1|I1', 'text', 'calendar days'), @('V08', 'Flags (FIRST_ADMIN)', 'INT', 'QC_Flags', 'VEC-V08|1|I1', 'flags', ''), @('V08', 'Flags (RESTART)', 'INT', 'QC_Flags', 'VEC-V08|1|I2', 'flags', ''),
    @('V09', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V09|1|I1', 'num', 60), @('V09', 'Unit', 'INT', 'Unit', 'VEC-V09|1|I1', 'text', 'minutes'), @('V09', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V09|1|I1', 'text', 'AFTER'), @('V09', 'Flags', 'INT', 'QC_Flags', 'VEC-V09|1|I1', 'flags', ''),
    @('V10', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V10|1|I1', 'num', 120), @('V10', 'Unit', 'INT', 'Unit', 'VEC-V10|1|I1', 'text', 'minutes (local clock)'), @('V10', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V10|1|I1', 'text', 'AFTER'), @('V10', 'Flags', 'INT', 'QC_Flags', 'VEC-V10|1|I1', 'flags', 'W_TZ_UNKNOWN'),
    @('V11', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V11|1|I1', 'num', 360), @('V11', 'Unit', 'INT', 'Unit', 'VEC-V11|1|I1', 'text', 'minutes'), @('V11', 'Reported timing retained', 'INT', 'Reported_Timing_Result', 'VEC-V11|1|I1', 'contains', 'two hours after the dose'), @('V11', 'Flags', 'INT', 'QC_Flags', 'VEC-V11|1|I1', 'flags', 'W_REPORTED_TIMING_CONFLICT'),
    @('V12', 'Version 1 kept (inspection aid)', 'INT', 'Interval_Value', 'VEC-V12|1|I1', 'num', 9), @('V12', 'Version 2 added (inspection aid)', 'INT', 'Interval_Value', 'VEC-V12|2|I1', 'num', 7),
    @('V13', '30 February', 'INT', 'QC_Flags', 'VEC-V13|1|I-DATE', 'flags', 'E_INVALID_DATE'), @('V13', 'Reversed bounds', 'INT', 'QC_Flags', 'VEC-V13|1|I-BOUNDS', 'flags', 'E_REVERSED_BOUNDS'), @('V13', 'Duplicate episode ID EX1', 'INT', 'QC_Flags', 'VEC-V13|1|I-IDS', 'flags', 'E_DUPLICATE_ID'), @('V13', 'Unlinked pair', 'INT', 'QC_Flags', 'VEC-V13|1|I-LINK', 'flags', 'E_MISSING_LINK'),
    @('V15', 'No value', 'INT', 'Interval_Value', 'VEC-V15|1|I1', 'blank', ''), @('V15', 'Reported timing flag (C6)', 'INT', 'C6_Code', 'VEC-V15|1|I1', 'text', 'I_APPROXIMATE_NOT_COMPARED'),
    @('V16', 'Fraction remaining', 'Q1', 'Fraction_Remaining', 'VEC-V16|1|Q1-V16', 'num', 0.125), @('V16', 'k per hour', 'Q1', 'k_per_hour', 'VEC-V16|1|Q1-V16', 'num', 0.028881133), @('V16', 'Flags', 'Q1', 'QC_Flags', 'VEC-V16|1|Q1-V16', 'flags', ''),
    @('V17', 'Time to fraction (hours)', 'Q2', 'Time_to_f', 'VEC-V17|1|Q2-V17', 'num', 79.726274), @('V17', 'Half-lives', 'Q2', 'Half_Lives', 'VEC-V17|1|Q2-V17', 'num', 3.321928), @('V17', 'Flags', 'Q2', 'QC_Flags', 'VEC-V17|1|Q2-V17', 'flags', ''),
    @('V18', 'Accumulation factor', 'Q3', 'Accumulation_Factor', 'VEC-V18|1|Q3-V18', 'num', 3.414214), @('V18', 'Flags', 'Q3', 'QC_Flags', 'VEC-V18|1|Q3-V18', 'flags', ''),
    @('V19', 'Half-life ratio', 'Q4', 'HalfLife_Ratio', 'VEC-V19|1|Q4-V19', 'num', 2), @('V19', 'Average Css ratio', 'Q4', 'Css_Ratio', 'VEC-V19|1|Q4-V19', 'num', 2), @('V19', 'Flags', 'Q4', 'QC_Flags', 'VEC-V19|1|Q4-V19', 'flags', ''),
    @('V20', 'Parent percent remaining', 'Q5', 'Parent_Percent', 'VEC-V20|1|Q5-V20', 'num', 0.390625), @('V20', 'Metabolite percent remaining', 'Q5', 'Metabolite_Percent', 'VEC-V20|1|Q5-V20', 'num', 62.996052), @('V20', 'Flags', 'Q5', 'QC_Flags', 'VEC-V20|1|Q5-V20', 'flags', ''),
    @('V21', 'Fraction replaced', 'Q6', 'Fraction_Q6', 'VEC-V21|1|Q6-V21', 'num', 0.75), @('V21', 'Flags', 'Q6', 'QC_Flags', 'VEC-V21|1|Q6-V21', 'flags', 'I_TEACHING_DEMONSTRATION'),
    @('V22', 'Fraction remaining', 'Q7', 'Fraction_Remaining', 'VEC-V22|1|Q7-V22', 'num', 0.25), @('V22', 'Flags', 'Q7', 'QC_Flags', 'VEC-V22|1|Q7-V22', 'flags', 'W_BIOLOGIC_MODEL_CAUTION'),
    @('V23', 'Fraction at lower t_half', 'Q8', 'Fraction_at_Lower', 'VEC-V23|1|Q8-V23', 'num', 0.0625), @('V23', 'Fraction at upper t_half', 'Q8', 'Fraction_at_Upper', 'VEC-V23|1|Q8-V23', 'num', 0.25), @('V23', 'Flags', 'Q8', 'QC_Flags', 'VEC-V23|1|Q8-V23', 'flags', 'I_SCENARIO_NOT_CI'),
    @('V24', 'Half-life 0', 'Q1', 'QC_Flags', 'VEC-V24|1|Q1-V24a', 'flags', 'E_NONPOSITIVE_PARAMETER'), @('V24', 'Unit fortnights', 'Q1', 'QC_Flags', 'VEC-V24|1|Q1-V24b', 'flags', 'E_INVALID_UNIT'), @('V24', 'Gate unconfirmed', 'Q1', 'QC_Flags', 'VEC-V24|1|Q1-V24c', 'flags', 'E_APPLICABILITY_UNCONFIRMED'), @('V24', 'Parameter status travels into Parameter_Set (V16 row)', 'Q1', 'Parameter_Set', 'VEC-V16|1|Q1-V16', 'contains', 'hypothetical example'),
    @('V25', 'Amount infused (mg)', 'Q9', 'Amount_mg', 'VEC-V25|1|Q9-V25', 'num', 58.333333), @('V25', 'Percent of planned dose', 'Q9', 'Percent_of_Planned', 'VEC-V25|1|Q9-V25', 'num', 11.666667), @('V25', 'Flags', 'Q9', 'QC_Flags', 'VEC-V25|1|Q9-V25', 'flags', ''),
    @('V26', 'Method A (mg/m2)', 'Q10', 'Method_A_mg_m2', 'VEC-V26|1|Q10-V26-first', 'num', 164.052288), @('V26', 'Method B, first BSA', 'Q10', 'Method_B_mg_m2', 'VEC-V26|1|Q10-V26-first', 'num', 161.111111), @('V26', 'Method B, last BSA', 'Q10', 'Method_B_mg_m2', 'VEC-V26|1|Q10-V26-last', 'num', 170.588235), @('V26', 'Flags', 'Q10', 'QC_Flags', 'VEC-V26|1|Q10-V26-first', 'flags', ''),
    @('V27', 'Case 1: male, 1.0 mg/dL', 'Q11', 'CrCl_mL_min', 'VEC-V27|1|Q11-V27a', 'num', 68.055556), @('V27', 'Case 2: male, 2.0 mg/dL', 'Q11', 'CrCl_mL_min', 'VEC-V27|1|Q11-V27b', 'num', 34.027778), @('V27', 'Case 3: female, 1.0 mg/dL', 'Q11', 'CrCl_mL_min', 'VEC-V27|1|Q11-V27c', 'num', 57.847222),
    @('V27', 'Case 4: sex not recorded, female estimate', 'Q11', 'CrCl_Female_If_Sex_Not_Recorded', 'VEC-V27|1|Q11-V27d', 'num', 57.847222), @('V27', 'Case 4: sex not recorded, male estimate', 'Q11', 'CrCl_Male_If_Sex_Not_Recorded', 'VEC-V27|1|Q11-V27d', 'num', 68.055556),
    @('V27', 'Case 5: no estimate', 'Q11', 'CrCl_mL_min', 'VEC-V27|1|Q11-V27e', 'blank', ''),
    @('V27', 'Flags case 1', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27a', 'flags', ''), @('V27', 'Flags case 2', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27b', 'flags', ''), @('V27', 'Flags case 3', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27c', 'flags', ''), @('V27', 'Flags case 4', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27d', 'flags', 'W_SEX_NOT_RECORDED'), @('V27', 'Flags case 5', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27e', 'flags', 'E_CREATININE_STABILITY_UNCONFIRMED'),
    @('V28', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V28|1|I1', 'num', 19), @('V28', 'Unit', 'INT', 'Unit', 'VEC-V28|1|I1', 'text', 'minutes'), @('V28', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V28|1|I1', 'text', 'AFTER'), @('V28', 'Interventions counted', 'INT', 'Interventions_Counted', 'VEC-V28|1|I1', 'text', 'MS-RX'), @('V28', 'Flags', 'INT', 'QC_Flags', 'VEC-V28|1|I1', 'flags', 'W_INTERVENTION_BEFORE_RECOVERY'), @('V28', 'Calculation_Mode', 'INT', 'Calculation_Mode_Used', 'VEC-V28|1|I1', 'text', 'C7 (C1)'),
    @('V29', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V29|1|I1', 'num', 3), @('V29', 'Unit', 'INT', 'Unit', 'VEC-V29|1|I1', 'text', 'calendar days'), @('V29', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V29|1|I1', 'text', 'AFTER'), @('V29', 'Interventions counted', 'INT', 'Interventions_Counted', 'VEC-V29|1|I1', 'text', 'MS-FLU'), @('V29', 'Flags', 'INT', 'QC_Flags', 'VEC-V29|1|I1', 'flags', 'W_INTERVENTION_BEFORE_RECOVERY'),
    @('V30', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V30|1|I1', 'num', 3), @('V30', 'Unit', 'INT', 'Unit', 'VEC-V30|1|I1', 'text', 'calendar days'), @('V30', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V30|1|I1', 'text', 'AFTER'), @('V30', 'Interventions counted (none)', 'INT', 'Interventions_Counted', 'VEC-V30|1|I1', 'blank', ''), @('V30', 'Flags', 'INT', 'QC_Flags', 'VEC-V30|1|I1', 'flags', ''),
    @('V31', 'Onset bounds (ISO 8601)', 'INT', 'End_ISO', 'VEC-V31|1|I1', 'text', '2026-03-15/2026-03-29'), @('V31', 'Interval_Minimum', 'INT', 'Interval_Minimum', 'VEC-V31|1|I1', 'num', 14), @('V31', 'Interval_Maximum', 'INT', 'Interval_Maximum', 'VEC-V31|1|I1', 'num', 28), @('V31', 'Unit', 'INT', 'Unit', 'VEC-V31|1|I1', 'text', 'calendar days'), @('V31', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V31|1|I1', 'text', 'AFTER'), @('V31', 'Flags', 'INT', 'QC_Flags', 'VEC-V31|1|I1', 'flags', ''),
    @('V32', 'Fraction of the change completed', 'Q6', 'Fraction_Q6', 'VEC-V32|1|Q6-V32', 'num', 0.875), @('V32', 'Output label (approach to a new level)', 'Q6', 'Output_Label', 'VEC-V32|1|Q6-V32', 'text', 'Fraction of the change to the new level completed'), @('V32', 'Flags', 'Q6', 'QC_Flags', 'VEC-V32|1|Q6-V32', 'flags', 'I_TEACHING_DEMONSTRATION'),
    @('V33', 'Selected anchor (ISO 8601)', 'INT', 'Anchor_ISO', 'VEC-V33|1|I1', 'text', '2026-03-02T08:00-05:00'), @('V33', 'Interval_Value', 'INT', 'Interval_Value', 'VEC-V33|1|I1', 'num', 720), @('V33', 'Unit', 'INT', 'Unit', 'VEC-V33|1|I1', 'text', 'minutes'), @('V33', 'Sequence', 'INT', 'Sequence_Status', 'VEC-V33|1|I1', 'text', 'AFTER'), @('V33', 'Flags', 'INT', 'QC_Flags', 'VEC-V33|1|I1', 'flags', ''),
    @('V34', 'No anchor selected', 'INT', 'Anchor_Record_ID_Used', 'VEC-V34|1|I1', 'blank', ''), @('V34', 'No value', 'INT', 'Interval_Value', 'VEC-V34|1|I1', 'blank', ''), @('V34', 'Flags', 'INT', 'QC_Flags', 'VEC-V34|1|I1', 'flags', 'E_MISSING_INPUT'),
    @('V35', 'Technical case: every E-code', 'Q1', 'QC_Flags', 'VEC-V35|1|Q1-V35a', 'flags', 'E_INVALID_UNIT; E_NONPOSITIVE_PARAMETER; E_APPLICABILITY_UNCONFIRMED'), @('V35', 'Interval case: every E-code', 'INT', 'QC_Flags', 'VEC-V35|1|I1', 'flags', 'E_INVALID_DATE; E_MISSING_LINK'),
    @('V36', 'Case 1: no value', 'Q1', 'Fraction_Remaining', 'VEC-V36|1|Q1-V36a', 'blank', ''), @('V36', 'Case 1: flags', 'Q1', 'QC_Flags', 'VEC-V36|1|Q1-V36a', 'flags', 'E_MISSING_INPUT; E_APPLICABILITY_UNCONFIRMED'),
    @('V36', 'Case 2: fraction remaining', 'Q1', 'Fraction_Remaining', 'VEC-V36|1|Q1-V36b', 'num', 0.125), @('V36', 'Case 2: flags', 'Q1', 'QC_Flags', 'VEC-V36|1|Q1-V36b', 'flags', ''),
    @('V36', 'Case 3: no value', 'Q1', 'Fraction_Remaining', 'VEC-V36|1|Q1-V36c', 'blank', ''), @('V36', 'Case 3: flags', 'Q1', 'QC_Flags', 'VEC-V36|1|Q1-V36c', 'flags', 'E_APPLICABILITY_UNCONFIRMED'),
    @('V37', 'Fraction at lower t_half', 'Q8', 'Fraction_at_Lower', 'VEC-V37|1|Q8-V37', 'num', 0.0625), @('V37', 'Fraction at upper t_half', 'Q8', 'Fraction_at_Upper', 'VEC-V37|1|Q8-V37', 'num', 0.25), @('V37', 'Flags', 'Q8', 'QC_Flags', 'VEC-V37|1|Q8-V37', 'flags', 'W_BIOLOGIC_MODEL_CAUTION; I_SCENARIO_NOT_CI'),
    @('V38', 'Method A (mg/m2)', 'Q10', 'Method_A_mg_m2', 'VEC-V38|1|Q10-V38-first', 'num', 164.052288), @('V38', 'Method B, first BSA by date', 'Q10', 'Method_B_mg_m2', 'VEC-V38|1|Q10-V38-first', 'num', 161.111111), @('V38', 'Method B, last BSA by date', 'Q10', 'Method_B_mg_m2', 'VEC-V38|1|Q10-V38-last', 'num', 170.588235), @('V38', 'Flags (first)', 'Q10', 'QC_Flags', 'VEC-V38|1|Q10-V38-first', 'flags', ''), @('V38', 'Flags (last)', 'Q10', 'QC_Flags', 'VEC-V38|1|Q10-V38-last', 'flags', '')
)

function LookupF([string]$t, [string]$col, [string]$key) { '=IFERROR(INDEX(n' + $t + '_' + $col + ',MATCH("' + $key + '",n' + $t + '_Key,0)),"(row not found)")' }

# ================================================================== BUILD
$SheetOrder = @('Read_Me', 'Cases', 'Exposures', 'Administrations', 'Events', 'Milestones', 'Drug_Event_Links', 'Intervals', 'Product_Dossier', 'Parameter_Register', 'PK_Exercises', 'PD_Profile', 'Temporal_Profile', 'Biology_Review', 'Dechallenge_Rechallenge', 'Alternatives_Followup', 'Assessment_QC', 'Summary', 'Export', 'Examples', 'Verification', 'Timepoints', 'Flags_Messages', 'Lists')
$t0 = Get-Date
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.ScreenUpdating = $false; $xl.EnableEvents = $false
    $xl.SheetsInNewWorkbook = 1
    $wb = $xl.Workbooks.Add()
    $xl.Calculation = -4135   # manual during build
    $S = @{}
    for ($i = 0; $i -lt $SheetOrder.Count; $i++) {
        if ($i -eq 0) { $ws = $wb.Worksheets.Item(1) } else { $ws = $wb.Worksheets.Add($Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
        $ws.Name = $SheetOrder[$i]; $ws.Cells.Font.Name = 'Calibri'; $ws.Cells.Font.Size = 10
        $S[$SheetOrder[$i]] = $ws
    }
    Write-Host "Sheets created"

    # ---------- Lists
    $ws = $S['Lists']; $ci = 1
    $ws.Range('A1').Value2 = 'Lists used by data validation (locked)'; $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Size = 13
    $ws.Range('A2').Value2 = 'Missing-data reason codes are the example codes proposed in Chapter 23 (not part of CALC-SPEC 1.1). Timing-statement kinds follow change register B14-13. OnsetPrecision adds LAB, the laboratory-defined onset of CALC-SPEC 1.1 C4.'
    foreach ($k in $Lists.Keys) {
        $L = ColL $ci; $vals = $Lists[$k]
        $ws.Range("$L" + '4').Value2 = $k; $ws.Range("$L" + '4').Font.Bold = $true
        $ws.Range("$L" + "5:$L" + (4 + $vals.Count)).NumberFormat = '@'
        for ($j = 0; $j -lt $vals.Count; $j++) { $ws.Range("$L" + (5 + $j)).Value2 = [string]$vals[$j] }
        [void]$wb.Names.Add("nLIST_$k", "='Lists'!`$$L`$5:`$$L`$" + (4 + $vals.Count))
        $ws.Columns($ci).ColumnWidth = [double][math]::Max(12, ($vals | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum + 2)
        $ci++
    }
    # ---------- Flags_Messages
    $ws = $S['Flags_Messages']
    $ws.Range('A1').Value2 = 'Flags and user messages (CALC-SPEC 1.1, Section 5; shown verbatim, [input] filled in; QC_Flags list every code in this order, rule G9)'; $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Size = 13
    $ws.Range('A2').Value2 = 'E-codes block a result (no number is shown). W-codes accompany a result and must be displayed with it. I-codes annotate it.'
    $hdr = @('Code', 'Type', 'Raised when', 'Message shown')
    for ($j = 0; $j -lt 4; $j++) { $c = $ws.Cells.Item(4, $j + 1); $c.Value2 = $hdr[$j]; $c.Font.Bold = $true; $c.Interior.Color = $C_FHdr; $c.Font.Color = RGBc 255 255 255 }
    for ($i = 0; $i -lt $Flags.Count; $i++) { for ($j = 0; $j -lt 4; $j++) { $ws.Cells.Item(5 + $i, $j + 1).Value2 = $Flags[$i][$j] } }
    $fl = 4 + $Flags.Count
    [void]$wb.Names.Add('nMSG_Code', "='Flags_Messages'!`$A`$5:`$A`$$fl"); [void]$wb.Names.Add('nMSG_Text', "='Flags_Messages'!`$D`$5:`$D`$$fl")
    $ws.Columns(1).ColumnWidth = [double]36; $ws.Columns(2).ColumnWidth = [double]9; $ws.Columns(3).ColumnWidth = [double]48; $ws.Columns(4).ColumnWidth = [double]95
    Write-Host "Lists and messages done"

    # ---------- record sheets
    $note = 'Yellow cells are inputs; grey cells are locked formulas. Enter what the source says: original text, precision and bounds. Nothing is imputed (CALC-SPEC G2).'
    $iCase = Layout-Table $S['Cases'] 1 'Cases: one row per case version' $note $CaseCols 20 $CasesEx $CasesVec 'CASE' $EX_LABEL $VEC_LABEL
    Def-Names $iCase 'nCASE_'
    $tpHelp = @{ S_Prec = 'TIMESTAMP: Date + Time + UTC offset (choose unknown if not recorded). DATE: Date only. MONTH: Year + Month. YEAR: Year. RANGE: earliest and latest dates as stated, with the source.' }
    $iExp = Layout-Table $S['Exposures'] 1 'Exposures: one row per exposure episode (Start, Stop and Restart are timepoints)' $note $ExpCols 30 $ExpEx $ExpVec 'EXP' $EX_LABEL $VEC_LABEL $tpHelp
    Def-Names $iExp 'nEXP_'
    $iAdm = Layout-Table $S['Administrations'] 1 'Administrations: one row per dose, administration or infusion rate segment' $note $AdmCols 60 $AdmEx $AdmVec 'ADM' $EX_LABEL $VEC_LABEL @{ A_Prec = $tpHelp.S_Prec; Kind = 'INFUSION_SEGMENT rows need a start timestamp, an end timestamp and a rate in mg/hour for Q9.' }
    Def-Names $iAdm 'nADM_'
    $iEvt = Layout-Table $S['Events'] 1 'Events: one row per event episode (Onset and First detection are timepoints; a laboratory-defined onset is precision LAB with two test milestones)' $note $EvtCols 15 $EvtEx $EvtVec 'EVT' $EX_LABEL $VEC_LABEL @{ O_Prec = ($tpHelp.S_Prec + ' LAB: laboratory-defined onset (CALC-SPEC 1.1 C4); enter the Milestone_IDs of the last normal and first abnormal tests.') }
    Def-Names $iEvt 'nEVT_'
    $iMs = Layout-Table $S['Milestones'] 1 'Milestones: tests, organ function, interventions, recovery and other dated records' $note $MsCols 40 $MsEx $MsVec 'MS' $EX_LABEL $VEC_LABEL @{ M_Prec = $tpHelp.S_Prec }
    Def-Names $iMs 'nMS_'
    $iLnk = Layout-Table $S['Drug_Event_Links'] 1 'Drug_Event_Links: pairs selected by the reviewer (an interval needs a link: E_MISSING_LINK)' $note $LnkCols 20 $LnkEx $LnkVec 'LNK' $EX_LABEL $VEC_LABEL
    Def-Names $iLnk 'nLNK_'
    Write-Host "Record sheets laid out"

    # ---------- Timepoints (derived index of every dated record)
    $wsT = $S['Timepoints']
    $tpCols = @('Key', 'Record_ID', 'Record_Type', 'Case_ID', 'Case_Version', 'Case_Key', 'Episode_ID', 'Precision', 'Offset', 'Err', 'HT', 'OK', 'Loc', 'UTC', 'Lo', 'Hi', 'ISO', 'Display', 'Original_Text', 'Source', 'Source_Location', 'Dup', 'Loc_n', 'UTC_n', 'Lo_n', 'Hi_n', 'RowN', 'Kind', 'Role', 'Label', 'Cand', 'CandKey', 'Exact', 'Lab_Note')
    $tpMap = @{}; for ($i = 0; $i -lt $tpCols.Count; $i++) { $tpMap[$tpCols[$i]] = ColL ($i + 1) }
    $segs = @(
        @{ info = $iExp; p = 'S'; id = 'Exposure_Episode_ID'; suf = '/START'; type = 'EXPOSURE_START'; ep = 'Exposure_Episode_ID'; label = 'Start' },
        @{ info = $iExp; p = 'T'; id = 'Exposure_Episode_ID'; suf = '/STOP'; type = 'EXPOSURE_STOP'; ep = 'Exposure_Episode_ID'; label = 'Stop' },
        @{ info = $iExp; p = 'R'; id = 'Exposure_Episode_ID'; suf = '/RESTART'; type = 'EXPOSURE_RESTART'; ep = 'Exposure_Episode_ID'; label = 'Restart' },
        @{ info = $iAdm; p = 'A'; id = 'Admin_ID'; suf = ''; type = 'ADMINISTRATION'; ep = 'Exposure_Episode_ID'; label = 'Administration' },
        @{ info = $iEvt; p = 'O'; id = 'Episode_ID'; suf = '/ONSET'; type = 'EVENT_ONSET'; ep = 'Episode_ID'; label = 'Onset' },
        @{ info = $iEvt; p = 'D'; id = 'Episode_ID'; suf = '/DETECTION'; type = 'EVENT_DETECTION'; ep = 'Episode_ID'; label = 'First detection' },
        @{ info = $iMs; p = 'M'; id = 'Milestone_ID'; suf = ''; type = 'MILESTONE'; ep = 'Related_Episode_ID'; label = 'Milestone' }
    )
    $r = 5
    foreach ($sg in $segs) { $sg.s = $r; $sg.e = $r + ($sg.info.Last - $sg.info.First); $r = $sg.e + 1 }
    $tpLast = $r - 1
    $wsT.Range('A1').Value2 = 'Timepoints: derived index of every dated record (locked). Use the Record_ID column as Anchor_Record_ID or End_Record_ID on Intervals.'
    $wsT.Range('A1').Font.Bold = $true; $wsT.Range('A1').Font.Size = 13
    $wsT.Range('A2').Value2 = 'Exposure timepoints are Episode/START, /STOP and /RESTART; event timepoints Episode/ONSET and /DETECTION; administrations and milestones use their own IDs. Loc and UTC are minutes since 1900 (integers), used only for arithmetic.'
    for ($i = 0; $i -lt $tpCols.Count; $i++) { $c = $wsT.Cells.Item(4, $i + 1); $c.Value2 = $tpCols[$i]; $c.Font.Bold = $true; $c.Interior.Color = $C_FHdr; $c.Font.Color = RGBc 255 255 255 }
    $wsT.Range("A5:" + (ColL $tpCols.Count) + $tpLast).Interior.Color = $C_F
    $wsT.Range("O5:P$tpLast").NumberFormat = 'd mmmm yyyy'
    foreach ($c in $tpCols) { [void]$wb.Names.Add("nTP_$c", "='Timepoints'!`$" + $tpMap[$c] + "`$5:`$" + $tpMap[$c] + "`$$tpLast") }
    Write-Host "Timepoints laid out ($tpLast rows)"

    # ---------- Intervals
    $iInt = Layout-Table $S['Intervals'] 1 'Intervals: one row per selected link and anchor (C1-C7)' 'Choose the link, the anchor and the rule; the row shows value, unit, sequence status and flags. Blocked rows show the E-code and no number. The workbook never picks the shortest interval and never assigns causality.' $IntCols 40 $IntEx $IntVec 'INT' $EX_LABEL $VEC_LABEL
    Def-Names $iInt 'nINT_'

    # ---------- PK_Exercises (stacked tables)
    $wsP = $S['PK_Exercises']
    $wsP.Range('A1').Value2 = 'PK_Exercises: gated basic PK arithmetic (Q1-Q5, Q7-Q11). Run only what the case question needs.'
    $wsP.Range('A1').Font.Bold = $true; $wsP.Range('A1').Font.Size = 14; $wsP.Range('A1').Font.Color = $C_Title
    $wsP.Range('A2').Value2 = 'Every result is gated, carries its Parameter_Set with each input''s status, and shows the CALC-SPEC mandatory text. These results describe exposure under a stated model; they never decide causality and are not dosing advice.'
    $qdefs = @(
        @{ t = 'Q1'; title = 'Q1 Elimination and fractional persistence'; cols = $Q1Cols; ex = @(); vec = $Q1Vec },
        @{ t = 'Q2'; title = 'Q2 Time to approach steady state'; cols = $Q2Cols; ex = @(); vec = $Q2Vec },
        @{ t = 'Q3'; title = 'Q3 Accumulation during regular dosing'; cols = $Q3Cols; ex = @(); vec = $Q3Vec },
        @{ t = 'Q4'; title = 'Q4 A change in clearance'; cols = $Q4Cols; ex = $Q4Ex; vec = $Q4Vec },
        @{ t = 'Q5'; title = 'Q5 Parent and metabolite persistence'; cols = $Q5Cols; ex = @(); vec = $Q5Vec },
        @{ t = 'Q7'; title = 'Q7 Long persistence of a monoclonal antibody (always W_BIOLOGIC_MODEL_CAUTION)'; cols = $Q7Cols; ex = $Q7Ex; vec = $Q7Vec },
        @{ t = 'Q8'; title = 'Q8 Parameter uncertainty and scenarios (a range, never a midpoint or confidence interval)'; cols = $Q8Cols; ex = $Q8Ex; vec = $Q8Vec },
        @{ t = 'Q9'; title = 'Q9 Dose received before onset during an infusion (rate segments from Administrations)'; cols = $Q9Cols; ex = $Q9Ex; vec = $Q9Vec },
        @{ t = 'Q10'; title = 'Q10 Cumulative dose normalized to body surface area (doses and BSA from Administrations)'; cols = $Q10Cols; ex = @(); vec = $Q10Vec },
        @{ t = 'Q11'; title = 'Q11 Estimated renal function (Cockcroft-Gault) with the creatinine-stability gate'; cols = $Q11Cols; ex = $Q11Ex; vec = $Q11Vec }
    )
    $row = 4; $pkInfos = @()
    foreach ($q in $qdefs) {
        $ni = Layout-Table $wsP $row $q.title 'Mandatory text and formula are shown on every result row. Common gate: confirm the model, the parameter used, the reference time and each input''s Parameter_Status (Q9, Q10 and Q11 have their own checks).' $q.cols 8 $q.ex $q.vec $q.t $EX_LABEL $VEC_LABEL
        Def-Names $ni ('n' + $q.t + '_')
        $pkInfos += $ni; $row = $ni.Last + 3
    }
    Set-Widths $wsP $pkInfos

    # ---------- PD_Profile
    $wsD = $S['PD_Profile']
    $iPDP = Layout-Table $wsD 1 'PD_Profile: structured prompts answered in words, with sources (no formula decides biology)' 'Prompts follow Stage 1 document B. Q6 below is a teaching demonstration only.' $PDPCols 15 $PDPEx @() 'PDP' $EX_LABEL $VEC_LABEL
    $iQ6 = Layout-Table $wsD ($iPDP.Last + 3) 'Q6 Biological replacement and recovery: TEACHING DEMONSTRATION (always I_TEACHING_DEMONSTRATION)' 'Shown only with the teaching-demonstration switch on. Hypothetical turnover illustration; not a measured recovery curve for platelets or any other clinical target.' $Q6Cols 8 @() $Q6Vec 'Q6' $EX_LABEL $VEC_LABEL
    Def-Names $iQ6 'nQ6_'
    Set-Widths $wsD @($iPDP, $iQ6)

    # ---------- reviewer and dossier sheets
    $rnote = 'Reviewer layer: judgments in words, each with a rationale and sources. No formula writes to these columns (CALC-SPEC G8).'
    $iPDS = Layout-Table $S['Product_Dossier'] 1 'Product_Dossier: one row per extracted statement (versioned; US and EU kept side by side)' 'Record document, version, revision, retrieval date and section for every statement.' $PDSCols 30 $PDSEx @() 'PDS' $EX_LABEL $VEC_LABEL
    $iPR = Layout-Table $S['Parameter_Register'] 1 'Parameter_Register: one row per parameter, with status, population and source' 'A value without a population and a half-life type is not ready for use.' $PRCols 25 $PREx @() 'PR' $EX_LABEL $VEC_LABEL
    $iTPR = Layout-Table $S['Temporal_Profile'] 1 'Temporal_Profile: expected course versus observed course, per drug-event pair' $rnote $TPRCols 15 $TPREx @() 'TPR' $EX_LABEL $VEC_LABEL
    $iBIO = Layout-Table $S['Biology_Review'] 1 'Biology_Review: the reviewer''s judgment of biological fit (four options; none is a relatedness category)' $rnote $BIOCols 15 $BIOEx @() 'BIO' $EX_LABEL $VEC_LABEL
    $iDCR = Layout-Table $S['Dechallenge_Rechallenge'] 1 'Dechallenge_Rechallenge: actions, responses and every intervention in between' $rnote $DCRCols 15 $DCREx @() 'DCR' $EX_LABEL $VEC_LABEL
    $iALT = Layout-Table $S['Alternatives_Followup'] 1 'Alternatives_Followup: competing explanations and targeted queries' $rnote $ALTCols 30 $ALTEx @() 'ALT' $EX_LABEL $VEC_LABEL
    $iAQC = Layout-Table $S['Assessment_QC'] 1 'Assessment_QC: narrative, method, unresolved issues (the formal causality assessment remains a human assessment)' $rnote $AQCCols 15 $AQCEx @() 'AQC' $EX_LABEL $VEC_LABEL
    Write-Host "All tables laid out"

    # ================================================================== PHASE 2: FORMULAS
    foreach ($inf in @($iCase, $iExp, $iAdm, $iEvt, $iMs, $iLnk)) { Write-Formulas $inf; Write-Host "  formulas: $($inf.Tid)" }
    # Timepoints formulas
    foreach ($sg in $segs) {
        $si = $sg.info; $sm = $si.Map; $p = $sg.p; $r0 = $si.First; $sn = $si.Sheet
        $ref = { param($col) "'" + $sn + "'!`$" + $sm[$col] + $r0 }
        $idx = & $ref $sg.id
        $cnt = "COUNTA('" + $sn + "'!`$" + $sm["${p}_Prec"] + $r0 + ":`$" + $sm["${p}_Upper"] + $r0 + ")"
        $fx = [ordered]@{}
        $fx['Record_ID'] = '=IF(AND(' + $idx + '<>"",' + $cnt + '>0),' + $idx + '&"' + $sg.suf + '","")'
        if ($sg.type -eq 'MILESTONE') { $fx['Record_Type'] = '=IF({Record_ID}="","","MILESTONE:"&' + (& $ref 'Type') + ')' } else { $fx['Record_Type'] = '=IF({Record_ID}="","","' + $sg.type + '")' }
        $fx['Case_ID'] = '=IF({Record_ID}="","",' + (& $ref 'Case_ID') + '&"")'
        $fx['Case_Version'] = '=IF({Record_ID}="","",' + (& $ref 'Case_Version') + '&"")'
        $fx['Episode_ID'] = '=IF({Record_ID}="","",' + (& $ref $sg.ep) + '&"")'
        $fx['Precision'] = '=IF({Record_ID}="","",' + (& $ref "${p}_Prec") + '&"")'
        $fx['Offset'] = '=IF({Record_ID}="","",' + (& $ref "${p}_Off") + '&"")'
        $fx['Err'] = '=IF({Record_ID}="","",' + (& $ref "${p}_Err") + '&"")'
        $fx['HT'] = '=IF({Record_ID}="",0,' + (& $ref "${p}_HT") + ')'
        $fx['OK'] = '=IF({Record_ID}="",0,' + (& $ref "${p}_OK") + ')'
        foreach ($z in @('Loc', 'UTC', 'Lo', 'Hi')) { $fx[$z] = '=IF({Record_ID}="","",' + (& $ref "${p}_$z") + ')' }
        $fx['Original_Text'] = '=IF({Record_ID}="","",' + (& $ref "${p}_Text") + '&"")'
        $fx['Source'] = '=IF({Record_ID}="","",' + (& $ref "${p}_Src") + '&"")'
        $fx['Source_Location'] = '=IF({Record_ID}="","","' + $sn + ' row "&ROW(' + (& $ref 'Case_ID') + ')&", ' + $sg.label + '")'
        $fx['Kind'] = $(if ($sg.type -eq 'ADMINISTRATION') { '=IF({Record_ID}="","",' + (& $ref 'Kind') + '&"")' } else { '=""' })
        if ($sg.info.Tid -eq 'EXP') { $fx['Role'] = '=IF({Record_ID}="","",' + (& $ref 'Role') + '&"")' }
        elseif ($sg.type -eq 'ADMINISTRATION') { $fx['Role'] = '=IF({Record_ID}="","",IFERROR(INDEX(nEXP_Role,MATCH(' + (& $ref 'Case_ID') + '&"|"&' + (& $ref 'Case_Version') + '&"|"&' + (& $ref 'Exposure_Episode_ID') + ',nEXP_Key,0))&"",""))' }
        else { $fx['Role'] = '=""' }
        switch ($sg.type) {
            'MILESTONE' { $fx['Label'] = '=IF({Record_ID}="","",' + (& $ref 'Description') + '&"")' }
            'ADMINISTRATION' { $fx['Label'] = '=IF({Record_ID}="","",' + (& $ref 'Kind') + '&"")' }
            'EVENT_ONSET' { $fx['Label'] = '=IF({Record_ID}="","",' + (& $ref 'Event_Term') + '&" (onset)")' }
            'EVENT_DETECTION' { $fx['Label'] = '=IF({Record_ID}="","",' + (& $ref 'Event_Term') + '&" (first detection)")' }
            default { $fx['Label'] = '=IF({Record_ID}="","",' + (& $ref 'Product') + '&": ' + $sg.label + '")' }
        }
        if ($sg.type -eq 'EVENT_ONSET') { $fx['Lab_Note'] = '=IF({Record_ID}="","",IF(' + (& $ref 'O_Prec') + '="LAB","Onset is laboratory-defined (CALC-SPEC 1.1, C4): lower bound = date of the last normal test ("&' + (& $ref 'O_LabNormal') + '&"), upper bound = date of the first abnormal test ("&' + (& $ref 'O_LabAbnormal') + '&"). The reviewer decides which results are normal.",""))' }
        else { $fx['Lab_Note'] = '=""' }
        foreach ($k in $fx.Keys) {
            $f = Expand $fx[$k] $tpMap $sg.s
            try { $wsT.Range($tpMap[$k] + $sg.s + ':' + $tpMap[$k] + $sg.e).Formula = $f } catch { throw "Timepoints $k rejected: $f :: $($_.Exception.Message)" }
        }
    }
    $uni = [ordered]@{
        Key      = '=IF({Record_ID}="","",{Case_Key}&"|"&{Record_ID})'
        Case_Key = '=IF({Record_ID}="","",{Case_ID}&"|"&{Case_Version})'
        ISO      = ('=IF(OR({Key}="",{Err}<>""),"",IF({Precision}="TIMESTAMP",' + (ISOD 'INT({Loc}/1440)') + '&"T"&' + (HHMM '{Loc}') + '&IF({OK}=1,{Offset},""),IF({Precision}="DATE",' + (ISOD '{Lo}') + ',IF({Precision}="MONTH",YEAR({Lo})&"-"&RIGHT("0"&MONTH({Lo}),2),IF({Precision}="YEAR",YEAR({Lo})&"",' + (ISOD '{Lo}') + '&"/"&' + (ISOD '{Hi}') + ')))))')
        Display  = ('=IF({Key}="","",IF({Err}<>"",{Err},IF({Precision}="TIMESTAMP",' + (DTXT 'INT({Loc}/1440)') + '&" "&' + (HHMM '{Loc}') + '&IF({OK}=1," (UTC"&{Offset}&")"," (offset unknown)"),IF({Precision}="DATE",' + (DTXT '{Lo}') + ',IF({Precision}="MONTH",INDEX(nLIST_Months,MONTH({Lo}))&" "&YEAR({Lo})&" ("&' + (DTXT '{Lo}') + '&" to "&' + (DTXT '{Hi}') + '&")",IF({Precision}="YEAR",YEAR({Lo})&" (1 January to 31 December)",' + (DTXT '{Lo}') + '&" to "&' + (DTXT '{Hi}') + '&IF({Precision}="LAB"," (laboratory-defined)","")))))))')
        Dup      = '=IF({Key}="",0,COUNTIF(nTP_Key,{Key}))'
        Loc_n    = '=IF({Loc}="",0,{Loc})'
        UTC_n    = '=IF({UTC}="",0,{UTC})'
        Lo_n     = '=IF({Lo}="",0,{Lo})'
        Hi_n     = '=IF({Hi}="",0,{Hi})'
        RowN     = '=ROW()-4'
        Cand     = '=IF(AND({Key}<>"",OR({Record_Type}="EXPOSURE_START",{Record_Type}="EXPOSURE_RESTART",AND({Record_Type}="ADMINISTRATION",{Kind}<>"INFUSION_SEGMENT"))),1,0)'
        CandKey  = '=IF({Cand}=1,{Case_Key}&"|"&{Episode_ID},"")'
        Exact    = '=IF(AND({Key}<>"",{Err}="",OR({Precision}="TIMESTAMP",{Precision}="DATE")),1,0)'
    }
    foreach ($k in $uni.Keys) {
        $f = Expand $uni[$k] $tpMap 5
        try { $wsT.Range($tpMap[$k] + '5:' + $tpMap[$k] + $tpLast).Formula = $f } catch { throw "Timepoints $k rejected: $f :: $($_.Exception.Message)" }
    }
    $tw = @{ Key = 26; Record_ID = 16; Record_Type = 22; Case_Key = 12; Display = 36; ISO = 24; Original_Text = 34; Source = 26; Source_Location = 26 }
    for ($i = 0; $i -lt $tpCols.Count; $i++) { $w = 9; if ($tw.ContainsKey($tpCols[$i])) { $w = $tw[$tpCols[$i]] }; $wsT.Columns($i + 1).ColumnWidth = [double]$w }
    Write-Host "  formulas: Timepoints"
    foreach ($inf in @($iInt) + $pkInfos + @($iPDP, $iQ6, $iPDS, $iPR, $iTPR, $iBIO, $iDCR, $iALT, $iAQC)) { Write-Formulas $inf; Write-Host "  formulas: $($inf.Tid)" }
    foreach ($inf in @($iCase, $iExp, $iAdm, $iEvt, $iMs, $iLnk, $iInt, $iPDS, $iPR, $iTPR, $iBIO, $iDCR, $iALT, $iAQC)) {
        for ($i = 0; $i -lt $inf.Cols.Count; $i++) { $inf.Ws.Columns($i + 1).ColumnWidth = [double]$inf.Cols[$i].w }
    }

    # ================================================================== EXAMPLES
    $wsE = $S['Examples']
    $wsE.Range('A1').Value2 = 'Examples: fictional teaching cases worked through the sheets'; $wsE.Range('A1').Font.Bold = $true; $wsE.Range('A1').Font.Size = 14; $wsE.Range('A1').Font.Color = $C_Title
    $wsE.Range('A2').Value2 = 'Cases EX-C01 (Chapter 3, Section 3.7, version 2), EX-C04 (Chapter 11, Section 11.8, version 2) and EX-C28 (Chapter 20, Section 20.4, version 2) are entered in the reserved blue rows of the input sheets with the home chapters'' dates, times and values. Each check below reads the workbook''s own result and compares it, rounded as the chapter states it, with the chapter''s number; the checks added in version 1.1 also follow the calculator''s example cases (docs/examples.js, tests/run_examples.js): EX-C28''s onset is laboratory-defined (CALC-SPEC 1.1 C4) and its Q8 rows are for a recorded monoclonal antibody. All cases and dates are fictional.'
    $wsE.Range('A2').WrapText = $true; $wsE.Range('A2:L2').Merge(); $wsE.Rows(2).RowHeight = [double]42
    $eh = @('#', 'Case', 'Home chapter', 'Item', 'Workbook row (sheet: key)', 'Chapter states', 'Check', 'Expected', 'Digits', 'Workbook value (unrounded)', 'Workbook display', 'Match')
    $ER0 = 6
    for ($j = 0; $j -lt $eh.Count; $j++) { $c = $wsE.Cells.Item($ER0, $j + 1); $c.Value2 = $eh[$j]; $c.Font.Bold = $true; $c.Interior.Color = $C_FHdr; $c.Font.Color = RGBc 255 255 255; $c.WrapText = $true }
    $sheetOf = @{ INT = 'Intervals'; Q4 = 'PK_Exercises (Q4)'; Q7 = 'PK_Exercises (Q7)'; Q8 = 'PK_Exercises (Q8)'; Q9 = 'PK_Exercises (Q9)'; Q10 = 'PK_Exercises (Q10)'; Q11 = 'PK_Exercises (Q11)' }
    for ($i = 0; $i -lt $ExChecks.Count; $i++) {
        $x = $ExChecks[$i]; $r = $ER0 + 1 + $i
        SetV $wsE.Cells.Item($r, 1) ($i + 1); $wsE.Cells.Item($r, 2).Value2 = $x[0]; $wsE.Cells.Item($r, 3).Value2 = $x[1]; $wsE.Cells.Item($r, 4).Value2 = $x[2]
        $wsE.Cells.Item($r, 5).Value2 = $sheetOf[$x[3]] + ': ' + $x[5] + ' / ' + $x[4]
        $wsE.Cells.Item($r, 6).NumberFormat = '@'; $wsE.Cells.Item($r, 6).Value2 = [string]$x[9]
        $wsE.Cells.Item($r, 7).Value2 = $x[6]
        if ($x[6] -in @('text', 'contains')) { $wsE.Cells.Item($r, 8).NumberFormat = '@' }
        SetV $wsE.Cells.Item($r, 8) $x[7]
        if ($null -ne $x[8]) { SetV $wsE.Cells.Item($r, 9) $x[8] }
        $wsE.Cells.Item($r, 10).Formula = LookupF $x[3] $x[4] $x[5]
        $wsE.Cells.Item($r, 11).Formula = LookupF $x[3] 'Result_Display' $x[5]
        switch ($x[6]) {
            'num' { $m = "=IF(ISNUMBER(J$r),ABS(ROUND(J$r,I$r)-H$r)<1E-9,FALSE)" }
            'text' { $m = "=EXACT(J$r&`"`",H$r&`"`")" }
            'contains' { $m = "=ISNUMBER(SEARCH(H$r,J$r))" }
            'lt' { $m = "=IF(ISNUMBER(J$r),J$r<H$r,FALSE)" }
        }
        $wsE.Cells.Item($r, 12).Formula = $m
    }
    $eLast = $ER0 + $ExChecks.Count
    [void]$wb.Names.Add('nEX_Match', "='Examples'!`$L`$$($ER0+1):`$L`$$eLast")
    $wsE.Range('A4').Value2 = 'Examples status:'; $wsE.Range('A4').Font.Bold = $true
    $wsE.Range('C4').Formula = '=IF(COUNTIF(nEX_Match,TRUE)=ROWS(nEX_Match),"MATCH: all "&ROWS(nEX_Match)&" example checks reproduce the home chapters'' numbers",COUNTIF(nEX_Match,TRUE)&" of "&ROWS(nEX_Match)&" example checks match; see FALSE rows")'
    $wsE.Range('C4').Font.Bold = $true; $wsE.Range('C4').Font.Size = 12
    [void]$wb.Names.Add('nEX_Status', "='Examples'!`$C`$4")
    $eRow = $eLast + 2
    $wsE.Cells.Item($eRow, 1).Value2 = 'Outside CALC-SPEC 1.1 (shown for Chapter 20 and 23 only): 24 calendar months after the last infusion, compared by calendar date (author''s convention; no CALC-SPEC rule adds calendar months).'
    $wsE.Cells.Item($eRow, 1).Font.Italic = $true
    $wsE.Cells.Item($eRow + 1, 4).Value2 = '24 calendar months after 3 March 2025'
    $wsE.Cells.Item($eRow + 1, 10).Formula = '=IFERROR(EDATE(INDEX(nTP_Lo,MATCH("EX-C28|2|ADM-RTX-6",nTP_Key,0)),24),"(row not found)")'
    $wsE.Cells.Item($eRow + 1, 10).NumberFormat = 'd mmmm yyyy'
    $wsE.Cells.Item($eRow + 1, 6).Value2 = '3 March 2027'
    $wsE.Cells.Item($eRow + 2, 4).Value2 = 'Calendar days from the last infusion to that date'
    $wsE.Cells.Item($eRow + 2, 10).Formula = "=IFERROR(J$($eRow+1)-INDEX(nTP_Lo,MATCH(`"EX-C28|2|ADM-RTX-6`",nTP_Key,0)),`"(row not found)`")"
    $wsE.Cells.Item($eRow + 2, 6).Value2 = '730 calendar days'
    $ew = @(4, 8, 22, 44, 34, 22, 8, 16, 6, 16, 50, 8); for ($j = 0; $j -lt $ew.Count; $j++) { $wsE.Columns($j + 1).ColumnWidth = [double]$ew[$j] }
    $wsE.Rows($ER0).RowHeight = [double]30
    $fc = $wsE.Range("L$($ER0+1):L$eLast").FormatConditions
    [void]$fc.Add(1, 3, '=TRUE'); $fc.Item(1).Interior.Color = $C_Pass
    [void]$fc.Add(1, 3, '=FALSE'); $fc.Item(2).Interior.Color = $C_Fail

    # ================================================================== VERIFICATION
    $wsV = $S['Verification']
    $wsV.Range('A1').Value2 = 'Verification: CALC-SPEC 1.1 vectors V01-V38 recomputed by this workbook'; $wsV.Range('A1').Font.Bold = $true; $wsV.Range('A1').Font.Size = 14; $wsV.Range('A1').Font.Color = $C_Title
    $wsV.Range('A2').Value2 = 'Each vector is entered as a reserved, locked case (VEC-Vnn) on the production sheets and computed by the same column formulas a user''s case uses. Expected values are the unrounded values of ' + $VecFile + ' (V01-V27 unchanged from CALC-SPEC 1.0; V28-V38 new in 1.1). Numbers are compared unrounded: relative tolerance 1E-6, or absolute 1E-9 when the expected value is zero; units, sequence status and flags must match exactly, and flag lists follow rule G9 (E-, then W-, then I-codes, in Section 5 order). V12 and V14 describe behavior and are checked by inspection.'
    $wsV.Range('A2:H2').Merge(); $wsV.Range('A2').WrapText = $true; $wsV.Rows(2).RowHeight = [double]54
    $wsV.Range('A4').Value2 = 'Overall status:'; $wsV.Range('A4').Font.Bold = $true; $wsV.Range('A4').Font.Size = 12
    $wsV.Range('A5').Value2 = 'Relative tolerance'; SetV $wsV.Range('B5') 1E-6; $wsV.Range('C5').Value2 = 'Absolute tolerance (expected = 0)'; SetV $wsV.Range('D5') 1E-9
    $wsV.Range('E5').Value2 = 'Source: ' + $VecFile + ' (spec CALC-SPEC 1.1, 2026-10-25)'
    [void]$wb.Names.Add('nTOL_REL', "='Verification'!`$B`$5"); [void]$wb.Names.Add('nTOL_ABS', "='Verification'!`$D`$5")
    $vh = @('ID', 'Rule', 'Where in this workbook', 'Type', 'Expected (JSON summary)', 'Checks', 'Passed', 'Result')
    $VR0 = 7
    for ($j = 0; $j -lt $vh.Count; $j++) { $c = $wsV.Cells.Item($VR0, $j + 1); $c.Value2 = $vh[$j]; $c.Font.Bold = $true; $c.Interior.Color = $C_FHdr; $c.Font.Color = RGBc 255 255 255 }
    $CR0 = $VR0 + $VerVectors.Count + 4
    $cLast = $CR0 + $VerChecks.Count
    for ($i = 0; $i -lt $VerVectors.Count; $i++) {
        $x = $VerVectors[$i]; $r = $VR0 + 1 + $i
        $ord = @(0, 1, 3, 2, 4); for ($j = 0; $j -lt 5; $j++) { $wsV.Cells.Item($r, $j + 1).Value2 = $x[$ord[$j]] }; $wsV.Cells.Item($r, 5).WrapText = $true; $wsV.Cells.Item($r, 3).WrapText = $true
        $wsV.Cells.Item($r, 6).Formula = "=COUNTIF(`$B`$$($CR0+1):`$B`$$cLast,A$r)"
        $wsV.Cells.Item($r, 7).Formula = "=COUNTIFS(`$B`$$($CR0+1):`$B`$$cLast,A$r,`$H`$$($CR0+1):`$H`$$cLast,TRUE)"
        $wsV.Cells.Item($r, 8).Formula = "=IF(D$r=`"behavioral`",`"INSPECT`",IF(AND(F$r>0,G$r=F$r),`"PASS`",`"FAIL`"))"
    }
    $vLast = $VR0 + $VerVectors.Count
    $wsV.Range("A$($VR0+1):H$vLast").VerticalAlignment = -4160
    [void]$wsV.Range("A$($VR0+1):H$vLast").Rows.AutoFit()
    $wsV.Range('C4').Formula = "=IF(COUNTIFS(`$D`$$($VR0+1):`$D`$$vLast,`"numeric`",`$H`$$($VR0+1):`$H`$$vLast,`"PASS`")=COUNTIF(`$D`$$($VR0+1):`$D`$$vLast,`"numeric`"),`"PASS: all `"&COUNTIF(`$D`$$($VR0+1):`$D`$$vLast,`"numeric`")&`" numeric vectors reproduce within the JSON tolerance`",`"FAIL: `"&COUNTIFS(`$D`$$($VR0+1):`$D`$$vLast,`"numeric`",`$H`$$($VR0+1):`$H`$$vLast,`"PASS`")&`" of `"&COUNTIF(`$D`$$($VR0+1):`$D`$$vLast,`"numeric`")&`" numeric vectors pass`")"
    $wsV.Range('C4').Font.Bold = $true; $wsV.Range('C4').Font.Size = 12
    [void]$wb.Names.Add('nVER_Overall', "='Verification'!`$C`$4")
    $wsV.Cells.Item($vLast + 2, 1).Value2 = 'Check details (one row per expected value in the JSON file)'; $wsV.Cells.Item($vLast + 2, 1).Font.Bold = $true
    $ch = @('#', 'Vector', 'What is checked', 'Kind', 'Workbook row (table: key / column)', 'Expected (unrounded, JSON)', 'Workbook value (unrounded)', 'Pass')
    for ($j = 0; $j -lt $ch.Count; $j++) { $c = $wsV.Cells.Item($CR0, $j + 1); $c.Value2 = $ch[$j]; $c.Font.Bold = $true; $c.Interior.Color = $C_FHdr; $c.Font.Color = RGBc 255 255 255 }
    $tabName = @{ INT = 'Intervals'; Q1 = 'Q1'; Q2 = 'Q2'; Q3 = 'Q3'; Q4 = 'Q4'; Q5 = 'Q5'; Q6 = 'Q6 (PD_Profile)'; Q7 = 'Q7'; Q8 = 'Q8'; Q9 = 'Q9'; Q10 = 'Q10'; Q11 = 'Q11' }
    for ($i = 0; $i -lt $VerChecks.Count; $i++) {
        $x = $VerChecks[$i]; $r = $CR0 + 1 + $i
        SetV $wsV.Cells.Item($r, 1) ($i + 1); $wsV.Cells.Item($r, 2).Value2 = $x[0]; $wsV.Cells.Item($r, 3).Value2 = $x[1]; $wsV.Cells.Item($r, 4).Value2 = $x[5]
        $wsV.Cells.Item($r, 5).Value2 = $tabName[$x[2]] + ': ' + $x[4] + ' / ' + $x[3]
        $ec = $wsV.Cells.Item($r, 6)
        if ($x[5] -eq 'num') { SetV $ec ([double]$x[6]); $ec.NumberFormat = 'General' }
        else { $ec.NumberFormat = '@'; if ($x[5] -eq 'flags' -and $x[6] -eq '') { $ec.Value2 = '(none)' } elseif ($x[5] -eq 'blank') { $ec.Value2 = '(blank)' } else { $ec.Value2 = $x[6] } }
        $wsV.Cells.Item($r, 7).Formula = LookupF $x[2] $x[3] $x[4]
        $wsV.Cells.Item($r, 7).NumberFormat = 'General'
        switch ($x[5]) {
            'num' { $m = "=IF(ISNUMBER(G$r),IF(F$r=0,ABS(G$r)<=nTOL_ABS,ABS(G$r-F$r)/ABS(F$r)<=nTOL_REL),FALSE)" }
            'text' { $m = "=EXACT(G$r&`"`",F$r&`"`")" }
            'flags' { $m = "=IF(F$r=`"(none)`",G$r&`"`"=`"`",EXACT(G$r&`"`",F$r&`"`"))" }
            'blank' { $m = "=G$r&`"`"=`"`"" }
            'contains' { $m = "=ISNUMBER(SEARCH(F$r,G$r))" }
        }
        $wsV.Cells.Item($r, 8).Formula = $m
    }
    $bRow = $cLast + 2
    $beh = @(
        'Behavioral vectors (inspection):',
        'V12 Versioning: Cases holds VEC-V12 version 1 and version 2 as separate rows, and Intervals keeps both results (9 calendar days for version 1, 7 for version 2; rows VEC-V12|1|I1 and VEC-V12|2|I1). Follow-up is added as a new Case_Version; earlier rows are never overwritten.',
        'V14 Export and re-import: on the Export sheet choose Case_ID VEC-V06, stamp the time, copy the rows as values and save them as CSV UTF-8 (the workbook has no macros). The row VEC-V06 I1 carries Interval_Minimum -5, Interval_Maximum 4, calendar days, SEQUENCE_UNCERTAIN, W_SEQUENCE_UNCERTAIN, the link L1 and, in Parameter_Set, the original partial-date text and precision; Tool_Version, Calculation_Version and Export_Timestamp are on every row. Re-open the CSV and compare: bounds, unit, original text and precision, link, sequence status and flags must be unchanged. Text that starts with =, +, -, @, a tab or a carriage return is exported with an apostrophe.',
        'Note on V15: the vector has no anchor or onset, so the interval row shows E_MISSING_INPUT and no value; the C6 check (C6_Code = I_APPROXIMATE_NOT_COMPARED) is the vector''s expected flag.',
        'Note on V35: the interval case has no Drug_Event_Links row and an onset of 30 February, so the row lists both E-codes (E_INVALID_DATE; E_MISSING_LINK), as rule G9 requires.'
    )
    for ($i = 0; $i -lt $beh.Count; $i++) { $wsV.Cells.Item($bRow + $i, 1).Value2 = $beh[$i]; if ($i -eq 0) { $wsV.Cells.Item($bRow, 1).Font.Bold = $true } }
    $vw = @(6, 10, 34, 9, 42, 30, 30, 9); for ($j = 0; $j -lt $vw.Count; $j++) { $wsV.Columns($j + 1).ColumnWidth = [double]$vw[$j] }
    foreach ($rg in @("H$($VR0+1):H$vLast")) {
        $fc = $wsV.Range($rg).FormatConditions
        [void]$fc.Add(1, 3, '="PASS"'); $fc.Item(1).Interior.Color = $C_Pass
        [void]$fc.Add(1, 3, '="FAIL"'); $fc.Item(2).Interior.Color = $C_Fail
    }
    $fc = $wsV.Range("H$($CR0+1):H$cLast").FormatConditions
    [void]$fc.Add(1, 3, '=TRUE'); $fc.Item(1).Interior.Color = $C_Pass
    [void]$fc.Add(1, 3, '=FALSE'); $fc.Item(2).Interior.Color = $C_Fail

    # ================================================================== SHARED: case selection and result lists
    $QT = @('Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11')
    function SelKeyF([string]$idCell, [string]$verCell) {
        '=IF(TRIM(' + $idCell + '&"")="","(none)",' + $idCell + '&"|"&IF(' + $verCell + '="",IFERROR(AGGREGATE(14,6,nCASE_Case_Version/(nCASE_Case_ID=' + $idCell + '),1),IFERROR(AGGREGATE(14,6,nINT_Case_Version/(nINT_Case_ID=' + $idCell + '),1),1)),' + $verCell + '))'
    }
    # k-th row (1-based within table T) whose case key equals $key; 0 when none
    function KRow([string]$t, [string]$key, [string]$k) {
        $ck = 'n' + $t + '_CaseKey'
        $mask = '(' + $ck + '=' + $key + ')'
        if ($t -eq 'INT') { $mask = $mask + '*(nINT_RowEmpty=0)' }
        'IFERROR(AGGREGATE(15,6,(ROW(' + $ck + ')-ROW(INDEX(' + $ck + ',1,1))+1)/(' + $mask + '),' + $k + '),0)'
    }
    function CountF([string]$t, [string]$key) { if ($t -eq 'INT') { 'COUNTIFS(nINT_CaseKey,' + $key + ',nINT_RowEmpty,0)' } else { 'COUNTIF(n' + $t + '_CaseKey,' + $key + ')' } }
    # writes a hidden block of per-table counts and cumulative starts; returns the two range addresses
    function UnionBlock($ws, [string]$col, [int]$row0, $tables, [string]$key) {
        $cN = ColL ($ws.Range($col + '1').Column + 1); $cS = ColL ($ws.Range($col + '1').Column + 2)
        for ($i = 0; $i -lt $tables.Count; $i++) {
            $r = $row0 + $i
            $ws.Range("$col$r").Value2 = $tables[$i]
            $ws.Range("$cN$r").Formula = '=' + (CountF $tables[$i] $key)
            if ($i -eq 0) { $ws.Range("$cS$r").Formula = '=0' } else { $ws.Range("$cS$r").Formula = "=$cS$($r-1)+$cN$($r-1)" }
        }
        $rl = $row0 + $tables.Count - 1
        return @{ N = "`$$cN`$$($row0):`$$cN`$$rl"; S = "`$$cS`$$($row0):`$$cS`$$rl"; T = "`$$col`$$($row0):`$$col`$$rl" }
    }
    $SAFE = { param($c) '=IF(ISNUMBER(' + $c + '),' + $c + ',IF(' + $c + '="","",IF(OR(LEFT(' + $c + ',1)="=",LEFT(' + $c + ',1)="+",LEFT(' + $c + ',1)="-",LEFT(' + $c + ',1)="@",LEFT(' + $c + ',1)=CHAR(9),LEFT(' + $c + ',1)=CHAR(13)),"' + "'" + '"&' + $c + ',' + $c + ')))' }

    # ================================================================== EXPORT (CALC-SPEC 1.1 Section 8; rule 11)
    $wsX = $S['Export']
    $XF = @('Case_ID', 'Case_Version', 'Result_ID', 'Label', 'Drug_Event_Link', 'Exposure_Episode_ID', 'Event_ID', 'End_Point', 'Interval_Value', 'Interval_Minimum', 'Interval_Maximum', 'Unit', 'Displayed_Value', 'Calculation_Mode', 'Anchor_Type', 'Anchor_Record_ID', 'Anchor_Description', 'Sequence_Status', 'Treatment_Day', 'Reported_Timing', 'Constraint_Source', 'Results', 'Parameter_Set', 'Assumptions', 'Calculation_Version', 'QC_Flags', 'Flag_Messages', 'Limitations', 'Export_Timestamp', 'Tool_Version')
    $XNum = @('Case_Version', 'Interval_Value', 'Interval_Minimum', 'Interval_Maximum', 'Treatment_Day')
    # source of each field: INT column, Q-table column ('' = blank); '=' prefix marks a raw expression with @T and @R placeholders
    $XMap = [ordered]@{
        Case_ID = @('Case_ID', 'Case_ID'); Case_Version = @('Case_Version', 'Case_Version'); Result_ID = @('Interval_ID', 'Calc_ID'); Label = @('Label', 'Label_Out')
        Drug_Event_Link = @('=IF(INDEX(nINT_LinkRow,@R)=0,"",INDEX(nINT_Link_ID,@R)&"")', ''); Exposure_Episode_ID = @('Exposure_Episode_ID', ''); Event_ID = @('Event_Episode_ID', '')
        End_Point = @('End_Record_ID_Used', ''); Interval_Value = @('Interval_Value', ''); Interval_Minimum = @('Interval_Minimum', ''); Interval_Maximum = @('Interval_Maximum', '')
        Unit = @('Unit', 'Unit_Out'); Displayed_Value = @('Displayed_Value', 'Displayed_Value'); Calculation_Mode = @('Calculation_Mode_Used', '="@T"'); Anchor_Type = @('Anchor_Type', '')
        Anchor_Record_ID = @('=IF(INDEX(nINT_K_Ok,@R)=1,INDEX(nINT_Anchor_Record_ID_Used,@R)&"","")', ''); Anchor_Description = @('Anchor_Description', ''); Sequence_Status = @('Sequence_Status', '')
        Treatment_Day = @('Treatment_Day', ''); Reported_Timing = @('Reported_Timing_Result', ''); Constraint_Source = @('=IF(INDEX(nINT_K_HasC,@R)=1,INDEX(nINT_Constraint_Source,@R)&"","")', '')
        Results = @('', 'Results_Text'); Parameter_Set = @('Parameter_Set', 'Parameter_Set'); Assumptions = @('Assumptions', 'Assumptions'); Calculation_Version = @('Calculation_Version', 'Calculation_Version')
        QC_Flags = @('QC_Flags', 'QC_Flags'); Flag_Messages = @('Messages', 'Messages'); Limitations = @('Limitations', 'Mandatory_Text')
    }
    $XN = 150; $XR0 = 8; $XLast = $XR0 + $XN - 1
    $wsX.Range('A1').Value2 = 'Export: one row per result, ready for CSV (CALC-SPEC 1.1 Section 8)'; $wsX.Range('A1').Font.Bold = $true; $wsX.Range('A1').Font.Size = 14; $wsX.Range('A1').Font.Color = $C_Title
    $wsX.Range('A2').Value2 = 'How to export: (1) enter the Case_ID in B4 (and a Case_Version in D4, or leave it blank for the latest version); (2) stamp the export time in F4 (Ctrl+; then a space, then Ctrl+Shift+;); (3) select the header row and the filled rows (from A' + ($XR0 - 1) + '), copy, and paste them as values into a new workbook (Paste Special > Values); (4) File > Save As > CSV UTF-8 (Comma delimited) (*.csv). Columns follow CALC-SPEC 1.1 Section 8 in order. Lists within a cell use "; " (QC_Flags) and " | " (Assumptions, Flag_Messages). Text that begins with =, +, -, @, a tab or a carriage return is already prefixed with an apostrophe so that a spreadsheet does not run it as a formula. Every row carries Tool_Version, Calculation_Version and Export_Timestamp. Intervals come first, then technical results (Q1-Q11).'
    $wsX.Range('A2:L2').Merge(); $wsX.Range('A2').WrapText = $true; $wsX.Rows(2).RowHeight = [double]66; $wsX.Range('A2').VerticalAlignment = -4160
    $wsX.Range('A4').Value2 = 'Case_ID'; $wsX.Range('C4').Value2 = 'Case_Version (blank = latest)'; $wsX.Range('E4').Value2 = 'Export timestamp (local time)'; $wsX.Range('G4').Value2 = 'Case key exported'
    foreach ($a in @('A4', 'C4', 'E4', 'G4')) { $wsX.Range($a).Font.Bold = $true; $wsX.Range($a).HorizontalAlignment = -4152 }
    foreach ($a in @('B4', 'D4', 'F4')) { $wsX.Range($a).Interior.Color = $C_In; $wsX.Range($a).Locked = $false; $wsX.Range($a).Borders.LineStyle = 1 }
    $wsX.Range('B4').NumberFormat = '@'; $wsX.Range('B4').Value2 = 'EX-C01'; $wsX.Range('D4').NumberFormat = '0'; $wsX.Range('F4').NumberFormat = 'yyyy-mm-dd hh:mm'
    Add-Val $wsX.Range('D4') 'posint' 'Leave blank to export the latest version of the case.'
    Add-Val $wsX.Range('F4') 'date' 'Type or stamp the date and time of the export (Ctrl+; space Ctrl+Shift+;).'
    $wsX.Range('H4').Formula = SelKeyF '$B$4' '$D$4'; $wsX.Range('H4').Font.Bold = $true
    $wsX.Range('I4').Formula = '=IF(H4="(none)","Enter a Case_ID",ROWS(A' + $XR0 + ':A' + $XLast + ')-COUNTBLANK(A' + $XR0 + ':A' + $XLast + ')&" result rows"&IF(ISNA(MATCH(H4,nCASE_Key,0)),"; no Cases row for this case version",""))'
    $wsX.Range('A6').Value2 = 'Rows below are formulas; copy them as values before saving as CSV. Capacity: ' + $XN + ' results per case version.'; $wsX.Range('A6').Font.Italic = $true
    $XH = $XR0 - 1
    for ($j = 0; $j -lt $XF.Count; $j++) { $c = $wsX.Cells.Item($XH, $j + 1); $c.Value2 = $XF[$j]; $c.Font.Bold = $true; $c.Interior.Color = $C_FHdr; $c.Font.Color = RGBc 255 255 255 }
    # hidden helpers: AF table index, AG k, AH source row; raw values from AJ; table block in CA:CC
    $ub = UnionBlock $wsX 'CA' $XR0 (@('INT') + $QT) '$H$4'
    $rawC0 = 36   # AJ
    $selTables = @('INT') + $QT
    for ($i = 0; $i -lt $XN; $i++) {
        $r = $XR0 + $i
        if ($i -eq 0) {
            $wsX.Range("AE$r").Formula = "=ROW()-$($XR0-1)"
            $wsX.Range("AF$r").Formula = "=IF(AE$r>SUM($($ub.N)),0,MATCH(AE$r-1,$($ub.S),1))"
            $wsX.Range("AG$r").Formula = "=IF(AF$r=0,0,AE$r-INDEX($($ub.S),AF$r))"
            $wsX.Range("AH$r").Formula = '=IF(AF' + $r + '=0,0,CHOOSE(AF' + $r + ',' + (($selTables | ForEach-Object { KRow $_ '$H$4' ('AG' + $r) }) -join ',') + '))'
            for ($j = 0; $j -lt $XF.Count; $j++) {
                $f = $XF[$j]; $L = ColL ($rawC0 + $j)
                if ($f -eq 'Export_Timestamp') { $e = '=IF($AF' + $r + '=0,"",IF(ISNUMBER($F$4),TEXT($F$4,"yyyy-mm-dd")&"T"&TEXT($F$4,"hh:mm")&":00","not stamped"))' }
                elseif ($f -eq 'Tool_Version') { $e = '=IF($AF' + $r + '=0,"","' + $ToolVersion + '")' }
                else {
                    $parts = @()
                    foreach ($t in $selTables) {
                        $src = $XMap[$f][$(if ($t -eq 'INT') { 0 } else { 1 })]
                        if (-not $src) { $parts += '""' }
                        elseif ($src.StartsWith('=')) { $parts += $src.Substring(1).Replace('@T', $t).Replace('@R', '$AH' + $r) }
                        elseif ($XNum -contains $f) { $parts += ('INDEX(n' + $t + '_' + $src + ',$AH' + $r + ')') }
                        else { $parts += ('INDEX(n' + $t + '_' + $src + ',$AH' + $r + ')&""') }
                    }
                    $e = '=IF($AF' + $r + '=0,"",CHOOSE($AF' + $r + ',' + ($parts -join ',') + '))'
                }
                try { $wsX.Range("$L$r").Formula = $e } catch { throw "Export formula rejected ($f): $e :: $($_.Exception.Message)" }
                $wsX.Range((ColL ($j + 1)) + $r).Formula = (& $SAFE ('$' + $L + $r))
            }
        }
    }
    $lastRaw = ColL ($rawC0 + $XF.Count - 1)
    $wsX.Range("AE$($XR0):$lastRaw$($XR0)").Copy($wsX.Range("AE$($XR0+1):$lastRaw$XLast")) | Out-Null
    $wsX.Range("A$($XR0):AD$($XR0)").Copy($wsX.Range("A$($XR0+1):AD$XLast")) | Out-Null
    $wsX.Range("A$($XR0):AD$XLast").Interior.Color = $C_F; $wsX.Range("A$($XR0):AD$XLast").Locked = $true
    $xw = @{ Case_ID = 10; Case_Version = 7; Result_ID = 12; Label = 26; Displayed_Value = 22; Calculation_Mode = 10; Parameter_Set = 40; Assumptions = 40; Flag_Messages = 40; Limitations = 30; Results = 30; QC_Flags = 24; Export_Timestamp = 18; Tool_Version = 22; Reported_Timing = 24 }
    for ($j = 0; $j -lt $XF.Count; $j++) { $w = 12; if ($xw.ContainsKey($XF[$j])) { $w = $xw[$XF[$j]] }; $wsX.Columns($j + 1).ColumnWidth = [double]$w }
    $wsX.Range("AE:$lastRaw").EntireColumn.Hidden = $true; $wsX.Range('CA:CC').EntireColumn.Hidden = $true
    [void]$wb.Names.Add('nEXPORT_Data', "='Export'!`$A`$$($XH):`$AD`$$XLast")
    Write-Host "Export sheet done"

    # ================================================================== SUMMARY (printable one-case summary)
    $wsS = $S['Summary']
    $sw = @(2, 13, 18, 30, 16, 26, 26, 16); for ($j = 0; $j -lt $sw.Count; $j++) { $wsS.Columns($j + 1).ColumnWidth = [double]$sw[$j] }
    $wsS.Range('B1').Value2 = 'TTO companion workbook: one-case summary'; $wsS.Range('B1').Font.Bold = $true; $wsS.Range('B1').Font.Size = 14; $wsS.Range('B1').Font.Color = $C_Title
    $wsS.Range('B2').Value2 = $VersionLabel + '. Educational tool; not validated for regulated production use; not a dosing tool; never assigns causality.'; $wsS.Range('B2').Font.Italic = $true
    $wsS.Range('B4').Value2 = 'Case_ID'; $wsS.Range('D4').Value2 = 'Case_Version (blank = latest)'; $wsS.Range('F4').Value2 = 'Case key shown'
    foreach ($a in @('B4', 'D4', 'F4')) { $wsS.Range($a).Font.Bold = $true }
    foreach ($a in @('C4', 'E4')) { $wsS.Range($a).Interior.Color = $C_In; $wsS.Range($a).Locked = $false; $wsS.Range($a).Borders.LineStyle = 1 }
    $wsS.Range('C4').NumberFormat = '@'; $wsS.Range('C4').Value2 = 'EX-C01'; $wsS.Range('E4').NumberFormat = '0'
    Add-Val $wsS.Range('C4') $null 'Type a Case_ID from the Cases sheet (for example EX-C01, EX-C04 or EX-C28).'
    Add-Val $wsS.Range('E4') 'posint' 'Leave blank to show the latest version of the case.'
    $wsS.Range('G4').Formula = SelKeyF '$C$4' '$E$4'; $wsS.Range('G4').Font.Bold = $true
    $wsS.Range('J4').Formula = '=IFERROR(MATCH($G$4,nCASE_Key,0),0)'
    $wsS.Range('H4').Formula = '=IF($G$4="(none)","Enter a Case_ID",IF($J$4=0,"No Cases row for this version","Case version found"))'
    function CaseF([string]$col, [bool]$date = $false) { if ($date) { '=IF($J$4=0,"",IF(INDEX(nCASE_' + $col + ',$J$4)="","",INDEX(nCASE_' + $col + ',$J$4)))' } else { '=IF($J$4=0,"",INDEX(nCASE_' + $col + ',$J$4)&"")' } }
    $kv = @(@('B5', 'Received', 'C5', 'Received_Date', $true), @('D5', 'Setting', 'E5', 'Setting', $false), @('G5', 'Reviewer', 'H5', 'Reviewer', $false),
            @('B6', 'Fictional', 'C6', 'Fictional', $false), @('D6', 'Review date', 'E6', 'Review_Date', $true), @('G6', 'Versions on Cases', 'H6', '', $false))
    foreach ($x in $kv) {
        $wsS.Range($x[0]).Value2 = $x[1]; $wsS.Range($x[0]).Font.Bold = $true
        if ($x[3]) { $wsS.Range($x[2]).Formula = CaseF $x[3] $x[4]; if ($x[4]) { $wsS.Range($x[2]).NumberFormat = 'd mmmm yyyy'; $wsS.Range($x[2]).HorizontalAlignment = -4131 } }
    }
    $wsS.Range('E5:F5').Merge()
    $wsS.Range('H6').Formula = '=IF($G$4="(none)","",COUNTIF(nCASE_Case_ID,$C$4))'; $wsS.Range('H6').HorizontalAlignment = -4131
    $wsS.Range('B7').Value2 = 'Summary'; $wsS.Range('B8').Value2 = 'This version adds'
    $wsS.Range('C7').Formula = CaseF 'Summary'; $wsS.Range('C8').Formula = CaseF 'Version_Note'
    foreach ($rr in 7, 8) { $wsS.Range("B$rr").Font.Bold = $true; $wsS.Range("C$($rr):H$rr").Merge(); $wsS.Range("C$rr").WrapText = $true; $wsS.Rows($rr).RowHeight = [double]26; $wsS.Range("B$($rr):H$rr").VerticalAlignment = -4160 }

    function SecHead([int]$r, [string]$t) { $c = $wsS.Range("B$r"); $c.Value2 = $t; $c.Font.Bold = $true; $c.Font.Size = 11; $c.Font.Color = $C_Title; $wsS.Range("B$($r):H$r").Borders.Item(9).LineStyle = 1 }
    function ColHeads([int]$r, $heads) { for ($j = 0; $j -lt $heads.Count; $j++) { $c = $wsS.Cells.Item($r, $j + 2); $c.Interior.Color = $C_FHdr; if ($heads[$j] -ne $null) { $c.Value2 = $heads[$j]; $c.Font.Bold = $true; $c.Font.Color = RGBc 255 255 255 } } }
    $KEY = '$G$4'
    # ---- exposures and links (6 rows)
    $rE = 10; SecHead $rE 'Exposures and drug-event links'
    ColHeads ($rE + 1) @('Exposure episode', 'Product', 'Role, route', 'Regimen (as reported)', 'Link', ('Exposure ' + $U_ARROW + ' event'), 'Reason')
    $nE = 6
    for ($i = 0; $i -lt $nE; $i++) {
        $r = $rE + 2 + $i; $k = $i + 1
        $wsS.Range("J$r").Formula = '=IFERROR(AGGREGATE(15,6,(ROW(nEXP_Key)-ROW(INDEX(nEXP_Key,1,1))+1)/((nEXP_Case_ID&"|"&nEXP_Case_Version=' + $KEY + ')*(nEXP_Exposure_Episode_ID<>"")),' + $k + '),0)'
        $wsS.Range("K$r").Formula = '=IFERROR(AGGREGATE(15,6,(ROW(nLNK_Key)-ROW(INDEX(nLNK_Key,1,1))+1)/((nLNK_Case_ID&"|"&nLNK_Case_Version=' + $KEY + ')*(nLNK_Link_ID<>"")),' + $k + '),0)'
        $wsS.Range("B$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nEXP_Exposure_Episode_ID,$J' + $r + ')&"")'
        $wsS.Range("C$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nEXP_Product,$J' + $r + ')&"")'
        $wsS.Range("D$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nEXP_Role,$J' + $r + ')&IF(INDEX(nEXP_Route,$J' + $r + ')&""="","",", "&INDEX(nEXP_Route,$J' + $r + ')))'
        $wsS.Range("E$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nEXP_Regimen_Text,$J' + $r + ')&"")'
        $wsS.Range("F$r").Formula = '=IF($K' + $r + '=0,"",INDEX(nLNK_Link_ID,$K' + $r + ')&"")'
        $wsS.Range("G$r").Formula = '=IF($K' + $r + '=0,"",INDEX(nLNK_Exposure_Episode_ID,$K' + $r + ')&" ' + $U_ARROW + ' "&INDEX(nLNK_Event_Episode_ID,$K' + $r + '))'
        $wsS.Range("H$r").Formula = '=IF($K' + $r + '=0,"",INDEX(nLNK_Reason,$K' + $r + ')&"")'
    }
    $rO = $rE + 2 + $nE
    $wsS.Range("B$rO").Formula = '=IF(OR(COUNTIFS(nEXP_Case_ID,$C$4,nEXP_Exposure_Episode_ID,"<>")>' + $nE + ',COUNTIFS(nLNK_Case_ID,$C$4,nLNK_Link_ID,"<>")>' + $nE + '),"More exposures or links are recorded than shown here; see the Exposures and Drug_Event_Links sheets.","")'
    # ---- dated records (26 rows)
    $rT = $rO + 1; SecHead $rT 'Dated records (Timepoints: every date and time as recorded)'
    ColHeads ($rT + 1) @('Record_ID', 'Record type', 'As recorded', 'Precision', 'Original text', $null, 'Source')
    $nT = 24
    for ($i = 0; $i -lt $nT; $i++) {
        $r = $rT + 2 + $i; $k = $i + 1
        $wsS.Range("J$r").Formula = '=IFERROR(AGGREGATE(15,6,nTP_RowN/(nTP_Case_Key=' + $KEY + '),' + $k + '),0)'
        $wsS.Range("B$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nTP_Record_ID,$J' + $r + ')&"")'
        $wsS.Range("C$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nTP_Record_Type,$J' + $r + ')&"")'
        $wsS.Range("D$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nTP_Display,$J' + $r + ')&"")'
        $wsS.Range("E$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nTP_Precision,$J' + $r + ')&IF(INDEX(nTP_Err,$J' + $r + ')<>""," ("&INDEX(nTP_Err,$J' + $r + ')&")",""))'
        $wsS.Range("F$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nTP_Original_Text,$J' + $r + ')&"")'
        $wsS.Range("H$r").Formula = '=IF($J' + $r + '=0,"",INDEX(nTP_Source,$J' + $r + ')&"")'
        $wsS.Range("F$($r):G$r").Merge()
    }
    $rO = $rT + 2 + $nT
    $wsS.Range("B$rO").Formula = '=IF(COUNTIF(nTP_Case_Key,' + $KEY + ')>' + $nT + ',COUNTIF(nTP_Case_Key,' + $KEY + ')-' + $nT + '&" more dated records are not shown; see the Timepoints sheet.","")'
    # ---- intervals (20 rows)
    $rI = $rO + 1; SecHead $rI 'Intervals (C1-C7) with flags'
    ColHeads ($rI + 1) @('Interval_ID', 'Link; anchor type', ('Anchor ' + $U_ARROW + ' end point'), 'Rule applied', 'Result', 'QC_Flags', 'Sequence')
    $nI = 16
    for ($i = 0; $i -lt $nI; $i++) {
        $r = $rI + 2 + $i; $k = $i + 1
        $wsS.Range("J$r").Formula = '=' + (KRow 'INT' $KEY $k)
        $J = '$J' + $r
        $wsS.Range("B$r").Formula = '=IF(' + $J + '=0,"",INDEX(nINT_Interval_ID,' + $J + ')&"")'
        $wsS.Range("C$r").Formula = '=IF(' + $J + '=0,"",INDEX(nINT_Link_ID,' + $J + ')&"; "&INDEX(nINT_Anchor_Type,' + $J + '))'
        $wsS.Range("D$r").Formula = '=IF(' + $J + '=0,"",IF(INDEX(nINT_K_Ok,' + $J + ')=1,INDEX(nINT_Label,' + $J + '),INDEX(nINT_Anchor_Record_ID_Used,' + $J + ')&" ' + $U_ARROW + ' "&INDEX(nINT_End_Record_ID_Used,' + $J + ')))'
        $wsS.Range("E$r").Formula = '=IF(' + $J + '=0,"",INDEX(nINT_Calculation_Mode_Used,' + $J + ')&"")'
        $wsS.Range("F$r").Formula = '=IF(' + $J + '=0,"",INDEX(nINT_Result_Display,' + $J + ')&"")'
        $wsS.Range("G$r").Formula = '=IF(' + $J + '=0,"",INDEX(nINT_QC_Flags,' + $J + ')&"")'
        $wsS.Range("H$r").Formula = '=IF(' + $J + '=0,"",INDEX(nINT_Sequence_Status,' + $J + ')&"")'
    }
    $rO = $rI + 2 + $nI
    $wsS.Range("B$rO").Formula = '=IF(' + (CountF 'INT' $KEY) + '>' + $nI + ',' + (CountF 'INT' $KEY) + '-' + $nI + '&" more interval rows are not shown; see the Intervals sheet.","")'
    # ---- technical results (12 rows)
    $rQ = $rO + 1; SecHead $rQ 'Technical results (Q1-Q11) with flags'
    ColHeads ($rQ + 1) @('Calc_ID', 'Rule', 'Question', 'Result', $null, 'QC_Flags', 'Unit')
    $ubS = UnionBlock $wsS 'R' 10 $QT $KEY
    $nQ = 10
    for ($i = 0; $i -lt $nQ; $i++) {
        $r = $rQ + 2 + $i; $k = $i + 1
        $wsS.Range("L$r").Formula = "=IF($k>SUM($($ubS.N)),0,MATCH($k-1,$($ubS.S),1))"
        $wsS.Range("M$r").Formula = "=IF(L$r=0,0,$k-INDEX($($ubS.S),L$r))"
        $wsS.Range("N$r").Formula = '=IF(L' + $r + '=0,0,CHOOSE(L' + $r + ',' + (($QT | ForEach-Object { KRow $_ $KEY ('M' + $r) }) -join ',') + '))'
        $pick = { param($col) '=IF($L' + $r + '=0,"",CHOOSE($L' + $r + ',' + (($QT | ForEach-Object { 'INDEX(n' + $_ + '_' + $col + ',$N' + $r + ')&""' }) -join ',') + '))' }
        $wsS.Range("B$r").Formula = (& $pick 'Calc_ID')
        $wsS.Range("C$r").Formula = '=IF($L' + $r + '=0,"",INDEX(' + $ubS.T + ',$L' + $r + '))'
        $wsS.Range("D$r").Formula = (& $pick 'Label_Out')
        $wsS.Range("E$r").Formula = (& $pick 'Result_Display')
        $wsS.Range("G$r").Formula = (& $pick 'QC_Flags')
        $wsS.Range("H$r").Formula = (& $pick 'Unit_Out')
        $wsS.Range("E$($r):F$r").Merge()
    }
    $rO = $rQ + 2 + $nQ
    $wsS.Range("B$rO").Formula = "=IF(SUM($($ubS.N))>$nQ,SUM($($ubS.N))-$nQ&"" more technical results are not shown; see PK_Exercises and PD_Profile."","""")"
    # ---- limitations and versions
    $rL = $rO + 1; SecHead $rL 'Limitations'
    $wsS.Range("B$($rL+1)").Value2 = $LIMIT_TOOL
    $wsS.Range("B$($rL+2)").Value2 = ('Intervals (CALC-SPEC 1.1 Section 3): ' + $LIMIT_C + $LIMIT_C4.Replace(' The true', ' For C4: the true') + $LIMIT_C7.Replace(' A recovery', ' For C7: a recovery'))
    $wsS.Range("B$($rL+3)").Value2 = 'Technical results (Section 4): each row carries its mandatory text, Parameter_Set and gate confirmations on PK_Exercises or PD_Profile. E-codes block a result; W-codes must be read with it. The formal causality assessment remains a human judgment.'
    foreach ($rr in ($rL + 1)..($rL + 3)) { $wsS.Range("B$($rr):H$rr").Merge(); $wsS.Range("B$rr").WrapText = $true; $wsS.Rows($rr).RowHeight = [double]24; $wsS.Range("B$rr").VerticalAlignment = -4160 }
    $rV = $rL + 5; SecHead $rV 'Versions'
    $wsS.Range("B$($rV+1)").Value2 = 'Tool version'; $wsS.Range("C$($rV+1)").Value2 = $ToolVersion
    $wsS.Range("E$($rV+1)").Value2 = 'Calculation version'; $wsS.Range("F$($rV+1)").Value2 = $SpecVersion
    $wsS.Range("B$($rV+2)").Value2 = 'Case version'; $wsS.Range("C$($rV+2)").Formula = '=IF($J$4=0,"",INDEX(nCASE_Case_Version,$J$4))'; $wsS.Range("C$($rV+2)").HorizontalAlignment = -4131
    $wsS.Range("E$($rV+2)").Value2 = 'Verification'; $wsS.Range("F$($rV+2)").Formula = '=LEFT(nVER_Overall,FIND(":",nVER_Overall&":")-1)&"; examples "&LEFT(nEX_Status,FIND(":",nEX_Status&":")-1)'
    $wsS.Range("B$($rV+3)").Value2 = 'Earlier versions of this case stay on the Cases sheet unchanged; choose a version in E4 to print it.'
    $wsS.Range("B$($rV+3):H$($rV+3)").Merge(); $wsS.Range("B$($rV+3)").Font.Italic = $true
    foreach ($rr in @(($rV + 1), ($rV + 2))) { $wsS.Range("B$rr").Font.Bold = $true; $wsS.Range("E$rr").Font.Bold = $true }
    $sLast = $rV + 3
    # formatting of the list areas
    foreach ($blk in @(@(($rE + 2), ($rE + 1 + $nE), 12), @(($rT + 2), ($rT + 1 + $nT), 12), @(($rI + 2), ($rI + 1 + $nI), 21), @(($rQ + 2), ($rQ + 1 + $nQ), 21))) {
        $rg = $wsS.Range("B$($blk[0]):H$($blk[1])")
        $rg.Font.Size = 8; $rg.VerticalAlignment = -4160; $rg.Borders.Item(12).LineStyle = 1; $rg.Borders.Item(12).Color = RGBc 217 217 217
        $wsS.Range("B$($blk[0]):B$($blk[1])").EntireRow.RowHeight = [double]$blk[2]
        if ($blk[2] -gt 12) { $rg.WrapText = $true }
    }
    $wsS.Range("B$($rE+1):H$($rE+1)").Font.Size = 8; $wsS.Range("B$($rT+1):H$($rT+1)").Font.Size = 8; $wsS.Range("B$($rI+1):H$($rI+1)").Font.Size = 8; $wsS.Range("B$($rQ+1):H$($rQ+1)").Font.Size = 8
    $wsS.Range("I:Z").EntireColumn.Hidden = $true
    $wsS.Range("A1:H$sLast").Locked = $true; $wsS.Range('C4').Locked = $false; $wsS.Range('E4').Locked = $false
    $SummaryLast = $sLast
    Write-Host "Summary sheet done"

    # ================================================================== READ_ME
    $wsR = $S['Read_Me']
    $wsR.Columns(1).ColumnWidth = [double]3; $wsR.Columns(2).ColumnWidth = [double]24; $wsR.Columns(3).ColumnWidth = [double]110
    $rm = New-Object System.Collections.ArrayList
    function AddRM($a, $b = '', $style = '') { [void]$rm.Add(@($a, $b, $style)) }
    AddRM 'Time to Onset in Drug Safety: companion Excel workbook' '' 'title'
    AddRM $VersionLabel '' 'sub'
    AddRM 'Companion to "Time to Onset in Drug Safety: A Technical Guide to ICSR Analysis Using Mechanism of Action, Pharmacokinetics and Pharmacodynamics" by Dr. Hafez Selim (Chapter 23, Use the Excel worksheets).' '' 'text'
    AddRM '' '' ''
    AddRM 'Educational tool: read before use' '' 'warnhead'
    AddRM 'This workbook is an educational tool for training, self-study and careful desk work on individual case safety reports (ICSRs). It is NOT validated for regulated production use. It is NOT a dosing calculator and gives no dosing recommendation. It NEVER assigns causality and never produces a causality category, a probability or a statement that an event is related to a drug. Every result describes timing or exposure under the assumptions shown beside it; the formal causality assessment remains a human judgment.' '' 'warn'
    AddRM 'Sheet protection only prevents accidental edits; it is not a security control. The workbook contains no macros, no VBA, no external links and no network connections, so case data stay on your computer. Use fictional or de-identified data for training.' '' 'warn'
    AddRM '' '' ''
    AddRM 'Version' ($VersionLabel + '. Calculation rules: CALC-SPEC 1.1 (Stage 1 document C), the same rules as the interactive calculator version 1.1. Verification vectors: ' + $VecFile + ' (V01-V38; 36 numeric).') 'kv'
    AddRM 'License' 'Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0), https://creativecommons.org/licenses/by-nc/4.0/' 'kv'
    AddRM 'Source and updates' ($RepoUrl + ' (the interactive calculator built on the same CALC-SPEC is at https://abdu94-hash.github.io/tto-companion-tools/)') 'kv'
    AddRM 'Verification status' '=nVER_Overall' 'kvf'
    AddRM 'Examples status' '=nEX_Status' 'kvf'
    AddRM 'Tested environment' ('Built and verified with Microsoft Excel for Windows, version ' + $xl.Version + ' (build ' + $xl.Build + '), on ' + (Get-Date -Format 'd MMMM yyyy') + ', and checked with verify_workbook.ps1 (the Verification and Examples status lines above are recalculated by the workbook itself). Formulas use only functions available in Excel 2010 and later (AGGREGATE, SUMPRODUCT, COUNTIFS; no LET, LAMBDA, dynamic arrays or TEXTJOIN). Not yet verified in Excel for the web, Excel for Mac, Google Sheets or LibreOffice. Some text formats (TEXT with yyyy-mm-dd on the Export sheet) assume English format codes.') 'kv'
    AddRM 'Changes in version 1.1' 'All 14 agreed rules of CALC-SPEC 1.1: every E-code listed (G9); the laboratory-defined onset (precision LAB); C5 by UTC instant with partial-date and tie blocking; C7 counting recorded interventions only, ends included; the complete common gate; Q6 named uses; Q8 antibody caution; Q10 Method B by date; units; display in powers of ten below 0.0001%; the finest rule by default; the C-rule limitations text; the Summary and Export sheets; vectors V28-V38.' 'kv'
    AddRM '' '' ''
    AddRM 'How to use (the workup of Chapter 23)' '' 'head'
    AddRM '1. Cases' 'Open a case version (Case_ID + Case_Version). Follow-up adds a new version; earlier rows stay (V12).' 'kv'
    AddRM '2. Record the sources' 'On Exposures, Administrations, Events and Milestones enter every date, time and timing statement as the source gives it: original text, precision, bounds and source. Never type 0, N/A or a typical value to make a row calculate. A laboratory-defined onset (CALC-SPEC 1.1 C4) is precision LAB on Events with the Milestone_IDs of the last normal test and the first abnormal test; no other date is used in their place.' 'kv'
    AddRM '3. Drug_Event_Links' 'Create the drug-event pairs you intend to compare, each with a reason. An interval without a link shows E_MISSING_LINK.' 'kv'
    AddRM '4. Intervals' 'One row per link and anchor. Leave the rule blank to use the finest rule both records support (C1 two timestamps, C2 two exact dates, C4 bounds), or choose a coarser one; a finer one is blocked (E_MISSING_INPUT). Leave the end point blank for the onset of the linked event. LAST_DOSE_BEFORE_ONSET with a blank anchor is selected by C5; FIRST_ADMIN with a blank anchor uses the episode Start. Reported timing (C6) sits beside the value. CESSATION to a RECOVERY milestone is a C7 recovery interval. Send a query for every E-code.' 'kv'
    AddRM '5. Product evidence' 'Product_Dossier (one row per extracted statement) and Parameter_Register (value, range, unit, population, half-life type, status, source).' 'kv'
    AddRM '6. Technical arithmetic' 'PK_Exercises (Q1-Q5, Q7-Q11) and PD_Profile (Q6 teaching demonstration). Run only what the question needs. The common gate (Q1-Q5, Q7, Q8) needs four confirmations (model, half-life used, reference time, Parameter_Status) and every detail: the half-life type, population and source, the reference time in words and each input''s Parameter_Status. Q8 shows a scenario range, never a midpoint or confidence interval. Q11 needs creatinine stability confirmed. The tools never calculate BSA; enter the BSA recorded at each administration.' 'kv'
    AddRM '7. Reviewer judgments' 'Temporal_Profile, Biology_Review (four options, each with a rationale; none is a relatedness category), Dechallenge_Rechallenge, Alternatives_Followup.' 'kv'
    AddRM '8. Assessment_QC, Summary and Export' 'Assessment_QC: narrative, method, unresolved issues and the counts of blocked and flagged interval rows. Summary: type a Case_ID (and a version, or blank for the latest) to see and print that case''s records, intervals with flags, technical results, limitations and versions on one or two pages. Export: one row per result with the CALC-SPEC 1.1 Section 8 columns, export-safe text, Tool_Version, Calculation_Version and Export_Timestamp; copy the rows as values and save them as CSV UTF-8.' 'kv'
    AddRM '' '' ''
    AddRM 'Sheets' '' 'head'
    AddRM 'Read_Me' 'This page.' 'kv'
    AddRM 'Cases' 'Input. One row per case version.' 'kv'
    AddRM 'Exposures' 'Input. One row per exposure episode; Start, Stop and Restart are timepoints (Episode/START, /STOP, /RESTART).' 'kv'
    AddRM 'Administrations' 'Input. One row per dose, dose change, other administration or infusion rate segment (start, end, rate); dose, BSA and date for Q10.' 'kv'
    AddRM 'Events' 'Input. One row per event episode; Onset and First detection are timepoints (Episode/ONSET, /DETECTION). Onset precision LAB = laboratory-defined onset.' 'kv'
    AddRM 'Milestones' 'Input. Tests, organ function, concentrations, interventions, recovery and other dated records.' 'kv'
    AddRM 'Drug_Event_Links' 'Reviewer. Selected pairs with a reason (no automatic pairing).' 'kv'
    AddRM 'Intervals' 'Derived. One row per link and anchor: value, unit, rule, sequence status, flags, messages, Parameter_Set, assumptions, limitations.' 'kv'
    AddRM 'Product_Dossier' 'Input. Versioned product statements, timing-statement kind, label discrepancies.' 'kv'
    AddRM 'Parameter_Register' 'Input. Parameters with status, population and source; a readiness check.' 'kv'
    AddRM 'PK_Exercises' 'Derived. Q1-Q5 and Q7-Q11 with gates, Parameter_Set, assumptions, mandatory text and flags.' 'kv'
    AddRM 'PD_Profile' 'Reviewer prompts; Q6 teaching demonstration (switch and use required; always I_TEACHING_DEMONSTRATION).' 'kv'
    AddRM 'Temporal_Profile, Biology_Review, Dechallenge_Rechallenge, Alternatives_Followup, Assessment_QC' 'Reviewer layer. Judgments in words, with rationales and sources.' 'kv'
    AddRM 'Summary' 'Printable one-case summary (choose the Case_ID in C4).' 'kv'
    AddRM 'Export' 'CSV-ready rows for one case version (choose the Case_ID in B4), CALC-SPEC 1.1 Section 8 columns.' 'kv'
    AddRM 'Examples' 'Fictional teaching cases EX-C01, EX-C04 and EX-C28 checked against their home chapters and the calculator''s example cases.' 'kv'
    AddRM 'Verification' 'All 36 numeric CALC-SPEC 1.1 vectors (V01-V38) recomputed with PASS/FAIL and an overall status; behavioral vectors V12 and V14 for inspection.' 'kv'
    AddRM 'Timepoints' 'Derived index of every dated record, with ISO 8601 and display forms; use its Record_ID values on Intervals.' 'kv'
    AddRM 'Flags_Messages' 'Every E-, W- and I-code with its verbatim message (CALC-SPEC 1.1 Section 5); QC_Flags list codes in this order.' 'kv'
    AddRM 'Lists' 'Validation lists.' 'kv'
    AddRM '' '' ''
    AddRM 'Conventions' '' 'head'
    AddRM 'Cells' 'Yellow = input (unlocked). Grey = formula (locked). Blue rows at the bottom of a sheet are reserved and locked: they hold the fictional teaching cases and the verification vectors read by the Examples and Verification sheets. Copy them into the yellow rows with a new Case_ID to experiment.' 'kv'
    AddRM 'Date precision' 'TIMESTAMP: Date, local Time and UTC offset (choose unknown if it is not recorded; C1 then uses local clock times with W_TZ_UNKNOWN). DATE: Date only, Time left blank (a date is never given a time). MONTH: Year and Month, expanded to the first and last day, leap years respected. YEAR: 1 January to 31 December. RANGE: earliest and latest dates as stated, with the source. LAB (event onset only): from the date of the last normal test to the date of the first abnormal test. A single date is never imputed from bounds.' 'kv'
    AddRM 'Dates and units' 'Dates are displayed as day month year (10 January 2026); ISO 8601 forms are on Timepoints and Intervals. Units are in the column headers. Only the conversions of CALC-SPEC Section 6 are automatic; weeks, months and years are never converted to days. A blank unit gives E_MISSING_INPUT; an unlisted unit E_INVALID_UNIT.' 'kv'
    AddRM 'Precision and display' 'Value columns hold unrounded results (G1). Displayed values follow CALC-SPEC 1.1 Section 7, as the calculator: whole minutes, hours to 1 decimal, whole calendar days, percentages to 3 significant figures (below 0.0001% in powers of ten, for example 3.41 x 10^-20 %), ratios to 2 decimals, mg and mg/m2 to 1 decimal, the percent of the planned dose (Q9) to 1 decimal, creatinine clearance in whole mL/minute.' 'kv'
    AddRM 'Flags' 'E-codes block a result: the row shows every E-code that applies and no number. W-codes are shown with the result; I-codes annotate it. QC_Flags lists each code once, E-codes then W-codes then I-codes, in Section 5 order (rule G9). Messages are verbatim from Flags_Messages; no spreadsheet error text is shown (G7).' 'kv'
    AddRM 'Protection' 'Every sheet is protected without a password (Review > Unprotect Sheet). Filtering is allowed. If you unprotect and change a formula, the Verification sheet no longer documents the workbook you are using.' 'kv'
    AddRM 'Export' 'Use the Export sheet (CALC-SPEC 1.1 Section 8): one row per result, export-safe text, Tool_Version, Calculation_Version and Export_Timestamp on every row; save as CSV UTF-8. Check the round trip (V14). Treatment-day numbers and analytical intervals are never written to E2B fields automatically.' 'kv'
    AddRM '' '' ''
    AddRM 'Implementation decisions (CALC-SPEC 1.1 Section 9: one agreed rule where the tools differed)' '' 'head'
    AddRM 'Choosing C1, C2 or C4' 'Agreed rule (Section 3): the finest rule both records support when the rule is left blank; a coarser rule on request; a finer one is blocked with E_MISSING_INPUT naming the missing precision (V01-V06).' 'kv'
    AddRM 'C5 last dose before onset' 'Agreed rule (C5 steps 1-4; V07, V33, V34): candidates are the episode''s Start and Restart and its DOSE, DOSE_CHANGE and OTHER administrations with an exact date or a timestamp (INFUSION_SEGMENT rows are Q9 rate segments, not candidates). Two timestamps are compared as instants, in UTC when both offsets are known, otherwise on the local clock; otherwise by date. The latest eligible candidate is selected. E_MISSING_INPUT when a position is uncertain, no candidate is eligible, the latest cannot be told apart (a tie), or a partial-date administration could be the last dose; a partial-date administration is never selected.' 'kv'
    AddRM 'C7 interventions' 'Agreed rule (C7; V28-V30): only INTERVENTION milestones of the case version count, never exposures (a treatment of the event recorded as an exposure is named in Assumptions, not counted). A recorded intervention counts when it can lie between cessation and recovery, ends included: as instants when all three are timestamps (UTC when both offsets of a comparison are known), otherwise by date. An intervention with an invalid or incomplete date blocks; one with no date is noted. The rule shows as C7 (C1), C7 (C2) or C7 (C4).' 'kv'
    AddRM 'Incomplete gate' 'Agreed rule (Section 4; V36): each blank gate detail (half-life type, population, source; reference time) and each missing Parameter_Status raises E_MISSING_INPUT; the gate is confirmed only when all four confirmations are Yes and nothing is blank, otherwise E_APPLICABILITY_UNCONFIRMED is also raised.' 'kv'
    AddRM 'Several errors' 'Agreed rule (G9; V35): every E-code that applies to a result''s own inputs is listed, in Section 5 order.' 'kv'
    AddRM 'Q8 for an antibody' 'Agreed rule (Q8; V37): when the molecule is recorded as a monoclonal antibody (column Antibody = Yes), W_BIOLOGIC_MODEL_CAUTION is raised before I_SCENARIO_NOT_CI.' 'kv'
    AddRM 'Q10 Method B' 'Agreed rule (Q10; V38): first and last are the earliest and latest administration by date; with no date at all, the order entered (stated in Assumptions); some dates missing, or the earliest or latest date shared with different BSAs, blocks with E_MISSING_INPUT. BSA is never calculated.' 'kv'
    AddRM 'Units and small percentages' 'Agreed rules (G5; Section 7): a blank unit is E_MISSING_INPUT and an unlisted unit E_INVALID_UNIT; percentages below 0.0001% are shown with three significant figures in powers of ten; the Q9 percent of the planned dose has one decimal.' 'kv'
    AddRM 'Export' 'Agreed rule (Section 8; V14): the Export sheet carries the Section 8 columns in order, with Tool_Version, Calculation_Version and Export_Timestamp on every row.' 'kv'
    AddRM 'Laboratory-defined onset' 'Input rule (C4; V31): onset precision LAB on Events, bounded by the date of the last normal test and the date of the first abnormal test (earliest and latest dates when partial). A missing or undated test gives E_MISSING_INPUT; a last normal test after the first abnormal test gives E_REVERSED_BOUNDS.' 'kv'
    AddRM 'Q6 named uses' 'Q6 (V21, V32): the use is required: replacement or recovery (fraction replaced) or approach to a new level (fraction of the change completed); one formula.' 'kv'
    AddRM 'Workbook layout choices' 'C1 timestamps are converted to whole minutes (date x 1440 + minutes of the day, minus the UTC offset), so intervals are exact integers. A C4 constraint applies to the anchor bounds and is shown on the same row beside the unconstrained result. C3 and C6 are not applied on C7 rows (as in the calculator, where recovery intervals have no reported timing or treatment day). Q7 accepts days only. With the Q6 switch off, Q6 shows E_APPLICABILITY_UNCONFIRMED. If any Q9 segment or the onset has an unknown offset, local clock times are used with W_TZ_UNKNOWN.' 'kv'
    $rr = 1
    foreach ($e in $rm) {
        $a = $e[0]; $b = $e[1]; $st = $e[2]
        switch ($st) {
            'title' { $wsR.Cells.Item($rr, 2).Value2 = $a; $wsR.Cells.Item($rr, 2).Font.Size = 18; $wsR.Cells.Item($rr, 2).Font.Bold = $true; $wsR.Cells.Item($rr, 2).Font.Color = $C_Title }
            'sub' { $wsR.Cells.Item($rr, 2).Value2 = $a; $wsR.Cells.Item($rr, 2).Font.Size = 13; $wsR.Cells.Item($rr, 2).Font.Bold = $true }
            'text' { $wsR.Cells.Item($rr, 2).Value2 = $a; $wsR.Range("B$($rr):C$rr").Merge(); $wsR.Cells.Item($rr, 2).WrapText = $true; $wsR.Rows($rr).RowHeight = [double]28 }
            'warnhead' { $wsR.Cells.Item($rr, 2).Value2 = $a; $wsR.Range("B$($rr):C$rr").Merge(); $wsR.Cells.Item($rr, 2).Font.Bold = $true; $wsR.Cells.Item($rr, 2).Font.Size = 13; $wsR.Range("B$($rr):C$rr").Interior.Color = $C_Warn; $wsR.Cells.Item($rr, 2).Font.Color = RGBc 156 0 6 }
            'warn' { $wsR.Cells.Item($rr, 2).Value2 = $a; $wsR.Range("B$($rr):C$rr").Merge(); $wsR.Cells.Item($rr, 2).WrapText = $true; $wsR.Range("B$($rr):C$rr").Interior.Color = $C_Warn; $wsR.Cells.Item($rr, 2).Font.Bold = $true; $wsR.Rows($rr).RowHeight = [double]58 }
            'head' { $wsR.Cells.Item($rr, 2).Value2 = $a; $wsR.Cells.Item($rr, 2).Font.Bold = $true; $wsR.Cells.Item($rr, 2).Font.Size = 12; $wsR.Cells.Item($rr, 2).Font.Color = $C_Title }
            'kv' { $wsR.Cells.Item($rr, 2).Value2 = $a; $wsR.Cells.Item($rr, 2).Font.Bold = $true; $wsR.Cells.Item($rr, 2).WrapText = $true; $wsR.Cells.Item($rr, 2).VerticalAlignment = -4160; $wsR.Cells.Item($rr, 3).Value2 = $b; $wsR.Cells.Item($rr, 3).WrapText = $true; $wsR.Cells.Item($rr, 3).VerticalAlignment = -4160; [void]$wsR.Rows($rr).AutoFit() }
            'kvf' { $wsR.Cells.Item($rr, 2).Value2 = $a; $wsR.Cells.Item($rr, 2).Font.Bold = $true; $wsR.Cells.Item($rr, 3).Formula = $b; $wsR.Cells.Item($rr, 3).Font.Bold = $true }
        }
        $rr++
    }
    $wsR.Range("B1:C$rr").VerticalAlignment = -4160
    [void]$wsR.Hyperlinks.Add($wsR.Range('C11'), $RepoUrl, $Missing, $Missing, $wsR.Range('C11').Text)

    # ================================================================== FINISH: calc, freeze panes, filters, protection
    $xl.Calculation = -4105
    $xl.CalculateFull()
    $xl.ScreenUpdating = $true
    $freeze = @{ Cases = 2; Exposures = 4; Administrations = 3; Events = 4; Milestones = 3; Drug_Event_Links = 3; Intervals = 3; Product_Dossier = 1; Parameter_Register = 1; Temporal_Profile = 3; Biology_Review = 3; Dechallenge_Rechallenge = 3; Alternatives_Followup = 2; Assessment_QC = 2 }
    $tblInfo = @{ Cases = $iCase; Exposures = $iExp; Administrations = $iAdm; Events = $iEvt; Milestones = $iMs; Drug_Event_Links = $iLnk; Intervals = $iInt; Product_Dossier = $iPDS; Parameter_Register = $iPR; Temporal_Profile = $iTPR; Biology_Review = $iBIO; Dechallenge_Rechallenge = $iDCR; Alternatives_Followup = $iALT; Assessment_QC = $iAQC }
    $wb.Activate()
    foreach ($name in $SheetOrder) {
        $ws = $S[$name]; [void]$ws.Activate()
        $win = $xl.ActiveWindow
        try { $win.FreezePanes = $false; $win.ScrollRow = 1; $win.ScrollColumn = 1 } catch {}
        if ($tblInfo.ContainsKey($name)) {
            $inf = $tblInfo[$name]
            try { $win.SplitColumn = $freeze[$name]; $win.SplitRow = $inf.HeadRow; $win.FreezePanes = $true } catch { Write-Host "freeze failed on $name" }
            try { [void]$ws.Range("A$($inf.HeadRow):$($inf.LastCol)$($inf.Last)").AutoFilter() } catch { Write-Host "autofilter failed on $name" }
        } elseif ($name -eq 'Export') {
            try { $win.SplitColumn = 3; $win.SplitRow = $XH; $win.FreezePanes = $true } catch {}
        } elseif ($name -eq 'Timepoints') {
            try { $win.SplitColumn = 2; $win.SplitRow = 4; $win.FreezePanes = $true } catch {}
            try { [void]$ws.Range("A4:" + (ColL $tpCols.Count) + $tpLast).AutoFilter() } catch {}
        }
        try { $win.Zoom = 90 } catch {}
        [void]$ws.Range('A1').Select()
        # protection without a password; filtering and column/row formatting allowed
        $ws.Protect($Missing, $true, $true, $true, $false, $false, $true, $true, $false, $false, $false, $false, $false, $false, $true, $false)
    }
    # page setup for printing Read_Me, Examples and Verification
    $xl.PrintCommunication = $false
    foreach ($name in @('Read_Me', 'Examples', 'Verification')) {
        $ps = $S[$name].PageSetup
        $ps.PrintArea = $S[$name].UsedRange.Address()
        $ps.Orientation = 2; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = $false
        $ps.LeftFooter = $VersionLabel + ' - educational tool; not validated for regulated use; not a dosing tool; never assigns causality'
        $ps.RightFooter = 'Page &P of &N'
    }
    # Summary: one case on one or two pages, landscape, fitted to the page width
    $ps = $wsS.PageSetup
    $ps.PrintArea = "`$A`$1:`$H`$$SummaryLast"
    $ps.Orientation = 2; $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = 2
    $ps.LeftMargin = $xl.InchesToPoints(0.4); $ps.RightMargin = $xl.InchesToPoints(0.4); $ps.TopMargin = $xl.InchesToPoints(0.5); $ps.BottomMargin = $xl.InchesToPoints(0.6)
    $ps.HeaderMargin = $xl.InchesToPoints(0.25); $ps.FooterMargin = $xl.InchesToPoints(0.3)
    $ps.CenterHorizontally = $true
    $ps.LeftFooter = $VersionLabel + ' - educational tool; not validated for regulated use; not a dosing tool; never assigns causality'
    $ps.RightFooter = 'Printed &D - page &P of &N'
    $ps.PrintTitleRows = ''
    $xl.PrintCommunication = $true
    try {
        $dp = $wb.BuiltinDocumentProperties
        foreach ($kvp in @(@('Title', 'TTO Companion Workbook v1.1'), @('Author', 'Dr. Hafez Selim'), @('Comments', ($VersionLabel + '. Educational tool; CC BY-NC 4.0; ' + $RepoUrl)))) {
            $prop = [System.__ComObject].InvokeMember('Item', [System.Reflection.BindingFlags]::GetProperty, $null, $dp, @($kvp[0]))
            [void][System.__ComObject].InvokeMember('Value', [System.Reflection.BindingFlags]::SetProperty, $null, $prop, @($kvp[1]))
        }
    } catch { Write-Host ("document properties not set: " + $_.Exception.Message) }
    [void]$S['Read_Me'].Activate()
    $wb.SaveAs($OutPath, 51)
    Write-Host ("Saved " + $OutPath + " in " + [int]((Get-Date) - $t0).TotalSeconds + " s")
    Write-Host ("Verification: " + $S['Verification'].Range('C4').Text)
    Write-Host ("Examples: " + $S['Examples'].Range('C4').Text)
    $wb.Close($false); $wb = $null
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {}; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    try { [TtoMessageFilter]::Revoke() } catch {}
}

# ================================================================== PACKAGE CLEAN-UP
# Excel stores the build machine's printer (paper size and a printerSettings part) in the page setup. Make the
# printed sheets printer-neutral: US Letter paper, no printerSettings parts (the reader's own printer is used).
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8 = New-Object System.Text.UTF8Encoding($false)
$zip = [System.IO.Compression.ZipFile]::Open($OutPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    function ReadEntry($e) { $sr = New-Object System.IO.StreamReader($e.Open(), $utf8); try { $sr.ReadToEnd() } finally { $sr.Dispose() } }
    function WriteEntry($zip, [string]$name, [string]$text) {
        $old = $zip.GetEntry($name); if ($old) { $old.Delete() }
        $ne = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
        $sw = New-Object System.IO.StreamWriter($ne.Open(), $utf8); try { $sw.Write($text) } finally { $sw.Dispose() }
    }
    $names = @($zip.Entries | ForEach-Object { $_.FullName })
    foreach ($n in $names) {
        if ($n -match '^xl/worksheets/sheet\d+\.xml$') {
            $t = ReadEntry $zip.GetEntry($n)
            if ($t -match '<pageSetup ') {
                $t2 = [regex]::Replace($t, '<pageSetup [^>]*/>', { param($m) ([regex]::Replace(([regex]::Replace($m.Value, ' paperSize="\d+"', ' paperSize="1"')), ' r:id="rId\d+"', '')) })
                if ($t2 -notmatch '<pageSetup [^>]*paperSize=') { $t2 = $t2.Replace('<pageSetup ', '<pageSetup paperSize="1" ') }
                if ($t2 -ne $t) { WriteEntry $zip $n $t2 }
            }
        } elseif ($n -match '^xl/worksheets/_rels/sheet\d+\.xml\.rels$') {
            $t = ReadEntry $zip.GetEntry($n)
            $t2 = [regex]::Replace($t, '<Relationship [^>]*relationships/printerSettings"[^>]*/>', '')
            if ($t2 -ne $t) { WriteEntry $zip $n $t2 }
        }
    }
    foreach ($n in $names) { if ($n -match '^xl/printerSettings/') { $zip.GetEntry($n).Delete() } }
    $ct = ReadEntry $zip.GetEntry('[Content_Types].xml')
    $ct2 = [regex]::Replace($ct, '<Default Extension="bin" ContentType="application/vnd\.openxmlformats-officedocument\.spreadsheetml\.printerSettings"/>', '')
    if ($ct2 -ne $ct) { WriteEntry $zip '[Content_Types].xml' $ct2 }
    Write-Host "Package: page setup set to US Letter; printer-specific settings removed"
} finally { $zip.Dispose() }
