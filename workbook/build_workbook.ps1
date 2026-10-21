<#
.SYNOPSIS
  Builds TTO_Companion_Workbook_v1.0.xlsx, the companion Excel workbook described in Chapter 23 of
  "Time to Onset in Drug Safety" (Dr. Hafez Selim), through Excel COM automation.

.DESCRIPTION
  TTO Companion Tools v1.0 (CALC-SPEC 1.0). Formulas only: no macros, no VBA, no external links.
  Rules C1-C7 and Q1-Q11, flags and messages follow CALC-SPEC 1.0 (Stage 1 document C).
  The fictional teaching cases (EX-C01, EX-C04, EX-C28) and the verification vectors (VEC-V01 ... VEC-V27)
  are written as reserved, locked rows on the production sheets, so every vector is computed by the
  same column formulas a user's own case uses. The Verification and Examples sheets read those rows.

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
if (-not $OutPath) { $OutPath = Join-Path $Here 'TTO_Companion_Workbook_v1.0.xlsx' }
$VersionLabel = 'TTO Companion Tools v1.0 (CALC-SPEC 1.0)'
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
function OFFMIN([string]$x) { '(IF(LEFT(' + $x + ',1)="-",-1,1)*(VALUE(MID(' + $x + ',2,2))*60+VALUE(MID(' + $x + ',5,2))))' }
function DTXT([string]$x) { 'DAY(' + $x + ')&" "&INDEX(nLIST_Months,MONTH(' + $x + '))&" "&YEAR(' + $x + ')' }
function ISOD([string]$x) { 'YEAR(' + $x + ')&"-"&RIGHT("0"&MONTH(' + $x + '),2)&"-"&RIGHT("0"&DAY(' + $x + '),2)' }
function HHMM([string]$m) { 'RIGHT("0"&INT(MOD(' + $m + ',1440)/60),2)&":"&RIGHT("0"&MOD(' + $m + ',60),2)' }
function SEQ([string]$x) { 'IF(' + $x + '>0,"AFTER",IF(' + $x + '<0,"BEFORE","SIMULTANEOUS"))' }
# Three significant figures as text (CALC-SPEC Section 7); scientific form below 0.0001
function P3([string]$x) {
    'IF(' + $x + '=0,"0",IF(ABS(' + $x + ')<0.0001,FIXED(' + $x + '/10^INT(LOG10(ABS(' + $x + '))),2,TRUE)&" x 10^"&INT(LOG10(ABS(' + $x + '))),FIXED(ROUND(' + $x + ',2-INT(LOG10(ABS(' + $x + ')))),MAX(0,2-INT(LOG10(ABS(' + $x + ')))),TRUE)))'
}
function MergeH($hs) { $r = [ordered]@{}; foreach ($h in $hs) { foreach ($k in $h.Keys) { $r[$k] = $h[$k] } }; $r }

# ------------------------------------------------------------------ column-spec helpers
# k: 'in' (input, unlocked in user rows) | 'f' (formula, locked)
function CIn($n, $h, $w = 12, $v = $null, $fmt = $null, $g = $null) { @{ n = $n; h = $h; k = 'in'; w = $w; v = $v; fmt = $fmt; g = $g } }
function CF($n, $h, $f, $w = 12, $fmt = $null, $g = $null) { @{ n = $n; h = $h; k = 'f'; w = $w; f = $f; fmt = $fmt; g = $g } }

# Timepoint block (CALC-SPEC Section 2): 10 inputs + 7 derived columns. '@' is replaced by the prefix.
function TPBlock([string]$p, [string]$label) {
    $tErr = '=IF(COUNTA({@_Prec}:{@_Upper})=0,"",IF({@_Prec}="","E_MISSING_INPUT",' +
      'IF({@_Prec}="TIMESTAMP",IF(OR({@_Date}="",{@_Time}="",{@_Off}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({@_Date})),NOT(ISNUMBER({@_Time}))),"E_INVALID_DATE",IF(OR({@_Date}<>INT({@_Date}),{@_Time}<0,{@_Time}>=1),"E_INVALID_DATE",IF(ISNA(MATCH({@_Off},nLIST_Offset,0)),"E_INVALID_UNIT","")))),' +
      'IF({@_Prec}="DATE",IF({@_Date}="","E_MISSING_INPUT",IF(NOT(ISNUMBER({@_Date})),"E_INVALID_DATE",IF({@_Date}<>INT({@_Date}),"E_INVALID_DATE",""))),' +
      'IF({@_Prec}="MONTH",IF(OR({@_Year}="",{@_Month}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({@_Year})),NOT(ISNUMBER({@_Month}))),"E_INVALID_DATE",IF(OR({@_Year}<1900,{@_Year}>9999,{@_Year}<>INT({@_Year}),{@_Month}<1,{@_Month}>12,{@_Month}<>INT({@_Month})),"E_INVALID_DATE",""))),' +
      'IF({@_Prec}="YEAR",IF({@_Year}="","E_MISSING_INPUT",IF(NOT(ISNUMBER({@_Year})),"E_INVALID_DATE",IF(OR({@_Year}<1900,{@_Year}>9999,{@_Year}<>INT({@_Year})),"E_INVALID_DATE",""))),' +
      'IF({@_Prec}="RANGE",IF(OR({@_Lower}="",{@_Upper}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({@_Lower})),NOT(ISNUMBER({@_Upper}))),"E_INVALID_DATE",IF(OR({@_Lower}<>INT({@_Lower}),{@_Upper}<>INT({@_Upper})),"E_INVALID_DATE",IF({@_Lower}>{@_Upper},"E_REVERSED_BOUNDS","")))),' +
      '"E_INVALID_UNIT")))))))'
    $tHT = '=IF(AND({@_Err}="",{@_Prec}="TIMESTAMP"),1,0)'
    $tOK = '=IF(AND({@_HT}=1,{@_Off}<>"unknown"),1,0)'
    $tLo = '=IF(OR({@_Err}<>"",{@_Prec}=""),"",IF(OR({@_Prec}="TIMESTAMP",{@_Prec}="DATE"),{@_Date},IF({@_Prec}="MONTH",DATE({@_Year},{@_Month},1),IF({@_Prec}="YEAR",DATE({@_Year},1,1),{@_Lower}))))'
    $tHi = '=IF(OR({@_Err}<>"",{@_Prec}=""),"",IF(OR({@_Prec}="TIMESTAMP",{@_Prec}="DATE"),{@_Date},IF({@_Prec}="MONTH",EOMONTH(DATE({@_Year},{@_Month},1),0),IF({@_Prec}="YEAR",DATE({@_Year},12,31),{@_Upper}))))'
    $tLoc = '=IF({@_HT}=1,{@_Date}*1440+ROUND({@_Time}*1440,0),"")'
    $tUTC = '=IF({@_OK}=1,{@_Loc}-' + (OFFMIN '{@_Off}') + ',"")'
    $cols = @(
        (CIn "${p}_Text" "${label}: original text (as reported)" 26 $null '@' "${label} (timepoint)"),
        (CIn "${p}_Prec" "${label}: precision" 11 'list:nLIST_Precision'),
        (CIn "${p}_Date" "${label}: date (TIMESTAMP, DATE)" 15 'date' 'd mmmm yyyy'),
        (CIn "${p}_Time" "${label}: local time hh:mm (TIMESTAMP only)" 9 'time' 'hh:mm'),
        (CIn "${p}_Off" "${label}: UTC offset or unknown (TIMESTAMP)" 9 'list:nLIST_Offset' '@'),
        (CIn "${p}_Year" "${label}: year (MONTH, YEAR)" 7 'year' '0'),
        (CIn "${p}_Month" "${label}: month 1-12 (MONTH)" 7 'month' '0'),
        (CIn "${p}_Lower" "${label}: earliest date (RANGE)" 15 'date' 'd mmmm yyyy'),
        (CIn "${p}_Upper" "${label}: latest date (RANGE)" 15 'date' 'd mmmm yyyy'),
        (CIn "${p}_Src" "${label}: source (document, record)" 22 $null '@'),
        (CF "${p}_Err" "${label}: check (E-code)" $tErr 12),
        (CF "${p}_HT" "${label}: has time (1/0)" $tHT 6 '0'),
        (CF "${p}_OK" "${label}: offset known (1/0)" $tOK 6 '0'),
        (CF "${p}_Lo" "${label}: lower bound (date)" $tLo 14 'd mmmm yyyy'),
        (CF "${p}_Hi" "${label}: upper bound (date)" $tHi 14 'd mmmm yyyy'),
        (CF "${p}_Loc" "${label}: local clock (minutes since 1900)" $tLoc 11 '0'),
        (CF "${p}_UTC" "${label}: UTC (minutes since 1900)" $tUTC 11 '0')
    )
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
    SCrUnit       = @('mg/dL', ([string][char]0x00B5 + 'mol/L'))
    AdminKind     = @('DOSE', 'INFUSION_SEGMENT', 'OTHER')
    RateUnit      = @('mg/hour', 'mL/hour')
    DoseUnit      = @('mg', 'micrograms', 'g', 'mL', 'IU', 'units', 'other (see regimen text)')
    Role          = @('suspect', 'concomitant', 'interacting', 'prophylaxis', 'treatment of the event', 'other')
    MilestoneType = @('BASELINE', 'TEST', 'ORGAN_FUNCTION', 'CONCENTRATION', 'INTERVENTION', 'RECOVERY', 'DIAGNOSIS', 'EXPOSURE_CHANGE', 'OTHER')
    MethodB       = @('FIRST', 'LAST', 'STATED')
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

# CALC-SPEC 1.0 Section 5, verbatim
$Flags = @(
    @('E_INVALID_DATE', 'Error', 'A date does not exist or cannot be read', 'This date does not exist. Check the source, or record it as a partial date.'),
    @('E_REVERSED_BOUNDS', 'Error', 'A lower bound is after or above its upper bound', 'The earliest value is after the latest value. Check the order of the bounds.'),
    @('E_DUPLICATE_ID', 'Error', 'Two records share an identifier that must be unique', 'This identifier is already used in this case. Each episode needs its own ID.'),
    @('E_MISSING_LINK', 'Error', 'An interval is requested without a Drug_Event_Link', 'Link this drug episode to this event before calculating an interval.'),
    @('E_MISSING_INPUT', 'Error', 'A required input is blank', 'This result needs [input]. It is left blank rather than estimated.'),
    @('E_INVALID_UNIT', 'Error', 'A unit not listed in Section 6', 'This unit is not supported here. Use one of the listed units.'),
    @('E_NONPOSITIVE_PARAMETER', 'Error', 'A half-life, clearance, volume, interval, rate, dose or BSA is zero or negative', 'This value must be greater than zero.'),
    @('E_OUT_OF_RANGE', 'Error', 'Target fraction outside 0-1, negative time, or age outside 18-120', 'This value is outside the range the equation accepts.'),
    @('E_APPLICABILITY_UNCONFIRMED', 'Error', 'The common gate is not confirmed', 'Confirm the model, the half-life used, the reference time and each input''s status before a result is shown.'),
    @('E_CREATININE_STABILITY_UNCONFIRMED', 'Error', 'Q11 without stability confirmation', 'Cockcroft-Gault assumes stable creatinine. Confirm stability before an estimate is shown.'),
    @('E_OVERLAPPING_SEGMENTS', 'Error', 'Infusion segments overlap', 'Two infusion segments overlap. Check the recorded start and stop times.'),
    @('W_TZ_UNKNOWN', 'Warning', 'A UTC offset is unknown', 'Offset unknown: this interval uses local clock times and could be wrong across a clock change or time zone.'),
    @('W_SAME_DATE_ORDER_UNKNOWN', 'Warning', 'Same date and no times', 'Same calendar date: the order within the day is unknown.'),
    @('W_SEQUENCE_UNCERTAIN', 'Warning', 'Bounds contain zero', 'The available dates allow onset before or after this anchor. Sequence is not established.'),
    @('W_ADMIN_SAME_DATE_AS_ONSET', 'Warning', 'The selected last dose shares the onset date without times', 'This dose was on the onset date; whether it came before onset is unknown.'),
    @('W_REPORTED_TIMING_CONFLICT', 'Warning', 'Reported timing does not overlap the computed interval', 'The reported timing conflicts with the recorded dates. Both are kept; query the source.'),
    @('W_INTERVENTION_BEFORE_RECOVERY', 'Warning', 'Treatment recorded between cessation and recovery', 'Treatment was given before recovery, so recovery cannot be attributed to withdrawal alone.'),
    @('W_BIOLOGIC_MODEL_CAUTION', 'Warning', 'Every Q7 result', 'Antibody disposition can be nonlinear or time-varying; this simple model may not apply.'),
    @('W_SEX_NOT_RECORDED', 'Warning', 'Q11 without recorded sex', 'Sex not recorded: both estimates are shown.'),
    @('W_PARTIAL_BSA', 'Warning', 'Q10 with a missing BSA', 'BSA missing for some administrations: the value shown is a lower bound.'),
    @('I_TREATMENT_DAY_UNDEFINED', 'Info', 'C2 is negative', 'No treatment-day number exists before the first dose.'),
    @('I_APPROXIMATE_NOT_COMPARED', 'Info', 'Reported timing has no bounds', 'Approximate timing is kept as reported and not converted.'),
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
function GY($params, $ref) { @{ G_Model = 'Yes'; G_Params = $params; G_RefTime = $ref; G_Status = 'Yes' } }
function PV($n, $v, $st, $src) { @{ $n = $v; "${n}_Status" = $st; "${n}_Source" = $src } }

$EX_LABEL = 'Reserved, locked rows: fictional teaching cases EX-C01 (case version 2), EX-C04 (version 2) and EX-C28 (version 2), read by the Examples sheet. Copy rows to the yellow area with a new Case_ID to experiment.'
$VEC_LABEL = 'Reserved, locked rows: CALC-SPEC 1.0 verification vectors (VEC-V01 ... VEC-V27), read by the Verification sheet. They use the same column formulas as your own rows.'

# ---------------- Cases
$CasesEx = @(
    @{ Case_ID = 'EX-C01'; Case_Version = 2; Received_Date = '2026-01-14'; Setting = 'Hospital infusion (fictional)'; Fictional = 'Yes'; Summary = 'Urticaria, wheeze and hypotension 15 minutes after the start of a first Drug A infusion (Chapter 3, Section 3.7).'; Version_Note = 'Version 2 (follow-up 14 January 2026) adds Drug B, the antiseptic, pump rate, first exposure, diagnosis and tryptase sample; version 1 facts unchanged.' },
    @{ Case_ID = 'EX-C04'; Case_Version = 2; Received_Date = '2026-07-20'; Setting = 'Primary care and emergency department (fictional)'; Fictional = 'Yes'; Summary = 'Nausea and symptomatic bradycardia on digoxin after renal deterioration (Chapter 11, Section 11.8; Chapter 23, Section 23.11).'; Version_Note = 'Version 2 (GP letter received 20 July 2026) adds dose, last dose time, ibuprofen, creatinine series, digoxin level with sample time and IV fluids.' },
    @{ Case_ID = 'EX-C28'; Case_Version = 2; Received_Date = '2026-06-15'; Setting = 'Hematology and hepatology (fictional)'; Fictional = 'Yes'; Summary = 'HBV reactivation months after completing rituximab (Chapter 20, Section 20.4; Chapter 23, Section 23.12).'; Version_Note = 'Version 2 (received 15 June 2026) adds cycle dates, serology, prophylaxis, HBV DNA monitoring, prednisolone course and B-cell/IgG results.' }
)
$CasesVec = @()
foreach ($v in @('V01', 'V02', 'V03', 'V04', 'V05', 'V06', 'V07', 'V08', 'V09', 'V10', 'V11', 'V12', 'V13', 'V15', 'V16', 'V17', 'V18', 'V19', 'V20', 'V21', 'V22', 'V23', 'V24', 'V25', 'V26', 'V27')) {
    $CasesVec += @{ Case_ID = "VEC-$v"; Case_Version = 1; Fictional = 'Yes'; Summary = "CALC-SPEC 1.0 verification vector $v (tto_calc_test_vectors_v1.0.json)"; Version_Note = 'Vector inputs as in the JSON file' }
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
$V = 'tto_calc_test_vectors_v1.0.json'
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
    (VX 'VEC-V26' 1 'EX1' @{ Notes = 'Administration dates not part of vector V26' })
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
    (MergeH @((CV 'VEC-V26' 1), @{ Admin_ID = 'D3'; Exposure_Episode_ID = 'EX1'; Kind = 'DOSE'; Dose = 90; Dose_Unit = 'mg'; BSA_m2 = 1.70 }))
)

# ---------------- Events
$EvtEx = @(
    (MergeH @((CV 'EX-C01' 2), @{ Event_ID = 'E1'; Episode_ID = 'EV-1'; Event_Term = 'Urticaria, wheeze and hypotension (acute infusion-associated reaction)'; Baseline_Status = '08:50 observations stable; skin clear' },
        (TS 'O' '2026-01-10' '09:15' '-05:00' '09:15 pt reports itching; urticaria chest and arms; audible wheeze; BP 84/52' 'Nursing observation chart'))),
    (MergeH @((CV 'EX-C04' 2), @{ Event_ID = 'E1'; Episode_ID = 'EV-1'; Event_Term = 'Nausea and symptomatic bradycardia'; Baseline_Status = 'Six months of tolerated treatment' },
        (DT 'O' '2026-07-05' 'Nausea since this morning, dizzy, pulse 38' 'Emergency department note'))),
    (MergeH @((CV 'EX-C28' 2), @{ Event_ID = 'E1'; Episode_ID = 'EV-HBV'; Event_Term = 'Hepatitis B virus reactivation'; Baseline_Status = '11/11/2024 HBsAg negative, anti-HBc positive, HBV DNA not detected' },
        (RG 'O' '2026-02-09' '2026-05-12' 'Not before the last undetectable HBV DNA (9 Feb 2026), not after the first positive (12 May 2026); reviewer-entered bounds' 'Test records M-HBV3 and M-HBV4 (hepatology letter)'),
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
    (VE 'VEC-V25' 1 'EV1' (TS 'O' '2026-01-10' '09:50' '-05:00' 'symptoms begin 09:50' "$V V25"))
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
foreach ($v in @('V01', 'V02', 'V03', 'V04', 'V05', 'V06', 'V07', 'V08', 'V09', 'V10', 'V11', 'V15')) { $LnkVec += (LK "VEC-$v" 1 'L1' 'EX1' 'EV1' "Vector $v pair") }
$LnkVec += (LK 'VEC-V12' 1 'L1' 'EX1' 'EV1' 'Vector V12 pair (version 1)')
$LnkVec += (LK 'VEC-V12' 2 'L1' 'EX1' 'EV1' 'Vector V12 pair (version 2)')
$LnkVec += (LK 'VEC-V13' 1 'L1' 'EX2' 'EV2' 'Vector V13 reversed-bounds pair')
$LnkVec += (LK 'VEC-V13' 1 'L2' 'EX1' 'EV2' 'Vector V13 duplicate-ID pair')
$LnkVec += (LK 'VEC-V13' 1 'L3' 'EX3' 'EV1' 'Vector V13 invalid-date pair')

# ---------------- Intervals
$IntEx = @(
    (IVR 'EX-C01' 2 'I01' 'L1' 'FIRST_ADMIN' '' 'EX-A/START' 'EV-1/ONSET' 'C1'),
    (IVR 'EX-C01' 2 'I02' 'L1' 'FIRST_ADMIN' '' 'EX-A/START' 'EV-1/ONSET' 'C2' @{ Treatment_Day_Convention = 'inclusive'; Reviewer_Note = 'Date-level result; C1 decides the order because both times are known' }),
    (IVR 'EX-C01' 2 'I03' 'L1' 'LAST_DOSE_BEFORE_ONSET' '' '' 'EV-1/ONSET' 'C1' @{ Reviewer_Note = 'The infusion running at onset; its 09:16 end belongs to the chronology, not the anchor' }),
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
    (IVR 'VEC-V15' 1 'I1' 'L1' 'FIRST_ADMIN' '' 'EX1/START' 'EV1/ONSET' 'C2' @{ Reported_Text = 'about one month after starting' })
)

# ================================================================== INTERVALS COLUMNS
$E5 = '(nADM_SelKey={CaseKey}&"|"&{Exposure_Episode_ID})*nADM_Valid*(nADM_A_HT*{O_HT}*nADM_A_OK*{O_OK}*(nADM_UTC_n<=N({O_UTC}))+nADM_A_HT*{O_HT}*(1-nADM_A_OK*{O_OK})*(nADM_Loc_n<=N({O_Loc}))+(1-nADM_A_HT*{O_HT})*(nADM_Lo_n<=N({O_Lo})))'
function LkA([string]$n, [string]$tp, [string]$def = '""') { CF "A_$n" "Anchor: $n" ('=IF({A_Row}=0,' + $def + ',INDEX(nTP_' + $tp + ',{A_Row}))') 8 }
function LkO([string]$n, [string]$tp, [string]$def = '""') { CF "O_$n" "End point: $n" ('=IF({O_Row}=0,' + $def + ',INDEX(nTP_' + $tp + ',{O_Row}))') 8 }

$IntMsg = '=IF({RowEmpty}=1,"",MID(' +
    'IF({K_Err}<>""," | "&{K_Err}&": "&' + (MSGX '{K_Err}' '{K_Needs}') + ',"")' +
    '&IF({f_TZ}=1," | W_TZ_UNKNOWN: "&' + (MSGX '"W_TZ_UNKNOWN"' '""') + ',"")' +
    '&IF({f_SameDate}=1," | W_SAME_DATE_ORDER_UNKNOWN: "&' + (MSGX '"W_SAME_DATE_ORDER_UNKNOWN"' '""') + ',"")' +
    '&IF({f_SeqUnc}=1," | W_SEQUENCE_UNCERTAIN: "&' + (MSGX '"W_SEQUENCE_UNCERTAIN"' '""') + ',"")' +
    '&IF({f_AdminSame}=1," | W_ADMIN_SAME_DATE_AS_ONSET: "&' + (MSGX '"W_ADMIN_SAME_DATE_AS_ONSET"' '""') + ',"")' +
    '&IF({f_Interv}=1," | W_INTERVENTION_BEFORE_RECOVERY: "&' + (MSGX '"W_INTERVENTION_BEFORE_RECOVERY"' '""') + ',"")' +
    '&IF(OR({C6_Code}="",{C6_Code}="OVERLAP"),""," | "&{C6_Code}&" (reported timing): "&' + (MSGX '{C6_Code}' '"numeric reported bounds and a reported unit"') + ')' +
    '&IF({f_TDundef}=1," | I_TREATMENT_DAY_UNDEFINED: "&' + (MSGX '"I_TREATMENT_DAY_UNDEFINED"' '""') + ',"")' +
    '&IF({K_CErr}<>""," | "&{K_CErr}&" (constraint): "&' + (MSGX '{K_CErr}' '"a source for the constraint"') + ',""),4,9999))'

$IntCols = @(
    (CIn 'Interval_ID' 'Interval_ID (unique in the case version)' 10 $null '@' 'INPUTS (yellow): one row per link and anchor'),
    (CIn 'Case_ID' 'Case_ID' 10 $null '@'),
    (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
    (CIn 'Link_ID' 'Link_ID (Drug_Event_Links)' 8 $null '@'),
    (CIn 'Anchor_Type' 'Anchor_Type (C5)' 16 'list:nLIST_AnchorType'),
    (CIn 'Anchor_Description' 'Anchor description (required for OTHER)' 22 $null '@'),
    (CIn 'Anchor_Record_ID' 'Anchor_Record_ID (see Timepoints; leave blank for LAST_DOSE_BEFORE_ONSET, selected by C5)' 16 $null '@'),
    (CIn 'End_Record_ID' 'End-point record ID (e.g. EV-1/ONSET, EV-1/DETECTION or a RECOVERY milestone)' 16 $null '@'),
    (CIn 'Calculation_Mode' 'Calculation mode: C1 (timestamps), C2 (dates), C4 (bounds)' 9 'list:nLIST_Mode'),
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

    (CF 'Result_Display' 'Result (display per CALC-SPEC Section 7; E-code and no number when blocked)' ('=IF({RowEmpty}=1,"",IF({K_Err}<>"",{K_Err}&": no result is shown",IF({Calculation_Mode}="C1",FIXED({K_Min},0,FALSE)&" "&{Unit}&IF(ABS({K_Min})>=60," ("&FIXED({K_Min}/60,1,FALSE)&" hours)",""),IF({Calculation_Mode}="C2",FIXED({K_Days},0,FALSE)&" calendar days"&IF(AND({K_Days}=0,{K_Min}<>"")," (order within the day from C1: "&{Sequence_Status}&")","")&IF({Treatment_Day}<>"","; treatment Day "&{Treatment_Day},""),FIXED({K_Min4},0,FALSE)&" to "&FIXED({K_Max4},0,FALSE)&" calendar days"&IF({Constrained_Minimum}<>"","; with constraint "&FIXED({Constrained_Minimum},0,FALSE)&" to "&FIXED({Constrained_Maximum},0,FALSE),"")))))') 34 $null 'RESULT (grey formulas, locked)'),
    (CF 'Interval_Value' 'Interval_Value (unrounded; C1 minutes or C2 calendar days)' '=IF({K_Err}<>"","",IF({Calculation_Mode}="C1",{K_Min},IF({Calculation_Mode}="C2",{K_Days},"")))' 10),
    (CF 'Interval_Minimum' 'Interval_Minimum (C4, calendar days)' '=IF({K_Min4}="","",{K_Min4})' 9),
    (CF 'Interval_Maximum' 'Interval_Maximum (C4, calendar days)' '=IF({K_Max4}="","",{K_Max4})' 9),
    (CF 'Unit' 'Unit' '=IF(OR({RowEmpty}=1,{K_Err}<>""),"",IF({Calculation_Mode}="C1",IF({K_TZU}=1,"minutes (local clock)","minutes"),"calendar days"))' 12),
    (CF 'Calculation_Mode_Used' 'Calculation_Mode (rule applied)' '=IF({RowEmpty}=1,"",IF({f_C7}=1,"C7 ("&{Calculation_Mode}&")",IF(AND({Calculation_Mode}="C2",{Treatment_Day_Convention}="inclusive"),"C2+C3",{Calculation_Mode})))' 9),
    (CF 'Sequence_Status' 'Sequence_Status' ('=IF(OR({RowEmpty}=1,{K_Err}<>""),"",IF({Calculation_Mode}="C1",' + (SEQ '{K_Min}') + ',IF({Calculation_Mode}="C2",IF({K_Days}>0,"AFTER",IF({K_Days}<0,"BEFORE",IF({K_Min}<>"",' + (SEQ '{K_Min}') + ',"SAME_DATE_ORDER_UNKNOWN"))),IF({K_Min4}>0,"AFTER",IF({K_Max4}<0,"BEFORE",IF(AND({K_Min4}=0,{K_Max4}=0),"SAME_DATE_ORDER_UNKNOWN","SEQUENCE_UNCERTAIN"))))))') 16),
    (CF 'Treatment_Day' 'Treatment day number (C3, inclusive)' '=IF(AND({K_Err}="",{Calculation_Mode}="C2",{Treatment_Day_Convention}="inclusive"),IF({K_Days}>=0,{K_Days}+1,""),"")' 8),
    (CF 'QC_Flags' 'QC_Flags (E-, W-, I-codes)' '=IF({RowEmpty}=1,"",IF({K_Err}<>"",{K_Err}&IF({C6_Code}="I_APPROXIMATE_NOT_COMPARED","; I_APPROXIMATE_NOT_COMPARED",""),MID(IF({f_TZ}=1,"; W_TZ_UNKNOWN","")&IF({f_SameDate}=1,"; W_SAME_DATE_ORDER_UNKNOWN","")&IF({f_SeqUnc}=1,"; W_SEQUENCE_UNCERTAIN","")&IF({f_AdminSame}=1,"; W_ADMIN_SAME_DATE_AS_ONSET","")&IF({f_Interv}=1,"; W_INTERVENTION_BEFORE_RECOVERY","")&IF(OR({C6_Code}="",{C6_Code}="OVERLAP"),"","; "&{C6_Code})&IF({f_TDundef}=1,"; I_TREATMENT_DAY_UNDEFINED","")&IF({K_CErr}<>"","; "&{K_CErr}&" (constraint)",""),3,999)))' 22),
    (CF 'Messages' 'Messages (CALC-SPEC Section 5, verbatim)' $IntMsg 40),
    (CF 'Constrained_Minimum' 'C4 with constraint: minimum' '=IF(OR({K_CLo}="",{K_CErr}<>""),"",{O_Lo}-{K_CHi})' 9),
    (CF 'Constrained_Maximum' 'C4 with constraint: maximum' '=IF(OR({K_CLo}="",{K_CErr}<>""),"",{O_Hi}-{K_CLo})' 9),
    (CF 'Reported_Timing_Result' 'Reported timing (C6): kept beside the computed value' '=IF({C6_Code}="","",IF({C6_Code}="OVERLAP","Overlaps the computed interval. ",IF({C6_Code}="W_REPORTED_TIMING_CONFLICT","Conflicts with the computed interval; both kept. ",IF({C6_Code}="I_APPROXIMATE_NOT_COMPARED","Not compared. ",{C6_Code}&": not compared. ")))&"Reported: "&CHAR(34)&{Reported_Text}&CHAR(34)&IF(AND(ISNUMBER({Reported_Lower}),ISNUMBER({Reported_Upper}))," ["&{Reported_Lower}&" to "&{Reported_Upper}&" "&{Reported_Unit}&"]",""))' 30),
    (CF 'Anchor_Record_ID_Used' 'Anchor_Record_ID used' '=IF({Anchor_Type}="LAST_DOSE_BEFORE_ONSET",{C5_Selected},{Anchor_Record_ID}&"")' 13),
    (CF 'Anchor_ISO' 'Anchor (ISO 8601)' '=IF({A_Row}=0,"",INDEX(nTP_ISO,{A_Row}))' 20),
    (CF 'Anchor_Display' 'Anchor (as recorded)' '=IF({A_Row}=0,"",INDEX(nTP_Display,{A_Row}))' 28),
    (CF 'End_ISO' 'End point (ISO 8601)' '=IF({O_Row}=0,"",INDEX(nTP_ISO,{O_Row}))' 20),
    (CF 'End_Display' 'End point (as recorded)' '=IF({O_Row}=0,"",INDEX(nTP_Display,{O_Row}))' 28),
    (CF 'Exposure_Episode_ID' 'Exposure_Episode_ID (from link)' '=IF({LinkRow}=0,"",INDEX(nLNK_Exposure_Episode_ID,{LinkRow})&"")' 10),
    (CF 'Event_Episode_ID' 'Event Episode_ID (from link)' '=IF({LinkRow}=0,"",INDEX(nLNK_Event_Episode_ID,{LinkRow})&"")' 10),
    (CF 'Assumptions' 'Assumptions' '=IF(OR({RowEmpty}=1,{K_Err}<>""),"",IF({Calculation_Mode}="C1",IF({K_TZU}=1,"Local clock times subtracted; the UTC offset of at least one time is unknown.","Both times converted to UTC with the recorded offsets."),IF({Calculation_Mode}="C2","Calendar-date difference of the recorded local dates; not elapsed days."&IF({Treatment_Day_Convention}="inclusive"," Treatment day uses the inclusive convention selected by the user; never written to E2B fields automatically.",""),"Bounded interval: minimum = end-point lower bound minus anchor upper bound; maximum = end-point upper bound minus anchor lower bound. No single date is imputed."))&IF({Anchor_Type}="LAST_DOSE_BEFORE_ONSET"," Anchor selected by C5: latest administration at or before the end point ("&{Anchor_Record_ID_Used}&").",""))' 40),
    (CF 'C7_Note' 'C7 note (not a CALC-SPEC code)' '=IF(AND({RowEmpty}=0,{f_C7}=1,{K_Err}="",{IntervCount}=0),"No intervention is recorded between cessation and recovery. The absence of a record is not evidence that none was given; CALC-SPEC 1.0 has no code for interventions unknown.","")' 24),
    (CF 'Calculation_Version' 'Calculation_Version' '=IF({RowEmpty}=1,"","CALC-SPEC 1.0")' 11),

    (CF 'RowEmpty' 'work: row empty' '=IF(AND({Interval_ID}="",{Link_ID}="",{Anchor_Type}="",{End_Record_ID}=""),1,0)' 6 '0' 'WORKING COLUMNS (locked; shown for inspection)'),
    (CF 'CaseKey' 'work: case key' '={Case_ID}&"|"&{Case_Version}' 10),
    (CF 'Key' 'work: interval key' '=IF({Interval_ID}="","",{CaseKey}&"|"&{Interval_ID})' 14),
    (CF 'Dup' 'work: key count' '=IF({Key}="",0,COUNTIF(nINT_Key,{Key}))' 5 '0'),
    (CF 'LinkRow' 'work: link row' '=IF({Link_ID}="",0,IFERROR(MATCH({CaseKey}&"|"&{Link_ID},nLNK_Key,0),0))' 5 '0'),
    (CF 'O_Row' 'work: end-point row in Timepoints' '=IF(OR({RowEmpty}=1,{End_Record_ID}=""),0,IFERROR(MATCH({CaseKey}&"|"&{End_Record_ID},nTP_Key,0),0))' 6 '0'),
    (LkO 'Dup' 'Dup' '0'), (LkO 'Type' 'Record_Type'), (LkO 'Prec' 'Precision'), (LkO 'Err' 'Err'),
    (LkO 'HT' 'HT' '0'), (LkO 'OK' 'OK' '0'), (LkO 'Loc' 'Loc'), (LkO 'UTC' 'UTC'), (LkO 'Lo' 'Lo'), (LkO 'Hi' 'Hi'),
    (CF 'C5_MaxKey' 'work: C5 latest eligible administration (sort key)' ('=IF(OR({RowEmpty}=1,{Anchor_Type}<>"LAST_DOSE_BEFORE_ONSET",{LinkRow}=0,{O_Row}=0),0,IF({O_Err}<>"",0,SUMPRODUCT(MAX(' + $E5 + '*nADM_Sort_n))))') 10 '0'),
    (CF 'C5_Row' 'work: C5 row' ('=IF({C5_MaxKey}=0,0,MATCH(1,INDEX(' + $E5 + '*(nADM_Sort_n={C5_MaxKey}),0),0))') 5 '0'),
    (CF 'C5_Selected' 'work: C5 selected administration' '=IF({C5_Row}=0,"",INDEX(nADM_Admin_ID,{C5_Row})&"")' 10),
    (CF 'A_Row' 'work: anchor row in Timepoints' '=IF(OR({RowEmpty}=1,{Anchor_Record_ID_Used}=""),0,IFERROR(MATCH({CaseKey}&"|"&{Anchor_Record_ID_Used},nTP_Key,0),0))' 6 '0'),
    (LkA 'Dup' 'Dup' '0'), (LkA 'Type' 'Record_Type'), (LkA 'Prec' 'Precision'), (LkA 'Err' 'Err'),
    (LkA 'HT' 'HT' '0'), (LkA 'OK' 'OK' '0'), (LkA 'Loc' 'Loc'), (LkA 'UTC' 'UTC'), (LkA 'Lo' 'Lo'), (LkA 'Hi' 'Hi'),
    (CF 'IntervCount' 'work: interventions between cessation and recovery (C7)' '=IF(OR({A_Row}=0,{O_Row}=0),0,SUMPRODUCT((nTP_Case_Key={CaseKey})*(nTP_Record_Type="MILESTONE:INTERVENTION")*(nTP_Err="")*(nTP_HT*{A_HT}*{O_HT}*(nTP_Loc_n>=N({A_Loc}))*(nTP_Loc_n<=N({O_Loc}))+(1-nTP_HT*{A_HT}*{O_HT})*(nTP_Lo_n>=N({A_Lo}))*(nTP_Hi_n<=N({O_Hi}))*(nTP_Lo_n>0))))' 6 '0'),
    (CF 'Pre_Err' 'work: link/record check' '=IF({RowEmpty}=1,"",IF({Interval_ID}="","E_MISSING_INPUT",IF({Dup}>1,"E_DUPLICATE_ID",IF({LinkRow}=0,"E_MISSING_LINK",IF(OR({Anchor_Type}="",AND({Anchor_Type}="OTHER",{Anchor_Description}="")),"E_MISSING_INPUT",IF({O_Row}=0,"E_MISSING_INPUT",IF({A_Row}=0,"E_MISSING_INPUT",IF(OR({A_Dup}>1,{O_Dup}>1),"E_DUPLICATE_ID",""))))))))' 12),
    (CF 'Pre_Needs' 'work: missing input named' '=IF({Pre_Err}<>"E_MISSING_INPUT","",IF({Interval_ID}="","an Interval_ID",IF({Anchor_Type}="","an Anchor_Type",IF(AND({Anchor_Type}="OTHER",{Anchor_Description}=""),"a description of the OTHER anchor",IF({O_Row}=0,"the end-point record "&IF({End_Record_ID}="","(End_Record_ID is blank)",{End_Record_ID}),IF({Anchor_Type}="LAST_DOSE_BEFORE_ONSET","an administration of exposure episode "&{Exposure_Episode_ID}&" dated at or before the end point (C5)","the anchor record "&IF({Anchor_Record_ID}="","(Anchor_Record_ID is blank)",{Anchor_Record_ID})))))))' 14),
    (CF 'K_Err' 'work: blocking E-code' '=IF({RowEmpty}=1,"",IF({Pre_Err}<>"",{Pre_Err},IF({A_Err}<>"",{A_Err},IF({O_Err}<>"",{O_Err},IF(AND({Calculation_Mode}<>"C1",{Calculation_Mode}<>"C2",{Calculation_Mode}<>"C4"),"E_MISSING_INPUT",IF(OR({A_Lo}="",{O_Lo}=""),"E_MISSING_INPUT",IF(AND({Calculation_Mode}="C1",{A_HT}*{O_HT}=0),"E_MISSING_INPUT",IF(AND({Calculation_Mode}="C2",OR(AND({A_Prec}<>"TIMESTAMP",{A_Prec}<>"DATE"),AND({O_Prec}<>"TIMESTAMP",{O_Prec}<>"DATE"))),"E_MISSING_INPUT",""))))))))' 12),
    (CF 'K_Needs' 'work: [input] for E_MISSING_INPUT' '=IF({K_Err}<>"E_MISSING_INPUT","",IF({Pre_Err}="E_MISSING_INPUT",{Pre_Needs},IF({A_Err}="E_MISSING_INPUT","the date fields required by the precision of anchor record "&{Anchor_Record_ID_Used},IF({O_Err}="E_MISSING_INPUT","the date fields required by the precision of end-point record "&{End_Record_ID},IF(AND({Calculation_Mode}<>"C1",{Calculation_Mode}<>"C2",{Calculation_Mode}<>"C4"),"a calculation mode (C1, C2 or C4)",IF(OR({A_Lo}="",{O_Lo}=""),"dated anchor and end-point records",IF({Calculation_Mode}="C1","a time of day and UTC offset (or unknown) on both records; C1 needs two timestamps, so use C2 or C4 for dates","an exact date on both records; use C4 for partial dates and ranges")))))))' 14),
    (CF 'K_TZU' 'work: offset unknown (1/0)' '=IF({A_OK}*{O_OK}=1,0,1)' 5 '0'),
    (CF 'K_Min' 'work: C1 minutes' '=IF(AND({K_Err}="",{A_HT}*{O_HT}=1),IF({K_TZU}=0,{O_UTC}-{A_UTC},{O_Loc}-{A_Loc}),"")' 8),
    (CF 'K_Days' 'work: C2 calendar days' '=IF(AND({K_Err}="",{Calculation_Mode}="C2"),{O_Lo}-{A_Lo},"")' 7),
    (CF 'K_Min4' 'work: C4 minimum' '=IF(AND({K_Err}="",{Calculation_Mode}="C4"),{O_Lo}-{A_Hi},"")' 7),
    (CF 'K_Max4' 'work: C4 maximum' '=IF(AND({K_Err}="",{Calculation_Mode}="C4"),{O_Hi}-{A_Lo},"")' 7),
    (CF 'K_HasC' 'work: constraint entered' '=IF(AND({Calculation_Mode}="C4",OR(ISNUMBER({Constraint_Not_Before}),ISNUMBER({Constraint_Not_After}))),1,0)' 5 '0'),
    (CF 'K_CLo' 'work: constrained anchor lower' '=IF(OR({K_Err}<>"",{K_HasC}=0),"",IF(ISNUMBER({Constraint_Not_Before}),MAX({A_Lo},{Constraint_Not_Before}),{A_Lo}))' 10 'd mmmm yyyy'),
    (CF 'K_CHi' 'work: constrained anchor upper' '=IF(OR({K_Err}<>"",{K_HasC}=0),"",IF(ISNUMBER({Constraint_Not_After}),MIN({A_Hi},{Constraint_Not_After}),{A_Hi}))' 10 'd mmmm yyyy'),
    (CF 'K_CErr' 'work: constraint check' '=IF({K_CLo}="","",IF({K_CLo}>{K_CHi},"E_REVERSED_BOUNDS",IF({Constraint_Source}="","E_MISSING_INPUT","")))' 10),
    (CF 'f_TZ' 'flag W_TZ_UNKNOWN' '=IF(AND({K_Err}="",{K_TZU}=1,{K_Min}<>"",OR({Calculation_Mode}="C1",AND({Calculation_Mode}="C2",{K_Days}=0))),1,0)' 4 '0'),
    (CF 'f_SameDate' 'flag W_SAME_DATE_ORDER_UNKNOWN' '=IF(AND({K_Err}="",OR(AND({Calculation_Mode}="C2",{K_Days}=0,{K_Min}=""),AND({Calculation_Mode}="C4",{K_Min4}=0,{K_Max4}=0))),1,0)' 4 '0'),
    (CF 'f_SeqUnc' 'flag W_SEQUENCE_UNCERTAIN' '=IF({Sequence_Status}="SEQUENCE_UNCERTAIN",1,0)' 4 '0'),
    (CF 'f_AdminSame' 'flag W_ADMIN_SAME_DATE_AS_ONSET' '=IF(AND({K_Err}="",{Anchor_Type}="LAST_DOSE_BEFORE_ONSET",{A_HT}*{O_HT}=0,{A_Lo}={O_Lo}),1,0)' 4 '0'),
    (CF 'f_TDundef' 'flag I_TREATMENT_DAY_UNDEFINED' '=IF(AND({K_Err}="",{Calculation_Mode}="C2",{Treatment_Day_Convention}="inclusive",{K_Days}<0),1,0)' 4 '0'),
    (CF 'f_C7' 'C7 recovery interval (CESSATION to RECOVERY)' '=IF(AND({Anchor_Type}="CESSATION",{O_Type}="MILESTONE:RECOVERY"),1,0)' 4 '0'),
    (CF 'f_Interv' 'flag W_INTERVENTION_BEFORE_RECOVERY' '=IF(AND({K_Err}="",{f_C7}=1,{IntervCount}>0),1,0)' 4 '0'),
    (CF 'C6_Factor' 'work: C6 unit factor (0 = not comparable)' '=IF({Calculation_Mode}="C1",IF({Reported_Unit}="minutes",1,IF({Reported_Unit}="hours",60,IF({Reported_Unit}="days",1440,0))),IF(OR({Calculation_Mode}="C2",{Calculation_Mode}="C4"),IF({Reported_Unit}="calendar days",1,0),0))' 6 '0'),
    (CF 'C6_Code' 'work: C6 result code' '=IF(COUNTA({Reported_Text}:{Reported_Unit})=0,"",IF(OR({Reported_Lower}="",{Reported_Upper}=""),"I_APPROXIMATE_NOT_COMPARED",IF(OR(NOT(ISNUMBER({Reported_Lower})),NOT(ISNUMBER({Reported_Upper}))),"E_MISSING_INPUT",IF({Reported_Lower}>{Reported_Upper},"E_REVERSED_BOUNDS",IF(OR({Reported_Unit}="weeks",{Reported_Unit}="months",{Reported_Unit}="years"),"I_APPROXIMATE_NOT_COMPARED",IF({Reported_Unit}="","E_MISSING_INPUT",IF({K_Err}<>"","",IF({C6_Factor}=0,"I_APPROXIMATE_NOT_COMPARED",IF(AND({Reported_Lower}*{C6_Factor}<=IF({Calculation_Mode}="C4",{K_Max4},{Interval_Value}),{Reported_Upper}*{C6_Factor}>=IF({Calculation_Mode}="C4",{K_Min4},{Interval_Value})),"OVERLAP","W_REPORTED_TIMING_CONFLICT")))))))))' 14)
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
$SegErr = '=IF(COUNTA({End_Date}:{End_Off})=0,"",IF(OR({End_Date}="",{End_Time}="",{End_Off}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({End_Date})),NOT(ISNUMBER({End_Time}))),"E_INVALID_DATE",IF(OR({End_Date}<>INT({End_Date}),{End_Time}<0,{End_Time}>=1),"E_INVALID_DATE",IF(ISNA(MATCH({End_Off},nLIST_Offset,0)),"E_INVALID_UNIT","")))))'
$AdmCols = @(
    (CIn 'Case_ID' 'Case_ID' 10 $null '@' 'INPUTS: one row per dose, administration or infusion rate segment'),
    (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'),
    (CIn 'Admin_ID' 'Admin_ID (record ID)' 12 $null '@'),
    (CIn 'Exposure_Episode_ID' 'Exposure_Episode_ID' 11 $null '@'),
    (CIn 'Kind' 'Kind (DOSE, INFUSION_SEGMENT, OTHER)' 13 'list:nLIST_AdminKind'),
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
    (CF 'Valid' 'Exact date, no error (C5 eligible)' '=IF(AND({Admin_ID}<>"",{A_Err}="",{A_Lo}<>"",{A_Lo}={A_Hi}),1,0)' 5 '0'),
    (CF 'Loc_n' 'Local minutes (0 if none)' '=IF({A_HT}=1,{A_Loc},0)' 9 '0'),
    (CF 'UTC_n' 'UTC minutes (0 if none)' '=IF({A_OK}=1,{A_UTC},0)' 9 '0'),
    (CF 'Lo_n' 'Date (0 if none)' '=IF({Valid}=1,{A_Lo},0)' 7 '0'),
    (CF 'Sort_n' 'Order key (UTC, else local, else date)' '=IF({Valid}=1,IF({A_OK}=1,{A_UTC},IF({A_HT}=1,{A_Loc},{A_Lo}*1440)),0)' 9 '0'),
    (CF 'Rate_n' 'Rate (0 if none)' '=IF(ISNUMBER({Rate}),{Rate},0)' 6),
    (CF 'Dose_n' 'Dose (0 if none)' '=IF(ISNUMBER({Dose}),{Dose},0)' 6),
    (CF 'HasDose' 'Has dose' '=IF(AND({Admin_ID}<>"",ISNUMBER({Dose})),1,0)' 5 '0'),
    (CF 'BSA_n' 'BSA (0 if none)' '=IF(ISNUMBER({BSA_m2}),{BSA_m2},0)' 6),
    (CF 'HasBSA' 'Has BSA' '=IF(AND({Admin_ID}<>"",ISNUMBER({BSA_m2})),1,0)' 5 '0'),
    (CF 'BSA_div' 'BSA divisor (1 if none)' '=IF(AND(ISNUMBER({BSA_m2}),{BSA_m2}>0),{BSA_m2},1)' 6),
    (CF 'Dose_per_BSA' 'Dose / BSA, mg/m2 (0 if either is missing)' '=IF(AND({HasDose}=1,{HasBSA}=1),{Dose_n}/{BSA_div},0)' 8),
    (CF 'SegOK' 'Segment: both offsets known' '=IF(AND({Kind}="INFUSION_SEGMENT",{A_OK}=1,{SE_OK}=1),1,0)' 5 '0'),
    (CF 'SegS_UTC_n' 'Segment start UTC' '=IF({SegOK}=1,{A_UTC},0)' 9 '0'),
    (CF 'SegE_UTC_n' 'Segment end UTC' '=IF({SegOK}=1,{SE_UTC},0)' 9 '0'),
    (CF 'SegS_Loc_n' 'Segment start local' '=IF(AND({Kind}="INFUSION_SEGMENT",{A_HT}=1),{A_Loc},0)' 9 '0'),
    (CF 'SegE_Loc_n' 'Segment end local' '=IF(AND({Kind}="INFUSION_SEGMENT",{SE_Loc}<>""),{SE_Loc},0)' 9 '0'),
    (CF 'SegS_cmp' 'Segment start (comparison)' '=IF({SegOK}=1,{SegS_UTC_n},{SegS_Loc_n})' 9 '0'),
    (CF 'SegE_cmp' 'Segment end (comparison)' '=IF({SegOK}=1,{SegE_UTC_n},{SegE_Loc_n})' 9 '0'),
    (CF 'Overlap' 'Overlapping segments (count)' '=IF(OR({Kind}<>"INFUSION_SEGMENT",{SegS_cmp}=0,{SegE_cmp}=0),0,COUNTIFS(nADM_SelKey,{SelKey},nADM_Kind,"INFUSION_SEGMENT",nADM_SegS_cmp,"<"&{SegE_cmp},nADM_SegE_cmp,">"&{SegS_cmp})-IF({SegE_cmp}>{SegS_cmp},1,0))' 5 '0'),
    (CF 'Seg_Check' 'Segment check (Q9)' '=IF({Kind}<>"INFUSION_SEGMENT","",IF({A_Err}<>"",{A_Err},IF({SE_Err}<>"",{SE_Err},IF(OR({A_HT}=0,{SE_Loc}=""),"E_MISSING_INPUT",IF(NOT(ISNUMBER({Rate})),"E_MISSING_INPUT",IF({Rate_Unit}<>"mg/hour",IF({Rate_Unit}="","E_MISSING_INPUT","E_INVALID_UNIT"),IF({Rate}<=0,"E_NONPOSITIVE_PARAMETER",IF({SegE_cmp}<{SegS_cmp},"E_REVERSED_BOUNDS",IF({Overlap}>0,"E_OVERLAPPING_SEGMENTS","")))))))))' 14),
    (CF 'Ep_Found' 'Exposure episode exists' '=IF({Admin_ID}="",1,IF(COUNTIF(nEXP_Key,{SelKey})>0,1,0))' 5 '0'),
    (CF 'Record_Check' 'Record check (E-code)' '=IF({Admin_ID}="",IF(COUNTA({Exposure_Episode_ID},{Dose},{Rate},{A_Prec},{A_Date})>0,"E_MISSING_INPUT",""),IF({Dup}>1,"E_DUPLICATE_ID",IF({Ep_Found}=0,"E_MISSING_INPUT",IF({A_Err}<>"",{A_Err},IF({Seg_Check}<>"",{Seg_Check},IF(AND(ISNUMBER({Dose}),{Dose}<=0),"E_NONPOSITIVE_PARAMETER",IF(AND(ISNUMBER({BSA_m2}),{BSA_m2}<=0),"E_NONPOSITIVE_PARAMETER","")))))))' 14),
    (CF 'Check_Message' 'Check message' ('=IF({Record_Check}="","",{Record_Check}&": "&' + (MSGX '{Record_Check}' 'IF({Admin_ID}="","an Admin_ID",IF({Ep_Found}=0,"an Exposures row for episode "&{Exposure_Episode_ID},IF({Seg_Check}="E_MISSING_INPUT","segment start and end timestamps and a rate in mg/hour","the date fields required by the selected precision")))') + ')') 40)
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
) + @(TPBlock 'O' 'Onset') + @(TPBlock 'D' 'First detection') + @(
    (CF 'Key' 'Event key' '=IF({Episode_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Episode_ID})' 14 $null 'CHECKS'),
    (CF 'Dup' 'Key count' '=IF({Key}="",0,COUNTIF(nEVT_Key,{Key}))' 5 '0'),
    (CF 'Record_Check' 'Record check (E-code)' '=IF({Episode_ID}="",IF(COUNTA({Event_Term},{O_Prec},{O_Date})>0,"E_MISSING_INPUT",""),IF({Dup}>1,"E_DUPLICATE_ID",IF({O_Err}<>"",{O_Err},{D_Err})))' 14),
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
$GATE = 'AND({G_Model}="Yes",{G_Params}<>"",{G_RefTime}<>"",{G_Status}="Yes")'
function QHead() { @((CIn 'Calc_ID' 'Calc_ID' 14 $null '@' 'INPUTS'), (CIn 'Case_ID' 'Case_ID' 10 $null '@'), (CIn 'Case_Version' 'Case_Version' 7 'posint' '0'), (CIn 'Question' 'Question this calculation answers' 30 $null '@')) }
function QGate() {
    @((CIn 'G_Model' 'Gate 1: simple model applies (one compartment, linear first-order, constant parameters)? Yes/No' 11 'list:nLIST_YesNo' $null 'COMMON GATE (CALC-SPEC Section 4)'),
      (CIn 'G_Params' 'Gate 2: half-life or parameter used, with type, population and source' 26 $null '@'),
      (CIn 'G_RefTime' 'Gate 3: reference time (what the result is a fraction of, and when)' 24 $null '@'),
      (CIn 'G_Status' 'Gate 4: each input''s Parameter_Status recorded and confirmed? Yes/No' 11 'list:nLIST_YesNo'))
}
function Par($n, $h, $w = 9) { @((CIn $n $h $w 'num'), (CIn "${n}_Status" "$n status" 13 'list:nLIST_Status'), (CIn "${n}_Source" "$n source (document, version, section)" 18 $null '@')) }
function PSET($n, $u) { '"' + $n + ' = "&{' + $n + '}&" "&' + $u + '&" ["&IF({' + $n + '_Status}="","status not recorded",{' + $n + '_Status})&IF({' + $n + '_Source}<>"","; "&{' + $n + '_Source},"")&"]"' }
function QTail($mand, $words, $flagsF, $msgF, $psetF, [bool]$gate = $true) {
    $a = @(
        (CF 'QC_Flags' 'QC_Flags' $flagsF 20),
        (CF 'Messages' 'Messages (verbatim)' $msgF 40),
        (CF 'Parameter_Set' 'Parameter_Set (value, unit, status, source)' $psetF 40)
    )
    if ($gate) { $a += (CF 'Assumptions' 'Assumptions (gate confirmations)' '=IF({RowEmpty}=1,"","Model: "&IF({G_Model}="Yes","confirmed","not confirmed")&". Parameter used: "&IF({G_Params}="","not stated",{G_Params})&". Reference time: "&IF({G_RefTime}="","not stated",{G_RefTime})&". Parameter_Status: "&IF({G_Status}="Yes","confirmed","not confirmed")&".")' 36) }
    $a += (CF 'Mandatory_Text' 'Mandatory text (CALC-SPEC)' ('=IF({RowEmpty}=1,"","' + $mand + '")') 40)
    $a += (CF 'Formula_in_words' 'Formula in words' ('=IF({RowEmpty}=1,"","' + $words + '")') 30)
    $a += (CF 'Calculation_Version' 'Calculation_Version' '=IF({RowEmpty}=1,"","CALC-SPEC 1.0")' 11)
    return $a
}
function QWork($emptyF, $errF, $needsF) {
    @((CF 'RowEmpty' 'work: row empty' $emptyF 5 '0' 'WORKING'),
      (CF 'Key' 'work: key' '=IF({Calc_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Calc_ID})' 14),
      (CF 'Err' 'work: blocking E-code' $errF 14),
      (CF 'Needs' 'work: [input]' $needsF 14),
      (CF 'Ok' 'work: result shown (1/0)' '=IF(AND({RowEmpty}=0,{Err}=""),1,0)' 5 '0'))
}
$UNITOK = 'ISNA(MATCH({Time_Unit},nLIST_TimeUnit,0))'
$SIMPLEMSG = '=IF({Err}="","",{Err}&": "&' + (MSGX '{Err}' '{Needs}') + ')'
function TD([string]$x) { 'IF({Time_Unit}="minutes",FIXED(' + $x + ',0,TRUE),FIXED(' + $x + ',1,TRUE))' }

$Q1Cols = @(QHead) + @(QGate) + @(
    (CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS (same time unit)')) + @(Par 't_half' 't_half (in Time_Unit)') + @(Par 't' 't, time since reference (in Time_Unit)') + @(
    (CF 'Result_Display' 'Result (display: percent to 3 significant figures)' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown",' + (P3 '{Percent_Remaining}') + '&"% of the reference remains after "&{t}&" "&{Time_Unit}&"; k = "&' + (P3 '{k_per_unit}') + '&" per "&LEFT({Time_Unit},LEN({Time_Unit})-1)))') 34 $null 'RESULT'),
    (CF 'Fraction_Remaining' 'Fraction remaining (unrounded)' '=IF({Ok}=1,2^(-{t}/{t_half}),"")' 12),
    (CF 'Percent_Remaining' 'Percent remaining (unrounded)' '=IF({Ok}=1,100*{Fraction_Remaining},"")' 12),
    (CF 'k_per_unit' 'k (per Time_Unit, unrounded)' '=IF({Ok}=1,LN(2)/{t_half},"")' 12),
    (CF 'k_per_hour' 'k (per hour, unrounded)' '=IF({Ok}=1,LN(2)/({t_half}*IF({Time_Unit}="minutes",1/60,IF({Time_Unit}="hours",1,24))),"")' 12)
) + @(QTail 'Fraction of the chosen reference amount or concentration under the stated model. It is not a causal or pharmacological threshold.' 'k = ln 2 / t_half; fraction remaining = 2^(-t / t_half) = exp(-k x t)' '=IF({RowEmpty}=1,"",{Err})' $SIMPLEMSG ('=IF({RowEmpty}=1,"",' + (PSET 't_half' '{Time_Unit}') + '&"; "&' + (PSET 't' '{Time_Unit}') + ')')) + @(QWork '=IF(AND({Calc_ID}="",{t_half}="",{t}=""),1,0)' `
    ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Time_Unit}="",{t_half}="",{t}=""),"E_MISSING_INPUT",IF(' + $UNITOK + ',"E_INVALID_UNIT",IF(OR(NOT(ISNUMBER({t_half})),NOT(ISNUMBER({t}))),"E_MISSING_INPUT",IF({t_half}<=0,"E_NONPOSITIVE_PARAMETER",IF(NOT(' + $GATE + '),"E_APPLICABILITY_UNCONFIRMED",IF(OR({t_half_Status}="",{t_Status}=""),"E_MISSING_INPUT",IF({t}<0,"E_OUT_OF_RANGE",""))))))))') `
    '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Time_Unit}="","a time unit (minutes, hours or days)",IF(NOT(ISNUMBER({t_half})),"a numeric t_half",IF(NOT(ISNUMBER({t})),"a numeric t",IF({t_half_Status}="","the Parameter_Status of t_half","the Parameter_Status of t"))))))')

$Q2Cols = @(QHead) + @(QGate) + @(
    (CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 't_half' 't_half (in Time_Unit)') + @(Par 'f_target' 'Target fraction f of steady state (0 < f < 1)') + @(
    (CF 'Result_Display' 'Result (display)' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown",' + (TD '{Time_to_f}') + '&" "&{Time_Unit}&" to reach "&' + (P3 '100*{f_target}') + '&"% of steady state ("&FIXED({Half_Lives},2,TRUE)&" half-lives)"))') 34 $null 'RESULT'),
    (CF 'Half_Lives' 'Half-lives needed n (unrounded)' '=IF({Ok}=1,-LOG(1-{f_target},2),"")' 12),
    (CF 'Time_to_f' 'Time to reach f (in Time_Unit, unrounded)' '=IF({Ok}=1,{Half_Lives}*{t_half},"")' 12)
) + @(QTail 'Assumes constant input from zero. Intermittent dosing compares corresponding points in the dosing cycle. Toxicity need not wait for steady state.' 'n = -log2(1 - f); time = n x t_half' '=IF({RowEmpty}=1,"",{Err})' $SIMPLEMSG ('=IF({RowEmpty}=1,"",' + (PSET 't_half' '{Time_Unit}') + '&"; "&' + (PSET 'f_target' '"(fraction)"') + ')')) + @(QWork '=IF(AND({Calc_ID}="",{t_half}="",{f_target}=""),1,0)' `
    ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Time_Unit}="",{t_half}="",{f_target}=""),"E_MISSING_INPUT",IF(' + $UNITOK + ',"E_INVALID_UNIT",IF(OR(NOT(ISNUMBER({t_half})),NOT(ISNUMBER({f_target}))),"E_MISSING_INPUT",IF({t_half}<=0,"E_NONPOSITIVE_PARAMETER",IF(NOT(' + $GATE + '),"E_APPLICABILITY_UNCONFIRMED",IF(OR({t_half_Status}="",{f_target_Status}=""),"E_MISSING_INPUT",IF(OR({f_target}<=0,{f_target}>=1),"E_OUT_OF_RANGE",""))))))))') `
    '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Time_Unit}="","a time unit",IF(NOT(ISNUMBER({t_half})),"a numeric t_half",IF(NOT(ISNUMBER({f_target})),"a numeric target fraction",IF({t_half_Status}="","the Parameter_Status of t_half","the Parameter_Status of the target fraction"))))))')

$Q3Cols = @(QHead) + @(QGate) + @(
    (CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 't_half' 't_half (in Time_Unit)') + @(Par 'tau' 'Dosing interval tau (in Time_Unit)') + @(
    (CF 'Result_Display' 'Result (display: 2 decimals)' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","Accumulation factor R = "&FIXED({Accumulation_Factor},2,TRUE)&" at corresponding points in the dosing cycle"))' 34 $null 'RESULT'),
    (CF 'Accumulation_Factor' 'Accumulation factor R (unrounded)' '=IF({Ok}=1,1/(1-2^(-{tau}/{t_half})),"")' 12)
) + @(QTail 'Model-based repeated-dose comparison. It is not a toxicity multiplier and does not give the time of clinical onset.' 'R = 1 / (1 - 2^(-tau / t_half))' '=IF({RowEmpty}=1,"",{Err})' $SIMPLEMSG ('=IF({RowEmpty}=1,"",' + (PSET 't_half' '{Time_Unit}') + '&"; "&' + (PSET 'tau' '{Time_Unit}') + ')')) + @(QWork '=IF(AND({Calc_ID}="",{t_half}="",{tau}=""),1,0)' `
    ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Time_Unit}="",{t_half}="",{tau}=""),"E_MISSING_INPUT",IF(' + $UNITOK + ',"E_INVALID_UNIT",IF(OR(NOT(ISNUMBER({t_half})),NOT(ISNUMBER({tau}))),"E_MISSING_INPUT",IF(OR({t_half}<=0,{tau}<=0),"E_NONPOSITIVE_PARAMETER",IF(NOT(' + $GATE + '),"E_APPLICABILITY_UNCONFIRMED",IF(OR({t_half_Status}="",{tau_Status}=""),"E_MISSING_INPUT","")))))))') `
    '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Time_Unit}="","a time unit",IF(NOT(ISNUMBER({t_half})),"a numeric t_half",IF(NOT(ISNUMBER({tau})),"a numeric tau",IF({t_half_Status}="","the Parameter_Status of t_half","the Parameter_Status of tau"))))))')

$CLH = 'IF({CL_Unit}="mL/minute",0.06,1)'
$Q4Cols = @(QHead) + @(QGate) + @(
    (CIn 'CL_Unit' 'Clearance unit (L/hour or mL/minute; same for both)' 10 'list:nLIST_CLUnit' $null 'INPUTS')) + @(Par 'CL_before' 'Clearance before') + @(Par 'CL_after' 'Clearance after') + @(Par 'V_L' 'Apparent volume V (L), optional') + @(
    (CF 'Result_Display' 'Result (display: ratios 2 decimals)' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","Half-life ratio "&FIXED({HalfLife_Ratio},2,TRUE)&"; average steady-state concentration ratio "&FIXED({Css_Ratio},2,TRUE)&IF({t_half_before_h}<>""," (t_half "&FIXED({t_half_before_h},1,TRUE)&" hours before, "&FIXED({t_half_after_h},1,TRUE)&" hours after)","")))' 34 $null 'RESULT'),
    (CF 'HalfLife_Ratio' 'Half-life ratio after/before (unrounded)' '=IF({Ok}=1,{CL_before}/{CL_after},"")' 12),
    (CF 'Css_Ratio' 'Average Css ratio after/before (unrounded)' '=IF({Ok}=1,{CL_before}/{CL_after},"")' 12),
    (CF 't_half_before_h' 't_half before, hours (needs V)' ('=IF(OR({Ok}=0,{V_L}=""),"",LN(2)*{V_L}/({CL_before}*' + $CLH + '))') 12),
    (CF 't_half_after_h' 't_half after, hours (needs V)' ('=IF(OR({Ok}=0,{V_L}=""),"",LN(2)*{V_L}/({CL_after}*' + $CLH + '))') 12)
) + @(QTail 'Describes the eventual average steady state. It does not date the patient''s onset; the transition depends on prior concentrations and when clearance changed.' 'ratio = CL_before / CL_after (V constant); with V: t_half = ln 2 x V / CL for each period' '=IF({RowEmpty}=1,"",{Err})' $SIMPLEMSG ('=IF({RowEmpty}=1,"",' + (PSET 'CL_before' '{CL_Unit}') + '&"; "&' + (PSET 'CL_after' '{CL_Unit}') + '&IF({V_L}<>"","; "&' + (PSET 'V_L' '"L"') + ',""))')) + @(QWork '=IF(AND({Calc_ID}="",{CL_before}="",{CL_after}=""),1,0)' `
    ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{CL_Unit}="",{CL_before}="",{CL_after}=""),"E_MISSING_INPUT",IF(ISNA(MATCH({CL_Unit},nLIST_CLUnit,0)),"E_INVALID_UNIT",IF(OR(NOT(ISNUMBER({CL_before})),NOT(ISNUMBER({CL_after})),AND({V_L}<>"",NOT(ISNUMBER({V_L})))),"E_MISSING_INPUT",IF(OR({CL_before}<=0,{CL_after}<=0,AND(ISNUMBER({V_L}),{V_L}<=0)),"E_NONPOSITIVE_PARAMETER",IF(NOT(' + $GATE + '),"E_APPLICABILITY_UNCONFIRMED",IF(OR({CL_before_Status}="",{CL_after_Status}="",AND({V_L}<>"",{V_L_Status}="")),"E_MISSING_INPUT","")))))))') `
    '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({CL_Unit}="","a clearance unit",IF(NOT(ISNUMBER({CL_before})),"a numeric clearance before",IF(NOT(ISNUMBER({CL_after})),"a numeric clearance after",IF(AND({V_L}<>"",NOT(ISNUMBER({V_L}))),"a numeric volume","the Parameter_Status of each input"))))))')

$Q5Cols = @(QHead) + @(QGate) + @(
    (CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 'Parent_t_half' 'Parent t_half') + @(Par 'Metabolite_t_half' 'Metabolite t_half') + @(Par 't' 't, time since reference') + @(
    (CF 'Result_Display' 'Result (display; no sum is calculated)' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","Parent "&' + (P3 '{Parent_Percent}') + '&"% and metabolite "&' + (P3 '{Metabolite_Percent}') + '&"% remain, each of its own starting value"))') 34 $null 'RESULT'),
    (CF 'Parent_Percent' 'Parent percent remaining (unrounded)' '=IF({Ok}=1,100*2^(-{t}/{Parent_t_half}),"")' 12),
    (CF 'Metabolite_Percent' 'Metabolite percent remaining (unrounded)' '=IF({Ok}=1,100*2^(-{t}/{Metabolite_t_half}),"")' 12)
) + @(QTail 'Assumes no further formation or input. Continued metabolite formation, potency, distribution and nonlinear kinetics can make this calculation invalid.' 'Two separate Q1 calculations, each normalized to its own starting value; their sum is never calculated' '=IF({RowEmpty}=1,"",{Err})' $SIMPLEMSG ('=IF({RowEmpty}=1,"",' + (PSET 'Parent_t_half' '{Time_Unit}') + '&"; "&' + (PSET 'Metabolite_t_half' '{Time_Unit}') + '&"; "&' + (PSET 't' '{Time_Unit}') + ')')) + @(QWork '=IF(AND({Calc_ID}="",{Parent_t_half}="",{t}=""),1,0)' `
    ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Time_Unit}="",{Parent_t_half}="",{Metabolite_t_half}="",{t}=""),"E_MISSING_INPUT",IF(' + $UNITOK + ',"E_INVALID_UNIT",IF(OR(NOT(ISNUMBER({Parent_t_half})),NOT(ISNUMBER({Metabolite_t_half})),NOT(ISNUMBER({t}))),"E_MISSING_INPUT",IF(OR({Parent_t_half}<=0,{Metabolite_t_half}<=0),"E_NONPOSITIVE_PARAMETER",IF(NOT(' + $GATE + '),"E_APPLICABILITY_UNCONFIRMED",IF(OR({Parent_t_half_Status}="",{Metabolite_t_half_Status}="",{t_Status}=""),"E_MISSING_INPUT",IF({t}<0,"E_OUT_OF_RANGE",""))))))))') `
    '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Time_Unit}="","a time unit",IF(NOT(ISNUMBER({Parent_t_half})),"a numeric parent t_half",IF(NOT(ISNUMBER({Metabolite_t_half})),"a numeric metabolite t_half",IF(NOT(ISNUMBER({t})),"a numeric t","the Parameter_Status of each input"))))))')

$Q7Cols = @(QHead) + @(QGate) + @(
    (CIn 'Time_Unit' 'Time unit (days only for Q7)' 9 'list:nLIST_DaysOnly' $null 'INPUTS')) + @(Par 't_half' 't_half (days)') + @(Par 't' 't (days since reference)') + @(
    (CF 'Result_Display' 'Result (display: percent to 3 significant figures)' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown",' + (P3 '{Percent_Remaining}') + '&"% of the post-infusion concentration remains after "&{t}&" days (W_BIOLOGIC_MODEL_CAUTION)"))') 34 $null 'RESULT'),
    (CF 'Fraction_Remaining' 'Fraction remaining (unrounded)' '=IF({Ok}=1,2^(-{t}/{t_half}),"")' 12),
    (CF 'Percent_Remaining' 'Percent remaining (unrounded)' '=IF({Ok}=1,100*{Fraction_Remaining},"")' 12)
) + @(QTail 'Antibody disposition can depend on dose, time, target burden and immunogenicity. Exposure persistence is not receptor occupancy, immune activity or biological persistence.' 'As Q1: fraction remaining = 2^(-t / t_half), t and t_half in days' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err},"W_BIOLOGIC_MODEL_CAUTION"))' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": "&' + (MSGX '{Err}' '{Needs}') + ',"W_BIOLOGIC_MODEL_CAUTION: "&' + (MSGX '"W_BIOLOGIC_MODEL_CAUTION"' '""') + '))') ('=IF({RowEmpty}=1,"",' + (PSET 't_half' '"days"') + '&"; "&' + (PSET 't' '"days"') + ')')) + @(QWork '=IF(AND({Calc_ID}="",{t_half}="",{t}=""),1,0)' `
    ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Time_Unit}="",{t_half}="",{t}=""),"E_MISSING_INPUT",IF({Time_Unit}<>"days","E_INVALID_UNIT",IF(OR(NOT(ISNUMBER({t_half})),NOT(ISNUMBER({t}))),"E_MISSING_INPUT",IF({t_half}<=0,"E_NONPOSITIVE_PARAMETER",IF(NOT(' + $GATE + '),"E_APPLICABILITY_UNCONFIRMED",IF(OR({t_half_Status}="",{t_Status}=""),"E_MISSING_INPUT",IF({t}<0,"E_OUT_OF_RANGE",""))))))))') `
    '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Time_Unit}="","a time unit (days)",IF(NOT(ISNUMBER({t_half})),"a numeric t_half",IF(NOT(ISNUMBER({t})),"a numeric t",IF({t_half_Status}="","the Parameter_Status of t_half","the Parameter_Status of t"))))))')

$Q8Cols = @(QHead) + @(QGate) + @(
    (CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 't_half_lower' 'Lower t_half') + @(Par 't_half_upper' 'Upper t_half') + @(Par 't' 't, time since reference') + @(
    (CF 'Result_Display' 'Result: scenario range (never a midpoint or confidence interval)' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","From "&' + (P3 '{Percent_at_Lower}') + '&"% (t_half "&{t_half_lower}&" "&{Time_Unit}&") to "&' + (P3 '{Percent_at_Upper}') + '&"% (t_half "&{t_half_upper}&" "&{Time_Unit}&") remaining; scenario range, not a confidence interval"))') 40 $null 'RESULT'),
    (CF 'Fraction_at_Lower' 'Fraction at lower t_half (unrounded)' '=IF({Ok}=1,2^(-{t}/{t_half_lower}),"")' 12),
    (CF 'Fraction_at_Upper' 'Fraction at upper t_half (unrounded)' '=IF({Ok}=1,2^(-{t}/{t_half_upper}),"")' 12),
    (CF 'Percent_at_Lower' 'Percent at lower t_half (unrounded)' '=IF({Ok}=1,100*{Fraction_at_Lower},"")' 12),
    (CF 'Percent_at_Upper' 'Percent at upper t_half (unrounded)' '=IF({Ok}=1,100*{Fraction_at_Upper},"")' 12)
) + @(QTail 'Assumption-based scenario range, not a confidence interval.' 'Q1 evaluated at each half-life bound; no midpoint or average is calculated' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err},"I_SCENARIO_NOT_CI"))' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": "&' + (MSGX '{Err}' '{Needs}') + ',"I_SCENARIO_NOT_CI: "&' + (MSGX '"I_SCENARIO_NOT_CI"' '""') + '))') ('=IF({RowEmpty}=1,"",' + (PSET 't_half_lower' '{Time_Unit}') + '&"; "&' + (PSET 't_half_upper' '{Time_Unit}') + '&"; "&' + (PSET 't' '{Time_Unit}') + ')')) + @(QWork '=IF(AND({Calc_ID}="",{t_half_lower}="",{t}=""),1,0)' `
    ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Time_Unit}="",{t_half_lower}="",{t_half_upper}="",{t}=""),"E_MISSING_INPUT",IF(' + $UNITOK + ',"E_INVALID_UNIT",IF(OR(NOT(ISNUMBER({t_half_lower})),NOT(ISNUMBER({t_half_upper})),NOT(ISNUMBER({t}))),"E_MISSING_INPUT",IF(OR({t_half_lower}<=0,{t_half_upper}<=0),"E_NONPOSITIVE_PARAMETER",IF(NOT(' + $GATE + '),"E_APPLICABILITY_UNCONFIRMED",IF(OR({t_half_lower_Status}="",{t_half_upper_Status}="",{t_Status}=""),"E_MISSING_INPUT",IF({t_half_lower}>{t_half_upper},"E_REVERSED_BOUNDS",IF({t}<0,"E_OUT_OF_RANGE","")))))))))') `
    '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Time_Unit}="","a time unit",IF(NOT(ISNUMBER({t_half_lower})),"a numeric lower t_half",IF(NOT(ISNUMBER({t_half_upper})),"a numeric upper t_half",IF(NOT(ISNUMBER({t})),"a numeric t","the Parameter_Status of each input"))))))')

# Q9: amount infused before onset, from Administrations rate segments
$MINE = '((nADM_SegE_@_n+{O_@})-ABS(nADM_SegE_@_n-{O_@}))/2'
function Q9Sum([string]$b) {
    $mi = $MINE.Replace('@', $b)
    'SUMPRODUCT((nADM_SelKey={SelKey})*(nADM_Kind="INFUSION_SEGMENT")*nADM_Rate_n*(((' + $mi + ')-nADM_SegS_' + $b + '_n)+ABS((' + $mi + ')-nADM_SegS_' + $b + '_n))/2)/60'
}
$BADSEG = '(nADM_SelKey={SelKey})*(nADM_Kind="INFUSION_SEGMENT")*(LEFT(nADM_Seg_Check,2)="E_")'
$Q9Cols = @(QHead) + @(
    (CIn 'Exposure_Episode_ID' 'Exposure_Episode_ID (rate segments on Administrations)' 12 $null '@' 'INPUTS (no common gate; segment checks)'),
    (CIn 'Onset_Record_ID' 'Onset record ID (timestamp, e.g. EV-1/ONSET)' 14 $null '@')) + @(Par 'Planned_Dose_mg' 'Planned dose (mg), optional') + @(
    (CF 'Result_Display' 'Result (display: mg 1 decimal)' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown"&IF({Err}="E_INVALID_UNIT"," (infusion rates must be in mg/hour; mL/hour needs a concentration and is not supported in CALC-SPEC 1.0)",""),FIXED({Amount_mg},1,TRUE)&" mg infused before onset"&IF({Percent_of_Planned}<>"","; "&FIXED({Percent_of_Planned},1,TRUE)&"% of the planned "&{Planned_Dose_mg}&" mg","")&" ("&{Segments_Used}&" segments; "&{Time_Basis}&")"))' 40 $null 'RESULT'),
    (CF 'Amount_mg' 'Amount infused before onset, mg (unrounded)' '=IF({Ok}=1,IF({AllKnown}=1,{Sum_UTC},{Sum_Loc}),"")' 12),
    (CF 'Percent_of_Planned' 'Percent of planned dose (unrounded)' '=IF(AND({Ok}=1,ISNUMBER({Planned_Dose_mg})),100*{Amount_mg}/{Planned_Dose_mg},"")' 12),
    (CF 'Segments_Used' 'Segments used' '=IF({RowEmpty}=1,"",{N_Seg})' 8),
    (CF 'Time_Basis' 'Time basis' '=IF({Ok}=0,"",IF({AllKnown}=1,"UTC","local clock"))' 10)
) + @(QTail 'Amount infused, not systemic or effect-site exposure. Assumes the recorded times without interruption and ignores line dead volume.' 'Amount = sum over segments of rate (mg/hour) x hours of that segment before onset' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err},IF({AllKnown}=0,"W_TZ_UNKNOWN","")))' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": "&' + (MSGX '{Err}' '{Needs}') + ',IF({AllKnown}=0,"W_TZ_UNKNOWN: "&' + (MSGX '"W_TZ_UNKNOWN"' '""') + ',"")))') '=IF({RowEmpty}=1,"","Rate segments of exposure episode "&{Exposure_Episode_ID}&" (Administrations); onset "&{Onset_Record_ID}&IF({Planned_Dose_mg}<>"","; planned dose = "&{Planned_Dose_mg}&" mg ["&{Planned_Dose_mg_Status}&IF({Planned_Dose_mg_Source}<>"","; "&{Planned_Dose_mg_Source},"")&"]",""))' $false) + @(
    (CF 'RowEmpty' 'work: row empty' '=IF(AND({Calc_ID}="",{Exposure_Episode_ID}="",{Onset_Record_ID}=""),1,0)' 5 '0' 'WORKING'),
    (CF 'Key' 'work: key' '=IF({Calc_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Calc_ID})' 14),
    (CF 'SelKey' 'work: episode key' '={Case_ID}&"|"&{Case_Version}&"|"&{Exposure_Episode_ID}' 14),
    (CF 'O_Row' 'work: onset row' '=IF({Onset_Record_ID}="",0,IFERROR(MATCH({Case_ID}&"|"&{Case_Version}&"|"&{Onset_Record_ID},nTP_Key,0),0))' 6 '0'),
    (CF 'O_HT' 'work: onset has time' '=IF({O_Row}=0,0,INDEX(nTP_HT,{O_Row}))' 5 '0'),
    (CF 'O_OK' 'work: onset offset known' '=IF({O_Row}=0,0,INDEX(nTP_OK,{O_Row}))' 5 '0'),
    (CF 'O_Loc' 'work: onset local minutes' '=IF({O_Row}=0,0,INDEX(nTP_Loc_n,{O_Row}))' 10 '0'),
    (CF 'O_UTC' 'work: onset UTC minutes' '=IF({O_Row}=0,0,INDEX(nTP_UTC_n,{O_Row}))' 10 '0'),
    (CF 'N_Seg' 'work: segments in episode' '=COUNTIFS(nADM_SelKey,{SelKey},nADM_Kind,"INFUSION_SEGMENT")' 5 '0'),
    (CF 'AllKnown' 'work: all offsets known' '=IF(AND({O_OK}=1,COUNTIFS(nADM_SelKey,{SelKey},nADM_Kind,"INFUSION_SEGMENT",nADM_SegOK,0)=0),1,0)' 5 '0'),
    (CF 'BadSeg' 'work: first segment problem' ('=IF(SUMPRODUCT(' + $BADSEG + ')=0,"",INDEX(nADM_Seg_Check,MATCH(1,INDEX(' + $BADSEG + ',0),0)))') 14),
    (CF 'Err' 'work: blocking E-code' '=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Exposure_Episode_ID}="",{Onset_Record_ID}=""),"E_MISSING_INPUT",IF({O_Row}=0,"E_MISSING_INPUT",IF(INDEX(nTP_Err,{O_Row})<>"",INDEX(nTP_Err,{O_Row}),IF({O_HT}=0,"E_MISSING_INPUT",IF({N_Seg}=0,"E_MISSING_INPUT",IF({BadSeg}<>"",{BadSeg},IF(AND({Planned_Dose_mg}<>"",NOT(ISNUMBER({Planned_Dose_mg}))),"E_MISSING_INPUT",IF(AND(ISNUMBER({Planned_Dose_mg}),{Planned_Dose_mg}<=0),"E_NONPOSITIVE_PARAMETER",IF(AND(ISNUMBER({Planned_Dose_mg}),{Planned_Dose_mg_Status}=""),"E_MISSING_INPUT",""))))))))))' 14),
    (CF 'Needs' 'work: [input]' '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Exposure_Episode_ID}="","an Exposure_Episode_ID",IF({Onset_Record_ID}="","an onset record ID",IF({O_Row}=0,"the onset record "&{Onset_Record_ID},IF({O_HT}=0,"an onset time of day (timestamp)",IF({N_Seg}=0,"INFUSION_SEGMENT rows on Administrations for "&{Exposure_Episode_ID},IF({BadSeg}<>"","segment start and end timestamps and a rate in mg/hour","a numeric planned dose with its Parameter_Status"))))))))' 14),
    (CF 'Ok' 'work: result shown (1/0)' '=IF(AND({RowEmpty}=0,{Err}=""),1,0)' 5 '0'),
    (CF 'Sum_UTC' 'work: amount (UTC basis)' ('=IF({O_Row}=0,0,' + (Q9Sum 'UTC') + ')') 10),
    (CF 'Sum_Loc' 'work: amount (local basis)' ('=IF({O_Row}=0,0,' + (Q9Sum 'Loc') + ')') 10)
)

# Q10
# $Q10A: every administration of the episode (infusion rate segments are Q9 rows, not administrations here).
# $Q10M: those with a dose. CALC-SPEC G2: a missing dose, or no BSA at all, blocks with E_MISSING_INPUT (as the calculator).
$Q10A = '(nADM_SelKey={SelKey})*(nADM_Kind<>"INFUSION_SEGMENT")'
$Q10M = $Q10A + '*nADM_HasDose'
$Q10Cols = @(QHead) + @(
    (CIn 'Exposure_Episode_ID' 'Exposure_Episode_ID (doses and BSA on Administrations)' 12 $null '@' 'INPUTS (Method A is the default)'),
    (CIn 'Method_B' 'Method B on request: FIRST, LAST or STATED BSA (blank = not requested)' 12 'list:nLIST_MethodB'),
    (CIn 'Method_B_Stated_BSA' ('Method B stated BSA (m' + [char]0x00B2 + ')') 9 'num'),
    (CIn 'Method_B_BSA_Source' 'Method B BSA source' 18 $null '@'),
    (CF 'Result_Display' 'Result (display: mg/m2 1 decimal, method named)' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown","Method A (sum of dose / BSA at each administration): "&FIXED({Method_A_mg_m2},1,TRUE)&" mg/m'+[char]0x00B2+'"&IF({N_With_BSA}<{N_Doses}," (lower bound from "&{N_With_BSA}&" of "&{N_Doses}&" administrations)","")&IF({Method_B}="","",IF({MethodB_Err}<>"","; Method B: "&{MethodB_Err},"; Method B (total "&{Total_Dose_mg}&" mg / "&{Method_B_BSA_m2}&" m'+[char]0x00B2+', "&{Method_B}&" BSA): "&FIXED({Method_B_mg_m2},1,TRUE)&" mg/m'+[char]0x00B2+'"))))') 44 $null 'RESULT'),
    (CF 'Method_A_mg_m2' 'Method A, mg/m2 (unrounded; lower bound if W_PARTIAL_BSA)' ('=IF({Ok}=1,SUMPRODUCT(' + $Q10M + '*nADM_Dose_per_BSA),"")') 12),
    (CF 'Total_Dose_mg' 'Total dose, mg' ('=IF({Ok}=1,SUMPRODUCT(' + $Q10M + '*nADM_Dose_n),"")') 10),
    (CF 'N_Doses' 'Administrations with a dose' ('=IF({RowEmpty}=1,"",SUMPRODUCT(' + $Q10M + '))') 8),
    (CF 'N_With_BSA' 'Administrations with a BSA' ('=IF({RowEmpty}=1,"",SUMPRODUCT(' + $Q10M + '*nADM_HasBSA))') 8),
    (CF 'Method_B_BSA_m2' 'Method B BSA used (m2)' ('=IF(OR({Ok}=0,{Method_B}=""),"",IF({Method_B}="STATED",IF(ISNUMBER({Method_B_Stated_BSA}),{Method_B_Stated_BSA},""),IF({Method_B}="FIRST",IF(INDEX(nADM_HasBSA,{FirstRow})=1,INDEX(nADM_BSA_n,{FirstRow}),""),IF(LOOKUP(2,1/(' + $Q10M + '*(nADM_Sort_n={MaxSort})),nADM_HasBSA)=1,LOOKUP(2,1/(' + $Q10M + '*(nADM_Sort_n={MaxSort})),nADM_BSA_n),""))))') 10),
    (CF 'Method_B_mg_m2' 'Method B, mg/m2 (unrounded)' '=IF(OR({Ok}=0,{Method_B}="",{MethodB_Err}<>""),"",{Total_Dose_mg}/{Method_B_BSA_m2})' 12)
) + @(QTail 'Describes cumulative exposure. It is not a toxicity threshold.' 'Method A (default): sum of dose_i / BSA_i; Method B (on request): total dose / one stated BSA' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err},MID(IF({N_With_BSA}<{N_Doses},"; W_PARTIAL_BSA","")&IF({MethodB_Err}<>"","; "&{MethodB_Err}&" (Method B)",""),3,99)))' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": "&' + (MSGX '{Err}' '{Needs}') + ',MID(IF({N_With_BSA}<{N_Doses}," | W_PARTIAL_BSA: "&' + (MSGX '"W_PARTIAL_BSA"' '""') + ',"")&IF({MethodB_Err}<>""," | "&{MethodB_Err}&" (Method B): "&' + (MSGX '{MethodB_Err}' '"a BSA for Method B"') + ',""),4,999)))') '=IF({RowEmpty}=1,"","Doses (mg) and BSA (m2) recorded at each administration of "&{Exposure_Episode_ID}&" on Administrations"&IF({Method_B}<>"","; Method B BSA: "&{Method_B}&IF({Method_B_BSA_Source}<>""," ("&{Method_B_BSA_Source}&")",""),""))' $false) + @(
    (CF 'RowEmpty' 'work: row empty' '=IF(AND({Calc_ID}="",{Exposure_Episode_ID}=""),1,0)' 5 '0' 'WORKING'),
    (CF 'Key' 'work: key' '=IF({Calc_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Calc_ID})' 14),
    (CF 'SelKey' 'work: episode key' '={Case_ID}&"|"&{Case_Version}&"|"&{Exposure_Episode_ID}' 14),
    (CF 'Err' 'work: blocking E-code' ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Exposure_Episode_ID}=""),"E_MISSING_INPUT",IF(SUMPRODUCT(' + $Q10A + ')=0,"E_MISSING_INPUT",IF(SUMPRODUCT(' + $Q10A + '*(1-nADM_HasDose))>0,"E_MISSING_INPUT",IF(SUMPRODUCT(' + $Q10M + '*(nADM_Dose_Unit<>"mg"))>0,"E_INVALID_UNIT",IF(SUMPRODUCT(' + $Q10M + '*(nADM_Dose_n<=0))>0,"E_NONPOSITIVE_PARAMETER",IF(SUMPRODUCT(' + $Q10M + '*nADM_HasBSA*(nADM_BSA_n<=0))>0,"E_NONPOSITIVE_PARAMETER",IF(SUMPRODUCT(' + $Q10M + '*nADM_HasBSA)=0,"E_MISSING_INPUT",""))))))))') 14),
    (CF 'Needs' 'work: [input]' ('=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Exposure_Episode_ID}="","an Exposure_Episode_ID",IF(SUMPRODUCT(' + $Q10A + ')=0,"administrations with a dose for "&{Exposure_Episode_ID},IF(SUMPRODUCT(' + $Q10A + '*(1-nADM_HasDose))>0,"the dose (mg) of every administration of "&{Exposure_Episode_ID}&" ("&SUMPRODUCT(' + $Q10A + '*(1-nADM_HasDose))&" of "&SUMPRODUCT(' + $Q10A + ')&" have no dose)","the BSA recorded at one or more administrations of "&{Exposure_Episode_ID}&" (none is recorded)")))))') 14),
    (CF 'Ok' 'work: result shown (1/0)' '=IF(AND({RowEmpty}=0,{Err}=""),1,0)' 5 '0'),
    (CF 'MinSort' 'work: first administration key' ('=IF({Ok}=0,0,1E+12-SUMPRODUCT(MAX(' + $Q10M + '*(1E+12-nADM_Sort_n))))') 10 '0'),
    (CF 'MaxSort' 'work: last administration key' ('=IF({Ok}=0,0,SUMPRODUCT(MAX(' + $Q10M + '*nADM_Sort_n)))') 10 '0'),
    (CF 'FirstRow' 'work: first administration row' ('=IF({Ok}=0,0,MATCH(1,INDEX(' + $Q10M + '*(nADM_Sort_n={MinSort}),0),0))') 6 '0'),
    (CF 'MethodB_Err' 'work: Method B check' '=IF(OR({Ok}=0,{Method_B}=""),"",IF({Method_B_BSA_m2}="","E_MISSING_INPUT",IF({Method_B_BSA_m2}<=0,"E_NONPOSITIVE_PARAMETER","")))' 12)
)

# Q11
$SCRU = ([string][char]0x00B5 + 'mol/L')
$Q11Cols = @(QHead) + @(
    (CIn 'Stability_Confirmed' 'Gate: serum creatinine stable at this time? Yes/No' 10 'list:nLIST_YesNo' $null 'GATE (Q11)'),
    (CIn 'Stability_Basis' 'Basis for stability (values and dates)' 26 $null '@'),
    (CIn 'Age_years' 'Age (years, 18-120)' 7 'num' $null 'INPUTS'),
    (CIn 'Age_Status' 'Age status' 13 'list:nLIST_Status'),
    (CIn 'Weight_kg' 'Weight (kg)' 7 'num'),
    (CIn 'Weight_Type' 'Weight type' 10 'list:nLIST_WeightType'),
    (CIn 'Weight_Status' 'Weight status' 13 'list:nLIST_Status'),
    (CIn 'SCr' 'Serum creatinine' 8 'num'),
    (CIn 'SCr_Unit' ('Creatinine unit (mg/dL or ' + $SCRU + ')') 9 'list:nLIST_SCrUnit'),
    (CIn 'SCr_Status' 'Creatinine status' 13 'list:nLIST_Status'),
    (CIn 'SCr_Source' 'Creatinine source (record, date)' 18 $null '@'),
    (CIn 'Sex' 'Sex as recorded' 10 'list:nLIST_Sex'),
    (CF 'Result_Display' 'Result (display: whole mL/minute)' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown",IF({Sex}="not recorded","Female "&FIXED({CrCl_Female_If_Sex_Not_Recorded},0,TRUE)&" mL/minute; male "&FIXED({CrCl_Male_If_Sex_Not_Recorded},0,TRUE)&" mL/minute (sex not recorded; both shown)",FIXED({CrCl_mL_min},0,TRUE)&" mL/minute estimated creatinine clearance ("&{Sex}&")")))' 40 $null 'RESULT'),
    (CF 'SCr_mg_dL' 'Creatinine, mg/dL (unrounded)' ('=IF({Ok}=1,IF({SCr_Unit}="' + $SCRU + '",{SCr}/88.4,{SCr}),"")') 10),
    (CF 'CrCl_mL_min' 'CrCl for recorded sex, mL/minute (unrounded)' '=IF(AND({Ok}=1,{Sex}<>"not recorded"),(140-{Age_years})*{Weight_kg}/(72*{SCr_mg_dL})*IF({Sex}="female",0.85,1),"")' 12),
    (CF 'CrCl_Female_If_Sex_Not_Recorded' 'If sex not recorded: female estimate (unrounded)' '=IF(AND({Ok}=1,{Sex}="not recorded"),0.85*(140-{Age_years})*{Weight_kg}/(72*{SCr_mg_dL}),"")' 12),
    (CF 'CrCl_Male_If_Sex_Not_Recorded' 'If sex not recorded: male estimate (unrounded)' '=IF(AND({Ok}=1,{Sex}="not recorded"),(140-{Age_years})*{Weight_kg}/(72*{SCr_mg_dL}),"")' 12)
) + @(QTail 'Assumes stable creatinine. It is not drug clearance and not a dosing recommendation.' 'Cockcroft-Gault: CrCl = (140 - age) x weight / (72 x SCr), x 0.85 when female sex is recorded; umol/L / 88.4 = mg/dL' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err},IF({Sex}="not recorded","W_SEX_NOT_RECORDED","")))' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": "&' + (MSGX '{Err}' '{Needs}') + ',IF({Sex}="not recorded","W_SEX_NOT_RECORDED: "&' + (MSGX '"W_SEX_NOT_RECORDED"' '""') + ',"")))') '=IF({RowEmpty}=1,"","age = "&{Age_years}&" years ["&{Age_Status}&"]; weight = "&{Weight_kg}&" kg, "&{Weight_Type}&" ["&{Weight_Status}&"]; serum creatinine = "&{SCr}&" "&{SCr_Unit}&" ["&{SCr_Status}&IF({SCr_Source}<>"","; "&{SCr_Source},"")&"]; sex: "&{Sex}&"; stability: "&IF({Stability_Confirmed}="Yes","confirmed ("&{Stability_Basis}&")","not confirmed"))' $false) + @(
    (CF 'RowEmpty' 'work: row empty' '=IF(AND({Calc_ID}="",{Age_years}="",{SCr}=""),1,0)' 5 '0' 'WORKING'),
    (CF 'Key' 'work: key' '=IF({Calc_ID}="","",{Case_ID}&"|"&{Case_Version}&"|"&{Calc_ID})' 14),
    (CF 'Err' 'work: blocking E-code' '=IF({RowEmpty}=1,"",IF({Calc_ID}="","E_MISSING_INPUT",IF({Stability_Confirmed}<>"Yes","E_CREATININE_STABILITY_UNCONFIRMED",IF(OR({Age_years}="",{Weight_kg}="",{Weight_Type}="",{SCr}="",{SCr_Unit}="",{Sex}=""),"E_MISSING_INPUT",IF(OR(NOT(ISNUMBER({Age_years})),NOT(ISNUMBER({Weight_kg})),NOT(ISNUMBER({SCr}))),"E_MISSING_INPUT",IF(OR({Weight_kg}<=0,{SCr}<=0),"E_NONPOSITIVE_PARAMETER",IF(OR({Age_years}<18,{Age_years}>120),"E_OUT_OF_RANGE",IF(ISNA(MATCH({SCr_Unit},nLIST_SCrUnit,0)),"E_INVALID_UNIT",IF(OR({Age_Status}="",{Weight_Status}="",{SCr_Status}=""),"E_MISSING_INPUT","")))))))))' 14),
    (CF 'Needs' 'work: [input]' '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF(NOT(ISNUMBER({Age_years})),"a numeric age",IF(NOT(ISNUMBER({Weight_kg})),"a numeric weight in kg",IF({Weight_Type}="","the weight type",IF(NOT(ISNUMBER({SCr})),"a numeric serum creatinine",IF({SCr_Unit}="","a creatinine unit",IF({Sex}="","sex as recorded (or not recorded)","the Parameter_Status of each input"))))))))' 14),
    (CF 'Ok' 'work: result shown (1/0)' '=IF(AND({RowEmpty}=0,{Err}=""),1,0)' 5 '0')
)

# Q6 (PD_Profile, teaching demonstration)
$Q6Cols = @(QHead) + @(
    (CIn 'Teaching_Switch' 'Teaching-demonstration switch on? Yes/No' 10 'list:nLIST_YesNo' $null 'GATE (Q6)'),
    (CIn 'Time_Unit' 'Time unit (minutes, hours, days)' 9 'list:nLIST_TimeUnit' $null 'INPUTS')) + @(Par 'Replacement_Half_Time' 'Replacement half-time (in Time_Unit)') + @(Par 't' 't (in Time_Unit)') + @(
    (CF 'Result_Display' 'Result (teaching demonstration; percent to 3 significant figures)' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": no result is shown"&IF({Err}="E_APPLICABILITY_UNCONFIRMED"," (Q6 is shown only with the teaching-demonstration switch on)",""),' + (P3 '{Percent_Replaced}') + '&"% replaced after "&{t}&" "&{Time_Unit}&" (teaching demonstration with hypothetical values)"))') 40 $null 'RESULT'),
    (CF 'Fraction_Replaced' 'Fraction replaced (unrounded)' '=IF({Ok}=1,1-2^(-{t}/{Replacement_Half_Time}),"")' 12),
    (CF 'Percent_Replaced' 'Percent replaced (unrounded)' '=IF({Ok}=1,100*{Fraction_Replaced},"")' 12)
) + @(QTail 'Hypothetical turnover illustration. It is not a measured recovery curve for platelets or any other clinical target.' 'fraction replaced = 1 - 2^(-t / replacement half-time)' '=IF({RowEmpty}=1,"",IF({Err}<>"",{Err},"I_TEACHING_DEMONSTRATION"))' ('=IF({RowEmpty}=1,"",IF({Err}<>"",{Err}&": "&' + (MSGX '{Err}' '{Needs}') + ',"I_TEACHING_DEMONSTRATION: "&' + (MSGX '"I_TEACHING_DEMONSTRATION"' '""') + '))') ('=IF({RowEmpty}=1,"",' + (PSET 'Replacement_Half_Time' '{Time_Unit}') + '&"; "&' + (PSET 't' '{Time_Unit}') + ')') $false) + @(QWork '=IF(AND({Calc_ID}="",{Replacement_Half_Time}="",{t}=""),1,0)' `
    ('=IF({RowEmpty}=1,"",IF(OR({Calc_ID}="",{Time_Unit}="",{Replacement_Half_Time}="",{t}=""),"E_MISSING_INPUT",IF(' + $UNITOK + ',"E_INVALID_UNIT",IF(OR(NOT(ISNUMBER({Replacement_Half_Time})),NOT(ISNUMBER({t}))),"E_MISSING_INPUT",IF({Replacement_Half_Time}<=0,"E_NONPOSITIVE_PARAMETER",IF({Teaching_Switch}<>"Yes","E_APPLICABILITY_UNCONFIRMED",IF(OR({Replacement_Half_Time_Status}="",{t_Status}=""),"E_MISSING_INPUT",IF({t}<0,"E_OUT_OF_RANGE",""))))))))') `
    '=IF({Err}<>"E_MISSING_INPUT","",IF({Calc_ID}="","a Calc_ID",IF({Time_Unit}="","a time unit",IF(NOT(ISNUMBER({Replacement_Half_Time})),"a numeric replacement half-time",IF(NOT(ISNUMBER({t})),"a numeric t","the Parameter_Status of each input")))))')

# ---------------- reserved rows for technical tables
$SV = 'hypothetical example'
function QR($tid, $case, $ver, $q, $extra) { MergeH @(@{ Calc_ID = $tid; Case_ID = $case; Case_Version = $ver; Question = $q }, $extra) }
$GV = GY 'Hypothetical teaching value from the vector file' 'Reference amount at time zero'
$Q1Vec = @(
    (QR 'Q1-V16' 'VEC-V16' 1 'Vector V16' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half' 24 $SV "$V V16"), (PV 't' 72 $SV "$V V16")))),
    (QR 'Q1-V24a' 'VEC-V24' 1 'Vector V24, case 1: half-life 0' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half' 0 $SV "$V V24"), (PV 't' 72 $SV "$V V24")))),
    (QR 'Q1-V24b' 'VEC-V24' 1 'Vector V24, case 2: unit fortnights' (MergeH @($GV, @{ Time_Unit = 'fortnights' }, (PV 't_half' 24 $SV "$V V24"), (PV 't' 72 $SV "$V V24")))),
    (QR 'Q1-V24c' 'VEC-V24' 1 'Vector V24, case 3: gate unconfirmed' (MergeH @($GV, @{ G_Model = 'No'; Time_Unit = 'hours' }, (PV 't_half' 24 $SV "$V V24"), (PV 't' 72 $SV "$V V24"))))
)
$Q2Vec = @((QR 'Q2-V17' 'VEC-V17' 1 'Vector V17' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half' 24 $SV "$V V17"), (PV 'f_target' 0.9 $SV "$V V17")))))
$Q3Vec = @((QR 'Q3-V18' 'VEC-V18' 1 'Vector V18' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half' 24 $SV "$V V18"), (PV 'tau' 12 $SV "$V V18")))))
$Q4Vec = @((QR 'Q4-V19' 'VEC-V19' 1 'Vector V19' (MergeH @((GY 'Hypothetical clearances from the vector file' 'Eventual average steady state'), @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 5 $SV "$V V19"), (PV 'CL_after' 2.5 $SV "$V V19")))))
$C04Q4G = GY 'Relative total clearance (baseline = 1, entered on the L/hour scale; only the ratio is used); fe 0.5-0.7 from healthy volunteers given IV digoxin [LANOXIN US 12.3]' 'Eventual average steady state after the renal change'
$Q4Ex = @(
    (QR 'Q4-C04-half-fe0.5' 'EX-C04' 2 'Renal elimination halved, fe 0.5: CL_after/CL_before = 0.5 + 0.5 x 0.5 = 0.75 (Chapter 11, Worked exercise 11.2)' (MergeH @($C04Q4G, @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 1 'model assumption' 'Chapter 11, Worked exercise 11.2'), (PV 'CL_after' 0.75 'model assumption' 'Chapter 11, Worked exercise 11.2')))),
    (QR 'Q4-C04-half-fe0.7' 'EX-C04' 2 'Renal elimination halved, fe 0.7: 0.3 + 0.7 x 0.5 = 0.65' (MergeH @($C04Q4G, @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 1 'model assumption' 'Chapter 11, Worked exercise 11.2'), (PV 'CL_after' 0.65 'model assumption' 'Chapter 11, Worked exercise 11.2')))),
    (QR 'Q4-C04-quarter-fe0.5' 'EX-C04' 2 'Renal elimination at a quarter, fe 0.5: 0.625' (MergeH @($C04Q4G, @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 1 'model assumption' 'Chapter 11, Worked exercise 11.2'), (PV 'CL_after' 0.625 'model assumption' 'Chapter 11, Worked exercise 11.2')))),
    (QR 'Q4-C04-quarter-fe0.7' 'EX-C04' 2 'Renal elimination at a quarter, fe 0.7: 0.475' (MergeH @($C04Q4G, @{ CL_Unit = 'L/hour' }, (PV 'CL_before' 1 'model assumption' 'Chapter 11, Worked exercise 11.2'), (PV 'CL_after' 0.475 'model assumption' 'Chapter 11, Worked exercise 11.2'))))
)
$Q5Vec = @((QR 'Q5-V20' 'VEC-V20' 1 'Vector V20' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 'Parent_t_half' 6 $SV "$V V20"), (PV 'Metabolite_t_half' 72 $SV "$V V20"), (PV 't' 48 $SV "$V V20")))))
$Q7Vec = @((QR 'Q7-V22' 'VEC-V22' 1 'Vector V22' (MergeH @($GV, @{ Time_Unit = 'days' }, (PV 't_half' 21 $SV "$V V22"), (PV 't' 42 $SV "$V V22")))))
$C28G = GY 'Rituximab terminal half-life, NHL patients, median 22 days (range 6.1-52) [RITUXAN US 12.3]; label fact for a population, not this patient' 'Post-infusion concentration after the last infusion (3 March 2025)'
$Q7Ex = @(
    (QR 'Q7-C28-435' 'EX-C28' 2 'Exposure persistence at the first positive test, 435 days after the last infusion (median half-life)' (MergeH @($C28G, @{ Time_Unit = 'days' }, (PV 't_half' 22 'label fact' 'RITUXAN US 12.3 (NHL, median)'), (PV 't' 435 'case measurement' 'Intervals I01')))),
    (QR 'Q7-C28-343' 'EX-C28' 2 'At the earliest possible onset, 343 days (median half-life)' (MergeH @($C28G, @{ Time_Unit = 'days' }, (PV 't_half' 22 'label fact' 'RITUXAN US 12.3 (NHL, median)'), (PV 't' 343 'case measurement' 'Intervals I02 (minimum)'))))
)
$Q8Vec = @((QR 'Q8-V23' 'VEC-V23' 1 'Vector V23' (MergeH @($GV, @{ Time_Unit = 'hours' }, (PV 't_half_lower' 18 $SV "$V V23"), (PV 't_half_upper' 36 $SV "$V V23"), (PV 't' 72 $SV "$V V23")))))
$C04G1 = GY 'Terminal elimination half-life 1.5-2 days, healthy volunteers with normal renal function [LANOXIN US 12.3]; not the patient''s population' 'Concentration when digoxin was withheld on 5 July 2026; no further doses; t = 3 days to the improvement on 8 July'
$C04G2 = GY 'Terminal elimination half-life 3.5-5 days, anuric patients [LANOXIN US 12.3]; not the patient''s population' 'Concentration when digoxin was withheld on 5 July 2026; no further doses; t = 3 days to the improvement on 8 July'
$Q8Ex = @(
    (QR 'Q8-C04-normal' 'EX-C04' 2 'Fraction remaining at the documented improvement, normal renal function half-lives (Chapter 23, Section 23.11)' (MergeH @($C04G1, @{ Time_Unit = 'days' }, (PV 't_half_lower' 1.5 'label fact' 'LANOXIN US 12.3'), (PV 't_half_upper' 2 'label fact' 'LANOXIN US 12.3'), (PV 't' 3 'case measurement' 'Intervals I06 (5 to 8 July)')))),
    (QR 'Q8-C04-anuric' 'EX-C04' 2 'Same, anuric-patient half-lives' (MergeH @($C04G2, @{ Time_Unit = 'days' }, (PV 't_half_lower' 3.5 'label fact' 'LANOXIN US 12.3'), (PV 't_half_upper' 5 'label fact' 'LANOXIN US 12.3'), (PV 't' 3 'case measurement' 'Intervals I06 (5 to 8 July)')))),
    (QR 'Q8-C28-435' 'EX-C28' 2 'Labeled NHL half-life range at the first positive test (435 days)' (MergeH @($C28G, @{ Time_Unit = 'days' }, (PV 't_half_lower' 6.1 'label fact' 'RITUXAN US 12.3 (NHL range)'), (PV 't_half_upper' 52 'label fact' 'RITUXAN US 12.3 (NHL range)'), (PV 't' 435 'case measurement' 'Intervals I01')))),
    (QR 'Q8-C28-343' 'EX-C28' 2 'Labeled NHL half-life range at the earliest possible onset (343 days)' (MergeH @($C28G, @{ Time_Unit = 'days' }, (PV 't_half_lower' 6.1 'label fact' 'RITUXAN US 12.3 (NHL range)'), (PV 't_half_upper' 52 'label fact' 'RITUXAN US 12.3 (NHL range)'), (PV 't' 343 'case measurement' 'Intervals I02 (minimum)'))))
)
$Q9Ex = @((QR 'Q9-C01' 'EX-C01' 2 'Amount of Drug A infused before onset (pump rate recorded in mL/hour only)' @{ Exposure_Episode_ID = 'EX-A'; Onset_Record_ID = 'EV-1/ONSET' }))
$Q9Vec = @((QR 'Q9-V25' 'VEC-V25' 1 'Vector V25' (MergeH @(@{ Exposure_Episode_ID = 'EX1'; Onset_Record_ID = 'EV1/ONSET' }, (PV 'Planned_Dose_mg' 500 $SV "$V V25")))))
$Q10Vec = @(
    (QR 'Q10-V26-first' 'VEC-V26' 1 'Vector V26, Method B with the first BSA' @{ Exposure_Episode_ID = 'EX1'; Method_B = 'FIRST'; Method_B_BSA_Source = 'BSA at the first administration' }),
    (QR 'Q10-V26-last' 'VEC-V26' 1 'Vector V26, Method B with the last BSA' @{ Exposure_Episode_ID = 'EX1'; Method_B = 'LAST'; Method_B_BSA_Source = 'BSA at the last administration' })
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
$Q6Vec = @((QR 'Q6-V21' 'VEC-V21' 1 'Vector V21' (MergeH @(@{ Teaching_Switch = 'Yes'; Time_Unit = 'days' }, (PV 'Replacement_Half_Time' 3 $SV "$V V21"), (PV 't' 6 $SV "$V V21")))))

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
    (CF 'Export_Safe_Description' 'Export-safe description (CSV formula-injection escape)' '=IF({Description}="","",IF(OR(LEFT({Description},1)="=",LEFT({Description},1)="+",LEFT({Description},1)="-",LEFT({Description},1)="@"),"''"&{Description},{Description}))' 40 $null 'EXPORT'))
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
    (CF 'Export_Safe_Narrative' 'Export-safe narrative (CSV escape)' '=IF({Narrative}="","",IF(OR(LEFT({Narrative},1)="=",LEFT({Narrative},1)="+",LEFT({Narrative},1)="-",LEFT({Narrative},1)="@"),"''"&{Narrative},{Narrative}))' 40))
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
    @('EX-C01', 'Ch 3, 3.7 (Question 3)', 'LAST_DOSE_BEFORE_ONSET selects the 09:00 infusion', 'INT', 'Anchor_Record_ID_Used', 'EX-C01|2|I03', 'text', 'ADM-A1', $null, 'infusion started 09:00'),
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
    @('EX-C28', 'Ch 20, 20.4 (Question 5)', 'Longest labeled t_half 52 days at 343 days', 'Q8', 'Percent_at_Upper', 'EX-C28|2|Q8-C28-343', 'num', 1.03, 2, '1.03%')
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
    @('V27', 'Q11', 'numeric', 'PK_Exercises (Q11)', '68.055556, 34.027778, female 57.847222; sex missing shows both; unconfirmed stability blocks')
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
    @('V21', 'Fraction replaced', 'Q6', 'Fraction_Replaced', 'VEC-V21|1|Q6-V21', 'num', 0.75), @('V21', 'Flags', 'Q6', 'QC_Flags', 'VEC-V21|1|Q6-V21', 'flags', 'I_TEACHING_DEMONSTRATION'),
    @('V22', 'Fraction remaining', 'Q7', 'Fraction_Remaining', 'VEC-V22|1|Q7-V22', 'num', 0.25), @('V22', 'Flags', 'Q7', 'QC_Flags', 'VEC-V22|1|Q7-V22', 'flags', 'W_BIOLOGIC_MODEL_CAUTION'),
    @('V23', 'Fraction at lower t_half', 'Q8', 'Fraction_at_Lower', 'VEC-V23|1|Q8-V23', 'num', 0.0625), @('V23', 'Fraction at upper t_half', 'Q8', 'Fraction_at_Upper', 'VEC-V23|1|Q8-V23', 'num', 0.25), @('V23', 'Flags', 'Q8', 'QC_Flags', 'VEC-V23|1|Q8-V23', 'flags', 'I_SCENARIO_NOT_CI'),
    @('V24', 'Half-life 0', 'Q1', 'QC_Flags', 'VEC-V24|1|Q1-V24a', 'flags', 'E_NONPOSITIVE_PARAMETER'), @('V24', 'Unit fortnights', 'Q1', 'QC_Flags', 'VEC-V24|1|Q1-V24b', 'flags', 'E_INVALID_UNIT'), @('V24', 'Gate unconfirmed', 'Q1', 'QC_Flags', 'VEC-V24|1|Q1-V24c', 'flags', 'E_APPLICABILITY_UNCONFIRMED'), @('V24', 'Parameter status travels into Parameter_Set (V16 row)', 'Q1', 'Parameter_Set', 'VEC-V16|1|Q1-V16', 'contains', 'hypothetical example'),
    @('V25', 'Amount infused (mg)', 'Q9', 'Amount_mg', 'VEC-V25|1|Q9-V25', 'num', 58.333333), @('V25', 'Percent of planned dose', 'Q9', 'Percent_of_Planned', 'VEC-V25|1|Q9-V25', 'num', 11.666667), @('V25', 'Flags', 'Q9', 'QC_Flags', 'VEC-V25|1|Q9-V25', 'flags', ''),
    @('V26', 'Method A (mg/m2)', 'Q10', 'Method_A_mg_m2', 'VEC-V26|1|Q10-V26-first', 'num', 164.052288), @('V26', 'Method B, first BSA', 'Q10', 'Method_B_mg_m2', 'VEC-V26|1|Q10-V26-first', 'num', 161.111111), @('V26', 'Method B, last BSA', 'Q10', 'Method_B_mg_m2', 'VEC-V26|1|Q10-V26-last', 'num', 170.588235), @('V26', 'Flags', 'Q10', 'QC_Flags', 'VEC-V26|1|Q10-V26-first', 'flags', ''),
    @('V27', 'Case 1: male, 1.0 mg/dL', 'Q11', 'CrCl_mL_min', 'VEC-V27|1|Q11-V27a', 'num', 68.055556), @('V27', 'Case 2: male, 2.0 mg/dL', 'Q11', 'CrCl_mL_min', 'VEC-V27|1|Q11-V27b', 'num', 34.027778), @('V27', 'Case 3: female, 1.0 mg/dL', 'Q11', 'CrCl_mL_min', 'VEC-V27|1|Q11-V27c', 'num', 57.847222),
    @('V27', 'Case 4: sex not recorded, female estimate', 'Q11', 'CrCl_Female_If_Sex_Not_Recorded', 'VEC-V27|1|Q11-V27d', 'num', 57.847222), @('V27', 'Case 4: sex not recorded, male estimate', 'Q11', 'CrCl_Male_If_Sex_Not_Recorded', 'VEC-V27|1|Q11-V27d', 'num', 68.055556),
    @('V27', 'Case 5: no estimate', 'Q11', 'CrCl_mL_min', 'VEC-V27|1|Q11-V27e', 'blank', ''),
    @('V27', 'Flags case 1', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27a', 'flags', ''), @('V27', 'Flags case 2', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27b', 'flags', ''), @('V27', 'Flags case 3', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27c', 'flags', ''), @('V27', 'Flags case 4', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27d', 'flags', 'W_SEX_NOT_RECORDED'), @('V27', 'Flags case 5', 'Q11', 'QC_Flags', 'VEC-V27|1|Q11-V27e', 'flags', 'E_CREATININE_STABILITY_UNCONFIRMED')
)

function LookupF([string]$t, [string]$col, [string]$key) { '=IFERROR(INDEX(n' + $t + '_' + $col + ',MATCH("' + $key + '",n' + $t + '_Key,0)),"(row not found)")' }

# ================================================================== BUILD
$SheetOrder = @('Read_Me', 'Cases', 'Exposures', 'Administrations', 'Events', 'Milestones', 'Drug_Event_Links', 'Intervals', 'Product_Dossier', 'Parameter_Register', 'PK_Exercises', 'PD_Profile', 'Temporal_Profile', 'Biology_Review', 'Dechallenge_Rechallenge', 'Alternatives_Followup', 'Assessment_QC', 'Examples', 'Verification', 'Timepoints', 'Flags_Messages', 'Lists')
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
    $ws.Range('A2').Value2 = 'Missing-data reason codes are the example codes proposed in Chapter 23 (not part of CALC-SPEC 1.0). Timing-statement kinds follow change register B14-13.'
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
    $ws.Range('A1').Value2 = 'Flags and user messages (CALC-SPEC 1.0, Section 5; shown verbatim, [input] filled in)'; $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Size = 13
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
    Def-Names $iCase 'nCASE_' @('Key')
    $tpHelp = @{ S_Prec = 'TIMESTAMP: Date + Time + UTC offset (choose unknown if not recorded). DATE: Date only. MONTH: Year + Month. YEAR: Year. RANGE: earliest and latest dates as stated, with the source.' }
    $iExp = Layout-Table $S['Exposures'] 1 'Exposures: one row per exposure episode (Start, Stop and Restart are timepoints)' $note $ExpCols 30 $ExpEx $ExpVec 'EXP' $EX_LABEL $VEC_LABEL $tpHelp
    Def-Names $iExp 'nEXP_' @('Key')
    $iAdm = Layout-Table $S['Administrations'] 1 'Administrations: one row per dose, administration or infusion rate segment' $note $AdmCols 60 $AdmEx $AdmVec 'ADM' $EX_LABEL $VEC_LABEL @{ A_Prec = $tpHelp.S_Prec; Kind = 'INFUSION_SEGMENT rows need a start timestamp, an end timestamp and a rate in mg/hour for Q9.' }
    Def-Names $iAdm 'nADM_'
    $iEvt = Layout-Table $S['Events'] 1 'Events: one row per event episode (Onset and First detection are timepoints)' $note $EvtCols 15 $EvtEx $EvtVec 'EVT' $EX_LABEL $VEC_LABEL @{ O_Prec = $tpHelp.S_Prec }
    Def-Names $iEvt 'nEVT_' @('Key')
    $iMs = Layout-Table $S['Milestones'] 1 'Milestones: tests, organ function, interventions, recovery and other dated records' $note $MsCols 40 $MsEx @() 'MS' $EX_LABEL $VEC_LABEL @{ M_Prec = $tpHelp.S_Prec }
    Def-Names $iMs 'nMS_' @('Key')
    $iLnk = Layout-Table $S['Drug_Event_Links'] 1 'Drug_Event_Links: pairs selected by the reviewer (an interval needs a link: E_MISSING_LINK)' $note $LnkCols 20 $LnkEx $LnkVec 'LNK' $EX_LABEL $VEC_LABEL
    Def-Names $iLnk 'nLNK_'
    Write-Host "Record sheets laid out"

    # ---------- Timepoints (derived index of every dated record)
    $wsT = $S['Timepoints']
    $tpCols = @('Key', 'Record_ID', 'Record_Type', 'Case_ID', 'Case_Version', 'Case_Key', 'Episode_ID', 'Precision', 'Offset', 'Err', 'HT', 'OK', 'Loc', 'UTC', 'Lo', 'Hi', 'ISO', 'Display', 'Original_Text', 'Source', 'Source_Location', 'Dup', 'Loc_n', 'UTC_n', 'Lo_n', 'Hi_n')
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
        foreach ($k in $fx.Keys) {
            $f = Expand $fx[$k] $tpMap $sg.s
            try { $wsT.Range($tpMap[$k] + $sg.s + ':' + $tpMap[$k] + $sg.e).Formula = $f } catch { throw "Timepoints $k rejected: $f :: $($_.Exception.Message)" }
        }
    }
    $uni = [ordered]@{
        Key      = '=IF({Record_ID}="","",{Case_Key}&"|"&{Record_ID})'
        Case_Key = '=IF({Record_ID}="","",{Case_ID}&"|"&{Case_Version})'
        ISO      = ('=IF(OR({Key}="",{Err}<>""),"",IF({Precision}="TIMESTAMP",' + (ISOD 'INT({Loc}/1440)') + '&"T"&' + (HHMM '{Loc}') + '&IF({OK}=1,{Offset},""),IF({Precision}="DATE",' + (ISOD '{Lo}') + ',IF({Precision}="MONTH",YEAR({Lo})&"-"&RIGHT("0"&MONTH({Lo}),2),IF({Precision}="YEAR",YEAR({Lo})&"",' + (ISOD '{Lo}') + '&"/"&' + (ISOD '{Hi}') + ')))))')
        Display  = ('=IF({Key}="","",IF({Err}<>"",{Err},IF({Precision}="TIMESTAMP",' + (DTXT 'INT({Loc}/1440)') + '&" "&' + (HHMM '{Loc}') + '&IF({OK}=1," (UTC"&{Offset}&")"," (offset unknown)"),IF({Precision}="DATE",' + (DTXT '{Lo}') + ',IF({Precision}="MONTH",INDEX(nLIST_Months,MONTH({Lo}))&" "&YEAR({Lo})&" ("&' + (DTXT '{Lo}') + '&" to "&' + (DTXT '{Hi}') + '&")",IF({Precision}="YEAR",YEAR({Lo})&" (1 January to 31 December)",' + (DTXT '{Lo}') + '&" to "&' + (DTXT '{Hi}') + '))))))')
        Dup      = '=IF({Key}="",0,COUNTIF(nTP_Key,{Key}))'
        Loc_n    = '=IF({Loc}="",0,{Loc})'
        UTC_n    = '=IF({UTC}="",0,{UTC})'
        Lo_n     = '=IF({Lo}="",0,{Lo})'
        Hi_n     = '=IF({Hi}="",0,{Hi})'
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
    $wsE.Range('A2').Value2 = 'Cases EX-C01 (Chapter 3, Section 3.7, version 2), EX-C04 (Chapter 11, Section 11.8, version 2) and EX-C28 (Chapter 20, Section 20.4, version 2) are entered in the reserved blue rows of the input sheets with the home chapters'' dates, times and values. Each check below reads the workbook''s own result and compares it, rounded as the chapter states it, with the chapter''s number. All cases and dates are fictional.'
    $wsE.Range('A2').WrapText = $true; $wsE.Range('A2:L2').Merge(); $wsE.Rows(2).RowHeight = [double]42
    $eh = @('#', 'Case', 'Home chapter', 'Item', 'Workbook row (sheet: key)', 'Chapter states', 'Check', 'Expected', 'Digits', 'Workbook value (unrounded)', 'Workbook display', 'Match')
    $ER0 = 6
    for ($j = 0; $j -lt $eh.Count; $j++) { $c = $wsE.Cells.Item($ER0, $j + 1); $c.Value2 = $eh[$j]; $c.Font.Bold = $true; $c.Interior.Color = $C_FHdr; $c.Font.Color = RGBc 255 255 255; $c.WrapText = $true }
    $sheetOf = @{ INT = 'Intervals'; Q4 = 'PK_Exercises (Q4)'; Q7 = 'PK_Exercises (Q7)'; Q8 = 'PK_Exercises (Q8)'; Q9 = 'PK_Exercises (Q9)'; Q11 = 'PK_Exercises (Q11)' }
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
    $wsE.Cells.Item($eRow, 1).Value2 = 'Outside CALC-SPEC 1.0 (shown for Chapter 20 and 23 only): 24 calendar months after the last infusion, compared by calendar date (author''s convention; no CALC-SPEC rule adds calendar months).'
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
    $wsV.Range('A1').Value2 = 'Verification: CALC-SPEC 1.0 vectors V01-V27 recomputed by this workbook'; $wsV.Range('A1').Font.Bold = $true; $wsV.Range('A1').Font.Size = 14; $wsV.Range('A1').Font.Color = $C_Title
    $wsV.Range('A2').Value2 = 'Each vector is entered as a reserved, locked case (VEC-Vnn) on the production sheets and computed by the same column formulas a user''s case uses. Expected values are the unrounded values of tto_calc_test_vectors_v1.0.json. Numbers are compared unrounded: relative tolerance 1E-6, or absolute 1E-9 when the expected value is zero. V12 and V14 describe behavior and are checked by inspection.'
    $wsV.Range('A2:H2').Merge(); $wsV.Range('A2').WrapText = $true; $wsV.Rows(2).RowHeight = [double]54
    $wsV.Range('A4').Value2 = 'Overall status:'; $wsV.Range('A4').Font.Bold = $true; $wsV.Range('A4').Font.Size = 12
    $wsV.Range('A5').Value2 = 'Relative tolerance'; SetV $wsV.Range('B5') 1E-6; $wsV.Range('C5').Value2 = 'Absolute tolerance (expected = 0)'; SetV $wsV.Range('D5') 1E-9
    $wsV.Range('E5').Value2 = 'Source: tto_calc_test_vectors_v1.0.json (spec CALC-SPEC 1.0, 2026-10-19)'
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
        'V14 Export and re-import: the workbook has no macros, so export is done by the user (File > Save As > CSV UTF-8 for a sheet, or copy the table). Inspect the V06 rows (Exposures VEC-V06 Start RANGE 1-10 January 2026 with original text; Events onset; Drug_Event_Links L1; Intervals VEC-V06|1|I1 with -5 to 4 calendar days, SEQUENCE_UNCERTAIN, W_SEQUENCE_UNCERTAIN) before export and after re-import: bounds, unit, original text and precision, link, sequence status and flags must be unchanged. Use the Export_Safe columns for free text that starts with =, +, - or @.',
        'Note on V15: the vector has no anchor or onset, so the interval row shows E_MISSING_INPUT and no value; the C6 check (C6_Code = I_APPROXIMATE_NOT_COMPARED) is the vector''s expected flag.'
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
    AddRM 'Version' ($VersionLabel + '. Calculation rules: CALC-SPEC 1.0 (Stage 1 document C). Verification vectors: tto_calc_test_vectors_v1.0.json (V01-V27).') 'kv'
    AddRM 'License' 'Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0), https://creativecommons.org/licenses/by-nc/4.0/' 'kv'
    AddRM 'Source and updates' ($RepoUrl + ' (the interactive calculator built on the same CALC-SPEC is at https://abdu94-hash.github.io/tto-companion-tools/)') 'kv'
    AddRM 'Verification status' '=nVER_Overall' 'kvf'
    AddRM 'Examples status' '=nEX_Status' 'kvf'
    AddRM 'Tested with' 'Microsoft 365 Excel for Windows (version 16.0, build 20430). Formulas use only functions available in Excel 2010 and later (no LET, LAMBDA, dynamic arrays or TEXTJOIN). Not yet verified in Excel for the web, Excel for Mac, Google Sheets or LibreOffice.' 'kv'
    AddRM '' '' ''
    AddRM 'How to use (the workup of Chapter 23)' '' 'head'
    AddRM '1. Cases' 'Open a case version (Case_ID + Case_Version). Follow-up adds a new version; earlier rows stay (V12).' 'kv'
    AddRM '2. Record the sources' 'On Exposures, Administrations, Events and Milestones enter every date, time and timing statement as the source gives it: original text, precision, bounds and source. Never type 0, N/A or a typical value to make a row calculate.' 'kv'
    AddRM '3. Drug_Event_Links' 'Create the drug-event pairs you intend to compare, each with a reason. An interval without a link shows E_MISSING_LINK.' 'kv'
    AddRM '4. Intervals' 'One row per link and anchor. Pick the anchor record (see Timepoints for IDs) and the end point, and the rule: C1 (two timestamps, minutes), C2 (two exact dates, calendar days; C3 treatment day on request) or C4 (bounds). LAST_DOSE_BEFORE_ONSET is selected by C5. Reported timing (C6) sits beside the value. CESSATION to a RECOVERY milestone is a C7 recovery interval. Send a query for every E-code.' 'kv'
    AddRM '5. Product evidence' 'Product_Dossier (one row per extracted statement) and Parameter_Register (value, range, unit, population, half-life type, status, source).' 'kv'
    AddRM '6. Technical arithmetic' 'PK_Exercises (Q1-Q5, Q7-Q11) and PD_Profile (Q6 teaching demonstration). Run only what the question needs, after confirming each gate. Q8 shows a scenario range, never a midpoint or confidence interval. Q11 needs creatinine stability confirmed.' 'kv'
    AddRM '7. Reviewer judgments' 'Temporal_Profile, Biology_Review (four options, each with a rationale; none is a relatedness category), Dechallenge_Rechallenge, Alternatives_Followup.' 'kv'
    AddRM '8. Assessment_QC' 'Narrative, method, unresolved issues and the counts of blocked and flagged interval rows. To print a one-case summary, filter Case_ID on each sheet and print.' 'kv'
    AddRM '' '' ''
    AddRM 'Sheets' '' 'head'
    AddRM 'Read_Me' 'This page.' 'kv'
    AddRM 'Cases' 'Input. One row per case version.' 'kv'
    AddRM 'Exposures' 'Input. One row per exposure episode; Start, Stop and Restart are timepoints (Episode/START, /STOP, /RESTART).' 'kv'
    AddRM 'Administrations' 'Input. One row per dose, administration or infusion rate segment (start, end, rate); dose and BSA for Q10.' 'kv'
    AddRM 'Events' 'Input. One row per event episode; Onset and First detection are timepoints (Episode/ONSET, /DETECTION).' 'kv'
    AddRM 'Milestones' 'Input. Tests, organ function, concentrations, interventions, recovery and other dated records.' 'kv'
    AddRM 'Drug_Event_Links' 'Reviewer. Selected pairs with a reason (no automatic pairing).' 'kv'
    AddRM 'Intervals' 'Derived. One row per link and anchor: value, unit, rule, sequence status, flags, messages, assumptions.' 'kv'
    AddRM 'Product_Dossier' 'Input. Versioned product statements, timing-statement kind, label discrepancies.' 'kv'
    AddRM 'Parameter_Register' 'Input. Parameters with status, population and source; a readiness check.' 'kv'
    AddRM 'PK_Exercises' 'Derived. Q1-Q5 and Q7-Q11 with gates, Parameter_Set, mandatory text and flags.' 'kv'
    AddRM 'PD_Profile' 'Reviewer prompts; Q6 teaching demonstration (switch required; always I_TEACHING_DEMONSTRATION).' 'kv'
    AddRM 'Temporal_Profile, Biology_Review, Dechallenge_Rechallenge, Alternatives_Followup, Assessment_QC' 'Reviewer layer. Judgments in words, with rationales and sources.' 'kv'
    AddRM 'Examples' 'Fictional teaching cases EX-C01, EX-C04 and EX-C28 checked against their home chapters.' 'kv'
    AddRM 'Verification' 'All 25 numeric CALC-SPEC vectors recomputed with PASS/FAIL and an overall status; behavioral vectors V12 and V14 for inspection.' 'kv'
    AddRM 'Timepoints' 'Derived index of every dated record, with ISO 8601 and display forms; use its Record_ID values on Intervals.' 'kv'
    AddRM 'Flags_Messages' 'Every E-, W- and I-code with its verbatim message (CALC-SPEC Section 5).' 'kv'
    AddRM 'Lists' 'Validation lists.' 'kv'
    AddRM '' '' ''
    AddRM 'Conventions' '' 'head'
    AddRM 'Cells' 'Yellow = input (unlocked). Grey = formula (locked). Blue rows at the bottom of a sheet are reserved and locked: they hold the fictional teaching cases and the verification vectors read by the Examples and Verification sheets. Copy them into the yellow rows with a new Case_ID to experiment.' 'kv'
    AddRM 'Date precision' 'TIMESTAMP: Date, local Time and UTC offset (choose unknown if it is not recorded; C1 then uses local clock times with W_TZ_UNKNOWN). DATE: Date only, Time left blank (a date is never given a time). MONTH: Year and Month, expanded to the first and last day, leap years respected. YEAR: 1 January to 31 December. RANGE: earliest and latest dates as stated, with the source. A single date is never imputed from bounds.' 'kv'
    AddRM 'Dates and units' 'Dates are displayed as day month year (10 January 2026); ISO 8601 forms are on Timepoints and Intervals. Units are in the column headers. Only the conversions of CALC-SPEC Section 6 are automatic; weeks, months and years are never converted to days.' 'kv'
    AddRM 'Precision and display' 'Value columns hold unrounded results (G1). Result_Display columns round per CALC-SPEC Section 7: whole minutes, hours to 1 decimal, whole calendar days, percentages to 3 significant figures, ratios to 2 decimals, mg and mg/m2 to 1 decimal, creatinine clearance in whole mL/minute.' 'kv'
    AddRM 'Flags' 'E-codes block a result: the row shows the E-code and no number. W-codes are shown with the result. I-codes annotate it. Messages are verbatim from Flags_Messages; no spreadsheet error text is shown (G7).' 'kv'
    AddRM 'Protection' 'Every sheet is protected without a password (Review > Unprotect Sheet). Filtering is allowed. If you unprotect and change a formula, the Verification sheet no longer documents the workbook you are using.' 'kv'
    AddRM 'Export' 'Export a table by copying it or saving the sheet as CSV UTF-8. Free text that starts with =, +, - or @ must be escaped (Export_Safe columns on Alternatives_Followup and Assessment_QC). Check the round trip (V14). Treatment-day numbers and analytical intervals are never written to E2B fields automatically.' 'kv'
    AddRM '' '' ''
    AddRM 'Implementation decisions (where CALC-SPEC 1.0 or Chapter 23 left a choice)' '' 'head'
    AddRM 'C1 arithmetic' 'Timestamps are converted to whole minutes (date x 1440 + minutes of the day, minus the UTC offset), so intervals are exact integers.' 'kv'
    AddRM 'C2 on the same date' 'When both records have times, C1 decides the order; if an offset is unknown the row also carries W_TZ_UNKNOWN.' 'kv'
    AddRM 'C5 selection' 'An administration is eligible when it is at or before the end point: by UTC time when both times and offsets are known, by local time when an offset is unknown, otherwise by date (for a bounded onset, on or before its earliest date). The latest eligible administration is selected.' 'kv'
    AddRM 'C6 comparison' 'Reported bounds in minutes, hours or days are compared with C1 results (converted to minutes); bounds in calendar days with C2 and C4 results. Weeks, months, years and other pairings are not converted or compared (I_APPROXIMATE_NOT_COMPARED).' 'kv'
    AddRM 'C7 interventions' 'An INTERVENTION milestone counts when it lies within cessation and recovery inclusive (by time when all three are timed, otherwise by date). CALC-SPEC 1.0 has no code for interventions unknown; the row shows a plain note instead.' 'kv'
    AddRM 'Q6 switch' 'With the teaching-demonstration switch off, Q6 shows E_APPLICABILITY_UNCONFIRMED, as in the reference implementation.' 'kv'
    AddRM 'Q9 offsets' 'If any segment or the onset has an unknown offset, local clock times are used and W_TZ_UNKNOWN is shown.' 'kv'
    AddRM 'Q10 Method B' 'Method B uses the BSA recorded at the FIRST or LAST administration, or a STATED value, chosen by the user.' 'kv'
    AddRM 'Q7 unit' 'Q7 accepts days only (CALC-SPEC: t_half and t in days).' 'kv'
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
    $xl.PrintCommunication = $true
    try {
        $wb.BuiltinDocumentProperties.Item('Title').Value = 'TTO Companion Workbook v1.0'
        $wb.BuiltinDocumentProperties.Item('Author').Value = 'Dr. Hafez Selim'
        $wb.BuiltinDocumentProperties.Item('Comments').Value = $VersionLabel + '. Educational tool; CC BY-NC 4.0; ' + $RepoUrl
    } catch { Write-Host "document properties not set" }
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
}
