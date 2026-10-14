<#
.SYNOPSIS
  Jiff Tweaks - free Windows tweaks for gaming and low latency. One file, all code visible.

.DESCRIPTION
  Run it and a window opens with a live dashboard, one-click profiles, a preview of
  the exact old and new value of every setting before you apply it, and one-click
  Restore. Or use it from the console.

  READ THIS FIRST - what it does and does not do
    * It is this one file. Nothing else is downloaded or installed.
    * Before it changes a value it saves the original (C:\JiffTweaks\journal.jsonl).
      Restore puts everything back (removed Store apps can be reinstalled from the Store).
    * The window is a local web page served on 127.0.0.1 (this PC only) with a secret
      token made at launch. It can only start tweaks from the list in this file.
    * No telemetry, no account. The internet is used only if you switch on the ping
      monitor or press "Check for updates", and for the one wallpaper in the branding tweak.

  RUN IT
    Open PowerShell as Administrator, then:
      irm https://raw.githubusercontent.com/jiff-dev/Jiff-free-Tweaks-V9/main/Jiff-Tweaks.ps1 | iex
    or save the file and run:
      powershell -ExecutionPolicy Bypass -File .\Jiff-Tweaks.ps1

  CONSOLE USE
      .\Jiff-Tweaks.ps1 -List                  every tweak and whether it is on
      .\Jiff-Tweaks.ps1 -Code reg_gaming       print the code of one tweak (the window does not show code)
      .\Jiff-Tweaks.ps1 -Preview svc_basic     show what a tweak would change
      .\Jiff-Tweaks.ps1 -Preset safe           safe | competitive | full
      .\Jiff-Tweaks.ps1 -Apply reg_gaming,privacy
      .\Jiff-Tweaks.ps1 -Restore               put every saved original back
      .\Jiff-Tweaks.ps1 -NoGui                 text menu
      .\Jiff-Tweaks.ps1 -SelfTest              checks everything the tweaks need works on this PC
      (add -Yes to skip the question)

  WHERE THINGS ARE IN THIS FILE
    line    77  Options (the parameters) and the program text below
    line    89  Engine: the verbs every tweak uses (Set-Reg, Disable-Svc ...)
    line   405  Engine: the only code that touches Windows (registry, services)
    line   747  Engine: helpers (network, power plans, boot, apps, Defender)
    line  1246  Tweaks: performance
    line  1500  Tweaks: services and scheduled tasks
    line  1627  Tweaks: network
    line  1734  Tweaks: privacy and debloat
    line  1878  Tweaks: GPU
    line  1999  Tweaks: expert lab and branding
    line  2200  Tweaks: restore and the Full profile
    line  2308  Engine: running tweaks, previews, the code view
    line  2548  Engine: live numbers (CPU, GPU, memory, disk, network)
    line  2735  Server: the local web server for the window
    line  3143  Command line: -List, -Apply, -Preset ...
    line  3427  The window itself (HTML, CSS, JavaScript)
    line  4932  Last: administrator rights, 64-bit check, launch

  Needs Windows 10 or 11 and Windows PowerShell 5.1 (or PowerShell 7) as administrator.
  Community: https://discord.gg/fuegoscrims
#>
[CmdletBinding()]
param(
  [string]$Preset,
  [string[]]$Apply,
  [string[]]$Preview,
  [string[]]$Code,
  [switch]$List,
  [switch]$Status,
  [switch]$Restore,
  [switch]$SelfTest,
  [switch]$NoGui,
  [switch]$Yes,
  [switch]$Browser,
  [switch]$NoOpen,
  [switch]$NoElevate,
  [int]$Port = 0
)

# ============================================================================
#  START
# ============================================================================
$JT_SELF_URL = 'https://raw.githubusercontent.com/jiff-dev/Jiff-free-Tweaks-V9/main/Jiff-Tweaks.ps1'

# "-Apply a,b" arrives as one text when the script is started with -File, so split the lists here
if ($Apply)   { $Apply   = @($Apply   | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ }) }
if ($Preview) { $Preview = @($Preview | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ }) }
if ($Code)    { $Code    = @($Code    | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ }) }

# ---------------------------------------------------------------- the program itself, as text (so each background worker can load its own copy)
$JT_ENGINE = @'
# ============================================================================
#  ENGINE - core
#  Everything a tweak can do goes through the small verbs defined here
#  (Set-Reg, Remove-Reg, Disable-Svc, Disable-Task, Invoke-Tool, Invoke-Step),
#  so each tweak is written once and can be APPLIED, CHECKED (is it on?) or
#  PREVIEWED (what would change?) from the same code. Nothing else touches
#  your system.
# ============================================================================

$script:IsWin = ($env:OS -eq 'Windows_NT')
$script:Mock  = ((-not $script:IsWin) -or ($env:JT_MOCK -eq '1'))

if (-not $script:Cfg) {
  $root = $env:JT_ROOT
  if (-not $root) { $root = Join-Path ([string]$env:SystemDrive + '\') 'JiffTweaks' }
  $script:Cfg = @{
    Version   = '9.2'
    Root      = $root
    PlanGuid  = '5a1f6d3e-8b2c-4d7a-9e41-0c3a6b7f1d25'
    ScriptDir = $env:JT_SCRIPTDIR
  }
}
$script:Cfg.LogDir   = Join-Path $script:Cfg.Root 'logs'
$script:Cfg.Journal  = Join-Path $script:Cfg.Root 'journal.jsonl'
$script:Cfg.PowerDir = Join-Path $script:Cfg.Root 'power'

# ---------------------------------------------------------------- state
$script:Tweaks   = New-Object System.Collections.Specialized.OrderedDictionary
$script:Mode     = 'apply'          # apply | check | preview
$script:Hit      = 0                # check mode: how many values already match
$script:Tot      = 0                # check mode: how many values were looked at
$script:Plan     = New-Object System.Collections.ArrayList   # preview mode: planned changes
$script:LogFile  = $null            # set while a run is active
$script:Echo     = $false           # CLI: also print log lines to the console
$script:Seen     = $null            # journal keys already recorded (first record wins)
$script:CntOk = 0; $script:CntFail = 0; $script:CntSkip = 0

# ---------------------------------------------------------------- logging
# Format: [HH:mm:ss] [TAG] message   (the UI reads this live)
function Log {
  param([string]$Tag, [string]$Msg)
  switch ($Tag) { 'OK' { $script:CntOk++ } 'FAIL' { $script:CntFail++ } 'SKIP' { $script:CntSkip++ } }
  $line = '[' + [DateTime]::Now.ToString('HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) + '] [' + $Tag + '] ' + $Msg
  if ($script:LogFile) { try { [IO.File]::AppendAllText($script:LogFile, $line + "`r`n") } catch { } }
  if ($script:Echo) {
    $c = 'Gray'
    switch ($Tag) { 'OK' { $c = 'Green' } 'FAIL' { $c = 'Red' } 'SKIP' { $c = 'DarkGray' } 'STEP' { $c = 'Cyan' } 'SEQ' { $c = 'Cyan' } 'TWEAK' { $c = 'Cyan' } 'NOTE' { $c = 'Yellow' } }
    Write-Host $line -ForegroundColor $c
  }
}

function Initialize-Jiff {
  foreach ($d in @($script:Cfg.Root, $script:Cfg.LogDir, $script:Cfg.PowerDir)) {
    if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Force -Path $d) }
  }
  if ($script:IsWin -and -not $script:Mock) {
    # only administrators and SYSTEM may touch the data folder (it holds the restore journal)
    try { & icacls.exe $script:Cfg.Root /inheritance:r /grant:r '*S-1-5-32-544:(OI)(CI)F' '*S-1-5-18:(OI)(CI)F' 2>&1 | Out-Null } catch { }
  }
}

# ---------------------------------------------------------------- journal (what the original values were)
# One JSON object per line: { k: kind, i: id, t: time, d: original data }.
# The first record for a given kind+id is the true original and is never overwritten.
function Journal-Load {
  $rows = New-Object System.Collections.ArrayList
  if (Test-Path -LiteralPath $script:Cfg.Journal) {
    foreach ($l in [IO.File]::ReadAllLines($script:Cfg.Journal)) {
      if (-not $l.Trim()) { continue }
      try { [void]$rows.Add(($l | ConvertFrom-Json)) } catch { }
    }
  }
  return ,$rows
}

function Journal-Init {
  if ($null -ne $script:Seen) { return }
  $script:Seen = @{}
  foreach ($r in (Journal-Load)) { $script:Seen[($r.k + '|' + $r.i)] = $true }
}

function Journal-Add {
  param([string]$Kind, [string]$Id, $Data)
  Journal-Init
  $key = $Kind + '|' + $Id
  if ($script:Seen.ContainsKey($key)) { return }
  $script:Seen[$key] = $true
  $obj = [ordered]@{ k = $Kind; i = $Id; t = (Get-Date).ToString('s'); d = $Data }
  $json = ConvertTo-Json -InputObject $obj -Compress -Depth 6
  [IO.File]::AppendAllText($script:Cfg.Journal, $json + "`r`n")
}

function Journal-Write {
  param($Rows)
  $sb = New-Object System.Text.StringBuilder
  foreach ($r in $Rows) { [void]$sb.AppendLine((ConvertTo-Json -InputObject $r -Compress -Depth 6)) }
  [IO.File]::WriteAllText($script:Cfg.Journal, $sb.ToString())
  $script:Seen = $null
}

function Journal-Counts {
  $c = @{ reg = 0; svc = 0; task = 0; bcd = 0; pnp = 0; power = 0; net = 0; mm = 0; cmd = 0; other = 0; total = 0 }
  foreach ($r in (Journal-Load)) { if ($c.ContainsKey($r.k)) { $c[$r.k]++ } else { $c.other++ }; $c.total++ }
  return $c
}

# ---------------------------------------------------------------- verbs used inside tweaks
# Set-Reg - registry value:  Set-Reg HKLM 'SOFTWARE\Key' 'Name' DWord 1
#   -Soft   report a protected value as skipped instead of failed
#   -Quiet  count the change but do not write a log line for it (used for hundreds of device flags)
function Set-Reg {
  param([string]$Hive, [string]$Key, [string]$Name, [string]$Type, $Value, [switch]$Soft, [switch]$Quiet)
  $cur = RegRead $Hive $Key $Name
  switch ($script:Mode) {
    'check' {
      $script:Tot++
      if ($cur -and (Test-RegEqual $cur $Type $Value)) { $script:Hit++ }
    }
    'preview' {
      [void]$script:Plan.Add([ordered]@{ kind = 'reg'; target = ($Hive + '\' + $Key); name = $Name; type = $Type
          current = $(if ($cur) { Format-RegValue $cur.Value $cur.Type } else { $null }); new = (Format-RegValue $Value $Type) })
    }
    default {
      try {
        Journal-Add 'reg' ($Hive + '\' + $Key + '|' + $Name) $(if ($cur) { @{ ex = $true; ty = $cur.Type; v = (ConvertTo-JournalValue $cur.Value $cur.Type) } } else { @{ ex = $false } })
        RegWrite $Hive $Key $Name $Type $Value
        if ($Quiet) { $script:CntOk++ } else { Log 'OK' ($Name + ' = ' + (Format-RegValue $Value $Type)) }
      } catch {
        if ($Quiet) { $script:CntSkip++ }
        elseif ($Soft) { Log 'SKIP' ($Name + ' is protected') }
        else { Log 'FAIL' ($Name + ': ' + $_.Exception.Message) }
      }
    }
  }
}

# Remove-Reg - delete a registry value (the original is journaled and can be restored)
function Remove-Reg {
  param([string]$Hive, [string]$Key, [string]$Name)
  $cur = RegRead $Hive $Key $Name
  switch ($script:Mode) {
    'check' { $script:Tot++; if (-not $cur) { $script:Hit++ } }
    'preview' {
      if ($cur) { [void]$script:Plan.Add([ordered]@{ kind = 'reg'; target = ($Hive + '\' + $Key); name = $Name; type = $cur.Type; current = (Format-RegValue $cur.Value $cur.Type); new = '(deleted)' }) }
    }
    default {
      if (-not $cur) { return }
      try {
        Journal-Add 'reg' ($Hive + '\' + $Key + '|' + $Name) @{ ex = $true; ty = $cur.Type; v = (ConvertTo-JournalValue $cur.Value $cur.Type) }
        RegDelete $Hive $Key $Name
        Log 'OK' ('removed ' + $Name)
      } catch { Log 'FAIL' ($Name + ': ' + $_.Exception.Message) }
    }
  }
}

# Disable-Svc - turn off one or more services (the original start type is journaled)
function Disable-Svc {
  param([string[]]$Name)
  foreach ($n in $Name) {
    $cur = SvcGetStart $n
    switch ($script:Mode) {
      'check' { if ($null -ne $cur) { $script:Tot++; if ($cur -eq 4) { $script:Hit++ } } }
      'preview' { if ($null -ne $cur -and $cur -ne 4) { [void]$script:Plan.Add([ordered]@{ kind = 'svc'; target = $n; name = 'Start'; type = 'DWord'; current = $cur; new = 4 }) } }
      default {
        if ($null -eq $cur) { Log 'SKIP' ($n + ' not present'); continue }
        if ($cur -eq 4) { continue }
        try {
          Journal-Add 'svc' $n @{ start = $cur }
          SvcDisable $n
          Log 'OK' ('service ' + $n)
        } catch { Log 'FAIL' ('service ' + $n + ': ' + $_.Exception.Message) }
      }
    }
  }
}

# Disable-Task - turn off scheduled tasks (full task paths)
function Disable-Task {
  param([string[]]$Path)
  foreach ($p in $Path) {
    $cur = TaskGetEnabled $p
    switch ($script:Mode) {
      'check' { if ($null -ne $cur) { $script:Tot++; if (-not $cur) { $script:Hit++ } } }
      'preview' { if ($cur) { [void]$script:Plan.Add([ordered]@{ kind = 'task'; target = $p; name = 'Enabled'; type = 'bool'; current = $true; new = $false }) } }
      default {
        if ($null -eq $cur) { Log 'SKIP' ($p + ' not present'); continue }
        if (-not $cur) { continue }
        try { Journal-Add 'task' $p @{ en = $true }; TaskSetEnabled $p $false; Log 'OK' ('task ' + $p) }
        catch { Log 'FAIL' ('task ' + $p + ': ' + $_.Exception.Message) }
      }
    }
  }
}

# Invoke-Tool - run a command line tool (apply mode only).
#   -Undo  the command that puts things back; it is journaled and run by Restore
#   -Soft  a non-zero exit code is reported as skipped (the setting may not exist on this Windows build)
function Invoke-Tool {
  param([string]$File, [string[]]$Arguments, [string[]]$Undo, [switch]$Soft)
  $text = ($File + ' ' + ($Arguments -join ' ')).Trim()
  switch ($script:Mode) {
    'check' { }
    'preview' { [void]$script:Plan.Add([ordered]@{ kind = 'cmd'; target = $text; name = ''; type = ''; current = $null; new = $null }) }
    default {
      if ($Undo) { Journal-Add 'cmd' $text @{ undo = @($Undo) } }
      $r = Exe $File $Arguments
      if ($r.Code -eq 0) { Log 'OK' $text } elseif ($Soft) { Log 'SKIP' ($text + ' (not supported here)') } else { Log 'FAIL' ($text + ' (exit ' + $r.Code + ')') }
    }
  }
}

# Invoke-Step - any other change. The text says what it does; the block is shown in "View code".
function Invoke-Step {
  param([string]$What, [scriptblock]$Run)
  switch ($script:Mode) {
    'check' { }
    'preview' { [void]$script:Plan.Add([ordered]@{ kind = 'step'; target = $What; name = ''; type = ''; current = $null; new = $null }) }
    default {
      try { & $Run; Log 'OK' $What } catch { Log 'FAIL' ($What + ': ' + $_.Exception.Message) }
    }
  }
}

# Write-Note - a line for the summary at the end of a run
function Write-Note {
  param([string]$Text)
  if ($script:Mode -eq 'apply') { Log 'NOTE' $Text }
}

# Show-Step - progress line for the multi-part Full profile
function Show-Step {
  param([int]$N, [int]$Total, [string]$Title)
  if ($script:Mode -ne 'apply') { return }
  if (Test-Cancel) { throw 'cancelled' }
  Log 'STEP' ([string]$N + '/' + [string]$Total + ' ' + $Title)
}

# Use-Tweak - run another tweak's code in the current mode (apply, check or preview)
function Use-Tweak {
  param([string]$Id)
  $t = $script:Tweaks[$Id]
  if (-not $t) { Log 'FAIL' ('unknown tweak ' + $Id); return }
  & $t.Apply
}

# ---------------------------------------------------------------- value helpers
function Format-RegValue {
  param($Value, [string]$Type)
  if ($null -eq $Value) { return '' }
  if ($Type -eq 'Binary') { if ($Value -is [byte[]]) { return (($Value | ForEach-Object { $_.ToString('X2') }) -join '') } return [string]$Value }
  if ($Type -eq 'MultiString') { return (@($Value) -join '; ') }
  if ($Type -eq 'DWord' -and ($Value -is [int]) -and $Value -lt 0) { return [string]([int64]$Value + 4294967296) }
  return [string]$Value
}

function ConvertTo-JournalValue {
  param($Value, [string]$Type)
  if ($Type -eq 'Binary' -and ($Value -is [byte[]])) { return (($Value | ForEach-Object { $_.ToString('X2') }) -join '') }
  if ($Type -eq 'MultiString') { return @($Value) }
  if ($Type -eq 'DWord' -and ($Value -is [int]) -and $Value -lt 0) { return ([int64]$Value + 4294967296) }
  return $Value
}

function Test-RegEqual {
  param($Cur, [string]$Type, $Want)
  $a = Format-RegValue $Cur.Value $Cur.Type
  $b = Format-RegValue $Want $Type
  if ($Type -eq 'DWord' -or $Type -eq 'QWord') { return ($a -eq $b) }
  return ($a -ieq $b)
}

# ---------------------------------------------------------------- run one tweak in a mode
function Get-TweakResult {
  param($T, [string]$Mode)
  $script:Mode = $Mode; $script:Hit = 0; $script:Tot = 0; $script:Plan = New-Object System.Collections.ArrayList
  $script:LastError = $null
  try { & $T.Apply } catch { $script:LastError = $_.Exception.Message } finally { $script:Mode = 'apply' }
  return @{ Hit = $script:Hit; Tot = $script:Tot; Plan = $script:Plan; Error = $script:LastError }
}

# state of a tweak: on | off | partial | $null (nothing to look at)
function Get-TweakState {
  param($T)
  try {
    if ($T.Check) { $script:Mode = 'check'; return (& $T.Check) }
    $r = Get-TweakResult $T 'check'
    if ($r.Tot -le 0) { return $null }
    $ratio = $r.Hit / $r.Tot
    if ($ratio -ge $T.OnAt) { return 'on' }
    if ($ratio -le $T.OffAt) { return 'off' }
    return 'partial'
  } catch { $script:LastError = $_.Exception.Message; return $null } finally { $script:Mode = 'apply' }
}

function Join-State {
  param($States)
  $s = @($States | Where-Object { $_ })
  if ($s.Count -eq 0) { return $null }
  if (@($s | Where-Object { $_ -eq 'on' }).Count -eq $s.Count) { return 'on' }
  if (@($s | Where-Object { $_ -eq 'off' }).Count -eq $s.Count) { return 'off' }
  return 'partial'
}

# Add-Tweak - register a tweak. Fields: Id Title Page Icon Risk Reboot Short Does Tags Apply Check
#   Note Vendor Covers OnAt OffAt Steps Kind ('tweak' | 'restore' | 'tool')
function Add-Tweak {
  param([hashtable]$T)
  $o = [ordered]@{
    Id = ''; Title = ''; Page = ''; Icon = 'zap'; Risk = 'safe'; Reboot = $false; Short = ''; Does = @(); Tags = @()
    Apply = $null; Check = $null; Note = ''; Vendor = ''; Covers = @(); OnAt = 1.0; OffAt = 0.0; Steps = 0; Kind = 'tweak'
  }
  foreach ($k in $T.Keys) { $o[$k] = $T[$k] }
  $script:Tweaks[$o.Id] = [pscustomobject]$o
}

# ============================================================================
#  ENGINE - platform layer
#  The only code that talks to Windows: registry, services, scheduled tasks and
#  command line tools. Everything else calls these functions.
#  (On a non-Windows machine an in-memory copy is used so the logic can be tested.)
# ============================================================================

# ---------------------------------------------------------------- mock (tests only, never used on Windows)
if ($script:Mock) {
  $m = [AppDomain]::CurrentDomain.GetData('JTMock')
  if (-not $m) {
    $m = [hashtable]::Synchronized(@{
        Reg = [hashtable]::Synchronized(@{}); Keys = [hashtable]::Synchronized(@{}); Task = [hashtable]::Synchronized(@{})
        Bcd = [hashtable]::Synchronized(@{}); Plans = [hashtable]::Synchronized(@{}); Misc = [hashtable]::Synchronized(@{})
      })
    [AppDomain]::CurrentDomain.SetData('JTMock', $m)
    $seed = $true
  } else { $seed = $false }
  $script:MockState = $m
  # folders that exist on every Windows PC (so test runs on other systems do not touch real paths)
  if (-not $env:ProgramData) { $env:ProgramData = Join-Path $script:Cfg.Root 'mock_programdata' }
  if (-not $env:SystemDrive) { $env:SystemDrive = 'C:' }
}

function Get-RegKind {
  param([string]$Type)
  switch ($Type) {
    'DWord' { return [Microsoft.Win32.RegistryValueKind]::DWord }
    'QWord' { return [Microsoft.Win32.RegistryValueKind]::QWord }
    'Binary' { return [Microsoft.Win32.RegistryValueKind]::Binary }
    'MultiString' { return [Microsoft.Win32.RegistryValueKind]::MultiString }
    'ExpandString' { return [Microsoft.Win32.RegistryValueKind]::ExpandString }
    default { return [Microsoft.Win32.RegistryValueKind]::String }
  }
}

function ConvertTo-RegData {
  param([string]$Type, $Value)
  switch ($Type) {
    'DWord' { $n = [int64]$Value; if ($n -gt 2147483647) { $n = $n - 4294967296 }; return [int]$n }
    'QWord' { return [int64]$Value }
    'Binary' {
      if ($Value -is [byte[]]) { return ,$Value }
      $h = ([string]$Value) -replace '[^0-9A-Fa-f]', ''
      $b = New-Object byte[] ([int]($h.Length / 2))
      for ($i = 0; $i -lt $b.Length; $i++) { $b[$i] = [Convert]::ToByte($h.Substring($i * 2, 2), 16) }
      return ,$b
    }
    'MultiString' { return ,([string[]]@($Value)) }
    default { return [string]$Value }
  }
}

# ---------------------------------------------------------------- registry
function Get-JTHive {
  param([string]$Hive)
  $h = [Microsoft.Win32.RegistryHive]::LocalMachine
  if ($Hive -eq 'HKCU') { $h = [Microsoft.Win32.RegistryHive]::CurrentUser }
  elseif ($Hive -eq 'HKU') { $h = [Microsoft.Win32.RegistryHive]::Users }
  elseif ($Hive -eq 'HKCR') { $h = [Microsoft.Win32.RegistryHive]::ClassesRoot }
  return [Microsoft.Win32.RegistryKey]::OpenBaseKey($h, [Microsoft.Win32.RegistryView]::Registry64)
}

function Mock-RegisterKey {
  param([string]$Hive, [string]$Key)
  $parts = $Key.Split('\'); $acc = ''
  foreach ($p in $parts) {
    if ($acc) { $acc = $acc + '\' + $p } else { $acc = $p }
    $script:MockState.Keys[($Hive + '|' + $acc).ToLower()] = ($Hive + '|' + $acc)
  }
}

# returns @{ Type; Value } or $null when the value does not exist
function RegRead {
  param([string]$Hive, [string]$Key, [string]$Name)
  if ($script:Mock) {
    $e = $script:MockState.Reg[($Hive + '|' + $Key + '|' + $Name).ToLower()]
    if ($e) { return @{ Type = $e.Type; Value = $e.Value } }
    return $null
  }
  $base = $null; $sk = $null
  try {
    $base = Get-JTHive $Hive
    $sk = $base.OpenSubKey($Key)
    if (-not $sk) { return $null }
    if (@($sk.GetValueNames()) -notcontains $Name) { return $null }
    $kind = $sk.GetValueKind($Name).ToString()
    $val = $sk.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    return @{ Type = $kind; Value = $val }
  } catch { return $null }
  finally { if ($sk) { $sk.Close() }; if ($base) { $base.Close() } }
}

function RegWrite {
  param([string]$Hive, [string]$Key, [string]$Name, [string]$Type, $Value)
  if ($script:Mock) {
    Mock-RegisterKey $Hive $Key
    $script:MockState.Reg[($Hive + '|' + $Key + '|' + $Name).ToLower()] = @{ Type = $Type; Value = (ConvertTo-RegData $Type $Value) }
    return
  }
  $base = Get-JTHive $Hive
  try {
    $sk = $base.CreateSubKey($Key)
    try { $sk.SetValue($Name, (ConvertTo-RegData $Type $Value), (Get-RegKind $Type)) } finally { $sk.Close() }
  } finally { $base.Close() }
}

function RegDelete {
  param([string]$Hive, [string]$Key, [string]$Name)
  if ($script:Mock) { $script:MockState.Reg.Remove(($Hive + '|' + $Key + '|' + $Name).ToLower()); return }
  $base = Get-JTHive $Hive
  try {
    $sk = $base.OpenSubKey($Key, $true)
    if ($sk) { try { $sk.DeleteValue($Name, $false) } finally { $sk.Close() } }
  } finally { $base.Close() }
}

function RegSubKeys {
  param([string]$Hive, [string]$Key)
  if ($script:Mock) {
    $pre = ($Hive + '|' + $Key + '\').ToLower(); $seen = @{}; $out = @()
    foreach ($k in @($script:MockState.Keys.Keys)) {
      if ($k.StartsWith($pre)) {
        $orig = $script:MockState.Keys[$k]; $rest = $orig.Substring($pre.Length); $seg = $rest.Split('\')[0]
        if ($seg -and -not $seen.ContainsKey($seg.ToLower())) { $seen[$seg.ToLower()] = $true; $out += $seg }
      }
    }
    return ,$out
  }
  $base = $null; $sk = $null
  try {
    $base = Get-JTHive $Hive; $sk = $base.OpenSubKey($Key)
    if (-not $sk) { return ,@() }
    return ,@($sk.GetSubKeyNames())
  } catch { return ,@() }
  finally { if ($sk) { $sk.Close() }; if ($base) { $base.Close() } }
}

# ---------------------------------------------------------------- services
function SvcGetStart {
  param([string]$Name)
  $r = RegRead 'HKLM' ('SYSTEM\CurrentControlSet\Services\' + $Name) 'Start'
  if ($r) { return [int]$r.Value }
  if ($script:Mock -and -not $script:MockState.Misc.ContainsKey('nosvc|' + $Name.ToLower())) { return 3 }
  return $null
}

function SvcSetStart {
  param([string]$Name, [int]$Start)
  try { RegWrite 'HKLM' ('SYSTEM\CurrentControlSet\Services\' + $Name) 'Start' 'DWord' $Start }
  catch {
    # some services are protected from registry writes; sc.exe can still change them
    if ($script:Mock) { throw }
    $word = @{ 0 = 'boot'; 1 = 'system'; 2 = 'auto'; 3 = 'demand'; 4 = 'disabled' }[$Start]
    $r = Exe 'sc.exe' @('config', $Name, 'start=', $word)
    if ($r.Code -ne 0) { throw }
  }
}

function SvcDisable {
  param([string]$Name)
  SvcSetStart $Name 4
  if (-not $script:Mock) { try { & sc.exe stop $Name 2>&1 | Out-Null } catch { } }
}

# ---------------------------------------------------------------- scheduled tasks (COM, works in every language)
$script:TaskSvc = $null
function Get-TaskService {
  if (-not $script:TaskSvc) { $script:TaskSvc = New-Object -ComObject 'Schedule.Service'; $script:TaskSvc.Connect() }
  return $script:TaskSvc
}

function TaskGetEnabled {
  param([string]$Path)
  if ($script:Mock) {
    $v = $script:MockState.Task[$Path]
    if ($null -eq $v) { return $true }
    return [bool]$v
  }
  try {
    $i = $Path.LastIndexOf('\'); $folder = $Path.Substring(0, $i); if (-not $folder) { $folder = '\' }
    $t = (Get-TaskService).GetFolder($folder).GetTask($Path.Substring($i + 1))
    return [bool]$t.Enabled
  } catch { return $null }
}

function TaskSetEnabled {
  param([string]$Path, [bool]$Enabled)
  if ($script:Mock) { $script:MockState.Task[$Path] = $Enabled; return }
  $i = $Path.LastIndexOf('\'); $folder = $Path.Substring(0, $i); if (-not $folder) { $folder = '\' }
  $t = (Get-TaskService).GetFolder($folder).GetTask($Path.Substring($i + 1))
  $t.Enabled = $Enabled
}

# ---------------------------------------------------------------- command line tools
function Exe {
  param([string]$File, [string[]]$Arguments)
  if ($script:Mock) { return (Mock-Exe $File $Arguments) }
  try {
    $out = & $File @Arguments 2>&1
    $code = $LASTEXITCODE; if ($null -eq $code) { $code = 0 }
    return @{ Code = $code; Out = (@($out) -join "`n") }
  } catch { return @{ Code = 1; Out = $_.Exception.Message } }
}

# ---------------------------------------------------------------- mock command line tools
function Mock-Exe {
  param([string]$File, [string[]]$A)
  $st = $script:MockState
  $f = ([IO.Path]::GetFileNameWithoutExtension($File)).ToLower()
  # test hooks: slow every command down, or make one tool fail
  if ($env:JT_MOCK_SLOW) { Start-Sleep -Milliseconds ([int]$env:JT_MOCK_SLOW) }
  if ($env:JT_MOCK_FAIL -and $f -eq $env:JT_MOCK_FAIL) { return @{ Code = 1; Out = 'mock failure' } }
  if ($f -eq 'powercfg') {
    if (-not $st.Plans.ContainsKey('init')) {
      $st.Plans['init'] = $true
      $st.Plans['381b4222-f694-41f0-9685-ff5bb260df2e'] = 'Balanced'
      $st.Plans['8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'] = 'High performance'
      $st.Plans['active'] = '381b4222-f694-41f0-9685-ff5bb260df2e'
    }
    switch -Regex (($A -join ' ')) {
      '^/?-?(getactivescheme|GETACTIVESCHEME)' { $g = $st.Plans['active']; return @{ Code = 0; Out = ('Power Scheme GUID: ' + $g + '  (' + $st.Plans[$g] + ')') } }
      '^/?-?list' {
        $o = @('Existing Power Schemes (* Active)')
        foreach ($k in @($st.Plans.Keys)) { if ($k -match '^[0-9a-f]{8}-') { $o += ('Power Scheme GUID: ' + $k + '  (' + $st.Plans[$k] + ')' + $(if ($k -eq $st.Plans['active']) { ' *' } else { '' })) } }
        return @{ Code = 0; Out = ($o -join "`n") }
      }
      '^/?-?duplicatescheme (\S+) (\S+)' { $st.Plans[$Matches[2]] = 'Copy of ' + $Matches[1]; return @{ Code = 0; Out = '' } }
      '^/?-?changename (\S+) "?([^"]+)"?' { $st.Plans[$Matches[1]] = $Matches[2]; return @{ Code = 0; Out = '' } }
      '^/?-?setactive (\S+)' { $st.Plans['active'] = $Matches[1]; return @{ Code = 0; Out = '' } }
      '^/?-?delete (\S+)' { if ($st.Plans['active'] -eq $Matches[1]) { return @{ Code = 1; Out = 'active' } }; $st.Plans.Remove($Matches[1]); return @{ Code = 0; Out = '' } }
      '^/?-?import (.+) (\S+)$' { $st.Plans[$Matches[2]] = 'Imported'; return @{ Code = 0; Out = '' } }
      default { return @{ Code = 0; Out = '' } }
    }
  }
  if ($f -eq 'fsutil') {
    $t = ($A -join ' ')
    if ($t -match 'behavior set (\S+) (\S+)') { $st.Misc['fsutil|' + $Matches[1].ToLower()] = $Matches[2]; return @{ Code = 0; Out = '' } }
    if ($t -match 'behavior query (\S+)') {
      $n = $Matches[1].ToLower(); $v = $st.Misc['fsutil|' + $n]
      if ($null -eq $v) { $v = @{ memoryusage = '1'; mftzone = '1'; disablelastaccess = '2'; disabledeletenotify = '0'; encryptpagingfile = '0' }[$n] }
      return @{ Code = 0; Out = ($n + ' = ' + $v) }
    }
  }
  if ($f -eq 'bcdedit') {
    $t = ($A -join ' ')
    $id = '{current}'
    if ($t -match '(\{[^}]+\})') { $id = $Matches[1].ToLower(); $t = $t.Replace($Matches[1], '').Replace('  ', ' ') }
    if ($t -match '/set (\S+) (.+)$') { $st.Bcd[$id + '|' + $Matches[1].ToLower()] = $Matches[2].Trim('"'); return @{ Code = 0; Out = '' } }
    if ($t -match '/deletevalue (\S+)') { $st.Bcd.Remove($id + '|' + $Matches[1].ToLower()); return @{ Code = 0; Out = '' } }
    if ($t -match '/enum') { $o = @(); foreach ($k in $st.Bcd.Keys) { if ($k.StartsWith($id + '|')) { $o += ($k.Substring($id.Length + 1) + '   ' + $st.Bcd[$k]) } }; return @{ Code = 0; Out = ($o -join "`n") } }
  }
  return @{ Code = 0; Out = '' }
}

function Initialize-Mock {
  $v = 'HKLM'; $nic = 'SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}'; $gpu = 'SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}'
  RegWrite $v ($nic + '\0007') 'MatchingDeviceId' 'String' 'pci\ven_8086&dev_15f3'
  RegWrite $v ($nic + '\0007') '*EEE' 'String' '1'
  RegWrite $v ($nic + '\0007') 'ITR' 'String' '1'
  RegWrite $v ($nic + '\0007') '*FlowControl' 'String' '3'
  RegWrite $v ($nic + '\0008') 'MatchingDeviceId' 'String' 'root\vmbus'
  RegWrite $v ($gpu + '\0000') 'MatchingDeviceId' 'String' 'pci\ven_10de&dev_2504'
  RegWrite $v ($gpu + '\0001') 'MatchingDeviceId' 'String' 'pci\ven_1002&dev_73ff'
  $ifb = 'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces'
  RegWrite $v ($ifb + '\{11111111-1111-1111-1111-111111111111}') 'EnableDHCP' 'DWord' 1
  RegWrite $v ($ifb + '\{22222222-2222-2222-2222-222222222222}') 'EnableDHCP' 'DWord' 1
  RegWrite $v 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' 'ProductName' 'String' 'Windows 11 Pro'
  RegWrite $v 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' 'DisplayVersion' 'String' '23H2'
  RegWrite $v 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' 'CurrentBuildNumber' 'String' '22631'
  RegWrite 'HKCU' 'Control Panel\Mouse' 'MouseSpeed' 'String' '1'
  RegWrite 'HKCU' 'Control Panel\Mouse' 'MouseThreshold1' 'String' '6'
  RegWrite $v 'SYSTEM\CurrentControlSet\Control\PriorityControl' 'Win32PrioritySeparation' 'DWord' 2
  RegWrite $v 'SYSTEM\CurrentControlSet\Enum\USB\VID_1234&PID_5678\0001\Device Parameters' 'EnableSelectiveSuspend' 'DWord' 1
  RegWrite $v 'SYSTEM\CurrentControlSet\Enum\USB\VID_1234&PID_5678\0001\Device Parameters' 'SelectiveSuspendOn' 'DWord' 0
  RegWrite $v 'SYSTEM\CurrentControlSet\Control\Power' 'HibernateEnabled' 'DWord' 1
  RegWrite 'HKCU' 'Control Panel\Desktop' 'MenuShowDelay' 'String' '400'
  $script:MockState.Misc['appx'] = @('Microsoft.BingWeather', 'Microsoft.Todos', 'Microsoft.XboxApp', 'Clipchamp.Clipchamp', 'Microsoft.WindowsCalculator', 'Microsoft.549981C3F5F10')
  $script:MockState.Bcd['{current}|description'] = 'Windows 11'
  $script:MockState.Bcd['{bootmgr}|description'] = 'Windows Boot Manager'
  $script:MockState.Bcd['{current}|useplatformclock'] = 'Yes'
}
if ($script:Mock -and $seed) { Initialize-Mock }

# ---------------------------------------------------------------- restore (reads the journal and puts originals back)
function Split-RegId {
  param([string]$Id)
  $bar = $Id.LastIndexOf('|'); $path = $Id.Substring(0, $bar); $name = $Id.Substring($bar + 1)
  $sl = $path.IndexOf('\')
  return @{ Hive = $path.Substring(0, $sl); Key = $path.Substring($sl + 1); Name = $name }
}

# one-line description of what restoring a journal row does (used by Preview and the Changes list)
function Describe-JournalRow {
  param($r)
  switch ($r.k) {
    'reg' { $x = Split-RegId $r.i; $now = RegRead $x.Hive $x.Key $x.Name
      return [ordered]@{ kind = 'reg'; target = ($x.Hive + '\' + $x.Key); name = $x.Name; type = $(if ($r.d.ex) { $r.d.ty } else { '' })
        current = $(if ($now) { Format-RegValue $now.Value $now.Type } else { $null }); new = $(if ($r.d.ex) { Format-RegValue $r.d.v $r.d.ty } else { '(deleted)' }) } }
    'svc' { return [ordered]@{ kind = 'svc'; target = $r.i; name = 'Start'; type = 'DWord'; current = (SvcGetStart $r.i); new = [int]$r.d.start } }
    'task' { return [ordered]@{ kind = 'task'; target = $r.i; name = 'Enabled'; type = 'bool'; current = (TaskGetEnabled $r.i); new = $true } }
    default { return [ordered]@{ kind = $r.k; target = [string]$r.i; name = ''; type = ''; current = $null; new = $null } }
  }
}

function Restore-Journal {
  param([string[]]$Kinds)
  $rows = Journal-Load; $keep = New-Object System.Collections.ArrayList; $n = 0
  foreach ($r in $rows) {
    if ($Kinds -notcontains $r.k) { [void]$keep.Add($r); continue }
    try {
      switch ($r.k) {
        'reg' {
          $x = Split-RegId $r.i; $cur = RegRead $x.Hive $x.Key $x.Name
          if ($r.d.ex) {
            if (-not ($cur -and (Test-RegEqual $cur $r.d.ty $r.d.v))) { RegWrite $x.Hive $x.Key $x.Name $r.d.ty $r.d.v; Log 'OK' ('restored ' + $x.Name) }
          } elseif ($cur) { RegDelete $x.Hive $x.Key $x.Name; Log 'OK' ('removed ' + $x.Name) }
        }
        'svc' { SvcSetStart $r.i ([int]$r.d.start); Log 'OK' ('restored service ' + $r.i) }
        'task' { TaskSetEnabled $r.i $true; Log 'OK' ('restored task ' + $r.i) }
        'bcd' {
          $bar = $r.i.LastIndexOf('|'); $bid = $r.i.Substring(0, $bar); $bn = $r.i.Substring($bar + 1)
          if ($r.d.present) { [void](Exe 'bcdedit' @('/set', $bid, $bn, [string]$r.d.value)) } else { [void](Exe 'bcdedit' @('/deletevalue', $bid, $bn)) }
          Log 'OK' ('restored boot setting ' + $bn)
        }
        'pnp' { Restore-Pnp $r.i; Log 'OK' ('restored device ' + $r.i) }
        'net' { Restore-Net $r.i $r.d; Log 'OK' ('restored network ' + $r.i) }
        'mm' { Restore-Mm $r.i $r.d; Log 'OK' ('restored memory setting ' + $r.i) }
        'power' { Restore-Power $r.d; Log 'OK' 'restored power plans' }
        'cmd' { $u = @($r.d.undo); $rr = Exe $u[0] @($u | Select-Object -Skip 1); if ($rr.Code -ne 0) { throw ('exit ' + $rr.Code) }; Log 'OK' ('undid ' + $r.i) }
        'mp' { Restore-Mp $r.d; Log 'OK' 'restored Defender preferences' }
        'label' { Restore-Label $r.d; Log 'OK' ('restored drive label ' + $r.i) }
        'host' { Restore-Host $r.d; Log 'OK' 'restored computer name (applies after restart)' }
        'file' { if (Test-Path -LiteralPath $r.i) { Remove-Item -LiteralPath $r.i -Force }; Log 'OK' ('removed ' + $r.i) }
      }
      $n++
    } catch { Log 'FAIL' ($r.k + ' ' + $r.i + ': ' + $_.Exception.Message); [void]$keep.Add($r) }
  }
  Journal-Write $keep
  return $n
}

# ============================================================================
#  ENGINE - helpers used by the tweaks
#  Each helper reads the current state first, writes the original into the
#  journal, then changes the setting. They all respect the three modes.
# ============================================================================

# ---------------------------------------------------------------- device class keys (network adapters, display adapters)
$script:NicClass = '{4d36e972-e325-11ce-bfc1-08002be10318}'
$script:GpuClass = '{4d36e968-e325-11ce-bfc1-08002be10318}'

function Get-ClassKeys {
  param([string]$ClassGuid, [string]$Pattern)
  $base = 'SYSTEM\CurrentControlSet\Control\Class\' + $ClassGuid
  $out = @()
  foreach ($n in (RegSubKeys 'HKLM' $base)) {
    if ($n -notmatch '^\d{4}$') { continue }
    $id = RegRead 'HKLM' ($base + '\' + $n) 'MatchingDeviceId'
    if ($id -and ([string]$id.Value) -match $Pattern) { $out += ($base + '\' + $n) }
  }
  return ,$out
}

# ---------------------------------------------------------------- Windows network settings
# Values are read through the Windows network cmdlets, so the original can be
# put back, and written with netsh (the same commands every guide uses).
$script:NetMap = @{
  'tcp.AutoTuningLevelLocal'         = @{ A = @('int', 'tcp', 'set', 'global'); K = 'autotuninglevel' }
  'tcp.EcnCapability'                = @{ A = @('int', 'tcp', 'set', 'global'); K = 'ecncapability' }
  'tcp.Timestamps'                   = @{ A = @('int', 'tcp', 'set', 'global'); K = 'timestamps' }
  'tcp.NonSackRttResiliency'         = @{ A = @('int', 'tcp', 'set', 'global'); K = 'nonsackrttresiliency' }
  'tcp.InitialRtoMs'                 = @{ A = @('int', 'tcp', 'set', 'global'); K = 'initialrto' }
  'tcp.MaxSynRetransmissions'        = @{ A = @('int', 'tcp', 'set', 'global'); K = 'maxsynretransmissions' }
  'tcp.ScalingHeuristics'            = @{ A = @('int', 'tcp', 'set', 'heuristics'); K = '' }
  'offload.ReceiveSideScaling'       = @{ A = @('int', 'tcp', 'set', 'global'); K = 'rss' }
  'offload.ReceiveSegmentCoalescing' = @{ A = @('int', 'tcp', 'set', 'global'); K = 'rsc' }
  'teredo.Type'                      = @{ A = @('int', 'teredo', 'set', 'state'); K = '' }
  '6to4.State'                       = @{ A = @('int', '6to4', 'set', 'state'); K = '' }
  'isatap.State'                     = @{ A = @('int', 'isatap', 'set', 'state'); K = '' }
}
$script:NetMockDefaults = @{
  'tcp.AutoTuningLevelLocal' = 'Normal'; 'tcp.EcnCapability' = 'Default'; 'tcp.Timestamps' = 'Disabled'; 'tcp.NonSackRttResiliency' = 'Enabled'
  'tcp.InitialRtoMs' = '3000'; 'tcp.MaxSynRetransmissions' = '2'; 'tcp.ScalingHeuristics' = 'Enabled'
  'offload.ReceiveSideScaling' = 'Enabled'; 'offload.ReceiveSegmentCoalescing' = 'Enabled'
  'teredo.Type' = 'Default'; '6to4.State' = 'Default'; 'isatap.State' = 'Default'
}

function NetGet {
  param([string]$Id)
  if ($script:Mock) {
    $v = $script:MockState.Misc['net|' + $Id]
    if ($null -ne $v) { return [string]$v }
    return $script:NetMockDefaults[$Id]
  }
  $dot = $Id.IndexOf('.'); $group = $Id.Substring(0, $dot); $prop = $Id.Substring($dot + 1)
  try {
    switch ($group) {
      'tcp'     { $o = Get-NetTCPSetting -SettingName InternetCustom -ErrorAction Stop }
      'offload' { $o = Get-NetOffloadGlobalSetting -ErrorAction Stop }
      'teredo'  { $o = Get-NetTeredoConfiguration -ErrorAction Stop }
      '6to4'    { $o = Get-Net6to4Configuration -ErrorAction Stop }
      'isatap'  { $o = Get-NetIsatapConfiguration -ErrorAction Stop }
    }
    $v = $o.$prop
    if ($null -eq $v) { return $null }
    return [string]$v
  } catch { return $null }
}

function NetSet {
  param([string]$Id, [string]$Value)
  if ($script:Mock) { $script:MockState.Misc['net|' + $Id] = $Value; return }
  $m = $script:NetMap[$Id]
  $v = $Value.ToLower()
  if ($Id -eq '6to4.State' -and $v -eq 'automatic') { $v = 'default' }
  $arg = $v; if ($m.K) { $arg = $m.K + '=' + $v }
  $r = Exe 'netsh' (@($m.A) + $arg)
  if ($r.Code -ne 0) { throw ('netsh ' + ($m.A -join ' ') + ' ' + $arg + ' failed') }
}

# Set-NetOpt - one Windows network setting:  Set-NetOpt 'tcp.Timestamps' Disabled
function Set-NetOpt {
  param([string]$Id, [string]$Want)
  $cur = NetGet $Id
  switch ($script:Mode) {
    'check' { if ($null -ne $cur) { $script:Tot++; if ($cur -ieq $Want) { $script:Hit++ } } }
    'preview' { if ($null -ne $cur -and $cur -ine $Want) { [void]$script:Plan.Add([ordered]@{ kind = 'net'; target = $Id; name = ''; type = 'setting'; current = $cur; new = $Want }) } }
    default {
      try {
        if ($null -ne $cur) { Journal-Add 'net' $Id @{ v = $cur } } else { Log 'NOTE' ($Id + ': original not readable, it will not be restored') }
        NetSet $Id $Want
        Log 'OK' ($Id + ' = ' + $Want)
      } catch { Log 'FAIL' ($Id + ': ' + $_.Exception.Message) }
    }
  }
}

function Restore-Net {
  param([string]$Id, $Data)
  NetSet $Id ([string]$Data.v)
}

# ---------------------------------------------------------------- power plans
function Get-PlanGuids {
  $r = Exe 'powercfg' @('/list')
  $out = @()
  foreach ($m in [regex]::Matches([string]$r.Out, '[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}')) { $out += $m.Value.ToLower() }
  return ,$out
}

function Get-ActivePlanGuid {
  $r = Exe 'powercfg' @('/getactivescheme')
  $m = [regex]::Match([string]$r.Out, '[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}')
  if ($m.Success) { return $m.Value.ToLower() }
  return $null
}

function Save-PowerPlans {
  Journal-Init
  if ($script:Seen.ContainsKey('power|plans')) { return }
  $guids = Get-PlanGuids; $active = Get-ActivePlanGuid
  $keep = @($guids | Where-Object { $_ -ne $script:Cfg.PlanGuid })
  foreach ($g in $keep) { [void](Exe 'powercfg' @('/export', (Join-Path $script:Cfg.PowerDir ('plan_' + $g + '.pow')), $g)) }
  Journal-Add 'power' 'plans' @{ active = $active; plans = $keep }
}

function Restore-Power {
  param($D)
  $have = Get-PlanGuids
  foreach ($g in @($D.plans)) {
    if ($have -notcontains $g) {
      $f = Join-Path $script:Cfg.PowerDir ('plan_' + $g + '.pow')
      if (Test-Path -LiteralPath $f) { [void](Exe 'powercfg' @('/import', $f, $g)) }
    }
  }
  if ($D.active) { $r = Exe 'powercfg' @('/setactive', [string]$D.active); if ($r.Code -ne 0) { [void](Exe 'powercfg' @('/restoredefaultschemes')) } }
  if ((Get-ActivePlanGuid) -ne $script:Cfg.PlanGuid) { [void](Exe 'powercfg' @('/delete', $script:Cfg.PlanGuid)) }
}

# create the JiffOS plan: your own JiffOSUltimate.pow if it sits next to the script, else a copy of Ultimate Performance
# returns 'exists' | 'imported' | 'created' | 'failed'
function New-JiffPlan {
  $p = $script:Cfg.PlanGuid
  if ((Get-PlanGuids) -contains $p) { return 'exists' }
  if ($script:Cfg.ScriptDir) {
    $c = Join-Path $script:Cfg.ScriptDir 'JiffOSUltimate.pow'
    if (Test-Path -LiteralPath $c) { $r = Exe 'powercfg' @('/import', $c, $p); if ($r.Code -eq 0) { return 'imported' } }
  }
  $ok = $false
  foreach ($src in @('e9a42b02-d5df-448d-aa00-03f14749eb61', '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c', '381b4222-f694-41f0-9685-ff5bb260df2e')) {
    $r = Exe 'powercfg' @('/duplicatescheme', $src, $p)
    if ($r.Code -eq 0) { $ok = $true; break }
  }
  if (-not $ok) { return 'failed' }
  [void](Exe 'powercfg' @('/changename', $p, 'JiffOS Ultimate', 'Maximum performance, no power saving'))
  return 'created'
}

# processor min/max 100%, boost aggressive, EPP 0, core parking off, cooling active, no standby/hibernate/hybrid sleep,
# disk idle never, PCIe link state off, USB selective suspend off, wireless adapter full performance  (AC only)
$script:PlanTune = @(
  @('54533251-82be-4824-96c1-47b60b740d00', '893dee8e-2bef-41e0-89c6-b55d0929964c', 100)
  @('54533251-82be-4824-96c1-47b60b740d00', 'bc5038f7-23e0-4960-96da-33abaf5935ec', 100)
  @('54533251-82be-4824-96c1-47b60b740d00', 'be337238-0d82-4146-a960-4f3749d470c7', 2)
  @('54533251-82be-4824-96c1-47b60b740d00', '36687f9e-e3a5-4dbf-b1dc-15eb381c6863', 0)
  @('54533251-82be-4824-96c1-47b60b740d00', '0cc5b647-c1df-4637-891a-dec35c318583', 100)
  @('54533251-82be-4824-96c1-47b60b740d00', 'ea062031-0e34-4ff1-9b6d-eb1059334028', 100)
  @('54533251-82be-4824-96c1-47b60b740d00', '94d3a615-a899-4ac5-ae2b-e4d8f634367f', 1)
  @('238c9fa8-0aad-41ed-83f4-97be242c8f20', '29f6c1db-86da-48c5-9fdb-f2b67b1f44da', 0)
  @('238c9fa8-0aad-41ed-83f4-97be242c8f20', '9d7815a6-7ee4-497e-8888-515a05f02364', 0)
  @('238c9fa8-0aad-41ed-83f4-97be242c8f20', '94ac6d29-73ce-41a6-809f-6363ba21b47e', 0)
  @('0012ee47-9041-4b5d-9b77-535fba8b1442', '6738e2c4-e8a5-4a42-b16a-e040e769756e', 0)
  @('501a4d13-42af-4429-9fd1-a8218c268e20', 'ee12f906-d277-404b-b6da-e5fa1a576df5', 0)
  @('2a737441-1930-4402-8d77-b2bebba308a3', '48e6b7a6-50f5-4782-a5d4-53bb8f07e226', 0)
  @('19cbb8fa-5279-450e-9fac-8a3d34e1e29d', '12bbebe6-58d6-4636-95bb-3217ef867c1a', 0)
)

function Set-PlanTuning {
  foreach ($t in $script:PlanTune) { [void](Exe 'powercfg' @('/setacvalueindex', $script:Cfg.PlanGuid, $t[0], $t[1], [string]$t[2])) }
}

# ---------------------------------------------------------------- boot settings (bcdedit)
function Get-BcdValue {
  param([string]$Id, [string]$Name)
  $r = Exe 'bcdedit' @('/enum', $Id)
  $m = [regex]::Match([string]$r.Out, '(?im)^\s*' + [regex]::Escape($Name) + '\s+(.+?)\s*$')
  if (-not $m.Success) { return $null }
  return $m.Groups[1].Value
}

# Set-Bcd - set a boot entry value, or delete it when -Value is $null:  Set-Bcd '{current}' disabledynamictick yes
function Set-Bcd {
  param([string]$Id, [string]$Name, $Value)
  $cur = Get-BcdValue $Id $Name
  switch ($script:Mode) {
    'check' { $script:Tot++; if ($null -eq $Value) { if ($null -eq $cur) { $script:Hit++ } } elseif ($cur -ieq [string]$Value) { $script:Hit++ } }
    'preview' {
      $same = ($null -eq $Value -and $null -eq $cur) -or ($null -ne $Value -and $cur -ieq [string]$Value)
      if (-not $same) { [void]$script:Plan.Add([ordered]@{ kind = 'bcd'; target = ($Id + ' ' + $Name); name = ''; type = 'boot'; current = $cur; new = $(if ($null -eq $Value) { '(deleted)' } else { [string]$Value }) }) }
    }
    default {
      if ($null -eq $Value -and $null -eq $cur) { return }
      Journal-Add 'bcd' ($Id + '|' + $Name) @{ present = ($null -ne $cur); value = $cur }
      if ($null -eq $Value) { $r = Exe 'bcdedit' @('/deletevalue', $Id, $Name) } else { $r = Exe 'bcdedit' @('/set', $Id, $Name, [string]$Value) }
      if ($r.Code -eq 0) { Log 'OK' ('boot ' + $Name + ' = ' + $(if ($null -eq $Value) { '(deleted)' } else { $Value })) } else { Log 'FAIL' ('bcdedit ' + $Name) }
    }
  }
}

# ---------------------------------------------------------------- device power management
function Get-PnpPowerDevices {
  if ($script:Mock) { return ,@() }
  try {
    if (Get-Command Get-WmiObject -ErrorAction SilentlyContinue) { return ,@(Get-WmiObject -Namespace 'root\wmi' -Class MSPower_DeviceEnable -ErrorAction Stop) }
    return ,@(Get-CimInstance -Namespace 'root\wmi' -ClassName MSPower_DeviceEnable -ErrorAction Stop)
  } catch { return ,@() }
}

function Set-PnpPower {
  param($Dev, [bool]$On)
  if ($Dev -is [System.Management.ManagementObject]) { $Dev.Enable = $On; [void]$Dev.psbase.Put() }
  else { Set-CimInstance -InputObject $Dev -Property @{ Enable = $On } }
}

function Restore-Pnp {
  param([string]$Instance)
  foreach ($d in (Get-PnpPowerDevices)) { if ($d.InstanceName -eq $Instance) { Set-PnpPower $d $true } }
}

# "Allow the computer to turn off this device" off for every device
function Disable-PnpPower {
  $n = 0
  foreach ($d in (Get-PnpPowerDevices)) {
    switch ($script:Mode) {
      'check' { $script:Tot++; if (-not $d.Enable) { $script:Hit++ } }
      'preview' { if ($d.Enable) { [void]$script:Plan.Add([ordered]@{ kind = 'pnp'; target = $d.InstanceName; name = 'Allow the computer to turn off this device'; type = 'bool'; current = $true; new = $false }) } }
      default {
        if (-not $d.Enable) { continue }
        try { Journal-Add 'pnp' $d.InstanceName @{ en = $true }; Set-PnpPower $d $false; $n++; $script:CntOk++ }
        catch { Log 'SKIP' ('device ' + $d.InstanceName + ' is protected') }
      }
    }
  }
  if ($script:Mode -eq 'apply') { Log 'INFO' ('device power saving off on ' + $n + ' devices') }
}

# every device registry value with one of these names (USB selective suspend and friends)
function Find-EnumValues {
  param([string[]]$Names)
  $found = New-Object System.Collections.ArrayList
  if ($script:Mock) {
    foreach ($k in @($script:MockState.Reg.Keys)) {
      if ($k -like 'hklm|system\currentcontrolset\enum\*') {
        $parts = $k.Split('|'); $e = $script:MockState.Reg[$k]
        foreach ($nm in $Names) { if ($parts[2] -ieq $nm -and $e.Value -is [int]) { [void]$found.Add(@{ Key = $parts[1]; Name = $nm; Value = $e.Value }) } }
      }
    }
    return ,$found
  }
  $base = Get-JTHive 'HKLM'
  try {
    $root = $base.OpenSubKey('SYSTEM\CurrentControlSet\Enum')
    if ($root) { Walk-Enum $root 'SYSTEM\CurrentControlSet\Enum' $Names $found; $root.Close() }
  } finally { $base.Close() }
  return ,$found
}

function Walk-Enum {
  param($Key, [string]$Path, [string[]]$Names, $Found)
  try {
    foreach ($vn in $Key.GetValueNames()) {
      if ($Names -contains $vn) {
        $v = $Key.GetValue($vn)
        if ($v -is [int]) { [void]$Found.Add(@{ Key = $Path; Name = $vn; Value = $v }) }
      }
    }
    foreach ($sn in $Key.GetSubKeyNames()) {
      $sk = $null
      try { $sk = $Key.OpenSubKey($sn); if ($sk) { Walk-Enum $sk ($Path + '\' + $sn) $Names $Found } } catch { } finally { if ($sk) { $sk.Close() } }
    }
  } catch { }
}

# ---------------------------------------------------------------- memory manager
function Get-MmState {
  if ($script:Mock) {
    $v = $script:MockState.Misc['mm']
    if ($v) { return $v }
    return @{ mc = $true; pc = $true }
  }
  try { $a = Get-MMAgent -ErrorAction Stop; return @{ mc = [bool]$a.MemoryCompression; pc = [bool]$a.PageCombining } } catch { return $null }
}

function Disable-MemoryCompression {
  $cur = Get-MmState
  if ($null -eq $cur) { if ($script:Mode -eq 'apply') { Log 'SKIP' 'memory agent not available' }; return }
  switch ($script:Mode) {
    'check' { $script:Tot += 2; if (-not $cur.mc) { $script:Hit++ }; if (-not $cur.pc) { $script:Hit++ } }
    'preview' {
      if ($cur.mc) { [void]$script:Plan.Add([ordered]@{ kind = 'mm'; target = 'Memory compression'; name = ''; type = 'bool'; current = $true; new = $false }) }
      if ($cur.pc) { [void]$script:Plan.Add([ordered]@{ kind = 'mm'; target = 'Page combining'; name = ''; type = 'bool'; current = $true; new = $false }) }
    }
    default {
      Journal-Add 'mm' 'agent' @{ mc = $cur.mc; pc = $cur.pc }
      if ($script:Mock) { $script:MockState.Misc['mm'] = @{ mc = $false; pc = $false } } else { Disable-MMAgent -MemoryCompression -PageCombining }
      Log 'OK' 'memory compression and page combining off'
    }
  }
}

function Restore-Mm {
  param([string]$Id, $D)
  if ($script:Mock) { $script:MockState.Misc['mm'] = @{ mc = [bool]$D.mc; pc = [bool]$D.pc }; return }
  if ($D.mc) { Enable-MMAgent -MemoryCompression }
  if ($D.pc) { Enable-MMAgent -PageCombining }
}

# ---------------------------------------------------------------- file system (fsutil)
# Set-Fsutil - fsutil behavior set NAME VALUE, with the old value recorded for Restore
function Set-Fsutil {
  param([string]$Name, [int]$Value)
  $q = Exe 'fsutil' @('behavior', 'query', $Name)
  $m = [regex]::Match([string]$q.Out, '=\s*(\d+)')
  if ($script:Mode -eq 'check') { if ($m.Success) { $script:Tot++; if ([int]$m.Groups[1].Value -eq $Value) { $script:Hit++ } }; return }
  $undo = $null
  if ($m.Success) { $undo = @('fsutil', 'behavior', 'set', $Name, $m.Groups[1].Value) }
  Invoke-Tool 'fsutil' @('behavior', 'set', $Name, [string]$Value) -Undo $undo
}

# ---------------------------------------------------------------- scheduled tasks by name pattern (e.g. NVIDIA telemetry)
function Find-Tasks {
  param([string]$Folder, [string]$Like)
  $out = @()
  if ($script:Mock) { return ,$out }
  try {
    $f = (Get-TaskService).GetFolder($Folder)
    foreach ($t in $f.GetTasks(0)) { if ($t.Name -like $Like) { $out += ($Folder.TrimEnd('\') + '\' + $t.Name) } }
  } catch { }
  return ,$out
}

# ---------------------------------------------------------------- AppX
$script:AppxAll = $null; $script:AppxProv = $null; $script:AppxAt = [datetime]::MinValue
function Load-Appx {
  param([switch]$Force)
  if (-not $Force -and $null -ne $script:AppxAll -and ((Get-Date) - $script:AppxAt).TotalSeconds -lt 90) { return }
  if ($script:Mock) {
    $names = @($script:MockState.Misc['appx'])
    $script:AppxAll = @($names | Where-Object { $_ } | ForEach-Object { [pscustomobject]@{ Name = $_; PackageFullName = ($_ + '_1.0') } })
    $script:AppxProv = @()
  } else {
    if ($PSVersionTable.PSEdition -eq 'Core') { try { Import-Module Appx -UseWindowsPowerShell -WarningAction SilentlyContinue -ErrorAction Stop } catch { } }
    $script:AppxAll = @(Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue)
    $script:AppxProv = @(Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue)
  }
  $script:AppxAt = Get-Date
}

# Remove-Apps - uninstall Store apps for all users and from the Windows image
function Remove-Apps {
  param([string[]]$Names)
  if ($script:Mode -eq 'apply') { Load-Appx -Force } else { Load-Appx }
  foreach ($n in $Names) {
    $pk = @($script:AppxAll | Where-Object { $_.Name -eq $n })
    $pv = @($script:AppxProv | Where-Object { $_.DisplayName -eq $n })
    $has = ($pk.Count + $pv.Count) -gt 0
    switch ($script:Mode) {
      'check' { $script:Tot++; if (-not $has) { $script:Hit++ } }
      'preview' { if ($has) { [void]$script:Plan.Add([ordered]@{ kind = 'app'; target = $n; name = ''; type = 'app'; current = 'installed'; new = '(removed)' }) } }
      default {
        if (-not $has) { Log 'SKIP' ($n + ' not installed'); continue }
        if ($script:Mock) { $script:MockState.Misc['appx'] = @($script:MockState.Misc['appx'] | Where-Object { $_ -ne $n }); Log 'OK' ('removed ' + $n); continue }
        foreach ($p in $pk) { try { Remove-AppxPackage -Package $p.PackageFullName -AllUsers -ErrorAction Stop } catch { try { Remove-AppxPackage -Package $p.PackageFullName -ErrorAction SilentlyContinue } catch { } } }
        foreach ($p in $pv) { try { [void](Remove-AppxProvisionedPackage -Online -PackageName $p.PackageName -ErrorAction SilentlyContinue) } catch { } }
        Log 'OK' ('removed ' + $n)
      }
    }
  }
  if ($script:Mode -eq 'apply') { $script:AppxAt = [datetime]::MinValue }
}

# ---------------------------------------------------------------- Windows Defender preferences
$script:MpProps = @('DisableRealtimeMonitoring', 'DisableBehaviorMonitoring', 'DisableIOAVProtection', 'DisableScriptScanning', 'DisableBlockAtFirstSeen', 'MAPSReporting', 'SubmitSamplesConsent')

function Get-MpState {
  if ($script:Mock) {
    $v = $script:MockState.Misc['mp']
    if ($v) { return $v }
    return @{ DisableRealtimeMonitoring = $false; DisableBehaviorMonitoring = $false; DisableIOAVProtection = $false; DisableScriptScanning = $false; DisableBlockAtFirstSeen = $false; MAPSReporting = 2; SubmitSamplesConsent = 1 }
  }
  try { $p = Get-MpPreference -ErrorAction Stop; $h = @{}; foreach ($n in $script:MpProps) { $h[$n] = $p.$n }; return $h } catch { return $null }
}

function Set-MpPref {
  param([hashtable]$Want)
  $cur = Get-MpState
  switch ($script:Mode) {
    'check' { if ($cur) { foreach ($k in $Want.Keys) { $script:Tot++; if ($cur[$k] -eq $Want[$k]) { $script:Hit++ } } } }
    'preview' { if ($cur) { foreach ($k in $Want.Keys) { if ($cur[$k] -ne $Want[$k]) { [void]$script:Plan.Add([ordered]@{ kind = 'mp'; target = 'Defender'; name = $k; type = 'setting'; current = [string]$cur[$k]; new = [string]$Want[$k] }) } } } }
    default {
      if (-not $cur) { Log 'SKIP' 'Defender preferences not available'; return }
      Journal-Add 'mp' 'prefs' $cur
      if ($script:Mock) { $m = @{}; foreach ($k in $cur.Keys) { $m[$k] = $cur[$k] }; foreach ($k in $Want.Keys) { $m[$k] = $Want[$k] }; $script:MockState.Misc['mp'] = $m }
      else { Set-MpPreference @Want -ErrorAction SilentlyContinue }
      $now = Get-MpState; $bad = 0
      foreach ($k in $Want.Keys) { if ($now[$k] -ne $Want[$k]) { $bad++ } }
      if ($bad -eq 0) { Log 'OK' 'Defender real-time, behavior, script and cloud protection off' }
      else { Log 'FAIL' ('Defender ignored ' + $bad + ' setting(s) - turn Tamper Protection off in Windows Security first') }
    }
  }
}

function Restore-Mp {
  param($D)
  $h = @{}
  foreach ($n in $script:MpProps) { if ($null -ne $D.$n) { $h[$n] = $D.$n } }
  if ($script:Mock) { $script:MockState.Misc['mp'] = $h; return }
  Set-MpPreference @h -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------- branding helpers
function Set-Wallpaper {
  param([string]$Path)
  if ($script:Mock) { return }
  if (-not ('JT.Native' -as [type])) {
    Add-Type -TypeDefinition 'using System.Runtime.InteropServices; namespace JT { public class Native { [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)] public static extern bool SystemParametersInfo(int a, int b, string c, int d); } }'
  }
  [void][JT.Native]::SystemParametersInfo(20, 0, $Path, 3)
}

function New-JiffLogo {
  param([string]$Path)
  if ($script:Mock) { [IO.File]::WriteAllText($Path, 'bmp'); return }
  Draw-JiffLogo $Path
}

function Draw-JiffLogo {
  param([string]$Path)
  Add-Type -AssemblyName System.Drawing
  $b = New-Object System.Drawing.Bitmap 120, 120
  $g = [System.Drawing.Graphics]::FromImage($b)
  $g.SmoothingMode = 'AntiAlias'; $g.TextRenderingHint = 'AntiAlias'
  $g.Clear([System.Drawing.Color]::FromArgb(14, 14, 14))
  $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(0, 200, 83), 6)
  $g.DrawEllipse($pen, 8, 8, 104, 104)
  $f = New-Object System.Drawing.Font('Segoe UI', 48, [System.Drawing.FontStyle]::Bold)
  $br = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(0, 200, 83))
  $sf = New-Object System.Drawing.StringFormat
  $sf.Alignment = 'Center'; $sf.LineAlignment = 'Center'
  $g.DrawString('J', $f, $br, (New-Object System.Drawing.RectangleF(0, 2, 120, 120)), $sf)
  $b.Save($Path, [System.Drawing.Imaging.ImageFormat]::Bmp)
  $g.Dispose(); $b.Dispose()
}

function Get-DriveLabel {
  param([string]$Drive)
  if ($script:Mock) { $v = $script:MockState.Misc['label|' + $Drive]; if ($null -ne $v) { return $v }; return 'Windows' }
  try { return ([IO.DriveInfo]($Drive + ':\')).VolumeLabel } catch { return $null }
}

function Set-DriveLabel {
  param([string]$Drive, [string]$Label)
  if ($script:Mock) { $script:MockState.Misc['label|' + $Drive] = $Label; return }
  ([IO.DriveInfo]($Drive + ':\')).VolumeLabel = $Label
}

function Restore-Label { param($D) Set-DriveLabel ([string]$D.drive) ([string]$D.label) }

function Get-HostName {
  if ($script:Mock) { $v = $script:MockState.Misc['host']; if ($v) { return $v }; return 'DESKTOP-MOCK' }
  return $env:COMPUTERNAME
}

function Set-HostName {
  param([string]$Name)
  if ($script:Mock) { $script:MockState.Misc['host'] = $Name; return }
  Rename-Computer -NewName $Name -Force -ErrorAction Stop
}

function Restore-Host { param($D) Set-HostName ([string]$D.name) }

# remember that the tweaks created this file, so Restore can remove it again
function Register-CreatedFile {
  param([string]$Path)
  Journal-Add 'file' $Path @{ created = $true }
}

# ---------------------------------------------------------------- restore as a tweak action (apply = do it, preview = list it)
function Invoke-RestoreKinds {
  param([string[]]$Kinds, [string]$What)
  switch ($script:Mode) {
    'check' { }
    'preview' { foreach ($r in (Journal-Load)) { if ($Kinds -contains $r.k) { [void]$script:Plan.Add((Describe-JournalRow $r)) } } }
    default {
      $n = @(Restore-Journal $Kinds)[-1]
      Log 'OK' ('restored ' + $n + ' ' + $What)
    }
  }
}

# ============================================================================
#  TWEAKS - performance: registry, power plan, device power, boot timer
#  Each tweak is one block of code. The same block is used to apply it,
#  to check whether it is on, and to preview what it would change.
# ============================================================================

Add-Tweak @{
  Id = 'restore_point'; Title = 'Create restore point'; Page = 'restore'; Icon = 'database'; Risk = 'safe'; Reboot = $false; Kind = 'tool'
  Short = 'Make a Windows restore point you can roll back to from Windows itself.'
  Tags = @('System Restore')
  Does = @('Turns on System Restore for the system drive if it is off', 'Creates a restore point named Jiff Tweaks', 'Windows allows one restore point per 24 hours by default')
  Check = { $null }
  Apply = {
    Invoke-Step 'Create a Windows restore point' {
      if ($script:Mock) { return }
      $cmd = "Enable-ComputerRestore -Drive ('" + $env:SystemDrive + "\'); Checkpoint-Computer -Description 'Jiff Tweaks' -RestorePointType MODIFY_SETTINGS -ErrorAction Stop"
      $r = Exe 'powershell.exe' @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $cmd)
      if ($r.Code -ne 0) { throw 'Windows did not create a restore point (it allows one every 24 hours)' }
    }
  }
}

Add-Tweak @{
  Id = 'reg_system'; Title = 'Latency and scheduler'; Page = 'optimize'; Icon = 'zap'; Risk = 'moderate'; Reboot = $true
  Short = 'MMCSS game priority, GPU scheduling, timers, memory and shutdown timeouts.'
  Tags = @('MMCSS', 'HAGS', 'TDR', 'Timers')
  Does = @(
    'MMCSS: game task priority, network throttling off'
    'Win32PrioritySeparation 0x26, IRQ8 priority'
    'Hardware GPU scheduling on, TDR timeouts, MPO off'
    'Power throttling off, Fast Startup off, Connected Standby off'
    'Paging executive and paging combining off, timer resolution requests'
    'Shutdown and hung-app timeouts set to 1 second'
    'USB selective suspend off, NTFS last-access updates off'
  )
  Apply = {
    $sp = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile'
    $gd = 'SYSTEM\CurrentControlSet\Control\GraphicsDrivers'
    $mm = 'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
    $kr = 'SYSTEM\CurrentControlSet\Control\Session Manager\kernel'
    $pw = 'SYSTEM\CurrentControlSet\Control\Power'
    $ex = 'Software\Microsoft\Windows\CurrentVersion'

    # MMCSS: give games the CPU and GPU first, stop throttling network traffic
    Set-Reg HKLM $sp NetworkThrottlingIndex DWord 0xffffffff
    Set-Reg HKLM $sp SystemResponsiveness DWord 0
    Set-Reg HKLM $sp AlwaysOn DWord 1
    Set-Reg HKLM $sp NoLazyMode DWord 1
    Set-Reg HKLM ($sp + '\Tasks\Games') 'GPU Priority' DWord 8
    Set-Reg HKLM ($sp + '\Tasks\Games') Priority DWord 6
    Set-Reg HKLM ($sp + '\Tasks\Games') 'Scheduling Category' String High
    Set-Reg HKLM ($sp + '\Tasks\Games') 'SFIO Priority' String High
    Set-Reg HKLM ($sp + '\Tasks\Games') 'Latency Sensitive' String True
    Set-Reg HKLM ($sp + '\Tasks\Games') 'Background Only' String False

    # scheduler: short, fixed quantum with foreground boost (0x26)
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Control\PriorityControl' Win32PrioritySeparation DWord 38
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Control\PriorityControl' IRQ8Priority DWord 1

    # GPU: hardware scheduling on, longer TDR timeouts, multi-plane overlay off
    Set-Reg HKLM $gd HwSchMode DWord 2
    Set-Reg HKLM $gd TdrDelay DWord 10
    Set-Reg HKLM $gd TdrDdiDelay DWord 20
    Set-Reg HKLM $gd DpiMapIommuContiguous DWord 1
    Set-Reg HKLM 'SOFTWARE\Microsoft\Windows\Dwm' OverlayTestMode DWord 5
    Set-Reg HKCU 'SOFTWARE\Microsoft\DirectX\UserGpuPreferences' DirectXUserGlobalSettings String 'VRROptimizeEnable=0;SwapEffectUpgradeEnable=1;'

    # GPU power and latency tolerance
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Services\DXGKrnl' MonitorLatencyTolerance DWord 1
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Services\DXGKrnl' MonitorRefreshLatencyTolerance DWord 1
    foreach ($n in 'DefaultD3TransitionLatencyActivelyUsed', 'DefaultD3TransitionLatencyIdleLongTime', 'DefaultD3TransitionLatencyIdleMonitorOff', 'DefaultD3TransitionLatencyIdleNoContext',
      'DefaultD3TransitionLatencyIdleShortTime', 'DefaultD3TransitionLatencyIdleVeryLongTime', 'DefaultLatencyToleranceIdle0', 'DefaultLatencyToleranceIdle0MonitorOff',
      'DefaultLatencyToleranceIdle1', 'DefaultLatencyToleranceIdle1MonitorOff', 'DefaultLatencyToleranceMemory', 'DefaultLatencyToleranceNoContext',
      'DefaultLatencyToleranceNoContextMonitorOff', 'DefaultLatencyToleranceOther', 'DefaultLatencyToleranceTimerPeriod', 'DefaultMemoryRefreshLatencyToleranceActivelyUsed',
      'DefaultMemoryRefreshLatencyToleranceMonitorOff', 'DefaultMemoryRefreshLatencyToleranceNoContext', 'Latency', 'MonitorLatencyTolerance', 'MonitorRefreshLatencyTolerance',
      'TransitionLatency') { Set-Reg HKLM ($gd + '\Power') $n DWord 1 }

    # power
    Set-Reg HKLM ($pw + '\PowerThrottling') PowerThrottlingOff DWord 1
    Set-Reg HKLM $pw HiberbootEnabled DWord 0
    Set-Reg HKLM $pw CsEnabled DWord 0
    Set-Reg HKLM $pw EnergyEstimationEnabled DWord 0
    Set-Reg HKLM $pw EventProcessorEnabled DWord 0

    # memory
    Set-Reg HKLM $mm DisablePagingExecutive DWord 1
    Set-Reg HKLM $mm LargeSystemCache DWord 0
    Set-Reg HKLM $mm DisablePagingCombining DWord 1

    # timers
    Set-Reg HKLM $kr GlobalTimerResolutionRequests DWord 1
    Set-Reg HKLM $kr DistributeTimers DWord 1

    # file system and USB
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Control\FileSystem' NtfsDisableLastAccessUpdate DWord 1
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Services\USB' DisableSelectiveSuspend DWord 1

    # shutdown and hung-app timeouts
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Control' WaitToKillServiceTimeout String 1000
    Set-Reg HKCU 'Control Panel\Desktop' WaitToKillAppTimeout String 1000
    Set-Reg HKCU 'Control Panel\Desktop' HungAppTimeout String 1000
    Set-Reg HKCU 'Control Panel\Desktop' AutoEndTasks String 1

    # no startup delay for desktop apps
    Set-Reg HKCU ($ex + '\Explorer\Serialize') StartupDelayInMSec DWord 0
    Set-Reg HKCU ($ex + '\Explorer\Serialize') WaitforIdleState DWord 0

    # background apps and the activity feed
    Set-Reg HKCU ($ex + '\BackgroundAccessApplications') GlobalUserDisabled DWord 1
    Set-Reg HKCU ($ex + '\Search') BackgroundAppGlobalToggle DWord 0
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\AppPrivacy' LetAppsRunInBackground DWord 2
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\System' EnableActivityFeed DWord 0
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\System' PublishUserActivities DWord 0
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\System' UploadUserActivities DWord 0
  }
}

Add-Tweak @{
  Id = 'reg_gaming'; Title = 'Gaming'; Page = 'optimize'; Icon = 'gamepad'; Risk = 'safe'; Reboot = $false
  Short = 'Game DVR off, fullscreen behavior and Game Mode.'
  Tags = @('Game DVR', 'Fullscreen', 'Game Mode')
  Does = @('Game DVR and background capture off', 'Fullscreen optimizations follow the per-app setting', 'Game Mode on, Game Bar startup panel off')
  Apply = {
    $gc = 'System\GameConfigStore'
    Set-Reg HKCU $gc GameDVR_Enabled DWord 0
    Set-Reg HKCU $gc GameDVR_FSEBehaviorMode DWord 2
    Set-Reg HKCU $gc GameDVR_HonorUserFSEBehaviorMode DWord 1
    Set-Reg HKCU $gc GameDVR_DXGIHonorFSEWindowsCompatible DWord 1
    Set-Reg HKCU $gc GameDVR_EFSEFeatureFlags DWord 0
    Set-Reg HKCU 'SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR' AppCaptureEnabled DWord 0
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\GameDVR' AllowGameDVR DWord 0
    Set-Reg HKCU 'SOFTWARE\Microsoft\GameBar' AllowAutoGameMode DWord 1
    Set-Reg HKCU 'SOFTWARE\Microsoft\GameBar' AutoGameModeEnabled DWord 1
    Set-Reg HKCU 'SOFTWARE\Microsoft\GameBar' ShowStartupPanel DWord 0
    Set-Reg HKCU 'SOFTWARE\Microsoft\GameBar' UseNexusForGameBarEnabled DWord 0
  }
}

Add-Tweak @{
  Id = 'reg_input_ui'; Title = 'Input and UI'; Page = 'optimize'; Icon = 'mouse'; Risk = 'safe'; Reboot = $false
  Short = '1:1 mouse input and snappier menus.'
  Tags = @('No acceleration', 'Menu delay 0')
  Does = @('Pointer acceleration off', 'Menu show delay and hover time 0', 'Window and taskbar animations and transparency off')
  Apply = {
    $ms = 'Control Panel\Mouse'
    Set-Reg HKCU $ms MouseSpeed String 0
    Set-Reg HKCU $ms MouseThreshold1 String 0
    Set-Reg HKCU $ms MouseThreshold2 String 0
    Set-Reg HKCU $ms MouseSensitivity String 10
    Set-Reg HKCU $ms MouseHoverTime String 0
    Set-Reg HKU '.DEFAULT\Control Panel\Mouse' MouseSpeed String 0
    Set-Reg HKU '.DEFAULT\Control Panel\Mouse' MouseThreshold1 String 0
    Set-Reg HKU '.DEFAULT\Control Panel\Mouse' MouseThreshold2 String 0
    Set-Reg HKCU 'Control Panel\Desktop' MenuShowDelay String 0
    Set-Reg HKCU 'Control Panel\Desktop\WindowMetrics' MinAnimate String 0
    Set-Reg HKCU 'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' EnableTransparency DWord 0
    Set-Reg HKCU 'Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' TaskbarAnimations DWord 0
  }
}

Add-Tweak @{
  Id = 'reg_priority'; Title = 'CPU and I/O priority'; Page = 'optimize'; Icon = 'cpu'; Risk = 'moderate'; Reboot = $true
  Short = 'Keep background tasks out of the way of your game.'
  Tags = @('Indexer', 'TrustedInstaller', 'Foreground boost')
  Does = @('Passive interrupt worker priority 18', 'Search Indexer, TrustedInstaller and wuauclt run at idle priority', 'Foreground boost decay off')
  Apply = {
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Control\Session Manager\I/O System' PassiveIntRealTimeWorkerPriority DWord 18
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Control\KernelVelocity' DisableFGBoostDecay DWord 1
    $ifeo = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options'
    foreach ($exe in 'SearchIndexer.exe', 'TrustedInstaller.exe', 'wuauclt.exe') {
      Set-Reg HKLM ($ifeo + '\' + $exe + '\PerfOptions') CpuPriorityClass DWord 1
      Set-Reg HKLM ($ifeo + '\' + $exe + '\PerfOptions') IoPriority DWord 0
    }
  }
}

Add-Tweak @{
  Id = 'reg_all'; Title = 'All registry tweaks'; Page = 'optimize'; Icon = 'layers'; Risk = 'moderate'; Reboot = $true
  Short = 'Latency, gaming, input and priority tweaks in one go.'
  Tags = @('MMCSS', 'HAGS', 'Timers', 'Input', 'Priority')
  Does = @('Latency and scheduler', 'Gaming', 'Input and UI', 'CPU and I/O priority')
  Covers = @('reg_system', 'reg_gaming', 'reg_input_ui', 'reg_priority')
  Apply = {
    Use-Tweak reg_system
    Use-Tweak reg_gaming
    Use-Tweak reg_input_ui
    Use-Tweak reg_priority
  }
}

Add-Tweak @{
  Id = 'power_plan'; Title = 'JiffOS power plan'; Page = 'optimize'; Icon = 'power'; Risk = 'moderate'; Reboot = $false
  Short = 'Maximum-performance plan for AC power. Your plans are exported first.'
  Tags = @('Ultimate', 'Core parking off', 'No sleep')
  Does = @(
    'Built from Ultimate Performance, or imports JiffOSUltimate.pow when it sits next to the script'
    'Processor 100% min and max, boost aggressive, core parking off'
    'Disk, USB, PCIe and Wi-Fi power saving off on AC'
    'Hibernate and Fast Startup off'
  )
  Check = {
    if ((Get-ActivePlanGuid) -eq $script:Cfg.PlanGuid) { return 'on' }
    if ((Get-PlanGuids) -contains $script:Cfg.PlanGuid) { return 'partial' }
    return 'off'
  }
  Apply = {
    Invoke-Step 'Save your current power plans (Restore brings them back)' { Save-PowerPlans }
    Invoke-Step 'Create the JiffOS Ultimate plan' {
      $script:PlanSrc = New-JiffPlan
      if ($script:PlanSrc -eq 'failed') { throw 'could not create the power plan' }
    }
    Invoke-Step 'Tune the plan: 100% CPU, core parking off, no disk, USB, PCIe or Wi-Fi power saving' {
      if ($script:PlanSrc -eq 'failed') { throw 'skipped, the plan was not created' }
      if ($script:PlanSrc -ne 'imported') { Set-PlanTuning }
    }
    Invoke-Step 'Make the JiffOS plan the active plan' {
      if ($script:PlanSrc -eq 'failed') { throw 'skipped, the plan was not created' }
      $r = Exe 'powercfg' @('/setactive', $script:Cfg.PlanGuid)
      if ($r.Code -ne 0) { throw 'powercfg /setactive failed' }
    }
    Invoke-Tool powercfg @('/hibernate', 'off') -Undo @('powercfg', '/hibernate', 'on')
  }
}

Add-Tweak @{
  Id = 'devpower'; Title = 'USB and device power saving'; Page = 'optimize'; Icon = 'plug'; Risk = 'moderate'; Reboot = $true
  Short = 'Stop Windows from suspending devices to save power.'
  Tags = @('USB', 'PnP')
  Does = @('Selective suspend and idle power flags cleared on every device', '"Allow the computer to turn off this device" off', 'Original values are saved')
  Apply = {
    $names = 'EnhancedPowerManagementEnabled', 'AllowIdleIrpInD3', 'EnableSelectiveSuspend', 'DeviceSelectiveSuspended', 'SelectiveSuspendEnabled', 'SelectiveSuspendOn',
    'WaitWakeEnabled', 'D3ColdSupported', 'WdfDirectedPowerTransitionEnable', 'EnableIdlePowerManagement', 'IdleInWorkingState'
    $n = 0
    foreach ($f in (Find-EnumValues $names)) {
      if ($script:Mode -ne 'check' -and $f.Value -eq 0) { continue }
      Set-Reg HKLM $f.Key $f.Name DWord 0 -Soft -Quiet
      $n++
    }
    if ($script:Mode -eq 'apply') { Log 'INFO' ('USB and device power flags cleared: ' + $n) }
    Disable-PnpPower
  }
}

Add-Tweak @{
  Id = 'bcd_timer'; Title = 'Boot timer'; Page = 'optimize'; Icon = 'clock'; Risk = 'moderate'; Reboot = $true
  Short = 'Dynamic tick off and the TSC timer.'
  Tags = @('bcdedit')
  Does = @('disabledynamictick set to yes', 'Platform clock override removed (uses TSC)')
  Apply = {
    Set-Bcd '{current}' useplatformclock $null
    Set-Bcd '{current}' disabledynamictick yes
  }
}

# ============================================================================
#  TWEAKS - services and scheduled tasks
#  Services are turned off by setting their start type to Disabled.
#  The original start type of every service is saved first.
#  Never touched: SamSs, Themes, TokenBroker, Defender and Security Center
#  (Defender is only touched by the Expert lab).
# ============================================================================

Add-Tweak @{
  Id = 'svc_basic'; Title = 'Basic services'; Page = 'services'; Icon = 'sliders'; Risk = 'safe'; Reboot = $false; OnAt = 0.9; OffAt = 0.1
  Short = 'Telemetry, maps, search indexing, SysMain and time sync.'
  Tags = @('DiagTrack', 'SysMain', 'WSearch')
  Does = @('Diagnostics, delivery optimization, push messaging', 'Maps, geolocation, offline files, retail demo', 'SysMain (Superfetch), Windows Search, Windows Time')
  Apply = {
    Disable-Svc 'DoSvc', 'diagsvc', 'DPS', 'dmwappushservice', 'MapsBroker', 'lfsvc', 'CscService', 'SEMgrSvc', 'PhoneSvc', 'RemoteRegistry', 'RetailDemo', 'SysMain', 'WalletService', 'WSearch', 'W32Time'
    Disable-Svc 'MessagingService', 'WpnUserService'
  }
}

Add-Tweak @{
  Id = 'svc_standard'; Title = 'Standard services'; Page = 'services'; Icon = 'sliders'; Risk = 'moderate'; Reboot = $true; OnAt = 0.9; OffAt = 0.1
  Short = 'Basic tier plus BITS, Remote Desktop, RAS and error reporting.'
  Tags = @('BITS', 'RDP', 'WerSvc')
  Does = @('Everything in Basic', 'BITS, Remote Desktop, RAS, Error Reporting, Performance Logs', 'Original start types are saved')
  Covers = @('svc_basic')
  Apply = {
    Use-Tweak svc_basic
    Disable-Svc 'AJRouter', 'tzautoupdate', 'BITS', 'DiagTrack', 'CDPSvc', 'DusmSvc', 'WdiServiceHost', 'WdiSystemHost', 'DisplayEnhancementService', 'fhsvc', 'HomeGroupListener', 'HomeGroupProvider', 'SmsRouter', 'pla', 'WpcMonSvc'
    Disable-Svc 'RasAuto', 'RasMan', 'SessionEnv', 'TermService', 'RpcLocator', 'WerSvc'
    Disable-Svc 'BcastDVRUserService'
  }
}

Add-Tweak @{
  Id = 'svc_xbox'; Title = 'Xbox services'; Page = 'services'; Icon = 'gamepad'; Risk = 'moderate'; Reboot = $false; OnAt = 0.9; OffAt = 0.1
  Short = 'Xbox Live auth, game save and networking services.'
  Tags = @('Game Pass', 'Xbox')
  Does = @('XboxGipSvc, xbgm, XblAuthManager, XblGameSave, XboxNetApiSvc', 'Game Pass and Xbox sign-in stop working')
  Apply = {
    Disable-Svc 'XboxGipSvc', 'xbgm', 'XblAuthManager', 'XblGameSave', 'XboxNetApiSvc'
  }
}

Add-Tweak @{
  Id = 'svc_advanced'; Title = 'Advanced services'; Page = 'services'; Icon = 'sliders'; Risk = 'aggressive'; Reboot = $true; OnAt = 0.9; OffAt = 0.1
  Short = 'About 150 services plus per-user templates. The most aggressive tier.'
  Tags = @('Update', 'Printing', 'Bluetooth', 'VPN', 'Hyper-V')
  Does = @(
    'Turns off Windows Update, printing, Bluetooth, VPN, file sharing, Remote Desktop, Hyper-V and WSL2, Store licensing'
    'Per-user service templates (Bluetooth, CDP, OneSync, ...) disabled'
    'SamSs, Themes, TokenBroker and Defender are always kept'
    'Original start type of every service is saved'
  )
  Covers = @('svc_standard', 'svc_basic', 'svc_xbox')
  Apply = {
    Use-Tweak svc_standard
    Use-Tweak svc_xbox

    # telemetry, diagnostics
    Disable-Svc 'diagnosticshub.standardcollector.service', 'IEEtwCollectorService', 'PcaSvc', 'wmiApSrv'
    # update, delivery, store
    Disable-Svc 'wuauserv', 'WaaSMedicSvc', 'UsoSvc', 'InstallService', 'ClipSVC', 'LicenseManager', 'AppXSvc', 'WpnService', 'wlidsvc'
    # search, fonts, shell, storage
    Disable-Svc 'FontCache', 'FontCache3.0.0.0', 'tiledatamodelsvc', 'TabletInputService', 'ShellHWDetection', 'StorSvc', 'TieringEngineService', 'defragsvc', 'TrkWks'
    # network, sharing, remote
    Disable-Svc 'AppMgmt', 'ALG', 'PeerDistSvc', 'fdPHost', 'FDResPub', 'lltdsvc', 'p2psvc', 'p2pimsvc', 'PNRPsvc', 'PNRPAutoReg', 'iphlpsvc', 'IpxlatCfgSvc', 'PolicyAgent', 'SharedAccess',
      'lmhosts', 'LanmanServer', 'LanmanWorkstation', 'WebClient', 'QWAVE', 'wcncsvc', 'WFDSConMgrSvc', 'icssvc', 'WMPNetworkSvc', 'spectrum'
    Disable-Svc 'RemoteAccess', 'UmRdpService', 'WinRM', 'Wecsvc', 'SNMPTRAP', 'MSiSCSI', 'irmon'
    # virtualization
    Disable-Svc 'HvHost', 'hns', 'vmickvpexchange', 'vmicguestinterface', 'vmicshutdown', 'vmicheartbeat', 'vmicvmsession', 'vmicrdv', 'vmictimesync', 'vmicvss', 'AppVClient', 'UevAgentService'
    # bluetooth, sensors, devices, printing, smart cards
    Disable-Svc 'BTAGService', 'bthserv', 'BthHFSrv', 'WPDBusEnum', 'Spooler', 'PrintNotify', 'SCardSvr', 'ScDeviceEnum', 'SCPolicySvc', 'StiSvc', 'WbioSrvc', 'FrameServer', 'SensorDataService',
      'SensrSvc', 'SensorService', 'NaturalAuthentication', 'RmSvc', 'TapiSrv', 'SDRSVC', 'wisvc'
    # security, enterprise, misc
    Disable-Svc 'BDESVC', 'wbengine', 'VaultSvc', 'CertPropSvc', 'EFS', 'EntAppSvc', 'AssignedAccessManagerSvc', 'MSDTC', 'smphost', 'seclogon', 'shpamsvc', 'SgrmBroker', 'WEPHOSTSVC',
      'SharedRealitySvc', 'perceptionsimulation'
    # per-user service templates
    Disable-Svc 'BluetoothUserService', 'CDPUserSvc', 'CaptureService', 'ConsentUxUserSvc', 'PimIndexMaintenanceSvc', 'DevicePickerUserSvc', 'DevicesFlowUserSvc', 'PrintWorkflowUserSvc', 'OneSyncSvc', 'UserDataSvc', 'UnistoreSvc'
  }
}

Add-Tweak @{
  Id = 'tasks_disable'; Title = 'Scheduled task cleanup'; Page = 'services'; Icon = 'clock'; Risk = 'safe'; Reboot = $false; OnAt = 0.9; OffAt = 0.1
  Short = 'Disable telemetry, CEIP, error reporting and maps tasks.'
  Tags = @('About 35 tasks')
  Does = @('Compatibility Appraiser, CEIP, Disk Diagnostics, Feedback', 'Maps, Family Safety, Work Folders, Settings sync', 'Only enabled tasks are touched, and each is recorded for Restore')
  Apply = {
    Disable-Task @(
      '\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser'
      '\Microsoft\Windows\Application Experience\ProgramDataUpdater'
      '\Microsoft\Windows\Application Experience\StartupAppTask'
      '\Microsoft\Windows\Application Experience\PcaPatchDbTask'
      '\Microsoft\Windows\Application Experience\MareBackup'
      '\Microsoft\Windows\Autochk\Proxy'
      '\Microsoft\Windows\Customer Experience Improvement Program\Consolidator'
      '\Microsoft\Windows\Customer Experience Improvement Program\UsbCeip'
      '\Microsoft\Windows\Customer Experience Improvement Program\KernelCeipTask'
      '\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticDataCollector'
      '\Microsoft\Windows\DiskDiagnostic\Microsoft-Windows-DiskDiagnosticResolver'
      '\Microsoft\Windows\DiskFootprint\Diagnostics'
      '\Microsoft\Windows\Feedback\Siuf\DmClient'
      '\Microsoft\Windows\Feedback\Siuf\DmClientOnScenarioDownload'
      '\Microsoft\Windows\Windows Error Reporting\QueueReporting'
      '\Microsoft\Windows\Maps\MapsToastTask'
      '\Microsoft\Windows\Maps\MapsUpdateTask'
      '\Microsoft\Windows\Shell\FamilySafetyMonitor'
      '\Microsoft\Windows\Shell\FamilySafetyRefreshTask'
      '\Microsoft\Windows\Power Efficiency Diagnostics\AnalyzeSystem'
      '\Microsoft\Windows\Flighting\FeatureConfig\ReconcileFeatures'
      '\Microsoft\Windows\Flighting\OneSettings\RefreshCache'
      '\Microsoft\Windows\Device Information\Device'
      '\Microsoft\Windows\NetTrace\GatherNetworkInfo'
      '\Microsoft\Windows\PI\Sqm-Tasks'
      '\Microsoft\Windows\Location\Notifications'
      '\Microsoft\Windows\Location\WindowsActionDialog'
      '\Microsoft\Windows\WDI\ResolutionHost'
      '\Microsoft\XblGameSave\XblGameSaveTask'
      '\Microsoft\Windows\Mobile Broadband Accounts\MNO Metadata Parser'
      '\Microsoft\Windows\Work Folders\Work Folders Logon Synchronization'
      '\Microsoft\Windows\Work Folders\Work Folders Maintenance Work'
      '\Microsoft\Windows\Windows Media Sharing\UpdateLibrary'
      '\Microsoft\Windows\SettingSync\BackgroundUploadTask'
      '\Microsoft\Windows\SettingSync\NetworkStateChangeTask'
    )
  }
}

# ============================================================================
#  TWEAKS - network
#  Windows network settings are read through the network cmdlets before they
#  are changed, so Restore can put the exact original values back.
# ============================================================================

Add-Tweak @{
  Id = 'net_global'; Title = 'TCP globals'; Page = 'network'; Icon = 'wifi'; Risk = 'safe'; Reboot = $false; OnAt = 0.9; OffAt = 0.2
  Short = 'Autotuning normal, RSC off, ECN off, timestamps off.'
  Tags = @('netsh', 'RSS', 'RSC')
  Does = @(
    'Autotuning normal, RSS on, RSC off, ECN off, timestamps off, heuristics off'
    'Teredo, 6to4 and ISATAP off'
    'Network reservation 0%, Delivery Optimization peer sharing off'
  )
  Note = 'DCA, UDP receive offload and the neighbor cache size are set but cannot be read back, so Restore leaves them alone.'
  Apply = {
    Set-NetOpt 'tcp.AutoTuningLevelLocal' Normal
    Set-NetOpt 'offload.ReceiveSideScaling' Enabled
    Set-NetOpt 'offload.ReceiveSegmentCoalescing' Disabled
    Set-NetOpt 'tcp.EcnCapability' Disabled
    Set-NetOpt 'tcp.Timestamps' Disabled
    Set-NetOpt 'tcp.NonSackRttResiliency' Disabled
    Set-NetOpt 'tcp.InitialRtoMs' 2000
    Set-NetOpt 'tcp.MaxSynRetransmissions' 2
    Set-NetOpt 'tcp.ScalingHeuristics' Disabled
    Set-NetOpt 'teredo.Type' Disabled
    Set-NetOpt '6to4.State' Disabled
    Set-NetOpt 'isatap.State' Disabled
    Invoke-Tool netsh @('int', 'tcp', 'set', 'global', 'dca=enabled') -Soft
    Invoke-Tool netsh @('int', 'udp', 'set', 'global', 'uro=disabled') -Soft
    Invoke-Tool netsh @('int', 'ip', 'set', 'global', 'neighborcachelimit=4096') -Soft

    $tp = 'SYSTEM\CurrentControlSet\Services\Tcpip'
    Set-Reg HKLM ($tp + '\Parameters') TcpTimedWaitDelay DWord 30
    Set-Reg HKLM ($tp + '\ServiceProvider') LocalPriority DWord 4
    Set-Reg HKLM ($tp + '\ServiceProvider') HostsPriority DWord 5
    Set-Reg HKLM ($tp + '\ServiceProvider') DnsPriority DWord 6
    Set-Reg HKLM ($tp + '\ServiceProvider') NetbtPriority DWord 7
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\Psched' NonBestEffortLimit DWord 0
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization' DODownloadMode DWord 0
  }
}

Add-Tweak @{
  Id = 'net_interfaces'; Title = 'Nagle off'; Page = 'network'; Icon = 'wifi'; Risk = 'safe'; Reboot = $true
  Short = 'TcpAckFrequency and TCPNoDelay on every network interface.'
  Tags = @('TcpAckFrequency', 'TCPNoDelay')
  Does = @('Writes the values under each interface GUID, not the parent key', 'Original values are saved before they are changed')
  Apply = {
    $base = 'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces'
    foreach ($guid in (RegSubKeys HKLM $base)) {
      Set-Reg HKLM ($base + '\' + $guid) TcpAckFrequency DWord 1
      Set-Reg HKLM ($base + '\' + $guid) TCPNoDelay DWord 1
      Set-Reg HKLM ($base + '\' + $guid) TcpDelAckTicks DWord 0
    }
  }
}

Add-Tweak @{
  Id = 'nic_tweaks'; Title = 'Adapter properties'; Page = 'network'; Icon = 'plug'; Risk = 'moderate'; Reboot = $true; OnAt = 0.95; OffAt = 0.2
  Short = 'EEE, Green Ethernet, flow control, interrupt moderation and LSO off.'
  Tags = @('EEE', 'LSO', 'Wake')
  Does = @('Only values your adapter already has are changed', 'Power saving and wake-on-LAN off', 'Takes effect after a restart or when the adapter is re-enabled')
  Apply = {
    # property names that start with * are standard keywords from the driver
    $props = '*EEE', 'EEE', 'AdvancedEEE', 'EEELinkAdvertisement', 'EnableGreenEthernet', 'GigaLite', 'PowerSavingMode', 'ULPMode', 'AutoDisableGigabit', 'EnablePME',
    '*FlowControl', '*InterruptModeration', 'ITR', '*LsoV1IPv4', '*LsoV2IPv4', '*LsoV2IPv6', '*PMARPOffload', '*PMNSOffload',
    '*WakeOnMagicPacket', '*WakeOnPattern', '*ModernStandbyWoLMagicPacket', 'S5WakeOnLan', 'WakeOnLink'
    foreach ($key in (Get-ClassKeys $script:NicClass '^pci\\ven_')) {
      # 24 = "allow the computer to turn off this device" and "allow this device to wake the computer" both off
      Set-Reg HKLM $key PnPCapabilities DWord 24
      foreach ($p in $props) {
        if (RegRead HKLM $key $p) { Set-Reg HKLM $key $p String 0 }
      }
    }
  }
}

Add-Tweak @{
  Id = 'net_all'; Title = 'Full network tuning'; Page = 'network'; Icon = 'wifi'; Risk = 'moderate'; Reboot = $true
  Short = 'TCP settings, Nagle off on every interface and adapter properties.'
  Tags = @('TCP', 'Nagle', 'NIC')
  Does = @('TCP globals', 'Nagle off per interface', 'Adapter properties and power saving')
  Covers = @('net_global', 'net_interfaces', 'nic_tweaks')
  Apply = {
    Use-Tweak net_global
    Use-Tweak net_interfaces
    Use-Tweak nic_tweaks
  }
}

Add-Tweak @{
  Id = 'net_reset'; Title = 'Network reset'; Page = 'network'; Icon = 'refresh'; Risk = 'moderate'; Reboot = $true; Kind = 'tool'
  Short = 'Winsock and IP stack reset plus a DNS flush.'
  Tags = @('winsock', 'ip reset')
  Does = @('netsh winsock reset, netsh int ip reset, ipconfig /flushdns', 'Clears static IP settings. VPN adapters may need re-adding')
  Note = 'A reset cannot be undone by Restore.'
  Check = { $null }
  Apply = {
    Invoke-Tool netsh @('winsock', 'reset')
    Invoke-Tool netsh @('int', 'ip', 'reset')
    Invoke-Tool ipconfig @('/flushdns')
    Write-Note 'Network stack reset. Restart required.'
  }
}

# ============================================================================
#  TWEAKS - privacy and debloat
#  Removed Store apps are not covered by Restore (reinstall them from the
#  Microsoft Store). Everything else here is journaled.
# ============================================================================

Add-Tweak @{
  Id = 'privacy'; Title = 'Privacy and telemetry'; Page = 'cleanup'; Icon = 'eye'; Risk = 'safe'; Reboot = $false; OnAt = 0.9
  Short = 'Telemetry, ads, suggestions and error reporting off.'
  Tags = @('Telemetry', 'Ads', 'Cortana')
  Does = @(
    'Telemetry policy 0, feedback prompts and OneSettings downloads off'
    'Bing search, activity feed, advertising ID and tailored experiences off'
    'Windows Error Reporting off'
    'Consumer features and suggested apps off'
    'Telemetry scheduled tasks disabled'
  )
  Apply = {
    $dc = 'SOFTWARE\Policies\Microsoft\Windows\DataCollection'
    $cv = 'Software\Microsoft\Windows\CurrentVersion'

    # telemetry
    Set-Reg HKLM $dc AllowTelemetry DWord 0
    Set-Reg HKLM 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection' AllowTelemetry DWord 0
    Set-Reg HKLM $dc LimitDiagnosticLogCollection DWord 1
    Set-Reg HKLM $dc DisableOneSettingsDownloads DWord 1
    Set-Reg HKLM $dc DoNotShowFeedbackNotifications DWord 1

    # search and Cortana
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\Windows Search' AllowCortana DWord 0
    Set-Reg HKCU ($cv + '\Search') BingSearchEnabled DWord 0
    Set-Reg HKCU ($cv + '\Search') CortanaConsent DWord 0
    Set-Reg HKCU ($cv + '\Search') HistoryViewEnabled DWord 0
    Set-Reg HKCU ($cv + '\Search') DeviceHistoryEnabled DWord 0
    Set-Reg HKCU 'Software\Policies\Microsoft\Windows\Explorer' DisableSearchBoxSuggestions DWord 1

    # advertising, location, ink
    Set-Reg HKCU ($cv + '\AdvertisingInfo') Enabled DWord 0
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo' DisabledByGroupPolicy DWord 1
    Set-Reg HKCU ($cv + '\Privacy') TailoredExperiencesWithDiagnosticDataEnabled DWord 0
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors' DisableLocation DWord 1
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\WindowsInkWorkspace' AllowWindowsInkWorkspace DWord 0

    # error reporting
    $we = 'SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting'
    Set-Reg HKLM $we Disabled DWord 1
    Set-Reg HKLM $we DoReport DWord 0
    Set-Reg HKLM $we LoggingDisabled DWord 1
    Set-Reg HKLM 'SOFTWARE\Microsoft\Windows\Windows Error Reporting' Disabled DWord 1

    # consumer features and suggestions
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\CloudContent' DisableWindowsConsumerFeatures DWord 1
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\CloudContent' DisableSoftLanding DWord 1
    foreach ($n in 'SystemPaneSuggestionsEnabled', 'SilentInstalledAppsEnabled', 'PreInstalledAppsEnabled', 'OemPreInstalledAppsEnabled', 'ContentDeliveryAllowed', 'SubscribedContentEnabled',
      'FeatureManagementEnabled', 'SoftLandingEnabled', 'RotatingLockScreenOverlayEnabled', 'SubscribedContent-310093Enabled', 'SubscribedContent-338387Enabled', 'SubscribedContent-338388Enabled',
      'SubscribedContent-338389Enabled', 'SubscribedContent-338393Enabled', 'SubscribedContent-353694Enabled', 'SubscribedContent-353696Enabled', 'SubscribedContent-353698Enabled') {
      Set-Reg HKCU ($cv + '\ContentDeliveryManager') $n DWord 0
    }

    # recent files
    Set-Reg HKCU ($cv + '\Explorer') ShowFrequent DWord 0
    Set-Reg HKCU ($cv + '\Explorer') ShowRecent DWord 0
    Set-Reg HKCU ($cv + '\Policies\Explorer') NoRecentDocsHistory DWord 1

    Use-Tweak tasks_disable
  }
}

Add-Tweak @{
  Id = 'debloat_curated'; Title = 'Consumer apps'; Page = 'cleanup'; Icon = 'trash'; Risk = 'moderate'; Reboot = $false; OnAt = 0.95
  Short = 'Weather, Maps, Solitaire, People, Feedback Hub and similar.'
  Tags = @('AppX')
  Does = @('Removes about 22 preinstalled apps for all users', 'Includes Xbox overlay and identity packages', 'Can be reinstalled from the Microsoft Store. Not covered by Restore')
  Apply = {
    Remove-Apps 'Microsoft.BingWeather', 'Microsoft.GetHelp', 'Microsoft.Getstarted', 'Microsoft.MicrosoftOfficeHub', 'Microsoft.MicrosoftSolitaireCollection', 'Microsoft.MixedReality.Portal',
      'Microsoft.People', 'Microsoft.SkypeApp', 'Microsoft.Wallet', 'Microsoft.WindowsAlarms', 'Microsoft.WindowsFeedbackHub', 'Microsoft.WindowsMaps', 'Microsoft.WindowsSoundRecorder',
      'Microsoft.YourPhone', 'Microsoft.ZuneMusic', 'Microsoft.ZuneVideo', 'Microsoft.Xbox.TCUI', 'Microsoft.XboxApp', 'Microsoft.XboxGameOverlay', 'Microsoft.XboxGamingOverlay',
      'Microsoft.XboxIdentityProvider', 'Microsoft.XboxSpeechToTextOverlay'
  }
}

Add-Tweak @{
  Id = 'debloat_extended'; Title = 'Extended apps'; Page = 'cleanup'; Icon = 'trash'; Risk = 'aggressive'; Reboot = $false; OnAt = 0.95
  Short = 'Teams, Copilot, Clipchamp, Bing apps, Mail and Calendar and more.'
  Tags = @('Teams', 'Copilot', 'Clipchamp')
  Does = @('Everything in Consumer apps', 'Removes about 25 more apps including Teams, Copilot, Clipchamp, Mail, Dev Home and the Bing apps', 'Game Pass and Xbox sign-in stop working', 'Can be reinstalled from the Microsoft Store. Not covered by Restore')
  Covers = @('debloat_curated')
  Apply = {
    Use-Tweak debloat_curated
    Invoke-Step 'Close Teams, Cortana and Widgets' {
      if ($script:Mock) { return }
      Stop-Process -Name 'Teams', 'ms-teams', 'Cortana', 'Widgets' -Force -ErrorAction SilentlyContinue
    }
    Remove-Apps 'Microsoft.549981C3F5F10', 'Microsoft.BingNews', 'Microsoft.BingSearch', 'Microsoft.BingFinance', 'Microsoft.BingSports', 'Microsoft.BingTranslator', 'Microsoft.GamingApp',
      'Microsoft.MicrosoftJournal', 'Microsoft.MicrosoftMahjong', 'Microsoft.MicrosoftCasualGames', 'Microsoft.Office.OneNote', 'Microsoft.OutlookForWindows', 'Microsoft.Todos',
      'MicrosoftCorporationII.MicrosoftFamily', 'microsoft.windowscommunicationsapps', 'Microsoft.PowerAutomateDesktop', 'MicrosoftCorporationII.WindowsSubsystemForAndroid',
      'Clipchamp.Clipchamp', 'Microsoft.Teams', 'MicrosoftTeams', 'MSTeams', 'Microsoft.Copilot', 'Microsoft.Windows.DevHome', 'MicrosoftWindows.Client.WebExperience'
  }
}

Add-Tweak @{
  Id = 'debloat_onedrive'; Title = 'Uninstall OneDrive'; Page = 'cleanup'; Icon = 'trash'; Risk = 'moderate'; Reboot = $false
  Short = 'Removes OneDrive. Your OneDrive folder is kept.'
  Tags = @('OneDrive')
  Does = @('Uninstalls OneDrive and blocks sync by policy', 'Your OneDrive folder is only removed if it is empty')
  Note = 'The uninstall itself is not covered by Restore. The sync policy and the startup entry are.'
  Check = {
    $p = RegRead HKLM 'SOFTWARE\Policies\Microsoft\Windows\OneDrive' DisableFileSyncNGSC
    $blocked = ($p -and [int]$p.Value -eq 1)
    $present = $false
    if (-not $script:Mock -and $env:LOCALAPPDATA) {
      foreach ($f in (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDrive.exe'), (Join-Path $env:ProgramFiles 'Microsoft OneDrive\OneDrive.exe')) { if (Test-Path -LiteralPath $f) { $present = $true } }
    }
    if ($blocked -and -not $present) { return 'on' }
    if ($blocked -or -not $present) { return 'partial' }
    return 'off'
  }
  Apply = {
    Invoke-Step 'Close OneDrive' {
      if ($script:Mock) { return }
      Stop-Process -Name 'OneDrive' -Force -ErrorAction SilentlyContinue
    }
    Invoke-Step 'Run the OneDrive uninstaller' {
      if ($script:Mock) { return }
      foreach ($f in (Join-Path $env:SystemRoot 'SysWOW64\OneDriveSetup.exe'), (Join-Path $env:SystemRoot 'System32\OneDriveSetup.exe')) {
        if (Test-Path -LiteralPath $f) { Start-Process -FilePath $f -ArgumentList '/uninstall' -Wait }
      }
    }
    Invoke-Step 'Remove the OneDrive folder only if it is empty (your files are never deleted)' {
      if ($script:Mock) { return }
      $d = Join-Path $env:USERPROFILE 'OneDrive'
      if (Test-Path -LiteralPath $d) { try { [IO.Directory]::Delete($d, $false) } catch { } }
    }
    Invoke-Step 'Remove leftover OneDrive program data' {
      if ($script:Mock) { return }
      foreach ($d in (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive'), (Join-Path $env:ProgramData 'Microsoft OneDrive')) {
        if ($d -and (Test-Path -LiteralPath $d)) { Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction SilentlyContinue }
      }
    }
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\OneDrive' DisableFileSyncNGSC DWord 1
    Remove-Reg HKCU 'Software\Microsoft\Windows\CurrentVersion\Run' OneDrive
  }
}

# ============================================================================
#  TWEAKS - GPU (NVIDIA and AMD driver settings)
#  Display adapters are found in the registry by vendor id (10DE = NVIDIA,
#  1002 = AMD); only the keys of those adapters are written.
# ============================================================================

Add-Tweak @{
  Id = 'nvidia_all'; Title = 'NVIDIA'; Page = 'gpu'; Icon = 'monitor'; Risk = 'moderate'; Reboot = $true; Vendor = 'nvidia'
  Short = 'HDCP off, P-state lock and telemetry off.'
  Tags = @('HDCP', 'P-state', 'Telemetry')
  Does = @('HDCP off (some protected video may not play)', 'P-state lock to the highest state, higher idle power', 'NVIDIA telemetry service and tasks off')
  Apply = {
    foreach ($key in (Get-ClassKeys $script:GpuClass 'ven_10de')) {
      Set-Reg HKLM $key RMHdcpKeyglobZero DWord 1
      Set-Reg HKLM $key DisableDynamicPstate DWord 1
    }
    Set-Reg HKLM 'SOFTWARE\NVIDIA Corporation\NvControlPanel2\Client' OptInOrOutPreference DWord 0
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Services\nvlddmkm\Global\Startup' SendTelemetryData DWord 0
    Disable-Svc 'NvTelemetryContainer'
    Disable-Task @((Find-Tasks '\' 'NvTm*') + (Find-Tasks '\' 'NvDriverUpdateCheck*') + (Find-Tasks '\' 'NvProfileUpdater*'))
    Write-Note 'NVIDIA tweaks applied. Restart required.'
  }
}

Add-Tweak @{
  Id = 'amd_all'; Title = 'AMD'; Page = 'gpu'; Icon = 'monitor'; Risk = 'moderate'; Reboot = $true; Vendor = 'amd'
  Short = 'Adrenalin, ReLive, Radeon Boost, Chill, Anti-Lag and driver defaults.'
  Tags = @('Adrenalin', 'ReLive', 'KMD', 'UMD')
  Does = @('Adrenalin auto-update, tray and toasts off', 'ReLive and DVR off', 'Radeon Boost, Chill, Anti-Lag and USU off', 'Driver UMD and DXVA defaults, AMD crash and events services off')
  Apply = {
    # Adrenalin software
    $cn = 'Software\AMD\CN'
    Set-Reg HKCU $cn AutoUpdateTriggered DWord 0
    Set-Reg HKCU $cn PowerSaverAutoEnable_CUR DWord 0
    Set-Reg HKCU $cn BuildType DWord 0
    Set-Reg HKCU $cn WizardProfile String PROFILE_CUSTOM
    Set-Reg HKCU $cn UserTypeWizardShown DWord 1
    Set-Reg HKCU $cn AutoUpdate DWord 0
    Set-Reg HKCU $cn RSXBrowserUnavailable String true
    Set-Reg HKCU $cn SystemTray String false
    Set-Reg HKCU $cn AllowWebContent String false
    Set-Reg HKCU $cn CN_Hide_Toast_Notification String true
    Set-Reg HKCU $cn AnimationEffect String false
    Set-Reg HKCU ($cn + '\OverlayNotification') AlreadyNotified DWord 1
    Set-Reg HKCU ($cn + '\VirtualSuperResolution') AlreadyNotified DWord 1
    Set-Reg HKLM 'Software\AMD\Install' AUEP DWord 1
    Set-Reg HKLM 'Software\AUEP' RSX_AUEPStatus DWord 2
    Set-Reg HKCU 'Software\ATI\ACE\Settings\ADL\AppProfiles' AplReloadCounter DWord 0

    # ReLive / DVR
    $dv = 'Software\AMD\DVR'
    Set-Reg HKCU $dv PerformanceMonitorOpacityWA DWord 0
    Set-Reg HKCU $dv DvrEnabled DWord 0
    Set-Reg HKCU $dv ActiveSceneId String 0
    Set-Reg HKCU $dv PrevInstantReplayEnable DWord 0
    Set-Reg HKCU $dv PrevInGameReplayEnabled DWord 0
    Set-Reg HKCU $dv PrevInstantGifEnabled DWord 0
    Set-Reg HKCU $dv RemoteServerStatus DWord 0
    Set-Reg HKCU $dv ShowRSOverlay String false

    # driver defaults: UMD (3D), DXVA (video). Binary values are given as hex bytes.
    $umd = [ordered]@{
      AnisoDegree_SET = '3020322034203820313600'; Main3D_SET = '302031203220332034203500'; Tessellation_OPTION = '3200'; Tessellation = '3100'
      AAF = '30000000'; GI = '31000000'; CatalystAI = '31000000'; TemporalAAMultiplier_NA = '3100'; ForceZBufferDepth = '30000000'
      EnableTripleBuffering = '3000'; ExportCompressedTex = '31000000'; PixelCenter = '30000000'; ZFormats_NA = '3100'; DitherAlpha_NA = '3100'
      SwapEffect_D3D_SET = '3020312032203320342038203900'; TFQ = '3200'; VSyncControl = '3100'; TextureOpt = '30000000'; TextureLod = '30000000'
      ASE = '3000'; ASD = '3000'; ASTT = '3000'; AntiAliasSamples = '3000'; AntiAlias = '3100'; AnisoDegree = '3000'; AnisoType = '30000000'
      AntiAliasMapping_SET = '3028303A302C313A3029203228303A322C313A3229203428303A342C313A3429203828303A382C313A382900'
      AntiAliasSamples_SET = '3020322034203800'; ForceZBufferDepth_SET = '3020313620323400'
      SwapEffect_OGL_SET = '3020312032203320342035203620372038203920313120313220313320313420313520313620313700'
      Tessellation_SET = '31203220342036203820313620333220363400'; HighQualityAF = '3100'; DisplayCrossfireLogo = '3000'
      AppGpuId = '300078003000310030003000'; SwapEffect = '30000000'; PowerState = '3000'; AntiStuttering = '3100'; TurboSync = '3000'
      SurfaceFormatReplacements = '3100'; EQAA = '3000'; ShaderCache = '3100'; MLF = '3000'; TruformMode_NA = '3100'; Main3D = '3100'
    }
    $dxva = [ordered]@{
      LRTCEnable = '30000000'; '3to2Pulldown' = '31000000'; MosquitoNoiseRemoval_ENABLE = '30000000'; Deblocking_ENABLE = '30000000'; DemoMode = '30000000'
      OverridePA = '30000000'; DynamicRange = '30000000'; StaticGamma_ENABLE = '30000000'; BlueStretch_ENABLE = '31000000'; BlueStretch = '31000000'
      LRTCCoef = '3100300030000000'; DynamicContrast_ENABLE = '30000000'; WhiteBalanceCorrection = '30000000'; Fleshtone_ENABLE = '30000000'
      ColorVibrance_ENABLE = '31000000'; ColorVibrance = '340030000000'; Detail_ENABLE = '30000000'; Detail = '310030000000'; Denoise_ENABLE = '30000000'
      TrueWhite = '30000000'; OvlTheaterMode = '30000000'; StaticGamma = '3100300030000000'; InternetVideo = '30000000'
    }
    foreach ($key in (Get-ClassKeys $script:GpuClass 'ven_1002')) {
      # kernel mode driver: Radeon Boost, Chill, Anti-Lag, USU off, DMA copy and power gating
      Set-Reg HKLM $key NotifySubscription Binary '3000'
      Set-Reg HKLM $key IsComponentControl Binary '00000000'
      Set-Reg HKLM $key KMD_USUEnable DWord 0
      Set-Reg HKLM $key KMD_RadeonBoostEnabled DWord 0
      Set-Reg HKLM $key IsAutoDefault Binary '01000000'
      Set-Reg HKLM $key KMD_ChillEnabled DWord 0
      Set-Reg HKLM $key KMD_DeLagEnabled DWord 0
      Set-Reg HKLM $key ACE Binary '3000'
      Set-Reg HKLM $key DisableDMACopy DWord 1
      Set-Reg HKLM $key DisableBlockWrite DWord 0
      Set-Reg HKLM $key DisableDrmdmaPowerGating DWord 1
      foreach ($name in $umd.Keys) { Set-Reg HKLM ($key + '\UMD') $name Binary $umd[$name] }
      Set-Reg HKLM ($key + '\UMD') Main3D_DEF String 1
      foreach ($name in $dxva.Keys) { Set-Reg HKLM ($key + '\UMD\DXVA') $name Binary $dxva[$name] }
    }

    # services
    Set-Reg HKLM 'System\CurrentControlSet\Services\amdwddmg' ChillEnabled DWord 0
    Disable-Svc 'AMD Crash Defender Service', 'AMD External Events Utility', 'amdfendr', 'amdfendrmgr', 'amdlog'
    Write-Note 'AMD tweaks applied. Restart required.'
  }
}

Add-Tweak @{
  Id = 'gpu_auto'; Title = 'Auto (detected GPU)'; Page = 'gpu'; Icon = 'monitor'; Risk = 'moderate'; Reboot = $true
  Short = 'Applies the right tweaks for the GPU vendor found in this PC.'
  Tags = @('NVIDIA', 'AMD')
  Does = @('NVIDIA adapters get the NVIDIA tweaks', 'AMD adapters get the AMD tweaks', 'Other GPUs are skipped')
  Covers = @('nvidia_all', 'amd_all')
  Apply = {
    $nv = Get-ClassKeys $script:GpuClass 'ven_10de'
    $amd = Get-ClassKeys $script:GpuClass 'ven_1002'
    if ($nv.Count -gt 0) { Use-Tweak nvidia_all }
    if ($amd.Count -gt 0) { Use-Tweak amd_all }
    if ($nv.Count -eq 0 -and $amd.Count -eq 0) { Write-Note 'No NVIDIA or AMD adapter found, nothing to do.' }
  }
}

# ============================================================================
#  TWEAKS - expert lab
#  These trade safety or features for speed. The UI asks before running them.
# ============================================================================

Add-Tweak @{
  Id = 'mitigations_off'; Title = 'CPU and OS mitigations off'; Page = 'expert'; Icon = 'shield'; Risk = 'security'; Reboot = $true
  Short = 'Spectre and Meltdown, SEHOP, CFG and VBS/HVCI switched off.'
  Tags = @('Spectre', 'CFG', 'VBS')
  Does = @('Spectre and Meltdown mitigations off', 'SEHOP, CFG and process mitigations off', 'Memory integrity (VBS / HVCI) off')
  Note = 'Lowers system security. Reboot required.'
  Apply = {
    $mm = 'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
    $kr = 'SYSTEM\CurrentControlSet\Control\Session Manager\kernel'
    Set-Reg HKLM $mm FeatureSettings DWord 1
    Set-Reg HKLM $mm FeatureSettingsOverride DWord 3
    Set-Reg HKLM $mm FeatureSettingsOverrideMask DWord 3
    Set-Reg HKLM $mm EnableCfg DWord 0
    Set-Reg HKLM $kr DisableExceptionChainValidation DWord 1
    Set-Reg HKLM $kr KernelSEHOPEnabled DWord 0
    Set-Reg HKLM $kr MitigationOptions Binary '222222222222222222222222222222222222222222222222'
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' EnableVirtualizationBasedSecurity DWord 0
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' HVCIMATRequired DWord 0
    Set-Reg HKLM 'SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' Enabled DWord 0
    Write-Note 'Mitigations off. Restart required.'
  }
}

Add-Tweak @{
  Id = 'defender_off'; Title = 'Disable Windows Defender'; Page = 'expert'; Icon = 'shield'; Risk = 'security'; Reboot = $true
  Short = 'Real-time protection, SmartScreen and Defender services off.'
  Tags = @('Defender', 'SmartScreen')
  Does = @(
    'Turn Tamper Protection off in Windows Security first, otherwise Windows blocks this'
    'Policies and Set-MpPreference: real-time, behavior and cloud protection off'
    'Defender and Security Center services disabled'
  )
  Note = 'Your PC has no antivirus afterward unless you install one.'
  Apply = {
    $df = 'SOFTWARE\Policies\Microsoft\Windows Defender'
    Set-Reg HKLM $df DisableAntiSpyware DWord 1
    Set-Reg HKLM $df DisableAntiVirus DWord 1
    Set-Reg HKLM $df DisableRoutinelyTakingAction DWord 1
    Set-Reg HKLM $df ServiceKeepAlive DWord 0
    Set-Reg HKLM ($df + '\Real-Time Protection') DisableBehaviorMonitoring DWord 1
    Set-Reg HKLM ($df + '\Real-Time Protection') DisableIOAVProtection DWord 1
    Set-Reg HKLM ($df + '\Real-Time Protection') DisableOnAccessProtection DWord 1
    Set-Reg HKLM ($df + '\Real-Time Protection') DisableRealtimeMonitoring DWord 1
    Set-Reg HKLM ($df + '\Real-Time Protection') DisableScriptScanning DWord 1
    Set-Reg HKLM ($df + '\Reporting') DisableEnhancedNotifications DWord 1
    Set-Reg HKLM ($df + '\Spynet') SpynetReporting DWord 0
    Set-Reg HKLM ($df + '\Spynet') SubmitSamplesConsent DWord 2
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows Defender Security Center\Notifications' DisableNotifications DWord 1
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\MRT' DontReportInfectionInformation DWord 1
    Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\System' EnableSmartScreen DWord 0
    Set-MpPref @{ DisableRealtimeMonitoring = $true; DisableBehaviorMonitoring = $true; DisableIOAVProtection = $true; DisableScriptScanning = $true; DisableBlockAtFirstSeen = $true; MAPSReporting = 0; SubmitSamplesConsent = 2 }
    Disable-Svc 'Sense', 'WdNisSvc', 'WinDefend', 'SecurityHealthService', 'wscsvc'
    Write-Note 'Defender off. Restart required.'
  }
}

Add-Tweak @{
  Id = 'ntfs_tweaks'; Title = 'NTFS tweaks'; Page = 'expert'; Icon = 'file'; Risk = 'moderate'; Reboot = $false
  Short = 'Memory usage, MFT zone, last-access off, TRIM on.'
  Tags = @('fsutil')
  Does = @('memoryusage 2, mftzone 4', 'Last-access updates off, delete notify (TRIM) on', 'Paging file encryption off')
  Apply = {
    Set-Fsutil memoryusage 2
    Set-Fsutil mftzone 4
    Set-Fsutil disablelastaccess 1
    Set-Fsutil disabledeletenotify 0
    Set-Fsutil encryptpagingfile 0
  }
}

Add-Tweak @{
  Id = 'memcomp_off'; Title = 'Memory compression off'; Page = 'expert'; Icon = 'cpu'; Risk = 'moderate'; Reboot = $true
  Short = 'Disable memory compression and page combining.'
  Tags = @('MMAgent')
  Does = @('Disable-MMAgent -MemoryCompression -PageCombining', 'Uses more RAM. Best with 16 GB or more')
  Apply = {
    Disable-MemoryCompression
  }
}

Add-Tweak @{
  Id = 'power_delete_others'; Title = 'Delete other power plans'; Page = 'expert'; Icon = 'trash'; Risk = 'moderate'; Reboot = $false
  Short = 'Keep only the JiffOS plan. Originals are exported first.'
  Tags = @('powercfg')
  Does = @('Requires the JiffOS plan to be active', 'Every other plan is exported, then deleted', 'Restore brings the plans back')
  Check = {
    $g = Get-PlanGuids
    if ($g -contains $script:Cfg.PlanGuid -and @($g | Where-Object { $_ -ne $script:Cfg.PlanGuid }).Count -eq 0) { return 'on' }
    return 'off'
  }
  Apply = {
    Invoke-Step 'Save your current power plans (Restore brings them back)' { Save-PowerPlans }
    Invoke-Step 'Delete every power plan except JiffOS' {
      if ((Get-ActivePlanGuid) -ne $script:Cfg.PlanGuid) { throw 'the JiffOS plan is not active - apply the JiffOS power plan first' }
      foreach ($g in (Get-PlanGuids)) {
        if ($g -ne $script:Cfg.PlanGuid) { [void](Exe 'powercfg' @('/delete', $g)); Log 'INFO' ('deleted power plan ' + $g) }
      }
    }
  }
}

Add-Tweak @{
  Id = 'branding'; Title = 'JiffOS branding'; Page = 'expert'; Icon = 'sparkle'; Risk = 'safe'; Reboot = $false
  Short = 'Wallpaper, lock screen, accent, system info and PC name.'
  Tags = @('Wallpaper', 'Theme')
  Does = @(
    'Downloads the JiffOS wallpaper (the only thing here that uses the internet), sets lock screen and dark accent theme'
    'Renames This PC, boot entry and the C: label, sets the PC name to JIFFOS-PC (applies after a restart)'
    'Adds system info logo and a desktop shortcut'
  )
  Note = 'Restore puts the registry values, boot names, drive label and PC name back and removes the files it created.'
  Apply = {
    $name = 'JiffOS'
    $pcName = 'JIFFOS-PC'
    $wallUrl = 'https://raw.githubusercontent.com/giannhs1101/wall/main/jiff_wallpaper.jpg'
    $discord = 'https://discord.gg/fuegoscrims'
    $dir = Join-Path $env:ProgramData 'JiffOS'
    $wall = Join-Path $dir 'wallpaper.jpg'
    $logo = Join-Path $dir 'logo.bmp'
    $cv = 'Software\Microsoft\Windows\CurrentVersion'

    Invoke-Step 'Download the wallpaper and set it as desktop and lock screen' {
      [void](New-Item -ItemType Directory -Force -Path $dir)
      if ($script:Mock) { [IO.File]::WriteAllText($wall, ('x' * 20000)) }
      else {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -UseBasicParsing -Uri $wallUrl -OutFile $wall
      }
      if ((Get-Item -LiteralPath $wall).Length -lt 10240) { Remove-Item -LiteralPath $wall -Force; throw 'wallpaper download failed' }
      Register-CreatedFile $wall
      Set-Reg HKCU 'Control Panel\Desktop' Wallpaper String $wall
      Set-Reg HKCU 'Control Panel\Desktop' WallpaperStyle String 10
      Set-Reg HKCU 'Control Panel\Desktop' TileWallpaper String 0
      Set-Wallpaper $wall
      $lk = 'SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP'
      Set-Reg HKLM $lk LockScreenImagePath String $wall
      Set-Reg HKLM $lk LockScreenImageUrl String $wall
      Set-Reg HKLM $lk LockScreenImageStatus DWord 1
      Set-Reg HKCU ($cv + '\ContentDeliveryManager') RotatingLockScreenEnabled DWord 0
      Set-Reg HKCU ($cv + '\ContentDeliveryManager') RotatingLockScreenOverlayEnabled DWord 0
      Set-Reg HKLM 'SOFTWARE\Policies\Microsoft\Windows\System' DisableAcrylicBackgroundOnLogon DWord 1
    }

    # dark theme with the green accent
    Set-Reg HKCU ($cv + '\Themes\Personalize') AppsUseLightTheme DWord 0
    Set-Reg HKCU ($cv + '\Themes\Personalize') SystemUsesLightTheme DWord 0
    Set-Reg HKCU ($cv + '\Themes\Personalize') ColorPrevalence DWord 0
    Set-Reg HKCU 'Software\Microsoft\Windows\DWM' AccentColor DWord 0xff53c800
    Set-Reg HKCU 'Software\Microsoft\Windows\DWM' ColorizationColor DWord 0xc400c853
    Set-Reg HKCU 'Software\Microsoft\Windows\DWM' ColorizationAfterglow DWord 0xc400c853
    Set-Reg HKCU 'Software\Microsoft\Windows\DWM' ColorPrevalence DWord 0
    Set-Reg HKCU ($cv + '\Explorer\Accent') AccentColorMenu DWord 0xff53c800
    Set-Reg HKCU ($cv + '\Explorer\Accent') StartColorMenu DWord 0xff53c800

    # names: This PC, boot menu, drive, computer
    Set-Reg HKCU ($cv + '\Explorer\CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}') '' String $name
    Set-Bcd '{current}' description $name
    Set-Bcd '{bootmgr}' description ($name + ' Boot Manager')
    Invoke-Step ('Name the C: drive ' + $name) {
      Journal-Add 'label' 'C' @{ drive = 'C'; label = (Get-DriveLabel 'C') }
      Set-DriveLabel 'C' $name
    }
    Invoke-Step ('Rename the PC to ' + $pcName + ' (applies after a restart)') {
      $cur = Get-HostName
      if ($cur -ine $pcName) { Journal-Add 'host' 'name' @{ name = $cur }; Set-HostName $pcName }
    }
    Set-Reg HKLM 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' RegisteredOwner String ($name + ' User')
    Set-Reg HKLM 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' RegisteredOrganization String $name

    # system info
    Invoke-Step 'Create the system info logo' {
      [void](New-Item -ItemType Directory -Force -Path $dir)
      New-JiffLogo $logo
      Register-CreatedFile $logo
    }
    $oem = 'SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation'
    Set-Reg HKLM $oem Manufacturer String $name
    Set-Reg HKLM $oem Model String $name
    Set-Reg HKLM $oem SupportURL String $discord
    if ($script:Mode -ne 'apply' -or (Test-Path -LiteralPath $logo)) { Set-Reg HKLM $oem Logo String $logo }

    Invoke-Step 'Add the JiffOS Community shortcut to your desktop' {
      $desk = [Environment]::GetFolderPath('Desktop')
      if (-not $desk) { return }
      $f = Join-Path $desk ($name + ' Community.url')
      [IO.File]::WriteAllText($f, "[InternetShortcut]`r`nURL=" + $discord + "`r`n")
      Register-CreatedFile $f
    }
    Invoke-Step 'Restart Explorer so the new names show up' {
      if ($script:Mock) { return }
      Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
      Start-Process explorer.exe
    }
  }
}

# ============================================================================
#  TWEAKS - restore and the Full profile
#  Restore reads the journal (journal.jsonl in the data folder). Every change
#  above recorded the original value there the first time it touched it.
# ============================================================================

Add-Tweak @{
  Id = 'restore_all'; Title = 'Restore everything'; Page = 'restore'; Icon = 'restore'; Risk = 'safe'; Reboot = $true; Kind = 'restore'
  Short = 'Every saved original goes back: services, tasks, registry, network, power plans, boot settings and more.'
  Tags = @('Services', 'Registry', 'Tasks', 'Power', 'Network')
  Does = @(
    'Services return to their original start types'
    'Registry values return to what they were, and values the tweaks created are deleted'
    'Scheduled tasks, power plans, boot settings, network settings and device flags are restored'
    'Removed apps are not reinstalled'
  )
  Check = { $null }
  Apply = { Invoke-RestoreKinds @('reg', 'svc', 'task', 'bcd', 'pnp', 'net', 'mm', 'cmd', 'mp', 'label', 'host', 'file', 'power') 'saved settings' }
}

Add-Tweak @{
  Id = 'restore_services'; Title = 'Restore services'; Page = 'restore'; Icon = 'restore'; Risk = 'safe'; Reboot = $true; Kind = 'restore'
  Short = 'Original start types for every service that was changed.'
  Tags = @('Services')
  Does = @('Writes the saved start type back for each service')
  Check = { $null }
  Apply = { Invoke-RestoreKinds @('svc') 'services' }
}

Add-Tweak @{
  Id = 'restore_tasks'; Title = 'Restore scheduled tasks'; Page = 'restore'; Icon = 'restore'; Risk = 'safe'; Reboot = $false; Kind = 'restore'
  Short = 'Re-enable every task the tweaks disabled.'
  Tags = @('Tasks')
  Does = @('Enables each task recorded by the cleanup')
  Check = { $null }
  Apply = { Invoke-RestoreKinds @('task') 'scheduled tasks' }
}

Add-Tweak @{
  Id = 'restore_registry'; Title = 'Restore registry'; Page = 'restore'; Icon = 'restore'; Risk = 'safe'; Reboot = $true; Kind = 'restore'
  Short = 'Put back registry values and delete the ones that were added.'
  Tags = @('Registry')
  Does = @('Each saved value is written back', 'Values that did not exist before are deleted')
  Check = { $null }
  Apply = { Invoke-RestoreKinds @('reg') 'registry values' }
}

Add-Tweak @{
  Id = 'restore_power'; Title = 'Restore power plans'; Page = 'restore'; Icon = 'restore'; Risk = 'safe'; Reboot = $false; Kind = 'restore'
  Short = 'Bring back your original power plans and active plan.'
  Tags = @('Power')
  Does = @('Exported plans are re-imported', 'Your original active plan is selected, JiffOS plan removed')
  Check = { $null }
  Apply = { Invoke-RestoreKinds @('power') 'power plan sets' }
}

Add-Tweak @{
  Id = 'forget_originals'; Title = 'Forget saved originals'; Page = 'restore'; Icon = 'trash'; Risk = 'danger'; Reboot = $false; Kind = 'restore'
  Short = 'Deletes the journal. Restore can no longer go back to the original state.'
  Tags = @('Journal')
  Does = @('Deletes the saved originals and the exported power plans', 'The next change saves a new baseline from the current state')
  Note = 'This cannot be undone.'
  Check = { $null }
  Apply = {
    Invoke-Step 'Delete the restore journal and the exported power plans' {
      if (Test-Path -LiteralPath $script:Cfg.Journal) { Remove-Item -LiteralPath $script:Cfg.Journal -Force }
      $script:Seen = $null
      if (Test-Path -LiteralPath $script:Cfg.PowerDir) { Get-ChildItem -LiteralPath $script:Cfg.PowerDir -Filter 'plan_*.pow' | Remove-Item -Force }
    }
  }
}

Add-Tweak @{
  Id = 'apply_full'; Title = 'Full profile'; Page = 'profile'; Icon = 'layers'; Risk = 'aggressive'; Reboot = $true; Steps = 11
  Short = 'Everything recommended in one run, with a restore point first.'
  Does = @(
    'Creates a Windows restore point'
    'Registry: latency, gaming, input and priority tweaks'
    'Services: advanced list plus per-user service templates'
    'Scheduled task cleanup'
    'Network: TCP settings, Nagle off, adapter properties'
    'JiffOS power plan, USB and device power saving off'
    'Privacy and telemetry policies'
    'GPU tweaks for the detected vendor, boot timer tweak'
    'Debloat: Store apps and OneDrive uninstall'
  )
  Note = 'Defender and CPU mitigations are not touched. Removed apps are not restored by Restore.'
  Covers = @('reg_all', 'reg_system', 'reg_gaming', 'reg_input_ui', 'reg_priority', 'svc_advanced', 'svc_standard', 'svc_basic', 'svc_xbox', 'tasks_disable', 'net_all', 'net_global',
    'net_interfaces', 'nic_tweaks', 'power_plan', 'devpower', 'bcd_timer', 'privacy', 'debloat_curated', 'debloat_extended', 'debloat_onedrive', 'gpu_auto', 'nvidia_all', 'amd_all', 'restore_point')
  Check = {
    $ids = 'reg_all', 'svc_advanced', 'tasks_disable', 'net_all', 'power_plan', 'devpower', 'privacy', 'gpu_auto', 'bcd_timer', 'debloat_extended', 'debloat_onedrive'
    Join-State ($ids | ForEach-Object { Get-TweakState $script:Tweaks[$_] })
  }
  Apply = {
    Show-Step 1 11 'Restore point';             Use-Tweak restore_point
    Show-Step 2 11 'Registry';                  Use-Tweak reg_all
    Show-Step 3 11 'Services';                  Use-Tweak svc_advanced
    Show-Step 4 11 'Scheduled tasks';           Use-Tweak tasks_disable
    Show-Step 5 11 'Network';                   Use-Tweak net_all
    Show-Step 6 11 'Power plan';                Use-Tweak power_plan
    Show-Step 7 11 'USB and device power';      Use-Tweak devpower
    Show-Step 8 11 'Privacy';                   Use-Tweak privacy
    Show-Step 9 11 'GPU';                       Use-Tweak gpu_auto
    Show-Step 10 11 'Boot timer';               Use-Tweak bcd_timer
    Show-Step 11 11 'Debloat';                  Use-Tweak debloat_extended; Use-Tweak debloat_onedrive
  }
}

# ============================================================================
#  ENGINE - running tweaks: profiles, sequences, states, previews, code view
# ============================================================================

# ---------------------------------------------------------------- pages and profiles
$script:Pages = [ordered]@{
  optimize = @{ Title = 'Performance'; Icon = 'zap'; Sub = 'Registry, scheduler, power plan and boot timer.'; Featured = 'reg_all'
    Ids = @('reg_all', 'reg_system', 'reg_gaming', 'reg_input_ui', 'reg_priority', 'power_plan', 'devpower', 'bcd_timer') }
  services = @{ Title = 'Services'; Icon = 'sliders'; Sub = 'Turn off background services and scheduled tasks you do not need.'; Featured = 'svc_advanced'
    Ids = @('svc_advanced', 'svc_standard', 'svc_basic', 'svc_xbox', 'tasks_disable') }
  network  = @{ Title = 'Network'; Icon = 'wifi'; Sub = 'Lower latency: TCP settings, Nagle, adapter properties.'; Featured = 'net_all'
    Ids = @('net_all', 'net_global', 'net_interfaces', 'nic_tweaks', 'net_reset') }
  cleanup  = @{ Title = 'Debloat and Privacy'; Icon = 'trash'; Sub = 'Remove preinstalled apps and switch off telemetry.'; Featured = 'privacy'
    Ids = @('privacy', 'debloat_curated', 'debloat_extended', 'debloat_onedrive') }
  gpu      = @{ Title = 'GPU'; Icon = 'monitor'; Sub = 'Driver-level tweaks for NVIDIA and AMD.'; Featured = 'gpu_auto'
    Ids = @('gpu_auto', 'nvidia_all', 'amd_all') }
  expert   = @{ Title = 'Expert Lab'; Icon = 'flask'; Sub = 'Powerful changes that trade safety or features for speed.'; Featured = ''
    Ids = @('mitigations_off', 'defender_off', 'ntfs_tweaks', 'memcomp_off', 'power_delete_others', 'branding') }
  restore  = @{ Title = 'Restore'; Icon = 'restore'; Sub = 'Put saved original values back.'; Featured = 'restore_all'
    Ids = @('restore_all', 'restore_point', 'restore_services', 'restore_tasks', 'restore_registry', 'restore_power', 'forget_originals') }
}

$script:Profiles = [ordered]@{
  safe = @{ Title = 'Safe'; Icon = 'shield'; Line = 'Low risk. Nothing that removes features or breaks apps.'
    Ids = @('restore_point', 'reg_gaming', 'reg_input_ui', 'privacy', 'tasks_disable', 'net_global', 'power_plan')
    Tags = @('Gaming', 'Input', 'Privacy', 'Tasks', 'TCP', 'Power plan') }
  competitive = @{ Title = 'Competitive'; Icon = 'zap'; Line = 'Low latency for games. Leaves Update, printing, VPN and Store alone.'
    Ids = @('restore_point', 'reg_all', 'net_all', 'power_plan', 'devpower', 'bcd_timer', 'privacy', 'tasks_disable', 'svc_basic')
    Tags = @('Latency', 'Network', 'Power', 'USB', 'Boot timer', 'Basic services') }
  full = @{ Title = 'Full'; Icon = 'layers'; Line = 'Everything: advanced services, GPU tweaks and debloat.'
    Ids = @('apply_full')
    Tags = @('Registry', 'Services', 'Network', 'Power', 'Privacy', 'GPU', 'Debloat') }
}

# the order tweaks run in when several are chosen together
$script:Order = @('restore_point', 'reg_all', 'reg_system', 'reg_gaming', 'reg_input_ui', 'reg_priority', 'svc_basic', 'svc_standard', 'svc_advanced', 'svc_xbox',
  'tasks_disable', 'net_all', 'net_global', 'net_interfaces', 'nic_tweaks', 'net_reset', 'power_plan', 'devpower', 'bcd_timer', 'privacy', 'debloat_curated',
  'debloat_extended', 'debloat_onedrive', 'gpu_auto', 'nvidia_all', 'amd_all', 'ntfs_tweaks', 'memcomp_off', 'power_delete_others', 'mitigations_off',
  'defender_off', 'branding')

# one line per verb, shown next to the code
$script:VerbHelp = [ordered]@{
  'Set-Reg'              = 'Writes a registry value. The original is saved first, or noted as missing so Restore can delete it.'
  'Remove-Reg'           = 'Deletes a registry value. The original is saved first.'
  'Disable-Svc'          = 'Sets services to Disabled and stops them. The original start type is saved first.'
  'Disable-Task'         = 'Disables scheduled tasks. Only tasks that are enabled are touched and each is recorded.'
  'Invoke-Tool'          = 'Runs a Windows command line tool. -Undo is the command Restore runs to put things back.'
  'Invoke-Step'          = 'One described action. Preview lists the description, Apply runs the code.'
  'Set-NetOpt'           = 'Reads a Windows network setting through the network cmdlets, saves it, then writes it with netsh.'
  'Set-Bcd'              = 'Sets (or deletes) a boot entry value with bcdedit. The original is saved first.'
  'Set-Fsutil'           = 'Runs fsutil behavior set and saves the old value for Restore.'
  'Remove-Apps'          = 'Uninstalls Store apps for all users and from the Windows image. Not restorable.'
  'Use-Tweak'            = 'Runs another tweak (used to build bigger tweaks out of smaller ones).'
  'Show-Step'            = 'Progress line for the Full profile.'
  'Write-Note'           = 'A line for the summary at the end of a run.'
}

function Resolve-Sequence {
  param([string[]]$Ids)
  $set = New-Object System.Collections.Generic.List[string]
  foreach ($i in $Ids) { if ($script:Tweaks.Contains($i) -and -not $set.Contains($i)) { $set.Add($i) } }
  $drop = @{}
  foreach ($i in $set) { foreach ($c in @($script:Tweaks[$i].Covers)) { if ($c) { $drop[$c] = $true } } }
  $keep = @($set | Where-Object { -not $drop.ContainsKey($_) })
  $ordered = @($script:Order | Where-Object { $keep -contains $_ })
  $rest = @($keep | Where-Object { $script:Order -notcontains $_ })
  return @($ordered + $rest)
}

# ---------------------------------------------------------------- control channel shared with the UI server
function Get-Ctl {
  $c = [AppDomain]::CurrentDomain.GetData('JTCtl')
  if (-not $c) { $c = [hashtable]::Synchronized(@{ Cancel = $false }); [AppDomain]::CurrentDomain.SetData('JTCtl', $c) }
  return $c
}

function Test-Cancel { return [bool](Get-Ctl).Cancel }

# ---------------------------------------------------------------- run a list of tweaks
# Log lines the UI reads:  [SEQ] start/end   [TWEAK] id|title   [STEP] n/total title   [OK] [FAIL] [SKIP] [NOTE] [INFO]   [DONE] id ok= fail= skip=
function Invoke-JiffSequence {
  param([string[]]$Ids, [string]$LogFile, [switch]$Echo)
  Initialize-Jiff
  $script:LogFile = $LogFile; $script:Echo = [bool]$Echo
  (Get-Ctl).Cancel = $false
  $ids = Resolve-Sequence $Ids
  $ok = 0; $fail = 0; $skip = 0; $reboot = $false; $done = 0
  Log 'SEQ' ('start ' + $ids.Count + ' ' + ($ids -join ','))
  foreach ($id in $ids) {
    if (Test-Cancel) { Log 'SEQ' 'cancelled'; break }
    $t = $script:Tweaks[$id]
    Log 'TWEAK' ($id + '|' + $t.Title)
    $script:Mode = 'apply'; $script:CntOk = 0; $script:CntFail = 0; $script:CntSkip = 0
    try { & $t.Apply }
    catch {
      if ($_.Exception.Message -eq 'cancelled') { Log 'SEQ' 'cancelled'; $script:Mode = 'apply'; break }
      Log 'FAIL' ($id + ': ' + $_.Exception.Message)
    }
    Log 'DONE' ($id + ' ok=' + $script:CntOk + ' fail=' + $script:CntFail + ' skip=' + $script:CntSkip)
    $ok += $script:CntOk; $fail += $script:CntFail; $skip += $script:CntSkip; $done++
    if ($t.Reboot -and $script:CntOk -gt 0) { $reboot = $true }
  }
  Log 'SEQ' ('end done=' + $done + ' ok=' + $ok + ' fail=' + $fail + ' skip=' + $skip + ' reboot=' + [int]$reboot)
  $script:LogFile = $null
  return @{ Done = $done; Ok = $ok; Fail = $fail; Skip = $skip; Reboot = $reboot }
}

# ---------------------------------------------------------------- states
function Get-StateMap {
  param([string[]]$Ids)
  if (-not $Ids -or $Ids.Count -eq 0) { $Ids = @($script:Tweaks.Keys) }
  $out = [ordered]@{}
  foreach ($id in $Ids) {
    $t = $script:Tweaks[$id]
    if ($t) { $out[$id] = Get-TweakState $t }
  }
  return $out
}

# ---------------------------------------------------------------- preview (what would change)
function Get-PreviewData {
  param([string[]]$Ids)
  $ids = Resolve-Sequence $Ids
  $out = @()
  foreach ($id in $ids) {
    $r = Get-TweakResult $script:Tweaks[$id] 'preview'
    $changes = @($r.Plan)
    $out += @{ id = $id; title = $script:Tweaks[$id].Title; count = $changes.Count; changes = $changes }
  }
  return $out
}

# ---------------------------------------------------------------- code view
function Format-Source {
  param([string]$Text)
  $lines = $Text -split "`r?`n"
  while ($lines.Count -gt 0 -and -not $lines[0].Trim()) { $lines = $lines | Select-Object -Skip 1 }
  while ($lines.Count -gt 0 -and -not $lines[$lines.Count - 1].Trim()) { $lines = $lines | Select-Object -First ($lines.Count - 1) }
  $min = 1000
  foreach ($l in $lines) { if ($l.Trim()) { $ind = $l.Length - $l.TrimStart().Length; if ($ind -lt $min) { $min = $ind } } }
  if ($min -eq 1000) { $min = 0 }
  return (($lines | ForEach-Object { if ($_.Length -ge $min) { $_.Substring($min) } else { $_.TrimStart() } }) -join "`n")
}

function Get-TweakCode {
  param([string]$Id)
  $t = $script:Tweaks[$Id]
  if (-not $t) { return $null }
  $apply = Format-Source $t.Apply.ToString()
  $check = $null
  if ($t.Check) { $check = Format-Source $t.Check.ToString() }
  $verbs = @()
  $all = $apply + "`n" + $check
  foreach ($fn in @(Get-ChildItem Function:) ) {
    $n = $fn.Name
    if ($n -match '^(Set-Reg|Remove-Reg|Disable-Svc|Disable-Task|Invoke-Tool|Invoke-Step|Set-NetOpt|Set-Bcd|Set-Fsutil|Remove-Apps|Use-Tweak|Show-Step|Write-Note|Disable-PnpPower|Disable-MemoryCompression|Set-MpPref|Find-EnumValues|Get-ClassKeys|Find-Tasks|Save-PowerPlans|New-JiffPlan|Set-PlanTuning|Invoke-RestoreKinds|Restore-Journal)$') {
      if ($all -match ('(?m)(^|[^\w-])' + [regex]::Escape($n) + '([^\w-]|$)')) {
        $help = ''; if ($script:VerbHelp.Contains($n)) { $help = $script:VerbHelp[$n] }
        $verbs += @{ name = $n; help = $help; source = ('function ' + $n + " {`n" + (Format-Source $fn.ScriptBlock.ToString()) + "`n}") }
      }
    }
  }
  return @{ id = $Id; title = $t.Title; apply = $apply; check = $check; verbs = $verbs }
}

# ---------------------------------------------------------------- what has been changed (the journal)
function Get-ChangeList {
  $rows = @()
  foreach ($r in (Journal-Load)) {
    $d = Describe-JournalRow $r
    $d['time'] = [string]$r.t
    $d['k'] = [string]$r.k
    $rows += $d
  }
  return $rows
}

# ---------------------------------------------------------------- system info
function Test-IsAdmin {
  if ($script:Mock) { return $true }
  try { return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch { return $false }
}

function Get-SystemInfo {
  $cv = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
  $prod = RegRead HKLM $cv ProductName; $dv = RegRead HKLM $cv DisplayVersion; $bn = RegRead HKLM $cv CurrentBuildNumber; $ubr = RegRead HKLM $cv UBR
  $os = 'Windows'; if ($prod) { $os = [string]$prod.Value }
  $build = 0; if ($bn) { $build = [int]$bn.Value }
  if ($build -ge 22000) { $os = $os -replace 'Windows 10', 'Windows 11' }
  $cpu = RegRead HKLM 'HARDWARE\DESCRIPTION\System\CentralProcessor\0' ProcessorNameString
  $cpuName = 'Unknown CPU'; if ($cpu) { $cpuName = ([string]$cpu.Value).Trim() }
  $ramMb = 0; $gpus = @()
  if ($script:Mock) { $ramMb = 32768; $cpuName = 'Mock CPU 8-core'; $gpus = @('NVIDIA GeForce RTX 3060', 'AMD Radeon RX 6700') }
  else {
    try { $ramMb = [int]([math]::Round((Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory / 1MB)) } catch { }
    try { $gpus = @(Get-CimInstance Win32_VideoController -ErrorAction Stop | ForEach-Object { $_.Name }) } catch { }
  }
  $nv = (Get-ClassKeys $script:GpuClass 'ven_10de').Count -gt 0
  $amd = (Get-ClassKeys $script:GpuClass 'ven_1002').Count -gt 0
  $edition = ''; if ($dv) { $edition = [string]$dv.Value }
  return @{
    os = $os; version = $edition; build = $build; ubr = $(if ($ubr) { [int]$ubr.Value } else { 0 })
    cpu = $cpuName; threads = [Environment]::ProcessorCount; ramMb = $ramMb; gpus = $gpus
    hasNvidia = $nv; hasAmd = $amd; host = (Get-HostName); admin = (Test-IsAdmin)
    psVersion = $PSVersionTable.PSVersion.ToString(); mock = $script:Mock; tool = $script:Cfg.Version
  }
}

# ---------------------------------------------------------------- before / after numbers
function Get-Snapshot {
  if ($script:Mock) {
    $off = 0
    foreach ($k in @($script:MockState.Reg.Keys)) { if ($k -like 'hklm|system\currentcontrolset\services\*|start' -and $script:MockState.Reg[$k].Value -eq 4) { $off++ } }
    return @{ t = (Get-Date).ToString('s'); services = [math]::Max(40, 118 - $off); processes = [math]::Max(60, 174 - [int]($off / 2)); ramUsedMb = [math]::Max(3000, 6900 - $off * 14); ramTotalMb = 32768 }
  }
  $svc = 0; $proc = 0; $used = 0; $total = 0
  try { $svc = @(Get-Service | Where-Object { $_.Status -eq 'Running' }).Count } catch { }
  try { $proc = @(Get-Process).Count } catch { }
  try { $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop; $total = [int]($os.TotalVisibleMemorySize / 1024); $used = [int](($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / 1024) } catch { }
  return @{ t = (Get-Date).ToString('s'); services = $svc; processes = $proc; ramUsedMb = $used; ramTotalMb = $total }
}

# ---------------------------------------------------------------- catalog for the UI (plain data, no code)
function Get-CatalogData {
  $tw = @()
  foreach ($t in $script:Tweaks.Values) {
    $tw += @{
      id = $t.Id; title = $t.Title; page = $t.Page; icon = $t.Icon; risk = $t.Risk; reboot = [bool]$t.Reboot; short = $t.Short
      does = @($t.Does); tags = @($t.Tags); note = $t.Note; vendor = $t.Vendor; covers = @($t.Covers); kind = $t.Kind; steps = $t.Steps
      hasState = [bool]($t.Kind -eq 'tweak')
    }
  }
  $pg = @()
  foreach ($k in $script:Pages.Keys) { $p = $script:Pages[$k]; $pg += @{ id = $k; title = $p.Title; icon = $p.Icon; sub = $p.Sub; featured = $p.Featured; ids = @($p.Ids) } }
  $pf = @()
  foreach ($k in $script:Profiles.Keys) { $p = $script:Profiles[$k]; $pf += @{ id = $k; title = $p.Title; icon = $p.Icon; line = $p.Line; ids = @($p.Ids); tags = @($p.Tags) } }
  $vb = @(); foreach ($k in $script:VerbHelp.Keys) { $vb += @{ name = $k; help = $script:VerbHelp[$k] } }
  return @{ version = $script:Cfg.Version; tweaks = $tw; pages = $pg; profiles = $pf; order = @($script:Order); verbs = $vb }
}

# ============================================================================
#  ENGINE - live sampler
#  Reads performance numbers from this PC once per second for the Live page.
#  Everything stays on this machine. The only network traffic is the optional
#  ping to a host you choose (off until you switch it on in Settings).
#  WMI performance classes are used because their names do not change with
#  the Windows display language.
# ============================================================================

$script:Samp = @{ Searchers = @{}; ProcPrev = @{}; ProcAt = $null; Net = $null; Ping = $null; PingHist = New-Object System.Collections.ArrayList; Boot = $null; Rnd = New-Object System.Random; Gpu = @{ gpu = 0; mem = 0; temp = 0; memUsed = 0; memTotal = 0; power = 0 }; Top = @() }

function Wmi-Query {
  param([string]$Query)
  try {
    $s = $script:Samp.Searchers[$Query]
    if (-not $s) {
      $s = New-Object System.Management.ManagementObjectSearcher('root\cimv2', $Query)
      $s.Options.ReturnImmediately = $true
      $script:Samp.Searchers[$Query] = $s
    }
    return @($s.Get())
  } catch { return @() }
}

# Math.Max/Min with a whole number first (Max(0, 0.35)) silently rounds to whole numbers in PowerShell, so these keep doubles
function Dmax { param([double]$a, [double]$b) if ($a -gt $b) { return $a } return $b }
function Dmin { param([double]$a, [double]$b) if ($a -lt $b) { return $a } return $b }

function Get-CoreCount { return [Environment]::ProcessorCount }

# ---------------------------------------------------------------- fake numbers for tests on non-Windows machines
function Get-MockSample {
  param([double]$T)
  $n = Get-CoreCount; $r = $script:Samp.Rnd
  $cores = @(); $sum = 0
  for ($i = 0; $i -lt $n; $i++) { $v = [math]::Round((Dmax (1) ((Dmin (100) (22 + 18 * [math]::Sin($T / 3 + $i) + $r.Next(-6, 9))))), 1); $cores += $v; $sum += $v }
  $cpu = [math]::Round($sum / $n, 1)
  $ramTotal = 32768; $ramUsed = [int](9800 + 400 * [math]::Sin($T / 11) + $r.Next(-30, 30))
  $top = @(
    @{ name = 'chrome'; pid = 4120; cpu = [math]::Round((Dmax (0.2) (6 + 4 * [math]::Sin($T / 4))), 1); ramMb = 1180 }
    @{ name = 'dwm'; pid = 1388; cpu = [math]::Round((Dmax (0.1) (2.4 + [math]::Sin($T / 2))), 1); ramMb = 210 }
    @{ name = 'Discord'; pid = 7712; cpu = [math]::Round((Dmax (0.1) (1.8 + [math]::Sin($T / 5))), 1); ramMb = 640 }
    @{ name = 'System'; pid = 4; cpu = [math]::Round((Dmax (0.1) (1.1 + 0.5 * [math]::Sin($T / 3))), 1); ramMb = 24 }
    @{ name = 'explorer'; pid = 5120; cpu = [math]::Round((Dmax (0.1) (0.8 + 0.4 * [math]::Sin($T))), 1); ramMb = 190 }
    @{ name = 'svchost'; pid = 1012; cpu = [math]::Round((Dmax (0.1) (0.6 + 0.3 * [math]::Sin($T / 6))), 1); ramMb = 68 }
  )
  $gpu = [math]::Round((Dmax (0) ((Dmin (100) (38 + 30 * [math]::Sin($T / 7) + $r.Next(-4, 5))))), 1)
  return @{
    cpu = $cpu; cores = $cores; clock = [int](4200 + 150 * [math]::Sin($T / 5)); dpc = [math]::Round((Dmax (0) (0.35 + 0.25 * [math]::Sin($T / 2) + $r.NextDouble() * 0.1)), 2); intr = [math]::Round((Dmax (0) (0.2 + 0.1 * [math]::Sin($T / 3))), 2)
    ram = [math]::Round(100 * $ramUsed / $ramTotal, 1); ramUsedMb = $ramUsed; ramTotalMb = $ramTotal
    disk = [math]::Round((Dmax (0) ((Dmin (100) (6 + 5 * [math]::Sin($T / 9) + $r.Next(0, 6))))), 1); diskR = [math]::Round((Dmax (0) (12 + 8 * [math]::Sin($T / 6))), 1); diskW = [math]::Round((Dmax (0) (5 + 3 * [math]::Sin($T / 4))), 1)
    netDown = [math]::Round((Dmax (0) (18 + 14 * [math]::Sin($T / 8) + $r.Next(0, 5))), 2); netUp = [math]::Round((Dmax (0) (2.4 + 1.6 * [math]::Sin($T / 5))), 2)
    gpu = $gpu; gpuTemp = [int](52 + 8 * [math]::Sin($T / 12)); gpuMemUsedMb = 3900; gpuMemTotalMb = 12288; gpuPower = [int](60 + 0.9 * $gpu)
    procs = 168; threads = 2410; services = 94; uptime = [int](3 * 3600 + 22 * 60 + $T % 3600); top = $top
  }
}

# ---------------------------------------------------------------- real numbers
function Get-RealSample {
  param([double]$T, [bool]$Slow)
  $out = @{}
  $s = $script:Samp

  # CPU: per core, total, DPC and interrupt time (these two matter for input latency)
  $cores = New-Object 'System.Collections.Generic.SortedDictionary[int,double]'
  foreach ($o in (Wmi-Query 'SELECT Name,PercentProcessorTime,PercentDPCTime,PercentInterruptTime FROM Win32_PerfFormattedData_PerfOS_Processor')) {
    if ($o.Name -eq '_Total') { $out.cpu = [double]$o.PercentProcessorTime; $out.dpc = [double]$o.PercentDPCTime; $out.intr = [double]$o.PercentInterruptTime }
    elseif ($o.Name -match '^\d+$') { $cores[[int]$o.Name] = [double]$o.PercentProcessorTime }
  }
  $out.cores = @($cores.Values)
  if ($null -eq $out.cpu) { $out.cpu = 0; $out.dpc = 0; $out.intr = 0 }

  # clock speed
  foreach ($o in (Wmi-Query "SELECT ProcessorFrequency,PercentProcessorPerformance FROM Win32_PerfFormattedData_Counters_ProcessorInformation WHERE Name='_Total'")) {
    $out.clock = [int]([double]$o.ProcessorFrequency * [double]$o.PercentProcessorPerformance / 100)
  }

  # memory
  foreach ($o in (Wmi-Query 'SELECT TotalVisibleMemorySize,FreePhysicalMemory FROM Win32_OperatingSystem')) {
    $out.ramTotalMb = [int]($o.TotalVisibleMemorySize / 1024); $out.ramUsedMb = [int](($o.TotalVisibleMemorySize - $o.FreePhysicalMemory) / 1024)
    $out.ram = [math]::Round(100 * $out.ramUsedMb / (Dmax (1) ($out.ramTotalMb)), 1)
  }

  # disk
  foreach ($o in (Wmi-Query "SELECT PercentDiskTime,DiskReadBytesPersec,DiskWriteBytesPersec FROM Win32_PerfFormattedData_PerfDisk_PhysicalDisk WHERE Name='_Total'")) {
    $out.disk = (Dmin (100) ([double]$o.PercentDiskTime)); $out.diskR = [math]::Round([double]$o.DiskReadBytesPersec / 1MB, 1); $out.diskW = [math]::Round([double]$o.DiskWriteBytesPersec / 1MB, 1)
  }

  # network (megabit per second, all physical adapters together)
  $rx = 0.0; $tx = 0.0
  foreach ($o in (Wmi-Query 'SELECT Name,BytesReceivedPersec,BytesSentPersec FROM Win32_PerfFormattedData_Tcpip_NetworkInterface')) {
    if ($o.Name -notmatch 'isatap|Loopback|Teredo|6to4|Pseudo') { $rx += [double]$o.BytesReceivedPersec; $tx += [double]$o.BytesSentPersec }
  }
  $out.netDown = [math]::Round($rx * 8 / 1MB, 2); $out.netUp = [math]::Round($tx * 8 / 1MB, 2)

  # system counters
  foreach ($o in (Wmi-Query 'SELECT Processes,Threads FROM Win32_PerfFormattedData_PerfOS_System')) { $out.procs = [int]$o.Processes; $out.threads = [int]$o.Threads }

  if ($null -eq $s.Boot) { foreach ($o in (Wmi-Query 'SELECT LastBootUpTime FROM Win32_OperatingSystem')) { try { $s.Boot = [Management.ManagementDateTimeConverter]::ToDateTime($o.LastBootUpTime) } catch { } } }
  if ($s.Boot) { $out.uptime = [int]((Get-Date) - $s.Boot).TotalSeconds }

  # slower numbers every other second
  if ($Slow) {
    $g = 0.0
    $eng = @{}
    foreach ($o in (Wmi-Query 'SELECT Name,UtilizationPercentage FROM Win32_PerfFormattedData_GPUPerformanceCounters_GPUEngine')) {
      if ($o.Name -match 'engtype_(\w+)$') { $et = $Matches[1]; if (-not $eng.ContainsKey($et)) { $eng[$et] = 0.0 }; $eng[$et] += [double]$o.UtilizationPercentage }
    }
    foreach ($v in $eng.Values) { if ($v -gt $g) { $g = $v } }
    $s.Gpu.gpu = [math]::Round((Dmin (100) ($g)), 1)
    $smi = Get-Command nvidia-smi -ErrorAction SilentlyContinue
    if ($smi) {
      try {
        $line = (& $smi.Source '--query-gpu=temperature.gpu,memory.used,memory.total,power.draw,utilization.gpu' '--format=csv,noheader,nounits' 2>$null | Select-Object -First 1)
        $p = ([string]$line).Split(',') | ForEach-Object { $_.Trim() }
        if ($p.Count -ge 5) { $s.Gpu.temp = [int]$p[0]; $s.Gpu.memUsed = [int]$p[1]; $s.Gpu.memTotal = [int]$p[2]; $s.Gpu.power = [int][double]$p[3]; if ([double]$p[4] -gt $s.Gpu.gpu) { $s.Gpu.gpu = [double]$p[4] } }
      } catch { }
    }

    # busiest processes
    $now = Get-Date; $n = Get-CoreCount; $list = @()
    $dt = 0; if ($s.ProcAt) { $dt = ($now - $s.ProcAt).TotalSeconds }
    $cur = @{}
    foreach ($p in (Get-Process -ErrorAction SilentlyContinue)) {
      try {
        $cpuT = $p.TotalProcessorTime.TotalSeconds
        $cur[$p.Id] = $cpuT
        $pct = 0.0
        if ($dt -gt 0 -and $s.ProcPrev.ContainsKey($p.Id)) { $pct = 100 * ($cpuT - $s.ProcPrev[$p.Id]) / ($dt * $n) }
        if ($pct -gt 0.05) { $list += @{ name = $p.ProcessName; pid = $p.Id; cpu = [math]::Round($pct, 1); ramMb = [int]($p.WorkingSet64 / 1MB) } }
      } catch { }
    }
    $s.ProcPrev = $cur; $s.ProcAt = $now
    $s.Top = @($list | Sort-Object { -$_.cpu } | Select-Object -First 6)
  }
  $out.gpu = $s.Gpu.gpu; $out.gpuTemp = $s.Gpu.temp; $out.gpuMemUsedMb = $s.Gpu.memUsed; $out.gpuMemTotalMb = $s.Gpu.memTotal; $out.gpuPower = $s.Gpu.power
  $out.top = $s.Top
  return $out
}

# ---------------------------------------------------------------- ping (only when switched on)
function Get-PingStats {
  param([string]$HostName)
  $s = $script:Samp
  if (-not $s.Ping) { $s.Ping = New-Object System.Net.NetworkInformation.Ping }
  $ms = $null
  if ($script:Mock) { $ms = [math]::Round(14 + 3 * $s.Rnd.NextDouble() + $(if ($s.Rnd.Next(0, 40) -eq 0) { 25 } else { 0 }), 1) }
  else {
    try { $r = $s.Ping.Send($HostName, 900); if ($r.Status -eq 'Success') { $ms = [double]$r.RoundtripTime } } catch { }
  }
  [void]$s.PingHist.Add($ms)
  while ($s.PingHist.Count -gt 30) { $s.PingHist.RemoveAt(0) }
  $ok = @($s.PingHist | Where-Object { $null -ne $_ })
  $loss = 0.0; if ($s.PingHist.Count -gt 0) { $loss = [math]::Round(100 * ($s.PingHist.Count - $ok.Count) / $s.PingHist.Count, 1) }
  $jit = 0.0
  if ($ok.Count -gt 1) { $d = 0.0; for ($i = 1; $i -lt $ok.Count; $i++) { $d += [math]::Abs($ok[$i] - $ok[$i - 1]) }; $jit = [math]::Round($d / ($ok.Count - 1), 1) }
  return @{ host = $HostName; ms = $ms; jitter = $jit; loss = $loss; samples = $s.PingHist.Count }
}

# ---------------------------------------------------------------- the loop (runs in its own runspace)
# $S is the synchronized state table shared with the server:  S.Live  S.Hist  S.Settings  S.Stop
function Invoke-SamplerLoop {
  param($S)
  $t0 = Get-Date; $tick = 0; $lastHost = ''
  while (-not $S.Stop) {
    $start = Get-Date
    $t = ($start - $t0).TotalSeconds
    try {
      if ($script:Mock) { $smp = Get-MockSample $t } else { $smp = Get-RealSample $t (($tick % 2) -eq 0) }
      if ($S.Settings.PingOn) {
        if ($lastHost -ne [string]$S.Settings.PingHost) { $script:Samp.PingHist.Clear(); $lastHost = [string]$S.Settings.PingHost }
        $smp.ping = Get-PingStats $lastHost
      }
      $smp.t = [int64]([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds())
      $S.Live = $smp
      $h = $S.Hist
      [void]$h.Add($smp)
      while ($h.Count -gt 120) { $h.RemoveAt(0) }
    } catch { $S.SamplerError = $_.Exception.Message }
    $tick++
    $wait = 1000 - [int]((Get-Date) - $start).TotalMilliseconds
    if ($wait -gt 20) { Start-Sleep -Milliseconds $wait }
  }
}
'@

$JT_SERVER = @'
# ============================================================================
#  SERVER - the window you see
#  A tiny web server that only listens on 127.0.0.1 (this PC, nothing else)
#  and serves the app to a browser window. Every request needs the secret
#  token of this launch (a cookie), the page can only ask for a fixed list of
#  things (see Handle-Request), and it can only run tweaks that exist in the
#  list above - it cannot run any other command.
# ============================================================================

# ---------------------------------------------------------------- settings (the only things saved besides the journal and logs)
function Get-Settings {
  $f = Join-Path $script:Cfg.Root 'settings.json'
  $d = @{ PingOn = $false; PingHost = '1.1.1.1'; UpdateCheck = $false; Motion = $true }
  if (Test-Path -LiteralPath $f) {
    try {
      $j = Get-Content -LiteralPath $f -Raw | ConvertFrom-Json
      if ($null -ne $j.PingOn) { $d.PingOn = [bool]$j.PingOn }
      if ($j.PingHost -and ([string]$j.PingHost) -match '^[A-Za-z0-9.\-:]{1,80}$') { $d.PingHost = [string]$j.PingHost }
      if ($null -ne $j.UpdateCheck) { $d.UpdateCheck = [bool]$j.UpdateCheck }
      if ($null -ne $j.Motion) { $d.Motion = [bool]$j.Motion }
    } catch { }
  }
  return $d
}

function Save-Settings {
  param($S)
  $f = Join-Path $script:Cfg.Root 'settings.json'
  [IO.File]::WriteAllText($f, (ConvertTo-Json -InputObject @{ PingOn = [bool]$S.Settings.PingOn; PingHost = [string]$S.Settings.PingHost; UpdateCheck = [bool]$S.Settings.UpdateCheck; Motion = [bool]$S.Settings.Motion }))
}

# ---------------------------------------------------------------- responses
function Send-Bytes {
  param($Ctx, [byte[]]$Bytes, [string]$Type, [int]$Code = 200)
  try {
    $r = $Ctx.Response
    $r.StatusCode = $Code
    $r.ContentType = $Type
    $r.ContentLength64 = $Bytes.Length
    $r.Headers.Set('Cache-Control', 'no-store')
    $r.Headers.Set('X-Content-Type-Options', 'nosniff')
    $r.Headers.Set('Referrer-Policy', 'no-referrer')
    $r.Headers.Set('Content-Security-Policy', "default-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; base-uri 'none'; form-action 'none'; frame-ancestors 'none'")
    $r.OutputStream.Write($Bytes, 0, $Bytes.Length)
  } catch { } finally { try { $Ctx.Response.OutputStream.Close() } catch { } }
}

function Send-Json {
  param($Ctx, $Obj, [int]$Code = 200)
  $json = ConvertTo-Json -InputObject $Obj -Depth 8 -Compress
  Send-Bytes $Ctx ([Text.Encoding]::UTF8.GetBytes($json)) 'application/json; charset=utf-8' $Code
}

function Send-Text {
  param($Ctx, [string]$Text, [string]$Type, [int]$Code = 200)
  Send-Bytes $Ctx ([Text.Encoding]::UTF8.GetBytes($Text)) $Type $Code
}

function Read-Body {
  param($Ctx)
  $req = $Ctx.Request
  if (-not $req.HasEntityBody) { return $null }
  if ($req.ContentLength64 -gt 65536) { throw 'request too large' }
  $sr = New-Object IO.StreamReader($req.InputStream, [Text.Encoding]::UTF8)
  try { $txt = $sr.ReadToEnd() } finally { $sr.Dispose() }
  if (-not $txt) { return $null }
  return ($txt | ConvertFrom-Json)
}

function Read-LogLines {
  param([string]$File, [int]$From)
  $out = New-Object System.Collections.ArrayList
  if (-not $File -or -not (Test-Path -LiteralPath $File)) { return @{ lines = @(); next = $From } }
  $fs = $null
  try {
    $fs = New-Object IO.FileStream($File, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $sr = New-Object IO.StreamReader($fs)
    $i = 0
    while (-not $sr.EndOfStream) { $l = $sr.ReadLine(); if ($i -ge $From -and $l) { [void]$out.Add($l) }; $i++ }
    return @{ lines = @($out); next = $i }
  } catch { return @{ lines = @(); next = $From } }
  finally { if ($fs) { $fs.Dispose() } }
}

# ---------------------------------------------------------------- one tweak run (in its own runspace)
function Start-Run {
  param($S, [string[]]$Ids)
  if ($S.Run -and $S.Run.Running) { return $null }
  $norm = @(Resolve-Sequence $Ids)
  if ($norm.Count -eq 0) { return $null }
  Initialize-Jiff
  $stamp = [DateTime]::Now.ToString('yyyyMMdd_HHmmss', [Globalization.CultureInfo]::InvariantCulture)
  $logFile = Join-Path $script:Cfg.LogDir ('jiff_' + $stamp + '.log')
  $run = [hashtable]::Synchronized(@{ Id = [guid]::NewGuid().ToString('N').Substring(0, 8); Running = $true; Ids = $norm; Log = $logFile; Started = (Get-Date).ToString('s'); Result = $null; Error = $null; Before = $null; After = $null; Printed = 0 })
  $run.Before = Get-Snapshot
  $S.Run = $run
  (Get-Ctl).Cancel = $false
  $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
  $ps = [powershell]::Create(); $ps.Runspace = $rs
  $body = {
    param($ids, $log, $run, $S)
    try { $run.Result = Invoke-JiffSequence -Ids $ids -LogFile $log }
    catch { $run.Error = $_.Exception.Message }
    finally {
      try { $run.After = Get-Snapshot } catch { }
      $run.Ended = (Get-Date).ToString('s'); $run.Running = $false; $S.StatesDirty = $true
    }
  }
  [void]$ps.AddScript($S.Src.Engine, $false).AddStatement()
  [void]$ps.AddScript($body.ToString(), $false).AddArgument($norm).AddArgument($logFile).AddArgument($run).AddArgument($S)
  $run.Ps = $ps; $run.Rs = $rs
  $run.Handle = $ps.BeginInvoke()
  return $run
}

# ---------------------------------------------------------------- update check (only when you press the button or switch it on)
function Get-UpdateInfo {
  $url = 'https://raw.githubusercontent.com/jiff-dev/Jiff-free-Tweaks-V9/main/version.txt'
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $txt = (Invoke-WebRequest -UseBasicParsing -Uri $url -TimeoutSec 6).Content.ToString().Trim()
    $latest = ($txt -split "\s+")[0]
    $newer = $false
    try { $newer = ([version]$latest) -gt ([version]$script:Cfg.Version) } catch { }
    return @{ ok = $true; current = $script:Cfg.Version; latest = $latest; newer = $newer; url = $url; page = 'https://github.com/jiff-dev/Jiff-free-Tweaks-V9' }
  } catch { return @{ ok = $false; current = $script:Cfg.Version; error = $_.Exception.Message; url = $url } }
}

function Get-DiagData {
  param($S)
  $edge = Find-Browser
  $count = 0; if (Test-Path -LiteralPath $script:Cfg.LogDir) { $count = @(Get-ChildItem -LiteralPath $script:Cfg.LogDir -Filter 'jiff_*.log').Count }
  return @{
    admin = (Test-IsAdmin); ps = $PSVersionTable.PSVersion.ToString(); edition = [string]$PSVersionTable.PSEdition; bits64 = [Environment]::Is64BitProcess
    browser = $edge; dataDir = $script:Cfg.Root; journalRows = (Journal-Counts).total; logs = $count; port = $S.Port; mock = $script:Mock
    sampler = $S.SamplerError; version = $script:Cfg.Version
  }
}

function Find-Browser {
  $c = @()
  foreach ($base in @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:LOCALAPPDATA)) {
    if ($base) {
      $c += (Join-Path $base 'Microsoft\Edge\Application\msedge.exe')
      $c += (Join-Path $base 'Google\Chrome\Application\chrome.exe')
      $c += (Join-Path $base 'BraveSoftware\Brave-Browser\Application\brave.exe')
    }
  }
  foreach ($p in $c) { if (Test-Path -LiteralPath $p) { return $p } }
  return $null
}

# ---------------------------------------------------------------- the request handler
function Handle-Request {
  param($Ctx, $S)
  $req = $Ctx.Request
  $path = $req.Url.AbsolutePath
  $method = $req.HttpMethod
  $q = $req.QueryString

  # only this PC, and only under our own address (blocks DNS rebinding)
  $h = [string]$req.Headers['Host']
  if ($h -ne ('127.0.0.1:' + $S.Port) -and $h -ne ('localhost:' + $S.Port)) { Send-Text $Ctx 'bad host' 'text/plain' 400; return }

  # first visit carries the token in the address, then it lives in a cookie
  if ($path -eq '/' -and $q['k']) {
    if ([string]$q['k'] -ceq [string]$S.Token) {
      $Ctx.Response.Headers.Add('Set-Cookie', ('jt=' + $S.Token + '; Path=/; HttpOnly; SameSite=Strict'))
      $Ctx.Response.StatusCode = 302; $Ctx.Response.Headers.Set('Location', '/'); $Ctx.Response.OutputStream.Close(); return
    }
    Send-Text $Ctx 'forbidden' 'text/plain' 403; return
  }
  $ck = $req.Cookies['jt']
  if (-not $ck -or $ck.Value -cne [string]$S.Token) { Send-Text $Ctx 'Open this app from the window the script started.' 'text/plain' 403; return }
  $S.LastSeen = Get-Date

  if ($method -eq 'POST') {
    $o = [string]$req.Headers['Origin']
    if ($o -and $o -ne ('http://127.0.0.1:' + $S.Port)) { Send-Text $Ctx 'bad origin' 'text/plain' 403; return }
    $sf = [string]$req.Headers['Sec-Fetch-Site']
    if ($sf -and $sf -ne 'same-origin' -and $sf -ne 'none') { Send-Text $Ctx 'bad origin' 'text/plain' 403; return }
    if (([string]$req.ContentType) -notlike 'application/json*') { Send-Text $Ctx 'json only' 'text/plain' 415; return }
  } elseif ($method -ne 'GET') { Send-Text $Ctx 'method' 'text/plain' 405; return }

  # ----- static files
  if ($method -eq 'GET' -and -not $path.StartsWith('/api/')) {
    $name = $path.TrimStart('/'); if (-not $name) { $name = 'index.html' }
    $a = $S.Assets[$name]
    if ($a) { Send-Bytes $Ctx $a.Bytes $a.Type; return }
    Send-Text $Ctx 'not found' 'text/plain' 404; return
  }

  $b = $null
  if ($method -eq 'POST') { $b = Read-Body $Ctx }

  switch ($path) {
    '/api/hello'   { Send-Json $Ctx @{ ok = $true; version = $script:Cfg.Version; started = $S.Started } }
    '/api/catalog' { Send-Text $Ctx $S.CatalogJson 'application/json; charset=utf-8' }
    '/api/sys'     { Send-Json $Ctx $S.Sys }
    '/api/live' {
      $o = @{ sample = $S.Live }
      if ($q['hist'] -eq '1') { $o.hist = @($S.Hist.ToArray()) }
      Send-Json $Ctx $o
    }
    '/api/states' {
      $m = @{}; foreach ($k in @($S.States.Keys)) { $m[$k] = $S.States[$k] }
      Send-Json $Ctx @{ map = $m; ver = $S.StatesVer; busy = [bool]$S.StatesBusy }
    }
    '/api/refresh' { $S.StatesDirty = $true; Send-Json $Ctx @{ ok = $true } }
    '/api/run' {
      if ($method -eq 'POST') {
        $ids = @($b.ids | ForEach-Object { [string]$_ })
        foreach ($i in $ids) { if (-not $script:Tweaks.Contains($i)) { Send-Json $Ctx @{ error = 'unknown tweak' } 400; return } }
        $run = Start-Run $S $ids
        if (-not $run) { Send-Json $Ctx @{ error = 'a run is already in progress' } 409; return }
        Send-Json $Ctx @{ ok = $true; id = $run.Id; ids = $run.Ids }
      } else {
        $run = $S.Run
        if (-not $run) { Send-Json $Ctx @{ running = $false; none = $true }; return }
        $from = 0; [void][int]::TryParse([string]$q['from'], [ref]$from)
        $ll = Read-LogLines $run.Log $from
        $o = @{ id = $run.Id; running = [bool]$run.Running; ids = $run.Ids; lines = $ll.lines; next = $ll.next; started = $run.Started }
        if (-not $run.Running) { $o.result = $run.Result; $o.error = $run.Error; $o.before = $run.Before; $o.after = $run.After }
        Send-Json $Ctx $o
      }
    }
    '/api/cancel' { (Get-Ctl).Cancel = $true; Send-Json $Ctx @{ ok = $true } }
    '/api/preview' {
      $ids = @($b.ids | ForEach-Object { [string]$_ })
      foreach ($i in $ids) { if (-not $script:Tweaks.Contains($i)) { Send-Json $Ctx @{ error = 'unknown tweak' } 400; return } }
      Send-Json $Ctx @{ items = @(Get-PreviewData $ids) }
    }
    '/api/changes' { Send-Json $Ctx @{ rows = @(Get-ChangeList); counts = (Journal-Counts) } }
    '/api/impact'  { Send-Json $Ctx @{ now = (Get-Snapshot); run = $(if ($S.Run) { @{ before = $S.Run.Before; after = $S.Run.After; running = [bool]$S.Run.Running } } else { $null }) } }
    '/api/logs' {
      $files = @()
      if (Test-Path -LiteralPath $script:Cfg.LogDir) {
        $files = @(Get-ChildItem -LiteralPath $script:Cfg.LogDir -Filter 'jiff_*.log' | Sort-Object LastWriteTime -Descending | Select-Object -First 40 | ForEach-Object { @{ name = $_.Name; size = $_.Length; time = $_.LastWriteTime.ToString('s') } })
      }
      Send-Json $Ctx @{ files = $files; dir = $script:Cfg.LogDir }
    }
    '/api/log' {
      $n = [string]$q['name']
      if ($n -notmatch '^jiff_[0-9_]+\.log$') { Send-Json $Ctx @{ error = 'bad name' } 400; return }
      $f = Join-Path $script:Cfg.LogDir $n
      if (-not (Test-Path -LiteralPath $f)) { Send-Json $Ctx @{ error = 'not found' } 404; return }
      $ll = Read-LogLines $f 0
      $lines = @($ll.lines); if ($lines.Count -gt 3000) { $lines = @($lines | Select-Object -Last 3000) }
      Send-Json $Ctx @{ name = $n; lines = $lines }
    }
    '/api/settings' {
      if ($method -eq 'POST') {
        if ($null -ne $b.PingOn) { $S.Settings.PingOn = [bool]$b.PingOn }
        if ($b.PingHost -and ([string]$b.PingHost) -match '^[A-Za-z0-9.\-:]{1,80}$') { if ([string]$b.PingHost -ne [string]$S.Settings.PingHost) { $S.PingReset = $true }; $S.Settings.PingHost = [string]$b.PingHost }
        if ($null -ne $b.UpdateCheck) { $S.Settings.UpdateCheck = [bool]$b.UpdateCheck }
        if ($null -ne $b.Motion) { $S.Settings.Motion = [bool]$b.Motion }
        Save-Settings $S
      }
      Send-Json $Ctx @{ PingOn = [bool]$S.Settings.PingOn; PingHost = [string]$S.Settings.PingHost; UpdateCheck = [bool]$S.Settings.UpdateCheck; Motion = [bool]$S.Settings.Motion }
    }
    '/api/update' { Send-Json $Ctx (Get-UpdateInfo) }
    '/api/diag'   { Send-Json $Ctx (Get-DiagData $S) }
    '/api/reboot' {
      if (-not $script:Mock) { [void](Exe 'shutdown.exe' @('/r', '/t', '20', '/c', 'Jiff Tweaks: restarting to finish applying changes')) }
      Send-Json $Ctx @{ ok = $true; mock = $script:Mock }
    }
    '/api/abort-reboot' { if (-not $script:Mock) { [void](Exe 'shutdown.exe' @('/a')) }; Send-Json $Ctx @{ ok = $true } }
    '/api/open-logs'    { if (-not $script:Mock) { Start-Process explorer.exe -ArgumentList ('"' + $script:Cfg.LogDir + '"') }; Send-Json $Ctx @{ ok = $true } }
    '/api/quit'         { Send-Json $Ctx @{ ok = $true }; $S.Quit = $true }
    default { Send-Json $Ctx @{ error = 'not found' } 404 }
  }
}

# every handler runspace runs this loop; HttpListener hands each request to one of them
function Invoke-HandlerLoop {
  param($S, $Listener)
  while ($Listener.IsListening -and -not $S.Stop) {
    $ctx = $null
    try { $ctx = $Listener.GetContext() } catch { break }
    try { Handle-Request $ctx $S }
    catch { try { Send-Json $ctx @{ error = $_.Exception.Message } 500 } catch { } }
  }
}

# background recomputation of "is this tweak on?" (runs in its own runspace so the page never waits)
function Invoke-StatusLoop {
  param($S)
  while (-not $S.Stop) {
    if ($S.StatesDirty) {
      $S.StatesDirty = $false; $S.StatesBusy = $true
      try {
        # the slow ones (Store app list, every device) go last so the rest shows up quickly
        $slow = @('devpower', 'debloat_curated', 'debloat_extended', 'debloat_onedrive', 'apply_full')
        $order = @(@($script:Tweaks.Keys | Where-Object { $slow -notcontains $_ }) + @($slow | Where-Object { $script:Tweaks.Contains($_) }))
        foreach ($id in $order) {
          if ($S.Stop) { break }
          $S.States[$id] = Get-TweakState $script:Tweaks[$id]
          if ($S.StatesDirty) { break }   # a newer request came in, start over
        }
        $S.StatesVer = [int]$S.StatesVer + 1
      } catch { $S.StatusError = $_.Exception.Message }
      $S.StatesBusy = $false
    }
    Start-Sleep -Milliseconds 250
  }
}

# ---------------------------------------------------------------- start everything
function New-Worker {
  param($S, [string]$Script, $Arg1, $Arg2)
  $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
  $ps = [powershell]::Create(); $ps.Runspace = $rs
  [void]$ps.AddScript($S.Src.Engine, $false).AddStatement()
  [void]$ps.AddScript($S.Src.Server, $false).AddStatement()
  $cmd = $ps.AddScript($Script, $false)
  [void]$cmd.AddArgument($Arg1)
  [void]$cmd.AddArgument($Arg2)
  return @{ Ps = $ps; Rs = $rs; Handle = $ps.BeginInvoke() }
}

function Start-JiffApp {
  param([hashtable]$Src, [hashtable]$Assets, [switch]$NoBrowser, [switch]$ForceBrowser, [int]$Port = 0)
  Initialize-Jiff
  # random loopback port, retry if taken
  $listener = $null; $S = $null
  $rnd = New-Object System.Random
  for ($try = 0; $try -lt 30; $try++) {
    $p = $Port; if ($p -le 0) { $p = $rnd.Next(49200, 65000) }
    $l = New-Object System.Net.HttpListener
    $l.Prefixes.Add('http://127.0.0.1:' + $p + '/')
    try { $l.Start(); $listener = $l; $Port = $p; break } catch { $l.Close(); if ($Port -gt 0) { throw } }
  }
  if (-not $listener) { throw 'could not open a local port' }

  $tok = New-Object byte[] 24; $rng = New-Object Security.Cryptography.RNGCryptoServiceProvider; $rng.GetBytes($tok)
  $token = ($tok | ForEach-Object { $_.ToString('x2') }) -join ''
  $settings = [hashtable]::Synchronized((Get-Settings))
  $S = [hashtable]::Synchronized(@{
      Port = $Port; Token = $token; Src = $Src; Assets = $Assets; Stop = $false; Quit = $false; Started = (Get-Date).ToString('s'); LastSeen = Get-Date
      Live = $null; Hist = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList)); Settings = $settings
      States = [hashtable]::Synchronized(@{}); StatesVer = 0; StatesDirty = $true; StatesBusy = $false
      Run = $null; CatalogJson = (ConvertTo-Json -InputObject (Get-CatalogData) -Depth 8 -Compress); Sys = (Get-SystemInfo)
    })

  $workers = @()
  $workers += New-Worker $S 'param($S, $L) Invoke-SamplerLoop $S' $S $null
  $workers += New-Worker $S 'param($S, $L) Invoke-StatusLoop $S' $S $null
  for ($i = 0; $i -lt 4; $i++) { $workers += New-Worker $S 'param($S, $L) Invoke-HandlerLoop $S $L' $S $listener }

  $url = 'http://127.0.0.1:' + $Port + '/?k=' + $token
  Write-Host ''
  Write-Host ('  Jiff Tweaks ' + $script:Cfg.Version) -ForegroundColor Green
  Write-Host ('  Window address (only this PC can reach it): http://127.0.0.1:' + $Port + '/')
  Write-Host '  Close the window or press Ctrl+C here to stop.'
  Write-Host ''

  $proc = $null; $hold = $false
  try {
    # open the window: a browser in app mode with its own profile, so closing it ends the script
    if (-not $NoBrowser) {
      $b = Find-Browser
      if ($b -and -not $ForceBrowser) {
        $prof = Join-Path $script:Cfg.Root 'window'
        try {
          $proc = Start-Process -FilePath $b -ArgumentList @('--app=' + $url, ('--user-data-dir=' + $prof), '--window-size=1320,860', '--no-first-run', '--no-default-browser-check', '--disable-extensions') -PassThru
          Start-Sleep -Milliseconds 2500
          if ($proc.HasExited) { $proc = $null; $hold = $true }   # the browser handed the window to a running instance
        } catch { $proc = $null; $hold = $true }
      } else {
        try { Start-Process $url } catch { Write-Host ('  Open this address in your browser: ' + $url) -ForegroundColor Cyan }
        $hold = $true
      }
    } else { Write-Host ('  Open: ' + $url) -ForegroundColor Cyan; $hold = $true }

    # keep running until the window closes
    while (-not $S.Quit) {
      Start-Sleep -Milliseconds 300
      if ($proc -and $proc.HasExited) { break }
      if ($hold -and $S.LastSeen -and ((Get-Date) - $S.LastSeen).TotalSeconds -gt 180) { Write-Host '  No page activity for 3 minutes, stopping.'; break }
      $run = $S.Run
      if ($run) {
        $ll = Read-LogLines $run.Log ([int]$run.Printed)
        foreach ($l in $ll.lines) {
          $c = 'Gray'
          if ($l -match '\[FAIL\]') { $c = 'Red' } elseif ($l -match '\[OK\]') { $c = 'Green' } elseif ($l -match '\[(TWEAK|STEP|SEQ)\]') { $c = 'Cyan' } elseif ($l -match '\[NOTE\]') { $c = 'Yellow' } elseif ($l -match '\[SKIP\]') { $c = 'DarkGray' }
          if ($l -notmatch '\[OK\]|\[SKIP\]' -or $script:VerboseConsole) { Write-Host ('  ' + $l) -ForegroundColor $c }
        }
        $run.Printed = $ll.next
      }
    }
  } finally {
    $S.Stop = $true
    (Get-Ctl).Cancel = $true
    try { $listener.Stop(); $listener.Close() } catch { }
    # let every background loop finish by itself, then release the runspaces
    $all = @($workers)
    if ($S.Run -and $S.Run.Ps) { $all += @{ Ps = $S.Run.Ps; Rs = $S.Run.Rs; Handle = $S.Run.Handle } }
    foreach ($w in $all) { try { [void]$w.Handle.AsyncWaitHandle.WaitOne(2500) } catch { } }
    foreach ($w in $all) {
      try { if (-not $w.Handle.IsCompleted) { $w.Ps.Stop() } } catch { }
      try { $w.Ps.Dispose() } catch { }
      try { $w.Rs.Dispose() } catch { }
    }
    if ($proc -and -not $proc.HasExited) { try { [void]$proc.CloseMainWindow() } catch { } }
    Write-Host '  Stopped.'
  }
}

# ============================================================================
#  COMMAND LINE - use the tweaks without the window
#    .\Jiff-Tweaks.ps1 -List                 every tweak with its state
#    .\Jiff-Tweaks.ps1 -Status               only what is on
#    .\Jiff-Tweaks.ps1 -Code reg_gaming      show the code of a tweak
#    .\Jiff-Tweaks.ps1 -Preview svc_basic    show what it would change
#    .\Jiff-Tweaks.ps1 -Preset safe         safe | competitive | full
#    .\Jiff-Tweaks.ps1 -Apply reg_gaming,privacy
#    .\Jiff-Tweaks.ps1 -Restore              put every saved original back
#    .\Jiff-Tweaks.ps1 -NoGui                text menu
#    .\Jiff-Tweaks.ps1 -SelfTest             checks that everything the tweaks need works on this PC
#  Add -Yes to skip the confirmation question.
# ============================================================================

function Write-StateWord {
  param($State)
  switch ($State) {
    'on'      { Write-Host ' ON      ' -NoNewline -ForegroundColor Green }
    'partial' { Write-Host ' PARTIAL ' -NoNewline -ForegroundColor Yellow }
    'off'     { Write-Host ' off     ' -NoNewline -ForegroundColor DarkGray }
    default   { Write-Host '         ' -NoNewline }
  }
}

function Show-TweakList {
  param([switch]$OnlyOn)
  Write-Host ''
  Write-Host '  Checking this PC ...' -ForegroundColor DarkGray
  foreach ($pk in $script:Pages.Keys) {
    $p = $script:Pages[$pk]; $rows = @()
    foreach ($id in $p.Ids) {
      $t = $script:Tweaks[$id]
      $st = $null; if ($t.Kind -eq 'tweak') { $st = Get-TweakState $t }
      if ($OnlyOn -and $st -ne 'on' -and $st -ne 'partial') { continue }
      $rows += @{ t = $t; st = $st }
    }
    if ($rows.Count -eq 0) { continue }
    Write-Host ''
    Write-Host ('  ' + $p.Title.ToUpper()) -ForegroundColor Cyan
    foreach ($r in $rows) {
      Write-Host ('  ' + $r.t.Id.PadRight(22)) -NoNewline
      Write-StateWord $r.st
      $rc = 'Gray'; if ($r.t.Risk -eq 'aggressive') { $rc = 'Yellow' } elseif ($r.t.Risk -eq 'security' -or $r.t.Risk -eq 'danger') { $rc = 'Red' }
      Write-Host ($r.t.Risk.PadRight(11)) -NoNewline -ForegroundColor $rc
      Write-Host $r.t.Title
    }
  }
  Write-Host ''
}

function Show-Preview {
  param([string[]]$Ids)
  foreach ($it in @(Get-PreviewData $Ids)) {
    Write-Host ''
    Write-Host ('  ' + $it.title + '  (' + $it.count + ' changes)') -ForegroundColor Cyan
    $n = 0
    foreach ($c in $it.changes) {
      $n++; if ($n -gt 25) { Write-Host ('    ... and ' + ($it.count - 25) + ' more') -ForegroundColor DarkGray; break }
      $left = [string]$c.target; if ($c.name) { $left += '  ' + $c.name }
      if ($null -ne $c.new) { Write-Host ('    ' + $left + '   ' + [string]$c.current + ' -> ' + [string]$c.new) } else { Write-Host ('    ' + $left) }
    }
  }
  Write-Host ''
}

function Confirm-Run {
  param([string[]]$Ids, [switch]$Yes)
  $ids = @(Resolve-Sequence $Ids)
  if ($ids.Count -eq 0) { Write-Host '  Nothing to do.'; return $null }
  Write-Host ''
  Write-Host '  This will apply:' -ForegroundColor Cyan
  $hard = $false; $reboot = $false
  foreach ($id in $ids) {
    $t = $script:Tweaks[$id]
    $rc = 'Gray'; if ($t.Risk -eq 'aggressive') { $rc = 'Yellow' } elseif ($t.Risk -eq 'security' -or $t.Risk -eq 'danger') { $rc = 'Red' }
    Write-Host ('    ' + $t.Title.PadRight(32)) -NoNewline
    Write-Host $t.Risk -ForegroundColor $rc
    if ($t.Risk -eq 'security' -or $t.Risk -eq 'danger') { $hard = $true }
    if ($t.Reboot) { $reboot = $true }
  }
  Write-Host ''
  Write-Host '  Every original value is saved first. Restore puts them back (removed Store apps excluded).' -ForegroundColor DarkGray
  if ($Yes) { return $ids }
  if ($hard) {
    Write-Host '  This includes changes that lower security or cannot be undone.' -ForegroundColor Red
    $a = Read-Host '  Type YES (capital letters) to continue'
    if ($a -cne 'YES') { Write-Host '  Cancelled.'; return $null }
  } else {
    $a = Read-Host '  Continue? [y/N]'
    if ($a -notmatch '^(y|yes)$') { Write-Host '  Cancelled.'; return $null }
  }
  return $ids
}

function Invoke-CliRun {
  param([string[]]$Ids, [switch]$Yes)
  $ids = Confirm-Run $Ids -Yes:$Yes
  if (-not $ids) { return }
  Initialize-Jiff
  $log = Join-Path $script:Cfg.LogDir ('jiff_' + ([DateTime]::Now.ToString('yyyyMMdd_HHmmss', [Globalization.CultureInfo]::InvariantCulture)) + '.log')
  Write-Host ''
  $r = Invoke-JiffSequence -Ids $ids -LogFile $log -Echo
  Write-Host ''
  Write-Host ('  Done: ' + $r.Ok + ' changes, ' + $r.Fail + ' failed, ' + $r.Skip + ' skipped.') -ForegroundColor Green
  Write-Host ('  Log: ' + $log)
  if ($r.Reboot) { Write-Host '  Restart Windows to finish applying.' -ForegroundColor Yellow }
}

function Show-Menu {
  while ($true) {
    Write-Host ''
    Write-Host ('  JIFF TWEAKS ' + $script:Cfg.Version) -ForegroundColor Green
    Write-Host '  1  Safe profile'
    Write-Host '  2  Competitive profile'
    Write-Host '  3  Full profile'
    Write-Host '  4  List every tweak and its state'
    Write-Host '  5  Apply tweaks by id'
    Write-Host '  6  Show the code of a tweak'
    Write-Host '  7  Restore everything'
    Write-Host '  0  Quit'
    $c = Read-Host '  Choose'
    switch ($c) {
      '1' { Invoke-CliRun $script:Profiles['safe'].Ids }
      '2' { Invoke-CliRun $script:Profiles['competitive'].Ids }
      '3' { Invoke-CliRun $script:Profiles['full'].Ids }
      '4' { Show-TweakList }
      '5' { $x = Read-Host '  Ids, separated by commas'; Invoke-CliRun @($x -split '[,\s]+' | Where-Object { $_ }) }
      '6' { $x = Read-Host '  Id'; Show-Code $x }
      '7' { Invoke-CliRun @('restore_all') }
      '0' { return }
    }
  }
}

function Show-Code {
  param([string]$Id)
  $c = Get-TweakCode $Id
  if (-not $c) { Write-Host ('  Unknown tweak: ' + $Id) -ForegroundColor Red; return }
  Write-Host ''
  Write-Host ('  ' + $c.title + '  (' + $c.id + ')') -ForegroundColor Cyan
  Write-Host ''
  foreach ($l in ($c.apply -split "`n")) { Write-Host ('    ' + $l) }
  if ($c.verbs.Count -gt 0) {
    Write-Host ''
    Write-Host '  Functions used:' -ForegroundColor Cyan
    foreach ($v in $c.verbs) { Write-Host ('    ' + $v.name.PadRight(26) + $v.help) -ForegroundColor DarkGray }
  }
  Write-Host ''
}

# entry point for command line use. $O is the parameter table of the script.
function Invoke-JiffCli {
  param($O)
  Initialize-Jiff
  if ($O.SelfTest) { Invoke-SelfTest; return }
  if ($O.Code) { foreach ($i in @($O.Code)) { Show-Code $i }; return }
  if ($O.List) { Show-TweakList; return }
  if ($O.Status) { Show-TweakList -OnlyOn; return }
  if ($O.Preview) { Show-Preview @($O.Preview); return }
  if ($O.Restore) { Invoke-CliRun @('restore_all') -Yes:$O.Yes; return }
  if ($O.Preset) {
    $k = ([string]$O.Preset).ToLower()
    if (-not $script:Profiles.Contains($k)) { Write-Host ('  Unknown profile. Use: ' + ($script:Profiles.Keys -join ', ')) -ForegroundColor Red; return }
    Invoke-CliRun $script:Profiles[$k].Ids -Yes:$O.Yes; return
  }
  if ($O.Apply) {
    $bad = @($O.Apply | Where-Object { -not $script:Tweaks.Contains($_) })
    if ($bad.Count -gt 0) { Write-Host ('  Unknown tweak id: ' + ($bad -join ', ') + '   (run -List to see the ids)') -ForegroundColor Red; return }
    Invoke-CliRun @($O.Apply) -Yes:$O.Yes; return
  }
  Show-Menu
}

# ---------------------------------------------------------------- self test: checks every kind of access the tweaks need, changes nothing that matters
#   .\Jiff-Tweaks.ps1 -SelfTest        (paste the output to someone if a line says FAIL)
function Invoke-SelfTest {
  $res = New-Object System.Collections.ArrayList
  $check = {
    param($Name, [scriptblock]$Body)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
      $r = & $Body
      $ok = $true; $msg = ''
      if ($r -is [hashtable] -and $r.ContainsKey('ok')) { $ok = [bool]$r.ok; $msg = [string]$r.msg } elseif ($null -ne $r) { $msg = [string]$r }
    } catch { $ok = $false; $msg = $_.Exception.Message }
    $sw.Stop()
    $tag = 'PASS'; $col = 'Green'; if (-not $ok) { $tag = 'FAIL'; $col = 'Red' }
    Write-Host ('  ' + $tag + '  ' + $Name.PadRight(44) + ' ' + $msg + '  (' + $sw.ElapsedMilliseconds + ' ms)') -ForegroundColor $col
    [void]$res.Add($ok)
  }
  Write-Host ''
  Write-Host ('  Jiff Tweaks ' + $script:Cfg.Version + ' self test') -ForegroundColor Cyan
  Write-Host ('  PowerShell ' + $PSVersionTable.PSVersion + ' (' + $PSVersionTable.PSEdition + '), admin: ' + (Test-IsAdmin) + ', 64-bit process: ' + [Environment]::Is64BitProcess + ', mock: ' + $script:Mock)
  Write-Host ''
  $si = Get-SystemInfo
  Write-Host ('  ' + $si.os + ' ' + $si.version + ' build ' + $si.build + ' | ' + $si.cpu + ' | ' + $si.ramMb + ' MB | GPU: ' + ($si.gpus -join ' / '))
  Write-Host ''

  # use a scratch journal so a real one is never touched
  $realJournal = $script:Cfg.Journal
  $script:Cfg.Journal = Join-Path ([IO.Path]::GetTempPath()) ('jt_selftest_' + [guid]::NewGuid().ToString('N') + '.jsonl')
  $script:Seen = $null
  $scratch = 'Software\JiffTweaksSelfTest'
  try {
    & $check 'registry: write and read all value types' {
      RegWrite 'HKCU' $scratch 'D' 'DWord' 4294967295
      RegWrite 'HKCU' $scratch 'S' 'String' 'hello'
      RegWrite 'HKCU' $scratch 'B' 'Binary' '3000ff'
      RegWrite 'HKCU' $scratch 'M' 'MultiString' @('a', 'b')
      $d = RegRead 'HKCU' $scratch 'D'; $b = RegRead 'HKCU' $scratch 'B'; $m = RegRead 'HKCU' $scratch 'M'
      $ok = ((Format-RegValue $d.Value $d.Type) -eq '4294967295') -and ((Format-RegValue $b.Value $b.Type) -eq '3000FF') -and ((Format-RegValue $m.Value $m.Type) -eq 'a; b')
      @{ ok = $ok; msg = 'DWord 0xFFFFFFFF, Binary, MultiString' }
    }
    & $check 'registry: Set-Reg, journal and Restore round trip' {
      RegWrite 'HKCU' $scratch 'Orig' 'DWord' 7
      $script:Mode = 'apply'
      Set-Reg HKCU $scratch Orig DWord 9
      Set-Reg HKCU $scratch NewValue String 'x'
      $a = (RegRead 'HKCU' $scratch 'Orig').Value -eq 9
      $n = Restore-Journal @('reg')
      $o = RegRead 'HKCU' $scratch 'Orig'; $nv = RegRead 'HKCU' $scratch 'NewValue'
      @{ ok = ($a -and $o.Value -eq 7 -and -not $nv); msg = ('restored ' + $n + ' values, original back, new value deleted') }
    }
    & $check 'registry: subkeys of Tcpip interfaces' { $k = RegSubKeys 'HKLM' 'SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces'; @{ ok = ($k.Count -gt 0); msg = ($k.Count.ToString() + ' interfaces') } }
    & $check 'registry: network and display adapter keys' { @{ ok = $true; msg = ((Get-ClassKeys $script:NicClass '^pci\\ven_').Count.ToString() + ' NIC, ' + (Get-ClassKeys $script:GpuClass 'ven_10de').Count + ' NVIDIA, ' + (Get-ClassKeys $script:GpuClass 'ven_1002').Count + ' AMD') } }
    & $check 'services: read start type (Spooler, BITS)' { $a = SvcGetStart 'Spooler'; $b = SvcGetStart 'BITS'; @{ ok = ($null -ne $a -or $null -ne $b); msg = ('Spooler=' + $a + ' BITS=' + $b) } }
    & $check 'scheduled tasks: read a task through COM' { $t = TaskGetEnabled '\Microsoft\Windows\Application Experience\Microsoft Compatibility Appraiser'; @{ ok = ($null -ne $t); msg = ('enabled=' + $t) } }
    & $check 'power plans: list and active plan' { $g = Get-PlanGuids; @{ ok = ($g.Count -gt 0 -and (Get-ActivePlanGuid)); msg = ($g.Count.ToString() + ' plans, active ' + (Get-ActivePlanGuid)) } }
    & $check 'boot settings: bcdedit read' { $v = Get-BcdValue '{current}' 'description'; @{ ok = [bool]$v; msg = $v } }
    & $check 'network: read settings through the network cmdlets' {
      $got = @(); $miss = @()
      foreach ($id in ($script:NetMap.Keys | Sort-Object)) { $v = NetGet $id; if ($null -ne $v) { $got += $id } else { $miss += $id } }
      @{ ok = ($got.Count -gt 0); msg = ($got.Count.ToString() + ' readable' + $(if ($miss.Count) { ', not readable: ' + ($miss -join ', ') } else { '' })) }
    }
    & $check 'network: read fsutil value' { $q = Exe 'fsutil' @('behavior', 'query', 'disablelastaccess'); @{ ok = ($q.Out -match '\d'); msg = ([string]$q.Out).Trim() } }
    & $check 'memory manager: Get-MMAgent' { $m = Get-MmState; @{ ok = ($null -ne $m); msg = ('compression=' + $m.mc + ' pageCombining=' + $m.pc) } }
    & $check 'Defender: Get-MpPreference' { $m = Get-MpState; @{ ok = ($null -ne $m); msg = $(if ($m) { 'realtime disabled=' + $m.DisableRealtimeMonitoring } else { 'not available (Defender may be off or replaced)' }) } }
    & $check 'devices: power management through WMI' { $d = Get-PnpPowerDevices; @{ ok = $true; msg = ($d.Count.ToString() + ' devices') } }
    & $check 'devices: USB and device flags in the registry' { $f = Find-EnumValues @('EnableSelectiveSuspend', 'SelectiveSuspendOn', 'EnhancedPowerManagementEnabled'); @{ ok = $true; msg = ($f.Count.ToString() + ' flags found') } }
    & $check 'Store apps: list for all users' { Load-Appx -Force; @{ ok = ($script:AppxAll.Count -gt 0); msg = ($script:AppxAll.Count.ToString() + ' apps') } }
    & $check 'live numbers: CPU, memory, disk, network (WMI)' {
      if ($script:Mock) { return @{ ok = $true; msg = 'skipped in test mode' } }
      [void](Get-RealSample 0 $true); Start-Sleep -Milliseconds 1100; $s = Get-RealSample 1 $true
      @{ ok = ($null -ne $s.cpu -and $s.ramTotalMb -gt 0); msg = ('cpu ' + $s.cpu + '% ram ' + $s.ram + '% disk ' + $s.disk + '% gpu ' + $s.gpu + '% cores ' + @($s.cores).Count) }
    }
    $script:Cfg.Journal = $realJournal
    # every tweak: check its state and preview it (these only read)
    $bad = @(); $t0 = Get-Date
    foreach ($t in $script:Tweaks.Values) {
      $script:LastError = $null
      $st = Get-TweakState $t
      if ($script:LastError) { $bad += ($t.Id + ' (state): ' + $script:LastError) }
      $r = Get-TweakResult $t 'preview'
      if ($r.Error) { $bad += ($t.Id + ' (preview): ' + $r.Error) }
    }
    $secs = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
    $ok = ($bad.Count -eq 0)
    Write-Host ('  ' + $(if ($ok) { 'PASS' } else { 'FAIL' }) + '  every tweak: read its state and preview it'.PadRight(46) + ' ' + $script:Tweaks.Count + ' tweaks in ' + $secs + ' s') -ForegroundColor $(if ($ok) { 'Green' } else { 'Red' })
    [void]$res.Add($ok)
    foreach ($b in $bad) { Write-Host ('        ' + $b) -ForegroundColor Red }
    & $check 'window server: start, answer, stop' {
      $l = New-Object System.Net.HttpListener; $port = 0
      foreach ($p in 51877..51890) { $l2 = New-Object System.Net.HttpListener; $l2.Prefixes.Add('http://127.0.0.1:' + $p + '/'); try { $l2.Start(); $l = $l2; $port = $p; break } catch { $l2.Close() } }
      if ($port -eq 0) { return @{ ok = $false; msg = 'could not open a port' } }
      $job = [powershell]::Create(); [void]$job.AddScript({ param($l) $c = $l.GetContext(); $b = [Text.Encoding]::UTF8.GetBytes('ok'); $c.Response.OutputStream.Write($b, 0, 2); $c.Response.Close() }).AddArgument($l)
      $h = $job.BeginInvoke()
      $wc = New-Object System.Net.WebClient; $wc.Proxy = $null
      $txt = $wc.DownloadString('http://127.0.0.1:' + $port + '/')
      $l.Stop(); $job.Dispose()
      @{ ok = ($txt -eq 'ok'); msg = ('port ' + $port) }
    }
  } finally {
    $script:Cfg.Journal = $realJournal; $script:Seen = $null
    try { $base = Get-JTHive 'HKCU'; $base.DeleteSubKeyTree($scratch, $false); $base.Close() } catch { }
    if ($script:Mock) { [void]$script:MockState.Reg.Clear() ; }
  }
  $fails = @($res | Where-Object { -not $_ }).Count
  Write-Host ''
  if ($fails -eq 0) { Write-Host ('  All ' + $res.Count + ' checks passed.') -ForegroundColor Green } else { Write-Host ('  ' + $fails + ' of ' + $res.Count + ' checks failed. Copy this output if you ask for help.') -ForegroundColor Red }
  Write-Host ''
}
'@

# ---------------------------------------------------------------- the window (HTML, CSS and JavaScript), served only to your own browser window
$JT_UI = @{}
$JT_UI['index.html'] = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Jiff Tweaks</title>
<link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Ccircle cx='32' cy='32' r='28' fill='%2307100b' stroke='%2300e676' stroke-width='5'/%3E%3Cpath d='M38 16v22a8 8 0 0 1-8 8 8 8 0 0 1-8-6' fill='none' stroke='%2300e676' stroke-width='6' stroke-linecap='round'/%3E%3C/svg%3E">
<link rel="stylesheet" href="app.css">
</head>
<body>

<svg width="0" height="0" style="position:absolute" aria-hidden="true" focusable="false">
  <defs>
    <linearGradient id="jg" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0" stop-color="#00e676"/>
      <stop offset="1" stop-color="#1de9b6"/>
    </linearGradient>
  </defs>
</svg>

<div id="bg" aria-hidden="true"><i class="orb o1"></i><i class="orb o2"></i><i class="orb o3"></i><div class="grid"></div><div class="grain"></div></div>

<div id="splash" aria-hidden="true">
  <div class="splash-inner">
    <svg class="splash-logo" viewBox="0 0 64 64" fill="none">
      <circle class="ring" cx="32" cy="32" r="28" stroke="url(#jg)" stroke-width="3.5" stroke-linecap="round"/>
      <path class="j" d="M38 16v22a8 8 0 0 1-8 8 8 8 0 0 1-8-6" stroke="url(#jg)" stroke-width="5" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>
    <div class="splash-word" id="splashWord"></div>
    <div class="splash-bar"><i></i></div>
  </div>
</div>

<div id="app">
  <aside id="side">
    <div class="brand">
      <svg viewBox="0 0 64 64" fill="none" aria-hidden="true">
        <circle cx="32" cy="32" r="28" stroke="url(#jg)" stroke-width="4"/>
        <path d="M38 16v22a8 8 0 0 1-8 8 8 8 0 0 1-8-6" stroke="url(#jg)" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
      <div><b>Jiff Tweaks</b><span id="verLabel">free</span></div>
    </div>
    <nav id="nav" aria-label="Sections"></nav>
    <div class="side-foot">
      <div class="foot-row"><span class="live-dot" id="liveDot"></span><span id="connLabel">Connecting</span></div>
      <div class="foot-row" id="adminRow"></div>
    </div>
  </aside>

  <main id="main">
    <div id="banners"></div>
    <header id="top">
      <div><h1 id="pageTitle">Cockpit</h1><p id="pageSub"></p></div>
      <div id="topRight"></div>
    </header>
    <div id="view" tabindex="-1"></div>
  </main>
</div>

<div id="queue" role="region" aria-label="Selected tweaks"></div>

<div id="drawer" aria-hidden="true">
  <div class="scrim" data-close="drawer"></div>
  <section class="panel" role="dialog" aria-modal="true" aria-labelledby="dwTitle">
    <header class="dw-head" id="dwHead">
      <div class="ring" id="dwRing">
        <svg viewBox="0 0 100 100" aria-hidden="true"><circle class="trk" cx="50" cy="50" r="42"/><circle class="arc" id="dwArc" cx="50" cy="50" r="42" stroke="url(#jg)"/></svg>
        <div class="mid"><b id="dwPct">0</b></div>
        <div id="dwCheck"></div>
      </div>
      <div class="grow">
        <div class="dw-kicker" id="dwKicker">Running</div>
        <h2 id="dwTitle">Applying</h2>
        <div class="dw-sub" id="dwSub"></div>
      </div>
    </header>
    <div class="dw-body">
      <div class="steps" id="dwSteps"></div>
      <div class="summary" id="dwSummary"></div>
      <div class="logbox" id="dwLog" aria-live="polite"></div>
    </div>
    <footer class="dw-foot" id="dwFoot"></footer>
  </section>
</div>

<div id="modal" aria-hidden="true">
  <div class="scrim" data-close="modal"></div>
  <section class="box" role="dialog" aria-modal="true" id="modalBox"></section>
</div>

<div id="palette" aria-hidden="true">
  <div class="scrim" data-close="palette"></div>
  <div class="pbox">
    <input id="palInput" placeholder="Search tweaks, pages and actions" autocomplete="off" spellcheck="false">
    <div id="palList"></div>
  </div>
</div>

<div id="toasts" aria-live="polite"></div>
<canvas id="confetti"></canvas>

<script src="core.js"></script>
<script src="charts.js"></script>
<script src="pages.js"></script>
<script src="flows.js"></script>
<script src="app.js"></script>
</body>
</html>
'@

$JT_UI['app.css'] = @'
/* ==========================================================================
   Jiff Tweaks - interface styles
   Dark control-room look: graphite surfaces, one green accent, data colors
   only on charts and gauges (text and labels always use the text colors).
   ========================================================================== */
:root {
  --bg: #050907;
  --bg-2: #0a110e;
  --surface: rgba(17, 26, 21, .74);
  --surface-solid: #0e1612;
  --surface-hi: #15211b;
  --surface-ring: #0e1612;
  --line: rgba(255, 255, 255, .07);
  --line-hi: rgba(255, 255, 255, .15);
  --text: #eaf3ee;
  --text-2: #aebfb5;
  --muted: #7c9186;
  --accent: #00e676;
  --accent-2: #1de9b6;
  --accent-ink: #03140a;
  --ok: #3ddc84;
  --warn: #ffb74d;
  --bad: #ff6b6b;
  --info: #4fc3f7;
  --c-cpu: #00e676;
  --c-gpu: #b388ff;
  --c-ram: #4fc3f7;
  --c-disk: #ffca28;
  --c-down: #4dd0e1;
  --c-up: #ff80ab;
  --c-dpc: #ff8a65;
  --c-int: #f48fb1;
  --c-ping: #80deea;
  --r: 18px;
  --r-sm: 11px;
  --ease: cubic-bezier(.2, .8, .2, 1);
  --spring: cubic-bezier(.34, 1.56, .64, 1);
  --shadow: 0 18px 50px -20px rgba(0, 0, 0, .7), 0 2px 0 rgba(255, 255, 255, .03) inset;
  --mono: ui-monospace, "Cascadia Code", "Cascadia Mono", Consolas, "SFMono-Regular", Menlo, monospace;
  --sans: "Segoe UI Variable Display", "Segoe UI Variable", "Segoe UI", system-ui, -apple-system, Roboto, sans-serif;
  color-scheme: dark;
}
html[data-motion="off"] *, html[data-motion="off"] *::before, html[data-motion="off"] *::after { animation-duration: .001ms !important; animation-delay: 0s !important; transition-duration: .001ms !important; }
@media (prefers-reduced-motion: reduce) { *, *::before, *::after { animation-duration: .001ms !important; transition-duration: .001ms !important; } }

* { box-sizing: border-box; }
html, body { height: 100%; }
body { margin: 0; background: var(--bg); color: var(--text); font: 14px/1.5 var(--sans); overflow: hidden; -webkit-font-smoothing: antialiased; }
button, input, select { font: inherit; color: inherit; }
button { cursor: pointer; }
h1, h2, h3, h4, p { margin: 0; }
svg { display: block; }
.ic { width: 18px; height: 18px; flex: none; }
.ic.sm { width: 14px; height: 14px; }
.ic.lg { width: 24px; height: 24px; }
:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; border-radius: 6px; }
::selection { background: rgba(0, 230, 118, .3); }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-thumb { background: rgba(255, 255, 255, .12); border-radius: 10px; border: 2px solid transparent; background-clip: content-box; }
::-webkit-scrollbar-thumb:hover { background: rgba(255, 255, 255, .22); background-clip: content-box; border: 2px solid transparent; }
::-webkit-scrollbar-track { background: transparent; }

/* ---------- backdrop ---------- */
#bg { position: fixed; inset: 0; z-index: 0; overflow: hidden; background: radial-gradient(1200px 700px at 78% -10%, #0c2a1a 0%, transparent 60%), radial-gradient(900px 600px at -5% 105%, #082219 0%, transparent 55%), var(--bg); }
#bg .orb { position: absolute; border-radius: 50%; filter: blur(80px); opacity: .5; will-change: transform; }
#bg .o1 { width: 520px; height: 520px; left: 52%; top: -160px; background: radial-gradient(circle, rgba(0, 230, 118, .35), transparent 70%); animation: drift1 26s ease-in-out infinite alternate; }
#bg .o2 { width: 460px; height: 460px; left: -120px; top: 55%; background: radial-gradient(circle, rgba(29, 233, 182, .22), transparent 70%); animation: drift2 31s ease-in-out infinite alternate; }
#bg .o3 { width: 380px; height: 380px; right: -100px; bottom: -120px; background: radial-gradient(circle, rgba(179, 136, 255, .16), transparent 70%); animation: drift3 36s ease-in-out infinite alternate; }
#bg .grid { position: absolute; inset: 0; background-image: linear-gradient(rgba(255, 255, 255, .025) 1px, transparent 1px), linear-gradient(90deg, rgba(255, 255, 255, .025) 1px, transparent 1px); background-size: 46px 46px; mask-image: radial-gradient(ellipse at 50% 30%, #000 0%, transparent 72%); -webkit-mask-image: radial-gradient(ellipse at 50% 30%, #000 0%, transparent 72%); }
#bg .grain { position: absolute; inset: -50%; opacity: .05; background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='160' height='160'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='.9' numOctaves='2' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)'/%3E%3C/svg%3E"); }
@keyframes drift1 { to { transform: translate(-140px, 120px) scale(1.15); } }
@keyframes drift2 { to { transform: translate(160px, -90px) scale(1.2); } }
@keyframes drift3 { to { transform: translate(-120px, -80px) scale(.9); } }

/* ---------- splash ---------- */
#splash { position: fixed; inset: 0; z-index: 200; display: grid; place-items: center; background: var(--bg); transition: opacity .6s var(--ease), visibility .6s; }
#splash.gone { opacity: 0; visibility: hidden; pointer-events: none; }
.splash-inner { display: grid; justify-items: center; gap: 22px; }
.splash-logo { width: 96px; height: 96px; }
.splash-logo .ring { stroke-dasharray: 176; stroke-dashoffset: 176; animation: draw 1.1s var(--ease) .1s forwards; }
.splash-logo .j { stroke-dasharray: 80; stroke-dashoffset: 80; animation: draw .8s var(--ease) .7s forwards; }
@keyframes draw { to { stroke-dashoffset: 0; } }
.splash-word { display: flex; gap: 8px; font-weight: 700; letter-spacing: .38em; font-size: 15px; color: var(--text-2); }
.splash-word span { opacity: 0; transform: translateY(8px); animation: rise .5s var(--ease) forwards; }
.splash-bar { width: 180px; height: 3px; border-radius: 3px; background: var(--line); overflow: hidden; }
.splash-bar i { display: block; height: 100%; width: 0; background: linear-gradient(90deg, var(--accent), var(--accent-2)); animation: load 1.4s var(--ease) .2s forwards; }
@keyframes load { to { width: 100%; } }
@keyframes rise { to { opacity: 1; transform: none; } }

/* ---------- shell ---------- */
#app { position: relative; z-index: 1; display: grid; grid-template-columns: 236px minmax(0, 1fr); height: 100vh; }
#side { display: flex; flex-direction: column; padding: 18px 12px 14px; border-right: 1px solid var(--line); background: linear-gradient(180deg, rgba(12, 19, 15, .85), rgba(8, 13, 10, .85)); backdrop-filter: blur(22px); min-height: 0; }
.brand { display: flex; align-items: center; gap: 11px; padding: 4px 8px 18px; }
.brand svg { width: 38px; height: 38px; }
.brand b { display: block; font-size: 16px; letter-spacing: .01em; }
.brand span { display: block; font-size: 11px; color: var(--muted); letter-spacing: .06em; text-transform: uppercase; }
#nav { position: relative; display: flex; flex-direction: column; gap: 2px; overflow-y: auto; flex: 1; padding-right: 2px; }
.nav-sep { margin: 12px 10px 4px; font-size: 10.5px; letter-spacing: .14em; text-transform: uppercase; color: var(--muted); }
.nav-item { position: relative; display: flex; align-items: center; gap: 11px; width: 100%; padding: 9px 11px; border: 0; border-radius: 11px; background: transparent; color: var(--text-2); text-align: left; transition: background .2s, color .2s, transform .2s var(--ease); }
.nav-item:hover { background: rgba(255, 255, 255, .05); color: var(--text); transform: translateX(2px); }
.nav-item.on { color: var(--text); background: linear-gradient(90deg, rgba(0, 230, 118, .16), rgba(0, 230, 118, .03)); box-shadow: inset 0 0 0 1px rgba(0, 230, 118, .22); }
.nav-item.on::before { content: ""; position: absolute; left: -12px; top: 9px; bottom: 9px; width: 3px; border-radius: 3px; background: var(--accent); box-shadow: 0 0 12px var(--accent); }
.nav-item .badge { margin-left: auto; font-size: 11px; min-width: 20px; padding: 1px 7px; border-radius: 20px; text-align: center; background: rgba(255, 255, 255, .07); color: var(--text-2); font-variant-numeric: tabular-nums; }
.nav-item .badge.ok { background: rgba(61, 220, 132, .16); color: var(--text); }
.side-foot { padding: 12px 8px 2px; display: grid; gap: 9px; border-top: 1px solid var(--line); margin-top: 8px; }
.foot-row { display: flex; align-items: center; gap: 8px; font-size: 12px; color: var(--muted); }
.live-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--ok); box-shadow: 0 0 0 0 rgba(61, 220, 132, .6); animation: pulse 2s infinite; }
.live-dot.off { background: var(--bad); animation: none; }
@keyframes pulse { 70% { box-shadow: 0 0 0 8px rgba(61, 220, 132, 0); } 100% { box-shadow: 0 0 0 0 rgba(61, 220, 132, 0); } }

#main { position: relative; min-width: 0; min-height: 0; display: flex; flex-direction: column; }
#banners { display: grid; gap: 8px; padding: 0 28px; }
.banner { display: flex; align-items: center; gap: 12px; margin-top: 14px; padding: 11px 14px; border-radius: 12px; border: 1px solid var(--line-hi); background: var(--surface); font-size: 13px; }
.banner.warn { border-color: rgba(255, 183, 77, .35); background: rgba(255, 183, 77, .08); }
.banner.bad { border-color: rgba(255, 107, 107, .4); background: rgba(255, 107, 107, .08); }
.banner .grow { flex: 1; }
#top { display: flex; align-items: flex-end; justify-content: space-between; gap: 20px; padding: 24px 28px 12px; }
#top h1 { font-size: 28px; letter-spacing: -.01em; font-weight: 700; }
#top p { color: var(--muted); margin-top: 3px; }
#topRight { display: flex; gap: 10px; align-items: center; }
#view { flex: 1; min-height: 0; overflow-y: auto; padding: 8px 28px 120px; scroll-behavior: smooth; }
.page-in { animation: pageIn .45s var(--ease) both; }
@keyframes pageIn { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: none; } }
.stagger > * { animation: pageIn .5s var(--ease) both; }
.stagger > *:nth-child(1) { animation-delay: .02s; } .stagger > *:nth-child(2) { animation-delay: .07s; } .stagger > *:nth-child(3) { animation-delay: .12s; } .stagger > *:nth-child(4) { animation-delay: .17s; }
.stagger > *:nth-child(5) { animation-delay: .22s; } .stagger > *:nth-child(6) { animation-delay: .27s; } .stagger > *:nth-child(7) { animation-delay: .32s; } .stagger > *:nth-child(8) { animation-delay: .37s; } .stagger > *:nth-child(n+9) { animation-delay: .42s; }

/* ---------- buttons, chips ---------- */
.btn { display: inline-flex; align-items: center; justify-content: center; gap: 8px; padding: 9px 16px; border-radius: 11px; border: 1px solid var(--line-hi); background: rgba(255, 255, 255, .05); color: var(--text); font-weight: 600; font-size: 13px; transition: transform .2s var(--spring), background .2s, border-color .2s, box-shadow .25s; white-space: nowrap; }
.btn:hover { background: rgba(255, 255, 255, .09); border-color: rgba(255, 255, 255, .26); transform: translateY(-1px); }
.btn:active { transform: translateY(0) scale(.97); }
.btn.primary { color: var(--accent-ink); border: 0; background: linear-gradient(135deg, var(--accent), var(--accent-2)); box-shadow: 0 8px 26px -10px rgba(0, 230, 118, .75); }
.btn.primary:hover { box-shadow: 0 12px 34px -8px rgba(0, 230, 118, .9); filter: brightness(1.06); }
.btn.danger { border-color: rgba(255, 107, 107, .5); background: rgba(255, 107, 107, .1); }
.btn.danger:hover { background: rgba(255, 107, 107, .2); }
.btn.ghost { background: transparent; border-color: transparent; color: var(--text-2); }
.btn.ghost:hover { background: rgba(255, 255, 255, .06); color: var(--text); }
.btn.sm { padding: 6px 11px; font-size: 12px; border-radius: 9px; }
.btn.icon { padding: 6px 8px; }
.btn[disabled] { opacity: .45; pointer-events: none; }
.kbd { font: 11px var(--mono); padding: 1px 6px; border-radius: 6px; border: 1px solid var(--line-hi); color: var(--text-2); background: rgba(255, 255, 255, .04); }

.chip { display: inline-flex; align-items: center; gap: 6px; padding: 3px 9px 3px 7px; border-radius: 20px; font-size: 11.5px; font-weight: 600; letter-spacing: .02em; border: 1px solid var(--line-hi); background: rgba(255, 255, 255, .04); color: var(--text-2); white-space: nowrap; }
.chip .ic { width: 13px; height: 13px; }
.chip.on { border-color: rgba(61, 220, 132, .45); background: rgba(61, 220, 132, .12); color: var(--text); }
.chip.on .ic { color: var(--ok); }
.chip.partial { border-color: rgba(255, 183, 77, .45); background: rgba(255, 183, 77, .1); color: var(--text); }
.chip.partial .ic { color: var(--warn); }
.chip.off .ic { color: var(--muted); }
.chip.r-safe .ic { color: var(--ok); }
.chip.r-moderate .ic { color: var(--info); }
.chip.r-aggressive { border-color: rgba(255, 183, 77, .4); } .chip.r-aggressive .ic { color: var(--warn); }
.chip.r-security, .chip.r-danger { border-color: rgba(255, 107, 107, .5); background: rgba(255, 107, 107, .1); color: var(--text); } .chip.r-security .ic, .chip.r-danger .ic { color: var(--bad); }
.chip.reboot .ic { color: var(--warn); }
.tag { font-size: 11px; padding: 2px 8px; border-radius: 6px; background: rgba(255, 255, 255, .05); color: var(--text-2); }
.skel { position: relative; overflow: hidden; background: rgba(255, 255, 255, .05); border-radius: 20px; color: transparent; min-width: 64px; }
.skel::after { content: ""; position: absolute; inset: 0; transform: translateX(-100%); background: linear-gradient(90deg, transparent, rgba(255, 255, 255, .12), transparent); animation: shimmer 1.4s infinite; }
@keyframes shimmer { to { transform: translateX(100%); } }

/* ---------- cards ---------- */
.card { position: relative; border-radius: var(--r); background: var(--surface); border: 1px solid var(--line); box-shadow: var(--shadow); backdrop-filter: blur(16px) saturate(1.2); padding: 18px; overflow: hidden; }
.card.glow::before { content: ""; position: absolute; inset: -1px; border-radius: inherit; padding: 1px; background: linear-gradient(135deg, rgba(0, 230, 118, .5), transparent 40%, transparent 60%, rgba(29, 233, 182, .3)); -webkit-mask: linear-gradient(#000 0 0) content-box, linear-gradient(#000 0 0); -webkit-mask-composite: xor; mask-composite: exclude; pointer-events: none; opacity: .7; }
.card h3 { font-size: 12px; text-transform: uppercase; letter-spacing: .12em; color: var(--muted); font-weight: 600; display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
.card h3 .grow { flex: 1; }
.lift { transition: transform .35s var(--ease), border-color .25s, box-shadow .35s; }
.lift:hover { transform: translateY(-3px); border-color: var(--line-hi); box-shadow: 0 26px 60px -24px rgba(0, 0, 0, .85), 0 0 0 1px rgba(0, 230, 118, .08); }
.sec { margin: 26px 0 12px; display: flex; align-items: baseline; gap: 12px; }
.sec h2 { font-size: 15px; letter-spacing: .01em; }
.sec span { color: var(--muted); font-size: 12.5px; }
.sec .grow { flex: 1; }
.muted { color: var(--muted); }
.mono { font-family: var(--mono); }
.row { display: flex; align-items: center; gap: 10px; }
.grow { flex: 1; min-width: 0; }

/* ---------- cockpit ---------- */
.g { display: grid; gap: 16px; }
.g4 { grid-template-columns: repeat(4, minmax(0, 1fr)); }
.g3 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
.g2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
.g-main { grid-template-columns: minmax(0, 2fr) minmax(0, 1fr); }
@media (max-width: 1180px) { .g4 { grid-template-columns: repeat(2, minmax(0, 1fr)); } .g3 { grid-template-columns: 1fr; } .g-main { grid-template-columns: 1fr; } }
.tile { padding: 16px 16px 12px; }
.tile .t-head { display: flex; align-items: center; gap: 8px; font-size: 11.5px; letter-spacing: .12em; text-transform: uppercase; color: var(--muted); font-weight: 600; }
.tile .t-head .sw { width: 9px; height: 9px; border-radius: 3px; }
.tile .t-body { display: flex; align-items: center; justify-content: space-between; gap: 6px; margin-top: 6px; }
.big { font-size: 52px; line-height: 1; font-weight: 700; letter-spacing: -.03em; font-variant-numeric: tabular-nums; }
.big small { font-size: 20px; font-weight: 600; color: var(--text-2); margin-left: 3px; letter-spacing: 0; }
.tile .sub { margin-top: 6px; font-size: 12px; color: var(--text-2); min-height: 18px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.tile svg.spark { width: 100%; height: 40px; margin-top: 8px; overflow: visible; }
.ring { position: relative; width: 84px; height: 84px; flex: none; }
.ring > svg { width: 100%; height: 100%; transform: rotate(-90deg); }
.ring .trk { fill: none; stroke: rgba(255, 255, 255, .07); stroke-width: 8; }
.ring .arc { fill: none; stroke-width: 8; stroke-linecap: round; transition: stroke-dashoffset .9s var(--ease); }
.ring.lg { width: 150px; height: 150px; }
.ring.lg .trk, .ring.lg .arc { stroke-width: 9; }
.ring .mid { position: absolute; inset: 0; display: grid; place-items: center; text-align: center; }
.ring .mid b { font-size: 36px; letter-spacing: -.03em; line-height: 1; font-variant-numeric: tabular-nums; }
.ring .mid small { display: block; font-size: 11px; color: var(--muted); letter-spacing: .1em; text-transform: uppercase; margin-top: 3px; }

.profile { display: flex; flex-direction: column; gap: 12px; padding: 18px; cursor: default; }
.profile .p-head { display: flex; align-items: center; gap: 12px; }
.p-ico { width: 44px; height: 44px; border-radius: 13px; display: grid; place-items: center; background: linear-gradient(135deg, rgba(0, 230, 118, .22), rgba(29, 233, 182, .06)); border: 1px solid rgba(0, 230, 118, .28); color: var(--accent); }
.p-ico .ic { width: 22px; height: 22px; }
.profile h4 { font-size: 17px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.profile .grow { min-width: 0; }
.profile p { color: var(--text-2); font-size: 13px; min-height: 40px; }
.profile .tags { display: flex; flex-wrap: wrap; gap: 5px; }
.profile .p-foot { display: flex; align-items: center; gap: 8px; margin-top: auto; }
.meter { height: 5px; border-radius: 5px; background: rgba(255, 255, 255, .07); overflow: hidden; }
.meter i { display: block; height: 100%; width: 0; border-radius: 5px; background: linear-gradient(90deg, var(--accent), var(--accent-2)); transition: width .9s var(--ease); }

.kv { display: grid; grid-template-columns: auto 1fr; gap: 6px 14px; font-size: 13px; }
.kv dt { color: var(--muted); } .kv dd { margin: 0; color: var(--text); text-align: right; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.score { display: flex; align-items: center; gap: 18px; }
.score .ring .mid b { font-size: 38px; }

.sug { display: flex; align-items: center; gap: 12px; padding: 10px 0; border-top: 1px solid var(--line); }
.sug:first-of-type { border-top: 0; }
.sug .s-t { font-weight: 600; font-size: 13.5px; }
.sug .s-d { font-size: 12px; color: var(--muted); }
.delta-row { display: flex; align-items: baseline; justify-content: space-between; padding: 8px 0; border-top: 1px solid var(--line); }
.delta-row:first-of-type { border-top: 0; }
.delta-row .d-l { color: var(--text-2); font-size: 13px; }
.delta-row .d-v { font-variant-numeric: tabular-nums; font-weight: 600; white-space: nowrap; display: inline-flex; align-items: center; gap: 6px; }
.delta-row .d-c { margin-left: 4px; font-size: 12px; padding: 1px 8px 1px 5px; border-radius: 6px; background: rgba(61, 220, 132, .14); display: inline-flex; align-items: center; gap: 3px; }
.delta-row .d-c.same { background: rgba(255, 255, 255, .06); padding-left: 8px; color: var(--text-2); font-weight: 500; }
.delta-row .d-c.up { background: rgba(255, 255, 255, .07); }

/* ---------- tweak cards ---------- */
.featured { display: grid; grid-template-columns: auto 1fr auto; gap: 20px; align-items: center; padding: 22px; margin-bottom: 18px; }
.featured .f-ico { width: 62px; height: 62px; border-radius: 18px; display: grid; place-items: center; color: var(--accent); background: linear-gradient(135deg, rgba(0, 230, 118, .25), rgba(29, 233, 182, .05)); border: 1px solid rgba(0, 230, 118, .3); }
.featured .f-ico .ic { width: 30px; height: 30px; }
.featured h2 { font-size: 20px; margin-bottom: 4px; }
.featured p { color: var(--text-2); max-width: 70ch; }
.featured .f-meta { display: flex; gap: 8px; margin-top: 10px; flex-wrap: wrap; }
.featured .f-act { display: flex; gap: 8px; align-items: center; }
.cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(330px, 1fr)); gap: 14px; }
.tw { display: flex; flex-direction: column; gap: 10px; padding: 16px; cursor: pointer; }
.tw.sel { border-color: rgba(0, 230, 118, .55); box-shadow: 0 0 0 1px rgba(0, 230, 118, .3), 0 18px 50px -22px rgba(0, 230, 118, .45); }
.tw.hide { display: none; }
.tw .t-top { display: flex; align-items: flex-start; gap: 12px; }
.tw .t-ico { width: 40px; height: 40px; border-radius: 12px; display: grid; place-items: center; background: rgba(255, 255, 255, .05); border: 1px solid var(--line); color: var(--text-2); transition: all .3s var(--ease); flex: none; }
.tw:hover .t-ico, .tw.sel .t-ico { color: var(--accent); background: rgba(0, 230, 118, .1); border-color: rgba(0, 230, 118, .3); }
.tw h4 { font-size: 15px; line-height: 1.25; }
.tw .t-short { color: var(--text-2); font-size: 13px; }
.tw .t-chips { display: flex; flex-wrap: wrap; gap: 6px; }
.tw .t-tags { display: flex; flex-wrap: wrap; gap: 5px; }
.tw .t-more { display: none; padding-top: 6px; border-top: 1px dashed var(--line-hi); }
.tw.open .t-more { display: block; animation: pageIn .3s var(--ease); }
.tw .t-more ul { margin: 0; padding-left: 18px; color: var(--text-2); font-size: 12.5px; display: grid; gap: 3px; }
.tw .t-more .note { margin-top: 8px; padding: 8px 10px; border-radius: 9px; background: rgba(255, 183, 77, .08); border: 1px solid rgba(255, 183, 77, .25); font-size: 12.5px; }
.tw .t-foot { display: flex; align-items: center; gap: 6px; margin-top: auto; padding-top: 4px; }
.tw .t-foot .sp { flex: 1; }
.chk { width: 22px; height: 22px; border-radius: 7px; border: 1.5px solid var(--line-hi); display: grid; place-items: center; transition: all .25s var(--spring); background: rgba(255, 255, 255, .03); flex: none; color: transparent; }
.chk .ic { width: 14px; height: 14px; }
.tw.sel .chk { background: var(--accent); border-color: var(--accent); color: var(--accent-ink); transform: scale(1.08); }
.tw .t-name { display: flex; flex-direction: column; min-width: 0; flex: 1; }
.filters { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; margin-bottom: 14px; }
.seg { display: inline-flex; padding: 3px; border-radius: 11px; background: rgba(255, 255, 255, .05); border: 1px solid var(--line); }
.seg button { border: 0; background: transparent; padding: 6px 12px; border-radius: 8px; color: var(--text-2); font-size: 12.5px; font-weight: 600; transition: all .2s; }
.seg button.on { background: rgba(255, 255, 255, .12); color: var(--text); box-shadow: 0 2px 8px rgba(0, 0, 0, .3); }
.inp { padding: 8px 12px; border-radius: 10px; border: 1px solid var(--line-hi); background: rgba(0, 0, 0, .25); color: var(--text); min-width: 0; }
.inp:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 3px rgba(0, 230, 118, .18); }

/* ---------- queue bar ---------- */
#queue { position: fixed; left: calc(236px + (100vw - 236px) / 2); bottom: 22px; transform: translate(-50%, 220px); visibility: hidden; z-index: 40; display: flex; align-items: center; gap: 14px; padding: 10px 12px 10px 18px; border-radius: 18px; background: rgba(14, 22, 18, .92); border: 1px solid rgba(0, 230, 118, .35); box-shadow: 0 24px 70px -20px rgba(0, 0, 0, .9), 0 0 40px -10px rgba(0, 230, 118, .35); backdrop-filter: blur(20px); transition: transform .5s var(--spring), visibility 0s .5s; max-width: calc(100vw - 300px); }
#queue.show { transform: translate(-50%, 0); visibility: visible; transition: transform .5s var(--spring), visibility 0s; }
#queue .q-n { font-size: 26px; font-weight: 700; letter-spacing: -.03em; font-variant-numeric: tabular-nums; }
#queue .q-t { font-size: 12px; color: var(--text-2); line-height: 1.3; max-width: 360px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#queue .q-t b { color: var(--text); display: block; font-size: 13px; }

/* ---------- drawer / modals ---------- */
#drawer, #modal { position: fixed; inset: 0; z-index: 80; visibility: hidden; }
#drawer.show, #modal.show { visibility: visible; }
.scrim { position: absolute; inset: 0; background: rgba(2, 5, 3, .66); backdrop-filter: blur(6px); opacity: 0; transition: opacity .35s; }
.show > .scrim { opacity: 1; }
#drawer .panel { position: absolute; top: 0; right: 0; bottom: 0; width: min(560px, 100vw); background: linear-gradient(180deg, #0e1713, #09100d); border-left: 1px solid var(--line-hi); transform: translateX(105%); transition: transform .55s var(--ease); display: flex; flex-direction: column; box-shadow: -40px 0 80px -30px #000; }
#drawer.show .panel { transform: none; }
.dw-head { display: flex; align-items: center; gap: 18px; padding: 24px 24px 16px; }
.dw-head .ring { width: 92px; height: 92px; }
.dw-head .ring .mid b { font-size: 26px; }
.dw-kicker { font-size: 11px; text-transform: uppercase; letter-spacing: .14em; color: var(--muted); }
.dw-head h2 { font-size: 20px; margin: 2px 0 4px; }
.dw-sub { color: var(--text-2); font-size: 13px; }
.dw-body { flex: 1; min-height: 0; display: flex; flex-direction: column; gap: 14px; padding: 0 24px; overflow: hidden; }
.steps { display: grid; gap: 6px; max-height: 38%; overflow-y: auto; padding-right: 4px; }
.step { display: flex; align-items: center; gap: 11px; padding: 8px 11px; border-radius: 11px; background: rgba(255, 255, 255, .03); border: 1px solid var(--line); font-size: 13px; transition: all .3s var(--ease); }
.step .s-i { width: 22px; height: 22px; border-radius: 50%; display: grid; place-items: center; border: 1.5px solid var(--line-hi); color: var(--muted); flex: none; }
.step .s-i .ic { width: 13px; height: 13px; }
.step .s-c { margin-left: auto; font-size: 11.5px; color: var(--muted); font-variant-numeric: tabular-nums; }
.step.sub { margin-left: 24px; padding: 5px 10px; font-size: 12.5px; background: transparent; border-style: dashed; }
.step.sub .s-i { width: 18px; height: 18px; }
.step.sub.done { opacity: .75; }
.step.run { border-color: rgba(0, 230, 118, .45); background: rgba(0, 230, 118, .07); }
.step.run .s-i { border-color: var(--accent); border-top-color: transparent; animation: spin .8s linear infinite; }
.step.done .s-i { background: var(--ok); border-color: var(--ok); color: var(--accent-ink); }
.step.err .s-i { background: var(--bad); border-color: var(--bad); color: #1a0505; }
.step.skip { opacity: .55; }
@keyframes spin { to { transform: rotate(360deg); } }
.logbox { flex: 1; min-height: 120px; overflow-y: auto; border-radius: 13px; background: #060b08; border: 1px solid var(--line); padding: 10px 12px; font: 12px/1.65 var(--mono); }
.ll { white-space: pre-wrap; word-break: break-word; display: flex; gap: 8px; }
.ll .tm { color: var(--muted); flex: none; }
.ll .tg { flex: none; width: 52px; font-weight: 600; letter-spacing: .02em; color: var(--text-2); }
.ll.t-OK .tg::before { content: "ok"; } .ll.t-OK .tg { font-size: 0; } .ll.t-OK .tg::before { font-size: 12px; color: var(--ok); }
.ll.t-FAIL .tg { color: var(--bad); } .ll.t-SKIP .tg { color: var(--muted); } .ll.t-NOTE .tg { color: var(--warn); }
.ll.t-TWEAK .tg, .ll.t-STEP .tg, .ll.t-SEQ .tg { color: var(--accent); }
.ll.t-FAIL .m { color: #ffb4b4; } .ll.t-TWEAK .m { color: var(--text); font-weight: 600; }
.dw-foot { display: flex; align-items: center; gap: 10px; padding: 16px 24px 22px; }
.summary { display: none; padding: 12px 14px; border-radius: 13px; border: 1px solid var(--line-hi); background: rgba(255, 255, 255, .04); }
.summary.show { display: block; animation: pageIn .4s var(--ease); }
.summary .s-row { display: flex; gap: 22px; flex-wrap: wrap; }
.summary .s-n { font-size: 28px; font-weight: 700; letter-spacing: -.02em; font-variant-numeric: tabular-nums; line-height: 1.1; }
.summary .s-l { font-size: 11px; letter-spacing: .1em; text-transform: uppercase; color: var(--muted); }
#dwCheck { position: absolute; inset: 0; display: grid; place-items: center; opacity: 0; transform: scale(.4); transition: all .5s var(--spring); color: var(--accent); }
#dwCheck .ic { width: 44px; height: 44px; stroke-width: 2.4; }
.finished #dwCheck { opacity: 1; transform: none; }
.finished .mid { opacity: 0; }
.dw-head.bad #dwCheck { color: var(--bad); }

#modal .box { position: absolute; left: 50%; top: 50%; transform: translate(-50%, -46%) scale(.96); width: min(980px, calc(100vw - 80px)); max-height: calc(100vh - 80px); display: flex; flex-direction: column; border-radius: 22px; background: linear-gradient(180deg, #0f1814, #09100d); border: 1px solid var(--line-hi); box-shadow: 0 60px 120px -30px #000; opacity: 0; transition: opacity .45s var(--ease), transform .45s var(--ease); overflow: hidden; }
#modal.show .box { opacity: 1; transform: translate(-50%, -50%) scale(1); }
.mh { display: flex; align-items: center; gap: 14px; padding: 18px 22px; border-bottom: 1px solid var(--line); }
.mh h2 { font-size: 18px; }
.mh .sub { color: var(--muted); font-size: 12.5px; }
.mb { flex: 1; min-height: 0; overflow: auto; padding: 18px 22px; }
.mf { display: flex; align-items: center; gap: 10px; padding: 14px 22px; border-top: 1px solid var(--line); }
.tabs { display: flex; gap: 4px; padding: 0 22px; border-bottom: 1px solid var(--line); }
.tabs button { border: 0; background: transparent; padding: 11px 14px; color: var(--muted); font-weight: 600; font-size: 13px; position: relative; }
.tabs button.on { color: var(--text); }
.tabs button.on::after { content: ""; position: absolute; left: 10px; right: 10px; bottom: -1px; height: 2px; background: var(--accent); border-radius: 2px; box-shadow: 0 0 10px var(--accent); }

/* code */
.code { position: relative; border-radius: 14px; background: #050a07; border: 1px solid var(--line); overflow: auto; }
.code pre { margin: 0; padding: 14px 16px 14px 0; font: 12.5px/1.7 var(--mono); counter-reset: ln; min-width: max-content; }
.code .l { display: block; padding-left: 58px; position: relative; white-space: pre; }
.code .l::before { counter-increment: ln; content: counter(ln); position: absolute; left: 0; width: 44px; text-align: right; color: #4c5f55; user-select: none; }
.code .l:hover { background: rgba(255, 255, 255, .03); }
.tk-c { color: #6b8576; font-style: italic; } .tk-s { color: #e6c07b; } .tk-v { color: #9ad1ff; } .tk-n { color: #ffab70; } .tk-k { color: #c792ea; } .tk-p { color: #7fdbca; } .tk-f { color: #4cffa0; font-weight: 600; } .tk-m { color: #9fe1c2; }
.copy { position: absolute; right: 10px; top: 10px; z-index: 2; }
.verb { padding: 12px 14px; border-radius: 13px; border: 1px solid var(--line); background: rgba(255, 255, 255, .03); margin-bottom: 10px; }
.verb summary { cursor: pointer; display: flex; gap: 12px; align-items: baseline; list-style: none; }
.verb summary::-webkit-details-marker { display: none; }
.verb summary b { font: 600 13px var(--mono); color: #4cffa0; }
.verb summary span { color: var(--text-2); font-size: 12.5px; }
.verb .code { margin-top: 10px; }

/* tables */
.tbl { width: 100%; border-collapse: collapse; font-size: 12.5px; }
.tbl th { text-align: left; padding: 8px 10px; color: var(--muted); font-weight: 600; font-size: 11px; letter-spacing: .1em; text-transform: uppercase; position: sticky; top: 0; background: #0b130f; z-index: 1; border-bottom: 1px solid var(--line-hi); }
.tbl td { padding: 7px 10px; border-bottom: 1px solid var(--line); vertical-align: top; word-break: break-word; }
.tbl tr:hover td { background: rgba(255, 255, 255, .025); }
.tbl .k { font: 11.5px var(--mono); color: var(--text-2); }
.tbl .old { color: var(--text-2); text-decoration: line-through; text-decoration-color: rgba(255, 107, 107, .7); }
.tbl .old.na { text-decoration: none; }
.tbl .new { color: var(--text); font-weight: 600; }
.pv-item { margin-bottom: 14px; border: 1px solid var(--line); border-radius: 14px; overflow: hidden; }
.pv-item > summary { display: flex; align-items: center; gap: 12px; padding: 12px 14px; cursor: pointer; list-style: none; background: rgba(255, 255, 255, .03); }
.pv-item > summary::-webkit-details-marker { display: none; }
.pv-item .pv-n { margin-left: auto; font-variant-numeric: tabular-nums; color: var(--text-2); }
.pv-item .chev { transition: transform .25s var(--ease); }
.pv-item[open] .chev { transform: rotate(90deg); }

/* ---------- live page ---------- */
.chart { position: relative; }
.chart svg { width: 100%; display: block; overflow: visible; }
.chart .gl { stroke: rgba(255, 255, 255, .07); stroke-width: 1; }
.chart .ax { fill: var(--muted); font-size: 10.5px; }
.chart .ln { fill: none; stroke-width: 2; stroke-linejoin: round; stroke-linecap: round; }
.chart .hov { stroke: rgba(255, 255, 255, .35); stroke-width: 1; stroke-dasharray: 3 3; }
.tip { position: absolute; pointer-events: none; z-index: 5; min-width: 130px; padding: 8px 10px; border-radius: 10px; background: rgba(8, 13, 10, .96); border: 1px solid var(--line-hi); box-shadow: 0 10px 30px rgba(0, 0, 0, .6); font-size: 12px; opacity: 0; transition: opacity .12s; }
.tip.show { opacity: 1; }
.tip .tt { color: var(--muted); font-size: 11px; margin-bottom: 4px; }
.tip .tr { display: flex; align-items: center; gap: 7px; justify-content: space-between; }
.tip .tr i { width: 8px; height: 8px; border-radius: 3px; display: inline-block; }
.tip .tr span { display: flex; align-items: center; gap: 7px; color: var(--text-2); }
.tip .tr b { font-variant-numeric: tabular-nums; }
.legend { display: flex; flex-wrap: wrap; gap: 16px; margin-bottom: 8px; font-size: 12.5px; color: var(--text-2); }
.legend span { display: inline-flex; align-items: center; gap: 7px; }
.legend i { width: 14px; height: 4px; border-radius: 3px; display: inline-block; }
.legend b { color: var(--text); font-variant-numeric: tabular-nums; margin-left: 2px; }
.cores { display: grid; grid-template-columns: repeat(auto-fill, minmax(54px, 1fr)); gap: 8px; }
.core { position: relative; height: 92px; border-radius: 10px; background: rgba(255, 255, 255, .04); border: 1px solid var(--line); overflow: hidden; }
.core i { position: absolute; left: 0; right: 0; bottom: 0; background: linear-gradient(180deg, var(--c-cpu), rgba(0, 230, 118, .35)); opacity: .85; transition: height .8s var(--ease); }
.core b { position: absolute; left: 0; right: 0; top: 7px; text-align: center; font-size: 12px; font-variant-numeric: tabular-nums; }
.core small { position: absolute; left: 0; right: 0; bottom: 5px; text-align: center; font-size: 10px; color: var(--text); text-shadow: 0 1px 2px #000; }
.explain { font-size: 12.5px; color: var(--muted); margin-top: 8px; max-width: 60ch; }
.bar { height: 6px; border-radius: 6px; background: rgba(255, 255, 255, .07); overflow: hidden; min-width: 70px; }
.bar i { display: block; height: 100%; border-radius: 6px; background: var(--c-cpu); transition: width .8s var(--ease); }

/* ---------- switch, toasts, palette ---------- */
.sw-t { position: relative; width: 44px; height: 26px; border-radius: 20px; background: rgba(255, 255, 255, .12); border: 0; transition: background .25s; flex: none; padding: 0; }
.sw-t::after { content: ""; position: absolute; left: 3px; top: 3px; width: 20px; height: 20px; border-radius: 50%; background: #fff; transition: transform .35s var(--spring); box-shadow: 0 2px 6px rgba(0, 0, 0, .5); }
.sw-t.on { background: var(--accent); }
.sw-t.on::after { transform: translateX(18px); }
.set-row { display: flex; align-items: center; gap: 16px; padding: 14px 0; border-top: 1px solid var(--line); }
.set-row:first-of-type { border-top: 0; }
.set-row .s-t { font-weight: 600; }
.set-row .s-d { color: var(--muted); font-size: 12.5px; max-width: 62ch; }
#toasts { position: fixed; right: 20px; bottom: 20px; z-index: 120; display: grid; gap: 8px; justify-items: end; }
.toast { padding: 10px 14px; border-radius: 12px; background: rgba(14, 22, 18, .96); border: 1px solid var(--line-hi); box-shadow: 0 16px 40px -10px #000; display: flex; align-items: center; gap: 10px; font-size: 13px; animation: toastIn .5s var(--spring); }
.toast.out { animation: toastOut .3s forwards; }
@keyframes toastIn { from { opacity: 0; transform: translateY(14px) scale(.94); } }
@keyframes toastOut { to { opacity: 0; transform: translateX(30px); } }
#palette { position: fixed; inset: 0; z-index: 110; visibility: hidden; }
#palette.show { visibility: visible; }
#palette .pbox { position: absolute; left: 50%; top: 14vh; transform: translate(-50%, -10px) scale(.97); width: min(620px, calc(100vw - 40px)); border-radius: 18px; background: #0d1511; border: 1px solid var(--line-hi); box-shadow: 0 50px 100px -20px #000; opacity: 0; transition: opacity .3s var(--ease), transform .3s var(--ease); overflow: hidden; }
#palette.show .pbox { opacity: 1; transform: translate(-50%, 0) scale(1); }
#palette input { width: 100%; padding: 18px 20px; border: 0; background: transparent; font-size: 16px; outline: none; border-bottom: 1px solid var(--line); }
#palList { max-height: 50vh; overflow: auto; padding: 8px; }
.pal-i { display: flex; align-items: center; gap: 12px; padding: 10px 12px; border-radius: 10px; cursor: pointer; }
.pal-i.on { background: rgba(0, 230, 118, .12); }
.pal-i .p-t { font-weight: 600; } .pal-i .p-s { color: var(--muted); font-size: 12px; margin-left: auto; }
#confetti { position: fixed; inset: 0; z-index: 150; pointer-events: none; }
.empty { display: grid; place-items: center; gap: 10px; padding: 50px 20px; color: var(--muted); text-align: center; }
.empty .ic { width: 40px; height: 40px; opacity: .5; }

.log-layout { display: grid; grid-template-columns: 260px minmax(0, 1fr); gap: 16px; height: calc(100vh - 220px); min-height: 360px; }
.log-list { overflow: auto; display: grid; gap: 6px; align-content: start; }
.log-item { text-align: left; padding: 10px 12px; border-radius: 11px; border: 1px solid var(--line); background: rgba(255, 255, 255, .03); }
.log-item.on { border-color: rgba(0, 230, 118, .45); background: rgba(0, 230, 118, .07); }
.log-item small { display: block; color: var(--muted); }
.diag { font: 12px/1.7 var(--mono); color: var(--text-2); }
.diag b { color: var(--text); }
'@

$JT_UI['core.js'] = @'
/* Jiff Tweaks UI - core: helpers, icons, API, state */
'use strict';

const $ = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const sleep = ms => new Promise(r => setTimeout(r, ms));
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const nf = (v, d) => (v == null || isNaN(v)) ? '-' : Number(v).toLocaleString('en-US', { minimumFractionDigits: d || 0, maximumFractionDigits: d || 0 });
const gb = mb => (mb / 1024).toFixed(mb >= 10240 ? 0 : 1);

function fmtUptime(s) {
  if (s == null) return '-';
  const d = Math.floor(s / 86400), h = Math.floor(s % 86400 / 3600), m = Math.floor(s % 3600 / 60);
  return (d ? d + 'd ' : '') + h + 'h ' + m + 'm';
}

/* ---------------------------------------------------------------- icons (24x24 stroke) */
const ICONS = {
  dashboard: '<rect x="3" y="3" width="7" height="9" rx="1.5"/><rect x="14" y="3" width="7" height="5" rx="1.5"/><rect x="14" y="12" width="7" height="9" rx="1.5"/><rect x="3" y="16" width="7" height="5" rx="1.5"/>',
  activity: '<polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>',
  zap: '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>',
  sliders: '<line x1="4" y1="21" x2="4" y2="14"/><line x1="4" y1="10" x2="4" y2="3"/><line x1="12" y1="21" x2="12" y2="12"/><line x1="12" y1="8" x2="12" y2="3"/><line x1="20" y1="21" x2="20" y2="16"/><line x1="20" y1="12" x2="20" y2="3"/><line x1="1" y1="14" x2="7" y2="14"/><line x1="9" y1="8" x2="15" y2="8"/><line x1="17" y1="16" x2="23" y2="16"/>',
  wifi: '<path d="M5 12.55a11 11 0 0 1 14.08 0"/><path d="M1.42 9a16 16 0 0 1 21.16 0"/><path d="M8.53 16.11a6 6 0 0 1 6.95 0"/><line x1="12" y1="20" x2="12.01" y2="20"/>',
  trash: '<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>',
  eye: '<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/>',
  cpu: '<rect x="4" y="4" width="16" height="16" rx="2"/><rect x="9" y="9" width="6" height="6"/><line x1="9" y1="1" x2="9" y2="4"/><line x1="15" y1="1" x2="15" y2="4"/><line x1="9" y1="20" x2="9" y2="23"/><line x1="15" y1="20" x2="15" y2="23"/><line x1="20" y1="9" x2="23" y2="9"/><line x1="20" y1="14" x2="23" y2="14"/><line x1="1" y1="9" x2="4" y2="9"/><line x1="1" y1="14" x2="4" y2="14"/>',
  monitor: '<rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>',
  flask: '<path d="M9 3h6"/><path d="M10 3v6.5L4.5 19a1.5 1.5 0 0 0 1.3 2.2h12.4a1.5 1.5 0 0 0 1.3-2.2L14 9.5V3"/><path d="M7.5 15h9"/>',
  restore: '<polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/>',
  file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/>',
  gear: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>',
  shield: '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>',
  power: '<path d="M18.36 6.64a9 9 0 1 1-12.73 0"/><line x1="12" y1="2" x2="12" y2="12"/>',
  plug: '<path d="M12 22v-5"/><path d="M9 8V2"/><path d="M15 8V2"/><path d="M18 8v5a4 4 0 0 1-4 4h-4a4 4 0 0 1-4-4V8z"/>',
  clock: '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>',
  mouse: '<path d="M3 3l7.07 16.97 2.51-7.39 7.39-2.51L3 3z"/><path d="M13 13l6 6"/>',
  gamepad: '<line x1="6" y1="12" x2="10" y2="12"/><line x1="8" y1="10" x2="8" y2="14"/><line x1="15" y1="13" x2="15.01" y2="13"/><line x1="18" y1="11" x2="18.01" y2="11"/><rect x="2" y="6" width="20" height="12" rx="3"/>',
  check: '<polyline points="20 6 9 17 4 12"/>',
  x: '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>',
  alert: '<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>',
  play: '<polygon points="6 3 20 12 6 21 6 3"/>',
  stop: '<rect x="6" y="6" width="12" height="12" rx="1.5"/>',
  refresh: '<polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>',
  layers: '<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>',
  database: '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>',
  sparkle: '<path d="M12 3l1.9 5.1L19 10l-5.1 1.9L12 17l-1.9-5.1L5 10l5.1-1.9z"/><path d="M19 17l.7 1.8L21.5 19.5l-1.8.7L19 22l-.7-1.8-1.8-.7 1.8-.7z"/>',
  info: '<circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>',
  arrow: '<line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>',
  chev: '<polyline points="9 18 15 12 9 6"/>',
  search: '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>',
  code: '<polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/>',
  copy: '<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
  list: '<line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/>',
  rocket: '<path d="M4.5 16.5c-1.5 1.26-2 5-2 5s3.74-.5 5-2c.71-.84.7-2.13-.09-2.91a2.18 2.18 0 0 0-2.91-.09z"/><path d="M12 15l-3-3a22 22 0 0 1 2-3.95A12.88 12.88 0 0 1 22 2c0 2.72-.78 7.5-6 11a22.35 22.35 0 0 1-4 2z"/><path d="M9 12H4s.55-3.03 2-4c1.62-1.08 5 0 5 0"/><path d="M12 15v5s3.03-.55 4-2c1.08-1.62 0-5 0-5"/>',
  skip: '<line x1="5" y1="12" x2="19" y2="12"/>',
  bolt: '<path d="M11 21l1-7H7l7-11-1 7h5z"/>',
  link: '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>',
  folder: '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
  down: '<line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/>',
  up: '<line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/>'
};
function ic(name, cls) {
  return '<svg class="ic ' + (cls || '') + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + (ICONS[name] || ICONS.info) + '</svg>';
}

/* ---------------------------------------------------------------- server calls */
let netFails = 0;
async function api(path, body) {
  const o = body === undefined ? {} : { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body || {}) };
  const r = await fetch(path, o);
  let j = null;
  try { j = await r.json(); } catch (e) { j = { error: 'bad response' }; }
  if (r.status === 403) throw new Error('forbidden');
  return j;
}

/* ---------------------------------------------------------------- state */
const S = {
  cat: null, byId: {}, sys: null, settings: {}, states: {}, statesVer: -1, statesBusy: false, statesLoaded: false,
  live: null, H: {}, page: 'cockpit', sel: new Set(), run: null, lastRun: null, changes: null, connected: null,
  filters: { risk: 'all', q: '' }
};
const HMAX = 120;
['cpu', 'gpu', 'ram', 'disk', 'diskR', 'diskW', 'netDown', 'netUp', 'dpc', 'intr', 'ping', 'jitter'].forEach(k => S.H[k] = []);

function pushHist(s) {
  const map = { cpu: s.cpu, gpu: s.gpu, ram: s.ram, disk: s.disk, diskR: s.diskR, diskW: s.diskW, netDown: s.netDown, netUp: s.netUp, dpc: s.dpc, intr: s.intr, ping: s.ping ? s.ping.ms : null, jitter: s.ping ? s.ping.jitter : null };
  for (const k in map) { const a = S.H[k]; a.push(map[k] == null ? null : Number(map[k])); while (a.length > HMAX) a.shift(); }
}

/* risk and state words always come with an icon, never color alone */
const RISK = {
  safe: { label: 'Safe', icon: 'check', note: 'Low impact, easy to undo.' },
  moderate: { label: 'Moderate', icon: 'info', note: 'Changes system behavior. Originals are saved.' },
  aggressive: { label: 'Aggressive', icon: 'alert', note: 'Turns off features some apps rely on. Originals are saved.' },
  security: { label: 'Security risk', icon: 'shield', note: 'Lowers protection. Only if you understand the trade-off.' },
  danger: { label: 'Destructive', icon: 'x', note: 'Cannot be undone.' }
};
const STATE = {
  on: { label: 'On', icon: 'check' },
  partial: { label: 'Partial', icon: 'skip' },
  off: { label: 'Off', icon: 'x' }
};
function riskChip(r) { const k = RISK[r] || RISK.safe; return '<span class="chip r-' + r + '" title="' + esc(k.note) + '">' + ic(k.icon) + k.label + '</span>'; }
function stateChip(id) {
  const t = S.byId[id]; if (!t || !t.hasState) return '';
  const st = S.states[id];
  if (!Object.prototype.hasOwnProperty.call(S.states, id)) return '<span class="chip skel" data-st="' + id + '">Checking</span>';
  if (!st) return '<span class="chip off" data-st="' + id + '">' + ic('skip') + 'n/a</span>';
  const k = STATE[st];
  return '<span class="chip ' + st + '" data-st="' + id + '">' + ic(k.icon) + k.label + '</span>';
}

/* the same dedupe rules as the server: a tweak that already contains others makes them redundant */
function normalize(ids) {
  const set = new Set(ids.filter(i => S.byId[i]));
  set.forEach(id => (S.byId[id].covers || []).forEach(c => set.delete(c)));
  const ord = S.cat.order;
  return ord.filter(x => set.has(x)).concat(Array.from(set).filter(x => ord.indexOf(x) < 0));
}

/* ---------------------------------------------------------------- toasts and confetti */
function toast(msg, kind) {
  const t = document.createElement('div');
  t.className = 'toast';
  t.innerHTML = ic(kind === 'bad' ? 'alert' : kind === 'info' ? 'info' : 'check') + '<span>' + esc(msg) + '</span>';
  $('#toasts').appendChild(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 320); }, 3600);
}

function confetti() {
  if (document.documentElement.dataset.motion === 'off') return;
  const c = $('#confetti'), x = c.getContext('2d');
  c.width = innerWidth; c.height = innerHeight;
  const cols = ['#00e676', '#1de9b6', '#b388ff', '#4fc3f7', '#ffca28', '#ff80ab'];
  const ps = Array.from({ length: 140 }, () => ({ x: innerWidth * (.3 + Math.random() * .4), y: innerHeight * .35, vx: (Math.random() - .5) * 15, vy: -Math.random() * 15 - 4, g: .38 + Math.random() * .2, w: 5 + Math.random() * 6, h: 3 + Math.random() * 5, r: Math.random() * 6, vr: (Math.random() - .5) * .4, c: cols[Math.floor(Math.random() * cols.length)] }));
  let f = 0;
  (function loop() {
    x.clearRect(0, 0, c.width, c.height);
    ps.forEach(p => { p.vy += p.g; p.x += p.vx; p.y += p.vy; p.vx *= .99; p.r += p.vr; x.save(); x.translate(p.x, p.y); x.rotate(p.r); x.fillStyle = p.c; x.globalAlpha = Math.max(0, 1 - f / 110); x.fillRect(-p.w / 2, -p.h / 2, p.w, p.h); x.restore(); });
    if (++f < 110) requestAnimationFrame(loop); else x.clearRect(0, 0, c.width, c.height);
  })();
}

function copyText(t) {
  const done = () => toast('Copied');
  if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(done, () => fallbackCopy(t, done));
  else fallbackCopy(t, done);
}
function fallbackCopy(t, done) { const a = document.createElement('textarea'); a.value = t; a.style.position = 'fixed'; a.style.opacity = '0'; document.body.appendChild(a); a.select(); try { document.execCommand('copy'); done(); } catch (e) { toast('Copy failed', 'bad'); } a.remove(); }

/* count-up for headline numbers */
function tween(el, to, fmt) {
  if (!el) return;
  const from = parseFloat(el.dataset.v || '0') || 0; el.dataset.v = to;
  if (document.documentElement.dataset.motion === 'off' || from === to) { el.textContent = fmt ? fmt(to) : Math.round(to); return; }
  const t0 = performance.now(), d = 700;
  (function f(t) { const k = clamp((t - t0) / d, 0, 1), e = 1 - Math.pow(1 - k, 3); const v = from + (to - from) * e; el.textContent = fmt ? fmt(v) : Math.round(v); if (k < 1) requestAnimationFrame(f); })(t0);
}
'@

$JT_UI['charts.js'] = @'
/* Jiff Tweaks UI - charts: rings, sparklines, line charts. Plain SVG, no libraries.
   Rules: 2px lines, an end dot with a surface-colored ring, hover tooltips with every value,
   a legend whenever there are two or more series, and labels never use the series color. */
'use strict';

const CIRC = 2 * Math.PI * 42;   // ring radius 42 in a 100 box

function ringHtml(id, color, big) {
  return '<div class="ring' + (big ? ' lg' : '') + '" id="' + id + '"><svg viewBox="0 0 100 100" aria-hidden="true"><circle class="trk" cx="50" cy="50" r="42"/>' +
    '<circle class="arc" cx="50" cy="50" r="42" stroke="' + color + '" stroke-dasharray="' + CIRC.toFixed(1) + '" stroke-dashoffset="' + CIRC.toFixed(1) + '"/></svg></div>';
}
function setRing(id, pct) {
  const a = $('#' + id + ' .arc'); if (!a) return;
  a.setAttribute('stroke-dashoffset', (CIRC * (1 - clamp(pct, 0, 100) / 100)).toFixed(1));
}

/* ---------------------------------------------------------------- sparkline (tile) */
function sparkHtml(id, color) {
  return '<div class="spk" id="' + id + '" style="position:relative"><svg class="spark" viewBox="0 0 120 40" preserveAspectRatio="none" aria-hidden="true">' +
    '<defs><linearGradient id="' + id + 'g" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="' + color + '" stop-opacity=".32"/><stop offset="1" stop-color="' + color + '" stop-opacity="0"/></linearGradient></defs>' +
    '<path class="sa" fill="url(#' + id + 'g)" d=""/><path class="sl" fill="none" stroke="' + color + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round" vector-effect="non-scaling-stroke" d=""/></svg>' +
    '<i class="edot" style="position:absolute;width:9px;height:9px;margin:-4.5px 0 0 -4.5px;border-radius:50%;background:' + color + ';border:2px solid var(--surface-ring);box-sizing:content-box;left:100%;top:50%;opacity:0;transition:top .8s var(--ease)"></i></div>';
}
function setSpark(id, vals, max) {
  const root = $('#' + id); if (!root) return;
  const v = vals.filter(x => x != null);
  if (v.length < 2) return;
  const n = clamp(vals.length, 8, HMAX), off = n - vals.length, mx = max || Math.max(1, Math.max.apply(null, v) * 1.15);
  let d = '';
  vals.forEach((x, i) => { if (x == null) return; const px = (off + i) / (n - 1) * 120, py = 38 - clamp(x / mx, 0, 1) * 36; d += (d ? 'L' : 'M') + px.toFixed(1) + ' ' + py.toFixed(1); });
  const first = (off + vals.findIndex(x => x != null)) / (n - 1) * 120;
  root.querySelector('.sl').setAttribute('d', d);
  root.querySelector('.sa').setAttribute('d', d + 'L120 40L' + first.toFixed(1) + ' 40Z');
  const dot = root.querySelector('.edot'), last = v[v.length - 1];
  dot.style.top = (100 * (38 - clamp(last / mx, 0, 1) * 36) / 40).toFixed(1) + '%'; dot.style.opacity = 1;
}

/* ---------------------------------------------------------------- line chart */
function niceMax(m) { if (m <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(m))), f = m / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p; }

const charts = {};
function chartMount(id, h) {
  return '<div class="chart" id="' + id + '" data-h="' + (h || 200) + '"><div class="legend"></div><svg></svg><div class="tip"></div></div>';
}
/* series: [{label,color,vals,fmt}]  opt: {max, unit, fmt} */
function chartUpdate(id, series, opt) {
  const el = $('#' + id); if (!el) return;
  opt = opt || {};
  const svg = el.querySelector('svg'), tip = el.querySelector('.tip'), leg = el.querySelector('.legend');
  const W = Math.max(280, el.clientWidth), H = Number(el.dataset.h) || 200, L = 44, R = 16, T = 10, B = 22;
  let longest = 0; series.forEach(s => { if (s.vals.length > longest) longest = s.vals.length; });
  const pw = W - L - R, ph = H - T - B, n = clamp(longest, 12, HMAX);
  let mx = opt.max;
  if (!mx) { let m = 0; series.forEach(s => s.vals.forEach(v => { if (v != null && v > m) m = v; })); mx = niceMax(m * 1.1); }
  const fmt = opt.fmt || (v => nf(v, v < 10 ? 1 : 0));
  const X = i => L + i / (n - 1) * pw, Y = v => T + ph - clamp(v / mx, 0, 1) * ph;
  let g = '';
  for (let k = 0; k <= 4; k++) { const yy = T + ph * k / 4, val = mx * (1 - k / 4); g += '<line class="gl" x1="' + L + '" x2="' + (W - R) + '" y1="' + yy + '" y2="' + yy + '"/><text class="ax" x="' + (L - 8) + '" y="' + (yy + 3.5) + '" text-anchor="end">' + esc(fmt(val)) + '</text>'; }
  g += '<text class="ax" x="' + L + '" y="' + (H - 5) + '">-' + n + 's</text><text class="ax" x="' + (L + pw / 2) + '" y="' + (H - 5) + '" text-anchor="middle">-' + Math.round(n / 2) + 's</text><text class="ax" x="' + (W - R) + '" y="' + (H - 5) + '" text-anchor="end">now</text>';
  let lines = '', dots = '';
  series.forEach(s => {
    const off = n - s.vals.length; let d = '', last = null, lx = 0;
    s.vals.forEach((v, i) => { if (v == null) { return; } d += (d && s.vals[i - 1] != null ? 'L' : 'M') + X(off + i).toFixed(1) + ' ' + Y(v).toFixed(1); last = v; lx = off + i; });
    if (d) lines += '<path class="ln" stroke="' + s.color + '" d="' + d + '"/>';
    if (last != null) dots += '<circle cx="' + X(lx).toFixed(1) + '" cy="' + Y(last).toFixed(1) + '" r="4" fill="' + s.color + '" stroke="var(--surface-ring)" stroke-width="2"/>';
  });
  svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H); svg.setAttribute('height', H);
  const hi = el._hov;
  let hov = '';
  if (hi != null) {
    hov = '<line class="hov" x1="' + X(hi) + '" x2="' + X(hi) + '" y1="' + T + '" y2="' + (T + ph) + '"/>';
    series.forEach(s => { const j = hi - (n - s.vals.length); const v = s.vals[j]; if (v != null) hov += '<circle cx="' + X(hi) + '" cy="' + Y(v) + '" r="4.5" fill="' + s.color + '" stroke="var(--surface-ring)" stroke-width="2"/>'; });
  }
  svg.innerHTML = g + lines + dots + hov + '<rect x="' + L + '" y="' + T + '" width="' + pw + '" height="' + ph + '" fill="transparent" class="hit"/>';
  // legend: swatch carries the color, text stays neutral; shows the latest value
  leg.innerHTML = series.map(s => { const lastV = s.vals.length ? s.vals[s.vals.length - 1] : null; return '<span><i style="background:' + s.color + '"></i>' + esc(s.label) + '<b>' + (lastV == null ? '-' : esc((s.fmt || fmt)(lastV)) + (opt.unit ? ' ' + esc(opt.unit) : '')) + '</b></span>'; }).join('');
  el._series = series; el._geo = { W, H, L, R, T, B, pw, ph, n, fmt, unit: opt.unit };
  if (!el._wired) {
    el._wired = true;
    const move = e => {
      const r = svg.getBoundingClientRect(), geo = el._geo, px = (e.clientX - r.left) * (geo.W / r.width);
      let i = Math.round((px - geo.L) / geo.pw * (geo.n - 1)); i = clamp(i, 0, geo.n - 1);
      el._hov = i; chartUpdate(id, el._series, el._opt);
      const rows = el._series.map(s => { const v = s.vals[i - (geo.n - s.vals.length)]; return '<div class="tr"><span><i style="background:' + s.color + '"></i>' + esc(s.label) + '</span><b>' + (v == null ? '-' : esc((s.fmt || geo.fmt)(v)) + (geo.unit ? ' ' + esc(geo.unit) : '')) + '</b></div>'; }).join('');
      tip.innerHTML = '<div class="tt">' + (geo.n - 1 - i === 0 ? 'now' : (geo.n - 1 - i) + ' s ago') + '</div>' + rows;
      const tx = (i / (geo.n - 1) * geo.pw + geo.L) / geo.W * r.width;
      tip.style.left = (tx > r.width * .6 ? tx - tip.offsetWidth - 14 : tx + 14) + 'px'; tip.style.top = '34px'; tip.classList.add('show');
    };
    svg.addEventListener('mousemove', move);
    svg.addEventListener('mouseleave', () => { el._hov = null; tip.classList.remove('show'); chartUpdate(id, el._series, el._opt); });
  }
  el._opt = opt;
}
'@

$JT_UI['pages.js'] = @'
/* Jiff Tweaks UI - pages */
'use strict';

const NAV_TOP = [{ id: 'cockpit', title: 'Cockpit', icon: 'dashboard', sub: 'Your PC right now, and what the tweaks changed.' }, { id: 'live', title: 'Live', icon: 'activity', sub: 'Real-time CPU, GPU, memory, disk, network and latency. Updates every second.' }];
const NAV_BOTTOM = [{ id: 'changes', title: 'Changes', icon: 'list', sub: 'Everything the tweaks changed, with the original value next to it.' }, { id: 'logs', title: 'Logs', icon: 'file', sub: 'What each run did, line by line.' }, { id: 'settings', title: 'Settings', icon: 'gear', sub: 'Online features are off until you switch them on.' }];

function allPages() {
  const tw = S.cat.pages.filter(p => p.id !== 'restore').map(p => ({ id: p.id, title: p.title, icon: p.icon, sub: p.sub, tweaks: true }));
  const rs = S.cat.pages.filter(p => p.id === 'restore').map(p => ({ id: p.id, title: p.title, icon: p.icon, sub: p.sub, tweaks: true }));
  return { top: NAV_TOP, tweaks: tw, bottom: NAV_BOTTOM.slice(0, 1).concat(rs, NAV_BOTTOM.slice(1)) };
}
function pageInfo(id) { const a = allPages(); return a.top.concat(a.tweaks, a.bottom).find(p => p.id === id) || a.top[0]; }

function renderNav() {
  const a = allPages();
  const item = p => '<button class="nav-item' + (S.page === p.id ? ' on' : '') + '" data-act="nav" data-page="' + p.id + '">' + ic(p.icon) + '<span>' + esc(p.title) + '</span>' + (p.tweaks && p.id !== 'restore' ? '<span class="badge" data-badge="' + p.id + '"></span>' : '') + '</button>';
  $('#nav').innerHTML = '<div class="nav-sep">Overview</div>' + a.top.map(item).join('') + '<div class="nav-sep">Tweaks</div>' + a.tweaks.map(item).join('') + '<div class="nav-sep">Manage</div>' + a.bottom.map(item).join('');
  paintBadges();
}
function paintBadges() {
  S.cat.pages.forEach(p => {
    const b = $('[data-badge="' + p.id + '"]'); if (!b) return;
    const ids = p.ids.filter(i => S.byId[i] && S.byId[i].hasState && !((S.byId[i].covers || []).length));
    if (!S.statesLoaded || !ids.length) { b.textContent = ''; return; }
    const on = ids.filter(i => S.states[i] === 'on').length;
    b.textContent = on + '/' + ids.length; b.classList.toggle('ok', on === ids.length);
  });
}

/* ---------------------------------------------------------------- router */
function go(id) {
  if (!S.cat) return;
  S.page = id; renderNav();
  const p = pageInfo(id);
  $('#pageTitle').textContent = p.title; $('#pageSub').textContent = p.sub || '';
  const v = $('#view');
  let html = '';
  if (id === 'cockpit') html = pageCockpit();
  else if (id === 'live') html = pageLive();
  else if (id === 'changes') html = pageChanges();
  else if (id === 'logs') html = pageLogs();
  else if (id === 'settings') html = pageSettings();
  else html = pageTweaks(id);
  v.innerHTML = '<div class="page-in">' + html + '</div>';
  v.scrollTop = 0;
  $('#topRight').innerHTML = '<button class="btn sm ghost" data-act="refresh" title="Check every tweak again">' + ic('refresh', 'sm') + 'Re-check</button><button class="btn sm" data-act="palette" title="Search">' + ic('search', 'sm') + 'Search <span class="kbd">Ctrl K</span></button>';
  const post = { cockpit: afterCockpit, live: afterLive, changes: afterChanges, logs: afterLogs, settings: afterSettings }[id];
  if (post) post(); else afterTweaks(id);
  paintLive(); paintStates();
  history.replaceState(null, '', '#' + id);
}

/* ---------------------------------------------------------------- cockpit */
function tile(k, name, color) {
  return '<div class="card tile lift"><div class="t-head"><i class="sw" style="background:' + color + '"></i>' + name + '</div>' +
    '<div class="t-body"><div class="big"><span id="v-' + k + '">-</span><small id="u-' + k + '">%</small></div>' + ringHtml('r-' + k, color) + '</div>' +
    '<div class="sub" id="s-' + k + '">&nbsp;</div>' + sparkHtml('sp-' + k, color) + '</div>';
}

function pageCockpit() {
  const rk = { safe: 0, moderate: 1, aggressive: 2, security: 3, danger: 4 };
  const prof = S.cat.profiles.map(p => {
    const norm = normalize(p.ids), worst = norm.reduce((a, i) => rk[S.byId[i].risk] > rk[a] ? S.byId[i].risk : a, 'safe'), rb = norm.some(i => S.byId[i].reboot);
    const count = p.id === 'full' ? 'About 450 settings' : norm.length + ' tweaks';
    return '<div class="card lift profile glow" data-profile="' + p.id + '"><div class="p-head"><div class="p-ico">' + ic(p.icon) + '</div><div class="grow"><h4>' + esc(p.title) + '</h4><div data-pstate="' + p.id + '" style="margin-top:3px"></div></div></div>' +
      '<p>' + esc(p.line) + '</p><div class="tags">' + p.tags.map(t => '<span class="tag">' + esc(t) + '</span>').join('') + '</div>' +
      '<div class="t-chips">' + riskChip(worst) + (rb ? '<span class="chip reboot">' + ic('refresh') + 'Restart</span>' : '') + '<span class="tag">' + count + '</span></div>' +
      '<div class="meter"><i data-pmeter="' + p.id + '"></i></div>' +
      '<div class="p-foot"><button class="btn sm" data-act="preview-ids" data-ids="' + p.ids.join(',') + '">' + ic('eye', 'sm') + 'Preview</button><button class="btn sm primary" data-act="apply-ids" data-ids="' + p.ids.join(',') + '" data-title="' + esc(p.title) + ' profile">' + ic('play', 'sm') + 'Apply</button></div></div>';
  }).join('');
  return '<div class="g g4 stagger">' + tile('cpu', 'CPU', 'var(--c-cpu)') + tile('gpu', 'GPU', 'var(--c-gpu)') + tile('ram', 'Memory', 'var(--c-ram)') + tile('disk', 'Disk', 'var(--c-disk)') + '</div>' +
    '<div class="sec"><h2>Profiles</h2><span>One click. Every original value is saved first, and Restore puts it back.</span></div>' +
    '<div class="g g-main"><div class="g g3 stagger">' + prof + '</div>' +
    '<div class="card" id="thisPc"></div></div>' +
    '<div class="g g3" style="margin-top:16px"><div class="card" id="cSug"></div><div class="card" id="cImpact"></div><div class="card" id="cLat"></div></div>';
}

function afterCockpit() {
  renderThisPc(); renderSuggestions(); renderImpact(); renderLatencyShell();
}

function tuneScore() {
  const ids = S.cat.tweaks.filter(t => t.hasState && t.risk !== 'security' && t.risk !== 'danger' && !(t.covers || []).length && t.id !== 'apply_full' && !(t.vendor === 'nvidia' && !S.sys.hasNvidia) && !(t.vendor === 'amd' && !S.sys.hasAmd) && t.id !== 'power_delete_others' && t.id !== 'branding');
  let on = 0, tot = 0;
  ids.forEach(t => { const st = S.states[t.id]; if (!st) return; tot++; on += st === 'on' ? 1 : st === 'partial' ? .5 : 0; });
  return { pct: tot ? Math.round(100 * on / tot) : 0, tot, on: Math.round(on) };
}

function renderThisPc() {
  const el = $('#thisPc'); if (!el) return; const y = S.sys;
  const sc = tuneScore();
  el.innerHTML = '<h3>' + ic('monitor', 'sm') + 'This PC<span class="grow"></span></h3>' +
    '<div class="score" style="margin-bottom:14px">' + ringHtml('r-score', 'url(#jg)', true).replace('</div>', '<div class="mid"><div><b id="v-score">0</b><small>tuned</small></div></div></div>') +
    '<div class="grow"><div style="font-weight:600;font-size:15px">' + (S.statesLoaded ? (sc.pct >= 70 ? 'Well tuned' : sc.pct >= 35 ? 'Partly tuned' : 'Mostly stock') : 'Checking') + '</div><div class="muted" style="font-size:12.5px;margin-top:3px">Safe and moderate tweaks that are on.</div></div></div>' +
    '<dl class="kv"><dt>Windows</dt><dd>' + esc(y.os + ' ' + (y.version || '') + ' (' + y.build + (y.ubr ? '.' + y.ubr : '') + ')') + '</dd><dt>CPU</dt><dd title="' + esc(y.cpu) + '">' + esc(y.cpu + ' | ' + y.threads + ' threads') + '</dd><dt>Memory</dt><dd>' + gb(y.ramMb) + ' GB</dd><dt>GPU</dt><dd title="' + esc((y.gpus || []).join(', ')) + '">' + esc((y.gpus || []).join(', ') || '-') + '</dd><dt>Uptime</dt><dd id="v-up">-</dd><dt>PC name</dt><dd>' + esc(y.host) + '</dd></dl>';
  paintScore();
}
function paintScore() {
  const sc = tuneScore();
  if ($('#r-score')) { setRing('r-score', sc.pct); tween($('#v-score'), sc.pct); }
}

function profileStatus(p) {
  const ids = normalize(p.ids);
  let on = 0, tot = 0;
  ids.forEach(i => { const t = S.byId[i]; if (!t || !t.hasState) return; const st = S.states[i]; if (!st) return; tot++; on += st === 'on' ? 1 : st === 'partial' ? .5 : 0; });
  return { on, tot, pct: tot ? on / tot : 0 };
}

function renderSuggestions() {
  const el = $('#cSug'); if (!el) return;
  let h = '<h3>' + ic('rocket', 'sm') + 'Suggested next</h3>';
  if (!S.statesLoaded) { el.innerHTML = h + '<div class="skel" style="height:90px;border-radius:12px"></div>'; return; }
  const prio = ['reg_gaming', 'reg_input_ui', 'power_plan', 'net_global', 'privacy', 'tasks_disable', 'reg_system', 'net_interfaces', 'bcd_timer', 'devpower', 'reg_priority', 'svc_basic', 'nic_tweaks'];
  const cand = prio.filter(i => S.byId[i] && S.states[i] && S.states[i] !== 'on' && !S.sel.has(i)).slice(0, 3);
  if (!cand.length) { el.innerHTML = h + '<div class="empty">' + ic('check') + '<div>Everything recommended is on.</div></div>'; return; }
  el.innerHTML = h + cand.map(i => { const t = S.byId[i]; return '<div class="sug"><div class="p-ico" style="width:36px;height:36px;border-radius:11px">' + ic(t.icon) + '</div><div class="grow"><div class="s-t">' + esc(t.title) + '</div><div class="s-d">' + esc(t.short) + '</div></div>' +
    '<button class="btn sm" data-act="queue-add" data-id="' + i + '">' + ic('play', 'sm') + 'Add</button></div>'; }).join('');
}

function deltaRow(label, a, b, unit, fmt) {
  const d = b - a, f = fmt || (v => nf(v));
  const chip = d === 0 ? '<span class="d-c same">no change</span>' : '<span class="d-c ' + (d < 0 ? '' : 'up') + '">' + ic(d < 0 ? 'down' : 'up', 'sm') + (d > 0 ? '+' : '') + f(d) + '</span>';
  return '<div class="delta-row"><span class="d-l">' + label + '</span><span class="d-v">' + f(a) + ' <span class="muted">&rarr;</span> ' + f(b) + (unit ? ' ' + unit : '') + chip + '</span></div>';
}
async function renderImpact() {
  const el = $('#cImpact'); if (!el) return;
  let r = null;
  try { r = await api('/api/impact'); } catch (e) { }
  const run = r && r.run && r.run.after ? r.run : (S.lastRun && S.lastRun.before && S.lastRun.after ? S.lastRun : null);
  let h = '<h3>' + ic('activity', 'sm') + 'Impact of the last run</h3>';
  if (run && run.before && run.after) {
    h += deltaRow('Services running', run.before.services, run.after.services) + deltaRow('Processes', run.before.processes, run.after.processes) +
      deltaRow('Memory in use', run.before.ramUsedMb, run.after.ramUsedMb, 'MB') +
      '<div class="muted" style="font-size:12px;margin-top:10px">Measured right before and after the run. Most changes show their full effect after a restart.</div>';
  } else if (r && r.now) {
    h += '<div class="delta-row"><span class="d-l">Services running</span><span class="d-v">' + nf(r.now.services) + '</span></div><div class="delta-row"><span class="d-l">Processes</span><span class="d-v">' + nf(r.now.processes) + '</span></div><div class="delta-row"><span class="d-l">Memory in use</span><span class="d-v">' + nf(r.now.ramUsedMb) + ' MB</span></div>' +
      '<div class="muted" style="font-size:12px;margin-top:10px">Run a profile and the before and after numbers show up here.</div>';
  }
  el.innerHTML = h;
}

function renderLatencyShell() {
  const el = $('#cLat'); if (!el) return;
  el.innerHTML = '<h3>' + ic('bolt', 'sm') + 'Latency health</h3>' +
    '<div class="row" style="align-items:flex-end;gap:22px"><div><div class="big" style="font-size:48px"><span id="v-dpc">-</span><small>%</small></div><div class="muted" style="font-size:12px">DPC time</div></div><div><div class="big" style="font-size:48px"><span id="v-intr">-</span><small>%</small></div><div class="muted" style="font-size:12px">Interrupt time</div></div></div>' +
    '<div id="latVerdict" style="margin-top:12px"></div>' + sparkHtml('sp-dpc', 'var(--c-dpc)') +
    '<div class="explain">DPC and interrupt time are the CPU time drivers take before your game gets to run. Staying under 1% keeps input and audio smooth.</div><div id="latPing" style="margin-top:10px"></div>';
}

/* ---------------------------------------------------------------- tweak pages */
function twCard(t) {
  const miss = (t.vendor === 'nvidia' && !S.sys.hasNvidia) || (t.vendor === 'amd' && !S.sys.hasAmd);
  return '<article class="card lift tw' + (S.sel.has(t.id) ? ' sel' : '') + '" data-id="' + t.id + '" data-risk="' + t.risk + '" data-q="' + esc((t.title + ' ' + t.short + ' ' + (t.tags || []).join(' ')).toLowerCase()) + '">' +
    '<div class="t-top"><div class="t-ico">' + ic(t.icon) + '</div><div class="t-name"><h4>' + esc(t.title) + '</h4></div><div class="chk">' + ic('check') + '</div></div>' +
    '<div class="t-short">' + esc(t.short) + '</div>' +
    '<div class="t-chips">' + riskChip(t.risk) + stateChip(t.id) + (t.reboot ? '<span class="chip reboot">' + ic('refresh') + 'Restart</span>' : '') + (miss ? '<span class="chip off">' + ic('skip') + 'Not detected</span>' : '') + '</div>' +
    '<div class="t-tags">' + (t.tags || []).slice(0, 5).map(x => '<span class="tag">' + esc(x) + '</span>').join('') + '</div>' +
    '<div class="t-more"><ul>' + (t.does || []).map(d => '<li>' + esc(d) + '</li>').join('') + '</ul>' + (t.note ? '<div class="note">' + esc(t.note) + '</div>' : '') + '</div>' +
    '<div class="t-foot"><button class="btn sm ghost icon" data-act="more" data-id="' + t.id + '" title="Details" aria-label="Details">' + ic('chev', 'sm') + '</button><span class="sp"></span>' +
    '<button class="btn sm" data-act="preview-ids" data-ids="' + t.id + '">' + ic('eye', 'sm') + 'Preview</button><button class="btn sm primary" data-act="apply-ids" data-ids="' + t.id + '" data-title="' + esc(t.title) + '">' + (t.kind === 'restore' ? ic('restore', 'sm') + 'Restore' : ic('play', 'sm') + 'Apply') + '</button></div></article>';
}

function pageTweaks(pid) {
  const p = S.cat.pages.find(x => x.id === pid); if (!p) return '';
  const f = p.featured && S.byId[p.featured];
  let h = '';
  if (pid === 'restore') h += '<div class="card" id="rsSummary" style="margin-bottom:16px"></div>';
  if (f) {
    h += '<div class="card glow featured lift" data-id="' + f.id + '"><div class="f-ico">' + ic(f.icon) + '</div><div><h2>' + esc(f.title) + '</h2><p>' + esc(f.short) + '</p><div class="f-meta">' + riskChip(f.risk) + stateChip(f.id) + (f.reboot ? '<span class="chip reboot">' + ic('refresh') + 'Restart</span>' : '') + (f.covers && f.covers.length ? '<span class="tag">Includes ' + f.covers.length + ' more</span>' : '') + '</div></div>' +
      '<div class="f-act"><button class="btn" data-act="preview-ids" data-ids="' + f.id + '">' + ic('eye', 'sm') + 'Preview</button><button class="btn primary" data-act="apply-ids" data-ids="' + f.id + '" data-title="' + esc(f.title) + '">' + (f.kind === 'restore' ? ic('restore', 'sm') + 'Restore' : ic('play', 'sm') + 'Apply') + '</button></div></div>';
  }
  const rest = p.ids.filter(i => i !== p.featured && S.byId[i]).map(i => S.byId[i]);
  if (rest.length > 3 && pid !== 'restore') {
    h += '<div class="filters"><div class="seg" id="riskSeg"><button class="on" data-act="risk" data-r="all">All</button><button data-act="risk" data-r="safe">Safe</button><button data-act="risk" data-r="moderate">Moderate</button><button data-act="risk" data-r="hard">Aggressive and up</button></div>' +
      '<input class="inp" id="twFilter" placeholder="Filter this page" style="min-width:220px" autocomplete="off"><span class="grow"></span><button class="btn sm ghost" data-act="sel-page" data-page="' + pid + '">Select all safe</button></div>';
  }
  h += '<div class="cards stagger" id="cards">' + rest.map(twCard).join('') + '</div>';
  return h;
}

function afterTweaks(pid) {
  if (pid === 'restore') renderRestoreSummary();
  const q = $('#twFilter');
  if (q) q.addEventListener('input', applyFilter);
}
function applyFilter() {
  const seg = $('#riskSeg .on'), r = seg ? seg.dataset.r : 'all', q = ($('#twFilter') || { value: '' }).value.toLowerCase().trim();
  $$('#cards .tw').forEach(c => {
    const rk = c.dataset.risk; let ok = true;
    if (r === 'safe') ok = rk === 'safe'; else if (r === 'moderate') ok = rk === 'moderate'; else if (r === 'hard') ok = rk === 'aggressive' || rk === 'security' || rk === 'danger';
    if (ok && q) ok = c.dataset.q.indexOf(q) >= 0;
    c.classList.toggle('hide', !ok);
  });
}

async function renderRestoreSummary() {
  const el = $('#rsSummary'); if (!el) return;
  let c = null; try { c = await api('/api/changes'); } catch (e) { }
  if (!el.isConnected) return;
  const n = c ? c.counts : null;
  if (!n || !n.total) { el.innerHTML = '<div class="row"><div class="p-ico">' + ic('database') + '</div><div><div style="font-weight:600">Nothing to restore yet</div><div class="muted">The first time a tweak changes something, its original value is saved here.</div></div></div>'; return; }
  const kinds = [['reg', 'Registry values'], ['svc', 'Services'], ['task', 'Scheduled tasks'], ['net', 'Network settings'], ['bcd', 'Boot settings'], ['power', 'Power plan sets'], ['pnp', 'Devices'], ['cmd', 'Commands']].filter(k => n[k[0]]);
  el.innerHTML = '<div class="row" style="gap:28px;flex-wrap:wrap"><div><div class="big">' + nf(n.total) + '</div><div class="muted" style="font-size:12px">original values saved</div></div>' + kinds.map(k => '<div><div class="big" style="font-size:30px">' + nf(n[k[0]]) + '</div><div class="muted" style="font-size:12px">' + k[1] + '</div></div>').join('') + '</div>';
}

/* ---------------------------------------------------------------- live page */
function metricCard(title, body) { return '<div class="card"><h3>' + title + '</h3>' + body + '</div>'; }

function pageLive() {
  return '<div class="card" style="margin-bottom:16px"><h3>' + ic('activity', 'sm') + 'Utilization</h3>' + chartMount('ch-util', 230) + '</div>' +
    '<div class="g g-main" style="margin-bottom:16px"><div class="card"><h3>' + ic('cpu', 'sm') + 'CPU cores<span class="grow"></span><span class="muted" id="coreInfo" style="text-transform:none;letter-spacing:0"></span></h3><div class="cores" id="cores"></div></div>' +
    '<div class="card"><h3>' + ic('bolt', 'sm') + 'Latency (drivers)</h3>' + chartMount('ch-lat', 150) + '<div class="explain">DPC and interrupt time: CPU time spent in drivers before your game can run. Keep both under 1%.</div></div></div>' +
    '<div class="g g3" style="margin-bottom:16px"><div class="card"><h3>' + ic('database', 'sm') + 'Disk throughput</h3>' + chartMount('ch-disk', 170) + '</div><div class="card"><h3>' + ic('wifi', 'sm') + 'Network</h3>' + chartMount('ch-net', 170) + '</div>' +
    '<div class="card"><h3>' + ic('link', 'sm') + 'Ping<span class="grow"></span><span id="pingSw"></span></h3><div id="pingBody"></div></div></div>' +
    '<div class="g g2"><div class="card"><h3>' + ic('monitor', 'sm') + 'GPU</h3><div class="row" style="gap:26px;flex-wrap:wrap" id="gpuBody"></div></div><div class="card"><h3>' + ic('layers', 'sm') + 'Busiest processes</h3><table class="tbl"><thead><tr><th>Process</th><th style="width:34%">CPU</th><th style="text-align:right">Memory</th></tr></thead><tbody id="topBody"></tbody></table></div></div>';
}
function afterLive() { renderPingControl(); }

function renderPingControl() {
  const sw = $('#pingSw'); if (!sw) return;
  sw.innerHTML = '<button class="sw-t' + (S.settings.PingOn ? ' on' : '') + '" data-act="ping-toggle" aria-label="Ping monitor"></button>';
  const b = $('#pingBody'); if (!b) return;
  if (!S.settings.PingOn) b.innerHTML = '<div class="empty" style="padding:26px 10px">' + ic('link') + '<div>The ping monitor is off.<br><span class="muted" style="font-size:12.5px">When on, this PC pings ' + esc(S.settings.PingHost || '1.1.1.1') + ' once per second. Nothing else is sent. Change the host in Settings.</span></div></div>';
  else b.innerHTML = chartMount('ch-ping', 130) + '<div class="row" style="gap:22px;margin-top:10px" id="pingStats"></div>';
}

function paintLiveCharts(s) {
  const H = S.H;
  chartUpdate('ch-util', [{ label: 'CPU', color: 'var(--c-cpu)', vals: H.cpu }, { label: 'GPU', color: 'var(--c-gpu)', vals: H.gpu }, { label: 'Memory', color: 'var(--c-ram)', vals: H.ram }, { label: 'Disk', color: 'var(--c-disk)', vals: H.disk }], { max: 100, unit: '%', fmt: v => nf(v, v < 10 && v % 1 ? 1 : 0) });
  chartUpdate('ch-lat', [{ label: 'DPC', color: 'var(--c-dpc)', vals: H.dpc, fmt: v => nf(v, 2) }, { label: 'Interrupt', color: 'var(--c-int)', vals: H.intr, fmt: v => nf(v, 2) }], { unit: '%', fmt: v => nf(v, 1) });
  chartUpdate('ch-disk', [{ label: 'Read', color: 'var(--c-disk)', vals: H.diskR }, { label: 'Write', color: 'var(--c-up)', vals: H.diskW }], { unit: 'MB/s' });
  chartUpdate('ch-net', [{ label: 'Down', color: 'var(--c-down)', vals: H.netDown }, { label: 'Up', color: 'var(--c-up)', vals: H.netUp }], { unit: 'Mbit/s' });
  if (S.settings.PingOn && $('#ch-ping')) {
    chartUpdate('ch-ping', [{ label: 'Ping', color: 'var(--c-ping)', vals: H.ping, fmt: v => nf(v, 0) }, { label: 'Jitter', color: 'var(--c-dpc)', vals: H.jitter, fmt: v => nf(v, 1) }], { unit: 'ms' });
    const p = s.ping; if (p && $('#pingStats')) $('#pingStats').innerHTML = '<div><div style="font-size:30px;font-weight:700;letter-spacing:-.02em">' + (p.ms == null ? '-' : nf(p.ms, 0)) + '<small style="font-size:14px;color:var(--text-2)"> ms</small></div><div class="muted" style="font-size:11.5px">to ' + esc(p.host) + '</div></div><div><div style="font-size:30px;font-weight:700;letter-spacing:-.02em">' + nf(p.jitter, 1) + '<small style="font-size:14px;color:var(--text-2)"> ms</small></div><div class="muted" style="font-size:11.5px">jitter</div></div><div><div style="font-size:30px;font-weight:700;letter-spacing:-.02em">' + nf(p.loss, 0) + '<small style="font-size:14px;color:var(--text-2)"> %</small></div><div class="muted" style="font-size:11.5px">loss</div></div>';
  }
  const cores = $('#cores');
  if (cores && s.cores) {
    if (cores.children.length !== s.cores.length) cores.innerHTML = s.cores.map((c, i) => '<div class="core"><b>' + i + '</b><i></i><small></small></div>').join('');
    s.cores.forEach((c, i) => { const el = cores.children[i]; if (!el) return; el.querySelector('i').style.height = clamp(c, 0, 100) + '%'; el.querySelector('small').textContent = Math.round(c) + '%'; });
    $('#coreInfo').textContent = (s.clock ? (s.clock / 1000).toFixed(2) + ' GHz' : '') + ' ' + (s.procs ? '| ' + s.procs + ' processes, ' + s.threads + ' threads' : '');
  }
  const gb_ = $('#gpuBody');
  if (gb_) gb_.innerHTML = '<div><div class="big" style="font-size:52px">' + nf(s.gpu, 0) + '<small>%</small></div><div class="muted" style="font-size:12px">engine load</div></div>' +
    (s.gpuTemp ? '<div><div class="big" style="font-size:52px">' + s.gpuTemp + '<small>C</small></div><div class="muted" style="font-size:12px">temperature</div></div>' : '') +
    (s.gpuMemTotalMb ? '<div><div class="big" style="font-size:52px">' + gb(s.gpuMemUsedMb) + '<small>/ ' + gb(s.gpuMemTotalMb) + ' GB</small></div><div class="muted" style="font-size:12px">video memory</div></div>' : '') +
    (s.gpuPower ? '<div><div class="big" style="font-size:52px">' + s.gpuPower + '<small>W</small></div><div class="muted" style="font-size:12px">power draw</div></div>' : '');
  const tb = $('#topBody');
  if (tb) tb.innerHTML = (s.top || []).map(p => '<tr><td>' + esc(p.name) + ' <span class="muted mono">' + p.pid + '</span></td><td><div class="row" style="gap:8px"><div class="bar grow"><i style="width:' + clamp(p.cpu * 3, 1, 100) + '%"></i></div><span style="font-variant-numeric:tabular-nums;width:42px;text-align:right">' + nf(p.cpu, 1) + '%</span></div></td><td style="text-align:right;font-variant-numeric:tabular-nums">' + nf(p.ramMb) + ' MB</td></tr>').join('') || '<tr><td colspan="3" class="muted">Collecting</td></tr>';
}

/* ---------------------------------------------------------------- live values on the cockpit */
function paintLive() {
  const s = S.live; if (!s) return;
  const set = (id, v) => { const e = $('#' + id); if (e) e.textContent = v; };
  if (S.page === 'cockpit') {
    set('v-cpu', nf(s.cpu, 0)); setRing('r-cpu', s.cpu); set('s-cpu', (s.clock ? (s.clock / 1000).toFixed(2) + ' GHz | ' : '') + S.sys.threads + ' threads'); setSpark('sp-cpu', S.H.cpu, 100);
    set('v-gpu', nf(s.gpu, 0)); setRing('r-gpu', s.gpu); set('s-gpu', s.gpuTemp ? s.gpuTemp + ' C | ' + gb(s.gpuMemUsedMb) + ' of ' + gb(s.gpuMemTotalMb) + ' GB VRAM' : (S.sys.gpus[0] || 'GPU')); setSpark('sp-gpu', S.H.gpu, 100);
    set('v-ram', nf(s.ram, 0)); setRing('r-ram', s.ram); set('s-ram', gb(s.ramUsedMb) + ' of ' + gb(s.ramTotalMb) + ' GB'); setSpark('sp-ram', S.H.ram, 100);
    set('v-disk', nf(s.disk, 0)); setRing('r-disk', s.disk); set('s-disk', 'Read ' + nf(s.diskR, 1) + ' | Write ' + nf(s.diskW, 1) + ' MB/s'); setSpark('sp-disk', S.H.disk, 100);
    set('v-up', fmtUptime(s.uptime));
    set('v-dpc', nf(s.dpc, 2)); set('v-intr', nf(s.intr, 2)); setSpark('sp-dpc', S.H.dpc);
    const v = $('#latVerdict');
    if (v) { const w = Math.max(s.dpc || 0, s.intr || 0); const k = w < 1 ? ['on', 'check', 'Good'] : w < 3 ? ['partial', 'skip', 'Elevated'] : ['off', 'alert', 'High']; v.innerHTML = '<span class="chip ' + k[0] + '">' + ic(k[1]) + k[2] + '</span>'; }
    const lp = $('#latPing');
    if (lp) lp.innerHTML = s.ping ? '<div class="row" style="gap:16px"><span class="chip">' + ic('link') + 'Ping ' + (s.ping.ms == null ? '-' : nf(s.ping.ms, 0) + ' ms') + '</span><span class="muted" style="font-size:12px">jitter ' + nf(s.ping.jitter, 1) + ' ms | loss ' + nf(s.ping.loss, 0) + '%</span></div>' : '<button class="btn sm ghost" data-act="nav" data-page="settings">' + ic('link', 'sm') + 'Turn on ping monitor</button>';
  } else if (S.page === 'live') paintLiveCharts(s);
}

/* ---------------------------------------------------------------- changes (the journal) */
const KIND = { reg: ['Registry', 'database'], svc: ['Service', 'sliders'], task: ['Task', 'clock'], net: ['Network', 'wifi'], bcd: ['Boot', 'power'], power: ['Power', 'plug'], pnp: ['Device', 'plug'], cmd: ['Command', 'code'], mm: ['Memory', 'cpu'], mp: ['Defender', 'shield'], file: ['File', 'file'], label: ['Drive', 'database'], host: ['PC name', 'monitor'], step: ['Step', 'info'] };
function pageChanges() { return '<div id="chBody"><div class="skel" style="height:160px;border-radius:16px"></div></div>'; }
async function afterChanges() {
  let c = null; try { c = await api('/api/changes'); } catch (e) { }
  const el = $('#chBody'); if (!el || !c) return;
  S.changes = c;
  if (!c.rows.length) { el.innerHTML = '<div class="card empty">' + ic('database') + '<div style="font-size:16px;color:var(--text)">Nothing changed yet</div><div>When a tweak changes something, the original value is written to a journal and listed here.</div></div>'; return; }
  const kinds = {}; c.rows.forEach(r => kinds[r.k] = (kinds[r.k] || 0) + 1);
  el.innerHTML = '<div class="filters"><div class="seg" id="kindSeg"><button class="on" data-act="kind" data-k="all">All ' + c.rows.length + '</button>' + Object.keys(kinds).map(k => '<button data-act="kind" data-k="' + k + '">' + ((KIND[k] || [k])[0]) + ' ' + kinds[k] + '</button>').join('') + '</div><input class="inp" id="chFilter" placeholder="Filter" autocomplete="off" style="min-width:200px"><span class="grow"></span><button class="btn sm danger" data-act="apply-ids" data-ids="restore_all" data-title="Restore everything">' + ic('restore', 'sm') + 'Restore everything</button></div>' +
    '<div class="card" style="padding:0"><div style="max-height:calc(100vh - 290px);overflow:auto"><table class="tbl"><thead><tr><th style="width:110px">Type</th><th>Setting</th><th style="width:22%">Now</th><th style="width:22%">Original</th></tr></thead><tbody id="chRows"></tbody></table></div></div>';
  S.chShow = 400; paintChangeRows(); $('#chFilter').addEventListener('input', paintChangeRows);
}
function paintChangeRows() {
  const c = S.changes; if (!c || !$('#chRows')) return;
  const k = ($('#kindSeg .on') || { dataset: { k: 'all' } }).dataset.k, q = ($('#chFilter') || { value: '' }).value.toLowerCase();
  let rows = c.rows.filter(r => (k === 'all' || r.k === k) && (!q || (r.target + ' ' + r.name).toLowerCase().indexOf(q) >= 0));
  const total = rows.length; rows = rows.slice(0, S.chShow);
  $('#chRows').innerHTML = rows.map(r => '<tr><td><span class="chip">' + ic((KIND[r.k] || [0, 'info'])[1]) + ((KIND[r.k] || [r.k])[0]) + '</span></td><td><div>' + esc(r.name || r.target) + '</div>' + (r.name ? '<div class="k">' + esc(r.target) + '</div>' : '') + '</td><td>' + (r.current == null ? '<span class="muted">not set</span>' : esc(r.current)) + '</td><td>' + (r.new == null ? '<span class="muted">-</span>' : r.new === '(deleted)' ? '<span class="muted">not set before</span>' : '<b>' + esc(r.new) + '</b>') + '</td></tr>').join('') +
    (total > rows.length ? '<tr><td colspan="4"><button class="btn sm" data-act="more-rows">Show more (' + (total - rows.length) + ' left)</button></td></tr>' : '');
}

/* ---------------------------------------------------------------- logs */
function pageLogs() { return '<div class="log-layout"><div class="log-list" id="logList"></div><div class="card" style="display:flex;flex-direction:column;min-height:0;padding:14px"><div class="filters" style="margin-bottom:8px"><div class="seg" id="logSeg"><button class="on" data-act="logf" data-f="all">All</button><button data-act="logf" data-f="bad">Problems</button><button data-act="logf" data-f="steps">Steps</button></div><span class="grow"></span><button class="btn sm ghost" data-act="open-logs">' + ic('folder', 'sm') + 'Open folder</button><button class="btn sm ghost" data-act="copy-log">' + ic('copy', 'sm') + 'Copy</button></div><div class="logbox" id="logView" style="flex:1"></div></div></div>'; }
async function afterLogs() {
  let r = null; try { r = await api('/api/logs'); } catch (e) { }
  const el = $('#logList'); if (!el || !r) return;
  if (!r.files.length) { el.innerHTML = '<div class="empty">' + ic('file') + '<div>No runs yet.</div></div>'; $('#logView').innerHTML = ''; return; }
  el.innerHTML = r.files.map((f, i) => '<button class="log-item' + (i === 0 ? ' on' : '') + '" data-act="log" data-name="' + esc(f.name) + '"><b>' + esc(f.time.replace('T', ' ')) + '</b><small>' + nf(f.size / 1024, 1) + ' KB</small></button>').join('');
  openLog(r.files[0].name);
}
async function openLog(name) {
  $$('.log-item').forEach(b => b.classList.toggle('on', b.dataset.name === name));
  const r = await api('/api/log?name=' + encodeURIComponent(name)); S.logLines = r.lines || []; S.logName = name; paintLog();
}
function paintLog() {
  const el = $('#logView'); if (!el) return;
  const f = ($('#logSeg .on') || { dataset: { f: 'all' } }).dataset.f;
  el.innerHTML = (S.logLines || []).filter(l => f === 'all' || (f === 'bad' && /\[(FAIL|NOTE)\]/.test(l)) || (f === 'steps' && /\[(TWEAK|STEP|SEQ|DONE)\]/.test(l))).map(logLine).join('');
}
function logLine(l) {
  const m = /^\[([\d:]+)\] \[(\w+)\] ?(.*)$/.exec(l);
  if (!m) return '<div class="ll"><span class="m">' + esc(l) + '</span></div>';
  return '<div class="ll t-' + m[2] + '"><span class="tm">' + m[1] + '</span><span class="tg">' + m[2] + '</span><span class="m">' + esc(m[3]) + '</span></div>';
}

/* ---------------------------------------------------------------- settings */
function pageSettings() {
  const s = S.settings;
  const row = (t, d, ctl) => '<div class="set-row"><div class="grow"><div class="s-t">' + t + '</div><div class="s-d">' + d + '</div></div>' + ctl + '</div>';
  return '<div class="g g-main"><div class="card"><h3>' + ic('link', 'sm') + 'Online features</h3><div class="muted" style="margin:-4px 0 6px">Everything works offline. These are the only things that can use the internet, and they are off until you switch them on.</div>' +
    row('Live ping monitor', 'Pings one host once per second and charts ping, jitter and loss on the Live page. Nothing else is sent.', '<div class="row"><input class="inp" id="pingHost" value="' + esc(s.PingHost || '1.1.1.1') + '" style="width:150px" spellcheck="false"><button class="sw-t' + (s.PingOn ? ' on' : '') + '" data-act="ping-toggle" aria-label="Ping monitor"></button></div>') +
    row('Check for updates', 'Reads a small text file (version.txt) from the project on GitHub and compares it with ' + esc(S.cat.version) + '. Nothing is downloaded or installed.', '<div class="row"><button class="btn sm" data-act="update-now">' + ic('refresh', 'sm') + 'Check now</button><button class="sw-t' + (s.UpdateCheck ? ' on' : '') + '" data-act="update-toggle" aria-label="Check on start"></button></div>') + '<div id="updRes" class="muted" style="font-size:12.5px;padding:0 0 4px"></div>' +
    '</div><div class="card"><h3>' + ic('gear', 'sm') + 'Interface</h3>' +
    row('Animations', 'Background motion, springy transitions and the confetti. Turn off for the lowest CPU use.', '<button class="sw-t' + (s.Motion !== false ? ' on' : '') + '" data-act="motion-toggle" aria-label="Animations"></button>') +
    row('Command palette', 'Press Ctrl+K anywhere to jump to a tweak or a page.', '<span class="kbd">Ctrl K</span>') + '</div></div>' +
    '<div class="g g2" style="margin-top:16px"><div class="card"><h3>' + ic('shield', 'sm') + 'What this app does and does not do</h3><ul class="muted" style="margin:0;padding-left:18px;display:grid;gap:7px;font-size:13px"><li>All of it is one plain-text PowerShell file. Open it in Notepad to read exactly what it does.</li><li>The window talks to this PC only (127.0.0.1) and needs a secret token created at launch.</li><li>It can only run tweaks from the built-in list. It cannot run any other command.</li><li>No telemetry, no accounts, nothing is uploaded.</li><li>Before a value changes, the original is saved in <span class="mono">' + esc((S.diag && S.diag.dataDir) || 'C:\\JiffTweaks') + '</span>.</li></ul></div>' +
    '<div class="card"><h3>' + ic('info', 'sm') + 'About<span class="grow"></span><button class="btn sm ghost" data-act="diag">' + ic('activity', 'sm') + 'Diagnostics</button></h3><dl class="kv"><dt>Version</dt><dd>' + esc(S.cat.version) + '</dd><dt>PowerShell</dt><dd>' + esc(S.sys.psVersion) + '</dd><dt>Administrator</dt><dd>' + (S.sys.admin ? 'Yes' : 'No') + '</dd><dt>Data folder</dt><dd class="mono" id="dataDir">-</dd></dl><div id="diagBox" class="diag" style="margin-top:10px"></div>' +
    '<div class="row" style="margin-top:14px;gap:8px"><button class="btn sm" data-act="open-logs">' + ic('folder', 'sm') + 'Open data folder</button><button class="btn sm danger" data-act="quit">' + ic('power', 'sm') + 'Stop Jiff Tweaks</button></div></div></div>';
}
async function afterSettings() {
  try { S.diag = await api('/api/diag'); const d = $('#dataDir'); if (d) d.textContent = S.diag.dataDir; } catch (e) { }
  const h = $('#pingHost');
  if (h) h.addEventListener('change', async () => { S.settings = await api('/api/settings', { PingHost: h.value.trim() }); toast('Ping host saved'); });
}
'@

$JT_UI['flows.js'] = @'
/* Jiff Tweaks UI - flows: queue, run drawer, preview, code, confirm, palette */
'use strict';

/* ---------------------------------------------------------------- selection queue */
function paintQueue() {
  const q = $('#queue'), ids = Array.from(S.sel);
  if (!ids.length) { q.classList.remove('show'); return; }
  const norm = normalize(ids), hidden = ids.length - norm.length;
  const names = norm.map(i => S.byId[i].title);
  q.innerHTML = '<div class="q-n">' + ids.length + '</div><div class="q-t"><b>selected</b>' + esc(names.join(', ')) + (hidden ? ' (' + hidden + ' already included)' : '') + '</div>' +
    '<button class="btn sm ghost" data-act="queue-clear">Clear</button><button class="btn" data-act="preview-ids" data-ids="' + ids.join(',') + '">' + ic('eye', 'sm') + 'Preview</button><button class="btn primary" data-act="apply-ids" data-ids="' + ids.join(',') + '" data-title="' + ids.length + ' selected">' + ic('play', 'sm') + 'Apply</button>';
  q.classList.add('show');
}
function toggleSel(id) {
  if (S.sel.has(id)) S.sel.delete(id); else S.sel.add(id);
  $$('.tw[data-id="' + id + '"]').forEach(c => c.classList.toggle('sel', S.sel.has(id)));
  paintQueue();
}

/* ---------------------------------------------------------------- modal plumbing */
function openModal(html) { $('#modalBox').innerHTML = html; $('#modal').classList.add('show'); $('#modal').setAttribute('aria-hidden', 'false'); }
function closeModal() { $('#modal').classList.remove('show'); $('#modal').setAttribute('aria-hidden', 'true'); if (S.modalResolve) { const r = S.modalResolve; S.modalResolve = null; r(false); } }
function modalHead(title, sub, icon) { return '<div class="mh"><div class="p-ico">' + ic(icon || 'info') + '</div><div class="grow"><h2>' + esc(title) + '</h2>' + (sub ? '<div class="sub">' + esc(sub) + '</div>' : '') + '</div><button class="btn sm ghost" data-close="modal" aria-label="Close">' + ic('x', 'sm') + '</button></div>'; }

/* ---------------------------------------------------------------- preview */
const SVC_START = { 0: 'Boot', 1: 'System', 2: 'Automatic', 3: 'Manual', 4: 'Disabled' };
function showVal(c, v) {
  if (v == null) return '<span class="muted">not set</span>';
  if (c.kind === 'svc') return esc(SVC_START[v] || v);
  if (c.kind === 'task') return v ? 'Enabled' : 'Disabled';
  if (c.kind === 'pnp' || c.kind === 'mm') return v === true || v === 'True' ? 'On' : 'Off';
  return esc(v);
}
async function openPreview(ids) {
  openModal(modalHead('Preview', 'Reading this PC ...', 'eye') + '<div class="mb"><div class="skel" style="height:140px;border-radius:14px"></div></div>');
  let r;
  try { r = await api('/api/preview', { ids }); } catch (e) { closeModal(); toast('Preview failed', 'bad'); return; }
  if (r.error) { closeModal(); toast(r.error, 'bad'); return; }
  const total = r.items.reduce((a, i) => a + i.count, 0);
  const body = r.items.map((it, n) => {
    const shown = it.changes.slice(0, 300);
    return '<details class="pv-item"' + (r.items.length === 1 || n === 0 ? ' open' : '') + '><summary>' + ic('chev', 'chev') + '<b>' + esc(it.title) + '</b><span class="pv-n">' + it.count + ' change' + (it.count === 1 ? '' : 's') + '</span></summary>' +
      (it.count ? '<div style="max-height:360px;overflow:auto"><table class="tbl"><thead><tr><th style="width:96px">Type</th><th>Setting</th><th style="width:20%">Now</th><th style="width:20%">After</th></tr></thead><tbody>' +
        shown.map(c => '<tr><td><span class="chip">' + ic((KIND[c.kind] || [0, 'info'])[1]) + ((KIND[c.kind] || [c.kind])[0]) + '</span></td><td><div>' + esc(c.name || c.target) + '</div>' + (c.name ? '<div class="k">' + esc(c.target) + '</div>' : '') + '</td><td>' + (c.kind === 'cmd' || c.kind === 'step' ? '<span class="muted">runs</span>' : '<span class="old' + (c.current == null ? ' na' : '') + '">' + showVal(c, c.current) + '</span>') + '</td><td>' + (c.kind === 'cmd' || c.kind === 'step' ? '' : '<span class="new">' + showVal(c, c.new) + '</span>') + '</td></tr>').join('') +
        (it.count > shown.length ? '<tr><td colspan="4" class="muted">... and ' + (it.count - shown.length) + ' more</td></tr>' : '') + '</tbody></table></div>' : '<div class="muted" style="padding:14px">Nothing to change. Everything here is already set.</div>') + '</details>';
  }).join('');
  const isRestore = ids.every(i => S.byId[i] && S.byId[i].kind === 'restore');
  openModal(modalHead('Preview', total + ' change' + (total === 1 ? '' : 's') + ' across ' + r.items.length + ' tweak' + (r.items.length === 1 ? '' : 's') + '. Nothing has been changed yet.', 'eye') +
    '<div class="mb">' + body + '</div><div class="mf"><span class="muted" style="font-size:12.5px">' + (isRestore ? 'Shows the values that will be written back.' : 'Every value in the Now column is saved before it is changed.') + '</span><span class="grow"></span><button class="btn" data-close="modal">Close</button><button class="btn primary" data-act="modal-apply" data-ids="' + ids.join(',') + '">' + ic('play', 'sm') + (isRestore ? 'Restore' : 'Apply') + '</button></div>');
}

/* ---------------------------------------------------------------- confirm for risky changes */
function confirmRisk(ids) {
  const norm = normalize(ids), isHard = i => ['aggressive', 'security', 'danger'].indexOf(S.byId[i].risk) >= 0;
  const hard = [];
  norm.forEach(i => { if (isHard(i) && hard.indexOf(i) < 0) hard.push(i); (S.byId[i].covers || []).forEach(c => { if (S.byId[c] && isHard(c) && hard.indexOf(c) < 0) hard.push(c); }); });
  if (!hard.length) return Promise.resolve(true);
  const heavy = hard.some(i => ['security', 'danger'].indexOf(S.byId[i].risk) >= 0);
  const list = hard.map(i => { const t = S.byId[i]; return '<div class="step" style="margin-bottom:6px"><div class="row grow">' + riskChip(t.risk) + '<b>' + esc(t.title) + '</b></div></div><div class="muted" style="font-size:12.5px;margin:-2px 0 10px 4px">' + esc(t.note || RISK[t.risk].note) + '</div>'; }).join('');
  return new Promise(res => {
    openModal(modalHead(heavy ? 'This lowers security or cannot be undone' : 'This turns features off', 'Read this before you continue', 'alert') + '<div class="mb">' + list +
      '<div class="muted" style="font-size:12.5px">Originals are saved first, so Restore can put them back. Removed Store apps cannot be restored, they can be reinstalled from the Store.</div></div>' +
      '<div class="mf"><span class="grow"></span><button class="btn" data-act="confirm-no">Cancel</button><button class="btn ' + (heavy ? 'danger' : 'primary') + '" data-act="confirm-yes">' + ic('play', 'sm') + 'I understand, apply</button></div>');
    S.modalResolve = res; S.confirmResolve = res;
  });
}

/* ---------------------------------------------------------------- run drawer */
const R = { ids: [], from: 0, timer: null, done: 0, stepFrac: 0, running: false, lines: [], current: null, fails: [], title: '', t0: 0 };

function openDrawer(ids, title) {
  R.ids = ids; R.from = 0; R.done = 0; R.stepFrac = 0; R.running = true; R.lines = []; R.current = null; R.fails = []; R.title = title; R.t0 = Date.now();
  $('#dwTitle').textContent = title || 'Applying';
  $('#dwKicker').textContent = 'Running';
  $('#dwSub').textContent = ids.length + ' tweak' + (ids.length === 1 ? '' : 's') + ' in order';
  $('#dwSteps').innerHTML = ids.map(i => '<div class="step" data-sid="' + i + '"><span class="s-i"></span><span>' + esc(S.byId[i].title) + '</span><span class="s-c"></span></div>').join('');
  $('#dwLog').innerHTML = ''; $('#dwSummary').className = 'summary'; $('#dwSummary').innerHTML = '';
  $('#dwHead').className = 'dw-head'; $('#dwRing').classList.remove('finished'); $('#dwCheck').innerHTML = '';
  $('#dwFoot').innerHTML = '<span class="muted" style="font-size:12.5px" id="dwTime">0 s</span><span class="grow"></span><button class="btn danger" data-act="dw-cancel">' + ic('stop', 'sm') + 'Stop after this step</button>';
  setDrawerPct(0);
  $('#drawer').classList.add('show'); $('#drawer').setAttribute('aria-hidden', 'false');
}
function closeDrawer() { if (R.running) return; $('#drawer').classList.remove('show'); $('#drawer').setAttribute('aria-hidden', 'true'); }
function setDrawerPct(p) {
  p = clamp(p, 0, 100);
  $('#dwArc').setAttribute('stroke-dasharray', CIRC.toFixed(1)); $('#dwArc').setAttribute('stroke-dashoffset', (CIRC * (1 - p / 100)).toFixed(1));
  $('#dwPct').textContent = Math.round(p);
}
function stepEl(id) { return $('#dwSteps [data-sid="' + id + '"]'); }

function handleRunLine(l) {
  R.lines.push(l);
  const box = $('#dwLog'), near = box.scrollHeight - box.scrollTop - box.clientHeight < 60;
  box.insertAdjacentHTML('beforeend', logLine(l));
  while (box.children.length > 700) box.removeChild(box.firstChild);
  if (near) box.scrollTop = box.scrollHeight;
  const m = /^\[[\d:]+\] \[(\w+)\] ?(.*)$/.exec(l); if (!m) return;
  const tag = m[1], msg = m[2];
  if (tag === 'TWEAK') {
    const id = msg.split('|')[0]; R.current = id; R.stepFrac = 0;
    const e = stepEl(id); if (e) { e.className = 'step run'; e.scrollIntoView({ block: 'nearest' }); }
    $('#dwSub').textContent = msg.split('|')[1] || '';
  } else if (tag === 'STEP') {
    const mm = /^(\d+)\/(\d+) (.*)$/.exec(msg);
    if (mm) {
      R.stepFrac = (Number(mm[1]) - 1) / Number(mm[2]); $('#dwSub').textContent = 'Step ' + mm[1] + ' of ' + mm[2] + ': ' + mm[3];
      const prev = $('#dwSteps .sub.run'); if (prev) { prev.className = 'step sub done'; prev.querySelector('.s-i').innerHTML = ic('check'); }
      const par = stepEl(R.current);
      if (par) { let at = par; while (at.nextElementSibling && at.nextElementSibling.classList.contains('sub')) at = at.nextElementSibling; const el = document.createElement('div'); el.className = 'step sub run'; el.innerHTML = '<span class="s-i"></span><span>' + esc(mm[3]) + '</span><span class="s-c">' + mm[1] + '/' + mm[2] + '</span>'; at.insertAdjacentElement('afterend', el); el.scrollIntoView({ block: 'nearest' }); }
    }
  } else if (tag === 'DONE') {
    const mm = /^(\S+) ok=(\d+) fail=(\d+) skip=(\d+)/.exec(msg);
    if (mm) {
      const e = stepEl(mm[1]); R.done++; R.stepFrac = 0;
      $$('#dwSteps .sub.run').forEach(x => { x.className = 'step sub done'; x.querySelector('.s-i').innerHTML = ic('check'); });
      if (e) { e.className = 'step ' + (Number(mm[3]) > 0 ? 'err' : 'done'); e.querySelector('.s-i').innerHTML = ic(Number(mm[3]) > 0 ? 'x' : 'check'); e.querySelector('.s-c').textContent = mm[2] + ' changed' + (Number(mm[3]) ? ', ' + mm[3] + ' failed' : '') + (Number(mm[4]) ? ', ' + mm[4] + ' skipped' : ''); }
    }
  } else if (tag === 'FAIL') R.fails.push(msg);
  else if (tag === 'SEQ' && /cancelled/.test(msg)) R.cancelled = true;
  const total = R.ids.length, cur = R.current && stepEl(R.current) && stepEl(R.current).classList.contains('run') ? R.stepFrac : 0;
  setDrawerPct((R.done + cur) / total * 100);
}

async function pollRun() {
  clearTimeout(R.timer);
  let r;
  try { r = await api('/api/run?from=' + R.from); } catch (e) { R.timer = setTimeout(pollRun, 900); return; }
  (r.lines || []).forEach(handleRunLine);
  R.from = r.next != null ? r.next : R.from;
  const tm = $('#dwTime'); if (tm) tm.textContent = Math.round((Date.now() - R.t0) / 1000) + ' s';
  if (r.running) { R.timer = setTimeout(pollRun, 350); return; }
  finishRun(r);
}

function finishRun(r) {
  R.running = false;
  const res = r.result || { Ok: 0, Fail: 0, Skip: 0, Done: 0, Reboot: false };
  const bad = res.Fail > 0 || r.error;
  S.lastRun = { before: r.before, after: r.after };
  setDrawerPct(R.cancelled ? ($('#dwPct').textContent * 1) : 100);
  $('#dwKicker').textContent = R.cancelled ? 'Stopped' : bad ? 'Finished with problems' : 'Finished';
  $('#dwTitle').textContent = R.cancelled ? 'Stopped before the end' : bad ? 'Done, with some problems' : 'All done';
  { const sec = Math.round((Date.now() - R.t0) / 1000); $('#dwSub').textContent = sec + (sec === 1 ? ' second' : ' seconds'); }
  $('#dwCheck').innerHTML = ic(bad ? 'alert' : 'check');
  $('#dwRing').classList.add('finished'); $('#dwHead').classList.toggle('bad', !!bad);
  $$('#dwSteps .step.run').forEach(e => { e.className = 'step skip'; });
  $$('#dwSteps .step:not(.sub)').forEach(e => { if (!e.querySelector('.s-i').innerHTML) { e.classList.add('skip'); e.querySelector('.s-c').textContent = 'not run'; } });
  let imp = '';
  if (r.before && r.after) imp = '<div style="margin-top:10px">' + deltaRow('Services running', r.before.services, r.after.services) + deltaRow('Processes', r.before.processes, r.after.processes) + '</div>';
  $('#dwSummary').innerHTML = '<div class="s-row"><div><div class="s-n">' + nf(res.Ok) + '</div><div class="s-l">changed</div></div><div><div class="s-n">' + nf(res.Fail) + '</div><div class="s-l">failed</div></div><div><div class="s-n">' + nf(res.Skip) + '</div><div class="s-l">skipped</div></div></div>' +
    (res.Reboot ? '<div class="chip reboot" style="margin-top:10px">' + ic('refresh') + 'Restart Windows to finish applying</div>' : '') + imp +
    (R.fails.length ? '<div class="muted" style="font-size:12px;margin-top:10px">First problem: ' + esc(R.fails[0]) + '</div>' : '');
  $('#dwSummary').classList.add('show');
  $('#dwFoot').innerHTML = '<button class="btn sm ghost" data-act="dw-copy">' + ic('copy', 'sm') + 'Copy report</button><button class="btn sm ghost" data-act="nav" data-page="logs" data-close="drawer">' + ic('file', 'sm') + 'Log</button><span class="grow"></span>' +
    (res.Reboot ? '<button class="btn" data-act="reboot">' + ic('refresh', 'sm') + 'Restart now</button>' : '') + '<button class="btn primary" data-close="drawer">Done</button>';
  if (!bad && res.Ok > 0 && !R.cancelled) confetti();
  R.cancelled = false;
  api('/api/refresh').catch(() => { });
  S.statesVer = -2;   // forces a repaint when the new states arrive
  if (S.page === 'cockpit') renderImpact();
}

async function startRun(ids, title) {
  ids = normalize(ids);
  if (!ids.length) { toast('Nothing to run'); return; }
  const ok = await confirmRisk(ids);
  closeModal(); S.modalResolve = null;
  if (!ok) return;
  const r = await api('/api/run', { ids });
  if (r.error) { toast(r.error, 'bad'); return; }
  S.sel.clear(); paintQueue(); $$('.tw.sel').forEach(c => c.classList.remove('sel'));
  openDrawer(r.ids, title);
  pollRun();
}

async function resumeRun() {
  const r = await api('/api/run?from=0');
  if (r && r.running) { openDrawer(r.ids, 'Run in progress'); pollRun(); }
}

function runReport() {
  const lines = ['Jiff Tweaks ' + S.cat.version + ' - ' + new Date().toLocaleString(), S.sys.os + ' ' + (S.sys.version || '') + ' build ' + S.sys.build, ''];
  lines.push(R.title + ': ' + R.ids.map(i => S.byId[i].title).join(', '));
  $$('#dwSteps .step').forEach(e => lines.push('  ' + e.textContent.trim().replace(/\s+/g, ' ')));
  if (R.fails.length) { lines.push('', 'Problems:'); R.fails.slice(0, 20).forEach(f => lines.push('  ' + f)); }
  return lines.join('\n');
}

/* ---------------------------------------------------------------- command palette */
let palItems = [], palIdx = 0;
function openPalette() { $('#palette').classList.add('show'); const i = $('#palInput'); i.value = ''; paintPalette(''); i.focus(); requestAnimationFrame(() => i.focus()); setTimeout(() => i.focus(), 60); }
function closePalette() { $('#palette').classList.remove('show'); }
function paintPalette(q) {
  q = q.toLowerCase().trim();
  const items = [];
  const a = allPages();
  a.top.concat(a.tweaks, a.bottom).forEach(p => items.push({ t: p.title, s: 'Page', icon: p.icon, score: 3, run: () => go(p.id) }));
  S.cat.profiles.forEach(p => items.push({ t: 'Apply ' + p.title + ' profile', s: 'Profile', icon: p.icon, score: 2, run: () => startRun(p.ids, p.title + ' profile') }));
  S.cat.tweaks.forEach(t => items.push({ t: t.title, s: t.page, icon: t.icon, score: 1, hay: (t.short + ' ' + (t.tags || []).join(' ') + ' ' + t.id).toLowerCase(), run: () => { go(t.page); setTimeout(() => { const c = $('.tw[data-id="' + t.id + '"]') || $('.featured[data-id="' + t.id + '"]'); if (c) { c.scrollIntoView({ block: 'center' }); c.animate([{ boxShadow: '0 0 0 0 rgba(0,230,118,.7)' }, { boxShadow: '0 0 0 14px rgba(0,230,118,0)' }], { duration: 900 }); } }, 80); } }));
  items.push({ t: 'Re-check all tweaks', s: 'Action', icon: 'refresh', score: 0, run: () => recheck() });
  const res = items.map(i => { let sc = -1; const tl = i.t.toLowerCase(); if (!q) sc = i.score; else if (tl.indexOf(q) === 0) sc = 100; else if (tl.indexOf(q) > 0) sc = 60; else if (i.hay && i.hay.indexOf(q) >= 0) sc = 20; return { i, sc }; }).filter(x => x.sc >= 0).sort((a, b) => b.sc - a.sc).slice(0, 9).map(x => x.i);
  palItems = res; palIdx = 0;
  $('#palList').innerHTML = res.map((i, n) => '<div class="pal-i' + (n === 0 ? ' on' : '') + '" data-pal="' + n + '">' + ic(i.icon) + '<span class="p-t">' + esc(i.t) + '</span><span class="p-s">' + esc(i.s) + '</span></div>').join('') || '<div class="empty">No match</div>';
}
function palMove(d) { if (!palItems.length) return; palIdx = (palIdx + d + palItems.length) % palItems.length; $$('#palList .pal-i').forEach((e, n) => e.classList.toggle('on', n === palIdx)); const e = $$('#palList .pal-i')[palIdx]; if (e) e.scrollIntoView({ block: 'nearest' }); }
function palRun(n) { const i = palItems[n == null ? palIdx : n]; if (!i) return; closePalette(); i.run(); }
'@

$JT_UI['app.js'] = @'
/* Jiff Tweaks UI - boot, polling, events */
'use strict';

/* ---------------------------------------------------------------- states */
function paintStates() {
  $$('[data-st]').forEach(e => { const n = document.createElement('div'); n.innerHTML = stateChip(e.dataset.st); if (n.firstChild) e.replaceWith(n.firstChild); });
  if (!S.cat) return;
  paintBadges();
  S.cat.profiles.forEach(p => {
    const st = profileStatus(p), el = $('[data-pstate="' + p.id + '"]'), m = $('[data-pmeter="' + p.id + '"]');
    if (el) el.innerHTML = !S.statesLoaded ? '<span class="chip skel">Checking</span>' : st.pct >= .98 ? '<span class="chip on">' + ic('check') + 'Applied</span>' : st.pct > .02 ? '<span class="chip partial">' + ic('skip') + Math.round(st.pct * 100) + '%</span>' : '<span class="chip off">' + ic('x') + 'Not applied</span>';
    if (m) m.style.width = Math.round(st.pct * 100) + '%';
  });
  if (S.page === 'cockpit') { paintScore(); renderSuggestions(); const t = $('#thisPc .grow div'); }
  if (S.page === 'cockpit' && $('#thisPc')) { const sc = tuneScore(); const w = $('#thisPc .score .grow div'); if (w && S.statesLoaded) w.textContent = sc.pct >= 70 ? 'Well tuned' : sc.pct >= 35 ? 'Partly tuned' : 'Mostly stock'; }
}

async function pollStates() {
  try {
    const r = await api('/api/states');
    const changed = r.ver !== S.statesVer || r.busy !== S.statesBusy;
    S.statesBusy = r.busy;
    if (changed || r.busy) {
      S.states = r.map || {};
      const was = S.statesLoaded; S.statesLoaded = r.ver > 0;
      S.statesVer = r.ver;
      paintStates();
    }
  } catch (e) { }
  setTimeout(pollStates, S.statesBusy ? 1200 : 3500);
}
async function recheck() { toast('Checking every tweak', 'info'); S.statesLoaded = S.statesLoaded; await api('/api/refresh'); S.statesVer = -2; }

/* ---------------------------------------------------------------- live */
let lastT = 0;
async function pollLive() {
  try {
    const r = await api('/api/live');
    if (r.sample) {
      if (r.sample.t !== lastT) { lastT = r.sample.t; S.live = r.sample; pushHist(r.sample); paintLive(); }
    }
    setConn(true);
  } catch (e) { setConn(false); }
  setTimeout(pollLive, 1000);
}
/* banners live in #banners, each under its own id so one never wipes another */
function setBanner(id, html) {
  let e = document.getElementById('bn-' + id);
  if (!html) { if (e) e.remove(); return; }
  if (!e) { e = document.createElement('div'); e.id = 'bn-' + id; $('#banners').appendChild(e); }
  e.innerHTML = html;
}
function setConn(ok) {
  if (ok) netFails = 0; else netFails++;
  const bad = netFails >= 3;
  if (S.connected === !bad) return;
  S.connected = !bad;
  $('#liveDot').classList.toggle('off', bad); $('#connLabel').textContent = bad ? 'Disconnected' : 'Live | 1 s';
  setBanner('conn', bad ? '<div class="banner bad">' + ic('alert') + '<div class="grow">The Jiff Tweaks script is not running any more. Start it again from PowerShell.</div></div>' : '');
}

/* ---------------------------------------------------------------- events */
document.addEventListener('click', async e => {
  const cl = e.target.closest('[data-close]');
  if (cl) { const w = cl.dataset.close; if (w === 'modal') closeModal(); else if (w === 'drawer') closeDrawer(); else if (w === 'palette') closePalette(); if (!cl.dataset.act) return; }
  const pal = e.target.closest('[data-pal]');
  if (pal) { palRun(Number(pal.dataset.pal)); return; }
  const a = e.target.closest('[data-act]');
  if (!a) {
    const card = e.target.closest('.tw');
    if (card && !e.target.closest('button')) toggleSel(card.dataset.id);
    return;
  }
  const act = a.dataset.act, id = a.dataset.id;
  const ids = a.dataset.ids ? a.dataset.ids.split(',') : [];
  switch (act) {
    case 'nav': go(a.dataset.page); break;
    case 'refresh': recheck(); break;
    case 'palette': openPalette(); break;
    case 'more': a.closest('.tw').classList.toggle('open'); a.querySelector('svg').style.transform = a.closest('.tw').classList.contains('open') ? 'rotate(90deg)' : ''; break;
    case 'preview-ids': openPreview(ids); break;
    case 'apply-ids': startRun(ids, a.dataset.title || 'Applying'); break;
    case 'modal-apply': startRun(ids, 'Applying'); break;
    case 'queue-add': S.sel.add(id); paintQueue(); renderSuggestions(); toast('Added to the queue'); break;
    case 'queue-clear': S.sel.clear(); paintQueue(); $$('.tw.sel').forEach(c => c.classList.remove('sel')); renderSuggestions(); break;
    case 'risk': $$('#riskSeg button').forEach(b => b.classList.toggle('on', b === a)); applyFilter(); break;
    case 'sel-page': { const p = S.cat.pages.find(x => x.id === a.dataset.page); p.ids.filter(i => S.byId[i].risk === 'safe' && S.byId[i].kind === 'tweak' && S.states[i] !== 'on').forEach(i => S.sel.add(i)); $$('.tw').forEach(c => c.classList.toggle('sel', S.sel.has(c.dataset.id))); paintQueue(); break; }
    case 'kind': $$('#kindSeg button').forEach(b => b.classList.toggle('on', b === a)); paintChangeRows(); break;
    case 'more-rows': S.chShow += 600; paintChangeRows(); break;
    case 'log': openLog(a.dataset.name); break;
    case 'logf': $$('#logSeg button').forEach(b => b.classList.toggle('on', b === a)); paintLog(); break;
    case 'open-logs': api('/api/open-logs', {}); toast((S.diag && S.diag.mock) ? 'Test mode: nothing opened' : 'Opening the folder', 'info'); break;
    case 'copy-log': copyText((S.logLines || []).join('\n')); break;
    case 'ping-toggle': S.settings = await api('/api/settings', { PingOn: !S.settings.PingOn }); $$('[data-act="ping-toggle"]').forEach(b => b.classList.toggle('on', S.settings.PingOn)); renderPingControl(); toast(S.settings.PingOn ? 'Ping monitor on' : 'Ping monitor off', 'info'); break;
    case 'update-toggle': S.settings = await api('/api/settings', { UpdateCheck: !S.settings.UpdateCheck }); a.classList.toggle('on', S.settings.UpdateCheck); break;
    case 'motion-toggle': S.settings = await api('/api/settings', { Motion: !(S.settings.Motion !== false) }); a.classList.toggle('on', S.settings.Motion !== false); document.documentElement.dataset.motion = S.settings.Motion === false ? 'off' : 'on'; break;
    case 'update-now': await checkUpdate(true); break;
    case 'diag': { const d = await api('/api/diag'); $('#diagBox').innerHTML = Object.keys(d).map(k => k + ': <b>' + esc(d[k]) + '</b>').join('<br>'); break; }
    case 'quit': if (confirm('Stop Jiff Tweaks and close the window?')) { await api('/api/quit', {}); document.body.innerHTML = '<div class="empty" style="height:100vh"><div>Jiff Tweaks stopped. You can close this window.</div></div>'; } break;
    case 'confirm-yes': if (S.confirmResolve) { const r = S.confirmResolve; S.confirmResolve = null; S.modalResolve = null; r(true); } break;
    case 'confirm-no': closeModal(); break;
    case 'dw-cancel': a.disabled = true; a.textContent = 'Stopping after this step'; api('/api/cancel', {}); break;
    case 'dw-copy': copyText(runReport()); break;
    case 'reboot': if (confirm('Restart Windows in 20 seconds? Save your work first.')) { await api('/api/reboot', {}); toast('Restarting in 20 seconds', 'info'); } break;
  }
});

document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); $('#palette').classList.contains('show') ? closePalette() : openPalette(); return; }
  if ($('#palette').classList.contains('show')) {
    if (e.key === 'Escape') closePalette(); else if (e.key === 'ArrowDown') { e.preventDefault(); palMove(1); } else if (e.key === 'ArrowUp') { e.preventDefault(); palMove(-1); } else if (e.key === 'Enter') { e.preventDefault(); palRun(); }
    return;
  }
  if (e.key === 'Escape') { if ($('#modal').classList.contains('show')) closeModal(); else if ($('#drawer').classList.contains('show')) closeDrawer(); }
});
$('#palInput').addEventListener('input', e => paintPalette(e.target.value));
let rz; window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { if (S.page === 'live') paintLive(); }, 120); });

async function checkUpdate(manual) {
  const r = await api('/api/update');
  const el = $('#updRes');
  if (r.ok) {
    const msg = r.newer ? 'Version ' + r.latest + ' is available (you have ' + r.current + ').' : 'You have the latest version (' + r.current + ').';
    if (el) el.textContent = msg;
    if (r.newer) setBanner('update', '<div class="banner warn">' + ic('rocket') + '<div class="grow">' + esc(msg) + ' Run the one-line command again to get it.</div><button class="btn sm" data-act="nav" data-page="settings">Details</button></div>');
    else if (manual) toast(msg);
  } else if (manual) { if (el) el.textContent = 'Could not check: ' + (r.error || 'offline'); toast('Could not check for updates', 'bad'); }
}

/* ---------------------------------------------------------------- boot */
(async function boot() {
  const t0 = Date.now();
  'JIFF TWEAKS'.split('').forEach((c, i) => { const s = document.createElement('span'); s.textContent = c === ' ' ? '\u00a0' : c; s.style.animationDelay = (0.9 + i * 0.045) + 's'; $('#splashWord').appendChild(s); });
  try {
    const [cat, sys, st] = await Promise.all([api('/api/catalog'), api('/api/sys'), api('/api/settings')]);
    S.cat = cat; S.sys = sys; S.settings = st;
    cat.tweaks.forEach(t => S.byId[t.id] = t);
    document.documentElement.dataset.motion = st.Motion === false ? 'off' : 'on';
    $('#verLabel').textContent = 'v' + cat.version + ' | free';
    api('/api/diag').then(d => { S.diag = d; }).catch(() => { });
    const lv = await api('/api/live?hist=1');
    (lv.hist || []).forEach(pushHist); S.live = lv.sample; if (lv.sample) lastT = lv.sample.t;
    $('#adminRow').innerHTML = sys.admin ? ic('shield', 'sm') + '<span>Administrator</span>' : ic('alert', 'sm') + '<span>Not administrator</span>';
    if (!sys.admin) setBanner('admin', '<div class="banner warn">' + ic('alert') + '<div class="grow">Not running as administrator. You can look around, but tweaks need administrator rights.</div></div>');
    if (sys.mock) setBanner('mock', '<div class="banner warn">' + ic('info') + '<div class="grow">Test mode: this is not Windows, so settings are simulated.</div></div>');
    setConn(true);
    const page = (location.hash || '').slice(1);
    go(page && pageInfo(page).id === page ? page : 'cockpit');
    pollLive(); pollStates();
    resumeRun();
    if (st.UpdateCheck) checkUpdate(false).catch(() => { });
  } catch (e) {
    $('#view').innerHTML = '<div class="card empty">' + ic('alert') + '<div style="font-size:16px;color:var(--text)">Cannot reach the script</div><div>Start Jiff Tweaks again from PowerShell and use the window it opens.</div></div>';
  }
  const wait = Math.max(0, 1700 - (Date.now() - t0));
  setTimeout(() => $('#splash').classList.add('gone'), wait);
})();
'@

# ============================================================================
#  CHECKS, ADMINISTRATOR RIGHTS, LAUNCH
# ============================================================================
$jtMock = ($env:JT_MOCK -eq '1')
if ($PSVersionTable.PSVersion.Major -lt 5) { Write-Host 'Jiff Tweaks needs PowerShell 5.1 or newer.' -ForegroundColor Red; return }
if ($env:OS -ne 'Windows_NT' -and -not $jtMock) { Write-Host 'Jiff Tweaks is for Windows.' -ForegroundColor Red; return }

$jtIsAdmin = $jtMock
if (-not $jtMock) { try { $jtIsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch { $jtIsAdmin = $false } }
$jtWow = ((-not $jtMock) -and [Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess)

# options typed on the command line, passed on when the script starts itself again
$jtArgs = New-Object System.Collections.Generic.List[string]
foreach ($k in $PSBoundParameters.Keys) {
  if ($k -eq 'NoElevate') { continue }
  $v = $PSBoundParameters[$k]
  if ($v -is [switch]) { if ($v.IsPresent) { [void]$jtArgs.Add('-' + $k) }; continue }
  $t = (@($v) -join ',')
  if ($t -notmatch '^[A-Za-z0-9_.,:\-]+$') { Write-Host ('Ignoring an unusual value for -' + $k) -ForegroundColor Yellow; continue }
  [void]$jtArgs.Add('-' + $k); [void]$jtArgs.Add($t)
}

if (-not $Code -and ((-not $jtIsAdmin -and -not $NoElevate) -or $jtWow)) {
  # start this script again: as administrator, and/or as a 64-bit process
  $exe = 'powershell.exe'
  if ($jtWow) { $exe = Join-Path $env:windir 'sysnative\WindowsPowerShell\v1.0\powershell.exe' }
  if ($PSCommandPath) { $call = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"')) + $jtArgs }
  else { $call = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ('"[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; & ([scriptblock]::Create((irm ''' + $JT_SELF_URL + '''))) ' + ($jtArgs -join ' ') + '"')) }
  $env:JT_RELAUNCH = '1'
  if (-not $jtIsAdmin) { Write-Host 'Asking for administrator rights ...' -ForegroundColor Cyan }
  try {
    if ($jtIsAdmin) { Start-Process -FilePath $exe -ArgumentList $call -Wait -NoNewWindow }
    else { Start-Process -FilePath $exe -ArgumentList $call -Verb RunAs }
  } catch { Write-Host 'Could not start with administrator rights. Open PowerShell as Administrator and run it again.' -ForegroundColor Red }
  return
}
if (-not $jtIsAdmin -and -not $Code) { Write-Host 'Run PowerShell as Administrator to use Jiff Tweaks.' -ForegroundColor Red; return }

# the files the window needs, as bytes
$jtTypes = @{ 'index.html' = 'text/html; charset=utf-8'; 'app.css' = 'text/css; charset=utf-8' }
$jtAssets = @{}
foreach ($n in @($JT_UI.Keys)) {
  $type = $jtTypes[$n]; if (-not $type) { $type = 'application/javascript; charset=utf-8' }
  $jtAssets[$n] = @{ Type = $type; Bytes = [Text.Encoding]::UTF8.GetBytes([string]$JT_UI[$n]) }
}

if ($PSCommandPath) { $env:JT_SCRIPTDIR = Split-Path -Parent $PSCommandPath } else { $env:JT_SCRIPTDIR = (Get-Location).Path }
$jtCli = [bool]($Preset -or $Apply -or $Preview -or $Code -or $List -or $Status -or $Restore -or $NoGui -or $SelfTest)
$jtOpts = @{ Preset = $Preset; Apply = $Apply; Preview = $Preview; Code = $Code; List = [bool]$List; Status = [bool]$Status; Restore = [bool]$Restore; SelfTest = [bool]$SelfTest; Yes = [bool]$Yes; NoOpen = [bool]$NoOpen; Browser = [bool]$Browser; Port = $Port }

# everything runs inside a module, so nothing it defines is left behind in your PowerShell session
try {
  [void](New-Module -Name JiffTweaks -ArgumentList $JT_ENGINE, $JT_SERVER, $jtAssets, $jtOpts, $jtCli -ScriptBlock {
      param($EngineSrc, $ServerSrc, $Assets, $O, $Cli)
      Export-ModuleMember -Function @()
      . ([scriptblock]::Create($EngineSrc))
      . ([scriptblock]::Create($ServerSrc))
      if ($Cli) { Invoke-JiffCli $O }
      else { Start-JiffApp -Src @{ Engine = $EngineSrc; Server = $ServerSrc } -Assets $Assets -NoBrowser:$O.NoOpen -ForceBrowser:$O.Browser -Port $O.Port }
    })
} catch {
  Write-Host ''
  Write-Host ('Jiff Tweaks stopped: ' + $_.Exception.Message) -ForegroundColor Red
  if ($PSCommandPath -or $env:JT_RELAUNCH) { [void](Read-Host 'Press Enter to close') }
} finally {
  Remove-Module JiffTweaks -ErrorAction SilentlyContinue
}
if (-not $PSCommandPath) { foreach ($n in 'JT_ENGINE', 'JT_SERVER', 'JT_UI', 'JT_SELF_URL', 'jtMock', 'jtIsAdmin', 'jtWow', 'jtArgs', 'jtAssets', 'jtTypes', 'jtCli', 'jtOpts') { Remove-Variable -Name $n -ErrorAction SilentlyContinue } }
