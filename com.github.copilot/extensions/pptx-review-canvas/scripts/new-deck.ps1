# Creates a new presentation with one slide and saves it as .pptx through desktop PowerPoint (COM).
# Without -Template it starts from PowerPoint's own blank presentation; with one, from a copy of that template or deck
# (the template file itself is never changed). The one slide uses the first layout (the title slide), whose empty placeholders draw nothing.
# Never overwrites: exit code 3 when -Path already exists, unless -Reserved says the caller created it empty to hold the name.
# Kept ASCII-only on purpose: Windows PowerShell 5.1 reads a script without a BOM in the system code page.
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$Template = '',
    [switch]$Reserved
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}

$wasRunning = [bool](Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
$app = $null
$pres = $null
$oldSecurity = $null
$oldAlerts = $null
$failed = $false
$exists = $false

try {
    if (-not $Reserved) {
        if (Test-Path -LiteralPath $Path) { $exists = $true; throw "exists: $Path" }
    } elseif ((Get-Item -LiteralPath $Path).Length -ne 0) {
        $exists = $true; throw "not empty: $Path"
    }
    $app = New-Object -ComObject PowerPoint.Application
    $oldSecurity = $app.AutomationSecurity
    $oldAlerts = $app.DisplayAlerts
    $app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: never run macros
    $app.DisplayAlerts = 1        # ppAlertsNone
    if ($Template) {
        $pres = $app.Presentations.Open($Template, $true, $true, $false)   # read-only, untitled (a copy), no window
    } else {
        $pres = $app.Presentations.Add(0)   # msoFalse: no window
    }
    if ($pres.Slides.Count -eq 0) {
        [void]$pres.Slides.AddSlide(1, $pres.SlideMaster.CustomLayouts.Item(1))
    }
    $pres.SaveAs($Path, 24)   # ppSaveAsOpenXMLPresentation
    Write-Output "CREATED $($pres.Slides.Count)"
}
catch {
    $failed = $true
    $hresult = '0x{0:X8}' -f $_.Exception.HResult
    [Console]::Error.WriteLine("HRESULT=$hresult " + $_.Exception.Message + " (script line " + $_.InvocationInfo.ScriptLineNumber + ")")
}
finally {
    if ($pres) { try { $pres.Close() } catch {} }
    if ($app) {
        try { if ($null -ne $oldSecurity) { $app.AutomationSecurity = $oldSecurity } } catch {}
        try { if ($null -ne $oldAlerts) { $app.DisplayAlerts = $oldAlerts } } catch {}
        try { if (-not $wasRunning -and $app.Presentations.Count -eq 0) { $app.Quit() } } catch {}
    }
    foreach ($o in @($pres, $app)) {
        if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} }
    }
}

if ($exists) { exit 3 }
if ($failed) { exit 1 }
exit 0
