# Renders every slide of a presentation to PNG through desktop PowerPoint (COM).
# Writes s0001.png (slide) and t0001.png (thumbnail) for each slide, plus slides.json with the titles and speaker notes.
# Stdout: "PROGRESS <done> <total>" lines. A file named cancel.flag in OutDir stops the run between slides (exit code 2).
# Kept ASCII-only on purpose: Windows PowerShell 5.1 reads a script without a BOM in the system code page.
param(
    [Parameter(Mandatory)][string]$Pptx,
    [Parameter(Mandatory)][string]$OutDir,
    [int]$Width = 1600,
    [int]$ThumbWidth = 360
)
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
try { [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false) } catch {}

function Normalize-Text([string]$text) {
    if ($null -eq $text) { return '' }
    return ($text -replace "(\r\n|\r|\n|\v)", "`n").Trim()
}

function Short-Text([string]$text, [int]$max) {
    $t = (Normalize-Text $text) -replace "\s+", ' '
    if ($t.Length -gt $max) { return $t.Substring(0, $max) + '...' }
    return $t
}

# Walks the shapes of a slide (groups included) and records where each one is and what it says.
# Child coordinates of a group are already slide coordinates in the PowerPoint object model.
function Collect-Shapes($shapes, [int]$depth, $list) {
    foreach ($shape in $shapes) {
        if ($list.Count -ge 500) { return }
        try {
            if (-not $shape.Visible) { continue }
            $type = [int]$shape.Type
            $kind = 'shape'
            $text = ''
            $alt = ''
            if ($type -eq 6) { $kind = 'group' }
            elseif ($type -eq 13 -or $type -eq 28) { $kind = 'picture' }
            elseif ($shape.HasTable) { $kind = 'table' }
            elseif ($type -eq 9) { $kind = 'line' }
            if ($kind -ne 'group') {
                if ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
                    $text = Short-Text $shape.TextFrame.TextRange.Text 100
                    $kind = 'text'
                }
                if ($kind -eq 'picture') { $alt = Short-Text $shape.AlternativeText 100 }
            }
            $list.Add([ordered]@{
                name  = [string]$shape.Name
                kind  = $kind
                depth = $depth
                x     = [Math]::Round([double]$shape.Left, 1)
                y     = [Math]::Round([double]$shape.Top, 1)
                w     = [Math]::Round([double]$shape.Width, 1)
                h     = [Math]::Round([double]$shape.Height, 1)
                text  = $text
                alt   = $alt
            })
            if ($kind -eq 'group') { Collect-Shapes $shape.GroupItems ($depth + 1) $list }
        } catch {}
    }
}

$wasRunning = [bool](Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
$app = $null
$pres = $null
$oldSecurity = $null
$oldAlerts = $null
$cancelled = $false
$failed = $false

try {
    $app = New-Object -ComObject PowerPoint.Application
    $oldSecurity = $app.AutomationSecurity
    $oldAlerts = $app.DisplayAlerts
    $app.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: never run macros
    $app.DisplayAlerts = 1        # ppAlertsNone
    $pres = $app.Presentations.Open($Pptx, $true, $false, $false)   # read-only, untitled, no window

    $slideW = [double]$pres.PageSetup.SlideWidth
    $slideH = [double]$pres.PageSetup.SlideHeight
    $height = [int][Math]::Round($Width * $slideH / $slideW)
    $thumbHeight = [int][Math]::Round($ThumbWidth * $slideH / $slideW)
    $total = $pres.Slides.Count
    Write-Output "PROGRESS 0 $total"

    $cancelFlag = Join-Path $OutDir 'cancel.flag'
    $slides = New-Object System.Collections.Generic.List[object]
    $shapeMap = [ordered]@{}
    $shapeMs = 0
    $i = 0
    foreach ($slide in $pres.Slides) {
        if (Test-Path -LiteralPath $cancelFlag) { $cancelled = $true; break }
        $i++
        $title = ''
        try { if ($slide.Shapes.HasTitle) { $title = Normalize-Text $slide.Shapes.Title.TextFrame.TextRange.Text } } catch {}
        $notes = ''
        try {
            foreach ($shape in $slide.NotesPage.Shapes) {
                # msoPlaceholder = 14, ppPlaceholderBody = 2
                if ($shape.Type -eq 14 -and $shape.PlaceholderFormat.Type -eq 2 -and $shape.HasTextFrame) {
                    $notes = Normalize-Text $shape.TextFrame.TextRange.Text
                    break
                }
            }
        } catch {}
        $name = '{0:D4}' -f $i
        $slide.Export((Join-Path $OutDir "s$name.png"), 'PNG', $Width, $height)
        $slide.Export((Join-Path $OutDir "t$name.png"), 'PNG', $ThumbWidth, $thumbHeight)
        $clock = [Diagnostics.Stopwatch]::StartNew()
        $found = New-Object System.Collections.Generic.List[object]
        try { Collect-Shapes $slide.Shapes 0 $found } catch {}
        $shapeMap["$i"] = $found.ToArray()
        $shapeMs += $clock.ElapsedMilliseconds
        $slides.Add([ordered]@{ n = $i; title = $title; notes = $notes })
        Write-Output "PROGRESS $i $total"
    }
    if ($env:PPTX_VIEWER_TIMING) { [Console]::Error.WriteLine("shapes: $shapeMs ms") }

    if (-not $cancelled) {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        $shapeJson = ConvertTo-Json -InputObject $shapeMap -Depth 5
        [System.IO.File]::WriteAllText((Join-Path $OutDir 'shapes.json'), $shapeJson, $utf8)
        $payload = [ordered]@{
            slideWidth  = $slideW
            slideHeight = $slideH
            imageWidth  = $Width
            imageHeight = $height
            thumbWidth  = $ThumbWidth
            thumbHeight = $thumbHeight
            slides      = $slides.ToArray()   # @($slides) fails in Windows PowerShell 5.1 for a List of ordered dictionaries
        }
        $json = ConvertTo-Json -InputObject $payload -Depth 6
        [System.IO.File]::WriteAllText((Join-Path $OutDir 'slides.json'), $json, $utf8)
    }
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

if ($failed) { exit 1 }
if ($cancelled) { Write-Output 'CANCELLED'; exit 2 }
exit 0
