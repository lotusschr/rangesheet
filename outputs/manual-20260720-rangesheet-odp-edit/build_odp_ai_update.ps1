$ErrorActionPreference = "Stop"

$source = "C:\Users\TH90383638\AppData\Roaming\Microsoft\Windows\Network Shortcuts\RANGESHEET PLATFORM present2.odp"
$outDir = "C:\Users\TH90383638\OneDrive - Lotus's\Documents\GitHub\rangesheet\outputs\manual-20260720-rangesheet-odp-edit\output"
$pptxOut = Join-Path $outDir "rangesheet_platform_present2_ai_updated.pptx"
$odpOut = Join-Path $outDir "rangesheet_platform_present2_ai_updated.odp"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null
if (Test-Path -LiteralPath $pptxOut) { Remove-Item -LiteralPath $pptxOut -Force }
if (Test-Path -LiteralPath $odpOut) { Remove-Item -LiteralPath $odpOut -Force }

function ColorRgb([int]$r, [int]$g, [int]$b) {
    return $r + ($g * 256) + ($b * 65536)
}

function Convert-HexColor([string]$hex) {
    $h = $hex.TrimStart("#")
    $r = [Convert]::ToInt32($h.Substring(0,2), 16)
    $g = [Convert]::ToInt32($h.Substring(2,2), 16)
    $b = [Convert]::ToInt32($h.Substring(4,2), 16)
    return ColorRgb $r $g $b
}

function Add-Rect($slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$fillHex, [string]$lineHex = "", [bool]$round = $false) {
    $shapeType = 1
    if ($round) { $shapeType = 5 }
    $shape = $slide.Shapes.AddShape($shapeType, $x, $y, $w, $h)
    $shape.Fill.Visible = -1
    $shape.Fill.ForeColor.RGB = Convert-HexColor $fillHex
    if ($lineHex -eq "") {
        $shape.Line.Visible = 0
    } else {
        $shape.Line.Visible = -1
        $shape.Line.ForeColor.RGB = Convert-HexColor $lineHex
        $shape.Line.Weight = 0.75
    }
    return $shape
}

function Add-Text($slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$text, [double]$size, [string]$colorHex, [bool]$bold = $false) {
    $box = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $box.TextFrame.TextRange.Text = $text
    $box.TextFrame.MarginLeft = 0
    $box.TextFrame.MarginRight = 0
    $box.TextFrame.MarginTop = 0
    $box.TextFrame.MarginBottom = 0
    $box.TextFrame.WordWrap = -1
    $font = $box.TextFrame.TextRange.Font
    $font.Name = "Aptos"
    $font.Size = $size
    $font.Color.RGB = Convert-HexColor $colorHex
    if ($bold) { $font.Bold = -1 } else { $font.Bold = 0 }
    return $box
}

function Add-Card($slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$accent, [string]$head, [string]$body) {
    Add-Rect $slide $x $y $w $h "FFFFFF" "DDE7E3" $true | Out-Null
    Add-Rect $slide $x $y 5 $h $accent "" $false | Out-Null
    Add-Text $slide ($x + 16) ($y + 14) ($w - 28) 22 $head 13 "111827" $true | Out-Null
    Add-Text $slide ($x + 16) ($y + 42) ($w - 28) ($h - 50) $body 10.5 "374151" $false | Out-Null
}

function Add-AISlide($presentation, [int]$index, [string]$kicker, [string]$title, [string]$subtitle, [object[]]$cards, [string]$note) {
    $blankLayout = 12
    $slide = $presentation.Slides.Add($index, $blankLayout)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.ForeColor.RGB = Convert-HexColor "F7FAF9"

    Add-Rect $slide 0 0 16 540 "2BBFA4" "" $false | Out-Null
    Add-Text $slide 38 30 210 24 $kicker 11 "2BBFA4" $true | Out-Null
    Add-Text $slide 38 62 690 50 $title 24 "111827" $true | Out-Null
    Add-Text $slide 40 122 690 42 $subtitle 12.5 "4B5563" $false | Out-Null

    $x0 = 48
    $y0 = 188
    $cardW = 205
    $cardH = 92
    $gapX = 18
    $gapY = 24
    for ($i = 0; $i -lt $cards.Count; $i++) {
        $col = $i % 3
        $row = [Math]::Floor($i / 3)
        $x = $x0 + ($col * ($cardW + $gapX))
        $y = $y0 + ($row * ($cardH + $gapY))
        Add-Card $slide $x $y $cardW $cardH $cards[$i][2] $cards[$i][0] $cards[$i][1]
    }

    if ($note -ne "") {
        Add-Rect $slide 48 488 650 34 "E8F8F5" "" $true | Out-Null
        Add-Text $slide 62 497 620 18 $note 10 "6B7280" $false | Out-Null
    }
}

$ppt = $null
$deck = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $deck = $ppt.Presentations.Open($source, $true, $false, $true)

    $insertAt = 14

    Add-AISlide $deck $insertAt "AI IN DEVELOPMENT" "AI made delivery possible within limited time and real business complexity." "Manual coding alone would not have been enough because the app handles real range data, large files, connected tables, business rules, reports, roles, audit logs, and exports." @(
        @("Limited timeline", "The development window was short, while the system scope kept expanding from table review into chatbot, reports, exports, and login roles.", "2BBFA4"),
        @("Real data scale", "Large CSV and parquet data created performance issues that are not visible in small classroom examples.", "F6D975"),
        @("Business logic", "Every dropdown change had to connect to status, cluster table, range architecture, autosave, report, and audit log.", "3B82F6"),
        @("Unexpected problems", "Issues appeared during real use: slow reruns, missing columns, stale state, wrong counts, and UI consistency problems.", "EF4444"),
        @("AI support", "AI helped draft code, debug errors, compare approaches, and explain why problems happened.", "8B5CF6"),
        @("Human control", "Final logic still came from business requirements and manual validation against the expected output.", "10B981")
    ) "AI was not replacing the developer. It helped shorten the learning and debugging cycle."
    $insertAt++

    Add-AISlide $deck $insertAt "AI WORKFLOW" "AI became a development partner during build, debug, and improve cycles." "The project used AI for practical engineering support: understand the issue, propose a fix, test the result, then adjust until the app behavior matched the business workflow." @(
        @("Code generation", "Helped create Streamlit UI, grid behavior, chatbot intent logic, reports, exports, and validation flow.", "2BBFA4"),
        @("Debugging", "Helped trace errors such as KeyError, import problems, dropdown state bugs, and report calculation mismatches.", "EF4444"),
        @("Performance", "Suggested caching, precomputed summaries, lighter recalculation, and reduced reruns to make the app faster.", "3B82F6"),
        @("Business mapping", "Translated actions such as Delete Some, Delete All, New Some, and NewNew into consistent output rules.", "F6D975"),
        @("UX iteration", "Helped refine login, chatbot, report preview, pinned columns, export files, and warning messages.", "8B5CF6"),
        @("Learning loop", "Each fix became a chance to learn the system deeper, not only copy code.", "10B981")
    ) ""
    $insertAt++

    Add-AISlide $deck $insertAt "KEY LEARNING" "The project taught how code, data, and business decisions depend on each other." "This was more than a programming task. It required understanding how one range decision affects products, planograms, clusters, reports, and decision risk." @(
        @("System thinking", "Learned to trace data from raw files through filters, editable grid, status calculation, and final report.", "2BBFA4"),
        @("State management", "Learned why autosave, refresh, switching DG, and role access must keep the same latest decision.", "3B82F6"),
        @("Data quality", "Learned how column names, blanks, numeric formats, and missing data can break business logic.", "F6D975"),
        @("Risk review", "Built warnings for Top 10% best seller deletion and Tail 10% lowest seller addition.", "EF4444"),
        @("Communication", "Converted detailed actions into manager-friendly report, audit log, and chatbot answers.", "8B5CF6"),
        @("Practical debugging", "Many issues appeared only after using real data and testing realistic workflows.", "10B981")
    ) ""
    $insertAt++

    Add-AISlide $deck $insertAt "VALUE CREATED" "RangeSheet became a connected review platform instead of separate manual spreadsheets." "The final workflow links review actions to business impact, warnings, reports, exports, and auditability." @(
        @("Connected output", "Dropdown edits update the main table, cluster table, range architecture, and report output.", "2BBFA4"),
        @("Decision support", "The app highlights risky changes before submission and records the confirmed action.", "EF4444"),
        @("Chatbot helper", "Users can ask common range questions and trigger supported item or planogram actions.", "8B5CF6"),
        @("Report ready", "Submitted DGs create manager-friendly summaries with SKU impact and sales risk review.", "F6D975"),
        @("Governance", "Login roles and audit logs help control access and track changes.", "3B82F6"),
        @("Business handoff", "Export files support downstream usage after review is complete.", "10B981")
    ) "The main value is consistency: one edit can flow through every related output."

    $deck.SaveAs($pptxOut, 24)
    try {
        $deck.SaveAs($odpOut, 35)
    } catch {
        Write-Output "ODP SaveAs skipped: $($_.Exception.Message)"
    }
}
finally {
    if ($deck -ne $null) { $deck.Close() | Out-Null }
    if ($ppt -ne $null) { $ppt.Quit() | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Output $pptxOut
if (Test-Path -LiteralPath $odpOut) { Write-Output $odpOut }
