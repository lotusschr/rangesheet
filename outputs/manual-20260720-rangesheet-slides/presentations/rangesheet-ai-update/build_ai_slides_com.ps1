$ErrorActionPreference = "Stop"

$source = "C:\Users\TH90383638\OneDrive - Lotus's\Microsoft Copilot Chat Files\Range Sheet Streamlit Architechture_share.pptx"
$working = "C:\Users\TH90383638\OneDrive - Lotus's\Documents\GitHub\rangesheet\outputs\manual-20260720-rangesheet-slides\presentations\rangesheet-ai-update\output\rangesheet_streamlit_architecture_ai_working_copy.pptx"
$output = "C:\Users\TH90383638\OneDrive - Lotus's\Documents\GitHub\rangesheet\outputs\manual-20260720-rangesheet-slides\presentations\rangesheet-ai-update\output\rangesheet_streamlit_architecture_ai_updated_openable.pptx"

New-Item -ItemType Directory -Force -Path (Split-Path -Parent $output) | Out-Null
Copy-Item -LiteralPath $source -Destination $working -Force
if (Test-Path -LiteralPath $output) { Remove-Item -LiteralPath $output -Force }

function ColorRgb([int]$r, [int]$g, [int]$b) {
    return $r + ($g * 256) + ($b * 65536)
}

function Add-Rect($slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$fillHex, [string]$lineHex = "", [bool]$round = $false) {
    $shapeType = 1
    if ($round) { $shapeType = 5 }
    $shape = $slide.Shapes.AddShape($shapeType, $x, $y, $w, $h)
    $shape.Fill.ForeColor.RGB = Convert-HexColor $fillHex
    $shape.Fill.Visible = -1
    if ($lineHex -eq "") {
        $shape.Line.Visible = 0
    } else {
        $shape.Line.Visible = -1
        $shape.Line.ForeColor.RGB = Convert-HexColor $lineHex
        $shape.Line.Weight = 0.75
    }
    return $shape
}

function Convert-HexColor([string]$hex) {
    $h = $hex.TrimStart("#")
    $r = [Convert]::ToInt32($h.Substring(0,2), 16)
    $g = [Convert]::ToInt32($h.Substring(2,2), 16)
    $b = [Convert]::ToInt32($h.Substring(4,2), 16)
    return ColorRgb $r $g $b
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

function Add-AISlide($presentation, [string]$kicker, [string]$title, [string]$subtitle, [object[]]$cards, [string]$note) {
    $blankLayout = 12
    $slide = $presentation.Slides.Add($presentation.Slides.Count + 1, $blankLayout)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.ForeColor.RGB = Convert-HexColor "F7FAF9"
    Add-Rect $slide 0 0 16 540 "2BBFA4" "" $false | Out-Null
    Add-Text $slide 38 30 190 24 $kicker 11 "2BBFA4" $true | Out-Null
    Add-Text $slide 38 62 690 46 $title 25 "111827" $true | Out-Null
    Add-Text $slide 40 122 690 40 $subtitle 13 "4B5563" $false | Out-Null

    $x0 = 48
    $y0 = 190
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

$deck = $null
$ppt = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = -1
    $deck = $ppt.Presentations.Open($working, $false, $false, $true)

    Add-AISlide $deck "AI IN DEVELOPMENT" "AI helped make a complex real-data web app feasible within the timeline." "The project moved beyond classroom examples: large CSV files, connected filters, editable grids, business rules, reports, roles, audit logs, and performance constraints had to work together." @(
        @("Time constraint", "Manual coding alone would not have been enough to build and debug all requested features in the short development window.", "2BBFA4"),
        @("Real business data", "The app handled large operational datasets where small logic errors could affect tables, reports, and user decisions.", "F6D975"),
        @("Higher complexity", "Changes in one cell needed to update status, cluster summaries, range architecture, autosave, chatbot behavior, and report output.", "3B82F6"),
        @("Unexpected issues", "Performance bottlenecks, column-name mismatches, stale state, and UI reruns appeared during actual use.", "EF4444"),
        @("AI as support", "AI was used to explain errors, suggest fixes, compare approaches, and speed up iteration.", "8B5CF6"),
        @("Human judgment", "Requirements, business rules, validation, and final decisions still had to be checked by the developer.", "10B981")
    ) "Key message: AI accelerated learning and delivery, but the logic was still guided by business requirements."

    Add-AISlide $deck "AI WORKFLOW" "AI supported coding, debugging, and performance improvement across the application." "Instead of using AI as a one-click generator, it was used as a development partner during build-test-fix cycles." @(
        @("Code writing", "Generated implementation drafts for Streamlit UI, grid behavior, chatbot logic, report sections, and export workflows.", "2BBFA4"),
        @("Debugging", "Helped trace errors such as KeyError, undefined variables, stale session values, and dropdown state not persisting.", "EF4444"),
        @("Business logic mapping", "Translated user actions such as Delete Some, Delete All, New Some, and NewNew into consistent table and report calculations.", "F6D975"),
        @("Performance thinking", "Suggested caching, lighter recalculation, reduced reruns, and precomputed summaries to improve responsiveness.", "3B82F6"),
        @("UX iteration", "Helped refine chatbot, login, report preview, export buttons, pinned tables, and full-screen views based on feedback.", "8B5CF6"),
        @("Learning loop", "Each AI suggestion became a chance to understand why the code behaved that way, not just copy the answer.", "10B981")
    ) ""

    Add-AISlide $deck "LEARNING OUTCOME" "The biggest learning was connecting code behavior with business meaning." "The project required both technical debugging and an understanding of how range decisions affect item movement, planograms, and management reporting." @(
        @("System thinking", "Learned to trace data from source files through filters, grids, status logic, cluster tables, and reports.", "2BBFA4"),
        @("Data quality awareness", "Learned why missing values, different column names, and inconsistent formats can break downstream logic.", "F6D975"),
        @("State management", "Learned how autosave, refresh, role access, and edited dropdown states must remain consistent across pages and DGs.", "3B82F6"),
        @("Risk logic", "Learned how business rules like Top 10% best seller and Tail 10% lowest seller can guide warnings and review points.", "EF4444"),
        @("Communication", "Learned to convert detailed app activity into manager-friendly reports and audit summaries.", "8B5CF6"),
        @("Practical coding", "Learned by debugging real problems that were not fully predictable at the design stage.", "10B981")
    ) ""

    Add-AISlide $deck "VALUE CREATED" "The final app turns range editing into a connected review workflow, not just a spreadsheet screen." "Users can edit range decisions, see impact, ask the chatbot, export outputs, and send management reports while keeping auditability." @(
        @("Connected tables", "Main table changes update status, cluster counts, range architecture, and report outputs.", "2BBFA4"),
        @("Decision support", "Warnings highlight risky actions such as deleting best sellers or adding lowest sellers.", "EF4444"),
        @("Faster review", "Chatbot answers common range questions and supports quick action commands.", "8B5CF6"),
        @("Manager summary", "Reports summarize DG changes, SKU impact, planogram impact, product movement, and sales risk.", "F6D975"),
        @("Governance", "Login roles and audit logs support controlled access and traceability.", "3B82F6"),
        @("Export readiness", "Range output and report files can be generated for business handoff.", "10B981")
    ) "Outcome: AI helped bridge the gap between limited time, unfamiliar technical areas, and real business complexity."

    $deck.SaveAs($output)
}
finally {
    if ($deck -ne $null) { $deck.Close() | Out-Null }
    if ($ppt -ne $null) { $ppt.Quit() | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Output $output
