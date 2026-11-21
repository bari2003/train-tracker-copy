# ==============================================================================
# Generate_Simple_PPT.ps1 — Clean, Easy-to-Explain PowerPoint (.ppt)
# Uses plain, accessible language and structured, high-contrast slides
# ==============================================================================

$pptApp = New-Object -ComObject PowerPoint.Application
$pres = $pptApp.Presentations.Add(-1)

# Set 16:9 Widescreen (960pt x 540pt)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function Get-RGB([int]$r, [int]$g, [int]$b) {
    return ($r + ($g * 256) + ($b * 65536))
}

# Theme Palette (Professional Dark Navy + Crisp White & Colored Accents)
$RGB_BG        = Get-RGB 15 23 42      # #0F172A (Navy Dark)
$RGB_CARD      = Get-RGB 30 41 59      # #1E293B (Card Fill)
$RGB_BORDER    = Get-RGB 71 85 105     # #475569 (Border)
$RGB_ACCENT    = Get-RGB 255 75 110    # Coral Red (Key terms)
$RGB_CYAN      = Get-RGB 6 182 212     # Cyan (Headers / Accents)
$RGB_GREEN     = Get-RGB 34 197 94     # Green (Solutions / Highlights)
$RGB_WHITE     = Get-RGB 255 255 255   # White (Main text)
$RGB_BODY      = Get-RGB 226 232 240   # #E2E8F0 (High readability text)
$RGB_MUTED     = Get-RGB 148 163 184   # #94A3B8 (Subtitles)
$RGB_HEADER_BG = Get-RGB 30 58 138     # #1E3A8A (Table Header)

# Helper: Standard Header with Category Tag and Clean Title
function Add-SlideHeader($slide, [string]$tagText, [string]$titleText, [int]$slideNum) {
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = $RGB_BG

    # Small Category Pill
    $tag = $slide.Shapes.AddTextbox(1, 45, 18, 500, 20)
    $tag.TextFrame.MarginLeft = 0; $tag.TextFrame.MarginTop = 0
    $tag.TextFrame.TextRange.Text = $tagText.ToUpper()
    $tag.TextFrame.TextRange.Font.Name = "Calibri"
    $tag.TextFrame.TextRange.Font.Size = 11
    $tag.TextFrame.TextRange.Font.Bold = -1
    $tag.TextFrame.TextRange.Font.Color.RGB = $RGB_ACCENT

    # Main Slide Title
    $title = $slide.Shapes.AddTextbox(1, 45, 38, 870, 44)
    $title.TextFrame.MarginLeft = 0; $title.TextFrame.MarginTop = 0
    $title.TextFrame.TextRange.Text = $titleText
    $title.TextFrame.TextRange.Font.Name = "Calibri"
    $title.TextFrame.TextRange.Font.Size = 24
    $title.TextFrame.TextRange.Font.Bold = -1
    $title.TextFrame.TextRange.Font.Color.RGB = $RGB_WHITE

    # Divider Line
    $line = $slide.Shapes.AddLine(45, 86, 915, 86)
    $line.Line.ForeColor.RGB = $RGB_BORDER
    $line.Line.Weight = 1

    # Footer Slide Number
    $ftr = $slide.Shapes.AddTextbox(1, 650, 508, 265, 20)
    $ftr.TextFrame.TextRange.Text = "Slide $slideNum of 10  |  Train Tracker"
    $ftr.TextFrame.TextRange.Font.Name = "Calibri"
    $ftr.TextFrame.TextRange.Font.Size = 10
    $ftr.TextFrame.TextRange.Font.Color.RGB = $RGB_MUTED
    $ftr.TextFrame.TextRange.ParagraphFormat.Alignment = 3 # Right
}

# Helper: Content Card Box
function Add-Card($slide, [float]$x, [float]$y, [float]$w, [float]$h, [string]$icon, [string]$title, [string]$bodyText) {
    $card = $slide.Shapes.AddShape(1, $x, $y, $w, $h) # Rectangle
    $card.Fill.Solid()
    $card.Fill.ForeColor.RGB = $RGB_CARD
    $card.Line.ForeColor.RGB = $RGB_BORDER
    $card.Line.Weight = 1

    # Card Title
    $hdr = $slide.Shapes.AddTextbox(1, $x + 16, $y + 12, $w - 32, 30)
    $hdr.TextFrame.MarginLeft = 0; $hdr.TextFrame.MarginTop = 0
    $hdr.TextFrame.TextRange.Text = "$icon  $title"
    $hdr.TextFrame.TextRange.Font.Name = "Calibri"
    $hdr.TextFrame.TextRange.Font.Size = 14
    $hdr.TextFrame.TextRange.Font.Bold = -1
    $hdr.TextFrame.TextRange.Font.Color.RGB = $RGB_CYAN

    # Card Body
    $txt = $slide.Shapes.AddTextbox(1, $x + 16, $y + 45, $w - 32, $h - 52)
    $txt.TextFrame.MarginLeft = 0; $txt.TextFrame.MarginTop = 0
    $txt.TextFrame.WordWrap = -1
    $txt.TextFrame.TextRange.Text = $bodyText
    $txt.TextFrame.TextRange.Font.Name = "Calibri"
    $txt.TextFrame.TextRange.Font.Size = 11.5
    $txt.TextFrame.TextRange.Font.Color.RGB = $RGB_BODY
    $txt.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3
}

# ==============================================================================
# SLIDE 1: Title & Project Overview
# ==============================================================================
Write-Host "Building Slide 1..."
$s1 = $pres.Slides.Add(1, 12)
$s1.FollowMasterBackground = 0
$s1.Background.Fill.Solid()
$s1.Background.Fill.ForeColor.RGB = $RGB_BG

$hero1 = $s1.Shapes.AddShape(1, 55, 45, 850, 450)
$hero1.Fill.Solid()
$hero1.Fill.ForeColor.RGB = $RGB_CARD
$hero1.Line.ForeColor.RGB = $RGB_BORDER
$hero1.Line.Weight = 1.5

$b1 = $s1.Shapes.AddTextbox(1, 95, 80, 500, 24)
$b1.TextFrame.TextRange.Text = "FINAL YEAR / CAPSTONE PROJECT PRESENTATION"
$b1.TextFrame.TextRange.Font.Name = "Calibri"
$b1.TextFrame.TextRange.Font.Size = 11
$b1.TextFrame.TextRange.Font.Bold = -1
$b1.TextFrame.TextRange.Font.Color.RGB = $RGB_ACCENT

$t1 = $s1.Shapes.AddTextbox(1, 95, 115, 770, 75)
$t1.TextFrame.TextRange.Text = "Real-Time Train Movement Tracker"
$t1.TextFrame.TextRange.Font.Name = "Calibri"
$t1.TextFrame.TextRange.Font.Size = 34
$t1.TextFrame.TextRange.Font.Bold = -1
$t1.TextFrame.TextRange.Font.Color.RGB = $RGB_WHITE

$sub1 = $s1.Shapes.AddTextbox(1, 95, 195, 770, 35)
$sub1.TextFrame.TextRange.Text = "Tracking Indian Railways Trains Live on Real Map Tracks"
$sub1.TextFrame.TextRange.Font.Name = "Calibri"
$sub1.TextFrame.TextRange.Font.Size = 16
$sub1.TextFrame.TextRange.Font.Color.RGB = $RGB_CYAN

$m1 = $s1.Shapes.AddTextbox(1, 95, 250, 770, 210)
$m1.TextFrame.WordWrap = -1
$m1.TextFrame.TextRange.Text = "• What it does: Shows live moving trains on a visual map with live speed, delays, and exact stations.`n`n• Author: Abdul Bari`n`n• Technologies Used: Java, Spring Boot, MySQL, Leaflet Maps, WebSockets`n`n• Live Demo Link: https://national-quad-alumni-fastest.trycloudflare.com"
$m1.TextFrame.TextRange.Font.Name = "Calibri"
$m1.TextFrame.TextRange.Font.Size = 13
$m1.TextFrame.TextRange.Font.Color.RGB = $RGB_BODY
$m1.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3

# ==============================================================================
# SLIDE 2: The Problem & Our Solution
# ==============================================================================
Write-Host "Building Slide 2..."
$s2 = $pres.Slides.Add(2, 12)
Add-SlideHeader $s2 "Why We Built This Project" "The Problem vs. Our Solution" 2

$tbl2 = $s2.Shapes.AddTable(4, 2, 45, 110, 870, 370)
$tbl2.Table.Columns.Item(1).Width = 435
$tbl2.Table.Columns.Item(2).Width = 435

$h2 = @("❌ What Traditional Apps Do (NTES / SMS)", "✅ What Our Train Tracker Does")
for ($c = 1; $c -le 2; $c++) {
    $cell = $tbl2.Table.Cell(1, $c)
    $cell.Shape.Fill.Solid()
    $cell.Shape.Fill.ForeColor.RGB = $RGB_HEADER_BG
    $cell.Shape.TextFrame.TextRange.Text = $h2[$c-1]
    $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
    $cell.Shape.TextFrame.TextRange.Font.Size = 13
    $cell.Shape.TextFrame.TextRange.Font.Bold = -1
    $cell.Shape.TextFrame.TextRange.Font.Color.RGB = if ($c -eq 1) { $RGB_ACCENT } else { $RGB_GREEN }
}

$rows2 = @(
    @("Only Text Updates:`nShows boring lines like 'Departed Surat at 2:15 PM'. You cannot see where the train is on a map.", "Visual Map on Real Tracks:`nRenders a moving train marker directly on top of real railway tracks, curves, and switches."),
    @("Slow Updates:`nInformation only updates when a train arrives at a large station. Between stations, you are blind.", "Live Continuous Streaming:`nPushes speed, delay, and position updates to your screen in real time without refreshing the page."),
    @("Website Freezes Under Heavy Traffic:`nMillions of users hitting 'Refresh' repeatedly crashes the railway servers.", "Smart Live Push:`nThe server automatically sends updates only to users watching that specific train, saving bandwidth.")
)

for ($r = 0; $r -lt 3; $r++) {
    for ($c = 1; $c -le 2; $c++) {
        $cell = $tbl2.Table.Cell($r + 2, $c)
        $cell.Shape.Fill.Solid()
        $cell.Shape.Fill.ForeColor.RGB = $RGB_CARD
        $cell.Shape.TextFrame.TextRange.Text = $rows2[$r][$c-1]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
        $cell.Shape.TextFrame.TextRange.Font.Size = 11.5
        $cell.Shape.TextFrame.TextRange.Font.Color.RGB = if ($c -eq 1) { $RGB_MUTED } else { $RGB_WHITE }
    }
}

# ==============================================================================
# SLIDE 3: How It Works (Simple 3 Steps)
# ==============================================================================
Write-Host "Building Slide 3..."
$s3 = $pres.Slides.Add(3, 12)
Add-SlideHeader $s3 "How It Works" "The Simple 3-Step Process" 3

Add-Card $s3 45 110 270 370 "1️⃣" "Step 1: Get Live Data" `
"• The backend calls the RailRadar API.`n`n• It fetches live train details: train speed, delay in minutes, departure station, and next stop.`n`n• It also downloads the real railway route with all station coordinates."

Add-Card $s3 345 110 270 370 "2️⃣" "Step 2: Calculate Spot" `
"• The Java backend calculates exactly where the train is along the track line.`n`n• It snaps the coordinates onto the physical rail line so the train never floats off the track.`n`n• It saves the latest status in MySQL."

Add-Card $s3 645 110 270 370 "3️⃣" "Step 3: Show on Map" `
"• The web map displays an animated train icon moving along the red railway line.`n`n• A live info panel displays: current speed, delay, and departure station.`n`n• WebSocket updates the position automatically without page reload."

# ==============================================================================
# SLIDE 4: Key Features
# ==============================================================================
Write-Host "Building Slide 4..."
$s4 = $pres.Slides.Add(4, 12)
Add-SlideHeader $s4 "Key Features" "What the User Experiences on Screen" 4

Add-Card $s4 45 110 420 175 "🚆" "Animated Train Marker" `
"An animated train icon with a glowing radar pulse moves along the track. The icon automatically rotates to face the direction the train is traveling."

Add-Card $s4 495 110 420 175 "🛤️" "Real Railway Tracks Map" `
"Uses OpenRailwayMap tiles to display real physical rail tracks, railway stations, junctions, and signals instead of a plain empty road map."

Add-Card $s4 45 305 420 175 "⏱️" "Live Delay & Speed HUD" `
"Displays real-time speed (e.g. 85 km/h) and color-coded delay status: green for 'On Time' and red for '+18 min late'."

Add-Card $s4 495 305 420 175 "🔍" "One-Click Train Search" `
"Search any train by number (e.g., 12951 Mumbai Rajdhani). Clicking any train instantly loads its complete route line and starts live tracking."

# ==============================================================================
# SLIDE 5: Smart Fallback System
# ==============================================================================
Write-Host "Building Slide 5..."
$s5 = $pres.Slides.Add(5, 12)
Add-SlideHeader $s5 "System Reliability" "Smart Fallback: Why Our App Never Freezes" 5

$d5 = $s5.Shapes.AddTextbox(1, 45, 100, 870, 40)
$d5.TextFrame.TextRange.Text = "Free railway APIs allow only a small number of requests per month. If the limit is reached, other apps crash. Our app uses a smart dual-mode system:"
$d5.TextFrame.TextRange.Font.Name = "Calibri"
$d5.TextFrame.TextRange.Font.Size = 12
$d5.TextFrame.TextRange.Font.Color.RGB = $RGB_BODY

Add-Card $s5 45 150 420 330 "🟢" "Mode A: Live Official Data" `
"• When the API has quota, the app uses 100% official live data.`n`n• Directly reads real-time GPS locations, official platform numbers, and live delay reports.`n`n• Caches the full railway track so it is saved for future visits."

Add-Card $s5 495 150 420 330 "🛡️" "Mode B: Smart Simulation Fallback" `
"• If the external API quota runs out (HTTP 429) or goes offline, this backup mode starts immediately.`n`n• Smoothly moves the train along the loaded physical track at realistic speeds (75-90 km/h).`n`n• The app never shows a blank error, never disconnects, and keeps moving smoothly."

# ==============================================================================
# SLIDE 6: Accurate Positioning
# ==============================================================================
Write-Host "Building Slide 6..."
$s6 = $pres.Slides.Add(6, 12)
Add-SlideHeader $s6 "Accurate Positioning" "How We Keep the Train on Track (Simple Math)" 6

Add-Card $s6 45 110 420 175 "🧲" "1. Snap to Track" `
"GPS signals can drift slightly to the left or right. Our backend automatically detects drift and pulls the train's marker right onto the nearest railway track line."

Add-Card $s6 495 110 420 175 "📏" "2. Percentage Progress" `
"Between two stations, the system calculates the percentage of the journey completed (e.g. 500 km out of 1380 km) and places the train at that exact spot."

Add-Card $s6 45 305 420 175 "🔄" "3. Smooth Train Switching" `
"When you click between different trains (e.g. Mumbai Rajdhani to Bhopal Shatabdi), the connection stays open and updates instantly without any disconnect glitch."

Add-Card $s6 495 305 420 175 "📐" "4. Automatic Rotation" `
"The train icon knows the angle of the track curve and rotates automatically so it always points forward as it moves around bends and junctions."

# ==============================================================================
# SLIDE 7: Technology Stack
# ==============================================================================
Write-Host "Building Slide 7..."
$s7 = $pres.Slides.Add(7, 12)
Add-SlideHeader $s7 "Technology Stack" "Tools & Frameworks Used to Build the App" 7

Add-Card $s7 45 110 200 370 "🌐" "Frontend" `
"• HTML5 & CSS3`nClean dark-mode interface`n`n• JavaScript`nInteractive controls & search`n`n• Leaflet.js`nFast map rendering engine`n`n• OpenRailwayMap`nReal railway track tiles"

Add-Card $s7 265 110 200 370 "⚙️" "Backend" `
"• Java 27`nModern high-speed runtime`n`n• Spring Boot 3`nREST API web framework`n`n• Tomcat Server`nEmbedded web server`n`n• Jackson`nFast JSON data parser"

Add-Card $s7 485 110 200 370 "💾" "Database" `
"• MySQL 8.0`nStores trains & stations`n`n• LONGTEXT Storage`nStores large route geometries with 220+ stations`n`n• H2 In-Memory DB`nFor quick offline testing"

Add-Card $s7 705 110 200 370 "📡" "Live Stream" `
"• WebSocket (STOMP)`nLive two-way data streaming`n`n• SockJS`nBrowser fallback support`n`n• Cloudflare Tunnel`nFree public HTTPS domain`n`n• RailRadar API`nLive transit data feed"

# ==============================================================================
# SLIDE 8: Database Design
# ==============================================================================
Write-Host "Building Slide 8..."
$s8 = $pres.Slides.Add(8, 12)
Add-SlideHeader $s8 "Database Design" "What Information We Store in MySQL" 8

$tbl8 = $s8.Shapes.AddTable(5, 3, 45, 110, 870, 370)
$tbl8.Table.Columns.Item(1).Width = 200
$tbl8.Table.Columns.Item(2).Width = 335
$tbl8.Table.Columns.Item(3).Width = 335

$h8 = @("Table Name", "What It Stores", "Example Data")
for ($c = 1; $c -le 3; $c++) {
    $cell = $tbl8.Table.Cell(1, $c)
    $cell.Shape.Fill.Solid()
    $cell.Shape.Fill.ForeColor.RGB = $RGB_HEADER_BG
    $cell.Shape.TextFrame.TextRange.Text = $h8[$c-1]
    $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
    $cell.Shape.TextFrame.TextRange.Font.Size = 12.5
    $cell.Shape.TextFrame.TextRange.Font.Bold = -1
    $cell.Shape.TextFrame.TextRange.Font.Color.RGB = $RGB_CYAN
}

$rows8 = @(
    @("trains", "Train list, train numbers, and names.", "12951 - Mumbai Rajdhani Express`n12001 - Bhopal Shatabdi Express"),
    @("stations", "Station names, short codes, and GPS coordinates.", "NDLS (New Delhi), BCT (Mumbai Central),`nBPL (Bhopal), HWH (Howrah)"),
    @("routes", "The complete curved railway line geometry connecting all intermediate stops (74 KB+).", "Full detailed path of 220+ station coordinates from New Delhi to Mumbai."),
    @("live_train_states", "Latest speed, delay, and current position for instant display when clicked.", "Speed: 85 km/h, Delay: 0 min,`nStatus: Departed Vadodara (BRC)")
)

for ($r = 0; $r -lt 4; $r++) {
    for ($c = 1; $c -le 3; $c++) {
        $cell = $tbl8.Table.Cell($r + 2, $c)
        $cell.Shape.Fill.Solid()
        $cell.Shape.Fill.ForeColor.RGB = $RGB_CARD
        $cell.Shape.TextFrame.TextRange.Text = $rows8[$r][$c-1]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
        $cell.Shape.TextFrame.TextRange.Font.Size = 11.5
        $cell.Shape.TextFrame.TextRange.Font.Color.RGB = if ($c -eq 1) { $RGB_WHITE } else { $RGB_BODY }
    }
}

# ==============================================================================
# SLIDE 9: Real-World Benefits
# ==============================================================================
Write-Host "Building Slide 9..."
$s9 = $pres.Slides.Add(9, 12)
Add-SlideHeader $s9 "Real-World Impact" "Who Benefits From This System?" 9

Add-Card $s9 45 110 420 175 "👨‍👩‍👦" "1. Passengers & Families" `
"See exactly where a train is so you don't wait for hours at the station. Especially useful during winter fog delays or late-night arrivals for family safety."

Add-Card $s9 495 110 420 175 "📦" "2. Parcel & Freight Logistics" `
"Companies shipping goods on express parcel trains can track their cargo live on the map and estimate accurate delivery times."

Add-Card $s9 45 305 420 175 "🏢" "3. Station Platform Staff" `
"Helps railway managers control platform crowds by knowing the exact arrival times of incoming trains in advance."

Add-Card $s9 495 305 420 175 "🚨" "4. Emergency Monitoring" `
"Gives control rooms instant visual awareness of where trains are located during maintenance blocks, heavy rains, or rail diversions."

# ==============================================================================
# SLIDE 10: Future Scope & Conclusion
# ==============================================================================
Write-Host "Building Slide 10..."
$s10 = $pres.Slides.Add(10, 12)
Add-SlideHeader $s10 "Conclusion" "Future Scope & Summary" 10

Add-Card $s10 45 110 270 230 "📡" "1. Multi-Train Radar" `
"Show hundreds of trains moving simultaneously across an entire zone on one master screen."

Add-Card $s10 345 110 270 230 "🤖" "2. AI Delay Predictor" `
"Use Machine Learning to predict delays before they happen based on weather and route traffic."

Add-Card $s10 645 110 270 230 "📱" "3. Mobile App (PWA)" `
"Allow passengers inside trains to see track progress on their phones even with weak internet."

# Bottom Summary Card
$sumCard = $s10.Shapes.AddShape(1, 45, 360, 870, 120)
$sumCard.Fill.Solid()
$sumCard.Fill.ForeColor.RGB = Get-RGB 20 35 60
$sumCard.Line.ForeColor.RGB = $RGB_CYAN
$sumCard.Line.Weight = 1.5

$stxt = $s10.Shapes.AddTextbox(1, 65, 372, 830, 95)
$stxt.TextFrame.WordWrap = -1
$stxt.TextFrame.TextRange.Text = "FINAL SUMMARY`n`nTrain Tracker turns complicated, raw railway data into an interactive, visual live map. It is easy to use, updates in real time, and never crashes even if external APIs run out of quota."
$stxt.TextFrame.TextRange.Font.Name = "Calibri"
$stxt.TextFrame.TextRange.Font.Size = 12.5
$stxt.TextFrame.TextRange.Font.Bold = -1
$stxt.TextFrame.TextRange.Font.Color.RGB = $RGB_WHITE
$stxt.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3

# ==============================================================================
# Save exclusively as .ppt format
# ==============================================================================
$outPpt = "C:\Users\ABDUL BARI\Documents\train-tracker - Copy\Real_Time_Train_Movement_Tracker.ppt"

if (Test-Path $outPpt) {
    Remove-Item -Path $outPpt -Force
}

Write-Host "Saving presentation to $outPpt..."
$pres.SaveAs($outPpt, 1) # 1 = ppSaveAsPresentation (.ppt format)

$pres.Close()
$pptApp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) | Out-Null

Write-Host "SUCCESS: Generated easy-to-explain .ppt presentation!" -ForegroundColor Green
