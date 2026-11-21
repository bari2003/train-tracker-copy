# ==============================================================================
# Generate_PPT.ps1 — Native PowerPoint Generator via Microsoft Office COM (.ppt)
# ==============================================================================

$pptApp = New-Object -ComObject PowerPoint.Application
# msoTrue = -1, msoFalse = 0
$pres = $pptApp.Presentations.Add(-1)

# Set 16:9 Widescreen (960pt x 540pt)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function Get-RGB([int]$r, [int]$g, [int]$b) {
    return ($r + ($g * 256) + ($b * 65536))
}

$RGB_BG        = Get-RGB 15 23 42      # #0F172A
$RGB_CARD      = Get-RGB 30 41 59      # #1E293B
$RGB_BORDER    = Get-RGB 71 85 105     # #475569
$RGB_ACCENT    = Get-RGB 255 42 95     # #FF2A5F
$RGB_CYAN      = Get-RGB 6 182 212     # #06B6D4
$RGB_WHITE     = Get-RGB 255 255 255   # White
$RGB_BODY      = Get-RGB 226 232 240   # #E2E8F0 (Crisp Light Text)
$RGB_MUTED     = Get-RGB 148 163 184   # #94A3B8
$RGB_HEADER_BG = Get-RGB 30 58 138     # #1E3A8A

function Add-SlideChrome($slide, [string]$category, [string]$title, [int]$slideNum) {
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = $RGB_BG

    # Category Pill Tag
    $tag = $slide.Shapes.AddTextbox(1, 40, 20, 500, 22)
    $tag.TextFrame.MarginLeft = 0
    $tag.TextFrame.MarginTop = 0
    $tag.TextFrame.TextRange.Text = $category.ToUpper()
    $tag.TextFrame.TextRange.Font.Name = "Segoe UI"
    $tag.TextFrame.TextRange.Font.Size = 10
    $tag.TextFrame.TextRange.Font.Bold = -1
    $tag.TextFrame.TextRange.Font.Color.RGB = $RGB_ACCENT

    # Title
    $titleBox = $slide.Shapes.AddTextbox(1, 40, 42, 800, 46)
    $titleBox.TextFrame.MarginLeft = 0
    $titleBox.TextFrame.MarginTop = 0
    $titleBox.TextFrame.TextRange.Text = $title
    $titleBox.TextFrame.TextRange.Font.Name = "Segoe UI"
    $titleBox.TextFrame.TextRange.Font.Size = 22
    $titleBox.TextFrame.TextRange.Font.Bold = -1
    $titleBox.TextFrame.TextRange.Font.Color.RGB = $RGB_WHITE

    # Divider Line
    $line = $slide.Shapes.AddLine(40, 92, 920, 92)
    $line.Line.ForeColor.RGB = $RGB_BORDER
    $line.Line.Weight = 1

    # Footer
    $footer = $slide.Shapes.AddTextbox(1, 600, 505, 320, 24)
    $footer.TextFrame.TextRange.Text = "Slide $slideNum of 10  |  Train Tracker"
    $footer.TextFrame.TextRange.Font.Name = "Segoe UI"
    $footer.TextFrame.TextRange.Font.Size = 9.5
    $footer.TextFrame.TextRange.Font.Color.RGB = $RGB_MUTED
    $footer.TextFrame.TextRange.ParagraphFormat.Alignment = 3 # Right
}

function Add-Card($slide, [float]$x, [float]$y, [float]$w, [float]$h, [string]$icon, [string]$title, [string]$body) {
    $card = $slide.Shapes.AddShape(1, $x, $y, $w, $h) # 1 = msoShapeRectangle
    $card.Fill.Solid()
    $card.Fill.ForeColor.RGB = $RGB_CARD
    $card.Line.ForeColor.RGB = $RGB_BORDER
    $card.Line.Weight = 1

    # Card Title
    $hdr = $slide.Shapes.AddTextbox(1, $x + 14, $y + 12, $w - 28, 28)
    $hdr.TextFrame.MarginLeft = 0
    $hdr.TextFrame.MarginTop = 0
    $hdr.TextFrame.TextRange.Text = "$icon  $title"
    $hdr.TextFrame.TextRange.Font.Name = "Segoe UI"
    $hdr.TextFrame.TextRange.Font.Size = 13
    $hdr.TextFrame.TextRange.Font.Bold = -1
    $hdr.TextFrame.TextRange.Font.Color.RGB = $RGB_CYAN

    # Card Body
    $txt = $slide.Shapes.AddTextbox(1, $x + 14, $y + 44, $w - 28, $h - 52)
    $txt.TextFrame.MarginLeft = 0
    $txt.TextFrame.MarginTop = 0
    $txt.TextFrame.WordWrap = -1
    $txt.TextFrame.TextRange.Text = $body
    $txt.TextFrame.TextRange.Font.Name = "Segoe UI"
    $txt.TextFrame.TextRange.Font.Size = 10.5
    $txt.TextFrame.TextRange.Font.Color.RGB = $RGB_BODY
    $txt.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.25
}

# ------------------------------------------------------------------------------
# SLIDE 1: Title
# ------------------------------------------------------------------------------
$s1 = $pres.Slides.Add(1, 12)
$s1.FollowMasterBackground = 0
$s1.Background.Fill.Solid()
$s1.Background.Fill.ForeColor.RGB = $RGB_BG

$hero = $s1.Shapes.AddShape(1, 50, 40, 860, 460)
$hero.Fill.Solid()
$hero.Fill.ForeColor.RGB = $RGB_CARD
$hero.Line.ForeColor.RGB = $RGB_BORDER
$hero.Line.Weight = 1.5

$b1 = $s1.Shapes.AddTextbox(1, 90, 80, 500, 24)
$b1.TextFrame.TextRange.Text = "ENTERPRISE GIS & LIVE TRANSIT TELEMETRY"
$b1.TextFrame.TextRange.Font.Name = "Segoe UI"
$b1.TextFrame.TextRange.Font.Size = 11
$b1.TextFrame.TextRange.Font.Bold = -1
$b1.TextFrame.TextRange.Font.Color.RGB = $RGB_ACCENT

$t1 = $s1.Shapes.AddTextbox(1, 90, 115, 780, 80)
$t1.TextFrame.TextRange.Text = "Real-Time Train Movement Tracker"
$t1.TextFrame.TextRange.Font.Name = "Segoe UI"
$t1.TextFrame.TextRange.Font.Size = 34
$t1.TextFrame.TextRange.Font.Bold = -1
$t1.TextFrame.TextRange.Font.Color.RGB = $RGB_WHITE

$sub1 = $s1.Shapes.AddTextbox(1, 90, 205, 780, 35)
$sub1.TextFrame.TextRange.Text = "Track-Level Visualisation, Live Telemetry Ingestion & WebSocket Streaming"
$sub1.TextFrame.TextRange.Font.Name = "Segoe UI"
$sub1.TextFrame.TextRange.Font.Size = 15
$sub1.TextFrame.TextRange.Font.Color.RGB = $RGB_CYAN

$m1 = $s1.Shapes.AddTextbox(1, 90, 260, 780, 190)
$m1.TextFrame.WordWrap = -1
$m1.TextFrame.TextRange.Text = "Author: Abdul Bari`n`nCore Stack: Java 27  |  Spring Boot 3.2.5  |  Leaflet.js  |  OpenRailwayMap  |  MySQL 8.0`n`nLive Web Deployment: https://national-quad-alumni-fastest.trycloudflare.com`nGitHub Repository: https://github.com/bari2003/train-tracker"
$m1.TextFrame.TextRange.Font.Name = "Segoe UI"
$m1.TextFrame.TextRange.Font.Size = 12
$m1.TextFrame.TextRange.Font.Color.RGB = $RGB_BODY
$m1.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3

# ------------------------------------------------------------------------------
# SLIDE 2: Problem Statement
# ------------------------------------------------------------------------------
$s2 = $pres.Slides.Add(2, 12)
Add-SlideChrome $s2 "CONTEXT & MOTIVATION" "Why Train Movement Tracking?" 2

$desc2 = $s2.Shapes.AddTextbox(1, 40, 102, 880, 35)
$desc2.TextFrame.TextRange.Text = "Traditional railway portals leave commuters with uncertainty between distant stations. Train Tracker bridges raw tabular feeds and authentic spatial rail navigation."
$desc2.TextFrame.TextRange.Font.Name = "Segoe UI"
$desc2.TextFrame.TextRange.Font.Size = 11
$desc2.TextFrame.TextRange.Font.Color.RGB = $RGB_BODY

$tbl2 = $s2.Shapes.AddTable(5, 2, 40, 145, 880, 335)
$tbl2.Table.Columns.Item(1).Width = 440
$tbl2.Table.Columns.Item(2).Width = 440

$tHeaders = @("Traditional Systems (NTES / SMS Portals)", "Real-Time Train Tracker Platform")
for ($c = 1; $c -le 2; $c++) {
    $cell = $tbl2.Table.Cell(1, $c)
    $cell.Shape.Fill.Solid()
    $cell.Shape.Fill.ForeColor.RGB = $RGB_HEADER_BG
    $cell.Shape.TextFrame.TextRange.Text = $tHeaders[$c-1]
    $cell.Shape.TextFrame.TextRange.Font.Name = "Segoe UI"
    $cell.Shape.TextFrame.TextRange.Font.Size = 12
    $cell.Shape.TextFrame.TextRange.Font.Bold = -1
    $cell.Shape.TextFrame.TextRange.Font.Color.RGB = if ($c -eq 1) { $RGB_CYAN } else { $RGB_ACCENT }
}

$tRows = @(
    @("Text-Only Feeds: Displays station names & timestamps without geographic spatial context.", "Physical Track Overlay: Train marker animates directly on OpenRailwayMap rails & switches."),
    @("Infrequent Logging: Data updated only when a train arrives at a manned junction.", "Continuous Telemetry Stream: Sub-second speed, delay, and position pushed via WebSocket."),
    @("No Intermediate Location: Commuters cannot see where a train is between two distant halts.", "Geodesic Interpolation: Precise coordinates estimated along 220+ high-resolution track vertices."),
    @("Server Overload: Millions of users polling HTTP REST APIs crash infrastructure.", "Pub/Sub Architecture: Single WebSocket broadcast channel reaches all active commuters.")
)

for ($r = 0; $r -lt 4; $r++) {
    for ($c = 1; $c -le 2; $c++) {
        $cell = $tbl2.Table.Cell($r + 2, $c)
        $cell.Shape.Fill.Solid()
        $cell.Shape.Fill.ForeColor.RGB = $RGB_CARD
        $cell.Shape.TextFrame.TextRange.Text = $tRows[$r][$c-1]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Segoe UI"
        $cell.Shape.TextFrame.TextRange.Font.Size = 10.5
        $cell.Shape.TextFrame.TextRange.Font.Color.RGB = if ($c -eq 1) { $RGB_MUTED } else { $RGB_BODY }
    }
}

# ------------------------------------------------------------------------------
# SLIDE 3: Architecture
# ------------------------------------------------------------------------------
$s3 = $pres.Slides.Add(3, 12)
Add-SlideChrome $s3 "TECHNICAL ARCHITECTURE" "System Architecture & Data Pipeline" 3

Add-Card $s3 40 120 275 365 "📡" "1. Live Ingestion" `
"• RailRadar v1 REST API client`n• Ingests running status, delays, and departure states`n• High-res GeoJSON route fetch with 220+ station vertices`n• Scheduled 15s polling + instant on-demand tracking trigger`n• Automatic schema normalization"

Add-Card $s3 342 120 275 365 "⚙️" "2. Processing Core" `
"• Spring Boot 3.2.5 on Java 27`n• MySQL 8.0 with LONGTEXT GeoJSON storage`n• Orthogonal track-snapping engine (< 500m threshold)`n• Quota-aware fallback simulation for zero downtime`n• Automatic historical telemetry snapshot sampling"

Add-Card $s3 645 120 275 365 "🗺️" "3. Presentation Layer" `
"• Leaflet.js with OpenRailwayMap layer`n• STOMP over SockJS WebSocket pub/sub`n• Smooth 800ms linear interpolation (LERP)`n• Pulsing radar beacon & dynamic bearing rotation`n• Zero-disconnect multi-train switching"

# ------------------------------------------------------------------------------
# SLIDE 4: Core Technical Innovations
# ------------------------------------------------------------------------------
$s4 = $pres.Slides.Add(4, 12)
Add-SlideChrome $s4 "SPATIAL ENGINEERING" "Core Technical Innovations" 4

Add-Card $s4 40 120 425 175 "📐" "Haversine Geodesic Interpolation" `
"Computes exact spherical distance along railway lines (R = 6,371 km). Translates journey completion fraction into sub-meter accurate GPS coordinates along track curves."

Add-Card $s4 495 120 425 175 "🎯" "Perpendicular Track Snapping" `
"Detects GPS reporting drift and orthogonally projects reported coordinates onto the nearest track vector segment within a 500m threshold, eliminating visual jitter."

Add-Card $s4 40 310 425 175 "🔄" "Zero-Downtime Train Switching" `
"Persistent WebSocket connection that dynamically swaps STOMP topic subscriptions (/topic/train/{id}) without connection teardowns or handshake overhead."

Add-Card $s4 495 310 425 175 "✨" "Smooth LERP Marker Motion" `
"Animates the train marker smoothly using requestAnimationFrame over 800ms, rotating the train icon to match the track heading angle (bearing) in real time."

# ------------------------------------------------------------------------------
# SLIDE 5: Resilience & Quota Fallback
# ------------------------------------------------------------------------------
$s5 = $pres.Slides.Add(5, 12)
Add-SlideChrome $s5 "SYSTEM RELIABILITY" "Resilience & Quota Fallback Engine" 5

$desc5 = $s5.Shapes.AddTextbox(1, 40, 102, 880, 35)
$desc5.TextFrame.TextRange.Text = "Third-party transit APIs enforce monthly request limits (HTTP 429). Train Tracker incorporates an intelligent dual-mode fallback architecture that guarantees zero application freezes."
$desc5.TextFrame.TextRange.Font.Name = "Segoe UI"
$desc5.TextFrame.TextRange.Font.Size = 11
$desc5.TextFrame.TextRange.Font.Color.RGB = $RGB_BODY

Add-Card $s5 40 150 425 335 "🟢" "Live API Mode (Primary)" `
"• Direct ingestion of real-time RailRadar v1 API`n• Live station arrival & departure provenance`n• Exact delay minutes (+18 min late / On time)`n• Real-time departure timestamps & platform info`n• Automatic high-precision route geometry caching`n• Runs whenever external quota is active"

Add-Card $s5 495 150 425 335 "🛡️" "Smart Kinematic Fallback (Resilient)" `
"• Triggers automatically on HTTP 429 or network timeouts`n• Computes realistic kinematic progression along physical track polyline`n• Generates accurate bearing and simulated speeds (78-100 km/h)`n• Prevents application freezing or STALE disconnections`n• Transitions back to live API when quota resets"

# ------------------------------------------------------------------------------
# SLIDE 6: Mapping Stack
# ------------------------------------------------------------------------------
$s6 = $pres.Slides.Add(6, 12)
Add-SlideChrome $s6 "USER EXPERIENCE & GIS" "Frontend Mapping & Visualisation Stack" 6

Add-Card $s6 40 120 275 365 "🗺️" "Standard OSM Base" `
"• Free, community-driven global cartography`n• Zero API key requirement (removes Carto watermark)`n• High contrast borders, junctions, and major cities`n• Fast vector tile rendering across all zoom levels`n• Seamless mobile responsive touch controls"

Add-Card $s6 342 120 275 365 "🚆" "OpenRailwayMap Overlay" `
"• Dedicated global railway infrastructure layer`n• Displays physical tracks, sidings, and electrified routes`n• Renders switches, signals, junctions, and platforms`n• Gives authentic spatial context to train movement`n• Free, publicly served transparent tiles"

Add-Card $s6 645 120 275 365 "📊" "Interactive Telemetry HUD" `
"• Current station indicator (e.g. Departed Limkheda LMK)`n• Next scheduled halt & remaining distance`n• Real-time speed gauge & color-coded delay status`n• Pulsing radar beacon with 'Follow Train' tracking lock`n• 0 ms instant visual data hydration"

# ------------------------------------------------------------------------------
# SLIDE 7: Database Architecture
# ------------------------------------------------------------------------------
$s7 = $pres.Slides.Add(7, 12)
Add-SlideChrome $s7 "PERSISTENCE LAYER" "Database Architecture & Data Model" 7

$desc7 = $s7.Shapes.AddTextbox(1, 40, 102, 880, 35)
$desc7.TextFrame.TextRange.Text = "Built on MySQL 8.0 with support for in-memory H2 profiles for testing. Upgraded to LONGTEXT GeoJSON storage to support massive Indian Railways routes containing 220+ stations without truncation."
$desc7.TextFrame.TextRange.Font.Name = "Segoe UI"
$desc7.TextFrame.TextRange.Font.Size = 11
$desc7.TextFrame.TextRange.Font.Color.RGB = $RGB_BODY

$tbl7 = $s7.Shapes.AddTable(5, 4, 40, 145, 880, 335)
$tbl7.Table.Columns.Item(1).Width = 150
$tbl7.Table.Columns.Item(2).Width = 200
$tbl7.Table.Columns.Item(3).Width = 265
$tbl7.Table.Columns.Item(4).Width = 265

$t7H = @("Table Name", "Primary / Foreign Keys", "Key Attributes & Types", "Description")
for ($c = 1; $c -le 4; $c++) {
    $cell = $tbl7.Table.Cell(1, $c)
    $cell.Shape.Fill.Solid()
    $cell.Shape.Fill.ForeColor.RGB = $RGB_HEADER_BG
    $cell.Shape.TextFrame.TextRange.Text = $t7H[$c-1]
    $cell.Shape.TextFrame.TextRange.Font.Name = "Segoe UI"
    $cell.Shape.TextFrame.TextRange.Font.Size = 11
    $cell.Shape.TextFrame.TextRange.Font.Bold = -1
    $cell.Shape.TextFrame.TextRange.Font.Color.RGB = $RGB_CYAN
}

$t7Rows = @(
    @("trains", "train_id (PK)`nsource/dest_id (FK)", "train_number (UK), train_name, status", "Train catalog and schedule status (Rajdhani, Shatabdi, Duronto)"),
    @("routes", "route_id (PK)`ntrain_id (FK, UK)", "geometry_geojson (LONGTEXT), distance_km", "Full high-resolution track polyline & stops (74 KB+ JSON)"),
    @("stations", "station_id (PK)", "station_code (UK), station_name, lat, lon", "Railway stations coordinate registry across Indian Railways network"),
    @("live_train_states", "state_id (PK)`ntrain_id (FK, UK)", "lat, lon, speed_kmh, delay_minutes, last_updated", "Latest telemetry snapshot for instant cache retrieval on click")
)

for ($r = 0; $r -lt 4; $r++) {
    for ($c = 1; $c -le 4; $c++) {
        $cell = $tbl7.Table.Cell($r + 2, $c)
        $cell.Shape.Fill.Solid()
        $cell.Shape.Fill.ForeColor.RGB = $RGB_CARD
        $cell.Shape.TextFrame.TextRange.Text = $t7Rows[$r][$c-1]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Segoe UI"
        $cell.Shape.TextFrame.TextRange.Font.Size = 10
        $cell.Shape.TextFrame.TextRange.Font.Color.RGB = if ($c -eq 1) { $RGB_WHITE } elseif ($c -eq 3 -and $r -eq 1) { $RGB_ACCENT } else { $RGB_BODY }
    }
}

# ------------------------------------------------------------------------------
# SLIDE 8: Cloud Deployment
# ------------------------------------------------------------------------------
$s8 = $pres.Slides.Add(8, 12)
Add-SlideChrome $s8 "DEVOPS & DEPLOYMENT" "Cloud Deployment & Edge Topology" 8

Add-Card $s8 40 120 425 210 "☕" "Application Runtime" `
"• Java 27 (Oracle JDK 27.0.1.1) runtime environment`n• Spring Boot 3.2.5 with Tomcat 10.1 embedded web server`n• MySQL 8.0 relational persistence with automatic schema seeding`n• Optimized Maven build pipeline packaging clean executable JARs"

Add-Card $s8 495 120 425 210 "☁️" "Cloudflare Zero-Trust Edge" `
"• Global public HTTPS ingress without router port-forwarding`n• Secure WebSocket proxying with sub-second message delivery`n• Automatic SSL/TLS termination and DDoS mitigation`n• Accessible from any mobile or desktop browser globally"

$linkCard = $s8.Shapes.AddShape(1, 40, 350, 880, 135)
$linkCard.Fill.Solid()
$linkCard.Fill.ForeColor.RGB = Get-RGB 14 42 56
$linkCard.Line.ForeColor.RGB = $RGB_CYAN
$linkCard.Line.Weight = 1.5

$lTxt = $s8.Shapes.AddTextbox(1, 60, 365, 840, 110)
$lTxt.TextFrame.WordWrap = -1
$lTxt.TextFrame.TextRange.Text = "🌐 LIVE PUBLIC ACCESS & SOURCE CODE`n`nLive Web App:  https://national-quad-alumni-fastest.trycloudflare.com`nGitHub Repository:  https://github.com/bari2003/train-tracker"
$lTxt.TextFrame.TextRange.Font.Name = "Segoe UI"
$lTxt.TextFrame.TextRange.Font.Size = 12
$lTxt.TextFrame.TextRange.Font.Bold = -1
$lTxt.TextFrame.TextRange.Font.Color.RGB = $RGB_WHITE
$lTxt.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3

# ------------------------------------------------------------------------------
# SLIDE 9: Real-World Applications
# ------------------------------------------------------------------------------
$s9 = $pres.Slides.Add(9, 12)
Add-SlideChrome $s9 "PRACTICAL USE CASES" "Real-World Applications & Industry Impact" 9

Add-Card $s9 40 120 425 175 "👨‍👩‍👧" "Passenger Safety & Convenience" `
"Provides commuters and families with exact track positions, mitigating anxiety during late-night travel, foggy winter delays, and monsoon disruptions."

Add-Card $s9 495 120 425 175 "📦" "Freight & Supply Chain Logistics" `
"Enables industrial logistics operators to monitor high-value cargo and parcel movements along Dedicated Freight Corridors (DFC) in real time."

Add-Card $s9 40 310 425 175 "🏢" "Station Platform Operations" `
"Assists station controllers in managing passenger crowd flow by displaying accurate estimated arrival times and platform assignments."

Add-Card $s9 495 310 425 175 "🚨" "Emergency Response & Maintenance" `
"Gives control rooms spatial awareness of train positions during block maintenance, signal failures, or emergency track diversions."

# ------------------------------------------------------------------------------
# SLIDE 10: Future Roadmap & Conclusion
# ------------------------------------------------------------------------------
$s10 = $pres.Slides.Add(10, 12)
Add-SlideChrome $s10 "CONCLUSION & VISION" "Future Roadmap & Project Conclusion" 10

Add-Card $s10 40 120 275 220 "📡" "Zonal Fleet Radar" `
"Scale the map engine to visualize hundreds of trains simultaneously across an entire railway zone using spatial clustering and WebGL."

Add-Card $s10 342 120 275 220 "🤖" "AI Predictive Delay" `
"Machine learning models trained on junction bottlenecks and weather forecasts to predict delays before they occur."

Add-Card $s10 645 120 275 220 "📱" "Offline-First PWA" `
"Service Worker caching allowing passengers inside moving trains to view track progression without active internet connectivity."

$cBanner = $s10.Shapes.AddShape(1, 40, 360, 880, 125)
$cBanner.Fill.Solid()
$cBanner.Fill.ForeColor.RGB = $RGB_CARD
$cBanner.Line.ForeColor.RGB = $RGB_BORDER
$cBanner.Line.Weight = 1

$cTxt = $s10.Shapes.AddTextbox(1, 60, 375, 840, 100)
$cTxt.TextFrame.WordWrap = -1
$cTxt.TextFrame.TextRange.Text = "EXECUTIVE SUMMARY`n`nThe Real-Time Train Movement Tracker successfully unites modern Java 27 backend engineering, real-time WebSocket pub/sub streaming, and open-source spatial cartography to deliver a scalable, track-accurate railway navigation platform."
$cTxt.TextFrame.TextRange.Font.Name = "Segoe UI"
$cTxt.TextFrame.TextRange.Font.Size = 11.5
$cTxt.TextFrame.TextRange.Font.Color.RGB = $RGB_WHITE
$cTxt.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3

# ------------------------------------------------------------------------------
# SAVE EXCLUSIVELY AS .PPT
# ------------------------------------------------------------------------------
$outPpt = "C:\Users\ABDUL BARI\Documents\train-tracker - Copy\Real_Time_Train_Movement_Tracker.ppt"

# 1 = ppSaveAsPresentation (.ppt format)
$pres.SaveAs($outPpt, 1)

$pres.Close()
$pptApp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp) | Out-Null

Write-Host "SUCCESS: Generated proper .ppt presentation at: $outPpt"
