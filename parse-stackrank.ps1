# IGI Command Center -- Stack Rank Parser
# Reads last 4 weeks of SOB JSON, Sales Revenue CSV, and DSM snapshots
# Outputs data/stackrank.json
# Usage: .\parse-stackrank.ps1

param(
 [string]$WeeklyRoot = "C:\Users\justi\.openclaw\workspace\_Ignite Growth Intelligence\Data\Weekly",
 [string]$DataDir = "C:\Users\justi\.openclaw\workspace\igi-command-center\data"
)

$nameMap = @{
 "Trenton"="Trenton/Princeton"; "Ft Collins"="Fort Collins"
 "Killeen-Temple"="Killeen/Temple"; "Evansville-Owensboro"="Evansville/Owensboro"
 "Faribault"="Faribault/Owatonna"; "St George"="St. George"
 "St Cloud"="St. Cloud"; "Odessa-Midland"="Odessa"; "Rochester"="Rochester MN"
}

$risdNames = @{
 "kathi"="Kathi Kirkland"; "taylor"="Taylor Wheeler"
 "tylerw"="Tyler Wille"; "jeroen"="Jeroen Corver"
 "nne"=""; "nj"=""; "ny"=""; "others"=""
}

function Parse-Num([string]$s) {
 $s = $s.Trim() -replace '[\$,\s]',''
 $neg = ($s -match '^\(') -or ($s -match '^-')
 $s = $s -replace '[^0-9.]',''
 if ($s -eq '') { return 0.0 }
 $v = try { [double]$s } catch { 0.0 }
 if ($neg) { return -$v } else { return $v }
}

function Read-SalesCSV([string]$path) {
 $adds = @{}
 try {
 $raw = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::Unicode)
 $lines = $raw -split "`r?`n"
 if ($lines.Count -lt 4) { return $adds }
 $header = $lines[2] -split "`t"
 $col4 = [Array]::IndexOf($header, '4')
 $col5 = [Array]::IndexOf($header, '5')
 $col6 = [Array]::IndexOf($header, '6')
 if ($col4 -lt 0 -or $col5 -lt 0 -or $col6 -lt 0) { return $adds }
 for ($i = 3; $i -lt $lines.Count; $i++) {
 $cols = $lines[$i] -split "`t"
 if ($cols.Count -lt 3) { continue }
 $mktRaw = $cols[1].Trim()
 if ($mktRaw -eq '' -or $mktRaw -match '(?i)grand total') { continue }
 $name = if ($nameMap.ContainsKey($mktRaw)) { $nameMap[$mktRaw] } else { $mktRaw }
 $apr = Parse-Num (if ($cols.Count -gt $col4) { $cols[$col4] } else { '' })
 $may = Parse-Num (if ($cols.Count -gt $col5) { $cols[$col5] } else { '' })
 $jun = Parse-Num (if ($cols.Count -gt $col6) { $cols[$col6] } else { '' })
 $adds[$name] = $apr + $may + $jun
 }
 } catch { }
 return $adds
}

# Get last 4 weekly folders sorted newest first
$folders = Get-ChildItem $WeeklyRoot -Directory | Sort-Object Name -Descending | Select-Object -First 4
$folders = $folders | Sort-Object Name # re-sort oldest first for delta calculation
$weekDates = $folders | ForEach-Object { $_.Name }
Write-Host "Analyzing weeks: $($weekDates -join ', ')"

# Load SOB data for each week
$sobByWeek = @{}
foreach ($wd in $weekDates) {
 $path = Join-Path $DataDir "sob-$wd.json"
 if (Test-Path $path) {
 $sobByWeek[$wd] = Get-Content $path -Raw | ConvertFrom-Json
 Write-Host "SOB loaded: $wd"
 } else {
 Write-Host "SOB missing: $wd"
 }
}

# Load Sales Revenue adds for each week
$addsByWeek = @{}
foreach ($folder in $folders) {
 $wd = $folder.Name
 $csvPath = Join-Path $folder.FullName "Sales Revenue.csv"
 if (Test-Path $csvPath) {
 $addsByWeek[$wd] = Read-SalesCSV $csvPath
 Write-Host "Sales CSV loaded: $wd ($($addsByWeek[$wd].Count) markets)"
 } else {
 $addsByWeek[$wd] = @{}
 Write-Host "Sales CSV missing: $wd"
 }
}

# Load DSM snapshots
$dsmByWeek = @{}
foreach ($wd in $weekDates) {
 $path = Join-Path $DataDir "dsm-$wd.json"
 if (Test-Path $path) {
 $dsmByWeek[$wd] = Get-Content $path -Raw | ConvertFrom-Json
 Write-Host "DSM loaded: $wd"
 } else {
 Write-Host "DSM missing: $wd"
 }
}

# Use most recent SOB for market list, focus rankings, and DSM/RISD lookups
$oldestWd = $weekDates[0]
$newestWd = $weekDates[$weekDates.Count - 1]
$latestSob = $sobByWeek[$newestWd]
if (-not $latestSob) { Write-Error "No SOB data for newest week $newestWd"; exit 1 }

# Build market -> RISD lookup
$marketRisd = @{}
foreach ($region in $latestSob.regions) {
 $rn = if ($risdNames.ContainsKey($region.id)) { $risdNames[$region.id] } else { "" }
 foreach ($mkt in $region.markets) { $marketRisd[$mkt.name] = $rn }
}

# Build market -> DSM lookup from most recent DSM snapshot
$marketDsm = @{}
$latestDsm = $dsmByWeek[$newestWd]
if ($latestDsm) {
 foreach ($dsm in $latestDsm.dsms) {
 foreach ($mkt in $dsm.markets) { $marketDsm[$mkt] = $dsm.name }
 }
}

# Build all market names from latest SOB
$allMarkets = @()
foreach ($region in $latestSob.regions) {
 foreach ($mkt in $region.markets) { $allMarkets += $mkt.name }
}

# Compute market momentum
$marketRows = @()
foreach ($mktName in $allMarkets) {
 # Q2 and FY pacing: oldest vs newest (values in $000s)
 $sobOld = 0; $sobNew = 0; $fyOld = 0; $fyNew = 0
 $q2PctBgt = 0; $q2Budget = 0; $q2PctPY = 0
 $fyBudget = 0; $fyPctBgt = 0; $fyGap = 0
 if ($sobByWeek[$oldestWd]) {
 foreach ($r in $sobByWeek[$oldestWd].regions) {
 foreach ($m in $r.markets) {
 if ($m.name -eq $mktName) {
  $sobOld = $m.q2.total.pacing
  $fyOld = $m.q2.total.pacing + $m.q3.total.pacing
  break
 }
 }
 }
 }
 foreach ($r in $latestSob.regions) {
 foreach ($m in $r.markets) {
 if ($m.name -eq $mktName) {
  $sobNew = $m.q2.total.pacing
  $q2PctBgt = $m.q2.total.pctBgt
  $q2Budget = $m.q2.total.budget
  $q2PctPY = $m.q2.total.pctPY
  $fyNew = $m.q2.total.pacing + $m.q3.total.pacing
  $fyBudget = $m.q2.total.budget + $m.q3.total.budget
  $fyPctBgt = if ($fyBudget -ne 0) { [math]::Round(($fyNew / $fyBudget) * 100, 1) } else { 0 }
  $fyGap = [math]::Round($fyNew - $fyBudget, 0)
  break
 }
 }
 }
 $sobDelta = $sobNew - $sobOld
 $sobDeltaPct = if ($sobOld -ne 0) { [math]::Round(($sobDelta / $sobOld) * 100, 1) } else { 0 }
 $fyPacingDelta = $fyNew - $fyOld

 # Cumulative adds over 4 weeks from Sales Revenue CSVs
 $cumAdds = 0
 foreach ($wd in $weekDates) {
 $v = if ($addsByWeek[$wd].ContainsKey($mktName)) { $addsByWeek[$wd][$mktName] } else { 0 }
 $cumAdds += $v
 }

 $marketRows += [ordered]@{
 name = $mktName
 dsm = if ($marketDsm.ContainsKey($mktName)) { $marketDsm[$mktName] } else { "" }
 risd = if ($marketRisd.ContainsKey($mktName)) { $marketRisd[$mktName] } else { "" }
 sobDelta = $sobDelta
 sobDeltaPct = $sobDeltaPct
 fyPacingDelta = $fyPacingDelta
 fyPctBgt = $fyPctBgt
 fyGap = $fyGap
 fyBudget = $fyBudget
 cumulativeAdds = [math]::Round($cumAdds, 0)
 q2PctBgt = $q2PctBgt
 q2Budget = $q2Budget
 q2PctPY = $q2PctPY
 }
}

# Sort by SOB delta descending for hot/cold
$sorted = $marketRows | Sort-Object { $_.sobDelta } -Descending
$hotMkts = $sorted | Select-Object -First 10
$coldMkts = ($sorted | Select-Object -Last 10) | Sort-Object { $_.sobDelta }

# Focus: markets below 85% Q2 AND FY $ gap worse than -$150K, ranked by FY $ gap
# (thresholds are in $000s to match SOB data units)
$focusQ2PctCeiling = 85      # Q2 % to budget must be below this
$focusFyGapFloor   = -150    # FY $ gap must be worse than this (in $000s = -$150K)
$coldNames = $coldMkts | ForEach-Object { $_.name }
$focusMktsSorted = ($marketRows | Where-Object {
 $_.q2PctBgt -lt $focusQ2PctCeiling -and $_.fyGap -lt $focusFyGapFloor
} | Sort-Object { $_.fyGap }) | Select-Object -First 15
$focusMkts = @()
foreach ($fm in $focusMktsSorted) {
 $focusMkts += [ordered]@{
  name = $fm.name
  dsm = $fm.dsm
  risd = $fm.risd
  fyPctBgt = $fm.fyPctBgt
  fyGap = $fm.fyGap
  fyBudget = $fm.fyBudget
  q2PctBgt = $fm.q2PctBgt
  q2PctPY = $fm.q2PctPY
  isAlsoCold = ($coldNames -contains $fm.name)
 }
}

# DSM rankings: FY total revenue change oldest->newest (Q1+Q2+Q3)
$dsmRows = @()
$oldestDsm = $dsmByWeek[$oldestWd]
$newestDsm = $dsmByWeek[$newestWd]
if ($oldestDsm -and $newestDsm) {
 $oldDsmMap = @{}
 foreach ($d in $oldestDsm.dsms) { $oldDsmMap[$d.name] = $d }
 foreach ($d in $newestDsm.dsms) {
 if (-not $oldDsmMap.ContainsKey($d.name)) { continue }
 $oldD = $oldDsmMap[$d.name]
 $fyStart = $oldD.totalRevenue.fyTotal
 $fyEnd = $d.totalRevenue.fyTotal
 $q2Delta = $fyEnd - $fyStart
 $q2Pct = if ($fyStart -ne 0) { [math]::Round(($q2Delta / $fyStart) * 100, 1) } else { 0 }
 $dsmRows += [ordered]@{
 name = $d.name
 markets = $d.markets
 q2Start = [math]::Round($fyStart, 0)
 q2End = [math]::Round($fyEnd, 0)
 q2Delta = [math]::Round($q2Delta, 0)
 q2DeltaPct = $q2Pct
 }
 }
}

$dsmSorted = $dsmRows | Sort-Object { $_.q2Delta } -Descending
$hotDsms = $dsmSorted | Select-Object -First 10
$coldDsms = ($dsmSorted | Select-Object -Last 10) | Sort-Object { $_.q2Delta }

# Build week labels for display
$weekLabels = $weekDates | ForEach-Object {
 $dt = [datetime]::ParseExact($_, 'yyyy-MM-dd', $null)
 "Apr " + $dt.Day.ToString()
}

$output = [ordered]@{
 lastRefreshed = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
 weeksAnalyzed = $weekLabels
 weekDates = $weekDates
 markets = [ordered]@{
 hot = $hotMkts
 cold = $coldMkts
 focus = $focusMkts
 }
 dsms = [ordered]@{
 hot = $hotDsms
 cold = $coldDsms
 }
}

$outPath = Join-Path $DataDir "stackrank.json"
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$jsonContent = $output | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($outPath, $jsonContent, $utf8NoBom)
Write-Host "Written: $outPath"
