# IGI Command Center — Weekly Data Parser
# Usage: .\parse-weekly.ps1

param(
  [string]$WeeklyRoot = "C:\Users\justi\.openclaw\workspace\_Ignite Growth Intelligence\Data\Weekly",
  [string]$OutputDir  = "C:\Users\justi\.openclaw\workspace\igi-command-center\data"
)

$regionMap = @{
  "Battle Creek"="kathi";"Flint"="kathi";"Grand Rapids"="kathi";"Kalamazoo"="kathi"
  "Killeen-Temple"="kathi";"Killeen/Temple"="kathi";"Lafayette"="kathi";"Lake Charles"="kathi"
  "Lansing"="kathi";"Lufkin"="kathi";"Rockford"="kathi";"Shreveport"="kathi"
  "Texarkana"="kathi";"Tyler"="kathi";"Victoria"="kathi"
  "Billings"="taylor";"Boise"="taylor";"Bozeman"="taylor";"Butte"="taylor"
  "Casper"="taylor";"Cheyenne"="taylor";"Fort Collins"="taylor";"Ft Collins"="taylor"
  "Great Falls"="taylor";"Laramie"="taylor";"Tri-Cities"="taylor";"Twin Falls"="taylor"
  "Wenatchee"="taylor";"Yakima"="taylor"
  "Abilene"="tylerw";"Amarillo"="tylerw";"El Paso"="tylerw";"Lawton"="tylerw"
  "Lubbock"="tylerw";"Odessa"="tylerw";"Odessa-Midland"="tylerw";"San Angelo"="tylerw"
  "Wichita Falls"="tylerw"
  "Bismarck"="jeroen";"Cedar Rapids"="jeroen";"Duluth"="jeroen";"Faribault"="jeroen"
  "Faribault/Owatonna"="jeroen";"Rochester"="jeroen";"Rochester MN"="jeroen"
  "Sedalia"="jeroen";"Waterloo"="jeroen"
}

$regionNames = @{
  "kathi"="Kathi Kirkland"; "taylor"="Taylor Wheeler"
  "tylerw"="Tyler Wille";   "jeroen"="Jeroen Corver"; "direct"="Direct Markets"
}

$aliases = @{
  "Ft Collins"="Fort Collins"; "Odessa-Midland"="Odessa"; "Killeen-Temple"="Killeen/Temple"
  "Evansville-Owensboro"="Evansville/Owensboro"; "Faribault"="Faribault/Owatonna"
  "Rochester"="Rochester MN"; "Quincy_Hannibal"="Quincy/Hannibal"
  "St George"="St. George"
}

$skipMarkets = @("NABCO","Backyard","Powell","Reno","Atlantic City")

function Normalize([string]$m) {
  $m = $m.Trim()
  if ($aliases.ContainsKey($m)) { return $aliases[$m] }
  return $m
}

function Get-Region([string]$m) {
  if ($regionMap.ContainsKey($m))  { return $regionMap[$m] }
  if ($regionMap.ContainsKey((Normalize $m))) { return $regionMap[(Normalize $m)] }
  return "direct"
}

function Clean-Num([string]$s) {
  $s = ($s -replace '[\$,"\s]','').Trim()
  if ($s -eq '' -or $s -eq '-' -or $s -eq 'null') { return 0.0 }
  try { return [double]$s } catch { return 0.0 }
}

# ── Blueprint CSV Parser ───────────────────────────────────────────────────────
function Parse-Blueprint([string]$path) {
  $result = @{}
  $lines = Get-Content $path -Encoding UTF8
  $inIgnite = $false; $headerSeen = $false
  foreach ($line in $lines) {
    $line = $line.Trim()
    if ($line -match '^IGNITE,') { $inIgnite = $true; $headerSeen = $false; continue }
    if ($inIgnite -and $line -match '^Pending Pitches') { continue }
    if ($inIgnite -and $line -match '^MARKET,') { $headerSeen = $true; continue }
    if ($inIgnite -and $headerSeen) {
      if ($line -match '^(BROADCAST|AMPED|EVENTS|STD|OLR|TSI|"2026)') { break }
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      # Parse CSV line with quoted fields: Market,"$x","$y","$z","$t"
      if ($line -match '^"?([^,"]+)"?,(.+)') {
        $market = Normalize ($matches[1].Trim('"').Trim())
        $rest = $matches[2]
        # Extract numeric values, handling quoted dollar amounts
        $vals = @()
        foreach ($token in ($rest -split ',')) {
          $vals += Clean-Num $token
          if ($vals.Count -ge 3) { break }
        }
        if ($vals.Count -ge 3 -and $market -ne '' -and $market -ne 'MARKET') {
          $result[$market] = @{ apr=$vals[0]; may=$vals[1]; jun=$vals[2] }
        }
      }
    }
  }
  return $result
}

# ── WO Analytics TSV Parser ───────────────────────────────────────────────────
function Parse-WO([string]$path) {
  $result = @{}
  $bytes = [System.IO.File]::ReadAllBytes($path)
  $text = [System.Text.Encoding]::Unicode.GetString($bytes)
  $lines = ($text -split "`r`n|`n") | Where-Object { $_.Trim() -ne '' }
  if ($lines.Count -lt 4) { return $result }

  # Line 2 (index 2) = header: "Actual As of Date  Market  Total  3  4  5  6..."
  $header = $lines[2] -split '\t'
  $aprIdx = $mayIdx = $junIdx = -1
  $firstFour = $false
  for ($c = 0; $c -lt $header.Count; $c++) {
    $v = $header[$c].Trim()
    if ($v -eq '4' -and $aprIdx -lt 0) { $aprIdx = $c }
    elseif ($v -eq '5' -and $mayIdx -lt 0) { $mayIdx = $c }
    elseif ($v -eq '6' -and $junIdx -lt 0) { $junIdx = $c }
  }

  if ($aprIdx -lt 0 -or $mayIdx -lt 0 -or $junIdx -lt 0) {
    Write-Warning "    Could not find month columns. Apr=$aprIdx May=$mayIdx Jun=$junIdx"
    return $result
  }

  # Data rows start at index 3
  for ($i = 3; $i -lt $lines.Count; $i++) {
    $cols = $lines[$i] -split '\t'
    if ($cols.Count -lt 3) { continue }
    $dateVal   = $cols[0].Trim()
    $marketVal = $cols[1].Trim()
    if ($dateVal -notmatch '^\d+/\d+/\d+') { continue }
    if ($marketVal -eq '' -or $marketVal -eq 'Grand Total' -or $marketVal -eq 'Total') { continue }
    $market = Normalize $marketVal
    if ($skipMarkets -contains $market) { continue }
    $apr = if ($aprIdx -lt $cols.Count) { Clean-Num $cols[$aprIdx] } else { 0.0 }
    $may = if ($mayIdx -lt $cols.Count) { Clean-Num $cols[$mayIdx] } else { 0.0 }
    $jun = if ($junIdx -lt $cols.Count) { Clean-Num $cols[$junIdx] } else { 0.0 }
    $result[$market] = @{ apr=$apr; may=$may; jun=$jun }
  }
  return $result
}

# ── Build JSON ─────────────────────────────────────────────────────────────────
function Build-WeekJson([string]$weekDate,[hashtable]$commits,[hashtable]$adds) {
  $allM = @{}
  foreach ($k in $commits.Keys) { $allM[$k]=$true }
  foreach ($k in $adds.Keys)    { $allM[$k]=$true }

  $byRegion = @{ kathi=@(); taylor=@(); tylerw=@(); jeroen=@(); direct=@() }
  $seen = @{}

  foreach ($market in ($allM.Keys | Sort-Object)) {
    if ($seen.ContainsKey($market)) { continue }
    $seen[$market] = $true
    $rid = Get-Region $market
    $c = if ($commits.ContainsKey($market)) { $commits[$market] } else { @{apr=0;may=0;jun=0} }
    $a = if ($adds.ContainsKey($market))    { $adds[$market] }    else { @{apr=0;may=0;jun=0} }
    $byRegion[$rid] += [ordered]@{
      name=$market
      apr=[ordered]@{ commit=[math]::Round($c.apr,2); adds=[math]::Round($a.apr,2) }
      may=[ordered]@{ commit=[math]::Round($c.may,2); adds=[math]::Round($a.may,2) }
      jun=[ordered]@{ commit=[math]::Round($c.jun,2); adds=[math]::Round($a.jun,2) }
    }
  }

  $dt = [datetime]::ParseExact($weekDate,'yyyy-MM-dd',$null)
  $months = @{1="January";2="February";3="March";4="April";5="May";6="June";
               7="July";8="August";9="September";10="October";11="November";12="December"}
  $weekLabel = "$($months[$dt.Month]) $($dt.Day), $($dt.Year)"

  $regionOrder = @("kathi","taylor","tylerw","jeroen","direct")
  $regionList = $regionOrder | ForEach-Object {
    [ordered]@{ id=$_; name=$regionNames[$_]; markets=$byRegion[$_] }
  }

  return [ordered]@{
    lastRefreshed=(Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    weekOf=$weekLabel; weekDate=$weekDate; regions=$regionList
  }
}

# ── Main ───────────────────────────────────────────────────────────────────────
$weeks = @()
foreach ($folder in (Get-ChildItem $WeeklyRoot -Directory | Sort-Object Name)) {
  $weekDate = $folder.Name
  Write-Host "Week: $weekDate"
  $bpFile = Get-ChildItem $folder.FullName "Blueprint*.csv" -ErrorAction SilentlyContinue | Select-Object -Last 1
  $woFile = Get-ChildItem $folder.FullName "Sales Revenue.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $bpFile -or -not $woFile) { Write-Warning "  Missing files, skipping"; continue }

  $commits = Parse-Blueprint $bpFile.FullName
  $adds    = Parse-WO        $woFile.FullName
  Write-Host "  Commits: $($commits.Count)  Adds: $($adds.Count)"

  $json = Build-WeekJson $weekDate $commits $adds
  $outPath = Join-Path $OutputDir "$weekDate.json"
  $json | ConvertTo-Json -Depth 10 | Set-Content $outPath -Encoding UTF8
  Write-Host "  -> $outPath"
  $weeks += [ordered]@{ date=$weekDate; label=$json.weekOf }
}

$weeksIndex = [ordered]@{ weeks=($weeks | Sort-Object { $_.date } -Descending) }
$weeksIndex | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $OutputDir "weeks.json") -Encoding UTF8

$latest = ($weeks | Sort-Object { $_.date } -Descending | Select-Object -First 1).date
if ($latest) {
  Copy-Item (Join-Path $OutputDir "$latest.json") (Join-Path $OutputDir "commits.json") -Force
  Write-Host "commits.json -> $latest"
}
Write-Host "Done. $($weeks.Count) weeks processed."
