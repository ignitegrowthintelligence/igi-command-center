# IGI Command Center — Weekly Data Parser
# Months are derived from the week date (current month + 2 forward)
# Usage: .\parse-weekly.ps1

param(
  [string]$WeeklyRoot = "C:\Users\justi\.openclaw\workspace\_Ignite Growth Intelligence\Data\Weekly",
  [string]$OutputDir  = "C:\Users\justi\.openclaw\workspace\igi-command-center\data"
)

# Complete master market list — all markets always shown regardless of weekly activity
$masterMarkets = @{
  "kathi"  = @("Flint","Grand Rapids","Kalamazoo","Killeen/Temple","Lafayette","Lake Charles","Lansing","Lufkin","Rockford","Shreveport","Texarkana","Tyler","Victoria")
  "taylor" = @("Billings","Boise","Bozeman","Butte","Casper","Cheyenne","Fort Collins","Great Falls","Laramie","Sierra Vista","St. George","Tri-Cities","Twin Falls","Wenatchee","Williston","Yakima")
  "tylerw" = @("Abilene","Amarillo","El Paso","Lawton","Lubbock","Odessa","San Angelo","Wichita Falls")
  "jeroen" = @("Bismarck","Cedar Rapids","Dubuque","Duluth","Faribault/Owatonna","Quad Cities","Quincy/Hannibal","Rochester MN","Sedalia","Waterloo")
  "nne"    = @("Augusta","Bangor","New Bedford","Portland","Portsmouth","Presque Isle")
  "nj"     = @("Atlantic City","Shore","Trenton/Princeton")
  "ny"     = @("Albany","Berkshires","Binghamton","Buffalo","Danbury","Oneonta","Poughkeepsie","Utica")
  "others" = @("Evansville/Owensboro","Grand Junction","Missoula","Montrose","Shelby","Sioux Falls","St. Cloud","Tuscaloosa")
}

# Region map — sourced from Regional coverage.xlsx (Market Realignment tab)
$regionMap = @{
  # Kathi Kirkland
  "Flint"="kathi";"Grand Rapids"="kathi";"Kalamazoo"="kathi"
  "Killeen-Temple"="kathi";"Killeen/Temple"="kathi";"Lafayette"="kathi";"Lake Charles"="kathi"
  "Lansing"="kathi";"Lufkin"="kathi";"Rockford"="kathi";"Shreveport"="kathi"
  "Texarkana"="kathi";"Tyler"="kathi";"Victoria"="kathi"
  # Taylor Wheeler (includes St George, Williston, Sierra Vista, Billings, Bozeman)
  "Billings"="taylor";"Boise"="taylor";"Bozeman"="taylor";"Butte"="taylor"
  "Casper"="taylor";"Cheyenne"="taylor";"Fort Collins"="taylor";"Ft Collins"="taylor"
  "Great Falls"="taylor";"Laramie"="taylor";"Tri-Cities"="taylor";"Twin Falls"="taylor"
  "Wenatchee"="taylor";"Yakima"="taylor"
  "St. George"="taylor";"St George"="taylor"
  "Williston"="taylor";"Sierra Vista"="taylor"
  # Tyler Wille
  "Abilene"="tylerw";"Amarillo"="tylerw";"El Paso"="tylerw";"Lawton"="tylerw"
  "Lubbock"="tylerw";"Odessa"="tylerw";"Odessa-Midland"="tylerw";"San Angelo"="tylerw"
  "Wichita Falls"="tylerw"
  # Jeroen Corver (includes Dubuque, Quad Cities, Quincy/Hannibal)
  "Bismarck"="jeroen";"Cedar Rapids"="jeroen";"Duluth"="jeroen"
  "Faribault"="jeroen";"Faribault/Owatonna"="jeroen"
  "Rochester"="jeroen";"Rochester MN"="jeroen"
  "Sedalia"="jeroen";"Waterloo"="jeroen"
  "Dubuque"="jeroen";"Quad Cities"="jeroen";"Quincy/Hannibal"="jeroen"
  # NNE
  "Augusta"="nne";"Bangor"="nne";"New Bedford"="nne"
  "Portland"="nne";"Portsmouth"="nne";"Presque Isle"="nne"
  # NJ
  "Atlantic City"="nj";"Shore"="nj";"Trenton/Princeton"="nj"
  "Princeton"="nj";"Trenton"="nj"
  # NY
  "Albany"="ny";"Berkshires"="ny";"Binghamton"="ny";"Buffalo"="ny"
  "Danbury"="ny";"Oneonta"="ny";"Poughkeepsie"="ny";"Utica"="ny"
  # Others
  "Evansville/Owensboro"="others";"Grand Junction"="others"
  "Missoula"="others";"Montrose"="others";"Shelby"="others"
  "Sioux Falls"="others";"St. Cloud"="others";"St Cloud"="others"
  "Tuscaloosa"="others"
}
$regionNames = @{
  "kathi"="Kathi Kirkland";"taylor"="Taylor Wheeler"
  "tylerw"="Tyler Wille";"jeroen"="Jeroen Corver"
  "nne"="NNE";"nj"="NJ";"ny"="NY";"others"="Others"
}
$aliases = @{
  "Ft Collins"="Fort Collins";"Odessa-Midland"="Odessa";"Killeen-Temple"="Killeen/Temple"
  "Evansville-Owensboro"="Evansville/Owensboro";"Faribault"="Faribault/Owatonna"
  "Rochester"="Rochester MN";"Quincy_Hannibal"="Quincy/Hannibal";"St George"="St. George"
}
$monthLongNames = @{
  1="January";2="February";3="March";4="April";5="May";6="June"
  7="July";8="August";9="September";10="October";11="November";12="December"
}
$monthKeys = @{
  "January"=1;"February"=2;"March"=3;"April"=4;"May"=5;"June"=6
  "July"=7;"August"=8;"September"=9;"October"=10;"November"=11;"December"=12
}
$skipMarkets = @("NABCO","Backyard","Powell","Reno","Atlantic City")

function Normalize([string]$m) {
  $m = $m.Trim()
  if ($aliases.ContainsKey($m)) { return $aliases[$m] }
  return $m
}
function Get-Region([string]$m) {
  if ($regionMap.ContainsKey($m)) { return $regionMap[$m] }
  $n = Normalize $m
  if ($regionMap.ContainsKey($n)) { return $regionMap[$n] }
  return "direct"
}
function Clean-Num([string]$s) {
  $s = ($s -replace '[\$,"\s]','').Trim()
  if ($s -eq '' -or $s -eq '-') { return 0.0 }
  try { return [double]$s } catch { return 0.0 }
}

# Proper CSV tokenizer that respects quoted fields
function Split-CsvLine([string]$line) {
  $tokens = @()
  $i = 0
  while ($i -lt $line.Length) {
    if ($line[$i] -eq '"') {
      # Quoted field
      $j = $line.IndexOf('"', $i + 1)
      if ($j -lt 0) { $j = $line.Length - 1 }
      $tokens += $line.Substring($i + 1, $j - $i - 1)
      $i = $j + 1
      if ($i -lt $line.Length -and $line[$i] -eq ',') { $i++ }
    } else {
      $j = $line.IndexOf(',', $i)
      if ($j -lt 0) { $j = $line.Length }
      $tokens += $line.Substring($i, $j - $i)
      $i = $j + 1
    }
  }
  return $tokens
}

# Derive the 3 active months from the week date
function Get-WeekMonths([string]$weekDate) {
  $dt = [datetime]::ParseExact($weekDate,'yyyy-MM-dd',$null)
  $result = @()
  for ($i = 0; $i -lt 3; $i++) {
    $mo = (($dt.Month - 1 + $i) % 12) + 1
    $label = $monthLongNames[$mo]
    $result += [ordered]@{ key="m$($i+1)"; label=$label; abbr=$label.Substring(0,3); woa_month=$mo }
  }
  return $result
}

# Blueprint CSV — returns @{ monthLabels=@(); markets=@{name=@{m1=;m2=;m3=}} }
# where m1/m2/m3 correspond to whatever 3 months Blueprint has in its columns
function Parse-Blueprint([string]$path) {
  $markets = @{}; $bpMonthLabels = @()
  $lines = Get-Content $path -Encoding UTF8
  $inIgnite = $false; $headerSeen = $false

  foreach ($line in $lines) {
    $line = $line.Trim()
    # TYPE header: extract the 3 month labels Blueprint uses
    if ($line -match '^TYPE,' -and $bpMonthLabels.Count -eq 0) {
      $parts = $line -split ','
      foreach ($p in ($parts[1..3])) {
        $label = ($p.Trim().Trim('"') -replace '\s+\d{4}$','').Trim()
        # Capitalize first letter
        if ($label.Length -gt 1) { $label = $label.Substring(0,1).ToUpper() + $label.Substring(1).ToLower() }
        if ($monthKeys.ContainsKey($label)) { $bpMonthLabels += $label }
      }
    }
    if ($line -match '^IGNITE,')              { $inIgnite = $true; $headerSeen = $false; continue }
    if ($inIgnite -and $line -match '^Pending Pitches') { continue }
    if ($inIgnite -and $line -match '^MARKET,') { $headerSeen = $true; continue }
    if ($inIgnite -and $headerSeen) {
      if ($line -match '^(BROADCAST|AMPED|EVENTS|STD|OLR|TSI|"2026)') { break }
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      # Use proper CSV tokenizer (handles mixed quoted/unquoted fields)
      $tokens = Split-CsvLine $line
      if ($tokens.Count -ge 4) {
        $market = Normalize $tokens[0].Trim()
        if ($market -ne '' -and $market -ne 'MARKET') {
          $markets[$market] = @{
            m1=(Clean-Num $tokens[1])
            m2=(Clean-Num $tokens[2])
            m3=(Clean-Num $tokens[3])
          }
        }
      }
    }
  }
  return @{ monthLabels=$bpMonthLabels; markets=$markets }
}

# WO Analytics TSV — extract adds by month number
function Parse-WO([string]$path) {
  $result = @{}
  $bytes = [System.IO.File]::ReadAllBytes($path)
  $text  = [System.Text.Encoding]::Unicode.GetString($bytes)
  $lines = ($text -split "`r`n|`n") | Where-Object { $_.Trim() -ne '' }
  if ($lines.Count -lt 4) { return $result }
  $header = $lines[2] -split '\t'
  $colMap = @{}
  for ($c = 0; $c -lt $header.Count; $c++) {
    $v = $header[$c].Trim()
    if ($v -match '^\d+$') { $mn = [int]$v; if (-not $colMap.ContainsKey($mn)) { $colMap[$mn] = $c } }
  }
  for ($i = 3; $i -lt $lines.Count; $i++) {
    $cols = $lines[$i] -split '\t'
    if ($cols.Count -lt 3) { continue }
    $dateVal = $cols[0].Trim(); $marketVal = $cols[1].Trim()
    if ($dateVal -notmatch '^\d+/\d+/\d+') { continue }
    if ($marketVal -eq '' -or $marketVal -eq 'Grand Total' -or $marketVal -eq 'Total') { continue }
    $market = Normalize $marketVal
    if ($skipMarkets -contains $market) { continue }
    $entry = @{}
    foreach ($mn in $colMap.Keys) {
      $ci = $colMap[$mn]
      $entry[$mn] = if ($ci -lt $cols.Count) { Clean-Num $cols[$ci] } else { 0.0 }
    }
    $result[$market] = $entry
  }
  return $result
}

# Build the final week JSON using derived months
function Build-WeekJson([string]$weekDate, [array]$months, [hashtable]$bpData, [hashtable]$woData) {
  $byRegion = @{ kathi=@(); taylor=@(); tylerw=@(); jeroen=@(); nne=@(); nj=@(); ny=@(); others=@() }

  # Always include every market from the master list, in defined order
  foreach ($rid in @("kathi","taylor","tylerw","jeroen","nne","nj","ny","others")) {
    foreach ($market in $masterMarkets[$rid]) {
      $mdata = [ordered]@{ name=$market }

    foreach ($mo in $months) {
      $mk  = $mo.key
      $wn  = $mo.woa_month
      $lbl = $mo.label

      # Commit: find the Blueprint slot that matches this month label
      $commit = 0.0
      if ($bpData.markets.ContainsKey($market)) {
        for ($j = 0; $j -lt $bpData.monthLabels.Count; $j++) {
          if ($bpData.monthLabels[$j] -eq $lbl) {
            $bpKey = "m$($j+1)"
            $commit = $bpData.markets[$market][$bpKey]
            break
          }
        }
      }

      # Adds: WO Analytics by month number
      $adds = 0.0
      if ($woData.ContainsKey($market) -and $woData[$market].ContainsKey($wn)) {
        $adds = $woData[$market][$wn]
      }

      $mdata[$mk] = [ordered]@{ commit=[math]::Round($commit,2); adds=[math]::Round($adds,2) }
    }
      $byRegion[$rid] += $mdata
    }  # end market loop
  }  # end region loop

  $dt = [datetime]::ParseExact($weekDate,'yyyy-MM-dd',$null)
  $weekLabel = "$($monthLongNames[$dt.Month]) $($dt.Day), $($dt.Year)"
  $regionOrder = @("kathi","taylor","tylerw","jeroen","nne","nj","ny","others")
  $regionList  = $regionOrder | ForEach-Object { [ordered]@{ id=$_; name=$regionNames[$_]; markets=$byRegion[$_] } }

  return [ordered]@{
    lastRefreshed=(Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    weekOf=$weekLabel; weekDate=$weekDate; months=$months; regions=$regionList
  }
}

# ── Main ───────────────────────────────────────────────────────────────────────
$weeks = @()
foreach ($folder in (Get-ChildItem $WeeklyRoot -Directory | Sort-Object Name)) {
  $weekDate = $folder.Name
  Write-Host "Week: $weekDate"
  $bpFile = Get-ChildItem $folder.FullName "Blueprint*.csv" -ErrorAction SilentlyContinue | Select-Object -Last 1
  $woFile = Get-ChildItem $folder.FullName "Sales Revenue.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $bpFile) { Write-Warning "  No Blueprint file, skipping"; continue }

  $months = Get-WeekMonths $weekDate
  $bp     = Parse-Blueprint $bpFile.FullName
  $wo     = if ($woFile) { Parse-WO $woFile.FullName } else { @{} }  # WO optional
  if (-not $woFile) { Write-Host "  No WO file - adds will be zero" }

  $mLabels = ($months | ForEach-Object { $_.label }) -join ', '
  Write-Host "  Months: $mLabels  BP months: $($bp.monthLabels -join ',')  Commits: $($bp.markets.Count)  WO: $($wo.Count)"

  $json = Build-WeekJson $weekDate $months $bp $wo
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
