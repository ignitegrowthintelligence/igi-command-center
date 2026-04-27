# IGI Command Center -- SOB Parser
# Scans weekly folders for SOB files, outputs per-week sob-{date}.json + sob.json (most recent)
# Usage: .\parse-sob.ps1

param(
  [string]$WeeklyRoot = "C:\Users\justi\.openclaw\workspace\_Ignite Growth Intelligence\Data\Weekly",
  [string]$OutputDir  = "C:\Users\justi\.openclaw\workspace\igi-command-center\data"
)

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
$regionNames = @{
  "kathi"="Kathi Kirkland";"taylor"="Taylor Wheeler";"tylerw"="Tyler Wille";"jeroen"="Jeroen Corver"
  "nne"="NNE";"nj"="NJ";"ny"="NY";"others"="Others"
}
$sheetMap = @{
  "Killeen/Temple"="Killeen-Temple"; "Fort Collins"="Ft Collins"
  "St. George"="St George"; "Faribault/Owatonna"="Faribault"
  "Quincy/Hannibal"="Quincy_Hannibal"; "Rochester MN"="Rochester"
  "Trenton/Princeton"="Trenton_Princeton"; "Evansville/Owensboro"="Evansville"
  "St. Cloud"="St Cloud"
}

function Get-SheetName([string]$market) {
  if ($sheetMap.ContainsKey($market)) { return $sheetMap[$market] }
  return $market
}

function Parse-Num([string]$s) {
  $s = $s.Trim().Replace(',','')
  $neg = ($s -match '^\(') -or ($s -match '^-')
  $s = $s -replace '[^0-9.]',''
  if ($s -eq '') { return 0.0 }
  $v = try { [double]$s } catch { 0.0 }
  if ($neg) { return -$v } else { return $v }
}

function Parse-PctFull([string]$s) {
  $s = $s.Trim()
  $neg = $s -match '^\('
  $s = $s -replace '[^0-9.]',''
  if ($s -eq '') { return $null }
  $v = try { [double]$s } catch { return $null }
  if ($neg) { return -$v } else { return $v }
}

function Get-MonthData($sheet, $baseRow, $col) {
  return [ordered]@{
    pacing   = Parse-Num     $sheet.Cells.Item($baseRow + 1, $col).Text
    budget   = Parse-Num     $sheet.Cells.Item($baseRow + 5, $col).Text
    pctBgt   = Parse-PctFull $sheet.Cells.Item($baseRow + 7, $col).Text
    forecast = Parse-Num     $sheet.Cells.Item($baseRow + 8, $col).Text
    pctFct   = Parse-PctFull $sheet.Cells.Item($baseRow + 10, $col).Text
    py       = Parse-Num     $sheet.Cells.Item($baseRow + 11, $col).Text
    pctPY    = Parse-PctFull $sheet.Cells.Item($baseRow + 13, $col).Text
  }
}

function Parse-SOB-File([string]$sobPath, [string]$weekDate) {
  Write-Host "  Parsing: $(Split-Path $sobPath -Leaf)"
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false; $excel.DisplayAlerts = $false
  $wb = $excel.Workbooks.Open($sobPath)

  $tsq = $wb.Sheets.Item("TSQ")
  $weekLabel = $tsq.Cells.Item(14,2).Text

  $regionOrder = @("kathi","taylor","tylerw","jeroen","nne","nj","ny","others")
  $regionList  = @()

  foreach ($rid in $regionOrder) {
    $markets = @()
    foreach ($market in $masterMarkets[$rid]) {
      $sheetName = Get-SheetName $market
      $mData = [ordered]@{ name=$market; q2=[ordered]@{}; q3=[ordered]@{} }
      $sheet = $null
      try { $sheet = $wb.Sheets.Item($sheetName) } catch { }

      if ($sheet) {
        $baseRow = 20
        for ($r = 18; $r -le 25; $r++) {
          if ($sheet.Cells.Item($r,1).Text -match 'Ignite Revenue' -or
              $sheet.Cells.Item($r,2).Text -match 'Ignite Revenue') { $baseRow = $r; break }
        }
        $mData.q2["apr"]   = Get-MonthData $sheet $baseRow 7
        $mData.q2["may"]   = Get-MonthData $sheet $baseRow 8
        $mData.q2["jun"]   = Get-MonthData $sheet $baseRow 9
        $mData.q2["total"] = Get-MonthData $sheet $baseRow 10
        $mData.q3["jul"]   = Get-MonthData $sheet $baseRow 11
        $mData.q3["aug"]   = Get-MonthData $sheet $baseRow 12
        $mData.q3["sep"]   = Get-MonthData $sheet $baseRow 13
        $mData.q3["total"] = Get-MonthData $sheet $baseRow 14
        $mData["fy"]       = Get-MonthData $sheet $baseRow 19
      }
      $markets += $mData
    }
    $regionList += [ordered]@{ id=$rid; name=$regionNames[$rid]; markets=$markets }
  }

  $wb.Close($false); $excel.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

  return [ordered]@{
    lastRefreshed = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    weekLabel     = $weekLabel
    weekDate      = $weekDate
    sobFile       = (Split-Path $sobPath -Leaf)
    regions       = $regionList
  }
}

# -- Main: scan weekly folders for SOB files --
$sobWeeks = @()
$latestDate = ""

foreach ($folder in (Get-ChildItem $WeeklyRoot -Directory | Sort-Object Name)) {
  $weekDate = $folder.Name
  $sobFile  = Get-ChildItem $folder.FullName -Filter "*SOB*.xlsx" -ErrorAction SilentlyContinue | Select-Object -First 1
  if (-not $sobFile) { continue }

  $outPath = Join-Path $OutputDir "sob-$weekDate.json"
  # Skip if output JSON is newer than the SOB source file
  if (Test-Path $outPath) {
    $outTime = (Get-Item $outPath).LastWriteTime
    if ($sobFile.LastWriteTime -le $outTime) {
      Write-Host "Week: $weekDate -- no changes, skipping"
      $existing = Get-Content $outPath | ConvertFrom-Json
      $sobWeeks += [ordered]@{ date=$weekDate; label=$existing.weekLabel; sobFile=$sobFile.Name }
      $latestDate = $weekDate
      continue
    }
  }

  Write-Host "Week: $weekDate"
  $json = Parse-SOB-File $sobFile.FullName $weekDate
  $utf8NoBom = New-Object System.Text.UTF8Encoding $false
  $jsonContent = $json | ConvertTo-Json -Depth 10
  [System.IO.File]::WriteAllText($outPath, $jsonContent, $utf8NoBom)
  Write-Host "  -> $outPath"
  $sobWeeks += [ordered]@{ date=$weekDate; label=$json.weekLabel; sobFile=$json.sobFile }
  $latestDate = $weekDate
}

# Write SOB weeks index
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
$sobIndex = [ordered]@{ weeks=($sobWeeks | Sort-Object { $_.date } -Descending) }
$indexContent = $sobIndex | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText((Join-Path $OutputDir "sob-weeks.json"), $indexContent, $utf8NoBom)

# Copy most recent as sob.json
if ($latestDate) {
  Copy-Item (Join-Path $OutputDir "sob-$latestDate.json") (Join-Path $OutputDir "sob.json") -Force
  Write-Host "sob.json -> $latestDate"
}

Write-Host "Done. $($sobWeeks.Count) SOB weeks processed."
