# IGI Command Center -- DSM Report Parser
# Reads the DSM Report Excel and outputs data/dsm.json
# Usage: .\parse-dsm.ps1

param(
  [string]$WeeklyRoot = "C:\Users\justi\.openclaw\workspace\_Ignite Growth Intelligence\Data\Weekly",
  [string]$InboxPath  = "C:\Users\justi\.openclaw\workspace\_Ignite Growth Intelligence\Data\_Inbox",
  [string]$OutputDir  = "C:\Users\justi\.openclaw\workspace\igi-command-center\data"
)

# DSM name clean-up: "BRADFORD, CRYSTAL K - X15573" -> "Crystal Bradford"
function Clean-DsmName([string]$sheetName) {
  $part = ($sheetName -split ' - ')[0].Trim()  # "BRADFORD, CRYSTAL K"
  $split = $part -split ',\s*'
  if ($split.Count -ge 2) {
    $last  = $split[0].Trim()
    $first = ($split[1].Trim() -split '\s+')[0].Trim()
    $last  = $last.Substring(0,1) + $last.Substring(1).ToLower()
    $first = $first.Substring(0,1) + $first.Substring(1).ToLower()
    return "$first $last"
  }
  return $sheetName
}

# DSM -> markets mapping (from Market Managers 2025 / ORG_STRUCTURE.md)
$dsmMarkets = @{
  "Crystal Bradford"    = @("Killeen/Temple","Tyler","Victoria")
  "Trey Bufkin"         = @("Lubbock")
  "Colton Bybee"        = @("Sierra Vista","St. George","Williston")
  "Bryce Clemens"       = @("Poughkeepsie")
  "Josh Cox"            = @("Grand Rapids")
  "Todd Cross"          = @("Faribault/Owatonna","Rochester MN")
  "Nicole Daily"        = @("Tri-Cities","Wenatchee","Yakima")
  "Jennelle Diggs"      = @("Amarillo","Lawton","Odessa","San Angelo","Wichita Falls")
  "Hillary Doyal"       = @("Lake Charles","Shreveport","Texarkana")
  "Jennifier Francis"   = @("Lafayette")
  "Christina Hawkins"   = @("Twin Falls")
  "Steve Horinka"       = @("Bismarck","Quad Cities","Quincy/Hannibal")
  "Nicholas Ineck"      = @("Billings","Bozeman")
  "Chelsea Jones"       = @("Evansville/Owensboro")
  "Alixzandra Jyawook"  = @("Lansing")
  "Kelly Katoski"       = @("Duluth")
  "Kathryn Kirkland"    = @("Flint","Grand Rapids","Kalamazoo","Killeen/Temple","Lafayette","Lake Charles","Lansing","Lufkin","Rockford","Shreveport","Texarkana","Tyler","Victoria")
  "Jeffrey Klein"       = @("Lufkin")
  "Jed Knapp"           = @("El Paso")
  "Paige Lauback"       = @("Binghamton","Oneonta","Utica")
  "Jason Longley"       = @("Boise")
  "Scott Mauser"        = @("Evansville/Owensboro")
  "Michael Miller"      = @("Missoula")
  "William Prieto"      = @("Buffalo")
  "Natalie Redding"     = @("Grand Junction","Montrose")
  "Alyssa Salisbury"    = @("Rockford")
  "Diana Scully"        = @("Shore","Trenton/Princeton")
  "Michelle Sellers"    = @("Butte","Great Falls")
  "John Shea"           = @("Albany","Berkshires")
  "Ryan Sheehy"         = @("New Bedford","Portsmouth")
  "Tyler Tholl"         = @("Augusta","Portland")
  "Angela Todd"         = @("Kalamazoo")
  "Tony Townsend"       = @("Cedar Rapids","Dubuque","Sedalia","Waterloo")
  "Jilian Watson"       = @("Flint")
  "Bryan Wheeler"       = @("Casper","Cheyenne","Fort Collins","Laramie")
  "Joshua Whinery"      = @("Bangor")
}

function Parse-Num([string]$s) {
  $s = $s.Trim() -replace '[\$,\s]',''
  $neg = ($s -match '^\(') -or ($s -match '^-')
  $s = $s -replace '[^0-9.]',''
  if ($s -eq '' -or $s -eq '-') { return 0.0 }
  $v = try { [double]$s } catch { 0.0 }
  if ($neg) { return -$v } else { return $v }
}

function Get-Period($sheet, $row) {
  return [ordered]@{
    jan = Parse-Num $sheet.Cells.Item($row, 8).Text
    feb = Parse-Num $sheet.Cells.Item($row, 9).Text
    mar = Parse-Num $sheet.Cells.Item($row, 10).Text
    q1  = Parse-Num $sheet.Cells.Item($row, 11).Text
    apr = Parse-Num $sheet.Cells.Item($row, 12).Text
    may = Parse-Num $sheet.Cells.Item($row, 13).Text
    jun = Parse-Num $sheet.Cells.Item($row, 14).Text
    q2  = Parse-Num $sheet.Cells.Item($row, 15).Text
    jul = Parse-Num $sheet.Cells.Item($row, 16).Text
    aug = Parse-Num $sheet.Cells.Item($row, 17).Text
    sep = Parse-Num $sheet.Cells.Item($row, 18).Text
    q3  = Parse-Num $sheet.Cells.Item($row, 19).Text
  }
}

# Find the row for IGNITE Total Revenue | Total (dynamically, since row varies by DSM)
function Find-TotalRevenueRow($sheet) {
  for ($r = 47; $r -le 120; $r++) {
    $prod   = $sheet.Cells.Item($r, 4).Text.Trim()
    $metric = $sheet.Cells.Item($r, 5).Text.Trim()
    $agg    = $sheet.Cells.Item($r, 6).Text.Trim()
    # Match IGNITE (exact) Total Revenue Total - not IGNITE DISPLAY, SEM, STV etc.
    if ($prod -eq 'IGNITE' -and $metric -eq 'Total Revenue' -and $agg -eq 'Total') { return $r }
  }
  return 59  # fallback
}

# Find the DSM file -- check weekly folders first, then inbox
function Find-DsmFile {
  # Check weekly folders (newest first)
  foreach ($folder in (Get-ChildItem $WeeklyRoot -Directory | Sort-Object Name -Descending)) {
    $f = Get-ChildItem $folder.FullName -Filter "*DSM*" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($f) { return $f }
  }
  # Fall back to inbox
  $f = Get-ChildItem $InboxPath -Filter "*DSM*" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if ($f) { return $f }
  return $null
}

$dsmFile = Find-DsmFile
if (-not $dsmFile) { Write-Error "No DSM file found"; exit 1 }
Write-Host "Using: $($dsmFile.Name)"

# Skip if output is newer than source
$outPath = Join-Path $OutputDir "dsm.json"
if (Test-Path $outPath) {
  if ($dsmFile.LastWriteTime -le (Get-Item $outPath).LastWriteTime) {
    Write-Host "No changes -- skipping"
    exit 0
  }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($dsmFile.FullName)

$dsmList = @()
foreach ($sheet in $wb.Sheets) {
  $name = $sheet.Name
  # DSM sheets have format "LASTNAME, FIRSTNAME M - CODE"
  if ($name -notmatch '^[A-Z]+,\s') { continue }

  $dsmName = Clean-DsmName $name
  Write-Host "  $dsmName"

  $soloBudget  = Get-Period $sheet 34
  $soloPacing  = Get-Period $sheet 35
  $totalRevRow  = Find-TotalRevenueRow $sheet
  $totalRevenue = Get-Period $sheet $totalRevRow

  # Force proper JSON array (PowerShell serializes single-item arrays as strings)
  $rawMarkets = if ($dsmMarkets.ContainsKey($dsmName)) { $dsmMarkets[$dsmName] } else { @() }
  $markets = [System.Collections.Generic.List[string]]::new()
  foreach ($m in $rawMarkets) { $markets.Add($m) }

  $dsmList += [ordered]@{
    sheetName    = $name
    name         = $dsmName
    markets      = $markets
    soloBudget   = $soloBudget
    soloPacing   = $soloPacing
    totalRevenue = $totalRevenue
  }
}

$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Sort alphabetically by first name
$dsmList = $dsmList | Sort-Object { $_.name }

$output = [ordered]@{
  lastRefreshed = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
  dsmFile       = $dsmFile.Name
  dsms          = $dsmList
}

$output | ConvertTo-Json -Depth 10 | Set-Content $outPath -Encoding UTF8
Write-Host "`nWritten: $outPath ($($dsmList.Count) DSMs)"
