# IGI Command Center — SOB Parser
# Reads the weekly SOB-IGNITE Excel and outputs data/sob.json
# Usage: .\parse-sob.ps1 [-SobPath "path/to/file.xlsx"]

param(
  [string]$SobPath   = "",
  [string]$OutputDir = "C:\Users\justi\.openclaw\workspace\igi-command-center\data"
)

# Auto-find most recent SOB file in inbox if not specified
if (-not $SobPath) {
  $inbox = "C:\Users\justi\.openclaw\workspace\_Ignite Growth Intelligence\Data\_Inbox"
  $file  = Get-ChildItem $inbox -Filter "*.xlsx" | Where-Object { $_.Name -match "SOB" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $file) { Write-Error "No SOB file found in inbox"; exit 1 }
  $SobPath = $file.FullName
  Write-Host "Using: $($file.Name)"
}

# Region master list (same as commit tracker)
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

# Map market name → Excel sheet name
$sheetMap = @{
  "Killeen/Temple"="Killeen-Temple"; "Fort Collins"="Ft Collins"
  "St. George"="St George"; "Faribault/Owatonna"="Faribault"
  "Quincy/Hannibal"="Quincy_Hannibal"; "Rochester MN"="Rochester"
  "Trenton/Princeton"="Trenton_Princeton"; "Evansville/Owensboro"="Evansville"
  "St. Cloud"="St Cloud"; "Missoula"="Missoula"
}

function Get-SheetName([string]$market) {
  if ($sheetMap.ContainsKey($market)) { return $sheetMap[$market] }
  return $market
}

function Parse-Pct([string]$s) {
  $s = $s.Trim().TrimEnd('%').Replace('(','').Replace(')','')
  if ($s -eq '' -or $s -eq '-') { return $null }
  try {
    $v = [double]$s
    # If original had parens, it's negative
    if ($s -notmatch '-' -and $s.Length -gt 0) {
      return $v  # Already positive
    }
    return $v
  } catch { return $null }
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
  return @{
    pacing   = Parse-Num $sheet.Cells.Item($baseRow + 1, $col).Text
    budget   = Parse-Num $sheet.Cells.Item($baseRow + 5, $col).Text
    pctBgt   = Parse-PctFull $sheet.Cells.Item($baseRow + 7, $col).Text
    forecast = Parse-Num $sheet.Cells.Item($baseRow + 8, $col).Text
    pctFct   = Parse-PctFull $sheet.Cells.Item($baseRow + 10, $col).Text
    py       = Parse-Num $sheet.Cells.Item($baseRow + 11, $col).Text
    pctPY    = Parse-PctFull $sheet.Cells.Item($baseRow + 13, $col).Text
  }
}

# Open Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false; $excel.DisplayAlerts = $false
$wb = $excel.Workbooks.Open($SobPath)

# Get week label from TSQ sheet
$tsq = $wb.Sheets.Item("TSQ")
$weekLabel = $tsq.Cells.Item(14,2).Text
Write-Host "SOB Week: $weekLabel"

$regionOrder = @("kathi","taylor","tylerw","jeroen","nne","nj","ny","others")
$regionList  = @()

foreach ($rid in $regionOrder) {
  $markets = @()
  foreach ($market in $masterMarkets[$rid]) {
    $sheetName = Get-SheetName $market
    $mData = [ordered]@{ name=$market; q2=[ordered]@{}; q3=[ordered]@{} }

    # Try to find the sheet
    $sheet = $null
    try { $sheet = $wb.Sheets.Item($sheetName) } catch { }

    if ($sheet) {
      # Find Ignite Revenue-Cash row (should be row 20, but search to be safe)
      $baseRow = 20
      for ($r = 18; $r -le 25; $r++) {
        if ($wb.Sheets.Item($sheetName).Cells.Item($r,1).Text -match 'Ignite Revenue' -or
            $wb.Sheets.Item($sheetName).Cells.Item($r,2).Text -match 'Ignite Revenue') {
          $baseRow = $r; break
        }
      }

      # Q2: Apr=col7, May=col8, Jun=col9, Q2total=col10
      $mData.q2["apr"]   = Get-MonthData $sheet $baseRow 7
      $mData.q2["may"]   = Get-MonthData $sheet $baseRow 8
      $mData.q2["jun"]   = Get-MonthData $sheet $baseRow 9
      $mData.q2["total"] = Get-MonthData $sheet $baseRow 10

      # Q3: Jul=col11, Aug=col12, Sep=col13, Q3total=col14
      $mData.q3["jul"]   = Get-MonthData $sheet $baseRow 11
      $mData.q3["aug"]   = Get-MonthData $sheet $baseRow 12
      $mData.q3["sep"]   = Get-MonthData $sheet $baseRow 13
      $mData.q3["total"] = Get-MonthData $sheet $baseRow 14

      Write-Host "  $market (sheet=$sheetName): Q2=$($mData.q2.total.pacing)K  Q3=$($mData.q3.total.pacing)K"
    } else {
      Write-Warning "  ${market}: sheet '$sheetName' not found"
    }
    $markets += $mData
  }
  $regionList += [ordered]@{ id=$rid; name=$regionNames[$rid]; markets=$markets }
}

$wb.Close($false); $excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$output = [ordered]@{
  lastRefreshed = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
  weekLabel     = $weekLabel
  sobFile       = (Split-Path $SobPath -Leaf)
  regions       = $regionList
}

$outPath = Join-Path $OutputDir "sob.json"
$output | ConvertTo-Json -Depth 10 | Set-Content $outPath -Encoding UTF8
Write-Host "`nWritten: $outPath"
