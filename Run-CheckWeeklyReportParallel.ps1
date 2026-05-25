param(
  [Parameter(Mandatory=$true)][string]$WorkbookPath,
  [Parameter(Mandatory=$true)][string]$TargetFolder,
  [int]$Workers = 2,
  [string]$MacroName = "CheckXlsxFiles_FromList",
  [string]$ResultCsvPath
)

$ErrorActionPreference = "Stop"

if (!(Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
if (!(Test-Path $TargetFolder)) { throw "TargetFolder not found: $TargetFolder" }
if ($Workers -lt 1) { $Workers = 1 }

$xlsxFiles = Get-ChildItem -Path $TargetFolder -Recurse -File -Filter *.xlsx | Select-Object -ExpandProperty FullName
if ($xlsxFiles.Count -eq 0) {
  Write-Host "No xlsx files found."
  exit 0
}

if ($Workers -gt $xlsxFiles.Count) { $Workers = $xlsxFiles.Count }

$base = Join-Path $env:TEMP ("checkweekly_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $base | Out-Null

$chunks = @()
for ($i = 0; $i -lt $Workers; $i++) { $chunks += ,(New-Object System.Collections.ArrayList) }

for ($i = 0; $i -lt $xlsxFiles.Count; $i++) {
  $bucket = $i % $Workers
  [void]$chunks[$bucket].Add($xlsxFiles[$i])
}

$listFiles = @()
$outFiles = @()
for ($i = 0; $i -lt $Workers; $i++) {
  $listPath = Join-Path $base ("list_{0}.txt" -f $i)
  $outPath  = Join-Path $base ("out_{0}.csv" -f $i)
  $chunks[$i] | Set-Content -Path $listPath -Encoding UTF8
  $listFiles += $listPath
  $outFiles += $outPath
}

$jobs = @()
for ($i = 0; $i -lt $Workers; $i++) {
  $lp = $listFiles[$i]
  $op = $outFiles[$i]
  $wb = $WorkbookPath
  $mn = $MacroName

  $jobs += Start-Job -ScriptBlock {
    param($WorkbookPath, $MacroName, $ListFile, $OutFile)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
      $wb = $excel.Workbooks.Open($WorkbookPath)
      $excel.Run($MacroName, $ListFile, $OutFile)
      $wb.Close($false)
    }
    finally {
      $excel.Quit()
      [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
      [gc]::Collect()
      [gc]::WaitForPendingFinalizers()
    }
  } -ArgumentList $wb, $mn, $lp, $op
}

$jobs | Wait-Job | Out-Null
$failed = $jobs | Where-Object { $_.State -ne 'Completed' }
if ($failed) {
  $failed | Receive-Job
  throw "One or more worker jobs failed."
}

$rows = @()
foreach ($file in $outFiles) {
  if (Test-Path $file) {
    $rows += Import-Csv -Path $file -Header Path,Status,Message
  }
}

$resultCsv = if ([string]::IsNullOrWhiteSpace($ResultCsvPath)) {
  Join-Path $base "merged_result.csv"
}
else {
  $ResultCsvPath
}

$resultCsvDir = Split-Path -Path $resultCsv -Parent
if (![string]::IsNullOrWhiteSpace($resultCsvDir) -and !(Test-Path $resultCsvDir)) {
  New-Item -ItemType Directory -Path $resultCsvDir -Force | Out-Null
}

$rows | Export-Csv -Path $resultCsv -NoTypeInformation -Encoding UTF8

Write-Host "Done."
Write-Host "Input files : $($xlsxFiles.Count)"
Write-Host "Workers     : $Workers"
Write-Host "Result CSV  : $resultCsv"
Write-Host "Temp folder : $base"
