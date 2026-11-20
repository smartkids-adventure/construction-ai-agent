param (
    [string]$Python = ".\.venv\Scripts\python.exe",
    [string]$DemoDirectory = ".\local\windows-commit-test"
)
$ErrorActionPreference = 'Stop'
$env:PYTHONPATH = (Get-Location).Path
$prepared = & $Python .\tests\make_windows_commit_case.py $DemoDirectory
if ($LASTEXITCODE -ne 0) { throw 'Synthetic commit setup failed.' }
$result = $prepared | ConvertFrom-Json
$actualHash = (Get-FileHash -LiteralPath $result.output -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actualHash -ne $result.output_sha) { throw 'Committed workbook fingerprint differs.' }

$excel = New-Object -ComObject Excel.Application
$book = $null
try {
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($result.output, 0, $true)
    $excel.CalculateFullRebuild()
    foreach ($sheet in $result.reconciliation.PSObject.Properties) {
        $worksheet = $book.Worksheets.Item(([int]$sheet.Name + 1))
        foreach ($expected in $sheet.Value.expected.PSObject.Properties) {
            $actual = $worksheet.Range($expected.Name).Value2
            if ($null -eq $expected.Value) {
                if ($null -ne $actual -and $actual -ne '') { throw "Expected blank at $($expected.Name)." }
            } elseif ($actual -isnot [double] -or [Math]::Abs([double]$actual - [double]$expected.Value) -gt 0.000001) {
                throw "Reconciliation failed at $($expected.Name)."
            }
        }
    }
    $site = $book.Worksheets.Item(6)
    if ($site.Range('K5').Value2 -ne 18 -or $site.Range('L5').Value2 -ne 7) {
        throw 'Reported and verified quantities did not remain separate.'
    }
    if ($null -ne $site.Range('L6').Value2) { throw 'Unverified quantity was populated.' }
    $readiness = $book.Worksheets.Item(20)
    foreach ($cell in 'F5','G5','L5','M5','N5') {
        if ($null -ne $readiness.Range($cell).Value2) {
            throw "Protected readiness field $cell was changed."
        }
    }
    Write-Output 'Desktop Excel commit acceptance passed.'
}
finally {
    if ($null -ne $book) { $book.Close($false) }
    $excel.Quit()
}
