param (
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)
$ErrorActionPreference = 'Stop'
$excel = New-Object -ComObject Excel.Application
$book = $null

function Assert-Cell([int]$Sheet, $Address, $Expected) {
    $actual = $book.Worksheets.Item($Sheet).Range($Address).Value2
    if ($actual -ne $Expected) {
        throw "$Sheet!$Address expected [$Expected], found [$actual]"
    }
}

try {
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $true)
    Assert-Cell 2 'A1' 'SYNTHETIC CAPACITY TEST ONLY'
    $site = $book.Worksheets.Item(6)
    if ($site.ListObjects.Item(1).ListRows.Count -ne 1000) {
        throw 'Expected 1,000 rows in the migrated daily register.'
    }
    foreach ($row in 5, 105, 1004) {
        $site.Range("A$row").Value2 = "SYN-$row"
        $site.Range("D$row").Value2 = 'SYN-A'
        $site.Range("E$row").Value2 = 'SYN-BQ'
        $site.Range("I$row").Value2 = 'm'
        $site.Range("K$row").Value2 = [double]30
        $site.Range("L$row").Value2 = [double]2
        $site.Range("R$row").Value2 = 'KVDT2-WF-000001'

        $progress = $book.Worksheets.Item(11)
        $progress.Range("A$row").Value2 = 'SYN-A'
        $progress.Range("E$row").Value2 = 'm'
        $quantity = $book.Worksheets.Item(13)
        $quantity.Range("A$row").Value2 = 'SYN-BQ'
        $quantity.Range("C$row").Value2 = 'm'

        $readiness = $book.Worksheets.Item(20)
        $readiness.Range("A$row").Value2 = 'KVDT2-WF-000001'
        $readiness.Range("F$row").Value2 = 'Blocked'
    }
    $excel.CalculateFullRebuild()
    Assert-Cell 1 'B5' 3
    foreach ($row in 5, 105, 1004) {
        Assert-Cell 11 "H$row" 6
        Assert-Cell 13 "E$row" 6
        Assert-Cell 20 "O$row" 'Blocked'
    }
    Write-Output 'Excel boundary and recalculation checks passed.'
}
finally {
    if ($null -ne $book) {
        $book.Close($false)
    }
    $excel.Quit()
}
