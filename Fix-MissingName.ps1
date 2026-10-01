param(
    [string]$CsvPath,
    [string]$StudentDb
)

$csvPath = $CsvPath
$dbPath  = $StudentDb

$rows = Import-Csv $csvPath
$db   = Import-Csv $dbPath -Delimiter "`t" -Header Barcode,MemberNumber,LastName,Names,Address,Email,Info1,Info2,Info3,Info4,Notes,Facility,Category,Status

foreach ($row in $rows) {
    if ([string]::IsNullOrWhiteSpace($row.'First Name')) {
        $match = $db | Where-Object { $_.Email -ieq $row.Email } | Select-Object -First 1
        if ($match) {
            $row.'First Name' = $match.Names
            $row.'Last Name'  = $match.LastName
        }
    }
}

$rows | Export-Csv $csvPath -NoTypeInformation
# strip the extra quotes Export-Csv adds around every field
(Get-Content $csvPath) -replace '"', '' | Set-Content $csvPath