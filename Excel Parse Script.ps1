$ExcelFile = '.\Calabrio Clean Up.xlsx'
$ExcelSheet = 'Calabrio'
$CSVFile = '.\Calabrio Clean Up - Working.csv'

$ImportedFile = Import-Excel -path $ExcelFile -WorksheetName $ExcelSheet | Where-Object `
{ $_."Employee ID" -like $null `
        -and $_."Deactivated" -like $null `
        -and $_."First Name" -notlike $null `
        -and $_."Last Name" -notlike $null }

Import-Excel `
    -path $ExcelFile `
    -WorksheetName $ExcelSheet | Export-Csv `
    -path $CSVFile

$FixedCSV = Import-Csv -path $CSVFile

foreach ($User in $ImportedFile) {
    $Name = "$($User."First Name") $($User."Last Name")"

    $UserInfo = Get-ADUser -Filter "Name -eq '$Name' -and Enabled -eq 'True'" -Properties mail, EmployeeID, SAMAccountName, givenName, sn `
    | Select-Object givenName, sn, mail, EmployeeID, `
    @{l="SAMAccountName";e={"HQ\" + $_.SAMAccountName}}

    $FixedCSV | Where-Object {"$($_."Login")" -eq "$($UserInfo.SAMAccountName)"} | foreach {
        $_."Employee ID" = $UserInfo.EmployeeID
        $_."Email" = $UserInfo.mail
        $_."Login" = $UserInfo.SAMAccountName
    } 
}

$FixedCSV | Export-Excel -Path '.\Calabrio Clean Up - Separate Names.xlsx' -WorksheetName $ExcelSheet