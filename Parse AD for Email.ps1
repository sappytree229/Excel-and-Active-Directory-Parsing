$ExcelFile = '.\Calabrio Clean Up- Finished.xlsx'
$ExcelSheet = 'Calabrio'
$CSVFile = '.\AD Email Query.csv'

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
    $Email = $User."Email"

    $UserInfo = Get-ADUser -Filter {UserPrincipalName -eq "$Email"} -Properties mail, EmployeeID, SAMAccountName, UserPrincipalName `
    | Select-Object UserPrincipalName, EmployeeID, @{l="SAMAccountName";e={"HQ\" + $_.SAMAccountName}}

    $FixedCSV | Where-Object {$Email -eq $UserInfo.EmailAddress} | ForEach {
        $_."Employee ID" = $UserInfo.EmployeeID 
        $_."Login" = $UserInfo.SAMAccountName
    } 
}

$FixedCSV | Export-Excel -Path '.\AD 2nd Fix Calabrio Clean Up- Finished.xlsx' -WorksheetName $ExcelSheet