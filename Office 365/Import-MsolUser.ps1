# Create a CSV with columns UserPrincipalName, FirstName, LastName, License
$Path = ''
$Import = Import-Csv -Path $Path

foreach ($User in $Import) {
    $DisplayName = $User.Firstname +  ' ' + $User.Lastname
    New-MsolUser -UserPrincipalName $User.UserPrincipalName -FirstName $User.Firstname -LastName $User.LastName -DisplayName $DisplayName -LicenseAssignment $User.License -UsageLocation 'US' -ForceChangePassword $false | Export-Csv 'C:\New-MsolUser.csv' -Append
}
