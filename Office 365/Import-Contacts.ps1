# Connect to Office and get Exchange Cmdlets
# Require MSOnline PowerShell module (Install-Module MSOnline -Force)

function Connect-Office {
    if (Get-PSSession) {
        $temp = Read-Host 'You are already logged in Office, continue with current account? (y/n)'
        switch ($temp) {
            'y' {
                return
            }
            'n' {
                Get-PSSession | Remove-PSSession
            }
            default {
                Write-Warning 'Invalid selection...'
                Connect-Office
            }
        }
    }
    $script:OfficeCredential = Get-Credential -Message 'Enter Office 365 login credentials'
    if ($null -eq $OfficeCredential -or $OfficeCredential -eq '') {
        return
    }
    else {
        try {
            Connect-MsolService -Credential $OfficeCredential -ErrorAction Stop
        }
        catch {
            Write-Warning 'Office credentials are invalid...' -ErrorAction Stop
            pause
            return
        }
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OfficeCredential -Authentication Basic -AllowRedirection
        Import-PSSession $Session
    }
}

Connect-Office

# Enter file path to .csv files with columns, FirstName, LastName, DisplayName, Email, Company
$ContactCSV = Import-Csv -Path ''

# Distribution lists to add to
$DistList = @(
)

# Import contacts into Office 365
$ContactCSV | ForEach-Object {New-MailContact -Name $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -DisplayName $_.DisplayName -ExternalEmailAddress $_.Email}

# Set contacts company attribute to their respectful companies
$ContactCSV | ForEach-Object {Set-Contact -Identity $_.Email -Company $_.Company}

Start-Sleep -Seconds 10

# Add contacts to $DistList
foreach ($User in $ContactCSV) {
    foreach ($Dist in $DistList) {
        Add-DistributionGroupMember -Identity $Dist -Member $User.'Email'
        Write-Host "$User added to $Dist"
    }
}
