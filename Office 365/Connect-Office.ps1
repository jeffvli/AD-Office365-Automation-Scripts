# Connect to Office only

# Add domain controller
$DC = ''

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

# Connect to both Office and AD
function Connect-ADOffice {
    Get-PSSession | Remove-PSSession
    $script:OfficeCredential = Get-Credential -Message 'Enter Office 365 login credentials'
    if ($null -eq $OfficeCredential -or $OfficeCredential -eq '') {
        return
    }
    else {
        $script:ADCredential = Get-Credential -Message 'Enter AD domain admin credentials (User@biolase.com)'
        if ($null -eq $ADCredential -or $ADCredential -eq '') {
            return
        }
        else {
            if ($null -eq $ADCredential -or $ADCredential -eq '') {
                Write-Error 'AD credentials required to run script'
            }
            else {
                try {
                    $locked_users = Search-ADAccount -LockedOut -Credential $ADCredential -Server $DC -ErrorAction Stop # Run AD function to check AD credentials
                }
                catch {
                    Write-Warning "AD credentials are invalid..."
                    pause
                    return # Stop script if above fails
                }
            }
            try {
                Connect-MsolService -Credential $OfficeCredential -ErrorAction Stop # Check if Office credentials are valid
            }
            catch {
                Write-Error 'Office credentials are invalid...' -ErrorAction Stop
                pause
                return
            }
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OfficeCredential -Authentication Basic -AllowRedirection
            Import-PSSession $Session
        }
    }
}

function Connect-AD {
    $script:ADCredential = $null
    $script:ADCredential = Get-Credential -Message 'Enter AD domain admin credentials (User@biolase.com)'
    if ($null -eq $ADCredential -or $ADCredential -eq '') {
        return
    }
    else {
        if ($null -eq $ADCredential -or $ADCredential -eq '') {
            Write-Error 'AD credentials required to run script'
        }
        else {
            try {
                $LockedUsers = Search-ADAccount -LockedOut -Credential $ADCredential -Server $DC -ErrorAction Stop # Run AD function to check AD credentials
            }
            catch {
                Write-Warning "AD credentials are invalid..."
                pause
                return # Stop script if above fails
            }
        }
    }
}
