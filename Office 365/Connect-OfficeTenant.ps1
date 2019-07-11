function Connect-OfficeTenant {
    [CmdletBinding()]
    param()

    Begin {
        function Input-OfficeCredential {
            if (!($UserCredential)) {
                $script:UserCredential = Get-Credential
	        }
	
            else {
                $InputSelection = Read-Host "Continue with current login as"$UserCredential.Username"? [Y/n/q]"
                if ($InputSelection -eq $null -or $InputSelection.ToLower() -eq 'y') {
                    # do nothing
                }
			
                if ($InputSelection.ToLower() -eq 'q') {
                    break
                }
			
                if ($InputSelection.ToLower() -eq 'n')  {
                    Get-PSSession | Remove-PSSession
                    $script:UserCredential = Get-Credential
                }
            }
        }

        Input-OfficeCredential
        Connect-MsolService -Credential $UserCredential -ErrorAction Stop
        $Tenants = Get-MsolPartnerContract -ErrorAction Stop
        $IndexCount = 1
        $TenantObject = @()
        $TenantID = $null
    }

    Process {
        foreach ($Tenant in $Tenants) {
            $TenantObject += New-Object -TypeName PSObject -Property @{
                Index = $IndexCount
                Name = $Tenant.Name
                TenantID = $Tenant.TenantID
            }
		
            $IndexCount++
        }

        # List tenants
        $TenantObject | Format-Table -Property @{Name='Index'; Expression={$_.Index}; Alignment='Center';}, Name, TenantID

        do {
            [int]$TenantSelection = Read-Host "Which tenant do you want to connect to? [#]"

            if ($TenantSelection -eq '') {
                Write-Warning 'Exiting...'
		Get-PSSession | Remove-PSSession
                break
            }

            if ($TenantSelection -ne (1..$Count)) {
                Write-Warning 'Select a valid index.'
            }
        } 
        
        while ($TenantSelection -ne (1..$Count) -or $TenantSelection -eq $null)
        
        $TenantSelection = $Tenants[$TenantSelection-1]
        $TenantID = $TenantSelection.TenantID.Guid
        Write-Output "Connecting to"$TenantSelection.Name"..."
        $Session = New-PSSession `
            -ConfigurationName Microsoft.Exchange `
            -ConnectionUri https://outlook.office365.com/powershell-liveid/?DelegatedOrg=$TenantID `
            -Credential $UserCredential `
            -Authentication Basic `
            -AllowRedirection:$true

        Import-PSSession $Session -AllowClobber
    }
}
