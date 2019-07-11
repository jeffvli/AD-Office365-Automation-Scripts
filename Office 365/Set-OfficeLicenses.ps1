### Office_license_modify manages licensing for Office 365 users
function Set-OfficeLicenses {
    do {
        Clear-Host
        Write-Host 'You have chosen to manage Office 365 licenses' -ForegroundColor Green
        Write-Host '1: View available licenses'
        Write-Host '2: Add licenses to user'
        Write-Host '3: Remove licenses from user'
        Write-Host '------------------------'
        Write-Host 'Q: Return to main menu'
        $menu_input = Read-Host 'Make a selection'

        switch ($menu_input) {
            '1' {
                Clear-Host
                Write-Host "Licenses for $company"
                Get-MsolAccountSku | Sort-Object AccountSkuID | Format-Table
                pause
            }
            '2' {            
                $email_name = $null
                $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "Users in $company"
                $email_name_selection = @($email_name.UserPrincipalName)

                if ($email_name -eq $null) {
                    return
                }

                else {
                    $license_name = $null
                    $license_name = Get-MsolAccountSku | Sort-Object AccountSkuID | Out-GridView -PassThru -Title "Licenses available in $company"
                    $license_name_selection = @($license_name.AccountSkuID)
                    if ($license_name -eq $null) {
                        return
                    }
                    
                    else {
                        foreach ($user in $email_name_selection) {
                            foreach ($license in $license_name_selection) {
                                Set-MsolUserLicense -UserPrincipalName $user -AddLicenses $license
                            }
                        }
                    }
                }
            }  
            '3' {
                $email_name = $null
                $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "Users in $company"
                $email_name_selection = @($email_name.UserPrincipalName)

                if ($email_name -eq $null) {
                    return
                }

                else {
                    $license_name = $null
                    $license_name = Get-MsolAccountSku | Sort-Object AccountSkuID | Out-GridView -PassThru -Title "Licenses available in $company"
                    $license_name_selection = @($license_name.AccountSkuID)
                    if ($license_name -eq $null) {
                        return
                    }
                    
                    else {
                        foreach ($user in $email_name_selection) {
                            foreach ($license in $license_name_selection) {
                                Set-MsolUserLicense -UserPrincipalName $user -RemoveLicenses $license
                            }
                        }
                    }
                }
            }
            'q' {
                return
            }
            Default {
                Write-Warning 'Invalid selection...'
                pause
            }
        }
    }
    until ($menu_input -eq 'q')
}
