function Remove-ADUser {
    <#
    .SYNOPSIS
      Name: Remove-ADUser.ps1
      The purpose of this script is to automate the  termination process.

    .DESCRIPTION
     This script is used to automate the termination process at . Given parameters: 

    .NOTES
      Version 1.0.0

      Release Date: 08-09-2018

      Author: Jeff Li

    .EXAMPLE
      Example 1: Remove user jli@domain.com and delegate shared mailbox access to mattm@domain.com
      PS C:\> Connect-ADOffice
      PS C:\> Remove-ADUser -EmailName jli@domain.com -Delegate mattm@domain.com

      Example 2: Remove user testuser@domain.com without delegating email
      PS C:\> Connect-ADOffice
      PS C:\> Remove-ADUser -EmailName testuser@domain.com
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$EmailName,
        [string]$Delegate,
        [string]$Domain
    )
        # Switch selection to determine if you want to continue the process after this point
        Get-User -Identity $EmailName | Select-Object DisplayName, UserPrincipalName | Out-Host
        $selection = Read-Host "Are you sure you want to terminate $EmailName? (y/n)"
        switch ($selection) {
            'y' {
                # Split the inputted email name to obtain the AD username
                $UserName = $EmailName.split('@')[0]

                # Create a random password using the format "Goodbye<date><symbol>"
                $PassDate = Get-Date -format 'MMddyyyy'
                $PassSymbol = Get-Random -InputObject !, ?
                $Pass = "Goodbye" + $PassDate + $PassSymbol

                #Set-MsolUserPassword will set the password as the previously generated one
                Set-MsolUserPassword -UserPrincipalName $EmailName -NewPassword $Pass
                Write-Host 'User password reset'

                # Add current user distribution groups to a temp array and loop to remove
                Write-Host 'Removing user groups/distlists...'
                $DistinguishedName = Get-User -Identity $EmailName | Select-Object -ExpandProperty DistinguishedName
                $Groups = Get-Recipient -Filter "Members -eq '$DistinguishedName'" -RecipientTypeDetails GroupMailbox | Select-Object Alias, PrimarySmtpAddress
                $DistributionLists = Get-Recipient -Filter "Members -eq '$DistinguishedName'" -RecipientTypeDetails MailUniversalDistributionGroup, MailUniversalSecurityGroup | Select-Object Name, PrimarySmtpAddress
                        
                foreach ($Group in $Groups.Alias) {
                    Remove-UnifiedGroupLinks -Identity $Group -Links $EmailName -LinkType Members -Confirm:$false
                }

                foreach ($List in $DistributionLists.Name) {
                    Remove-DistributionGroupMember -Identity $List -Member $EmailName -Confirm:$false
                }
                            
                Write-Host 'User groups/distlists removed'

                # Remove mailbox forwarding
                Write-Host 'Removing email forwarding...'
                $Forwarding = Get-Mailbox -Identity $EmailName | Where-Object {$_.ForwardingAddress -ne $null}

                if ($null -ne $Forwarding) {
                    Set-Mailbox $EmailName -ForwardingAddress $null
                    Write-Host 'Email forwarding removed'
                }

                else {
                    Write-Host 'Email forwarding not found'
                }

                # Remove mobile devices
                Write-Host 'Removing mobile devices...'
				$Mobile = Get-MobileDevice -mailbox $EmailName | Select-Object DeviceModel, Identity
				if ($null -ne $Mobile) {
					foreach ($Device in $Mobile) {
						$Device | Remove-MobileDevice -Identity $_.Identity -Confirm:$false
					}
					Write-Host 'Mobile devices removed...'
				}

				else {
					Write-Host 'No mobile devices found'
				}

                # Convert to shared mailbox
                Write-Host 'Converting to shared mailbox... (may take some time)'
                Set-Mailbox $EmailName -Type Shared

                # Assign temp variable to null and have it write RecipientTypeDetails to another variable
                $Temp = $null
                $Temp = Get-MailBox -Identity $EmailName
                $EmailDetails = $Temp.RecipientTypeDetails 
                       
                # Use a while loop to continuously check (every 2 seconds) the temp variable until SharedMailBox type is assigned to it
                while ($EmailDetails -eq 'UserMailbox') {
                    $Temp = Get-Mailbox -Identity $EmailName
                    $EmailDetails = $Temp.RecipientTypeDetails
                    Write-Host 'Currently converting $EmailName to shared mailbox. Please wait...'
                    Start-Sleep -Seconds 1
                }
    
                Write-Host "$EmailName has successfully been converted to shared mailbox"

				if ($Delegate) {
						Write-Host 'You have chosen to delegate mailbox access'
						Write-Host 'Current mailbox permissions'
						Get-MailboxPermission -Identity $EmailName | Format-Table

						# Add-MailboxPermission will set the terminated employee's mailbox to give Ownership rights to the delegation_name
						Add-MailboxPermission -Identity $EmailName -User $Delegate -AccessRights FullAccess -Confirm:$False

						# Write user licenses into object and remove them with a ForEach loop
						Write-Host 'Removing licenses...'
						(get-MsolUser -UserPrincipalName $EmailName).licenses.AccountSkuId |
						ForEach-Object {
							Set-MsolUserLicense -UserPrincipalName $EmailName -RemoveLicenses $_
						}
						Write-Host 'Removed user licenses'
						Write-Host 'Blocking user account...'

						# Block user account login
						Set-MsolUser -UserPrincipalName $EmailName -BlockCredential $true
						Write-Host 'Blocked user account'
						Write-Host "You have successfully terminated $EmailName!"
				}

				else {
					# Write user licenses into object and remove them with a ForEach loop
					Write-Host 'Removing user licenses...'
					(Get-MsolUser -UserPrincipalName $EmailName).Licenses.AccountSkuId |
					ForEach-Object {
						Set-MsolUserLicense -UserPrincipalName $EmailName -RemoveLicenses $_
					}
					Write-Host 'Removed user licenses'

					# Block user account login
					Set-MsolUser -UserPrincipalName $EmailName -BlockCredential $true
					Write-Host 'Blocked user account...'
					Write-Host "You have successfully terminated $EmailName!"
				}
							
			# Start AD account disable process
			Write-Host "Office 365 account for $EmailName has successfully been terminated!" -ForegroundColor Green
			Write-Host 'Starting AD account termination process...'

			# Write AD user properties into temp variable
			$Temp = Get-ADUser -Credential $ADCredential -Server $Domain -Identity $UserName -Properties Description
			$Description = $temp.Description
			$Date = Get-Date -format 'MM-dd-yyyy'

			# Reset AD password
			# Set AD user description with termination date
			# Clear IP phones from AD profile
			# Disable AD account
			Write-Host 'Resetting AD password...'
			Set-ADAccountPassword -Identity $UserName -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Pass -Force) -Credential $ADCredential -Server $Domain
			Write-Host 'AD password reset'
			Write-Host 'Setting termination date to description...'
			Set-ADUser -Identity $UserName -Description "$Description Terminated $Date" <#-Clear ipPhone#> -Credential $ADCredential -Server $Domain
			Write-Host 'Termination date set'
			Write-Host 'Disabling AD account...'
			Disable-ADAccount -Identity $UserName -Credential $ADCredential -Server $Domain
			Write-Host 'AD account disabled'
			Write-Host 'AD account termination process completed' -ForegroundColor Green
			pause
		}
		'n' {
			return
		}
		
		default {
			Write-Warning 'Invalid selection...'
			pause
		}
	}
}
