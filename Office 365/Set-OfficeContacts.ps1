function Set-OfficeContacts {
    function contact_add {
        Write-Host 'You have chosen to add a new contact' -ForegroundColor Green
        $contact_name = Read-Host 'Enter the full name of the user'
        if ($contact_name -eq '' -or $contact_name -eq $null) {
            return
        }

        else {
            $contact_first_name, $contact_last_name = $contact_name.Split(' ')
            $contact_email = Read-Host 'Enter the full email of the user'
            $company_name = Read-Host 'Enter the company/location of the user' 
            # New-MailContact will add a new Office 365 contact given contact name and email
            New-MailContact -Name $contact_name -ExternalEmailAddress $contact_email -FirstName $contact_first_name -LastName $contact_last_name -Confirm
            Set-MailContact -Identity $contact_email -Company $company_name   
        }
    }
        
    function contact_remove {
        Write-Host 'You have chosen to remove an existing contact' -ForegroundColor Green
        #$contact_name = Read-Host 'Enter full email of user'
        $contact_selection = Get-MailContact -SortBy Name | Select-Object Name, PrimarySmtpAddress | Out-GridView -Title "Contacts for $company" -PassThru
        if ($contact_selection -eq $null) {
            return
        }

        else {
            $contact_name = @($contact_selection.PrimarySmtpAddress)
            # Remove-MailContact will remove an Office 365 contact given contact email
            foreach ($cont in $contact_name) {
                Remove-MailContact -Identity $cont -Confirm:$true
            }
        }
   
    }

    function contact_search {
        Write-Host 'You have chosen to search contacts' -ForegroundColor Green
        $contact_name = Read-Host "Enter the name of the user"
        if ($contact_name -eq '' -or $contact_name -eq $null) {
            return
        }

        else {
            # Get-MailContact will search contacts by name and display Name, PrimarySmtpAddress in a table format
            Get-MailContact -Identity "*$contact_name*" -SortBy Name | Select-Object Name, PrimarySmtpAddress | Out-GridView -Title "Contacts for $company"
        }
            
    }

    function contact_view_all {
        Write-Host 'You have chosen to view all contacts' -ForegroundColor Green
        Get-MailContact -SortBy Name | Select-Object Name, PrimarySmtpAddress | Out-GridView -Title "Contacts for $company"
            
    }

    do {
        Clear-Host
        Write-Host 'You have chosen to manage Office 365 contacts' -ForegroundColor Green
        Write-Host '1: View all contacts'
        Write-Host '2: Search contacts'
        Write-Host '3: Add a new contact'
        Write-Host '4: Remove an existing contact'
        Write-Host '------------------------'
        Write-Host 'Q: Return to main menu'
        $selection_input = Read-Host "Make your selection"

        # Switch on selection_input to choose which function to run
        switch ($selection_input) {
            '1' {
                contact_view_all
            }
            '2' {
                contact_search
            }
            '3' {
                contact_add
            }
            '4' {
                contact_remove
            }
            'q' {
                return
            }
            default {
            Write-Warning 'Invalid Selection'
            pause
            }
        }
    }
    until ($selection_input -eq 'q') 
}
