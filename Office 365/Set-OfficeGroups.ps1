function Set-OfficeGroups {
    # View_user_group function will display user's current groups and distribution lists every time it is called
    function view_user_group {
        Write-Host "User $dist_email_name is a member of the following:"
        $user_info_output = Get-User $dist_email_name | Select-Object -ExpandProperty DistinguishedName
        Write-Host 'Groups' -ForegroundColor Green
        Get-Recipient -Filter "Members -eq '$user_info_output'" -RecipientTypeDetails GroupMailbox | Select-Object Alias, PrimarySmtpAddress | Out-Host
        Write-Host 'Distribution Lists' -ForegroundColor Green
        Get-Recipient -Filter "Members -eq '$user_info_output'" -RecipientTypeDetails MailUniversalDistributionGroup, MailUniversalSecurityGroup | Sort-Object Name | Select-Object Name, PrimarySmtpAddress | Out-Host
    }

    function view_contact_group {
        Write-Host "User $dist_email_name is a member of the following:"
        $user_info = Get-MailContact $dist_email_name | Select-Object DistinguishedName
        $user_info_output = $user_info.DistinguishedName
        Write-Host 'Groups' -ForegroundColor Green
        Write-Host 'Distribution Lists' -ForegroundColor Green
        Get-DistributionGroup -ResultSize unlimited -Filter "Members -like '$user_info_output'" | Sort-Object Name | Select Name, PrimarySmtpAddress | Out-Host
    }

    function add_contact_group {
        Clear-Host
        Write-Host 'You have chosen to add contacts to groups/distribution lists' -ForegroundColor Green
        $dist_email_name = (Get-MailContact | Select DisplayName, PrimarySmtpAddress | Sort-Object DisplayName | Out-GridView -PassThru -Title 'Contacts').PrimarySmtpAddress
        if ($dist_email_name -eq '' -or $dist_email_name -eq $null) {
            return
        }

        else {
            # Display user's current groups
            view_contact_group

            # Declare empty array to store distribution list names
            # Out-GridView all distribution lists and write to $dist_temp if selected
            $dist_temp = @()
            $dist_temp = Get-DistributionGroup -SortBy Name | Select-Object DisplayName,PrimarySmtpAddress | Out-GridView -Title 'Add user $dist_email_name to the following distribution lists' -PassThru

            if ($dist_temp -eq $null) {
                return
            }
                
            else {
                # Create new $dist_name array and write the selected distribution lists' PrimarySmtpAddress to $dist_name
                $dist_name = @()
                $dist_name = @($dist_temp.PrimarySmtpAddress)

                # Nested loop to add each contact to each selected distlist
                foreach ($contact in $dist_email_name) {
                    foreach ($distlist in $dist_name) {
                        Add-DistributionGroupMember -Identity $distlist -Member $contact
                    }
                } 

                view_contact_group
                pause
            }
        }
    }

    function remove_contact_group {  
        Clear-Host
        Write-Host 'You have chosen to remove a contact from groups/distribution lists' -ForegroundColor Green
        $dist_email = Get-MailContact | Select DisplayName, PrimarySmtpAddress | Sort-Object DisplayName | Out-GridView -PassThru -Title 'Contacts'
        if ($dist_email -eq '' -or $dist_email -eq $null) {
            return
        }

        else {
            # Declare an empty array to store group names
            $dist_email_name = $dist_email.PrimarySmtpAddress
            $user_info = Get-MailContact $dist_email_name | Select-Object DistinguishedName
            $user_info_output = $user_info.DistinguishedName
            $dist_temp = @()
            $dist_temp = Get-Recipient -Filter "Members -eq '$user_info_output'" -RecipientTypeDetails GroupMailbox | Select-Object Alias, PrimarySmtpAddress | Out-GridView -Title 'Current groups' -PassThru
            
            if ($dist_temp -eq $null) {
                #do nothing
            }

            else {
                
                $dist_name = @()
                $dist_name = @($dist_temp.Alias)

                # Loop until the end of the array to remove all listed groups
                foreach ($contact in $dist_email_name) {
                    foreach ($distlist in $dist_name) {
                        Remove-UnifiedGroupLinks -Identity $distlist -LinkType Member -Links $contact -Confirm:$false
                    }
                }
            }
                

            # Declare empty array to store distribution group names
            $dist_temp = @()
            $dist_temp = Get-DistributionGroup -ResultSize unlimited -Filter "Members -like '$user_info_output'" | Select DisplayName, PrimarySmtpAddress | Out-GridView -Title "Distribution lists for $dist_email_name" -PassThru

            if ($dist_temp -eq $null) {
                #do nothing
            }

            else {
                $dist_name = @()
                $dist_name = @($dist_temp.PrimarySmtpAddress)

                #Loop through array to remove selected distribution groups
                foreach ($contact in $dist_email_name) {
                    foreach ($distlist in $dist_name) {
                        Remove-DistributionGroupMember -Identity $distlist -Member $contact -Confirm:$false
                    }
                }

                view_contact_group
                pause
            }
        }
    }  

    function view_all_group {
        Clear-Host
        Write-Host 'Groups' -ForegroundColor Green
        Get-UnifiedGroup | Select-Object Alias, PrimarySmtpAddress | Format-Table | Out-Host
        Write-Host 'Distribution Lists' -ForegroundColor Green
        Get-DistributionGroup | Select-Object DisplayName, PrimarySmtpAddress | Format-Table | Out-Host
        pause
    }

    # Add user to groups/distribution lists
    function add_user_group {
        Clear-Host
        Write-Host 'You have chosen to add user to groups/distribution lists' -ForegroundColor Green
        #$dist_email_name = Read-Host 'What is the full email of the user?'
        $dist_email_name = (Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "User to add to groups").UserPrincipalName
        if ($dist_email_name -eq '' -or $dist_email_name -eq $null) {
            return
        }

        else {
            # Display user's current groups
            view_user_group

            # Declare an empty array to store group names
            # Out-GridView all domain groups and write to $dist_temp if selected
            $dist_temp = @()
            $dist_temp = Get-UnifiedGroup -SortBy Alias | Select-Object Alias, DistinguishedName, PrimarySmtpAddress | Out-GridView -Title "Add user $dist_email_name to the following groups" -PassThru

            if ($dist_temp -eq $null) { 
                #do nothing
            }

            else {
                # Create new $dist_name array and write the selected groups' Alias to $dist_name
                $dist_name = @()
                $dist_name = @($dist_temp.Alias)

                # Loop on $dist_name to add the user into the groups using Add-UnifiedGroupLinks
                for ($i = 0; $i -lt ($dist_name.Count); $i++) {
                    Add-UnifiedGroupLinks -Identity $dist_name[$i] -LinkType Member -Links $dist_email_name
                }
            }
            # Declare empty array to store distribution list names
            # Out-GridView all distribution lists and write to $dist_temp if selected
            $dist_temp = @()
            $dist_temp = Get-DistributionGroup -SortBy Name | Select-Object DisplayName,PrimarySmtpAddress | Out-GridView -Title "Add user $dist_email_name to the following distribution lists" -PassThru

            if ($dist_temp -eq $null) {
                return
            }
                
            else {
                # Create new $dist_name array and write the selected distribution lists' PrimarySmtpAddress to $dist_name
                $dist_name = @()
                $dist_name = @($dist_temp.PrimarySmtpAddress)

                # Loop on $dist_name to add the user into the distribution lists using Add-DistributionGroupMember
                for ($i = 0; $i -lt ($dist_name.Count); $i++) {
                    Add-DistributionGroupMember -Identity $dist_name[$i] -Member $dist_email_name
                }

                view_user_group
                pause
            }
        }
    }

    # Remove user from groups/distribution lists
    function remove_user_group {  
        Clear-Host
        Write-Host 'You have chosen to remove a user from groups/distribution lists' -ForegroundColor Green
        #$dist_email_name = Read-Host 'What is the full email of the user?'
        $dist_email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -OutputMode Single -Title "User to remove from distribution lists"
        if ($dist_email_name -eq '' -or $dist_email_name -eq $null) {
            return
        }

        else {
            $dist_email = $dist_email_name.UserPrincipalName
            $user_info_output = Get-User -Identity $dist_email | Select-Object -ExpandProperty DistinguishedName

            # Declare an empty array to store group names
            $dist_temp = @()
            $dist_temp = Get-Recipient -Filter "Members -eq '$user_info_output'" -RecipientTypeDetails GroupMailbox | Select-Object Alias, PrimarySmtpAddress | Out-GridView -Title 'Current groups' -PassThru
            
            if ($dist_temp -eq $null) {
                #do nothing
            }

            else {
                $dist_name = @()
                $dist_name = @($dist_temp.Alias)

                # Loop until the end of the array to remove all listed groups
                for ($i = 0; $i -lt ($dist_name.Count); $i++) {
                    Remove-UnifiedGroupLinks -Identity $dist_name[$i] -LinkType Member -Links $dist_email_name -Confirm:$false
                }
            }

            #Declare empty array to store distribution group names
            $dist_temp = @()
            $dist_temp = Get-Recipient -Filter "Members -eq '$user_info_output'" -RecipientTypeDetails MailUniversalDistributionGroup, MailUniversalSecurityGroup | Select-Object Name, PrimarySmtpAddress | Out-GridView -Title 'Current Distribution Lists' -PassThru

            if ($dist_temp -eq $null) {
                #do nothing
            }

            else {
                $dist_name = @()
                $dist_name = @($dist_temp.Name)

                #Loop through array to remove selected distribution groups
                for ($i = 0; $i -lt ($dist_name.Count); $i++) {
                    Remove-DistributionGroupMember -Identity $dist_name[$i] -Member $dist_email_name -Confirm:$false
                }

                view_user_group
                pause
            }
        }
    }  
        
    function view_selected_contact_group {
        Clear-Host
        Write-Host 'You have chosen to view a contacts'' distribution lists' -ForegroundColor Green
        $dist_email_name = Read-Host 'What is the full email of the user?'
        if ($dist_email_name -eq '' -or $dist_email_name -eq $null) {
            return
        }

        else {
            view_contact_group
            pause
        }
    }

    # Function to view a selected user's groups/distribution lists
    function view_selected_user_group {
        Clear-Host
        Write-Host 'You have chosen to view a user''s distribution lists.' -ForegroundColor Green
        $dist_email_name = Read-Host 'What is the full email of the user?'
        if ($dist_email_name -eq '' -or $dist_email_name -eq $null) {
            return
        }

        else {
            view_user_group
            pause
        }
    }

    do {
        $group_menu = {
            Clear-Host
            Write-Host 'You have chosen to manage distribution groups' -ForegroundColor Green
            Write-Host '1: View users distribution groups'
            Write-Host '2: Add users to distribution groups'
            Write-Host '3: Remove users from distribution groups'
            Write-Host '4: View contacts distribution groups'
            Write-Host '5: Add contacts to distribution groups'
            Write-Host '6: Remove contacts from distribution groups'
            Write-Host '7: View all groups/distribution lists'
            Write-Host '------------------------'
            Write-Host 'Q: Return to main menu'
            $menu_input = Read-Host 'Make a selection'
            switch ($menu_input) {
                '2' {
                    add_user_group
                    view_group_menu
                }
                '3' {
                    remove_user_group
                    view_group_menu 
                }
                '1' {
                    view_selected_user_group
                    view_group_menu   
                }
                '4' {
                    view_selected_contact_group
                    view_group_menu
                }
                '5' {
                    add_contact_group
                    view_group_menu
                }
                '6' {
                    remove_contact_group
                    view_group_menu
                }
                '7' {
                    view_all_group
                }
                'q' {
                    return
                    Clear-Host
                }
                default {
                    Write-Warning "Please enter a valid selection"
                    pause
                    .$group_menu
                }
            }
        }
        .$group_menu
    }
    until ($menu_input -eq 'q')
}
