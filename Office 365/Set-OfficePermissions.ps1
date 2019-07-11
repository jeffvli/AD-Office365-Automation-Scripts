function Set-OfficePermissions {
    <#
    .SYNOPSIS
        Assign or remove calender/mailbox permissions on Office 365

    .DESCRIPTION
        Allows bulk and single set calendar/mailbox permissions for users on Office 365
    #>

    # View calendar permissions of a specified user
    function calendar_view_permissions {
        Write-Host 'You have chosen to view a user''s calendar permissions' -ForegroundColor Green
        $email_name = $null
        $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "Users in $company"
        $email_name_selection = @($email_name.UserPrincipalName)

        if ($email_name -eq $null) {
            return
        }

        else {
            foreach ($email in $email_name_selection) {
                Write-Host "Calendar permissions for $email" -ForegroundColor Green
                Get-MailboxFolderPermission -Identity ($email + ':\calendar') | Select-Object Folder, User, AccessRights, SharingPermissionFlags | Format-Table
            }
        }
    }

    # Add calendar permissions for a specified user
        function calendar_add_permissions {
        Write-Host 'You have chosen to add calendar permissions' -ForegroundColor Green
        $email_name = $null
        $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "$company : Users' calendar you want to share"
        $email_name_selection = @($email_name.UserPrincipalName)

        if ($email_name -eq $null) {
            return
        }
            
        else {
            foreach ($email in $email_name_selection) {
                Write-Host "Calendar permissions for $email" -ForegroundColor Green
                Get-MailboxFolderPermission -Identity ($email + ':\calendar') | Select-Object Folder, User, AccessRights, SharingPermissionFlags
            }
                
            $email_name_give = $null
            $email_name_give = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "$company : Users you want to share calendar to"
            $email_permission_name = @($email_name_give.UserPrincipalName)

            if ($email_name_give -eq $null) {
                return
            }

            else {   
                $access_rights = Read-Host 'Enter the access rights (owner, editor, reviewer, availabilityonly)'
                foreach ($email in $email_name_selection) {
                    foreach ($email_give in $email_permission_name) {
                        Add-MailboxFolderPermission -Identity ($email + ':\calendar') -User $email_give -AccessRights $access_rights -Confirm
                    }
                }
            }
        }
    }

    function calendar_add_delegate_permissions {
        Write-Host 'You have chosen to add calendar editor-delegate permissions' -ForegroundColor Green
        $email_name = $null
        $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "$company : Users' calendar you want to share"
        $email_name_selection = @($email_name.UserPrincipalName)

        if ($email_name -eq $null) {
            return
        }
            
        else {
            foreach ($email in $email_name_selection) {
                Write-Host "Calendar permissions for $email" -ForegroundColor Green
                Get-MailboxFolderPermission -Identity ($email + ':\calendar') | Select-Object Folder, User, AccessRights, SharingPermissionFlags
            }
                
            $email_name_give = $null
            $email_name_give = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "$company : Users you want to share calendar to"
            $email_permission_name = @($email_name_give.UserPrincipalName)

            if ($email_name_give -eq $null) {
                return
            }

            else {   
                foreach ($email in $email_name_selection) {
                    foreach ($email_give in $email_permission_name) {
                        Add-MailboxFolderPermission -Identity ($email + ':\calendar') -User $email_give -AccessRights Editor -SharingPermissionFlags Delegate -Confirm

                    }
                }
            }
        }
    }

        # Remove calendar permissions for a specified user
        function calendar_remove_permissions {
        Write-Host 'You have chosen to remove calendar permissions' -ForegroundColor Green
        $email_name = $null
        $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -OutputMode Single -Title "$company : Users' mailbox to remove permissions from"
        $email_name_selection = @($email_name.UserPrincipalName)

        if ($email_name -eq $null) {
            return
        }

        else {
            Write-Host 'Current calendar permissions'
            foreach ($email in $email_name_selection) {
                Get-MailboxFolderPermission -Identity ($email + ':\calendar') | Select-Object Folder, User, AccessRights, SharingPermissionFlags | Format-Table
                $email_name_remove = Get-MailboxFolderPermission -Identity ($email + ':\calendar') | ForEach-Object {$_.User.ADRecipient.UserPrincipalName} | Out-GridView -PassThru -Title "Current permissions on $email_name. Remove which ones?"
            }
            
            $email_permission_name = @($email_name_remove)
            $access_rights = @($email_name_remove.AccessRights)
            $sharing_flags = @($email_name_remove.SharingPermissionFlags)

            foreach ($email in $email_name_selection) {
                for ($i = 0; $i -lt $email_permission_name.Count; $i++) {
                    Remove-MailboxFolderPermission -Identity ($email + ':\calendar') -User $email_permission_name[$i] -Confirm
                }   
            }
        }
    }
    
    # View mailbox permissions for a specified user
    function mailbox_view_permissions {
        Write-Host 'You have chosen to view a user''s mailbox permissions' -ForegroundColor Green
        $email_name = $null
        $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "Users in $company"
        $email_name_selection = @($email_name.UserPrincipalName)

        if ($email_name_selection -eq $null) {
            return
        }

        else {
            foreach ($email in $email_name_selection) {
                Write-Host "Mailbox permissions for $email" -ForegroundColor Green
                Get-MailboxPermission -Identity $email | Format-Table
            }
        }
    }

    # Add mailbox permissions for a specified user
    function mailbox_add_permissions {
        Write-Host 'You have chosen to add mailbox permissions' -ForegroundColor Green
        $email_name = $null
        $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "$company : Users' mailbox you want to share"
        $email_name_selection = @($email_name.UserPrincipalName)

        if ($email_name -eq $null) {
            return
        }
            
        else {
            foreach ($email in $email_name_selection) {
                Write-Host "Mailbox permissions for $email" -ForegroundColor Green
                Get-MailboxPermission -Identity $email | Format-Table
            }

            $email_name_give = $null
            $email_name_give = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -PassThru -Title "$company : Users you want to share mailbox to"
            $email_permission_name = @($email_name_give.UserPrincipalName)

            if ($email_name_give -eq $null) {
                return
            }

            else {
                $access_rights = Read-Host 'Enter access rights (fullaccess, readpermission)'
                foreach ($email in $email_name_selection) {
                    foreach ($email_give in $email_permission_name) {
                        Add-MailboxPermission -Identity $email -User $email_give -AccessRights $access_rights -Confirm
                    }
                }
            }
        }
        pause   
    }

    # Remove mailbox permissions for a specified user
    function mailbox_remove_permissions {
        Write-Host 'You have chosen to remove mailbox permissions' -ForegroundColor Green
        $email_name = $null
        $email_name = Get-MsolUser -all | Sort-Object DisplayName | Select-Object DisplayName, UserPrincipalName, Licenses | Out-GridView -OutputMode Single -Title "$company : Users' mailbox to remove permissions from"
        $email_name_selection = @($email_name.UserPrincipalName)

        if ($email_name -eq $null) {
            return
        }

        else {
            Write-Host 'Current mailbox permissions'
            foreach ($email in $email_name_selection) {
                Get-MailboxPermission -Identity $email| Format-Table
                $email_name_remove = Get-MailboxPermission -Identity $email | Where-Object {$_.User -like '*@*' -and $_.User -ne $email} | Out-GridView -PassThru -Title "Current permissions on $email_name"
            }
            
            $email_permission_name = @($email_name_remove.User)
            $access_rights = @($email_name_remove.AccessRights)

            foreach ($email in $email_name_selection) {
                for ($i = 0; $i -lt $email_permission_name.Count; $i++) {
                    Remove-MailboxPermission -Identity $email -User $email_permission_name[$i] -AccessRights $access_rights[$i] -Confirm
                }   
            }
        }
    }
    do {
        Clear-Host
        Write-Host 'You have chosen to manage Office 365 permissions' -ForegroundColor Green
        Write-Host '1: View mailbox permissions'
        Write-Host '2: Add mailbox permissions'
        Write-Host '3: Remove mailbox permissions'
        Write-Host '4: View calendar permissions'
        Write-Host '5: Add calendar permissions'
        Write-Host '6: Remove calendar permissions'
        Write-Host '------------------------' 
        Write-Host '7: Add calendar editor-delegate permissions (lets user manage appointments)'
        Write-Host '------------------------'
        Write-Host 'Q: Exit and remove PSSession'

        $selection_temp = Read-Host 'Make a selection'
        switch ($selection_temp) {
            '1' {
                Clear-Host
                mailbox_view_permissions
                pause
            }
            '2' {
                Clear-Host
                mailbox_add_permissions
            }
            '3' {
                Clear-Host
                mailbox_remove_permissions
            }
            '4' {
                Clear-Host
                calendar_view_permissions
                pause
            }
            '5' {
                Clear-Host
                calendar_add_permissions
            }
            '6' {
                Clear-Host
                calendar_remove_permissions
            }
            '7' {
                Clear-Host
                calendar_add_delegate_permissions
            }
            'q' {
                return
            }
            default {
                Write-Warning 'Invalid selection...'
                pause
            }
        }
    }
    until ($selection_temp -eq 'q')
}
