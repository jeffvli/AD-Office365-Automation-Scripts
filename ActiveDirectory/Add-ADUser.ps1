function Add-ADUser {

    <#
    .SYNOPSIS
        Name: Add-ADUser.ps1
        The purpose of this script is to add a new hire to AD/Office 365.

    .DESCRIPTION
        This script is used to automate the new hire process at . Given parameters: FirstName, LastName,
        Description, ADUserCopy, and Expire, the script will create an Office 365 User and AD account. Groups
        and distribution lists for both O365 and AD will be copied from ADUserCopy onto the newly created user.

    .NOTES
        Version 1.0.0

        Release Date: 08-09-2018

        Author: Jeff Li

    .EXAMPLE
        Example 1: Create a new account for a new Account Manager who requires password to not expire since they are remote.
        PS C:\> Connect-ADOffice
        PS C:\> Add-User -FirstName Harold -LastName Tanner -Description 'Account Manager' -ADUserCopy gwalsh -Domain test.com -EmailDomain test.com

        Example 2: Create a new account for an internal HR hire.
        PS C:\> Connect-ADOffice
        PS C:\> Add-User -FirstName Jenny -LastName Smith -Description 'HR Manager' -ADUserCopy hjenkins -Domain test.com -EmailDomain test.com -Expire 

        Example 3: Create account account for internal Warehouse hire
        PS C:\> Connect-ADOffice
        PS C:\> Add-User Connor Luckey 'Warehouse Temp' arivera test.com test.com -Expire
    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$FirstName,

        [Parameter(Mandatory=$true, Position=1)]
        [string]$LastName,

        [Parameter(Mandatory=$true, Position=2)]
        [string]$Description,

        [Parameter(Mandatory=$true, Position=3)]
        [string]$ADUserCopy,

        [Parameter(Mandatory=$true, Position=4)]
        [string]$Domain,

        [Parameter(Mandatory=$true, Position=5)]
        [string]$EmailDomain,

        [switch]$Expire,

        [string]$ExportCsv
    )

    function New-RandomPassword {
        # Generate a random password following default format
        $script:PassNumber = Get-Random -InputObject 123, 234, 345, 456, 567, 678, 789, 890
        $script:PassSymbol = Get-Random -InputObject '!', '?'
        $script:Password = $FirstName.Substring(0,1).ToLower() + $LastName.Substring(0,1).ToLower() + $PassNumber + $FirstName.Substring(0,1).ToLower() + $LastName.Substring(0,1).ToLower() + $PassSymbol
    }

    # Import modules
    Import-Module MSOnline -Force
    Import-Module ActiveDirectory -Force
    
    # Set display name and email name using first and last name variables
    $DisplayName = "$FirstName $LastName"
    $EmailName = $FirstName.substring(0,1).ToLower() + $LastName.ToLower() + $EmailDomain

    # Check if email already exists. If so, create email in a different format
    $DuplicateCheck = Get-MsolUser -UserPrincipalName $EmailName -ErrorAction SilentlyContinue

    if ($null -ne $DuplicateCheck) {
        Write-Warning "$EmailName already exists."
        $DuplicateCheck | Out-Host
        $EmailName = $FirstName.ToLower() + $LastName.Substring(0,1).ToLower() + $EmailDomain
        Write-Warning "Email name set to $EmailName."
    }

    # Generate  password
    New-RandomPassword

    # Params to create new Office365 account
    $OfficeParams = @{
        FirstName            = $FirstName
        LastName             = $LastName
        DisplayName          = $DisplayName
        UserPrincipalName    = $EmailName
        Password             = $Password
        UsageLocation        = 'US'
        ForceChangePassword  = $false
        PasswordNeverExpires = $true
        LicenseAssignment    = 'Inc:ENTERPRISEPACK'
    }

    # Copy copyuser groups/distribution lists to variables $Groups and $DistLists
    $CopyUser = $ADUserCopy + $EmailDomain
    $UserInfo = Get-User $CopyUser | Select-Object -ExpandProperty DistinguishedName
    $Groups = @(Get-Recipient -Filter "Members -eq '$UserInfo'" -RecipientTypeDetails GroupMailBox)
    $DistLists = @(Get-Recipient -Filter "Members -eq '$UserInfo'" -RecipientTypeDetails MailUniversalDistributionGroup, MailUniversalSecurityGroup)

    # Assign DistinguishedName as the distinguished name of the AD user you want to copy
    $DistinguishedName = (Get-ADUser $ADUserCopy).DistinguishedName

    # Assign Path as the proper format for AD path
    $PathArray = $DistinguishedName.Split(',')
    $Path = $PathArray[1..$PathArray.Count] -join ","

    # Write random password into Secure String to make it compatible with AD standards
    $SPassword = ConvertTo-SecureString $Password -AsPlainText -Force

    # Set SamAccountName
    $SamName = $FirstName.Substring(0,1).ToLower() + $LastName.ToLower()

    $ADParams = @{
        Credential           = $ADCredential
        Server               = $Domain
        Name                 = "$FirstName $LastName"
        DisplayName          = "$FirstName $LastName"
        SamAccountName       = ($FirstName.Substring(0,1).ToLower() + $LastName.ToLower())
        Path                 = $Path 
        GivenName            = $FirstName
        Surname              = $LastName
        Description          = $Description
        UserPrincipalName    = ($FirstName.Substring(0,1).ToLower() + $LastName.ToLower() + $EmailDomain)
        EmailAddress         = ($FirstName.Substring(0,1).ToLower() + $LastName.ToLower() + $EmailDomain)
        AccountPassword      = $SPassword
        PasswordNeverExpires = $true
        Instance             = $DistinguishedName
    }

    $ADExpireParams = @{
        Credential           = $ADCredential
        Server               = $Domain
        Name                 = "$FirstName $LastName"
        DisplayName          = "$FirstName $LastName"
        SamAccountName       = ($FirstName.Substring(0,1).ToLower() + $LastName.ToLower())
        Path                 = $Path 
        GivenName            = $FirstName
        Surname              = $LastName
        Description          = $Description
        UserPrincipalName    = ($FirstName.Substring(0,1).ToLower() + $LastName.ToLower() + '@.net')
        EmailAddress         = ($FirstName.Substring(0,1).ToLower() + $LastName.ToLower() + '@.com')
        AccountPassword      = $SPassword
        PasswordNeverExpires = $false
        Instance             = $DistinguishedName
    }

    # Write new user params to console
    Write-Host 'New Office user details' -ForegroundColor Green
    $OfficeParams.GetEnumerator() | Sort-Object -Property Name

    if ($Expire) {
        Write-Host 'New AD user details' -ForegroundColor Green
        $ADExpireParams.GetEnumerator() | Sort-Object -Property Name
    }

    else {
        Write-Host 'New AD user details' -ForegroundColor Green
        $ADParams.GetEnumerator() | Sort-Object -Property Name
    }

    # Write copy user details to console
    Write-Host 'Copied user details' -ForegroundColor Green
    Get-MsolUser -UserPrincipalName $CopyUser | Out-Host
    Write-Host 'Groups'
    $Groups.Alias
    $DistLists.PrimarySmtpAddress
    
    # Confirm check if user info is correct
    $Confirm = Read-Host 'Create  user with this information? (y/n)'

    if ($Confirm -eq 'y') {
        New-MsolUser @OfficeParams | Select-Object DisplayName, UserPrincipalName, Password | Out-Host
    }

    # Exit script if confirm not equal to Y
    else {
        Write-Warning 'Exiting Add-User script...'
        return
    }
    
    while (!(Get-MsolUser -UserPrincipalName $EmailName -ErrorAction Ignore)) {
        Write-Host "Waiting for email account creation to complete..."
        Start-Sleep -Seconds 5
    }    
     
    Write-Host 'Now starting AD account creation...'
               
    # Password expire true/false
    if ($Expire) {
        New-ADUser @ADExpireParams
        Enable-ADAccount -Credential $ADCredential -Server $Domain -Identity $SamName

        # Copy MemberOf groups from the specified $ad_user
        Get-ADUser -Identity $ADUserCopy -Properties MemberOf | Select-Object MemberOf -ExpandProperty MemberOf | Add-ADGroupMember -Members $SamName
    }
 
    else {
        New-ADUser @ADParams
        Enable-ADAccount -Credential $ADCredential -Server $Domain -Identity $SamName

        # Copy MemberOf groups from the specified $ad_user
        Get-ADUser -Identity $ADUserCopy -Properties MemberOf | Select-Object MemberOf -ExpandProperty MemberOf | Add-ADGroupMember -Members $SamName
    }

	Start-Sleep -Seconds 45
	 # Loop through $Groups and $DistLists to add the copied groups/distribution lists
		for ($i = 0; $i -lt ($Groups.Count); $i++) {
				Add-UnifiedGroupLinks -Identity $Groups[$i].Alias -LinkType Member -Links $EmailName
		}

		for ($i = 0; $i -lt ($DistLists.Count); $i++) {
				Add-DistributionGroupMember -Identity $DistLists[$i].PrimarySmtpAddress -Member $EmailName
		}
	
                  
    Write-Host 'New user email and AD have been created!'
    pause
}
