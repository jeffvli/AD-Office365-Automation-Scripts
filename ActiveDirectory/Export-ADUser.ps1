function Export-ADUser {

<#
.SYNOPSIS
  Name: Export-ADUser.ps1
  The purpose of this script is to search and/or export AD users.

.DESCRIPTION
 Export-ADUser.ps1 provides a log-in function Connect-AD to provide your domain admin credentials.
 Following login, you can use function Export-ADUser with a mandatory parameter Name, to provide
 either * (all AD accounts) or a string which searches by GivenName, SurName, and SamAccountName. 

 The search output is given in an Out-Gridview, which allows options to filter by the selected object
 parameters.

 In addition, there is an optional parameter, ExportCsv, which will allow you to specify a path
 in which to export the selected AD user output to a .csv file.

.NOTES
  Version 1.0.0

  Release Date: 09-19-2018

  Author: Jeff Li

.EXAMPLE
  Example 1: Connect to AD and search all AD users without exporting to CSV. Select users from
  gridview and output to console.
  PS C:\> Connect-AD
  PS C:\> Export-ADUser -Name *

  Example 2: Export-ADUser with name Jeff to C:\test.csv
  PS C:\> Export-ADUser -Name Jeff -ExportCsv C:\test.csv
#>

    [CmdletBinding()]
    Param(
            [Parameter(Mandatory=$true)]
            [string]$Name,
            [Parameter(Mandatory=$false)]
            [string]$ExportCsv,
            [string]$DC
    )
	
    $ADName = $Name
    # Set the number of days within expiration
    $DaysWithinExpiration = 30
 
    # Set the days where the password is already expired and needs to change
    $MaxPwdAge   = (Get-ADDefaultDomainPasswordPolicy -Credential $ADCredential -Server $DC).MaxPasswordAge.Days
    $ExpiredDate = (Get-Date).addDays(-$MaxPwdAge)

    # Select params for AD search
    $ADSelect = @(
        'GivenName',
        'SurName',
        'SamAccountName',
        'EmailAddress',
        'LockedOut',
        'Enabled',
        'Description', 
        'DistinguishedName',
        'PasswordNeverExpires',
        'PasswordLastSet',
        @{name = "DaysUntilExpired"; Expression = {$_.PasswordLastSet - $ExpiredDate | Select-Object -ExpandProperty Days}}
    )

    if ($ADName -eq '*') {
        $ADUser = Get-ADUser -Server $DC -Credential $ADCredential -Filter * -Properties * | Select-Object $ADSelect | Out-GridView -Title 'AD Users' -PassThru
    } 

    else {
        $ADUser = Get-ADUser -Server $DC -Credential $ADCredential -Filter "Name -like '*$ADName*'" -Properties * | Select-Object $ADSelect | Out-GridView -Title 'AD Users' -PassThru
    }

    # Write selection to console
    $ADUser | Format-Table

    if ($ExportCsv -eq $null -or $ExportCsv -eq '') {
        return
    }

    else {
        $ADUser | Export-Csv $ExportCsv -Force -NoTypeInformation
        Write-Host "Wrote csv to $ExportCsv"
    }
}
