function Export-JabberCList {

<#
.SYNOPSIS
  Name: Export-JabberCList.ps1
  The purpose of this script is to create a CSV to import contacts into Jabber.

.DESCRIPTION
  Create an xlsx excel spreadsheet as a master database containing columns: User ID, User Domain, Nickname (can be empty), Group Name
  Whenever you want to update the master list, save first as xlsx (master database xlsx)vol, and save another copy as csv (master database csv). Do not make changes and save on the CSV directly

  Point the $inputCsv at the master database csv file. Point $outputCsv to the file you want to upload to CUCM IM & Presence Administration

  In CUCM IMP, go to Bulk Administration -> Upload/Download Files and click Add New
  Choose your csv database, target contact lists, select import users' contacts - custom file

  After uploading, go to Bulk Administration -> Contact List -> Update Contact List
  Choose your newly uploaded contact list from the dropdown list, select Run Immediately, and click submit

.NOTES
  Version 1.0.0

  Release Date: 09-18-2018

  Author: Jeff Li

.EXAMPLE
  Example 1: Create and export Jabber contact list
  PS C:\> Export-JabberCList -InputPath C:\Users\Jeff\Documents\Jabbercontacts.csv -OutputPath C:\Users\Jeff\Documents
#>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$InputPath,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    $CurrentDate = (Get-Date).ToString('MM-dd-yyyy')
    $FileName = "Jabber Contact List $CurrentDate.csv"
    
    # Use these variables if you want to set a static input and output path, remove mandatory params
    #$inputCsv = $FileBrowser.FileName #'C:\Users\jli\Documents\PowerShell\Jabber users.csv' # Input path
    #$outputCsv = "C:\Users\jli\Documents\PowerShell\Jabber contacts $CurrentDate.csv" # Output path

    # Import csv to $jabberCsv
    $JabberCsv = Import-Csv -Path $InputPath
    # Create empty array to hold Jabber contact list
    $JabberArray = @() 

    # Loop through csv data twice to write all contacts to each user
    foreach ($User in $JabberCsv) {
        foreach ($UserProfile in $JabberCsv) {
            if ($UserProfile.'User ID' -eq $User.'User ID') { 
                continue  # Do not write contact if contact is self
            }
            else {
                $Jabber = @{}
                $Jabber.'User ID'          = $User.'User ID'
                $Jabber.'User Domain'      = $UserProfile.'User Domain'
                $Jabber.'Contact ID'       = $UserProfile.'User ID'
                $Jabber.'Contact Domain'   = $UserProfile.'User Domain'
                $Jabber.'Nickname'         = $UserProfile.'Nickname'
                $Jabber.'Group Name'       = $UserProfile.'Group Name'

                # Create temp object $tempObject from above csv data
                $TempObject = New-Object PSObject -Property $Jabber

                # Append temp object value to array $jabberArray
                $JabberArray += $TempObject 
            }
        }
    }
    $JabberArray | Select-Object 'User ID', 'User Domain', 'Contact ID', 'Contact Domain', 'Nickname', 'Group Name' | Export-Csv  -Path "$OutputPath$FileName" -NoTypeInformation -Force
}
