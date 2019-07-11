function Unlock-ADUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Server,
        [Parameter(Mandatory=$true, Position=1)]
        [string]$LogPath
    )
    $TimeStamp = Get-Date -Format "MM/dd/yyyy - HH:mm:ss"
    $Credentials = Get-Credential -Message 'Enter AD admin credentials (User@Domain.com or Domain\Username)'
    try {
        $LockedUserSearch = Search-ADAccount `
        -LockedOut `
        -Credential $Credentials `
        -Server $Server `
        | Select-Object SamAccountName, LockedOut, DistinguishedName -ErrorAction Stop
    }

    catch {
        if ($null -eq $Credentials -or $Credentials -eq '') {
            break
        }
        else {
            Write-Warning 'Please re-enter valid credentials.' | Tee-Object -FilePath $LogPath -Append
            Unlock-ADUser
        }
    }

    Write-Output "($TimeStamp) AD unlock utility has been started" | Tee-Object -FilePath $LogPath -Append

    while ($true) {
        $TimeStamp = Get-Date -Format "MM/dd/yyyy - HH:mm:ss"
        if ($null -eq $LockedUserSearch -or $LockedUserSearch -eq "") {
            Write-Output "($TimeStamp) No users are locked." | Tee-Object -FilePath $LogPath -Append
            pause
        }
 
        else {
            $LockedUserSelection = @($LockedUserSearch.SamAccountName)
            foreach ($User in $LockedUserSelection) {
                Unlock-ADAccount -Identity $User -Credential $Credentials -Server $ADServer
                Write-Output "($TimeStamp) $User has been unlocked." | Tee-Object -FilePath $LogPath -Append
                pause
            }
        }
    }
}
