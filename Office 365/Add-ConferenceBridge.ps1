# First create new Shoretel conference bridge in Shoretel portal
# Copy the participant and host code for reference

function Add-ConferenceRoom {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$true, Position=0)]
    [String]$FirstName,

    [Parameter(Mandatory=$true, Position=1)]
    [String]$LastName,

    [Parameter(Mandatory=$true, Position=2)]
    [String]$ParticipantCode,

    [String]$Domain
    )

    $Name = "$FirstName $LastName $ParticipantCode#"
    $Phone = "$ParticipantCode#"
    $UPN = "$FirstName.$LastName$Domain"

    $Params = @{
        Name = $Name
        DisplayName = $Name
        Phone = $Phone
        Room = $true
    }

    New-Mailbox @Params # Create the conference bridge as a room in Office 365
    Set-Mailbox "$Name" -EmailAddress "SMTP:$UPN" # Set the conference room email to Firstname.LastName@Domain
}