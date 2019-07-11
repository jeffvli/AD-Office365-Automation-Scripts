# Add Veeam cmdlets
Add-PSSnapin -Name VeeamPSSnapin
# $Date = Get-Date # Current date
# $ExpiredTape = @() # Array for expired tapes

# Declare the pool to add expired tapes to (Free)
$FreePool = Get-VBRTapeMediaPool -Name 'Free'

# Declare the pools to find expired tapes from
# $MonthlyPool = Get-VBRTapeMediaPool -Name 'Monthly Pool'
$WeeklyPool = Get-VBRTapeMediaPool -Name 'Weekly Pool'

# Declare and find expired tapes from $MonthlyPool and $WeeklyPool
# $ExpiredMonthlyTape = Get-VBRTapeMedium -MediaPool $MonthlyPool | Where-Object {$_.isExpired -eq $true}
$ExpiredWeeklyTape = Get-VBRTapeMedium -MediaPool $WeeklyPool | Where-Object {$_.isExpired -eq $true}


<#
foreach ($Tape in $ExpiredWeeklyTape) {
    # If the expired tape is younger than 4 months, do nothing
    if ($Tape.ExpirationDate -gt $Date.AddDays(-120)) {
        continue
    }

    # If the expired tape is older than 4 months, add to $ExpiredTapes array
    else {
        $ExpiredTape += $Tape
    }
}
#>

# Move all expired tapes from to the declared free pool
Move-VBRTapeMedium -Medium $ExpiredWeeklyTape -MediaPool $FreePool
