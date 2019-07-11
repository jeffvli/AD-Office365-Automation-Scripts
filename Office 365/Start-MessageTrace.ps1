function Start-MessageTrace {
    function time_selection {
        $script:time_select = Read-Host 'How many days in the past do you want to search? (Use numerals 1-10)'
        $script:time_check = $script:time_select -in 1..10
        if ($script:time_check -eq $false) {
            Write-Warning 'Invalid range...'
            time_selection
        }
        else {
            $script:date_end = (Get-Date).AddDays(1).ToString('MM/dd/yyyy') 
            $script:date_start = (Get-Date).AddDays(-$time_select).ToString('MM/dd/yyyy')
        }
    }
    do {
        Clear-Host
        # Calculate a time frame for which to perform the message trace
        Write-Host 'You have chosen to trace Office 365 emails' -ForegroundColor Green
        Write-Host '1: Search by sender'
        Write-Host '2: Search by recipient'
        Write-Host '3: Search by both sender and recipient'
        Write-Host '------------------------'
        Write-Host 'Q: Return to previous menu'
        $selection_temp = Read-Host 'Make your selection'

        switch ($selection_temp) {
            # Perform message trace by the sender
            '1' {
                Clear-Host
                Write-Host 'You have chosen to search by sender' -ForegroundColor Green
                $sender = Read-Host 'Enter the full email of the sender'
                if ($sender -eq '') {
                    return
                }

                else {
                    time_selection
                    Get-MessageTrace -SenderAddress $sender -StartDate $date_start -EndDate $date_end | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, Size, MessageID, MessageTraceID | Out-GridView
                }
            }
            # Perform message trace by the recipient
            '2' {
                Clear-Host
                Write-Host 'You have chosen to search by recipient' -ForegroundColor Green
                $recipient = Read-Host 'Enter the full email of the recipient'
                if ($recipient -eq '') {
                    return
                }

                else {
                    time_selection
                    Get-MessageTrace -RecipientAddress $recipient -StartDate $date_start -EndDate $date_end | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, MessageID, MessageTraceID | Out-GridView
                }
            }
            # Perform message trace by both sender and recipient
            '3' {
                Clear-Host
                Write-Host 'You have chosen to search by both sender and recipient' -ForegroundColor Green
                $sender = Read-Host 'Enter the full email of the sender'
                if ($sender -eq '') {
                    return
                }

                else {
                    $recipient = Read-Host 'Enter the full email of the recipient'
                    if ($recipient -eq '') {
                        return
                    }

                    else {
                        time_selection
                        Get-MessageTrace -SenderAddress $sender -RecipientAddress $recipient -StartDate $date_start -EndDate $date_end | Select-Object Received, SenderAddress, RecipientAddress, Subject, Status, ToIP, FromIP, MessageID, MessageTraceID | Out-GridView
                    }
                }
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
