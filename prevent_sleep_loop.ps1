
# add Windows Forms assembly for cursor position
Add-Type -AssemblyName System.Windows.Forms

# add shell to  send keyboard combinations
$wsh = New-Object -ComObject WScript.Shell

# variables to track mouse position and timing
$lastMousePos = [System.Windows.Forms.Cursor]::Position
$lastMoveTime = Get-Date
$checkInterval = 1000  # check every 1000 ms (Start-Sleep approach has bad precision)
$timeToTrigger = 3     # seconds

Write-Host "`rRunning..." -NoNewline

try {
    while($true) {
        # get current mouse position
        $currentMousePos = [System.Windows.Forms.Cursor]::Position
        $currentTime = Get-Date

        # check if mouse has moved
        if ($currentMousePos.X -ne $lastMousePos.X -or $currentMousePos.Y -ne $lastMousePos.Y) {
            # mouse moved - reset timer
            $lastMoveTime = $currentTime
            $lastMousePos = $currentMousePos
        }
        
        # calculate time since last movement
        $timeSinceMove = $currentTime - $lastMoveTime
        $secondsSinceMove = [math]::Floor($timeSinceMove.TotalSeconds)

        # debug
        Write-Host "`rIdle timer: $timeToTrigger seconds. Time since last activity: $secondsSinceMove seconds." -NoNewline

        # time is up - make action
        if ($secondsSinceMove -ge $timeToTrigger) {
            # move cursor square to prevent cursor stuck in screen corners
            $Pos = [System.Windows.Forms.Cursor]::Position
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($Pos.X + 1), ($Pos.Y - 1))
            Start-Sleep -Milliseconds 50
            
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($Pos.X, ($Pos.Y + 2))
            Start-Sleep -Milliseconds 50
            
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($Pos.X - 2), $Pos.Y)
            Start-Sleep -Milliseconds 50
            
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($Pos.X, ($Pos.Y - 2))
            Start-Sleep -Milliseconds 50
            
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($Pos.X + 2), $Pos.Y)
            Start-Sleep -Milliseconds 50

            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(($Pos.X - 1), ($Pos.Y + 1))
            Start-Sleep -Milliseconds 50

            # Shift + F15 -- the least intrusive key combination I can think of
            # $wsh.SendKeys('+{F15}')

            $lastMoveTime = $currentTime
        }

        # Wait before next check
        Start-Sleep -Milliseconds $checkInterval

    }
}
catch {
    Write-Host "`nScript interrupted - exception has occurred." -ForegroundColor Red
}

