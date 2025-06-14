# add Windows Forms assembly for cursor position
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# add shell to send keyboard combinations
$Global:wsh = New-Object -ComObject WScript.Shell

# variables to track mouse position and timing
$Global:lastMousePos = [System.Windows.Forms.Cursor]::Position
$Global:timeToTrigger = 3      # seconds
$Global:secondsSinceMove = 0

# Define the action script block BEFORE using it
$actionScript = {
    Write-Host "`rRunning..." -NoNewline
    
    # get current mouse position
    $currentMousePos = [System.Windows.Forms.Cursor]::Position
    
    # check if mouse has moved
    if ($currentMousePos.X -ne $Global:lastMousePos.X -or $currentMousePos.Y -ne $Global:lastMousePos.Y) {
        # mouse moved - reset timer, make current position - last
        $Global:secondsSinceMove = 0
        $Global:lastMousePos = $currentMousePos
    } else {    
        # calculate time since last movement
        $Global:secondsSinceMove++
    }
    
    # debug
    Write-Host "`rIdle timer: $($Global:timeToTrigger) seconds. Time since last activity: $($Global:secondsSinceMove) seconds." -NoNewline
    
    # time is up - make action
    if ($Global:secondsSinceMove -ge $Global:timeToTrigger) {
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
        # $Global:wsh.SendKeys('+{F15}')
        $Global:secondsSinceMove = 0
    }
}

# timer setup
$timer = New-Object System.Timers.Timer
$timer.Interval = 1000 # ms
$timer.AutoReset = $true

# create the event subscription (using the properly defined script block)
Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action $actionScript

# Start the timer
$timer.Enabled = $true
Write-Host "Mouse idle monitor started. Press Ctrl+C to stop." -ForegroundColor Green


# Keep the script running
try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
}
finally {
    # Cleanup when script ends
    $timer.Stop()
    $timer.Dispose()
    Get-EventSubscriber | Unregister-Event
    Write-Host "`nTimer stopped and cleaned up." -ForegroundColor Yellow
}






<#

# Start the timer
$timer.Start()
Write-Host "Timer started. Press Ctrl+C to stop or wait 10 seconds for auto-stop..." -ForegroundColor Yellow

# Let it run for 10 seconds as demonstration
Start-Sleep -Seconds 10

# Stop and dispose
$timer.Stop()
$timer.Dispose()
Write-Host "Timer stopped." -ForegroundColor Red

# ============================================
# Method 2: Simple loop approach (Alternative)
Write-Host "`nAlternative method using loop:" -ForegroundColor Cyan

$counter = 0
$maxTicks = 5

while ($counter -lt $maxTicks) {
    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "Loop tick $($counter + 1): $timestamp" -ForegroundColor Blue
    
    Start-Sleep -Seconds 1
    $counter++
}

# ============================================
# Method 3: Advanced timer with event cleanup
Write-Host "`nAdvanced timer with proper cleanup:" -ForegroundColor Magenta

$timer2 = New-Object System.Timers.Timer
$timer2.Interval = 1000
$timer2.AutoReset = $true

$eventJob = Register-ObjectEvent -InputObject $timer2 -EventName Elapsed -Action {
    $Global:TickCount++
    Write-Host "Advanced timer tick #$Global:TickCount at $(Get-Date -Format 'HH:mm:ss.fff')" -ForegroundColor Magenta
    
    # Stop after 5 ticks
    if ($Global:TickCount -ge 5) {
        $Event.MessageData.Stop()
    }
} -MessageData $timer2

# Initialize counter
$Global:TickCount = 0

# Start timer
$timer2.Start()

# Wait for completion
do {
    Start-Sleep -Milliseconds 100
} while ($timer2.Enabled)

# Cleanup
Unregister-Event -SourceIdentifier $eventJob.Name
$timer2.Dispose()
Write-Host "Advanced timer completed and cleaned up." -ForegroundColor Green

#>