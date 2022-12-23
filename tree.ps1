<#
.SYNOPSIS
   This is a script that runs a Simple Holiday Tree
.DESCRIPTION
   This is a re-write and optimization of my first PowerShell script that
   creates a Holiday tree in the console complete with expensive
   blinking lights.  It serves no purpose other than to be
   pretty and as a demo of PowerShell
.EXAMPLE
   #Display Tree with default output for message
   .\tree.ps1 
.EXAMPLE
   #Display Tree with default output for message
   .\tree.ps1 -message 'Something fun to display'
.NOTES
   This is just a script to demo using Write-Host to extremes
   in PowerShell.  It serves no other purpose than to brighten
   up the console with blinking lights
.FUNCTIONALITY
   Completely frivolous and pointless fun
#>
param($Message = 'Happy Holidays PowerShell')
# PowershellTree-V-whatever.PS1
# 
# more information
# First verify we are NOT running in the ISE, this will NOT work in the ISE
If ($Host.Name -notmatch 'ISE') {
    #Clear the Screen

    clear-host

    #Move it all down the line

    write-host

#region Beginning of variables
    # character for 'LIGHTS'

    $starchar = @('0')

    # Number of rows Deep for the Tree

    $Rows = ([console]::WindowHeight) - 7
    $WidthofConsole = $Host.UI.RawUI.WindowSize
    $WidthOfConsole.Width = (($rows + 1) * 2) - 1 + 4
    $WidthOfBuffer = $Host.UI.RawUI.BufferSize
    $WidthofBuffer.Width = $WidthofConsole.Width

    $Host.UI.RawUI.WindowSize = $WidthofConsole
    $Host.Ui.RawUI.BufferSize = $WidthofBuffer

    # Change the title of the Window
    $Host.ui.RawUI.WindowTitle = $Message

    # These variables are for the Bouncing Marquee at the bottom
    # Column, number of Columns to move (relative to size of tree)
    # and Direction it will move

    $BottomRow = $Rows + 4
    $BottomColumn = 0
    $Direction = 1
    $Min = 0
    $Max = $Rows

    # Get all the console colors available except the background
    # so you don't get "Burned out lights"
    $colors = [enum]::GetValues([System.ConsoleColor]) | Where-object { $_ -notmatch $Host.ui.RawUI.BackgroundColor }

    # Get where the Cursor was

    $oldpos = $host.ui.RawUI.CursorPosition
#endregion variables
    # BsponPosh’s ORIGINAL Tree building Algorithm 🙂
    # None of this would be possible if it weren’t for him

    # Here we define the primary loop
    Foreach ($r in ($rows..1)) {
        write-host $(' ' * $r) -NoNewline
        <#
        Slight change to improve performance.  
        The old process was a for loop.   Used bsonposh approach with Math
        to make it all faster
        #>
        write-Host $('*' * (($rows + 1 - $r) * 2 - 1)) -ForegroundColor Darkgreen  -nonewline
        write-host
    }           

    # trunk
    # A slight change, an extra row on the stump of the tree
    # and Red (Trying to make it look like a brown trunk

    write-host $('{0}***' -f (' ' * ($Rows - 1) ))  -ForegroundColor DarkRed
    write-host $('{0}***' -f (' ' * ($Rows - 1) ))  -ForegroundColor DarkRed
    write-host $('{0}***' -f (' ' * ($Rows - 1) ))  -ForegroundColor DarkRed

    $host.ui.RawUI.CursorPosition = $oldpos

    # New Addins by Sean “The Energized Tech” Kearney

    # Compute the possible number of stars in tree (Number of Rows Squared)

    $numberstars = [math]::pow($Rows, 2)

    # Number of lights to give to tree.  %25 percent of the number of green stars.  You pick

    [int]$numberlights = $numberstars * .10

    # Initialize an array to remember all the “Star Locations”

    [System.Array]$Starlocation = @()

    for ($i = 0; $i -lt $numberlights; $i++) {
        $Starlocation += @($host.ui.Rawui.CursorPosition)
    }

    # Probably redundant, but just in case, remember where the  heck I am!

    $oldpos = $host.ui.RawUI.CursorPosition

    # New change.  Create an Array of positions to place lights on and off

    foreach ($light in ($numberlights..1)) {
        # Pick a Random Row

        $row = (get-random -min 1 -max (($Rows) + 1))

        # Pick a Random Column – Note The Column Position is
        # Relative to the Row vs Number of Rows

        $column = ($Rows - $row) + (get-random -min 1 -max ($row * 2))

        #Grab the current position and store that away in a $Temp Variable
        $temppos = $host.ui.rawui.CursorPosition

        # Now Build new location of X,Y into $HOST
        $temppos.x = $column
        $temppos.y = $row

        # Store this away for later
        $Starlocation[(($light) - 1)] = $temppos

        # Now update that “STAR” with a Colour
    }
    # Repeat this OVER and OVER and OVER and OVER

    while ($true) {

        # Now we just pull all those stars up and blank em back
        # on or off randomly 7 at a time

        for ($light = 1; $light -lt 7; $light++) {
            # Set cursor to random location within Predefined 'Star Location Array'

            $host.ui.RawUI.CursorPosition = ($Starlocation | get-random)
            # Pick a random number between 1 and 1000
            # if 500 or higher, turn it to a light
            # Else turn it off

            write-Host ($starchar | Get-random) -ForegroundColor  ($colors | get-random) -nonewline
        }

        # Remember where we are

        $temppos = $oldpos

        # Set a position for the row and Column

        $oldpos.X = $BottomColumn
        $oldpos.Y = $BottomRow

        # update the console

        $host.ui.Rawui.CursorPosition = $oldpos

        # Bump up the column position based upon direction

        $BottomColumn = $BottomColumn + $Direction

        # Boolean trick to flip the Direction from 1 to -1 without using if/then/else
        $direction = $Direction * ( - [int]($BottomColumn -eq $min)) + $Direction * ( - [int]($BottomColumn -eq $max)) + $Direction * [int]($BottomColumn -gt $min -and $BottomColumn -lt $max)

        # Print greeting.  Space must be before and after to avoid
        # Trails.  Output in Random Colour

        write-host " $Message " -ForegroundColor  ($colors | get-random)
        $WindowTitle = "$(' ' * ($Max-$BottomColumn))$Message"
        $Host.UI.Rawui.WindowTitle = $WindowTitle
        start-sleep -Milliseconds 50
        # End of the loop, keep doin’ in and go “loopy!”

    }
}
else {
    
    Write-Host 'Sorry, this script cannot work in the ISE due to heavy use of $Host.UI.Rawui'
}