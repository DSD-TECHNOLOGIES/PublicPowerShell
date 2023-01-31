# Returns a list of apps found within  theShell.Application namespace and determines which are pinned to the Microsoft Windows taskbar

Function Get-PinnedHEX{
    # Taskband Key
    $Taskband = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    
    # Taskband Binary to HEX
    $TaskbandHEX = (Get-ItemProperty -Path $Taskband -Name FavoritesResolve).FavoritesResolve | Format-Hex

    # empty HEXString
    $HEXString = ""    

    # format the relevant HEX outputs into a single string
    foreach($Line in $TaskbandHEX)
    {
        $HEXString += $line.tostring().Replace($line.tostring().Substring(0,60),"")
    }

    Return $HEXString
}

Function Get-PinnedApps{
param($HEX)

    # Results Obj
    $ResultsCollection = @()

    # Get Apps using the COM object that also allows the unpin method
    $Apps = (New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items()
    
    # For each App found
    Foreach($App in $Apps)
    {
        # empty the TempAppName
        [string]$TempAppName = ""

        # Split the App's Path to a character array and for each item in that array 
        # populate the TempAppName and also add a fullstop to match the Taskband format
        $App.Path.ToCharArray() | foreach{$TempAppName = $TempAppName + $_ + "." }

        # Create a temporary object to populate results
        $TempObj = New-Object -TypeName PSCustomObject

        # Add the AppName name to the temp object
        $TempObj | Add-Member -MemberType NoteProperty -Name AppName -Value $App.Name
        
        # If the formatted app name, check if the TempAppName is found in the HEXString and add result to temp object
        $TempObj | Add-Member -MemberType NoteProperty -Name Pinned -Value ($HEXString -like "*$TempAppName*")

        # store current apps detail to an object of results
        $ResultsCollection += $TempObj
    }

    # return details form all found apps
    Return $ResultsCollection
}

Get-PinnedApps -HEX Get-PinnedHEX
