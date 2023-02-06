Function Get-PinnedApps{

    # Results Obj
    $ResultsCollection = @()

    # set Found state is false by default
    $found = $False

    # Get Apps using the Shell.Application COM object 
    $Apps = (New-Object -Com Shell.Application).NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}').Items()
    
    # For each App found
    Foreach($App in $Apps)
    {
        # Create a temporary object to populate results
        $TempObj = New-Object -TypeName PSCustomObject

        # Add the AppName name to the temp object
        $TempObj | Add-Member -MemberType NoteProperty -Name App -Value $App.Name
       
        # Check the verbs available to the current app
        Foreach($verb in $app.Verbs())
        {
            # if the Unpin verb is found then the item is pinned
            if($verb.Name -eq 'Unpin from tas&kbar')
            {
                # set Found to True
                $found = $true
            }
        }
        
        # Add the Found result to the temp object
        $TempObj | Add-Member -MemberType NoteProperty -Name Pinned -Value $found

        # store current app results to a collection of results
        $ResultsCollection += $TempObj

         # Return back to not found
        $found = $false
    }

    # return app collection results
    Return $ResultsCollection
}

Get-PinnedApps
