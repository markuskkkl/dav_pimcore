# Get-PimcoreGroupEvents.ps1
# v1 2025-04-15
#
#
# Extract events (Gruppentermine) from Pimcore backend, cause frontend cannot display required data in one list
# Script needs cookie and CSRF token from a valid browser session
#
# Author: Markus Bossert
#
# Please follow the steps below to retrieve valid auth data:
# 1) Open you webbrowser and navigate to https://dav360redaktion.alpenverein.de/admin/
# 2) Login to your account (if necessary)
# 3) Open any Dokument or Datenobjekt
# 4) Enter Developer-Mode in your browser (F12 key) and switch to network tab
# 5) Close the document/site in Pimcore by hitting X button
# 6) Inspect the 'unlock-element' event in the network tab
# 7) From Headers copy the values of 'Cookie' and 'x-pimcore-csrf-token'
#      Examples:
#           PHPSESSID=63tnaodiheu2025phjd4scuddq   <-- this is your [AuthCookie]
#           c373a1d22e33cc72b12bbe44e2a8aa3548e32d31   <-- this is your [PimcoreCsrfToken]
#
# Execute script with following syntax:
# .\Get-PimcoreGroupEvents.ps1 -AuthCookie "[AuthCookie]" -PimcoreCsrfToken "[PimcoreCsrfToken]"

param (
    [Parameter(Mandatory=$true)]
    [string]$AuthCookie,
    [Parameter(Mandatory=$true)]
    [string]$PimcoreCsrfToken
)


######
# Configuration

# $AuthCookie = "PHPSESSID=..."
# $PimcoreCsrfToken = "bd33a9...."
$GruppenFolderId = 70514 # ID of folder "Gruppen" (Datenobjekt)


#####
# Required modules
Import-Module PSWriteHTML -Force -ErrorAction Stop


#####
# Functions

# Function to remove HTML tags from code (very simple solution)
Function StripHTML () {
    param (
        [string]$Text
    )

    # replace HTML special chars and remove HTML tags <>
    If ($Text) {
        $Text = [System.Web.HttpUtility]::HtmlDecode($Text)
        $Text = $Text -replace '<p>',''
        $Text = $Text -replace '</p>',"`r`n"
        $Text = $Text -replace '<[^>]+>',''
        return $Text
    }
}

# Function to test connectivity to Pimcore backend
Function Test-PimcoreBackend () {
    param (
        [Parameter(Mandatory=$true)]
        [string]$AuthCookie,
        [Parameter(Mandatory=$true)]
        [string]$PimcoreCsrfToken
    )

    Write-Host "Testing access to Pimcore backend .."

    # Run test with PHPSESSID and CSRF Token (last one requires PUT operation, cause Pimcore does not validate CSRF for read operations)
    $TestItemString = "id=27322&type=document" # ID of "113 - Sektion Hannover" main website
    $TestUrl = "https://dav360redaktion.alpenverein.de/admin/element/unlock-element"
    $Headers = @{
            Cookie=$AuthCookie;
            "Referer"="https://dav360redaktion.alpenverein.de/admin/";
            "X-Pimcore-Csrf-Token"=$PimcoreCsrfToken;
            "X-Requested-With"="XMLHttpRequest"
    }

    Try {
        $TestQuery = Invoke-WebRequest -Uri $TestUrl -Method PUT -Headers $Headers -Body $TestItemString -ContentType 'application/x-www-form-urlencoded; charset=UTF-8' -ErrorAction SilentlyContinue
    }
    Catch {
        Write-Error "Connection to backend failed"
        Write-Host "Backend response: $($_)"
        return $false
    }

    $TestQuery
    # Lets check if test was successful
    If ($TestQuery.StatusCode -eq 200 -and (($TestQuery.Content | ConvertFrom-Json -ErrorAction SilentlyContinue).success -eq "true")) {
        return $true
    }
    Else {
        return $false
    }
}


# Function to get list of events from Pimcore backend
Function Get-PimcoreGridJson () {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FolderId,
        [Parameter(Mandatory=$true)]
        [string]$ClassId,
        [Parameter(Mandatory=$true)]
        [string]$AuthCookie,
        [Parameter(Mandatory=$true)]
        [string]$PimcoreCsrfToken,
        [bool]$SkipUnlock=$false
    )

    $BaseUrl = "https://dav360redaktion.alpenverein.de/admin/object/grid-proxy?xaction=read"
    $ClassName = 'EventItemTour' # seems to be static for all kind of object types
 
    $QueryUrl = "$($BaseUrl)&classId=$($ClassId)&folderId=$($FolderId)&_dc=$([DateTimeOffset]::Now.ToUnixTimeSeconds())"
    $Headers = @{
        Cookie=$AuthCookie;
        "Referer"="https://dav360redaktion.alpenverein.de/admin/";
        "X-Pimcore-Csrf-Token"=$PimcoreCsrfToken;
        "X-Requested-With"="XMLHttpRequest"
    }
    $Body = "language=en&class=$($ClassName)&fields%5B%5D=id&only_direct_children=false&query=&page=1&start=0&limit=999999"

    $Query = Invoke-WebRequest -Uri $QueryUrl -Method Post -Headers $Headers  -ContentType 'application/x-www-form-urlencoded; charset=UTF-8' -Body $Body

    # Get list of IDs from Json
    If ($Query.StatusCode -eq 200 -and $Query.Headers."Content-Type" -eq 'application/json') {

        Write-Host "Retrieved grid list for class '$($ClassId)' in folder '$($FolderId)'"

        # convert to Json and return IDs
        $Json = $Query.Content | ConvertFrom-Json
        $Result = $Json.data | ? { $_.published -eq $true } | Select id,fullpath,type,subtype,classname,filename,creationDate,modificationDate

        return $Result
    }
    Else {
        Write-Host "Failed to get grid list for class '$($ClassId)' in folder '$($FolderId)'"
        return $false
    }
}


# Function to get single event data from Pimcore backend
Function Get-PimcoreEventJson () {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ItemId,
        [Parameter(Mandatory=$true)]
        [string]$AuthCookie,
        [Parameter(Mandatory=$true)]
        [string]$PimcoreCsrfToken,
        [bool]$SkipUnlock=$false
    )

    $BaseUrl = "https://dav360redaktion.alpenverein.de/admin/object/get"
    $UnlockUrl = "https://dav360redaktion.alpenverein.de/admin/element/unlock-element"

    $QueryUrl = "$($BaseUrl)?_dc=$([DateTimeOffset]::Now.ToUnixTimeSeconds())&id=$($ItemId)"
    $Query = Invoke-WebRequest -Uri $QueryUrl -Headers @{Cookie=$AuthCookie}

    # Get event data from Json
    If ($Query.StatusCode -eq 200 -and $Query.Headers."Content-Type" -eq 'application/json') {
        $Json = $Query.Content | ConvertFrom-Json

        # Test if edit/read for this object is locked and try to remove lock
        If ($Json.editlock -ne $null) {
            If ($SkipUnlock -eq $false) {
                Write-Host "Object is locked for editing, trying to unlock"
                $UnlockQuery = Invoke-WebRequest -Uri $UnlockUrl -Method PUT -Headers @{Cookie=$AuthCookie; "Referer"="https://dav360redaktion.alpenverein.de/admin/"; "X-Pimcore-Csrf-Token"=$PimcoreCsrfToken; "X-Requested-With"="XMLHttpRequest"} -ContentType 'application/x-www-form-urlencoded; charset=UTF-8' -Body "id=$($ItemId)&type=object"
                
                If ($UnlockQuery.StatusCode -eq 200 -and (($UnlockQuery.Content | ConvertFrom-Json -ErrorAction SilentlyContinue).success -eq "true")) {
                    # Retry query to retrieve event data
                    Get-PimCoreEventJson -ItemId $ItemId -AuthCookie $AuthCookie -PimcoreCsrfToken $PimcoreCsrfToken -SkipUnlock:$true
                }
            }
        }
        Else {
            Write-Host "Retrieved event data for ItemId $($ItemId)"
            return $Json.data
        }
    }
    Else {
        Write-Warning "Unable to retrieve event data for ItemId $($ItemId)"
    }
}


# Function to interpret and format event data
Function Get-EventData () {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ItemId
    )

    $EventData = Get-PimCoreEventJson -ItemId $ItemId -AuthCookie $SCRIPT:AuthCookie -PimcoreCsrfToken $SCRIPT:PimcoreCsrfToken

    If ($EventData) {
        $Record = New-Object -TypeName PSCustomObject

        # ID
        $Record | Add-Member -MemberType NoteProperty -Name 'ID' -Value $ItemId

        # Gruppenname
        # $EventData.assignedGroups[0].fullpath.split('/')[-1]
        If ($EventData.assignedGroups) {
            $Record | Add-Member -MemberType NoteProperty -Name 'Gruppe' -Value $EventData.assignedGroups[0].fullpath.split('/')[-1]
        }


        # Titel
        # $EventData.title
        $Record | Add-Member -MemberType NoteProperty -Name 'Titel' -Value $EventData.title


        # Tourenleitung (Referenz zu Objekt)
        # $EventData.leaders[0].fullpath.split('/')[-1] # erste Person
        # ($EventData.leaders.fullpath | % { $_.split('/')[-1] }) -join '; ' # alle Personen
        # (($EventData.leaders | % { "$($_.firstName) $($_.lastName)"}) -join '; ') # alle Personen mit "Vorname Nachname" - funktioniert nicht bei Touren ohne weiteres Query
        If ($EventData.leaders) {
            $Record | Add-Member -MemberType NoteProperty -Name 'Tourenleitung' -Value (($EventData.leaders.fullpath | % { $_.split('/')[-1] }) -join '; ')
        }


        # Veranstaltungsort (Referenz zu Objekt)
        $Record | Add-Member -MemberType NoteProperty -Name 'Veranstaltungsort' -Value $EventData.locations.name
        

        # Treffpunkt
        # $EventData.meetingPoint
        $Record | Add-Member -MemberType NoteProperty -Name 'Treffpunkt' -Value (StripHTML -Text $EventData.meetingPoint)


        # Uhrzeit Start / Ende
        $LocalTimeZone = Get-TimeZone
        If ($EventData.dates.data[0].dateStart) {
            $StartTimeUTC = Get-Date ([datetime]'1/1/1970').AddSeconds($EventData.dates.data[0].dateStart)
            $StartTimeLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($StartTimeUTC, $LocalTimeZone)
        }
        If ($EventData.dates.data[0].dateEnd) {
            $EndTimeUTC = Get-Date ([datetime]'1/1/1970').AddSeconds($EventData.dates.data[0].dateEnd)
            $EndTimeLocal = [System.TimeZoneInfo]::ConvertTimeFromUtc($EndTimeUTC, $LocalTimeZone)
        }

        # Publish time for start/end dates only if value (for start) is not 00:00
        If ((Get-Date $StartTimeLocal -Format 'HH:mm:ss') -ne '00:00:00') {
            $Record | Add-Member -MemberType NoteProperty -Name 'Termin_Start' -Value (Get-Date $StartTimeLocal -Format 'yyyy-MM-dd HH:mm')
            If ($EndTimeLocal) {
                $Record | Add-Member -MemberType NoteProperty -Name 'Termin_Ende' -Value (Get-Date $EndTimeLocal -Format 'yyyy-MM-dd HH:mm')
            }
        }
        Else {
            $Record | Add-Member -MemberType NoteProperty -Name 'Termin_Start' -Value (Get-Date $StartTimeLocal -Format 'yyyy-MM-dd')
            If ($EndTimeLocal) {
                $Record | Add-Member -MemberType NoteProperty -Name 'Termin_Ende' -Value (Get-Date $EndTimeLocal -Format 'yyyy-MM-dd')
            }
        }

        # Beschreibung
        # $EventData.description
        $Record | Add-Member -MemberType NoteProperty -Name 'Beschreibung' -Value (StripHTML -Text $EventData.description)
        $Record | Add-Member -MemberType NoteProperty -Name 'Beschreibung_HTML' -Value $EventData.description

        # return EventData (only if present)
        return $Record
    }

}



#################
### MAIN

# Test connectivity to Pimcore backend
$ConnectivityCheck = Test-PimcoreBackend -AuthCookie $AuthCookie -PimcoreCsrfToken $PimcoreCsrfToken

If (!($ConnectivityCheck)) {
    Write-Host "`n`n"
    Write-Error "Please validate connectivity to Pimcore"
    Write-Host "`n`n"

    Write-Host -ForegroundColor Yellow "Please follow the steps below to retrieve valid auth data:"
    Write-Host "1) Open you webbrowser and navigate to https://dav360redaktion.alpenverein.de/admin/"
    Write-Host "2) Login to your account (if necessary)"
    Write-Host "3) Open any Dokument or Datenobjekt"
    Write-Host "4) Enter Developer-Mode in your browser (F12 key) and switch to network tab"
    Write-Host "5) Close the document/site in Pimcore by hitting X button"
    Write-Host "6) Inspect the 'unlock-element' event in the network tab"
    Write-Host "7) From Headers copy the values of 'Cookie' and 'x-pimcore-csrf-token'"
    Write-Host "     Examples:"
    Write-Host "          PHPSESSID=63tnaodiheu2025phjd4scuddq   <-- this is your [AuthCookie]"
    Write-Host "          c373a1d22e33cc72b12bbe44e2a8aa3548e32d31   <-- this is your [PimcoreCsrfToken]"
    Write-Host "`n`n"

    Write-Host -ForegroundColor Yellow "Execute script with following syntax:"
    Write-Host ".\Get-PimcoreGroupEvents.ps1 -AuthCookie ""[AuthCookie]"" -PimcoreCsrfToken ""[PimcoreCsrfToken]"""
    Write-Host "`n`n"
}
Else {
# If auth is working, continue ...

    # Get list of all events (multiple event types)
    # Type "Tour": eit
    # Type "Veranstaltung": ei
    $EventList = @()
    $EventList += Get-PimcoreGridJson -FolderId $GruppenFolderId -ClassId eit -AuthCookie $AuthCookie -PimcoreCsrfToken $PimcoreCsrfToken
    $EventList += Get-PimcoreGridJson -FolderId $GruppenFolderId -ClassId ei -AuthCookie $AuthCookie -PimcoreCsrfToken $PimcoreCsrfToken

    # Get event data for all events in list
    $EventReport = @()

    ForEach ($Event in $EventList) {
        # Get event data for single event
        $EventData = Get-EventData -ItemId $Event.id
        If ($EventData) {
            $EventReport += $EventData
        }

        # Add some delay to avoid throttling (not yet happend)
        Start-Sleep -Milliseconds 100
    }

    # Repot generation
    $ReportFilePath = "C:\Temp\DAVRedaktion_Gruppentermine_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').html"
    $EventReport | Sort-Object Termin_Start | Select Gruppe,Titel,Termin_Start,Termin_Ende,Tourenleitung,Veranstaltungsort,Treffpunkt,Beschreibung,ID | Out-HtmlView -HideFooter -PagingLength 10000 -FilePath $ReportFilePath
    Write-Host "Report saved to $($ReportFilePath)"
}


