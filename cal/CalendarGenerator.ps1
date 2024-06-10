## Load some private values from another file
# This shifts them off into a file ignored by git so I'm not uploading some PII for all to see :)
# $WorkDir = <local directory I'm working from >
# $CalendarUserID = <My personal email address>
. .\CalendarPrivateInfo.ps1


If (-not $GenerateCalFor){ 
    $GenerateCalFor = "2024-06"
}



$OrgMap = @"
Code,Name,NameShort,Site
ABGG,Adelaide Board Game Group,ABGG,https://abgg.au
TLD,The Lost Dice,The Lost Dice,https://thelostdice.com.au
SSG,South Side Games,South Side Games,https://
AM,Adelaide Megagames,Adelaide Megagames,https://linktr.ee/adelaidemegagames
SMB,Smorgasboard,Smorgasboard,
GAMES,Adelaide Uni GAMES Club,GAMES,
AVCon,AVCon,AVCon,https://avcon.org.au
"@ | ConvertFrom-Csv

$TermTable = @"
Context,Term,Replacement
*,North Adelaide Community Centre And Library,NACC
calendar,North Adelaide Community Centre And Library,North Adelaide
agenda,North Adelaide Community Centre And Library,North Adelaide (NACC)
calendar,Payneham Community Centre,Payneham
calendar,Parks Library,Parks
*,297 Diagonal Road,297 Diagonal Road
calendar,Hartley Building,Adelaide Uni
"@ | ConvertFrom-Csv

$NoCaseChange = @("avcon")

$T = @{
    "special"="Special";
    "extra"="Extra";
    "regular"="Recurring Schedule";
    "official"="ABGG Official Events"
    "community"="Community Events";
}

Function Term ([string]$Term,[switch]$Wildcard,[string]$Context="*") {
    if ($Wildcard) {
        $Replacements = $TermTable | Where-Object {$_.Term -like "*$Term*"}
    } else {
        $Replacements = $TermTable | Where-Object {$_.Term -eq $Term}
    }
    ## No replacement found
    ## OR the needed context isn't in the TermTable + there's no wild context present
    if ($null -eq $Replacements -or ($Replacements.context -notcontains $Context -and $Replacements.context -notcontains "*")) {
        $Out = $Term
    } else {
        $Replacement = $Replacements | Where-Object {$_.Context -eq $Context}
        ## If null use the wildcard
        if ($null -eq $Replacement) {$Replacement = $Replacements | Where-Object {$_.Context -eq "*"}}
        $Out = $Replacement.Replacement
    }
    If ($NoCaseChange -contains $Out) {
        $Out
    } else {
        (Get-Culture).TextInfo.ToTitleCase($Out).trim()
    }
}

Function MapOrg($Code) {
    $OrgMap | Where-Object {$_.Code -eq $code}
}

Function GetNumberSuffix ($Num) {
    switch ($Num) {
        1{"st"}
        2{"nd"}
        3{"rd"}
        default{"th"}
    } 
}

Function GetDayOFWeekIndex ([string]$DoW) {
    switch -Wildcard ($DoW){
        "Sun*"{7}
        "Mon*"{1}
        "Tue*"{2}
        "Wed*"{3}
        "Thu*"{4}
        "Fri*"{5}
        "Sat*"{6}
    }
}
Function GetMonthMetaPack ([string]$DateCode) {
    $Year = $DateCode.Split("-")[0]
    $Month = $DateCode.Split("-")[1]
    $FirstDay = Get-Date -Year $Year -Month $Month -Day 1
    $DatesInMonth = 0..31 | ForEach-Object {Get-Date $FirstDay.AddDays($_) -Format "yyyy-MM-dd"} | Where-Object {$_.StartsWith($DateCode) -eq $true}

    [pscustomobject]@{
        DateCode = $DateCode
        Year = $Year
        Month = $Month
        FirstDayOfWeek = $FirstDay.DayOfWeek
        DatesInMonth = $DatesInMonth
        DayCount = $DatesInMonth.Count
        MonthName = Get-Date $FirstDay -Format "MMMM"
    }
}
Function ExtractNotesFromBody ($EventInstance) {
    $Body = $EventInstance.Body.Content
    $Body = $Body.Replace("`r","").Replace("`n","")
    ## Wildcard search for an opening brace, followed by HTML-encoded quote, any amount of content, then a closing brace
    If ($Body -like "*{&quot;*}*") {
        ## Replace HTML encoded quotes with real ones; then use RegEx to pull out the JSON from the event body.
        ## This is... not great, but it works.
        $Body = $Body.Replace("&quot;","`"")
        if ($Body -match ">.*(?<json>{.*}).*<") {
            $Matches.json | ConvertFrom-Json
        } 
    }
}

Function Tabs ([int]$TL) {
    "`r`n" + (0..$TL | ForEach-Object {"  "}) -join ''
}
Function getHourWritten([int]$Time) {
    $Time = $Time%12
    Switch ($Time) {1{"One"};2{"Two"};3{"Three"};4{"Four"};5{"Five"};6{"Six"};7{"Seven"};8{"Eight"};9{"Nine"};10{"Ten"};11{"Eleven"};12{"Twelve"};0{"Twelve"}}
}

Function RenderAddress ($Location) {
    $Place = $Location.Location
    $DN = $Place.DisplayName
    $A = $Place.Address
    if ($null -ne $A.Street -and $null -ne $A.City) {
        ## Street info, with suburb
        $AddressStandardised = "$($A.Street), $($A.City)$(if ($null -ne $A.PostalCode) {" $($A.PostalCode)"})"
    } elseif ($null -ne $A.Street) {
        ## Street info, but no suburb
        $AddressStandardised = "$($A.Street)"
    } else {
        $AddressStandardised = $null
    }
    
    if ($null -eq $DN) {
        ## No displayname
        if ($null -eq $Place.Address) {
            #No address
            "No location specified"
        } else {
            # No DN, but had addres
            $AddressStandardised
        }
    } else {
        if ($null -ne $AddressStandardised) {
            ## Gold standard - both DN with address
            "$DN ($AddressStandardised)"
        } else {
            #Just a DN
            $DN
        }
    }
}

Function GetAgendaCell ($IconClasses,$Content,$ContentClasses=$Null,[switch]$Force2Col) {
    [pscustomobject]@{
        html = "<div><span class='mdi $($IconClasses -join ' ')'></span></div><div$(if ($null -ne $ContentClasses) {" class='$($ContentClasses -join ' ')'"})>$($Content)</div>"
        ContentLength = $Content.Length
        order = $EventCells.count
        Force2Col = $Force2Col
    }
}

Function GetAgendaListItem ($EventGroup,[int]$TL) {
    ## At this point events should have been filtered to make sure that Subject, Organiser, and Location, and Price are all the same

    #Therefore sample first item to get these key details
    $EventInstance = $_.Group[0]
    $icon = GetEventIcon $EventInstance
    $A = $_.Group[0].Location.Address

    $EventCells = @()
    $EventCells += GetAgendaCell -IconClasses "mdi-$icon","effects-$icon" -ContentClasses "eventtitle" -Content "$(Term -Context "agenda" $EventInstance.DispTitle)" -Force2Col
    $EventCells += GetAgendaCell -IconClasses "mdi-account-group" -Content "$(Term -Context "agenda"  (MapOrg $_.Group[0].ManagedBy).Name)" -Force2Col

    ## Then list out sets of cells based on each combination of Day, and time seen -- Summarising all 
    $_.Group | Group-Object DayOfWeek,DispTimeSlot | ForEach-Object {
        $TimeInstance = $_.Group[0]
        $EventCells += GetAgendaCell -IconClasses "mdi-calendar-week" -Content "$($TimeInstance.DayOfWeek) $(($_.Group | Select-Object -ExpandProperty dispDay) -Join "/")" -Force2Col
        $EventCells += GetAgendaCell -IconClasses "mdi-clock-time-$((getHourWritten  $_.Group[0].DispStart.Hour).toLower())-outline" -Content "$(Term  -Context "agenda" $TimeInstance.dispTimeSlot)" -Force2Col 
    }
    $EventCells += GetAgendaCell -IconClasses "mdi-office-building-marker" -Content "$(Term $_.Group[0].DispLocation -Context "agenda")" -Force2Col
    if ($null -ne $A.Street) {
        $EventCells += GetAgendaCell -IconClasses "mdi-map" -Content "$(Term -Context "agenda" $A.Street), $(Term -Context "agenda" $A.City) $(Term -Context "agenda" $A.Postalcode)"
    }
    if ($null -ne $EventInstance.Price) {
        $EventCells += GetAgendaCell -IconClasses "mdi-$(if ($EventInstance.IsFree) {"currency-usd-off"} else {"currency-usd"})" -Content ($EventInstance.price)
    }
                                        
    if ($null -ne $EventInstance.notes -and $EventInstance.Notes -ne "") {
        $EventCells += GetAgendaCell -IconClasses "mdi-note-text" -content ($EventInstance.notes)
    }  

    ## Render content
    $TL = $TL+1
    "$(Tabs $TL)<div class='agenda-event-container'>"
     $TL = $TL+1
        "$(Tabs $TL)<div class='agenda-event $(GetEventCSSClassNames $EventInstance)'>"
            $TL = $TL+1
            "$(Tabs $TL)<div class='agenda-event-grid-2col'>"
            $TL = $TL+1
                $EventCells | Where-Object {$_.ContentLength -le $MaxCharsForTwoCol -or $_.Force2Col -eq $true} | Sort-Object Order | ForEach-Object {
                    "$(Tabs $TL)$($_.html)"
                }
            $TL = $TL-1
            "$(Tabs $TL)</div>"
            $TL = $TL-1

            "$(Tabs $TL)<div class='agenda-event-grid-1col'>"
            $TL = $TL+1
            $EventCells | Where-Object {$_.ContentLength -gt $MaxCharsForTwoCol -and $_.Force2Col -ne $true} | Sort-Object Order  | ForEach-Object {
                "$(Tabs $TL)$($_.html)"
            }

            "$(Tabs $TL)</div>"
            $TL = $TL-1
        "$(Tabs $TL)</div>"
        $TL = $TL-1
    "$(Tabs $TL)</div>"
    $TL = $TL-1
}

Function GenerateEventAgendaList ($Cal, $MaxCharsForTwoCol=45) {
    $TL = 0
    $Cal | Group-Object IsOfficial | Sort-Object name -Descending | ForEach-Object {
        $GrpIsOfficial = $_
        $TL = $TL+1
        ## Two groups -- ABGG Official, then Community events
        $GroupHeader = If ($GrpIsOfficial.Group[0].IsOfficial) {$T['official']} else {$T['community']}
        "$(Tabs $TL)<div class='agenda-management-group'><h1>$GroupHeader</h1>"

        $GrpIsOfficial.Group | Group-Object IsSpecial | Sort-Object name -Descending | ForEach-Object {
            $GrpIsSpecial = $_
            ## Group by special, then all non-special events
            $GrpIsSpecial.Group | Group-Object IsRecurring | Sort-Object name | ForEach-Object {
                $GrpIsRecurring = $_
                $TL = $TL+1
                ##Group by non-recurring, then recurring
                $Grouping = $_.Group[0].Grouping
                $GroupingCount = ($GrpIsOfficial.Group | Select-Object -ExpandProperty Grouping | Sort-Object -Unique | Measure-Object).Count
                If ($GroupingCount -gt 1){
                    "$(Tabs $TL)<div class='agenda-eventtype-group'>"#<h2>$Grouping</h2>"
                }
                $GrpIsRecurring.Group | Sort-Object DayIndex | Group-Object Subject,ManagedBy,DispLocation,Price | ForEach-Object {
                    $TL = $TL+1
                        GetAgendaListItem -EventGroup $_ -TL $TL
                    $TL = $TL-1
                }
                If ($GroupingCount -gt 1){
                    "$(Tabs $TL)</div>"
                }
                $TL = $TL-1
            }
        }
        "$(Tabs $TL)</div>"
        $TL = $TL-1
    }
}

Function GetEventCSSClassNames ($EventInstance) {
    $Org = MapOrg $EventInstance.ManagedBy
    $Classes = @() 
    $Classes += "venue-"+$(switch -Wildcard ($EventInstance.DispLocation) {
        "North Adelaide Community Centre*" {"nacc"}          
        "Parks Library" {"parks"}  
        "Payneham Community Centre" {"payneham"}  
        "San Churro*" {"sanchurro"}  
        "The Lost Dice" {"tld"}  
        default {"$($EventInstance.DispLocation.Replace(" ",'')) venue-unknown"}
    })
    $Classes += switch ($EventInstance.IsOfficial) {
        $true {"IsOfficial"}  
        $false {"IsCommunity"}  
    }
    $Classes += switch ($EventInstance.IsSpecial) {
        $true {"IsSpecial"}  
    }
    $Classes += switch ($EventInstance.IsExtra) {
        $true {"IsExtra"}   
    }
    $Classes += switch ($EventInstance.IsRecurring) {
        $true {"IsRecurring"}   
    }
    $Classes += "managedby-$($EventInstance.ManagedBy)"
    $Classes += "logo-image"
    $Classes += "logo-image-$($Org.Code)"
    if ($null -ne $EventInstance.Price) {
        if ($EventInstance.IsFree) {
            "IsFree"
        } else {
            "IsNonFree"
        }
    }

    $Classes -join " "
}

Function GetEventIcon($EventInstance) {
    Function EvaluteIcons ($EventInstance) {
        if ($EventInstance.ForceIcon -ne "" -and $null -ne $EventInstance.ForceIcon) {
            $EventInstance.ForceIcon
        }
        if ($EventInstance.IsSpecial) {
            "calendar-star"
        }
        if ($EventInstance.IsExtra) {
            "calendar-plus"
        }
        if ($EventInstance.IsRecurring) {
            "calendar-sync"
        }
    }
    $IconList = EvaluteIcons $EventInstance
    if ($null -eq $IconList) {
        
    } else {
        $IconList | Select-Object -first 1
    }
}

Function GetCalendarEventCell ($EventInstance) {
    $MB = Term -Context "calendar" (MapOrg $EventInstance.ManagedBy).NameShort
    $Loc = Term -Context "calendar" $EventInstance.DispLocation
    
    if ($MB -eq $Loc -or $null -eq $MB) {
        ## Avoid redundant 'By XYZ at XYZ' and give a custom line
        $CoreInfo = "At $Loc"
    } else {
        $CoreInfo = "$MB at $Loc"
    }

    $TL = 3
    $Classes = @("cal-event")
    $Classes += GetEventCSSClassNames $EventInstance
    "$(Tabs $TL)<div class='$($Classes -join " ")'>"
    $Icon = GetEventIcon $EventInstance
    if ($Icon -ne "calendar-sync") {$icon = "mdi mdi-$icon effects-$icon"} else {$icon = $null}
    if ($null -ne $icon) {
        ## Only include this if there's actually an icon to show.  tbc if this is a good idea tbeh
        "$(Tabs ($TL+1))<div class='cal-event-icon'><span class='$($icon)'></span></div>"
    }
    "$(Tabs ($TL))<div class='cal-event-content'>"
    "$(Tabs ($TL+1))<div class='cal-event-title'>$(Term -Context "cal" $EventInstance.DispTitle)</div>"
    if ($EventInstance.Subtitle -ne "" -and $null -ne $EventInstance.Subtitle) {
        "$(Tabs ($TL+1))<div class='cal-event-subtitle'>$($EventInstance.Subtitle)</div>"
    }
    "$(Tabs ($TL+1))<div class='cal-event-core-info'>$CoreInfo</div>"
    "$(Tabs $TL)</div></div>"
    #"  <div class='event $(GetEventCSSClassNames $EventInstance)'><span class='eventicon mdi mdi-$(GetEventIcon $EventInstance)'></span>$($EventInstance.DispTitle)</div>"
}

Function GenerateCalendarCells ($Cal, $MonthMeta) {
    $CellsToFill = 6*7
    #$EmptyCells = $CellsToFill-$MonthMeta.DayCount
    $FirstCellToFill = (GetDayOFWeekIndex $MonthMeta.FirstDayOfWeek)
    $TL = 4
    (1..$CellsToFill | ForEach-Object {
        $Cell = $_
        $Date = (Get-Date $MonthMeta.DatesInMonth[0]).AddDays((($FirstCellToFill*-1)+$Cell)).Day
        If ($Cell -lt $FirstCellToFill -or $Cell -ge $FirstCellToFill+$MonthMeta.DayCount) {
            ## Blank Cells for the empty dates
            "$(Tabs $TL)<div>
                $(Tabs ($TL+1))<div class='inactive-month-header'>$Date</div>
                $(Tabs ($TL+1))<div class='inactive-month-body'></div>
            $(Tabs ($TL))</div>"
        } else {
            $EventsToday = $Cal | Where-Object {$_.DispStart.Day -eq $date} | Sort-Object IsOfficial -Descending
            "$(Tabs $TL)<div>
                $(Tabs ($TL+1))<div class='active-month-header'>$Date</div>
                $(Tabs ($TL+1))<div class='active-month-body'>$($EventsToday | ForEach-Object {Tabs ($TL+1);(GetCalendarEventCell $_)})
                $(Tabs ($TL+1))</div>
            $(Tabs ($TL))</div>"
        }
    }) -join "`r`n"
}







Import-Module Microsoft.Graph.Calendar

Write-Verbose -Verbose "Connecting to MS Graph"
$MGConnection = Connect-MgGraph -Scopes "Calendars.ReadBasic","Calendars.Read","Calendars.Read.Shared" 
Write-Verbose -Verbose "Getting Calendars for $CalendarUserID"
$CalendarList = Get-MgUserCalendar -UserId $CalendarUserID
#$CalendarGroupList = Get-MgUserCalendarGroup -UserId $CalendarUserID


$CalMeta = $CalendarList | Where-Object {$_.Name -eq "ABGG-Public"}
Write-Verbose -Verbose "Getting Events (MgUserCalendarEvent)"
$Events = Get-MgUserCalendarEvent -CalendarId $CalMeta.Id -UserId $CalendarUserID -PageSize 999 

Function GetAndProcessEventsFromGraph ($GenerateCalFor) {

    Write-Verbose -Verbose "Getting Calendar View (MgUserCalendarView)"
    $queryStart = (Get-Date $MonthMeta.DatesInMonth[0] -Hour 0 -Minute 0 -Second 0 -Format 'o') + "+09:30"
    $queryEnd = (Get-Date $MonthMeta.DatesInMonth[-1] -Hour 23 -Minute 59 -Second 59 -Format 'o') + "+09:30"
    $Cal = Get-MgUserCalendarView -CalendarId $CalMeta.Id -UserId $CalendarUserID -StartDateTime $queryStart -EndDateTime $queryEnd -PageSize 999 -Headers @{'Prefer'='outlook.timezone="Cen. Australia Standard Time"'}

    $Cal | ForEach-Object {
        $Item = $_       
        $BodyExtract = ExtractNotesFromBody $Item
        ## Include it here with all properties for potential bespoke usage later
        $Item | Add-Member -MemberType NoteProperty -Name 'BodyExtract' -Value $BodyExtract
        ## Extract specific fields with some structure behind them
        $Item | Add-Member -MemberType NoteProperty -Name 'Price' -Value $BodyExtract.price
        $Item | Add-Member -MemberType NoteProperty -Name 'Notes' -Value $BodyExtract.notes
        $Item | Add-Member -MemberType NoteProperty -Name 'Subtitle' -Value $BodyExtract.Subtitle
        $Item | Add-Member -MemberType NoteProperty -Name 'Repeats' -Value $BodyExtract.Repeats
        $Item | Add-Member -MemberType NoteProperty -Name 'ForceIcon' -Value $BodyExtract.ForceIcon
        $Item | Add-Member -MemberType NoteProperty -Name 'IsFree' -Value ($Item.Price -like "*free*" -or $Item.Price -like "*donation*")                         
        $Item | Add-Member -MemberType NoteProperty -Name 'dispStart' -Value (Get-Date $Item.Start.DateTime)
        $Item | Add-Member -MemberType NoteProperty -Name 'dispEnd' -Value (Get-Date $Item.End.DateTime)
        $Item | Add-Member -MemberType NoteProperty -Name 'Date' -Value (Get-Date $Item.Start.DateTime -format "yyyy-MM-dd")
        $Item | Add-Member -MemberType NoteProperty -Name 'dispLocation' -Value $Item.Location.DisplayName
        $DispTimeSlot = "$(Get-Date $Item.dispStart -Format "h:mm tt")-$(Get-Date $Item.DispEnd -Format "h:mm tt")"
        If ($Item.dispStart.Minute -eq 0 -and $Item.dispEnd.Minute -eq 0) {
            ## If we start, and end on an hour then don't show the minutes
            $DispTimeSlot = "$(Get-Date $Item.dispStart -Format "h tt") to $(Get-Date $Item.DispEnd -Format "h tt")"
        }
        $Item | Add-Member -MemberType NoteProperty -Name 'dispTimeslot' -Value $DispTimeSlot 
        $Item | Add-Member -MemberType NoteProperty -Name 'dispDay' -Value "$($Item.dispStart.day)$(GetNumberSuffix $Item.dispStart.Day)"   
        $Item | Add-Member -MemberType NoteProperty -Name 'DayIndex' -Value $Item.dispStart.day
        $Item | Add-Member -MemberType NoteProperty -Name 'dispDaySuffix' -Value "$Suffix"   
        $Item | Add-Member -MemberType NoteProperty -Name 'DayOfWeek' -Value (Get-Date $Item.Start.DateTime).DayOfWeek
        $Item | Add-Member -MemberType NoteProperty -Name 'Address' -Value $Item.Location.Address
        $Item | Add-Member -MemberType NoteProperty -Name 'IsOfficial' -Value ($Item.Subject -like "ABGG |*")
        $Item | Add-Member -MemberType NoteProperty -Name 'IsSpecial' -Value ($Item.Categories -contains "Special Events")
        $Item | Add-Member -MemberType NoteProperty -Name 'IsRecurring' -Value ($null -ne $Item.Repeats -or $Item.Categories -contains "Recurring Events" -or ($Cal | Where-Object {$_.Subject -eq $Item.Subject} | Measure-Object).Count -gt 1 -or $null -ne ($Events | Where-Object {$_.Subject -eq $Item.Subject}).Recurrence.Pattern.Interval -and $Item.BodyExtract.BreakReccurance -ne $true)
        $Grouping = if ($Item.IsSpecial) {$T['special']} elseif (-not $Item.IsRecurring) {$T['extra']} else {$T['regular']}
            $Item | Add-Member -MemberType NoteProperty -Name 'IsExtra' -Value ($Grouping -eq $T['extra'])
        $Item | Add-Member -MemberType NoteProperty -Name 'Grouping' -Value $Grouping
        $SubSplit = $Item.Subject.Split("|").Trim()
        $Item | Add-Member -MemberType NoteProperty -Name 'ManagedBy' -Value $SubSplit[0]
        $Item | Add-Member -MemberType NoteProperty -Name 'DispTitle' -Value $SubSplit[1]
    }
    $Cal
}

Function GenerateCalendar ($GenerateCalFor) {
    if ($GenerateCalFor -like "*.P") {
        $IsPreview = $true
        $GenerateCalFor = $GenerateCalFor.Split(".")[0]
    } 
    $MonthMeta = GetMonthMetaPack $GenerateCalFor
    $Cal = GetAndProcessEventsFromGraph $GenerateCalFor
    $Month = "$($MonthMeta.MonthName) $($MonthMeta.Year)"
    Write-Verbose -Verbose "Creating HTML output to .\$($GenerateCalFor).htm"
    $Template = Get-Content .\HTML_Template.htm
    $Template = $Template.Replace("REPLACEME_CALBODY",(GenerateCalendarCells $Cal $MonthMeta))
    $Template = $Template.Replace("REPLACEME_AGENDA",(GenerateEventAgendaList $Cal))
    If ($IsPreview) {$Month += "&nbsp; <span class='previewtext'>(Preview)</span><div class='previewbanner'>PREVIEW</div>"}
    $Template = $Template.Replace("REPLACEME_MONTHNAME",$Month)
    $Template = $Template.Replace("REPLACEME_ANNOUNCEMENTS","")

    $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
    [System.IO.File]::WriteAllLines((Join-Path $pwd ".\$($GenerateCalFor).htm"), $Template, $Utf8NoBomEncoding)
}



GenerateCalendar "2024-06"; GenerateCalendar "2024-07.P"