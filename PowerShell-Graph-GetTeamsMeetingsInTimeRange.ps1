# PowerShell-Graph-GetTeamsMeetingsInTimeRange.ps1
# This script demonstrates how to retrieve a list of Teams meetings within a specified time range for a user via Microsoft Graph API.
# Generated initially with CoPilot.
#  Prompt: create a powershell graph sample which will show the IDs and subjects for a teams meetings within a time range.  Add verbose comments.
#  Read before trying: https://learn.microsoft.com/en-us/graph/api/resources/onlinemeeting?view=graph-rest-1.0
#  Need Application permissions granted:
#    OnlineMeeting.Read.All
#    Be sure to do an administrtor grant after setting the specifc permissions.

<#
.SYNOPSIS
  List Teams meetings (Graph calendar events) within a specified time range and output their Graph event IDs and subjects.

.DESCRIPTION
  This script:
   1) Acquires an OAuth token (client credentials flow).
   2) Calls Microsoft Graph /users/{id|upn}/calendarView with startDateTime/endDateTime.
   3) Selects only the fields needed to detect Teams meetings:
        - id
        - subject
        - start/end
        - isOnlineMeeting
        - onlineMeetingProvider
        - onlineMeeting (contains joinUrl)
   4) Filters to Teams meetings:
        - isOnlineMeeting == $true
        - onlineMeetingProvider == "teamsForBusiness" (common Teams provider value)
          OR onlineMeeting.joinUrl contains "teams.microsoft.com" (practical validation)

.NOTES
  - This outputs the Graph *event* ID (the "id" property returned by /calendarView).
  - For the join link, use onlineMeeting.joinUrl (Graph guidance says not to rely on onlineMeetingUrl, which is being deprecated). [1](https://learn.microsoft.com/en-us/graph/outlook-calendar-online-meetings)
#>

# =========================
# CONFIGURATION (EDIT ME)
# =========================

# Tenant/app auth (App Registration with appropriate permissions, e.g. Calendars.Read)
$TenantId     = "<YOUR_TENANT_ID_GUID>"
$ClientId     = "<YOUR_APP_ID_GUID>"
$ClientSecret = "<YOUR_CLIENT_SECRET>"

# Target mailbox to read calendar from (UPN or user id)
$UserIdOrUpn  = "<someone@contoso.com>"

# Time range (ISO 8601). Use explicit offset or 'Z' for UTC.
# Example range:
$StartDateTime = "2026-03-03T00:00:00-05:00"
$EndDateTime   = "2026-03-06T23:59:59-05:00"

# OPTIONAL: Force Graph to interpret times in a specific timezone for the response.
# (This affects how start/end are returned; the *filter window* is still your provided start/end.)
$PreferTimeZone = "Eastern Standard Time"

# =========================
# 1) GET APP-ONLY TOKEN
# =========================
Write-Verbose "Requesting app-only token via client credentials flow..." -Verbose

$TokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

$TokenBody = @{
  grant_type    = "client_credentials"
  client_id     = $ClientId
  client_secret = $ClientSecret
  scope         = "https://graph.microsoft.com/.default"
}

$TokenResponse = Invoke-RestMethod -Method POST -Uri $TokenUri -Body $TokenBody -ContentType "application/x-www-form-urlencoded"
$AccessToken   = $TokenResponse.access_token

if (-not $AccessToken) {
  throw "Failed to obtain access token."
}

# Common headers for Graph calls
$Headers = @{
  Authorization = "Bearer $AccessToken"
}

# Prefer header for timezone (optional)
if ($PreferTimeZone) {
  # The value is quoted per Graph convention
  $Headers["Prefer"] = "outlook.timezone=`"$PreferTimeZone`""
}

# =========================
# 2) CALL /calendarView IN A LOOP (PAGING)
# =========================

# /calendarView requires startDateTime & endDateTime query params.
# We'll request only the fields we care about to keep the payload smaller.
# We also ask for a larger page size with $top (Graph may still cap it).
$Select = "id,subject,start,end,isOnlineMeeting,onlineMeetingProvider,onlineMeeting"
$Top    = 100

# Build initial URL
# NOTE: We URL-encode query parameter values.
$encodedStart = [System.Web.HttpUtility]::UrlEncode($StartDateTime)
$encodedEnd   = [System.Web.HttpUtility]::UrlEncode($EndDateTime)

$Url = "https://graph.microsoft.com/v1.0/users/$UserIdOrUpn/calendarView" +
       "?startDateTime=$encodedStart" +
       "&endDateTime=$encodedEnd" +
       "&`$select=$Select" +
       "&`$top=$Top"

Write-Verbose "Initial calendarView URL: $Url" -Verbose

$AllEvents = New-Object System.Collections.Generic.List[object]

while ($Url) {
  Write-Verbose "Fetching page: $Url" -Verbose

  $Resp = Invoke-RestMethod -Method GET -Uri $Url -Headers $Headers

  if ($Resp.value) {
    foreach ($e in $Resp.value) {
      $AllEvents.Add($e) | Out-Null
    }
    Write-Verbose ("Collected {0} events so far..." -f $AllEvents.Count) -Verbose
  }

  # Paging: if Graph returns @odata.nextLink, keep going
  $Url = $Resp.'@odata.nextLink'
}

Write-Verbose ("Total events retrieved in window: {0}" -f $AllEvents.Count) -Verbose

# =========================
# 3) FILTER TO TEAMS MEETINGS
# =========================
<#
Teams meetings are typically:
 - isOnlineMeeting == true
 - onlineMeetingProvider == "teamsForBusiness"
And they usually have:
 - onlineMeeting.joinUrl containing "teams.microsoft.com"

We filter conservatively:
  A) Must be online meeting
  B) Must be Teams provider OR have Teams joinUrl
#>

$TeamsMeetings =
  $AllEvents |
  Where-Object {
    $_.isOnlineMeeting -eq $true -and (
      $_.onlineMeetingProvider -eq "teamsForBusiness" -or
      ($_.onlineMeeting.joinUrl -and $_.onlineMeeting.joinUrl -match "teams\.microsoft\.com")
    )
  }

Write-Verbose ("Teams meetings found: {0}" -f ($TeamsMeetings | Measure-Object | Select-Object -ExpandProperty Count)) -Verbose

# =========================
# 4) OUTPUT IDS + SUBJECTS
# =========================

<# 
# Output the Graph event ID + subject (plus joinUrl for validation/troubleshooting)
$TeamsMeetings |
  Sort-Object { $_.start.dateTime } |
  Select-Object `
    @{Name="EventId"; Expression = { $_.id }}, `
    @{Name="Subject"; Expression = { $_.subject }}, `
    @{Name="Start";   Expression = { $_.start.dateTime }}, `
    @{Name="End";     Expression = { $_.end.dateTime }}, `
    @{Name="Provider";Expression = { $_.onlineMeetingProvider }}, `
    @{Name="JoinUrl"; Expression = { $_.onlineMeeting.joinUrl }} |
  Format-Table -AutoSize
#>

  # Export to CSV (uncomment if needed)
  $OutputCsvPath = "c:\test\TeamsMeetings_$($UserIdOrUpn)_$((Get-Date).ToString("yyyyMMdd_HHmmss")).csv"
  $TeamsMeetings |
  Sort-Object { $_.start.dateTime } |
  Select-Object `
    @{Name="EventId"; Expression = { $_.id }}, `
    @{Name="Subject"; Expression = { $_.subject }}, `
    @{Name="Start";   Expression = { $_.start.dateTime }}, `
    @{Name="End";     Expression = { $_.end.dateTime }}, `
    @{Name="Provider";Expression = { $_.onlineMeetingProvider }}, `
    @{Name="JoinUrl"; Expression = { $_.onlineMeeting.joinUrl }} |
  Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

 

 
