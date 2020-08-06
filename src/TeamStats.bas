Attribute VB_Name = "TeamStats"
''
' VBA-ScrumTeamStats v1
' (c) Ben Doughton
'
' JIRA Scrum Team Stats VBA
'
' @Dependencies:
'               Mod - Base64Encoding
'               Mod - JsonConverter
'
' Note: This is designed to be a standalone module for TeamStats so if there are other modules
'       from the same family of Jira apicalls there could be duplication of code
'

' @author ben.doughton@lch.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Public RemainingSprintTime As Long
Public DaysInSprint As Long
Public LastSprintName As String
Public LastSprintId As Integer

Sub GetTeamStats()

''
' This should be run by the user and sets up all the underlying api calls to get the teams stats

'' Known limitations with this macro:
' (1) Work in progress - run each call individually by commenting out the other api calls

Dim callResult As Long
Dim url As String
Dim l As Integer
Dim boardJql As String
Dim blnRoll As Boolean
Dim msgCaption As String

'Fetch the base url
url = PublicVariables.JiraBaseUrl
'Fetch the board query
boardJql = "Team = 81 AND CATEGORY = calm AND NOT issuetype in (Initiative) ORDER BY Rank ASC"

' Allow the user to confirm the base url before executing the code
If MsgBox("You are using this Jira url: " & url, vbOKCancel) = vbCancel Then
    msgCaption = "Cancelled action"
    Exit Sub
End If

If MsgBox("Do you want to roll the previous data?", vbYesNo) = vbYes Then
    blnRoll = True
Else
    blnRoll = False
End If

'Check if a user is logged in and if not request login up to 3 times
'The MyCredentials API is forcing a system login so my custom form only
'shows and captures login details if the system login fails
l = 0
Do While RestApiCalls.MyCredentials(encodedAuth, url) <> 200
    Frm_JiraLogin.Show
    l = l + 1
    If l = 3 Then
        MsgBox ("Too Many Unsuccessful Login Attempts")
        Exit Sub
    End If
Loop

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

''Rollstats True of False
funcRollStats (blnRoll)

''Fetch Data from Api
callResult = TeamStats.funcGet3MonthsOfDoneJiras(encodedAuth, url, boardJql, "In Progress", "Done", 0, 2)
'callResult = TeamStats.funcGetIncompleteJiras(encodedAuth, url, boardJql, 0, 2)
'callResult = TeamStats.funcGetVelocity(encodedAuth, url, CStr(RapidBoardId))
'callResult = TeamStats.funcPostTeamsFind(encodedAuth, url)

''Need to save the RemainingSprintTime to the right cell
'callResult = TeamStats.funcGetSprintBurnDown(encodedAuth, url, CStr(RapidBoardId), CStr(ws_VelocityData.Range("A2").value))

'Output error if call was not successful
If callResult <> 200 Then
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox ("Error")
    Exit Sub
End If

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

MsgBox ("Success")

End Sub

Private Function funcRollStats(ByVal Enabled As Boolean)

''
' Function will roll over previous data from the ws_TeamStats worksheet ready to add new data
'
''

'' Known limitations with this macro:
' (1) ws_TeamStats worksheet has to exist (with headings and formatting) for the macro to run

If Enabled Then
    With ws_TeamStats
        .Range("AW22:BA44").value = .Range("AX22:BB44").value
    End With
End If

End Function

Private Function funcGet3MonthsOfDoneJiras(ByVal auth As String, ByVal baseUrl As String, _
            ByVal boardJql As String, ByVal inProgressState As String, ByVal endProgressState As String, _
            ByRef startAtVal, r As Integer) As Long

''
' Source Jiras that are in a Done state and were included in a sprint and had a fixVersion that was updated in the last 24 weeks
' Then cycle through and get the sub-tasks for all the issues from the first api call

'
' @param {String} auth, baseUrl, boardJql, inProgressState, endProgressState
' @param {Integer} startAtVal, r
' @write {ws_LeadTimeData} & {ws_WiPData}
' @apicalls 1x{get search standardissuetypes} ?x{get search subtaskissuetypes}
' @return {Long} status of first apicall
''

'' Known limitations with this macro:
' (1) Can't handle different inProgressState and endProgressState by issue type
' (2) needs to be updated to run for a smaller number of maxresults
' (3) worksheets are not scrubbed the first time the macro is called
' (4) worksheets have to exist (with headings) for the macro to run
' (5) there is no error handling around the second api call to get the subtasks which could fail
' (6) ws_LeadTimeData sheet needs to be made active

Dim apicall As String
Dim jql As String
Dim http, JSON, Item, fixversion, history, changeitem, sprint As Object
Dim i, h, c As Integer
Dim CleanJSON As String
Dim rng_author As Range
Dim rng_issue As Range
Dim WiPRow As Integer
Dim rng_Parent As Range
Dim col As Integer
Dim dictResourceNm As Dictionary
Dim dictTimeLoggedToStory As Dictionary
Dim collIssueKey As New Collection

jql = "fixversion changed after -24w AND " & _
        "fixVersion is not EMPTY AND " & _
        "Sprint is not EMPTY AND " & _
        "NOT issuetype in (Initiative,Epic,Test,subTaskIssueTypes()) AND " & _
        "statusCategory in (Done) AND " & _
        boardJql

apicall = _
    baseUrl & "rest/api/latest/search?jql=" _
    & jql _
    & "&fields=" _
        & "key," _
        & "issuetype," _
        & "fixVersions," _
        & "resolutiondate," _
        & sprints & "," _
        & "created," _
        & "changelog" _
        & "&startAt=" _
        & 0 _
    & "&maxResults=1000&expand=changelog"
            
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

funcGet3MonthsOfDoneJiras = http.Status

Debug.Print (apicall)

If http.Status = 200 Then
    Set dictTimeLoggedToStory = New Scripting.Dictionary
    CleanJSON = CleanSprintsAndCustomFields(http.responseText)
    Set JSON = ParseJson(CleanJSON)
    startAtVal = startAtVal + 1000 'Increment the next start position based on maxResults above
    i = 1 'reset the issue to 1
    WiPRow = r
    For Each Item In JSON("issues")
        Set dictResourceNm = New Dictionary
        h = 1 'reset the change history to 1
        If CDate(JSON("issues")(i)("fields")("fixVersions")(1)("releaseDate")) >= DateAdd("m", -3, "01/" & Month(Now()) & "/" & Year(Now())) Then 'Only include if the release date was in the last 3 months
            With ws_LeadTimeData
                .Cells(r, 1).value = JSON("issues")(i)("id")
                .Cells(r, 2).value = JSON("issues")(i)("key")
                .Cells(r, 3).value = JSON("issues")(i)("fields")("issuetype")("name")
                .Cells(r, 4).value = JSON("issues")(i)("fields")("created")
                .Cells(r, 5).value = sprint_ParseString(JSON("issues")(i)("fields")("sprints")(1), "startDate")
                If JSON("issues")(i)("fields")("fixVersions")(1)("releaseDate") > JSON("issues")(i)("fields")("created") Then
                    .Cells(r, 6).value = JSON("issues")(i)("fields")("fixVersions")(1)("releaseDate") 'Always use the 1st fixVersion, even if there are multiple
                Else
                    .Cells(r, 6).value = Left(JSON("issues")(i)("fields")("resolutiondate"), 10) 'use the resolution date if there is no fixVersion. Note: this can lead to incorrect deployment frequency
                    
                End If
                For Each history In JSON("issues")(i)("changelog")("histories")
                    c = 1 'reset the change item to 1
                    For Each changeitem In JSON("issues")(i)("changelog")("histories")(h)("items")
                        If JSON("issues")(i)("changelog")("histories")(h)("items")(c)("field") = "status" Then
                            Select Case JSON("issues")(i)("changelog")("histories")(h)("items")(c)("toString")
                                Case inProgressState 'enter the date the issue transitioned to its inProgressState
                                    ws_WiPData.Cells(WiPRow, 1).value = JSON("issues")(i)("id")
                                    ws_WiPData.Cells(WiPRow, 2).value = JSON("issues")(i)("key")
                                    ws_WiPData.Cells(WiPRow, 3).value = JSON("issues")(i)("fields")("issuetype")("name")
                                    For Each fixversion In JSON("issues")(i)("fields")("fixVersions")
                                        ws_WiPData.Cells(WiPRow, 6).value = JSON("issues")(i)("fields")("fixVersions")(1)("releaseDate")  'Always use the 1st fixVersion, even if there are multiple
                                    Next
                                    ws_WiPData.Cells(WiPRow, 4).value = JSON("issues")(i)("changelog")("histories")(h)("created")
                                Case endProgressState 'enter the date the issue transitioned to its endProgressState
                                    ws_WiPData.Cells(WiPRow, 5).value = JSON("issues")(i)("changelog")("histories")(h)("created")
                                    WiPRow = WiPRow + 1
                            End Select
                        ElseIf JSON("issues")(i)("changelog")("histories")(h)("items")(c)("field") = "timespent" Then
                            dictResourceNm(JSON("issues")(i)("changelog")("histories")(h)("author")("key")) = Val(JSON("issues")(i)("changelog")("histories")(h)("items")(c)("toString"))
                            Set rng_author = ws_LeadTimeData.Rows(1).Find(JSON("issues")(i)("changelog")("histories")(h)("author")("key"), LookIn:=xlValues, LookAt:=xlWhole)
                            If rng_author Is Nothing Then
                                ws_LeadTimeData.Range("A1").End(xlToRight).Offset(0, 1).value = JSON("issues")(i)("changelog")("histories")(h)("author")("key")
                            End If
                        End If
                        c = c + 1
                    Next
                    h = h + 1
                Next
            collIssueKey.Add dictResourceNm, JSON("issues")(i)("key")
            End With
            r = r + 1 'increment the row
        End If
        i = i + 1 'increment the issue
    Next
    
    '' This next section cycles through all the sub-tasks and adds up the time logged to each
    
    For Each rng_Parent In ws_LeadTimeData.Range("B2:B" & ws_LeadTimeData.Range("A1").End(xlDown).Row)
        apicall = _
        baseUrl & "rest/api/latest/search?jql=" _
        & "Parent = " & rng_Parent.value _
        & "&fields=" _
            & "key," _
            & "issuetype," _
            & "changelog" _
            & "&startAt=" _
            & 0 _
        & "&maxResults=1000&expand=changelog"
        
        Debug.Print apicall
        
        Set http = CreateObject("MSXML2.XMLHTTP")
    
        http.Open "GET", apicall, False
        http.setRequestHeader "Content-Type", "application/json"
        http.setRequestHeader "X-Atlassian-Token", "no-check"
        http.setRequestHeader "Authorization", "Basic " & auth
        http.Send
        Set JSON = ParseJson(http.responseText)
        i = 1 'reset the issue to 1
        For Each Item In JSON("issues")
            h = 1 'reset the change history to 1
            With ws_LeadTimeData
                For Each history In JSON("issues")(i)("changelog")("histories")
                    c = 1 'reset the change item to 1
                    For Each changeitem In JSON("issues")(i)("changelog")("histories")(h)("items")
                        If JSON("issues")(i)("changelog")("histories")(h)("items")(c)("field") = "timespent" Then
                            Set rng_author = ws_LeadTimeData.Rows(1).Find(JSON("issues")(i)("changelog")("histories")(h)("author")("key"), LookIn:=xlValues, LookAt:=xlWhole)
                            If rng_author Is Nothing Then
                                ws_LeadTimeData.Range("A1").End(xlToRight).Offset(0, 1).value = JSON("issues")(i)("changelog")("histories")(h)("author")("key")
                            End If
                            'set the new value for the story to be the old value for the story + the new value for the sub-task - the old value for the sub task
                            If Not JSON("issues")(i)("changelog")("histories")(h)("items")(c)("fromString") = "" Then
                                collIssueKey(rng_Parent.value)(JSON("issues")(i)("changelog")("histories")(h)("author")("key")) = _
                                    collIssueKey(rng_Parent.value)(JSON("issues")(i)("changelog")("histories")(h)("author")("key")) _
                                    + Val(JSON("issues")(i)("changelog")("histories")(h)("items")(c)("toString")) _
                                    - Val(JSON("issues")(i)("changelog")("histories")(h)("items")(c)("fromString"))
                            Else
                                collIssueKey(rng_Parent.value)(JSON("issues")(i)("changelog")("histories")(h)("author")("key")) = _
                                    collIssueKey(rng_Parent.value)(JSON("issues")(i)("changelog")("histories")(h)("author")("key")) _
                                    + Val(JSON("issues")(i)("changelog")("histories")(h)("items")(c)("toString"))
                            End If
                        End If
                        c = c + 1
                    Next
                    h = h + 1
                Next
            End With
            i = i + 1 'increment the issue
        Next
        '' Totals
        For Each rng_author In ws_LeadTimeData.Range(Cells(1, 9), Cells(1, ws_LeadTimeData.Range("A1").End(xlToRight).Column))
            ws_LeadTimeData.Cells(rng_Parent.Row, rng_author.Column) = collIssueKey(rng_Parent.value)(rng_author.value)
        Next rng_author
        
        col = ws_LeadTimeData.Range("A1").End(xlToRight).Column
        rng_Parent.Offset(0, 5).value = Application.WorksheetFunction.Sum(Range(Cells(rng_Parent.Row, 9), Cells(rng_Parent.Row, col)))
        rng_Parent.Offset(0, 6).value = PublicVariables.jiratime(rng_Parent.Offset(0, 5).value)
    Next rng_Parent
End If

End Function
Private Function funcGetIncompleteJiras(ByVal auth As String, ByVal baseUrl As String, _
            ByVal boardJql As String, ByRef startAtVal, r As Integer) As Long

''
' Source Jiras that are not in a done state and not subTasks

'
' @param {String} auth, baseUrl, boardJql
' @param {Integer} startAtVal, r
' @write {ws_IncompleteIssuesData}
' @apicalls 1x{get search standardissuetypes}
' @return {Long} status of apicall
''

'' Known limitations with this macro:
' (1) needs to be updated to run for a smaller number of maxresults
' (2) worksheet is not scrubbed (clear contents) the first time the macro is called
' (3) worksheet has to exist (with headings) for the macro to run

Dim apicall As String
Dim jql As String
Dim http, JSON, Item, sprint As Object
Dim CleanJSON As String
Dim i, s As Integer

jql = "statusCategory not in (Done) AND " & _
        "issuetype not in subTaskIssueTypes() AND " & _
        boardJql

apicall = _
    baseUrl & "rest/api/latest/search?jql=" _
    & jql _
    & "&fields=" _
        & "key," _
        & "issuetype," _
        & "project," _
        & "status," _
        & epiclink & "," _
        & storypoints & "," _
        & "aggregatetimeestimate," _
        & sprints & "," _
        & "&startAt=" _
        & 0 _
    & "&maxResults=1000"
            
Debug.Print apicall
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

funcGetIncompleteJiras = http.Status

If http.Status = 200 Then
    CleanJSON = CleanSprintsAndCustomFields(http.responseText)
    Set JSON = ParseJson(CleanJSON)
    startAtVal = startAtVal + 1000 'Increment the next start position based on maxResults above
    i = 1 'reset the issue to 1
    For Each Item In JSON("issues")
        With ws_IncompleteIssuesData
            .Cells(r, 1).value = JSON("issues")(i)("id")
            .Cells(r, 2).value = JSON("issues")(i)("key")
            .Cells(r, 3).value = JSON("issues")(i)("fields")("issuetype")("name")
            .Cells(r, 4).value = JSON("issues")(i)("fields")("project")("key")
            .Cells(r, 5).value = JSON("issues")(i)("fields")("epiclink")
            .Cells(r, 6).value = JSON("issues")(i)("fields")("storypoints")
            .Cells(r, 7).value = JSON("issues")(i)("fields")("status")("name")
            .Cells(r, 8).value = JSON("issues")(i)("fields")("status")("statusCategory")("name")
            .Cells(r, 9).value = JSON("issues")(i)("fields")("aggregatetimeestimate")
            If JSON("issues")(i)("fields")("sprints").Count > 0 Then
                s = JSON("issues")(i)("fields")("sprints").Count
                .Cells(r, 10).value = sprint_ParseString(JSON("issues")(i)("fields")("sprints")(s), "state") 'Find the last sprint's state
            Else
                .Cells(r, 10).value = "BACKLOG"
            End If
        End With
        i = i + 1 'increment the issue
        r = r + 1 'increment the row
    Next Item
End If

End Function
Private Function funcGetVelocity(ByVal auth As String, ByVal baseUrl As String, _
                            ByVal rapidViewId As String) As Long

''
' Source the last velocity data from the team's last seven sprints

'
' @param {String} auth
' @param {String} baseUrl
' @param {String} rapidViewId

' @write {ws_VelocityData}
' @apicalls 1x{get Velocity report}
' @return {Long} status of apicall
''

'' Known limitations with this macro:
' (1) worksheet is not scrubbed (clear contents) as it is assumed that you always get seven sprints and just overwrite previous data

Dim http, JSON, Item As Object
Dim apicall As String
Dim r, s As Integer

apicall = baseUrl & "rest/greenhopper/latest/rapid/charts/velocity.json?rapidViewId=" & rapidViewId
             
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

funcGetVelocity = http.Status

If http.Status = 200 Then
    Debug.Print apicall
    Set JSON = ParseJson(http.responseText)
    'Debug.Print http.responseText
    r = 2
    s = 1
    For Each Item In JSON("sprints")
        With ws_VelocityData
            .Cells(r, 1).value = JSON("sprints")(s)("id") ' SprintId
            .Cells(r, 2).value = JSON("sprints")(s)("name") 'SprintName
            .Cells(r, 3).value = JSON("sprints")(s)("state") 'SprintState
            .Cells(r, 4).value = JSON("velocityStatEntries")(CStr(.Cells(r, 1).value))("estimated")("value") 'Commitment
            .Cells(r, 5).value = JSON("velocityStatEntries")(CStr(.Cells(r, 1).value))("completed")("value") 'Completed
        End With
        r = r + 1
        s = s + 1
    Next
End If

Debug.Print (strDebug("GetVelocity", "rapidViewId: " & rapidViewId, http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function
Private Function funcPostTeamsFind(ByVal auth As String, ByVal baseUrl As String) As Long

''
' Source the Teams from Portfolio for Jira

'
' @param {String} auth, baseUrl
' @write {ws_TeamsData}
' @apicalls 1x{post teams find}
' @return {Long} status of apicall
''

'' Known limitations with this macro:
' (1) is hardcoded to a maximum of 50 teams in JsonPost

Dim http, JSON, Team, Resource, Person As Object
Dim t, p, r, l As Integer
Dim JsonPost As String
Dim apicall As String

JsonPost = "{" & Chr(34) & "maxResults" & Chr(34) & ":50}"
apicall = baseUrl & "rest/teams/1.0/teams/find"

Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "POST", apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send (JsonPost)

PostTeamsFind = http.Status

If http.Status = 200 Then
'    Debug.Print apicall
'    Debug.Print JsonPost
    Set JSON = ParseJson(http.responseText)
    t = 1 'reset the teams to 1
    r = 2
    With ws_TeamsData
        .Activate
        .Range(Cells(2, 1), Cells(.Range("A1048576").End(xlUp).Row, 6)).ClearContents ' clear existing data
        For Each Team In JSON("teams")
            p = 1
            For Each Resource In JSON("teams")(t)("resources")
                .Cells(r, 1).value = JSON("teams")(t)("id")
                .Cells(r, 2).value = JSON("teams")(t)("title")
                .Cells(r, 3).value = JSON("teams")(t)("resources")(p)("id")
                .Cells(r, 4).value = JSON("teams")(t)("resources")(p)("personId")
                p = p + 1
                r = r + 1
            Next Resource
            t = t + 1
        Next Team
        r = 2
        p = 1
        For Each Person In JSON("persons")
            .Cells(r, 5).value = JSON("persons")(p)("personId")
            .Cells(r, 6).value = JSON("persons")(p)("jiraUser")("jiraUsername")
            p = p + 1
            r = r + 1
        Next Person
    End With
End If

Debug.Print (strDebug("PostTeamsFind", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Private Function funcGetSprintBurnDown(ByVal auth As String, ByVal baseUrl As String, _
                            ByVal rapidViewId As String, ByVal SprintId As String) As Long

'' This function records a log of time spent against each issue during a sprint (from taken from the Sprint BurnDown Chart)
'' It also updates the RemainingSprintTime public variable

' @param {String} auth
' @param {String} baseUrl
' @param {String} rapidViewId
' @param {String} TeamId
' @param {String} SprintId

' @write {ws_Work}
' @return {Long} status of apicall
''

Dim http, JSON, time, change, rates As Object
Dim apicall As String
Dim c%, r%, d As Integer

apicall = baseUrl & "rest/greenhopper/1.0/rapid/charts/scopechangeburndownchart?rapidViewId=" _
        & rapidViewId & "&sprintId=" & SprintId
             
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

funcGetSprintBurnDown = http.Status

Debug.Print (apicall)

If http.Status = 200 Then
    
    With ws_Work
        .Cells.ClearContents
        .Cells(1, 1).value = "Key"
        .Cells(1, 2).value = "rapidViewId"
        .Cells(1, 3).value = "timeSpent"
    End With

    Set JSON = ParseJson(http.responseText)
    
    RemainingSprintTime = 0
    DaysInSprint = 0
    
    For Each time In JSON("changes")
        c = 1
        For Each change In JSON("changes")(time)
            If JSON("changes")(time)(c).Exists("timeC") Then
                'If JSON("changes")(time)(c)("timeC")("changeDate") < JSON("endTime") Then
                '' To align with the Sprint Burndown chart I am changing this to take the effective date of the work log rather than the changedate _
                which is when the worklog was created, thus allowing users to backvalue burndown data and comparing this to the time the sprint actually _
                ended in Jira rather than the time it was expected to end. Updated if condition below:
                If Val(time) < JSON("completeTime") Then
                    RemainingSprintTime = RemainingSprintTime + _
                        (JSON("changes")(time)(c)("timeC")("newEstimate") - JSON("changes")(time)(c)("timeC")("oldEstimate"))
                    '' This next statement records a log of the issues that have had work logged to them during the sprint
                    If Val(time) > JSON("startTime") Then
                        If JSON("changes")(time)(c)("timeC").Exists("timeSpent") Then
                            With ws_Work.Range("A1048576").End(xlUp)
                                .Offset(1).value = JSON("changes")(time)(c)("key")
                                .Offset(1, 1).value = rapidViewId
                                .Offset(1, 2).value = JSON("changes")(time)(c)("timeC")("timeSpent")
                            End With
                        End If
                    End If
                End If
            End If
            c = c + 1
        Next change
    Next time
End If

r = 1
For Each rates In JSON("workRateData")("rates")
    If JSON("workRateData")("rates")(r)("rate") = 1 Then
        DaysInSprint = DaysInSprint + JSON("workRateData")("rates")(r)("end") - JSON("workRateData")("rates")(r)("start")
    End If
    r = r + 1
Next rates

DaysInSprint = DaysInSprint / 86400 / 1000

Debug.Print (strDebug("GetSprintBurnDown", "rapidViewId: " & rapidViewId & " | SprintId = " & SprintId & " | RemainingSprintTime = " & CStr(RemainingSprintTime) & " | Days In Sprint = " & CStr(DaysInSprint), http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Function funcPredictabilitySprintsEstimated()

'' Update the TeamStats worksheet with the *Sprints Estimated* calcualtion
' The value for stories is added both to the sparkline graph
' Then the value for subtasks, stories and epics is added to the board for display
'
' Dependent on function: funcGetIncompleteJiras & funcGetVelocity
'
''
Dim SubTaskEstimate&, StoryPointEstimate&, TShirtEstimate As Long

'' To be removed and added as an input variable
DaysInSprint = 9

'SubTaskEstimate calculated as _
    = aggregateTimeEstimate from backlog / teamsize / working DaysInSprint / working hours in day / seconds in hour
SubTaskEstimate = Excel.WorksheetFunction.Sum(ws_IncompleteIssuesData.Range("I:I")) _
                    / Excel.WorksheetFunction.CountIf(ws_TeamsData.Range("A:A"), CInt(TeamId)) _
                    / DaysInSprint _
                    / 8 _
                    / 3600
                    
'StoryPointEstimate calculated as _
    = Total StoryPoints from backlog (excluding Epics and stories greater than 20) / Average Velocity from last 7 sprints
StoryPointEstimate = Excel.WorksheetFunction.SumIfs(ws_IncompleteIssuesData.Range("F:F"), ws_IncompleteIssuesData.Range("F:F"), "<20", ws_IncompleteIssuesData.Range("C:C"), "<>Epic") _
                    / Excel.WorksheetFunction.Average(ws_VelocityData.Range("E:E"))

''Note: TShirtEstimate is taken to be all Stories with a size of 20 or more. We are not taking Epic estimates into account

'TShirtEstimate calculated as _
    = Total StoryPoints from backlog (excluding Epics and stories less than 20) / Average Velocity from last 7 sprints
TShirtEstimate = Excel.WorksheetFunction.SumIfs(ws_IncompleteIssuesData.Range("F:F"), ws_IncompleteIssuesData.Range("F:F"), ">=20", ws_IncompleteIssuesData.Range("C:C"), "<>Epic") _
                    / Excel.WorksheetFunction.Average(ws_VelocityData.Range("E:E"))

With ws_TeamStats
    .Range("BB25").value = Round(StoryPointEstimate, 0)
    .Range("J10").value = CStr(Round(SubTaskEstimate, 0)) & "/" & CStr(Round(StoryPointEstimate, 0)) & "/" & CStr(Round(TShirtEstimate, 0))
End With

End Function
Function funcPredictabilityVelocity()

'' Update the TeamStats worksheet with the *Sprint Velocity* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetVelocity
''

'Sprint Velocity calculated as _
    = Completed / Committed for the most recent sprint
With ws_TeamStats
    .Range("BB26").value = ws_VelocityData.Range("E2").value / ws_VelocityData.Range("D2").value
    .Range("S10").value = ws_VelocityData.Range("E2").value / ws_VelocityData.Range("D2").value
End With

End Function
Function funcPredictabilityTiPVariability()

'' Update the TeamStats worksheet with the *TiP Variability* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras & funcResponsivenessTiP
''

'' Known limitations with this macro:
' (1) ws_LeadTimeData sheet needs to be made active

Dim rng As Range
Dim TiPCol As Integer
Dim TipRow As Long

TiPCol = ws_LeadTimeData.Range("1:1").Find("TiP").Column
TipRow = ws_LeadTimeData.Cells(1, TiPCol).End(xlDown).Row

Set rng = ws_LeadTimeData.Range(Cells(2, TiPCol), Cells(TipRow, TiPCol))

With ws_TeamStats
    .Range("BB27").value = Excel.WorksheetFunction.StDev_P(rng) / Excel.WorksheetFunction.Average(rng)
    .Range("J16").value = Excel.WorksheetFunction.StDev_P(rng) / Excel.WorksheetFunction.Average(rng)
End With

End Function
Function funcPredictabilitySprintOutputVariability()

'' Update the TeamStats worksheet with the *Sprint Output Variability* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetVelocity
'
''

With ws_TeamStats
    .Range("BB28").value = Excel.WorksheetFunction.StDev_P(ws_VelocityData.Range("E2:E8")) / Excel.WorksheetFunction.Average(ws_VelocityData.Range("E2:E8"))
    .Range("S16").value = Excel.WorksheetFunction.StDev_P(ws_VelocityData.Range("E2:E8")) / Excel.WorksheetFunction.Average(ws_VelocityData.Range("E2:E8"))
End With

End Function
Function funcResponsivenessLeadTime()

'' Update the TeamStats worksheet with the *Lead Time* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''

Dim col As Integer
Dim rng_LeadTime As Range
Dim c As Range

With ws_LeadTimeData

    col = .Range("A1").End(xlToRight).Column + 1
    .Cells(1, col).value = "leadTime"

    Set rng_LeadTime = .Range(Cells(2, col), Cells(.Range("A1").End(xlDown).Row, col))

    For Each c In rng_LeadTime
        c.value = CDate(.Cells(c.Row, 6).value) - CDate(Left(.Cells(c.Row, 4).value, 10))
    Next c
    
End With


With ws_TeamStats
    .Range("BB30").value = Excel.WorksheetFunction.Average(rng_LeadTime)
    .Range("AD10").value = Round(Excel.WorksheetFunction.Average(rng_LeadTime), 0)
End With

End Function
Function funcResponsivenessDeploymentFrequency()

'' Update the TeamStats worksheet with the *Deployment Frequency* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''

'Deployment Frequency calculated as _
    = Number of unique releaseDates in the previous month from today

Dim dict As Dictionary
Dim cell As Range

Set dict = New Dictionary
        
For Each cell In ws_LeadTimeData.Range(Cells(2, 6), Cells(ws_LeadTimeData.Range("F1").End(xlDown).Row, 6))
    If cell.value >= DateAdd("m", -1, "01/" & Month(Now()) & "/" & Year(Now())) Then ' only count if after start of previous month
        If cell.value < CDate("01/" & Month(Now()) & "/" & Year(Now())) Then ' only count if before start of current month
            If Not dict.Exists(cell.value) Then
                dict.Add cell.value, 0
            End If
        End If
    End If
Next

With ws_TeamStats
    .Range("BB31").value = dict.Count
    .Range("AM10").value = dict.Count
End With

End Function
Function funcResponsivenessTiP()

'' Update the TeamStats worksheet with the *TiP* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''

Dim col As Integer
Dim rng_TiP As Range
Dim c As Range

With ws_LeadTimeData

    col = .Range("A1").End(xlToRight).Column + 1
    .Cells(1, col).value = "TiP"

    Set rng_TiP = .Range(Cells(2, col), Cells(.Range("A1").End(xlDown).Row, col))

    For Each c In rng_TiP
        c.value = CDate(.Cells(c.Row, 6).value) - CDate(Left(.Cells(c.Row, 5).value, 10))
    Next c
    
End With


With ws_TeamStats
    .Range("BB32").value = Excel.WorksheetFunction.Average(rng_TiP)
    .Range("AD16").value = Round(Excel.WorksheetFunction.Average(rng_TiP), 0)
End With

End Function
Function funcResponsivenessWiP()

'' Update the TeamStats worksheet with the *WiP* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''

With ws_TeamStats
    .Range("BB33").value = 0 ' Forumla to be updated
    .Range("AM16").value = 0 ' Forumla to be updated
End With

End Function
Function funcProductivityReleaseVelocity()

'' Update the TeamStats worksheet with the *Release Velocity* data
'
' Dependent on function: funcGetDoneJiras
'
''

With ws_TeamStats
    .Range("AT5").value = 0 ' Forumla to be updated - Feature Velocity
    .Range("AU5").value = 0 ' Forumla to be updated - Defects Velocity
    .Range("AV5").value = 0 ' Forumla to be updated - Risks Velocity
    .Range("AW5").value = 0 ' Forumla to be updated - Debts Velocity
    .Range("AX5").value = 0 ' Forumla to be updated - Enablers Velocity
    
    .Range("AT6").value = 0 ' Forumla to be updated - Feature Baseline
    .Range("AU6").value = 0 ' Forumla to be updated - Defects Baseline
    .Range("AV6").value = 0 ' Forumla to be updated - Risks Baseline
    .Range("AW6").value = 0 ' Forumla to be updated - Debts Baseline
    .Range("AX6").value = 0 ' Forumla to be updated - Enablers Baseline
End With

End Function
Function funcProductivityEfficiency()

'' Update the TeamStats worksheet with the *Efficiency* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''

With ws_TeamStats
    .Range("BB35").value = 0 ' Forumla to be updated
    .Range("AB24").value = 0 ' Forumla to be updated
End With

End Function
Function funcProductivityDistribution()

'' Update the TeamStats worksheet with the *Distribution* data
'
' Dependent on function: funcGetDoneJiras & funcGetIncompleteJiras
'
''

''' Should not rely on funcGetDoneJiras as this is only last 3 months of data. Need to go back 12 months so new api call


With ws_TeamStats
    .Range("AT11:BF15").value = 0 ' Forumla to be updated - Data
    .Range("AT16:BE16").value = "Date" ' Formula to be updated - Release Months
End With

End Function
Function funcQualityTimeToResolve()

'' Update the TeamStats worksheet with the *Time To Resolve* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''

With ws_TeamStats
    .Range("BB37").value = 0 ' Forumla to be updated
    .Range("AM24").value = 0 ' Forumla to be updated
End With

End Function
Function funcQualityDefectDentisy()

'' Update the TeamStats worksheet with the *Defect Density* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: **AffectsVersionJQL**
'
''

With ws_TeamStats
    .Range("BB38").value = "TBC" ' Forumla to be updated
    .Range("AM30").value = "TBC" ' Forumla to be updated
End With

End Function
Function funcQualityFailRate()

'' Update the TeamStats worksheet with the *Fail Rate* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: **AffectsVersionJQL**
'
''

With ws_TeamStats
    .Range("BB39").value = "TBC" ' Forumla to be updated
    .Range("AM36").value = "TBC" ' Forumla to be updated
End With

End Function
Function funcScrumTeamStability()

'' Update the TeamStats worksheet with the Team Stability
' The value is added both to the sparkline graph
' Then to the board for display
'
' Dependent on function: funcPostTeamsFind
'
''

With ws_TeamStats
    .Range("BB22").value = 0 ' Forumla to be updated
    .Range("J5").value = 0 ' Forumla to be updated
End With

End Function
Function funcScrumUnplannedWork()

'' Update the TeamStats worksheet with the *Uplanned Work* (RemainingSprintTime)
' The the value in milliseconds is used for the sparkline graph
' This is then converted into a string for display in weeks, days, hours on the board
'
' Dependent on function: funcGetSprintBurnDown
'
''

With ws_TeamStats
    .Range("BB23").value = RemainingSprintTime
    .Range("AD5").value = TeamStats.jiratime(RemainingSprintTime)
End With

End Function
Function funcJiraAdminCorrectStatus()

'' Update the TeamStats worksheet with the *Correct Status* calcualtion
' The value is added both to the sparkline graph
' Then to the board for display
'
' Dependent on function: funcGetIncompleteJiras
'
''

With ws_TeamStats
    .Range("BB41").value = 0 ' Forumla to be updated
    .Range("J43").value = 0 ' Forumla to be updated
End With

End Function
Function funcJiraAdminCorrectEpicLink()

'' Update the TeamStats worksheet with the *Correct Project & Epic Link* calcualtion
' The value is added both to the sparkline graph
' Then to the board for display
'
' Dependent on function: funcGetIncompleteJiras
'
''

With ws_TeamStats
    .Range("BB42").value = 0 ' Forumla to be updated
    .Range("T43").value = 0 ' Forumla to be updated
End With

End Function
Function funcJiraAdminDoneInSprint()

'' Update the TeamStats worksheet with the *Done In Sprint* calcualtion
' The value is added both to the sparkline graph
' Then to the board for display
'
' Dependent on function: **NewSprintReport**
'
''

With ws_TeamStats
    .Range("BB43").value = "TBC" ' Forumla to be updated
    .Range("AC43").value = "TBC" ' Forumla to be updated
End With

End Function
Function funcJiraAdminActiveTime()

'' Update the TeamStats worksheet with the *Active Time %* calcualtion
' The value is added both to the sparkline graph
' Then to the board for display
'
' Dependent on function: funcGetSprintBurnDown & funcPostTeamsFind
'
''

With ws_TeamStats
    .Range("BB41").value = 0 ' Forumla to be updated
    .Range("J43").value = 0 ' Forumla to be updated
End With

End Function


'' The following Functions are used by the apicalls above



Private Function CleanSprintsAndCustomFields(FullJasonStr As String)
' This function replaces certain aspects of the default Jira Json result so that
' it is easier to work with

Dim str As String
Dim oIssue As Object
Dim oIssueSprints As Object
Dim oRE As Object
Dim s, e As Long

Set oRE = CreateObject("VBScript.RegExp")
' replace the default sprint pattern
With oRE
    .Global = True
    .Pattern = "com.atlassian.greenhopper.service.sprint.Sprint@[0-9a-zA-Z]+\["
    FullJasonStr = .Replace(FullJasonStr, "[")
End With

'rename the custom fields usiing the Public Constants in the PublicVariables module
FullJasonStr = Replace(FullJasonStr, sprints, "sprints")
FullJasonStr = Replace(FullJasonStr, parentlink, "parentlink")
FullJasonStr = Replace(FullJasonStr, epiclink, "epiclink")
FullJasonStr = Replace(FullJasonStr, Team, "team")
FullJasonStr = Replace(FullJasonStr, storypoints, "storypoints")
'ensure that the sprints object is passed as an array even when there are no sprints
FullJasonStr = Replace(FullJasonStr, """sprints""" & ":null", """sprints""" & ":[]")

CleanSprintsAndCustomFields = FullJasonStr

Set oRE = Nothing

End Function

Private Function sprint_ParseString(ByVal sprint_String As String, sprint_Field As String) As String

'This function parses out the sprint fields which are stored as a long comma seperate string within an array
    
    Dim StartPos, EndPos As Long
    
    StartPos = InStr(1, sprint_String, sprint_Field) + Len(sprint_Field) + 1
        Select Case sprint_Field
            Case "id"
                EndPos = InStr(1, sprint_String, "rapidViewId") - 1
            Case "rapidViewId"
                EndPos = InStr(1, sprint_String, "state") - 1
            Case "state"
                EndPos = InStr(1, sprint_String, "name") - 1
            Case "name"
                EndPos = InStr(1, sprint_String, "startDate") - 1
            Case "startDate"
                EndPos = InStr(1, sprint_String, "endDate") - 1
            Case "endDate"
                EndPos = InStr(1, sprint_String, "completeDate") - 1
            Case "completeDate"
                EndPos = InStr(1, sprint_String, "sequence") - 1
            Case "sequence"
                EndPos = InStr(1, sprint_String, "goal") - 1
            Case "goal"
                EndPos = Len(sprint_String) - 1
            Case Else
                sprint_ParseString = ""
        End Select
    
    sprint_ParseString = Mid(sprint_String, StartPos, EndPos - StartPos)

End Function
Private Function strDebug(ByVal api As String, s As String, r As Integer, userName As String) As String

'' This function records a log of all the api calls to a Log tab and returns a string for use with debugging

' @param {String} api = name of the api call
' @param {String} s = some description you want included in the debug text
' @param {Integer} r = is the response status code i.e. 200
' @param {String} userName = user that was authenticated for the api call

' @write {ws_Log}
' @return {String} string i.e. to be printed to the Immediate window
''

    Dim strLog As String
    Dim c As Range
    
    strLog = Now() & ": " & api & ": " & userName & " | " & s & " | - " & r
    Set c = ws_Log.Cells(ws_ProjectData.Rows.Count, "A").End(xlUp)
    c.Offset(1, 0).value = strLog
    strDebug = strLog
    
End Function

Private Function jiratime(ByVal timeseries As Double) As String

'' This function converts a time in milliseconds (as used by Jira) into a string value of weeks, days and hours

Dim h As Double
Dim w As Double
Dim d As Double

h = WorksheetFunction.RoundDown(timeseries / 3600, 0)

Select Case h
    Case Is >= 40
        w = WorksheetFunction.RoundDown(h / 40, 0)
        d = WorksheetFunction.RoundDown((h Mod 40) / 8, 0)
        h = h Mod 40 Mod 8
        If d > 0 Then
            If h > 0 Then
                jiratime = w & "w " & d & "d " & h & "h"
            Else
                jiratime = w & "w " & d & "d "
            End If
        Else
            jiratime = w & "w " & h & "h"
        End If
    Case Is >= 8
        d = WorksheetFunction.RoundDown(h / 8, 0)
        h = h Mod 8
        If h > 0 Then
            jiratime = d & "d " & h & "h"
        Else
            jiratime = d & "d "
        End If
    Case Else
        jiratime = h & "h"
        
End Select
End Function

Private Function funcRAG()

'' This function should update the RAG triangles for each of the values displayed on the Dashboard

End Function

Private Function funcAsOfDateTeamName()

'' This function should update the AsOfDate and Team Name displayed on the Dashboard

End Function
