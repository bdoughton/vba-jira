Attribute VB_Name = "JiraScrumTeamStats"
''
' VBA-JiraScrumTeamStats v2
' (c) Ben Doughton - https://github.com/bdoughton/vba-jira
'
' JIRA Scrum Team Stats VBA
'
' @Dependencies:
'               Mod - WebHelpers
'               Mod - Jira
'               Class - JiraResponse
'               Class - WebRequest
'               Class - WebResponse
'               Class - WebClient
'               Class - Dictionary
'
' Note: This is designed to be a standalone module for TeamStats so if there are other modules
'       from the same family of Jira apicalls there could be duplication of code
'
' @module JiraScrumTeamStats
' @author bdoughton@me.com
' @license GNU General Public License v3.0 (https://opensource.org/licenses/GPL-3.0)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
 
Option Explicit
Option Base 1
Public RemainingSprintTime As Long
Public DaysInSprint As Long
Public LastSprintName As String
Public LastSprintId As Integer

Sub GetTeamStats(control As IRibbonControl)
 
''
' This should be run by the user and sets up all the underlying api calls to get the teams stats
 
'' Known limitations with this macro:
' (1) Work in progress - run each call individually by commenting out the other api calls
 
'Pause calculations and screen updating and make read-only worksheets visible
'These actions are reversed at the end of the macro
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
 
    ' --- Comment out the respective value to enable or suspend logging
    WebHelpers.EnableLogging = True
'    WebHelpers.EnableLogging = False
   
    'Check if a user is logged in and if not perform login, if login fails exit
    If Not IsLoggedIn Then
        If Not LoginUser Then Exit Sub
    End If
   
    ''Rollstats True of False
    'Dim blnRoll As Boolean
    'If MsgBox("Do you want to roll the previous data?", vbYesNo) = vbYes Then
    '    blnRoll = True
    'Else
    '    blnRoll = False
    'End If
    'funcRollStats (blnRoll)
 
    
    ''Fetch Data from Api
    Dim callResult(1 To 5) As WebStatusCode
    callResult(1) = funcGet3MonthsOfDoneJiras(boardJql, "In Progress", "Done", 0, 2)
    callResult(2) = funcGetIncompleteJiras(boardJql, 0, 2)
    callResult(3) = funcGetVelocity(rapidViewId)
    callResult(4) = funcPostTeamsFind()
'
'    ''Need to save the RemainingSprintTime to the right cell
    callResult(5) = funcGetSprintBurnDown(rapidViewId, CStr(ws_VelocityData.Range("A2").Value))
 
 
'Reverse the opening statements that paused calculations and screen updating
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
        .Range("AW22:BA44").Value = .Range("AX22:BB44").Value
    End With
End If
 
End Function
 
Private Function funcGet3MonthsOfDoneJiras(ByVal boardJql As String, ByVal inProgressState As String, ByVal endProgressState As String, _
            ByRef startAtVal, r As Integer) As WebStatusCode
 
''
' Source Jiras that are in a Done state and were included in a sprint and had a fixVersion that was updated in the last 24 weeks
' Then cycle through and get the sub-tasks for all the issues from the first api call
 
'
' @param {String} boardJql, inProgressState, endProgressState
' @param {Integer} startAtVal, r
' @write {ws_LeadTimeData} & {ws_WiPData}
' @apicalls 1x{get search standardissuetypes} ?x{get search subtaskissuetypes}
' @return {WebStatusCode} status of first apicall
''
 
'' Known limitations with this macro:
' (1) Can't handle different inProgressState and endProgressState by issue type
' (2) needs to be updated to run for a smaller number of maxresults
' (3) there is no error handling around the second api call to get the subtasks which could fail
 
Dim jql As String
jql = "fixversion changed after -24w AND " & _
        "fixVersion is not EMPTY AND " & _
        "Sprint is not EMPTY AND " & _
        "NOT issuetype in (Theme,Initiative,Epic,Test,subTaskIssueTypes()) AND " & _
        "statusCategory in (Done) AND " & _
        boardJql
 
Dim apiFields As String
apiFields = "key," _
        & "issuetype," _
        & "fixVersions," _
        & "resolutiondate," _
        & sprints & "," _
        & "created," _
        & "changelog"
 
'Define the new JQLRequest
Dim JQL_PBI_Request As New WebRequest
With JQL_PBI_Request
    .Resource = "api/2/search"
    .Method = WebMethod.HttpGet
    .AddQuerystringParam "jql", jql
    .AddQuerystringParam "fields", apiFields
    .AddQuerystringParam "startAt", startAtVal
    .AddQuerystringParam "maxResults", "1000"
    .AddQuerystringParam "expand", "changelog"
End With
           
Dim JQL_PBI_Search_Response As New JiraResponse
Dim JQL_Search_Response As New WebResponse
Dim Item As Object
Dim history As Object
Dim changeitem As Object
Dim fixversion As Object
Dim i As Integer
Dim h As Integer
Dim c As Integer
Dim rng_author As Range
Dim WiPRow As Integer
Dim rng_Parent As Range
Dim col As Integer
Dim dictResourceNm As Dictionary
Dim dictTimeLoggedToStory As Dictionary
Dim collIssueKey As New Collection
 
Set JQL_Search_Response = JQL_PBI_Search_Response.JiraCall(JQL_PBI_Request)
 
funcGet3MonthsOfDoneJiras = JQL_Search_Response.StatusCode
 
If funcGet3MonthsOfDoneJiras = Ok Then
    clearOldData ws_LeadTimeData
    clearOldData ws_WiPData
    Set dictTimeLoggedToStory = New Dictionary
    startAtVal = startAtVal + 1000 'Increment the next start position based on maxResults above -- making this smaller will speed up the API calls
    i = 1 'reset the issue to 1
    WiPRow = r
    For Each Item In JQL_Search_Response.Data("issues")
        Set dictResourceNm = New Dictionary
        h = 1 'reset the change history to 1
        If CDate(JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")(1)("releaseDate")) >= DateAdd("m", -3, "01/" & Month(Now()) & "/" & Year(Now())) Then 'Only include if the release date was in the last 3 months
            With ws_LeadTimeData
                .Cells(r, 1).Value = JQL_Search_Response.Data("issues")(i)("id")
                .Cells(r, 2).Value = JQL_Search_Response.Data("issues")(i)("key")
                .Cells(r, 3).Value = JQL_Search_Response.Data("issues")(i)("fields")("issuetype")("name")
                .Cells(r, 4).Value = JQL_Search_Response.Data("issues")(i)("fields")("created")
                .Cells(r, 5).Value = sprint_ParseString(JQL_Search_Response.Data("issues")(i)("fields")(sprints)(1), "startDate")
                If JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")(1)("releaseDate") > JQL_Search_Response.Data("issues")(i)("fields")("created") Then
                    .Cells(r, 6).Value = JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")(1)("releaseDate") 'Always use the 1st fixVersion, even if there are multiple
                Else
                    .Cells(r, 6).Value = Left(JQL_Search_Response.Data("issues")(i)("fields")("resolutiondate"), 10) 'use the resolution date if there is no fixVersion. Note: this can lead to incorrect deployment frequency
                   
                End If
                For Each history In JQL_Search_Response.Data("issues")(i)("changelog")("histories")
                    c = 1 'reset the change item to 1
                    For Each changeitem In JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")
                        If JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("field") = "status" Then
                            Select Case JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("toString")
                                Case inProgressState 'enter the date the issue transitioned to its inProgressState
                                    ws_WiPData.Cells(WiPRow, 1).Value = JQL_Search_Response.Data("issues")(i)("id")
                                    ws_WiPData.Cells(WiPRow, 2).Value = JQL_Search_Response.Data("issues")(i)("key")
                                    ws_WiPData.Cells(WiPRow, 3).Value = JQL_Search_Response.Data("issues")(i)("fields")("issuetype")("name")
                                    For Each fixversion In JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")
                                        ws_WiPData.Cells(WiPRow, 6).Value = JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")(1)("releaseDate")  'Always use the 1st fixVersion, even if there are multiple
                                    Next
                                    ws_WiPData.Cells(WiPRow, 4).Value = JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("created")
                                Case endProgressState 'enter the date the issue transitioned to its endProgressState
                                    ws_WiPData.Cells(WiPRow, 5).Value = JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("created")
                                    WiPRow = WiPRow + 1
                            End Select
                        ElseIf JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("field") = "timespent" Then
                            dictResourceNm(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key")) = Val(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("toString"))
                            Set rng_author = ws_LeadTimeData.Rows(1).Find(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key"), LookIn:=xlValues, LookAt:=xlWhole)
                            If rng_author Is Nothing Then
                                ws_LeadTimeData.Range("A1").End(xlToRight).Offset(0, 1).Value = JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key")
                            End If
                        End If
                        c = c + 1
                    Next
                    h = h + 1
                Next
            collIssueKey.Add dictResourceNm, JQL_Search_Response.Data("issues")(i)("key")
            End With
            r = r + 1 'increment the row
        End If
        i = i + 1 'increment the issue
    Next
   
    '' This next section cycles through all the sub-tasks and adds up the time logged to each
 
    For Each rng_Parent In ws_LeadTimeData.Range("B2:B" & ws_LeadTimeData.Range("A1").End(xlDown).row)
        jql = "Parent = " & rng_Parent.Value
   
        apiFields = "key," _
            & "issuetype," _
            & "changelog"
 
        Dim JQL_SubTask_Request As New WebRequest
        With JQL_SubTask_Request
            .Resource = "api/2/search"
            .Method = WebMethod.HttpGet
            .AddQuerystringParam "jql", jql
            .AddQuerystringParam "fields", apiFields
            .AddQuerystringParam "startAt", "0"
            .AddQuerystringParam "maxResults", "1000"
            .AddQuerystringParam "expand", "changelog"
        End With
 
        Dim JQL_SubTask_Search_Response As New JiraResponse
        Set JQL_Search_Response = JQL_SubTask_Search_Response.JiraCall(JQL_SubTask_Request)
 
        i = 1 'reset the issue to 1
        For Each Item In JQL_Search_Response.Data("issues")
            h = 1 'reset the change history to 1
            With ws_LeadTimeData
                For Each history In JQL_Search_Response.Data("issues")(i)("changelog")("histories")
                    c = 1 'reset the change item to 1
                    For Each changeitem In JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")
                        If JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("field") = "timespent" Then
                            Set rng_author = ws_LeadTimeData.Rows(1).Find(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key"), LookIn:=xlValues, LookAt:=xlWhole)
                            If rng_author Is Nothing Then
                                ws_LeadTimeData.Range("A1").End(xlToRight).Offset(0, 1).Value = JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key")
                            End If
                            'set the new value for the story to be the old value for the story + the new value for the sub-task - the old value for the sub task
                            If Not JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("fromString") = "" Then
                                collIssueKey(rng_Parent.Value)(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key")) = _
                                    collIssueKey(rng_Parent.Value)(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key")) _
                                    + Val(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("toString")) _
                                    - Val(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("fromString"))
                            Else
                                collIssueKey(rng_Parent.Value)(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key")) = _
                                    collIssueKey(rng_Parent.Value)(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("author")("key")) _
                                    + Val(JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("items")(c)("toString"))
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
        ws_LeadTimeData.Activate
        For Each rng_author In ws_LeadTimeData.Range(Cells(1, 9), Cells(1, ws_LeadTimeData.Range("A1").End(xlToRight).column))
            ws_LeadTimeData.Cells(rng_Parent.row, rng_author.column) = collIssueKey(rng_Parent.Value)(rng_author.Value)
        Next rng_author
       
        col = ws_LeadTimeData.Range("A1").End(xlToRight).column
        rng_Parent.Offset(0, 5).Value = Application.WorksheetFunction.Sum(Range(Cells(rng_Parent.row, 9), Cells(rng_Parent.row, col)))
        rng_Parent.Offset(0, 6).Value = Jira.jiratime(rng_Parent.Offset(0, 5).Value)
        Set JQL_SubTask_Request = Nothing
    Next rng_Parent
End If
 
End Function
Private Function funcGetIncompleteJiras(ByVal boardJql As String, ByRef startAtVal, r As Integer) As WebStatusCode
 
''
' Source Jiras that are not in a done state and not subTasks
 
'
' @param {String} boardJql
' @param {Integer} startAtVal, r
' @write {ws_IncompleteIssuesData}
' @apicalls 1x{get search standardissuetypes}
' @return {WebStatusCode} status of apicall
''
 
'' Known limitations with this macro:
' (1) needs to be updated to run for a smaller number of maxresults
 
Dim jql As String
jql = "statusCategory not in (Done) AND " & _
        "issuetype not in subTaskIssueTypes() AND " & _
        boardJql
       
Dim apiFields As String
apiFields = "key," _
        & "issuetype," _
        & "project," _
        & "status," _
        & epiclink & "," _
        & storypoints & "," _
        & "aggregatetimeestimate," _
        & sprints
           
'Define the new Request
Dim JQL_PBI_Request As New WebRequest
With JQL_PBI_Request
    .Resource = "api/2/search"
    .Method = WebMethod.HttpGet
    .AddQuerystringParam "jql", jql
    .AddQuerystringParam "fields", apiFields
    .AddQuerystringParam "startAt", startAtVal
    .AddQuerystringParam "maxResults", "1000"
    .AddQuerystringParam "expand", "changelog"
End With
           
Dim JQL_PBI_Search_Response As New JiraResponse
Dim JQL_Search_Response As New WebResponse
 
Set JQL_Search_Response = JQL_PBI_Search_Response.JiraCall(JQL_PBI_Request)
 
funcGetIncompleteJiras = JQL_Search_Response.StatusCode
 
Dim i%, s As Integer
Dim Item As Object
 
If funcGetIncompleteJiras = Ok Then
    clearOldData ws_IncompleteIssuesData
    startAtVal = startAtVal + 1000 'Increment the next start position based on maxResults above
    i = 1 'reset the issue to 1
    For Each Item In JQL_Search_Response.Data("issues")
        With ws_IncompleteIssuesData
           .Cells(r, 1).Value = JQL_Search_Response.Data("issues")(i)("id")
            .Cells(r, 2).Value = JQL_Search_Response.Data("issues")(i)("key")
            .Cells(r, 3).Value = JQL_Search_Response.Data("issues")(i)("fields")("issuetype")("name")
            .Cells(r, 4).Value = JQL_Search_Response.Data("issues")(i)("fields")("project")("key")
            .Cells(r, 5).Value = JQL_Search_Response.Data("issues")(i)("fields")(epiclink)
            .Cells(r, 6).Value = JQL_Search_Response.Data("issues")(i)("fields")(storypoints)
            .Cells(r, 7).Value = JQL_Search_Response.Data("issues")(i)("fields")("status")("name")
            .Cells(r, 8).Value = JQL_Search_Response.Data("issues")(i)("fields")("status")("statusCategory")("name")
            .Cells(r, 9).Value = JQL_Search_Response.Data("issues")(i)("fields")("aggregatetimeestimate")
            If VarType(JQL_Search_Response.Data("issues")(i)("fields")(sprints)) = vbObject Then
                s = JQL_Search_Response.Data("issues")(i)("fields")(sprints).Count
                .Cells(r, 10).Value = sprint_ParseString(JQL_Search_Response.Data("issues")(i)("fields")(sprints)(s), "state") 'Find the last sprint's state
            Else
                .Cells(r, 10).Value = "BACKLOG"
            End If
        End With
        i = i + 1 'increment the issue
        r = r + 1 'increment the row
    Next Item
End If
 
End Function
Private Function funcGetVelocity(ByVal rapidViewId As String) As WebStatusCode
 
''
' Source the last velocity data from the team's last seven sprints
'
' @param {String} rapidViewId
' @write {ws_VelocityData}
' @apicalls 1x{get Velocity report}
' @return {WebStatusCode} status of apicall
''
          
'Define the new Request
Dim VelocityChartRequest As New WebRequest
With VelocityChartRequest
    .Resource = "greenhopper/latest/rapid/charts/velocity.json"
    .Method = WebMethod.HttpGet
    .AddQuerystringParam "rapidViewId", rapidViewId
End With
           
Dim VelocityChartResponse As New JiraResponse
Dim VelocityResponse As New WebResponse
 
Set VelocityResponse = VelocityChartResponse.JiraCall(VelocityChartRequest)
 
funcGetVelocity = VelocityResponse.StatusCode
 
Dim Item As Object
Dim r%, s As Integer
 
If funcGetVelocity = Ok Then
    r = 2
    s = 1
    clearOldData ws_VelocityData
    For Each Item In VelocityResponse.Data("sprints")
        With ws_VelocityData
            .Cells(r, 1).Value = VelocityResponse.Data("sprints")(s)("id") ' SprintId
            .Cells(r, 2).Value = VelocityResponse.Data("sprints")(s)("name") 'SprintName
            .Cells(r, 3).Value = VelocityResponse.Data("sprints")(s)("state") 'SprintState
            .Cells(r, 4).Value = VelocityResponse.Data("velocityStatEntries")(CStr(.Cells(r, 1).Value))("estimated")("value") 'Commitment
            .Cells(r, 5).Value = VelocityResponse.Data("velocityStatEntries")(CStr(.Cells(r, 1).Value))("completed")("value") 'Completed
        End With
        r = r + 1
        s = s + 1
    Next
End If
 
End Function
Private Function funcPostTeamsFind() As WebStatusCode
 
''
' Source the Teams from Portfolio for Jira
'
' @write {ws_TeamsData}
' @apicalls 1x{post teams find}
' @return {WebStatusCode} status of apicall
''
 
'' Known limitations with this macro:
' (1) is hardcoded to a maximum of 50 teams in JsonPost
 
'Dim JsonPost As String
'JsonPost = "{" & Chr(34) & "maxResults" & Chr(34) & ":50}"
Dim JiraBody As New Dictionary
JiraBody.Add "maxResults", 50
 
'Define the new JQLRequest
Dim PostTeamsRequest As New WebRequest
With PostTeamsRequest
    .Resource = "teams/1.0/teams/find"
    .Method = WebMethod.HttpPost
    Set .Body = JiraBody
End With
           
Dim PostTeamsFindResponse As New JiraResponse
Dim PostTeamsResponse As New WebResponse
 
Set PostTeamsResponse = PostTeamsFindResponse.JiraCall(PostTeamsRequest)
 
funcPostTeamsFind = PostTeamsResponse.StatusCode
 
Dim jiraTeam, jiraResource, jiraPerson As Object
Dim t%, p%, r%, l As Integer
 
If funcPostTeamsFind = Ok Then
    t = 1 'reset the teams to 1
    r = 2
    clearOldData ws_TeamsData
    With ws_TeamsData
        .Activate
        For Each jiraTeam In PostTeamsResponse.Data("teams")
            p = 1
            For Each jiraResource In PostTeamsResponse.Data("teams")(t)("resources")
                .Cells(r, 1).Value = PostTeamsResponse.Data("teams")(t)("id")
                .Cells(r, 2).Value = PostTeamsResponse.Data("teams")(t)("title")
                .Cells(r, 3).Value = PostTeamsResponse.Data("teams")(t)("resources")(p)("id")
                .Cells(r, 4).Value = PostTeamsResponse.Data("teams")(t)("resources")(p)("personId")
                p = p + 1
                r = r + 1
            Next jiraResource
            t = t + 1
        Next jiraTeam
        r = 2
        p = 1
        For Each jiraPerson In PostTeamsResponse.Data("persons")
            .Cells(r, 5).Value = PostTeamsResponse.Data("persons")(p)("personId")
            .Cells(r, 6).Value = PostTeamsResponse.Data("persons")(p)("jiraUser")("jiraUsername")
            p = p + 1
            r = r + 1
        Next jiraPerson
    End With
End If
 
End Function
 
Private Function funcGetSprintBurnDown(ByVal rapidViewId As String, ByVal sprintId As String) As WebStatusCode
 
'' This function records a log of time spent against each issue during a sprint (from taken from the Sprint BurnDown Chart)
'' It also updates the RemainingSprintTime public variable
 
' @param {String} rapidViewId
' @param {String} TeamId
' @param {String} SprintId
 
' @write {ws_Work}
' @return {WebStatusCode} status of apicall
''
 
'Define the new Request
Dim SprintBurnDownRequest As New WebRequest
With SprintBurnDownRequest
    .Resource = "greenhopper/1.0/rapid/charts/scopechangeburndownchart"
    .Method = WebMethod.HttpGet
    .AddQuerystringParam "rapidViewId", rapidViewId
    .AddQuerystringParam "sprintId", sprintId
End With
           
Dim SprintBurnDownChartResponse As New JiraResponse
Dim SprintBurnDownResponse As New WebResponse
 
Set SprintBurnDownResponse = SprintBurnDownChartResponse.JiraCall(SprintBurnDownRequest)
 
funcGetSprintBurnDown = SprintBurnDownResponse.StatusCode
 
If funcGetSprintBurnDown = Ok Then
    clearOldData ws_Work
    RemainingSprintTime = 0
    DaysInSprint = 0
   
    Dim time, change, rates As Object
    Dim c%, r%, d As Integer
 
    For Each time In SprintBurnDownResponse.Data("changes").Keys
        c = 1
        For Each change In SprintBurnDownResponse.Data("changes")(time)
            If SprintBurnDownResponse.Data("changes")(time)(c).Exists("timeC") Then
                If Val(time) < SprintBurnDownResponse.Data("completeTime") Then
                    RemainingSprintTime = RemainingSprintTime + _
                        (SprintBurnDownResponse.Data("changes")(time)(c)("timeC")("newEstimate") - SprintBurnDownResponse.Data("changes")(time)(c)("timeC")("oldEstimate"))
                    '' This next statement records a log of the issues that have had work logged to them during the sprint
                    If Val(time) > SprintBurnDownResponse.Data("startTime") Then
                        If SprintBurnDownResponse.Data("changes")(time)(c)("timeC").Exists("timeSpent") Then
                            With ws_Work.Range("A1048576").End(xlUp)
                                .Offset(1).Value = SprintBurnDownResponse.Data("changes")(time)(c)("key")
                                .Offset(1, 1).Value = rapidViewId
                                .Offset(1, 2).Value = SprintBurnDownResponse.Data("changes")(time)(c)("timeC")("timeSpent")
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
For Each rates In SprintBurnDownResponse.Data("workRateData")("rates")
    If SprintBurnDownResponse.Data("workRateData")("rates")(r)("rate") = 1 Then
        DaysInSprint = DaysInSprint + SprintBurnDownResponse.Data("workRateData")("rates")(r)("end") - SprintBurnDownResponse.Data("workRateData")("rates")(r)("start")
    End If
    r = r + 1
Next rates
 
DaysInSprint = DaysInSprint / 86400 / 1000
 
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
    .Range("BB25").Value = Round(StoryPointEstimate, 0)
    .Range("J10").Value = CStr(Round(SubTaskEstimate, 0)) & "/" & CStr(Round(StoryPointEstimate, 0)) & "/" & CStr(Round(TShirtEstimate, 0))
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
    .Range("BB26").Value = ws_VelocityData.Range("E2").Value / ws_VelocityData.Range("D2").Value
    .Range("S10").Value = ws_VelocityData.Range("E2").Value / ws_VelocityData.Range("D2").Value
End With
 
End Function
Function funcPredictabilityTiPVariability()
 
'' Update the TeamStats worksheet with the *TiP Variability* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras & funcResponsivenessTiP (see below)
''
 
'' Known limitations with this macro:
' (1) ws_LeadTimeData sheet needs to be made active
 
Dim rng As Range
Dim TiPCol As Integer
Dim TipRow As Long
 
ws_LeadTimeData.Activate
TiPCol = ws_LeadTimeData.Range("1:1").Find("TiP").column
TipRow = ws_LeadTimeData.Cells(1, TiPCol).End(xlDown).row
 
Set rng = ws_LeadTimeData.Range(Cells(2, TiPCol), Cells(TipRow, TiPCol))
 
With ws_TeamStats
    .Range("BB27").Value = Excel.WorksheetFunction.StDev_P(rng) / Excel.WorksheetFunction.Average(rng)
    .Range("J16").Value = Excel.WorksheetFunction.StDev_P(rng) / Excel.WorksheetFunction.Average(rng)
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
    .Range("BB28").Value = Excel.WorksheetFunction.StDev_P(ws_VelocityData.Range("E2:E8")) / Excel.WorksheetFunction.Average(ws_VelocityData.Range("E2:E8"))
    .Range("S16").Value = Excel.WorksheetFunction.StDev_P(ws_VelocityData.Range("E2:E8")) / Excel.WorksheetFunction.Average(ws_VelocityData.Range("E2:E8"))
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
    .Activate
    col = .Range("A1").End(xlToRight).column + 1
    .Cells(1, col).Value = "leadTime"
 
    Set rng_LeadTime = .Range(Cells(2, col), Cells(.Range("A1").End(xlDown).row, col))
 
    For Each c In rng_LeadTime
        c.Value = CDate(.Cells(c.row, 6).Value) - CDate(Left(.Cells(c.row, 4).Value, 10))
    Next c
   
End With
 
 
With ws_TeamStats
    .Range("BB30").Value = Excel.WorksheetFunction.Average(rng_LeadTime)
    .Range("AD10").Value = Round(Excel.WorksheetFunction.Average(rng_LeadTime), 0)
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
       
For Each cell In ws_LeadTimeData.Range(Cells(2, 6), Cells(ws_LeadTimeData.Range("F1").End(xlDown).row, 6))
    If cell.Value >= DateAdd("m", -1, "01/" & Month(Now()) & "/" & Year(Now())) Then ' only count if after start of previous month
        If cell.Value < CDate("01/" & Month(Now()) & "/" & Year(Now())) Then ' only count if before start of current month
            If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, 0
            End If
        End If
    End If
Next
 
With ws_TeamStats
    .Range("BB31").Value = dict.Count
    .Range("AM10").Value = dict.Count
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
 
    col = .Range("A1").End(xlToRight).column + 1
    .Cells(1, col).Value = "TiP"
 
    Set rng_TiP = .Range(Cells(2, col), Cells(.Range("A1").End(xlDown).row, col))
 
    For Each c In rng_TiP
        c.Value = CDate(.Cells(c.row, 6).Value) - CDate(Left(.Cells(c.row, 5).Value, 10))
    Next c
   
End With
 
 
With ws_TeamStats
    .Range("BB32").Value = Excel.WorksheetFunction.Average(rng_TiP)
    .Range("AD16").Value = Round(Excel.WorksheetFunction.Average(rng_TiP), 0)
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
Dim startDatesRange As Range, endDatesRange As Range
Set startDatesRange = ws_WiPData.Range(Cells(2, 4), Cells(ws_WiPData.Range("D2").End(xlDown).row, 4))
Set endDatesRange = ws_WiPData.Range(Cells(2, 5), Cells(ws_WiPData.Range("E2").End(xlDown).row, 5))

'' Determine the headings of a grid as the min and max dates
Dim HeadingsArr As Variant
Dim HeadingsRange As Range
HeadingsArr = ArrayOfDates(MinMaxDate(startDatesRange, "Min"), MinMaxDate(endDatesRange, "Max"))
Set HeadingsRange = ws_WiPData.Range("H1").Resize(1, UBound(HeadingsArr))
HeadingsRange.Value = HeadingsArr ' assumes a one dimensional array; base 1
 
'' Create a 2 dimensional array to hold values for when the issue was actively in progress
Dim r As Range, c As Range
Dim WipGridArr As Variant
ReDim WipGridArr(1 To startDatesRange.Rows.Count, 1 To HeadingsRange.Columns.Count) As Integer
Dim x%, y As Integer
Dim startDate&, endDate As Long
y = 1
For Each r In startDatesRange
    x = 1
    For Each c In HeadingsRange
        startDate = DateValue(Left(r.Value, 10))
        endDate = DateValue(Left(r.Offset(0, 1).Value, 10))
        WipGridArr(y, x) = inProgressForDate(startDate, endDate, c.Value)
        x = x + 1
    Next c
    y = y + 1
Next r

'' Write the results of the 2 dimensional array to the sheet
Dim WipGridRange As Range
Set WipGridRange = ws_WiPData.Range("H2").Resize(startDatesRange.Rows.Count, HeadingsRange.Columns.Count)
WipGridRange.Value = WipGridArr

'' Calculate the WiP for each issue
For Each r In startDatesRange
    r.Offset(0, 3) = WiP(r.row)
Next r
 
With ws_TeamStats
    .Range("BB33").Value = WorksheetFunction.Average(ws_WiPData.Range(Cells(2, 7), Cells(startDatesRange.Rows.Count + 1, 7))) ' Forumla to be updated
    .Range("AM16").Value = WorksheetFunction.Average(ws_WiPData.Range(Cells(2, 7), Cells(startDatesRange.Rows.Count + 1, 7))) ' Forumla to be updated
End With
 
MsgBox ("Done")
 
End Function


Function funcProductivityReleaseVelocity()
 
'' Update the TeamStats worksheet with the *Release Velocity* data
'
' Dependent on function: funcGetDoneJiras
'
''
 
With ws_TeamStats
    .Range("AT5").Value = 0 ' Forumla to be updated - Feature Velocity
    .Range("AU5").Value = 0 ' Forumla to be updated - Defects Velocity
    .Range("AV5").Value = 0 ' Forumla to be updated - Risks Velocity
    .Range("AW5").Value = 0 ' Forumla to be updated - Debts Velocity
    .Range("AX5").Value = 0 ' Forumla to be updated - Enablers Velocity
   
    .Range("AT6").Value = 0 ' Forumla to be updated - Feature Baseline
    .Range("AU6").Value = 0 ' Forumla to be updated - Defects Baseline
    .Range("AV6").Value = 0 ' Forumla to be updated - Risks Baseline
    .Range("AW6").Value = 0 ' Forumla to be updated - Debts Baseline
    .Range("AX6").Value = 0 ' Forumla to be updated - Enablers Baseline
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
    .Range("BB35").Value = 0 ' Forumla to be updated
    .Range("AB24").Value = 0 ' Forumla to be updated
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
    .Range("AT11:BF15").Value = 0 ' Forumla to be updated - Data
    .Range("AT16:BE16").Value = "Date" ' Formula to be updated - Release Months
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
    .Range("BB37").Value = 0 ' Forumla to be updated
    .Range("AM24").Value = 0 ' Forumla to be updated
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
    .Range("BB38").Value = "TBC" ' Forumla to be updated
    .Range("AM30").Value = "TBC" ' Forumla to be updated
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
    .Range("BB39").Value = "TBC" ' Forumla to be updated
    .Range("AM36").Value = "TBC" ' Forumla to be updated
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
    .Range("BB22").Value = 0 ' Forumla to be updated
    .Range("J5").Value = 0 ' Forumla to be updated
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
    .Range("BB23").Value = RemainingSprintTime
    .Range("AD5").Value = Jira.jiratime(RemainingSprintTime)
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
    .Range("BB41").Value = 0 ' Forumla to be updated
    .Range("J43").Value = 0 ' Forumla to be updated
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
    .Range("BB42").Value = 0 ' Forumla to be updated
    .Range("T43").Value = 0 ' Forumla to be updated
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
    .Range("BB43").Value = "TBC" ' Forumla to be updated
    .Range("AC43").Value = "TBC" ' Forumla to be updated
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
    .Range("BB41").Value = 0 ' Forumla to be updated
    .Range("J43").Value = 0 ' Forumla to be updated
End With
 
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
 
 
Private Function funcRAG()
 
'' This function should update the RAG triangles for each of the values displayed on the Dashboard
 
End Function
 
Private Function funcAsOfDateTeamName()
 
'' This function should update the AsOfDate and Team Name displayed on the Dashboard
 
End Function
Private Function clearOldData(ByVal ws As Worksheet)
 
Dim rngOldData As Range
    With ws
        Set rngOldData = .Range("A1").CurrentRegion
        If rngOldData.Rows.Count > 1 Then
            Set rngOldData = rngOldData.Resize(rngOldData.Rows.Count - 1).Offset(1)
            rngOldData.ClearContents ' clear existing data
        End If
    End With
   
End Function
Function CreateWorkSheet(ByVal name As String, Optional ByRef headings As Variant) As Worksheet
 
'' Checks if a Worksheet exists and creates one if it doesn't
 
Dim ws_exists As Boolean
Dim ws As Worksheet
 
    For Each ws In ActiveWorkbook.Worksheets
        If ws.name = name Then
            ws_exists = True
            Exit For
        Else
            ws_exists = False
        End If
    Next ws
 
    If ws_exists Then
        Set CreateWorkSheet = ActiveWorkbook.Worksheets(name)
    Else
        Set CreateWorkSheet = ActiveWorkbook.Sheets.Add
        CreateWorkSheet.name = name
        If Not IsMissing(headings) Then
            CreateWorkSheet.Range("A1").Resize(1, UBound(headings)).Value = headings ' assumes a one dimensional array; base 1
        End If
    End If
 
End Function
Function ws_TeamStats() As Worksheet
    Set ws_TeamStats = CreateWorkSheet("ws_TeamStats")
End Function
Function ws_LeadTimeData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "issueType", "createdDate", "sprintStartDate", "releaseDate", "totalTime", "totalString")
    Set ws_LeadTimeData = CreateWorkSheet("ws_LeadTimeData", HeadingsArr)
End Function
Function ws_WiPData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "issueType", "inProgressDate", "endProgressDate", "releaseDate", "WiP")
    Set ws_WiPData = CreateWorkSheet("ws_WiPData", HeadingsArr)
End Function
Function ws_IncompleteIssuesData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "issueType", "project", "epicKey", "storyPoints", "status", "statusCategory", "aggregateTimeEstimate", "sprintState")
    Set ws_IncompleteIssuesData = CreateWorkSheet("ws_IncompleteIssuesData", HeadingsArr)
End Function
Function ws_VelocityData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("SprintId", "SprintName", "State", "Committed", "Completed")
   Set ws_VelocityData = CreateWorkSheet("ws_VelocityData", HeadingsArr)
End Function
Function ws_TeamsData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "title", "resourceId", "personId", "personId2", "JiraUserName")
    Set ws_TeamsData = CreateWorkSheet("ws_TeamsData", HeadingsArr)
End Function
Function ws_Work() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("key", "rapidViewId", "timeSpent")
    Set ws_Work = CreateWorkSheet("ws_Work", HeadingsArr)
End Function
Function ws_ProjectData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "projectName", "category")
    Set ws_ProjectData = CreateWorkSheet("ws_ProjectData", HeadingsArr)
End Function
Function TeamId() As String
''Placeholder to define other values
    TeamId = "81"
End Function
Function rapidViewId() As String
    rapidViewId = InputBox("rapidViewId?")
End Function
Function boardJql() As String
''Placeholder to define other values
    boardJql = "Team = 81 AND CATEGORY = calm AND NOT issuetype in (Initiative) ORDER BY Rank ASC"
End Function
Function ArrayOfDates(ByVal startDate As Long, ByVal endDate As Long) As Variant()

    Dim arr() As Variant
    Dim DateLoop As Variant
    Dim i%, totalDays As Integer
    DateLoop = startDate
    totalDays = endDate - startDate
    ReDim ArrayOfDates(1 To totalDays + 1)
    ReDim arr(1 To totalDays + 1)
    i = 1
    Do While DateLoop <= endDate
        arr(i) = DateLoop
        DateLoop = DateLoop + 1
        i = i + 1
    Loop
    ArrayOfDates = arr
    
End Function

Function MinMaxDate(ByVal dateRange As Range, ByVal MType As String) As Variant
    Dim c As Range
    Dim arr() As Long
    Dim totalDays As Integer
    totalDays = dateRange.Rows.Count
    ReDim arr(1 To totalDays)
    Dim i As Integer
    i = 1
    For Each c In dateRange
        arr(i) = DateValue(Left(c.Value, 10))
        i = i + 1
    Next c
    
    If MType = "Max" Then
        MinMaxDate = WorksheetFunction.Max(arr)
    ElseIf MType = "Min" Then
        MinMaxDate = WorksheetFunction.Min(arr)
    Else
        MinMaxDate = 0
    End If
    
End Function
Function inProgressForDate(ByVal startDate As Long, ByVal endDate As Long, currentDate As Long) As Integer
    If currentDate >= startDate Then
        If currentDate <= endDate Then
            inProgressForDate = 1
        Else
            inProgressForDate = 0
        End If
    Else
        inProgressForDate = 0
    End If
End Function

Function WiP(ByVal row As Long) As Integer
    
    Dim headers As Range
    Set headers = ws_WiPData.Range(Cells(1, 8), Cells(1, ws_WiPData.Range("H1").End(xlToRight).column))

    Dim countRows As Long
    countRows = ws_WiPData.Range("H1").End(xlDown).row - 1
    
    Dim arr() As Integer
    ReDim arr(1 To headers.Columns.Count)
    Dim c As Range
    Dim i As Integer
    i = 1
    For Each c In headers
        If c.Offset(row).Value = 1 Then
            arr(i) = WorksheetFunction.Sum(c.Resize(countRows, 1).Offset(1))
        End If
        i = i + 1
    Next c
    
    WiP = WorksheetFunction.Max(arr)
    
End Function

