Attribute VB_Name = "JiraScrumTeamStats"
''
' VBA-JiraScrumTeamStats v1.0
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

Sub UpdateIssueTypeMapping(control As IRibbonControl)

    '' Need to use this data in other macros
    '' needs refactoring to remove duplication of code into sub functions
    
    Dim arr() As String
    Dim i As Integer
    Dim strFeatures$, strDefects$, strRisks$, strDebts As String
     
    '' Set the Default values based on either the Jira Default or the Last Saved values
    If vbaJiraProperties.Range("K1") = 0 Then
        strFeatures = "Story"
    Else
        strFeatures = vbaJiraProperties.Range("K1").Value
    End If
    If vbaJiraProperties.Range("K2") = 0 Then
        strDefects = "Bug"
    Else
        strDefects = vbaJiraProperties.Range("K2").Value
    End If
    If vbaJiraProperties.Range("K3") = 0 Then
        strRisks = ""
    Else
        strRisks = vbaJiraProperties.Range("K3").Value
    End If
    If vbaJiraProperties.Range("K4") = 0 Then
        strDebts = ""
    Else
        strDebts = vbaJiraProperties.Range("K4").Value
    End If
    
    '' Define the new API Request
    Dim IssueTypeRequest As New WebRequest
    With IssueTypeRequest
        .Resource = "api/2/issuetype"
        .Method = WebMethod.HttpGet
    End With
               
    Dim IssueTypeJiraResponse As New JiraResponse
    Dim IssueTypeResponse As New WebResponse
     
    Set IssueTypeResponse = IssueTypeJiraResponse.JiraCall(IssueTypeRequest)
     
    If IssueTypeResponse.StatusCode = 200 Then
        Dim JiraIssueTypeIndex As Long
        Dim foundIssue() As Boolean
        
        vbaJiraProperties.Range("J1:K1").Value = Array("Features:", _
            InputBox("Name of Jira IssueType(s) that maps to Features? Seperate multiple Issue Types with a comma (,)", _
            "Features", strFeatures))
        arr = Split(vbaJiraProperties.Range("K1"), ",")
        For i = LBound(arr) To UBound(arr)
            ReDim foundIssue(0 To UBound(arr)) As Boolean
            foundIssue(i) = False
            For JiraIssueTypeIndex = 1 To IssueTypeResponse.Data.Count
                    If IssueTypeResponse.Data(JiraIssueTypeIndex)("name") = arr(i) Then
                        If IssueTypeResponse.Data(JiraIssueTypeIndex)("subtask") = True Then
                            MsgBox ("The following Issue Type is a Sub-Task, please try again: " & arr(i))
                            Exit Sub
                        Else
                            foundIssue(i) = True
                            Exit For
                        End If
                    End If
            Next JiraIssueTypeIndex
            If Not foundIssue(i) Then ' If the issue type could not be found in Jira
                MsgBox ("The following Issue Type could not be found in Jira, please try again: " & arr(i))
                Exit Sub
            End If
        Next i
            
        vbaJiraProperties.Range("J2:K2").Value = Array("Defects:", _
            InputBox("Name of Jira IssueType(s) that maps to Defects? Seperate multiple Issue Types with a comma (,)", _
            "Defects", strDefects))
        arr = Split(vbaJiraProperties.Range("K2"), ",")
        For i = LBound(arr) To UBound(arr)
            ReDim foundIssue(0 To UBound(arr)) As Boolean
            foundIssue(i) = False
            For JiraIssueTypeIndex = 1 To IssueTypeResponse.Data.Count
                    If IssueTypeResponse.Data(JiraIssueTypeIndex)("name") = arr(i) Then
                        If IssueTypeResponse.Data(JiraIssueTypeIndex)("subtask") = True Then
                            MsgBox ("The following Issue Type is a Sub-Task, please try again: " & arr(i))
                            Exit Sub
                        Else
                            foundIssue(i) = True
                            Exit For
                        End If
                    End If
            Next JiraIssueTypeIndex
            If Not foundIssue(i) Then ' If the issue type could not be found in Jira
                MsgBox ("The following Issue Type could not be found in Jira, please try again: " & arr(i))
                Exit Sub
            End If
        Next i
        
        vbaJiraProperties.Range("J3:K3").Value = Array("Risks:", _
            InputBox("Name of Jira IssueType(s) that maps to Risks? Seperate multiple Issue Types with a comma (,)", _
            "Risks", strRisks))
        arr = Split(vbaJiraProperties.Range("K3"), ",")
        For i = LBound(arr) To UBound(arr)
            ReDim foundIssue(0 To UBound(arr)) As Boolean
            foundIssue(i) = False
            For JiraIssueTypeIndex = 1 To IssueTypeResponse.Data.Count
                    If IssueTypeResponse.Data(JiraIssueTypeIndex)("name") = arr(i) Then
                        If IssueTypeResponse.Data(JiraIssueTypeIndex)("subtask") = True Then
                            MsgBox ("The following Issue Type is a Sub-Task, please try again: " & arr(i))
                            Exit Sub
                        Else
                            foundIssue(i) = True
                            Exit For
                        End If
                    End If
            Next JiraIssueTypeIndex
            If Not foundIssue(i) Then ' If the issue type could not be found in Jira
                MsgBox ("The following Issue Type could not be found in Jira, please try again: " & arr(i))
                Exit Sub
            End If
        Next i
        
        vbaJiraProperties.Range("J4:K4").Value = Array("Debts:", _
            InputBox("Name of Jira IssueType(s) that maps to Debts? Seperate multiple Issue Types with a comma (,)", _
            "Debts", strDebts))
        arr = Split(vbaJiraProperties.Range("K4"), ",")
        For i = LBound(arr) To UBound(arr)
            ReDim foundIssue(0 To UBound(arr)) As Boolean
            foundIssue(i) = False
            For JiraIssueTypeIndex = 1 To IssueTypeResponse.Data.Count
                    If IssueTypeResponse.Data(JiraIssueTypeIndex)("name") = arr(i) Then
                        If IssueTypeResponse.Data(JiraIssueTypeIndex)("subtask") = True Then
                            MsgBox ("The following Issue Type is a Sub-Task, please try again: " & arr(i))
                            Exit Sub
                        Else
                            foundIssue(i) = True
                            Exit For
                        End If
                    End If
            Next JiraIssueTypeIndex
            If Not foundIssue(i) Then ' If the issue type could not be found in Jira
                MsgBox ("The following Issue Type could not be found in Jira, please try again: " & arr(i))
                Exit Sub
            End If
        Next i
        MsgBox ("Successfully Updated")
    Else
        MsgBox ("Error Getting Issue Types from Jira: " & IssueTypeResponse.StatusCode)
    End If
End Sub
Sub GetTeamStats(control As IRibbonControl)
 
''
' This should be run by the user and sets up all the underlying api calls to get the teams stats
 
'' Known limitations with this macro:
' (0) Still a Work in Progress - use with caution!
' (1) Needs some user feedback (i.e. progress bar) to show progress
' (2) Some of the API calls would benefit from being looped over a smaller set of maxResults
' (3) Error logging
' (4) Capture veriables, such as TeamId and allow for user configuration
' (5) Validate that the rapidViewId has time tracking available before running certain functions
 
'Pause calculations and screen updating and make read-only worksheets visible
'These actions are reversed at the end of the macro
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
ws_TeamStats.Unprotect ("KM_e@UyRnMtTqvWpd3NG")
 
    ' --- Comment out the respective value to enable or suspend logging
'    WebHelpers.EnableLogging = True
    WebHelpers.EnableLogging = False
   
    'Check if a user is logged in and if not perform login, if login fails exit
    If Not IsLoggedIn Then
        If Not LoginUser Then Exit Sub
    End If
   
    'Rollstats True of False (copy across values from the last time the macros were run or False will overwrite)
    Dim blnRoll As Boolean
    If MsgBox("Do you want to roll the previous data?", vbYesNo) = vbYes Then
        blnRoll = True
    Else
        blnRoll = False
    End If
    funcRollStats (blnRoll) 'Roll the stats
    funcAsOfDateTeamName 'Update the TeamName and As Of Data
        
    ''Fetch Data from Api Calls
    Dim callResult(1 To 9) As WebStatusCode
        Debug.Print (Now())
    callResult(1) = funcGet3MonthsOfDoneJiras(boardJql, "In Progress", "Done", 0, 2)
        Debug.Print ("funcGet3MonthsOfDoneJiras: " & callResult(1) & " : " & Now())
    callResult(2) = funcGetIncompleteJiras(boardJql, 0, 2)
        Debug.Print ("funcGetIncompleteJiras: " & callResult(2) & " : " & Now())
    callResult(3) = funcGet12MonthDoneJiras(boardJql, 0, 2)
        Debug.Print ("funcGet12MonthDoneJiras: " & callResult(3) & " : " & Now())
    callResult(4) = funcGetDefects(boardJql, 0, 2)
        Debug.Print ("funcGetDefects: " & callResult(4) & " : " & Now())
    callResult(5) = funcGetVelocity(rapidViewId)
        Debug.Print ("funcGetVelocity: " & callResult(5) & " : " & Now())
    callResult(6) = funcPostTeamsFind()
        Debug.Print ("funcPostTeamsFind: " & callResult(6) & " : " & Now())
    callResult(7) = funcGetSprintBurnDown(rapidViewId, CStr(ws_VelocityData.Range("A2").Value))
        Debug.Print ("funcGetSprintBurnDown: " & callResult(7) & " : " & Now())
    callResult(8) = funcGetSprintDetails(CStr(ws_VelocityData.Range("A2").Value))
        Debug.Print ("funcGetSprintDetails: " & callResult(8) & " : " & Now())
    callResult(9) = funcGetSprintWorkLog(TeamResourcesString, CStr(ws_VelocityData.Range("F2").Value), CStr(ws_VelocityData.Range("A2").Value))
        Debug.Print ("funcGetSprintWorkLog: " & callResult(9) & " : " & Now())
    ' Input the most recent sprint name
    ws_TeamStats.Range("AS3").Value = ws_VelocityData.Range("B2").Value
        Debug.Print (Now())
    'Run the calculations - note the order is specfic
    funcPredictabilitySprintsEstimated
        Debug.Print ("funcPredictabilitySprintsEstimated: " & Now())
    funcPredictabilityVelocity
        Debug.Print ("funcPredictabilityVelocity: " & Now())
    funcPredictabilitySprintOutputVariability
        Debug.Print ("funcPredictabilitySprintOutputVariability: " & Now())
    funcResponsivenessLeadTime
        Debug.Print ("funcResponsivenessLeadTime: " & Now())
    funcResponsivenessDeploymentFrequency
        Debug.Print ("funcResponsivenessDeploymentFrequency: " & Now())
    funcResponsivenessTiP
        Debug.Print ("funcResponsivenessTiP: " & Now())
    funcPredictabilityTiPVariability
        Debug.Print ("funcPredictabilityTiPVariability: " & Now())
    funcScrumUnplannedWork
        Debug.Print ("funcScrumUnplannedWork: " & Now())
    funcResponsivenessWiP
        Debug.Print ("funcResponsivenessWiP: " & Now())
    funcProductivityReleaseVelocity
        Debug.Print ("funcProductivityReleaseVelocity: " & Now())
    funcProductivityEfficiency
        Debug.Print ("funcProductivityEfficiency: " & Now())
    funcProductivityDistribution
        Debug.Print ("funcProductivityDistribution: " & Now())
    funcQualityTimeToResolve
        Debug.Print ("funcQualityTimeToResolve: " & Now())
    funcQualityDefectDentisy
        Debug.Print ("funcQualityDefectDentisy: " & Now())
    funcQualityFailRate
        Debug.Print ("funcQualityFailRate: " & Now())
    funcScrumTeamStability
        Debug.Print ("funcScrumTeamStability: " & Now())
    funcJiraAdminActiveTime
        Debug.Print ("funcJiraAdminActiveTime: " & Now())

ws_TeamStats.Activate
    
'Reverse the opening statements that paused calculations and screen updating
ws_TeamStats.Protect ("KM_e@UyRnMtTqvWpd3NG")
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
        .Range("AW22:BA43").Value = .Range("AX22:BB43").Value
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
 
Dim JQL As String
JQL = "fixversion changed after -24w AND " & _
        "fixVersion is not EMPTY AND " & _
        "Sprint is not EMPTY AND " & _
        "issuetype in (" & issueTypeSearchString & ") AND " & _
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
    .AddQuerystringParam "jql", JQL
    .AddQuerystringParam "fields", apiFields
    .AddQuerystringParam "startAt", startAtVal
    .AddQuerystringParam "maxResults", "1000"
    .AddQuerystringParam "expand", "changelog"
End With
           
Dim JQL_PBI_Search_Response As New JiraResponse
Dim JQL_Search_Response As New WebResponse
Dim item As Object
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
 
If funcGet3MonthsOfDoneJiras = OK Then
    clearOldData ws_LeadTimeData
    clearOldData ws_WiPData
    Set dictTimeLoggedToStory = New Dictionary
    startAtVal = startAtVal + 1000 'Increment the next start position based on maxResults above -- making this smaller will speed up the API calls
    i = 1 'reset the issue to 1
    WiPRow = r
    For Each item In JQL_Search_Response.Data("issues")
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
                                    If ws_WiPData.Cells(WiPRow, 1).Value = 0 Then 'If issue transitioned straight to its endProgressState
                                        ws_WiPData.Cells(WiPRow, 1).Value = JQL_Search_Response.Data("issues")(i)("id")
                                        ws_WiPData.Cells(WiPRow, 2).Value = JQL_Search_Response.Data("issues")(i)("key")
                                        ws_WiPData.Cells(WiPRow, 3).Value = JQL_Search_Response.Data("issues")(i)("fields")("issuetype")("name")
                                        For Each fixversion In JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")
                                            ws_WiPData.Cells(WiPRow, 6).Value = JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")(1)("releaseDate")  'Always use the 1st fixVersion, even if there are multiple
                                        Next
                                        ws_WiPData.Cells(WiPRow, 4).Value = JQL_Search_Response.Data("issues")(i)("changelog")("histories")(h)("created")
                                    End If
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
        JQL = "Parent = " & rng_Parent.Value
   
        apiFields = "key," _
            & "issuetype," _
            & "changelog"
 
        Dim JQL_SubTask_Request As New WebRequest
        With JQL_SubTask_Request
            .Resource = "api/2/search"
            .Method = WebMethod.HttpGet
            .AddQuerystringParam "jql", JQL
            .AddQuerystringParam "fields", apiFields
            .AddQuerystringParam "startAt", "0"
            .AddQuerystringParam "maxResults", "1000"
            .AddQuerystringParam "expand", "changelog"
        End With
 
        Dim JQL_SubTask_Search_Response As New JiraResponse
        Set JQL_Search_Response = JQL_SubTask_Search_Response.JiraCall(JQL_SubTask_Request)
 
        i = 1 'reset the issue to 1
        For Each item In JQL_Search_Response.Data("issues")
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
        For Each rng_author In ws_LeadTimeData.Range(Cells(1, 9), Cells(1, ws_LeadTimeData.Range("A1").End(xlToRight).Column))
            ws_LeadTimeData.Cells(rng_Parent.row, rng_author.Column) = collIssueKey(rng_Parent.Value)(rng_author.Value)
        Next rng_author
       
        col = ws_LeadTimeData.Range("A1").End(xlToRight).Column
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
 
Dim JQL As String
JQL = "statusCategory not in (Done) AND " & _
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
    .AddQuerystringParam "jql", JQL
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
Dim item As Object
 
If funcGetIncompleteJiras = OK Then
    clearOldData ws_IncompleteIssuesData
    startAtVal = startAtVal + 1000 'Increment the next start position based on maxResults above
    i = 1 'reset the issue to 1
    For Each item In JQL_Search_Response.Data("issues")
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
    Next item
End If
 
End Function
Private Function funcGet12MonthDoneJiras(ByVal boardJql As String, ByRef startAtVal, r As Integer) As WebStatusCode
 
''
' Source 12 months of Jiras that done state and not subTasks
 
'
' @param {String} boardJql
' @param {Integer} startAtVal, r
' @write {ws_DoneData}
' @apicalls 1x{get search standardissuetypes}
' @return {WebStatusCode} status of apicall
''
 
'' Known limitations with this macro:
' (1) needs to be updated to run for a smaller number of maxresults
 
Dim JQL As String
JQL = "fixversion changed after -60w AND " & _
        "fixVersion is not EMPTY AND " & _
        "Sprint is not EMPTY AND " & _
        "issuetype in (" & issueTypeSearchString & ") AND " & _
        "statusCategory in (Done) AND " & _
        boardJql
       
Dim apiFields As String
apiFields = "key," _
        & "issuetype," _
        & "fixVersions"
           
'Define the new Request
Dim JQL_PBI_Request As New WebRequest
With JQL_PBI_Request
    .Resource = "api/2/search"
    .Method = WebMethod.HttpGet
    .AddQuerystringParam "jql", JQL
    .AddQuerystringParam "fields", apiFields
    .AddQuerystringParam "startAt", startAtVal
    .AddQuerystringParam "maxResults", "1000"
    .AddQuerystringParam "expand", "changelog"
End With
           
Dim JQL_PBI_Search_Response As New JiraResponse
Dim JQL_Search_Response As New WebResponse
 
Set JQL_Search_Response = JQL_PBI_Search_Response.JiraCall(JQL_PBI_Request)
 
funcGet12MonthDoneJiras = JQL_Search_Response.StatusCode
 
Dim i%, s As Integer
Dim item As Object
 
If funcGet12MonthDoneJiras = OK Then
    clearOldData ws_DoneData
    startAtVal = startAtVal + 1000 'Increment the next start position based on maxResults above
    i = 1 'reset the issue to 1
    For Each item In JQL_Search_Response.Data("issues")
        If CDate(JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")(1)("releaseDate")) >= DateAdd("m", -12, "01/" & Month(Now()) & "/" & Year(Now())) Then 'Only include if the release date was in the last 12 months
            With ws_DoneData
                .Cells(r, 1).Value = JQL_Search_Response.Data("issues")(i)("id")
                .Cells(r, 2).Value = JQL_Search_Response.Data("issues")(i)("key")
                .Cells(r, 3).Value = JQL_Search_Response.Data("issues")(i)("fields")("issuetype")("name")
                If JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")(1)("releaseDate") > JQL_Search_Response.Data("issues")(i)("fields")("created") Then
                    .Cells(r, 4).Value = JQL_Search_Response.Data("issues")(i)("fields")("fixVersions")(1)("releaseDate") 'Always use the 1st fixVersion, even if there are multiple
                Else
                    .Cells(r, 4).Value = Left(JQL_Search_Response.Data("issues")(i)("fields")("resolutiondate"), 10) 'use the resolution date if there is no fixVersion.
                End If
            End With
            r = r + 1 'increment the row
        End If
        i = i + 1 'increment the issue
    Next item
End If
 
End Function
Private Function funcGetDefects(ByVal boardJql As String, ByRef startAtVal, r As Integer) As WebStatusCode
 
''
' Source Bugs and Defefcts that have been created in the last 3 months
'
' @param {String} boardJql
' @param {Integer} startAtVal, r
' @write {ws_DefectData}
' @apicalls 1x{get search standardissuetypes}
' @return {WebStatusCode} status of apicall
''
 
'' Known limitations with this macro:
' (1) needs to be updated to run for a smaller number of maxresults
 
Dim JQL As String
''Get the string of Defect issue types from the vbaJiraProperties
'' This no longer includes SubTasks for the calculation of defect density
JQL = "issuetype in (" & vbaJiraProperties.Range("K2").Value & ") AND created >= startOfMonth(-3) AND " & _
        boardJql
       
Dim apiFields As String
apiFields = "key," _
        & "issuetype," _
        & "versions," _
        & "fixVersions"
           
'Define the new Request
Dim JQL_PBI_Request As New WebRequest
With JQL_PBI_Request
    .Resource = "api/2/search"
    .Method = WebMethod.HttpGet
    .AddQuerystringParam "jql", JQL
    .AddQuerystringParam "fields", apiFields
    .AddQuerystringParam "startAt", startAtVal
    .AddQuerystringParam "maxResults", "1000"
    .AddQuerystringParam "expand", "changelog"
End With
           
Dim JQL_PBI_Search_Response As New JiraResponse
Dim JQL_Search_Response As New WebResponse
 
Set JQL_Search_Response = JQL_PBI_Search_Response.JiraCall(JQL_PBI_Request)
 
funcGetDefects = JQL_Search_Response.StatusCode
 
Dim i%, s As Integer
Dim item As Object
  
If funcGetDefects = OK Then
    clearOldData ws_DefectData
    startAtVal = startAtVal + 1000 'Increment the next start position based on maxResults above
    i = 1 'reset the issue to 1
    For Each item In JQL_Search_Response.Data("issues")
            With ws_DefectData
                .Cells(r, 1).Value = JQL_Search_Response.Data("issues")(i)("id")
                .Cells(r, 2).Value = JQL_Search_Response.Data("issues")(i)("key")
                .Cells(r, 3).Value = JQL_Search_Response.Data("issues")(i)("fields")("issuetype")("name")
                If JQL_Search_Response.Data("issues")(i)("fields").Exists("versions") Then
                    .Cells(r, 4).Value = JQL_Search_Response.Data("issues")(i)("fields")("versions")(1)("name") 'Always use the 1st Affects Version, even if there are multiple
                    .Cells(r, 5).Value = JQL_Search_Response.Data("issues")(i)("fields")("versions")(1)("releaseDate") 'Always use the 1st Affects Version, even if there are multiple
                End If
            End With
            r = r + 1 'increment the row
        i = i + 1 'increment the issue
    Next item
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
 
Dim item As Object
Dim r%, s As Integer
 
If funcGetVelocity = OK Then
    r = 2
    s = 1
    clearOldData ws_VelocityData
    For Each item In VelocityResponse.Data("sprints")
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
' (2) need to add start date, end date and holiday accountability to the teams data
 
'Dim JsonPost As String
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
 
If funcPostTeamsFind = OK Then
    t = 1 'reset the teams to 1
    r = 2
    clearOldData ws_TeamsData
    With ws_TeamsData
        .Activate
        For Each jiraTeam In PostTeamsResponse.Data("teams")
            p = 1
            If PostTeamsResponse.Data("teams")(t)("id") = CStr(TeamId) Then ' Only import the Team that matches the one in the properties
                For Each jiraResource In PostTeamsResponse.Data("teams")(t)("resources")
                    .Cells(r, 1).Value = PostTeamsResponse.Data("teams")(t)("id")
                    .Cells(r, 2).Value = PostTeamsResponse.Data("teams")(t)("title")
                    .Cells(r, 3).Value = PostTeamsResponse.Data("teams")(t)("resources")(p)("id")
                    .Cells(r, 4).Value = PostTeamsResponse.Data("teams")(t)("resources")(p)("personId")
                    If PostTeamsResponse.Data("teams")(t)("resources")(p).Exists("weeklyHours") Then
                        .Cells(r, 5).Value = PostTeamsResponse.Data("teams")(t)("resources")(p)("weeklyHours")
                    Else
                        .Cells(r, 5).Value = 40
                    End If
                    p = p + 1
                    r = r + 1
                Next jiraResource
            End If
            t = t + 1
        Next jiraTeam
        r = 2
        p = 1
        For Each jiraPerson In PostTeamsResponse.Data("persons")
            If PostTeamsResponse.Data("persons")(p)("personId") = CStr(.Cells(r, 4).Value) Then
                .Cells(r, 6).Value = PostTeamsResponse.Data("persons")(p)("personId")
                .Cells(r, 7).Value = PostTeamsResponse.Data("persons")(p)("jiraUser")("jiraUsername")
                r = r + 1
            End If
            p = p + 1
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
 
If funcGetSprintBurnDown = OK Then
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

Private Function funcGetSprintDetails(ByVal requestedSprintId As String) As WebStatusCode
'' This is technically only required with older versions of Jira that don't have the updated Velocity Chart, otherwise this could be added to funcGetVelocity. I've kept it seperate for backwards compatibility

'Define the new Sprint Request
Dim Sprint_Request As New WebRequest
With Sprint_Request
    .Resource = "agile/1.0/sprint/{sprintId}"
    .Method = WebMethod.HttpGet
    .AddUrlSegment "sprintId", requestedSprintId
End With

Dim Jira_Sprint_Response As New JiraResponse
Dim Sprint_Response As New WebResponse
 
Set Sprint_Response = Jira_Sprint_Response.JiraCall(Sprint_Request)
 
funcGetSprintDetails = Sprint_Response.StatusCode
 
If Sprint_Response.StatusCode = OK Then
    With ws_VelocityData
        .Cells(1, 6).Value = "Start"
        .Cells(1, 7).Value = "End"
        .Cells(2, 6).Value = Sprint_Response.Data("startDate")
        .Cells(2, 7).Value = Sprint_Response.Data("endDate")
    End With
    

End If

End Function

Private Function funcGetSprintWorkLog(ByVal teamMembers As String, ByVal startDate As String, ByVal requestedSprintId As String, Optional ByVal endDate As String) As WebStatusCode
 
''
' Source Worklog for Team and Board
'
' @param {String} teamMembers (comma seperated list of team members)
' @param {String} startDate (String for startDate of workLog in Jira date format)
' @param {String} teamMembers (String for endDate of workLog in Jira date format) - optional
' @write {ws_Work}
' @apicalls 2x{get agile/1.0/board/{id}/sprint/{id}/issues}
' @apicalls 1x{get search}
' @apicalls x{get agile/1.0/issue/{id}
' @return {WebStatusCode} status of last successful apicall
''
 
'' Known limitations with this macro:
 
Dim JQLdateRange As String
Dim JQLauthors As String
Dim JQLothers As String
Dim JQLshrink As String

JQLdateRange = "worklogDate >= " & Left(startDate, 10)
       
'If endDate Is Not Null Then
'    JQLdateRange = JQLrange & " AND worklogDate <= " & endDate
'End If

JQLauthors = JQLdateRange & " AND worklogAuthor in (" & teamMembers & ")"
JQLothers = JQLdateRange & " AND worklogAuthor not in (" & teamMembers & ")"
JQLshrink = JQLdateRange & " AND worklogAuthor in (" & teamMembers & ") AND Sprint not in (" & requestedSprintId & ")"
                
'Define the new TeamIssuesForSprint Request
Dim TeamIssuesForSprint_Request As New WebRequest
With TeamIssuesForSprint_Request
    .Resource = "agile/1.0/board/{boardId}/sprint/{sprintId}/issue"
    .Method = WebMethod.HttpGet
    .AddUrlSegment "boardId", rapidViewId
    .AddUrlSegment "sprintId", requestedSprintId
    .AddQuerystringParam "jql", JQLauthors
    .AddQuerystringParam "fields", "worklog"
    .AddQuerystringParam "startAt", 0
    .AddQuerystringParam "maxResults", "1000"
End With

'Define the new GrowthIssuesForSprint Request
Dim GrowthIssuesForSprint_Request As New WebRequest
With GrowthIssuesForSprint_Request
    .Resource = "agile/1.0/board/{boardId}/sprint/{sprintId}/issue"
    .Method = WebMethod.HttpGet
    .AddUrlSegment "boardId", rapidViewId
    .AddUrlSegment "sprintId", requestedSprintId
    .AddQuerystringParam "jql", JQLothers
    .AddQuerystringParam "fields", "worklog"
    .AddQuerystringParam "startAt", 0
    .AddQuerystringParam "maxResults", "1000"
End With
           
'Define the new Search (Shrinkage) Request
Dim Search_Request As New WebRequest
With Search_Request
    .Resource = "api/2/search"
    .Method = WebMethod.HttpGet
    .AddQuerystringParam "jql", JQLshrink
    .AddQuerystringParam "fields", "key"
    .AddQuerystringParam "startAt", 0
    .AddQuerystringParam "maxResults", "1000"
End With
           
Dim Jira_TeamIssuesForSprint_Response As New JiraResponse
Dim Jira_GrowthIssuesForSprint_Response As New JiraResponse
Dim Jira_Search_Response As New JiraResponse

Dim TeamIssuesForSprint_Response As New WebResponse
Dim GrowthIssuesForSprint_Response As New WebResponse
Dim Search_Response As New WebResponse
 
Set TeamIssuesForSprint_Response = Jira_TeamIssuesForSprint_Response.JiraCall(TeamIssuesForSprint_Request)
Set GrowthIssuesForSprint_Response = Jira_GrowthIssuesForSprint_Response.JiraCall(GrowthIssuesForSprint_Request)
Set Search_Response = Jira_Search_Response.JiraCall(Search_Request)
 
If TeamIssuesForSprint_Response.StatusCode = OK Then
    If GrowthIssuesForSprint_Response.StatusCode = OK Then
        funcGetSprintWorkLog = Search_Response.StatusCode
    Else
        funcGetSprintWorkLog = GrowthIssuesForSprint_Response.StatusCode
        Debug.Print "Error with Jira_GrowthIssuesForSprint: " & GrowthIssuesForSprint_Response.StatusCode
        Exit Function
    End If
Else
    funcGetSprintWorkLog = TeamIssuesForSprint_Response.StatusCode
    Debug.Print "Error with Jira_TeamIssuesForSprint: " & TeamIssuesForSprint_Response.StatusCode
    Exit Function
End If

Dim Worklog_Request As New WebRequest
Dim Jira_Worklog_Response As New JiraResponse
Dim Worklog_Response As New WebResponse

Dim i%, w As Integer
Dim r As Long
Dim item As Object
Dim worklog As Object
  
If funcGetSprintWorkLog = OK Then
    r = 2
    i = 1 'reset the issue to 1
    For Each item In TeamIssuesForSprint_Response.Data("issues")
            w = 1 'reset the worklog to 1
            With ws_Work
                .Cells(r, 6).Value = TeamIssuesForSprint_Response.Data("issues")(i)("key")
                    If 1 = 1 Then
                    If TeamIssuesForSprint_Response.Data("issues")(i).Exists("fields") Then
                        If TeamIssuesForSprint_Response.Data("issues")(i)("fields").Exists("worklog") Then
                            If TeamIssuesForSprint_Response.Data("issues")(i)("fields")("worklog").Exists("worklogs") Then
                                For Each worklog In TeamIssuesForSprint_Response.Data("issues")(i)("fields")("worklog")("worklogs")
                                    .Cells(r, 7).Value = TeamIssuesForSprint_Response.Data("issues")(i)("fields")("worklog")("worklogs")(w)("author")("name")
                                    .Cells(r, 8).Value = TeamIssuesForSprint_Response.Data("issues")(i)("fields")("worklog")("worklogs")(w)("started")
                                    .Cells(r, 9).Value = TeamIssuesForSprint_Response.Data("issues")(i)("fields")("worklog")("worklogs")(w)("timeSpentSeconds")
                                    w = w + 1
                                Next worklog
                            End If
                        End If
                    End If
                End If
            End With
            r = r + 1 'increment the row
        i = i + 1 'increment the issue
    Next item
    r = 2 'reset row
    i = 1 'reset the issue to 1
    For Each item In GrowthIssuesForSprint_Response.Data("issues")
            w = 1 'reset the worklog to 1
            With ws_Work
                .Cells(r, 11).Value = GrowthIssuesForSprint_Response.Data("issues")(i)("key")
                If GrowthIssuesForSprint_Response.Data("issues")(i).Exists("fields") Then
                    If GrowthIssuesForSprint_Response.Data("issues")(i)("fields").Exists("worklog") Then
                        If GrowthIssuesForSprint_Response.Data("issues")(i)("fields")("worklog").Exists("worklogs") Then
                            For Each worklog In GrowthIssuesForSprint_Response.Data("issues")(i)("fields")("worklog")("worklogs")
                                .Cells(r, 12).Value = GrowthIssuesForSprint_Response.Data("issues")(i)("fields")("worklog")("worklogs")(w)("author")("name")
                                .Cells(r, 13).Value = GrowthIssuesForSprint_Response.Data("issues")(i)("fields")("worklog")("worklogs")(w)("started")
                                .Cells(r, 14).Value = GrowthIssuesForSprint_Response.Data("issues")(i)("fields")("worklog")("worklogs")(w)("timeSpentSeconds")
                                w = w + 1
                            Next worklog
                        End If
                    End If
                End If
            End With
            r = r + 1 'increment the row
        i = i + 1 'increment the issue
    Next item
    r = 2 'reset row
    i = 1 'reset the issue to 1
    For Each item In Search_Response.Data("issues")
            w = 1 'reset the worklog to 1
            With ws_Work
                              
                'Define the new Worklog (Shrinkage) Request
                With Worklog_Request
                    .Resource = "/agile/1.0/issue/{issueIdOrKey}"
                    .Method = WebMethod.HttpGet
                    .AddUrlSegment "issueIdOrKey", Search_Response.Data("issues")(i)("key")
                    .AddQuerystringParam "fields", "worklog"
                End With
              
                Set Worklog_Response = Jira_Worklog_Response.JiraCall(Worklog_Request)
                 
                If Worklog_Response.StatusCode = OK Then
                    If Worklog_Response.Data.Exists("fields") Then
                        If Worklog_Response.Data("fields").Exists("worklog") Then
                            If Worklog_Response.Data("fields")("worklog").Exists("worklogs") Then
                                For Each worklog In Worklog_Response.Data("fields")("worklog")("worklogs")
                                    .Cells(r, 16).Value = Search_Response.Data("issues")(i)("key")
                                    .Cells(r, 17).Value = Worklog_Response.Data("fields")("worklog")("worklogs")(w)("author")("name")
                                    .Cells(r, 18).Value = Worklog_Response.Data("fields")("worklog")("worklogs")(w)("started")
                                    .Cells(r, 19).Value = Worklog_Response.Data("fields")("worklog")("worklogs")(w)("timeSpentSeconds")
                                    w = w + 1
                                Next worklog
                            End If
                        End If
                    End If
                End If
                
                Set Worklog_Request = Nothing
                
            End With
            r = r + 1 'increment the row
        i = i + 1 'increment the issue
    Next item
Else
    Debug.Print "Error with Jira_Search: " & Search_Response.StatusCode
End If
 
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
    = Total StoryPoints from backlog (excluding Epics) / Average Velocity from last 7 sprints
StoryPointEstimate = Excel.WorksheetFunction.SumIfs(ws_IncompleteIssuesData.Range("F:F"), ws_IncompleteIssuesData.Range("C:C"), "<>Epic") _
                    / Excel.WorksheetFunction.Average(ws_VelocityData.Range("E:E"))
 
''Note: TShirtEstimate is taken to be all Stories with a size of 20 or more. We are not taking Epic estimates into account
 
'TShirtEstimate calculated as _
    = Total StoryPoints from Epics / Average Velocity from last 7 sprints
TShirtEstimate = Excel.WorksheetFunction.SumIfs(ws_IncompleteIssuesData.Range("F:F"), ws_IncompleteIssuesData.Range("C:C"), "Epic") _
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
Private Function funcPredictabilityTiPVariability()
 
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
TiPCol = ws_LeadTimeData.Range("1:1").Find("TiP").Column
If ws_LeadTimeData.Range("A2").Value = 0 Then ' If no LeadTimeData then end function
    Debug.Print "No LeadTimeData for funcPredicatabiltyTiPVariability"
    Exit Function
Else
    TipRow = ws_LeadTimeData.Cells(1, TiPCol).End(xlDown).row
    Set rng = ws_LeadTimeData.Range(Cells(2, TiPCol), Cells(TipRow, TiPCol))
End If

With ws_TeamStats
    .Range("BB27").Value = Excel.WorksheetFunction.StDev_P(rng) / Excel.WorksheetFunction.Average(rng)
    .Range("J16").Value = Excel.WorksheetFunction.StDev_P(rng) / Excel.WorksheetFunction.Average(rng)
End With
 
End Function
Private Function funcPredictabilitySprintOutputVariability()
 
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
Private Function funcResponsivenessLeadTime()
 
'' Update the TeamStats worksheet with the *Lead Time* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''
 
Dim rng_LeadTime As Range
Dim c As Range
 
With ws_LeadTimeData
    .Activate
    If .Range("A2").Value = 0 Then ' If no LeadTimeData then exit the function
        Debug.Print "No LeadTimeData for funcResponsivenessLeadTime"
        Exit Function
    Else
       Set rng_LeadTime = .Range(Cells(2, 9), Cells(.Range("A1").End(xlDown).row, 9))
       For Each c In rng_LeadTime
           c.Value = CDate(.Cells(c.row, 6).Value) - CDate(Left(.Cells(c.row, 4).Value, 10))
       Next c
    End If
End With
 
With ws_TeamStats
    .Range("BB30").Value = WorksheetFunction.Median(rng_LeadTime)
    .Range("AD10").Value = Round(WorksheetFunction.Median(rng_LeadTime), 0)
End With

End Function
Private Function funcResponsivenessDeploymentFrequency()
 
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
       
If ws_LeadTimeData.Range("A2").Value = 0 Then ' If no LeadTimeData then exit the function
    Debug.Print "No LeadTimeData for funcResponsivenessDeploymentFrequency"
    Exit Function
End If

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
Private Function funcResponsivenessTiP()
 
'' Update the TeamStats worksheet with the *TiP* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''
 
Dim rng_TiP As Range
Dim c As Range
 
With ws_LeadTimeData
 
    If .Range("A2").Value = 0 Then ' If no LeadTimeData then exit the function
        Debug.Print "No LeadTimeData for funcResponsivenessTiP"
        Exit Function
    End If
 
    Set rng_TiP = .Range(Cells(2, 10), Cells(.Range("A1").End(xlDown).row, 10))
 
    For Each c In rng_TiP
        c.Value = CDate(.Cells(c.row, 6).Value) - CDate(Left(.Cells(c.row, 5).Value, 10))
    Next c
   
End With
 
With ws_TeamStats
    .Range("BB32").Value = WorksheetFunction.Median(rng_TiP)
    .Range("AD16").Value = Round(WorksheetFunction.Median(rng_TiP), 0)
End With
 
End Function
Private Function funcResponsivenessWiP()
 
'' Update the TeamStats worksheet with the *WiP* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDoneJiras
'
''
Dim startDatesRange As Range, endDatesRange As Range
ws_WiPData.Activate

If ws_WiPData.Range("A2").Value = 0 Then ' If no WiPData then exit the function
    Debug.Print "No WiPData for funcResponsivenessWiP"
    Exit Function
End If

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
    .Range("BB33").Value = WorksheetFunction.Average(ws_WiPData.Range(Cells(2, 7), Cells(startDatesRange.Rows.Count + 1, 7)))
    .Range("AM16").Value = WorksheetFunction.Average(ws_WiPData.Range(Cells(2, 7), Cells(startDatesRange.Rows.Count + 1, 7)))
End With

End Function
Private Function funcProductivityReleaseVelocity()
 
'' Update the TeamStats worksheet with the *Release Velocity* data
'
' Dependent on function: funcGetDoneJiras
'
''
 
Dim c As Range
Dim releaseDateRange As Range
Dim issueTypeRange As Range
Dim currentReleaseDate As Long
Dim countOfReleases As Integer
Dim dict As Dictionary
Set dict = New Dictionary

currentReleaseDate = 0
ws_LeadTimeData.Activate

If ws_LeadTimeData.Range("A2").Value = 0 Then ' If no LeadTimeData then exit the function
    Debug.Print "No LeadTimeData for funcProductivityReleaseVelocity"
    Exit Function
End If


Set releaseDateRange = ws_LeadTimeData.Range(Cells(2, 6), Cells(ws_LeadTimeData.Range("F2").End(xlDown).row, 6))
Set issueTypeRange = ws_LeadTimeData.Range(Cells(2, 3), Cells(ws_LeadTimeData.Range("C2").End(xlDown).row, 3))

For Each c In releaseDateRange
    If Not dict.Exists(c.Value) Then
        dict.Add c.Value, 0
    End If
    If testDate(c) > currentReleaseDate Then
        If testDate(c) < CLng(Date) Then
            currentReleaseDate = testDate(c)
        End If
    End If
Next c

countOfReleases = dict.Count

Dim arr As Variant
Dim i As Integer

With ws_TeamStats
    .Range("AS4").Value = currentReleaseDate
    .Range("AT4:AX4").Value = Array("Feature", "Defects", "Risks", "Debts")
    .Range("AS5").Value = "Velocity"
    .Range("AT5").Value = funcVelocity(Split(vbaJiraProperties.Range("K1").Value, ","), _
                                        issueTypeRange, releaseDateRange, currentReleaseDate) '' Features Velocity
    .Range("AU5").Value = funcVelocity(Split(vbaJiraProperties.Range("K2").Value, ","), _
                                        issueTypeRange, releaseDateRange, currentReleaseDate) '' Defects Velocity
    .Range("AV5").Value = funcVelocity(Split(vbaJiraProperties.Range("K3").Value, ","), _
                                        issueTypeRange, releaseDateRange, currentReleaseDate) '' Risks Velocity
    .Range("AW5").Value = funcVelocity(Split(vbaJiraProperties.Range("K4").Value, ","), _
                                        issueTypeRange, releaseDateRange, currentReleaseDate) '' Debts Velocity
    .Range("AS6").Value = "Baseline"
    .Range("AT6").Value = funcBaselineVelocity(Split(vbaJiraProperties.Range("K1").Value, ","), _
                                        issueTypeRange, countOfReleases) '' Features Baseline
    .Range("AU6").Value = funcBaselineVelocity(Split(vbaJiraProperties.Range("K2").Value, ","), _
                                        issueTypeRange, countOfReleases) '' Defects Baseline
    .Range("AV6").Value = funcBaselineVelocity(Split(vbaJiraProperties.Range("K3").Value, ","), _
                                        issueTypeRange, countOfReleases) '' Risks Baseline
    .Range("AW6").Value = funcBaselineVelocity(Split(vbaJiraProperties.Range("K4").Value, ","), _
                                        issueTypeRange, countOfReleases) '' Debts Baseline
End With
 
End Function

Private Function funcVelocity(ByVal arr As Variant, ByVal issueTypeRange As Range, _
                                                                                ByVal releaseDateRange As Range, _
                                                                                ByVal currentReleaseDate As Long) As Long
    Dim i As Integer

    For i = LBound(arr) To UBound(arr)
        If i = LBound(arr) Then
            funcVelocity = WorksheetFunction.CountIfs(issueTypeRange, arr(i), releaseDateRange, currentReleaseDate)
        Else
            funcVelocity = funcVelocity + WorksheetFunction.CountIfs(issueTypeRange, arr(i), releaseDateRange, currentReleaseDate)
        End If
    Next i

End Function

Private Function funcBaselineVelocity(ByVal arr As Variant, ByVal issueTypeRange As Range, _
                                                                                ByVal countOfReleases As Integer) As Long
    Dim i As Integer
    
    For i = LBound(arr) To UBound(arr)
        If i = LBound(arr) Then
            funcBaselineVelocity = WorksheetFunction.CountIf(issueTypeRange, arr(i))
        Else
            funcBaselineVelocity = funcBaselineVelocity + WorksheetFunction.CountIf(issueTypeRange, arr(i))
        End If
    Next i
    funcBaselineVelocity = funcBaselineVelocity / countOfReleases
    
End Function

Private Function funcProductivityEfficiency()
 
'' Update the TeamStats worksheet with the *Efficiency* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGet3MonthsDoneJiras
'
''

Dim avgActiveTime_ms As Long
Dim avgActiveTime_days As Long
Dim avgLeadTime_days As Long

With ws_LeadTimeData
    If .Range("A2").Value = 0 Then ' If there is no LeadTimeData exit the function
        Debug.Print "No ws_LeadTimeData for funcProductivityEfficiency"
        Exit Function
    End If
    avgActiveTime_ms = Application.WorksheetFunction.Average(.Range("G:G"))
    avgLeadTime_days = Application.WorksheetFunction.Average(.Range("I:I"))
End With

    avgActiveTime_days = avgActiveTime_ms / 3600 / 8
 
With ws_TeamStats
    .Range("BB43").Value = avgActiveTime_days
    .Range("BB35").Value = avgActiveTime_days / avgLeadTime_days
    .Range("AB24").Value = Round(avgActiveTime_days / avgLeadTime_days, 1)
End With
 
End Function
Private Function funcProductivityDistribution()
 
'' Update the TeamStats worksheet with the *Distribution* data
'
' Dependent on function: funcGet12MonthDoneJiras & funcGetIncompleteJiras
'
''
 
Dim arr(1 To 5, 1 To 13) As Long
Dim x%, y As Integer
Dim issueTypeDoneRange As Range
Dim issueTypeBacklogRange As Range
Dim issueType As String
Dim releaseDate As Long
Dim i As Range

If ws_DoneData.Range("C3").Value = 0 Then ' If there is no DoneData exit the function
    Debug.Print "No ws_DoneData for funcProductivityDistribution"
    Exit Function
End If

ws_DoneData.Activate
Set issueTypeDoneRange = ws_DoneData.Range(Cells(2, 3), Cells(ws_DoneData.Range("C2").End(xlDown).row, 3))

If ws_IncompleteIssuesData.Range("C3").Value = 0 Then ' If there is no IncompleteData exit the function
    Debug.Print "No ws_IncompleteIssuesData for funcProductivityDistribution"
    Exit Function
End If

ws_IncompleteIssuesData.Activate
Set issueTypeBacklogRange = ws_IncompleteIssuesData.Range(Cells(2, 3), Cells(ws_IncompleteIssuesData.Range("C2").End(xlDown).row, 3))

For Each i In issueTypeDoneRange
    issueType = i.Value
    If Contains(Split(vbaJiraProperties.Range("K1").Value, ","), issueType) Then ' Feature
        y = 1
    ElseIf Contains(Split(vbaJiraProperties.Range("K2").Value, ","), issueType) Then ' Defect
        y = 2
    ElseIf Contains(Split(vbaJiraProperties.Range("K3").Value, ","), issueType) Then ' Risk
        y = 3
    ElseIf Contains(Split(vbaJiraProperties.Range("K4").Value, ","), issueType) Then 'Debt
        y = 4
    Else
        y = 0
    End If
    releaseDate = i.Offset(0, 1).Value
    Select Case releaseDate
    Case WorksheetFunction.EoMonth(Date, -13) To WorksheetFunction.EoMonth(Date, -12)
        x = 1
    Case WorksheetFunction.EoMonth(Date, -12) To WorksheetFunction.EoMonth(Date, -11)
        x = 2
    Case WorksheetFunction.EoMonth(Date, -11) To WorksheetFunction.EoMonth(Date, -10)
        x = 3
    Case WorksheetFunction.EoMonth(Date, -10) To WorksheetFunction.EoMonth(Date, -9)
        x = 4
    Case WorksheetFunction.EoMonth(Date, -9) To WorksheetFunction.EoMonth(Date, -8)
        x = 5
    Case WorksheetFunction.EoMonth(Date, -8) To WorksheetFunction.EoMonth(Date, -7)
        x = 6
    Case WorksheetFunction.EoMonth(Date, -7) To WorksheetFunction.EoMonth(Date, -6)
        x = 7
    Case WorksheetFunction.EoMonth(Date, -6) To WorksheetFunction.EoMonth(Date, -5)
        x = 8
    Case WorksheetFunction.EoMonth(Date, -5) To WorksheetFunction.EoMonth(Date, -4)
        x = 9
    Case WorksheetFunction.EoMonth(Date, -4) To WorksheetFunction.EoMonth(Date, -3)
        x = 10
    Case WorksheetFunction.EoMonth(Date, -3) To WorksheetFunction.EoMonth(Date, -2)
        x = 11
    Case WorksheetFunction.EoMonth(Date, -2) To WorksheetFunction.EoMonth(Date, -1)
        x = 12
    Case Else
        x = 0
    End Select
    If x > 0 And y > 0 Then arr(y, x) = arr(y, x) + 1
Next i

For Each i In issueTypeBacklogRange
    issueType = i.Value
    If Contains(Split(vbaJiraProperties.Range("K1").Value, ","), issueType) Then ' Feature
        y = 1
    ElseIf Contains(Split(vbaJiraProperties.Range("K2").Value, ","), issueType) Then ' Defect
        y = 2
    ElseIf Contains(Split(vbaJiraProperties.Range("K3").Value, ","), issueType) Then ' Risk
        y = 3
    ElseIf Contains(Split(vbaJiraProperties.Range("K4").Value, ","), issueType) Then 'Debt
        y = 4
    Else
        y = 0
    End If
    x = 13
    If y > 0 Then arr(y, x) = arr(y, x) + 1
Next i

With ws_TeamStats
    .Range("AS11").Value = "Features"
    .Range("AS12").Value = "Defects"
    .Range("AS13").Value = "Risks"
    .Range("AS14").Value = "Debts"
    .Range("AT11:BF14").Value = arr ' Data
    .Range("AT15:BF15").Value = Array(WorksheetFunction.EoMonth(Date, -12), _
                                                            WorksheetFunction.EoMonth(Date, -11), _
                                                            WorksheetFunction.EoMonth(Date, -10), _
                                                            WorksheetFunction.EoMonth(Date, -9), _
                                                            WorksheetFunction.EoMonth(Date, -8), _
                                                            WorksheetFunction.EoMonth(Date, -7), _
                                                            WorksheetFunction.EoMonth(Date, -6), _
                                                            WorksheetFunction.EoMonth(Date, -5), _
                                                            WorksheetFunction.EoMonth(Date, -4), _
                                                            WorksheetFunction.EoMonth(Date, -3), _
                                                            WorksheetFunction.EoMonth(Date, -2), _
                                                            WorksheetFunction.EoMonth(Date, -1), _
                                                            "Backlog") ' Release Months
    .Range("AT16:BE16").NumberFormat = "mmm yy"
End With
 
End Function
Private Function Contains(ByVal arr As Variant, ByVal v As String) As Boolean
Dim rv As Boolean, lb As Long, ub As Long, i As Long
    lb = LBound(arr)
    ub = UBound(arr)
    For i = lb To ub
        If arr(i) = v Then
            rv = True
            Exit For
        End If
    Next i
    Contains = rv
End Function

Private Function funcQualityTimeToResolve()
 
'' Update the TeamStats worksheet with the *Time To Resolve* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGet3MonthsDoneJiras & funcResponsivenessLeadTime
'
''

Dim leadTimeRange As Range
Dim issueTypeRange As Range
Dim nBugs As Integer

If ws_LeadTimeData.Range("A2").Value = 0 Then ' If no LeadTimeData then exit the function
    Debug.Print "No LeadTimeData for funcQualityTimeToResolve"
    Exit Function
End If

ws_LeadTimeData.Activate
Set leadTimeRange = ws_LeadTimeData.Range(Cells(2, 9), Cells(ws_LeadTimeData.Range("F2").End(xlDown).row, 9))
Set issueTypeRange = ws_LeadTimeData.Range(Cells(2, 3), Cells(ws_LeadTimeData.Range("C2").End(xlDown).row, 3))
 
' Make sure there are some Bugs to report
nBugs = WorksheetFunction.CountIf(issueTypeRange, "Bug")
 
With ws_TeamStats
    If nBugs > 0 Then
        .Range("BB37").Value = WorksheetFunction.AverageIf(issueTypeRange, "Bug", leadTimeRange)
        .Range("AM24").Value = Round(WorksheetFunction.AverageIf(issueTypeRange, "Bug", leadTimeRange), 0)
    Else
        .Range("BB37").Value = 0
        .Range("AM24").Value = 0
    End If
End With
 
End Function
Private Function funcQualityDefectDentisy()
 
'' Update the TeamStats worksheet with the *Defect Density* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDefects
'
''

'Count of all issues on the ws_DefectData sheet that have an affects version within the last 12 weeks / the man days over the same period

If ws_DefectData.Range("A2").Value = 0 Then ' If no DefectsData then exit the function
    Debug.Print "No ws_DefectsData for funcQualityDefectDentisy"
    Exit Function
End If
If ws_TeamsData.Range("A2").Value = 0 Then ' If no TeamsData then exit the function
    Debug.Print "No ws_TeamsData for funcQualityDefectDentisy"
    Exit Function
End If

Dim countOfDefects As Long
Dim countOfManDays As Long
Dim TeamId As Long

TeamId = vbaJiraProperties.Range("N1").Value

countOfDefects = Application.WorksheetFunction.CountIf(ws_DefectData.Range("E:E"), ">=" & Now() - (12 * 7))
countOfManDays = Application.WorksheetFunction.SumIf(ws_TeamsData.Range("A:A"), TeamId, ws_TeamsData.Range("E:E")) * 12 / 8
 
With ws_TeamStats
    .Range("BB38").Value = countOfDefects / countOfManDays
    .Range("AM30").Value = Round(countOfDefects / countOfManDays, 1)
End With
 
End Function
Private Function funcQualityFailRate()
 
'' Update the TeamStats worksheet with the *Fail Rate* calcualtion
' The value is added both to the sparkline graph
' Then the to the board for display
'
' Dependent on function: funcGetDefects & funcGet3MonthsOfDoneJiras
'
''

' total number of bugs from the last 3 months of releases / total number of issuesreleased over the same period

Dim totalBugs&, totalPBIs As Long
Dim period As Long
 
period = DateAdd("m", -3, DateSerial(Year(Date), Month(Date), 1))
 
totalBugs = WorksheetFunction.CountIf(ws_DefectData.Range("E:E"), ">" & period)
totalPBIs = WorksheetFunction.CountIf(ws_LeadTimeData.Range("F:F"), ">" & period)

With ws_TeamStats
    .Range("BB39").Value = totalBugs / totalPBIs
    .Range("AM36").Value = Round(totalBugs / totalPBIs, 2)
End With
 
End Function
Private Function funcScrumTeamStability()
 
'' Update the TeamStats worksheet with the Team Stability
' The value is added both to the sparkline graph
' Then to the board for display
'
' This is an indication of the team's stability. For example, given:
' If User Portfolio the get the baseline form the Teams data (otherwise we could use the previous period), i.e:
'   George: 100% dedicated (based on hours allocated per week in portfolio - default 40 - 100%)
'   Joe: 50%
'   Jen: 80%
' Check current reporting period - i.e. most recent completed sprint by searching for time booked to the issues on the team board with:
' Get agile/1.0/board/{id}/sprint/{id}/issue - which gets all issues from the sprint, including their worklog
' AND Jira Search jql=worklogDate =>date (start of sprint) and worklogDate =< date (end of sprint) and worklogAuthor in (team members)
' i.e.:
'   George: 85% (-15% delta)
'   Jen: 100% (+20%)
'   Jeff: 25% (new) (+25%)
'   Joe: missing (-50)
' Calculate the TeamGrowth for the team as .2 (Jen) + .25 (Jeff) = .45 divided by the current team size (2) or 22.5%.
' Calculate the TeamShrinkage for the team would be |-.15| (George) + |-.5| (Joe) = .65 divided by the old team size (2.7) or 24.07%.
' The total volatility would be the sum of the two prior metrics or 46.57% and Team Stability would be 100 - 46.57/2 = 76.715%
''
 
'Deployment Frequency calculated as _
    = Number of unique releaseDates in the previous month from today
 
Dim Current_Team_Size As Long
Dim Old_Team_Size As Long
Dim Team_Growth As Long
Dim Team_Shrinkage As Long
       
Current_Team_Size = Application.WorksheetFunction.Sum(ws_Work.Range("I:I")) ' sum of seconds worked
Old_Team_Size = Application.WorksheetFunction.Sum(ws_TeamsData.Range("E:E")) * 60 * 60 ' sum of hours expected x 60 minutes x 60 seconds

Team_Growth = Application.WorksheetFunction.Sum(ws_Work.Range("N:N")) / Current_Team_Size ' sum of growth seconds / sum of seconds worked
Team_Shrinkage = Application.WorksheetFunction.Sum(ws_Work.Range("S:S")) / Old_Team_Size ' sum of shrinkage seconds / sum of seconds expected
 
With ws_TeamStats
    .Range("BB22").Value = 1 - ((Team_Growth + Team_Shrinkage) / 2)
    .Range("J5").Value = Round(1 - ((Team_Growth + Team_Shrinkage) / 2), 2)
End With
 
End Function
Private Function funcScrumUnplannedWork()
 
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
    .Range("BB42").Value = Application.WorksheetFunction.Sum(ws_Work.Range("C:C")) ' Total Active Time
End With
 
End Function

Private Function funcJiraAdminActiveTime()
 
'' Update the TeamStats worksheet with the *Active Time %* calcualtion
' The value is added both to the sparkline graph
' Then to the board for display
'
' Dependent on function: funcGetSprintBurnDown & funcPostTeamsFind
'
''

' Known Limitiations:
' (1) need to add start date, end date and holiday accountability to the funcPostTeamsFind and incorporate that in here

'Sum the timeSpent/3600 to get the number of hours spent then /8 to get the number of days
'Then / (number of days in sprint * number of hours allocated per week/ 5))
 
Dim daysLogged As Long
daysLogged = WorksheetFunction.Sum(ws_Work.Range("C:C")) / 3600 / 8
Dim daysAllocated As Long
daysAllocated = DaysInSprint * WorksheetFunction.SumIf(ws_TeamsData.Range("A:A"), CLng(TeamId), ws_TeamsData.Range("E:E")) / 5
 
With ws_TeamStats
    .Range("BB41").Value = daysLogged / daysAllocated
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
 
'' This function updates the As Of Date and Team Name displayed on the Dashboard
 
With ws_TeamStats
    .Range("B2").Value = vbaJiraProperties.Range("N2") & " - As Of: " & Format(Now(), "dd-mmm-yy")
End With
 
End Function
Private Function clearOldData(ByVal ws As Worksheet)
 
Dim rngOldData As Range
    With ws
        Set rngOldData = .Range("A1").CurrentRegion ' Define the extend of the data range
        If rngOldData.Rows.Count > 1 Then ' First row is the headings soonly reset if more than one row
            Set rngOldData = rngOldData.Resize(rngOldData.Rows.Count - 1).Offset(1) ' Remove the headings from the range
            rngOldData.ClearContents ' clear existing data
        End If
    End With
   
End Function
Private Function CreateWorkSheet(ByVal name As String, Optional ByRef headings As Variant) As Worksheet
 
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
Private Function ws_TeamStats() As Worksheet
    Set ws_TeamStats = CreateWorkSheet("Team Metrics")
End Function
Private Function ws_LeadTimeData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "issueType", "createdDate", "sprintStartDate", "releaseDate", "totalTime", "totalString", "leadTime", "TiP")
    Set ws_LeadTimeData = CreateWorkSheet("ws_LeadTimeData", HeadingsArr)
End Function
Private Function ws_WiPData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "issueType", "inProgressDate", "endProgressDate", "releaseDate", "WiP")
    Set ws_WiPData = CreateWorkSheet("ws_WiPData", HeadingsArr)
End Function
Private Function ws_IncompleteIssuesData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "issueType", "project", "epicKey", "storyPoints", "status", "statusCategory", "aggregateTimeEstimate", "sprintState")
    Set ws_IncompleteIssuesData = CreateWorkSheet("ws_IncompleteIssuesData", HeadingsArr)
End Function
Private Function ws_VelocityData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("SprintId", "SprintName", "State", "Committed", "Completed")
   Set ws_VelocityData = CreateWorkSheet("ws_VelocityData", HeadingsArr)
End Function
Private Function ws_TeamsData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "title", "resourceId", "personId", "weeklyHours", "personId2", "JiraUserName")
    Set ws_TeamsData = CreateWorkSheet("ws_TeamsData", HeadingsArr)
End Function
Private Function ws_Work() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("burndown:key", "burndown:rapidViewId", "burndown:timeSpent", , , "board&team:key", "board&team:author", "board&team:date", "board&team:timeSpent", , "growth:key", "growth:author", "growth:date", "growth:timeSpent", , "shrinkage:key", "shrinkage:author", "shrinkage:date", "shrinkage:timeSpent")
    Set ws_Work = CreateWorkSheet("ws_Work", HeadingsArr)
End Function
Private Function ws_ProjectData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "projectName", "category")
    Set ws_ProjectData = CreateWorkSheet("ws_ProjectData", HeadingsArr)
End Function
Private Function ws_DoneData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "issueType", "releaseDate")
    Set ws_DoneData = CreateWorkSheet("ws_DoneData", HeadingsArr)
End Function
Private Function ws_DefectData() As Worksheet
    Dim HeadingsArr As Variant
    HeadingsArr = Array("id", "key", "issueType", "affectsVersion", "affectsVersionReleaseDate")
    Set ws_DefectData = CreateWorkSheet("ws_DefectData", HeadingsArr)
End Function
Private Function ArrayOfDates(ByVal startDate As Long, ByVal endDate As Long) As Variant()

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
Private Function MinMaxDate(ByVal dateRange As Range, ByVal MType As String) As Variant
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
Private Function inProgressForDate(ByVal startDate As Long, ByVal endDate As Long, currentDate As Long) As Integer
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
Private Function WiP(ByVal row As Long) As Integer
    
    Dim headers As Range
    Set headers = ws_WiPData.Range(Cells(1, 8), Cells(1, ws_WiPData.Range("H1").End(xlToRight).Column))

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
Private Function testDate(ByVal cell As Range) As Long
    ''returns the excel date as a Long from a cell value that is formated as a test string 'i.e. "30/05/2020" -->43981
    testDate = CLng(DateValue(cell.Value))
End Function
Private Function issueTypeSearchString() As String
    Dim strFeatures$, strDefects$, strRisks$, strDebts As String
     
    '' Get the individual issue type strings
    strFeatures = vbaJiraProperties.Range("K1").Value
    strDefects = vbaJiraProperties.Range("K2").Value
    strRisks = vbaJiraProperties.Range("K3").Value
    strDebts = vbaJiraProperties.Range("K4").Value

    '' Join all the issue type strings together
    issueTypeSearchString = strFeatures & "," & strDefects & "," & strRisks & "," & strDebts
    
    '' Remove the duplicates using the dictionary method
    Dim arr As Variant
    arr = Split(issueTypeSearchString, ",")
    Dim i As Long
    Dim d As Dictionary
    Set d = New Dictionary
    With d
        For i = LBound(arr) To UBound(arr)
            If Not d.Exists(arr(i)) Then
                d.Add arr(i), 1
                If d.Count = 1 Then
                    issueTypeSearchString = arr(i)
                Else
                    issueTypeSearchString = issueTypeSearchString & "," & arr(i)
                End If
            End If
        Next
    End With

End Function
Private Function TeamResourcesString() As String

Dim rg As Range
Set rg = ws_TeamsData.Range("G2").CurrentRegion
Set rg = rg.Resize(rg.Rows.Count - 1, 1).Offset(1, 6)

TeamResourcesString = Join(WorksheetFunction.Transpose(rg.Value), ",")

End Function

