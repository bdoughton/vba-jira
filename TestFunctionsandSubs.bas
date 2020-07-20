Attribute VB_Name = "TestFunctionsandSubs"
Option Explicit
Sub unhidesheets()
Dim ws As Worksheet

For Each ws In Worksheets
    ws.Visible = xlSheetVisible
Next ws

End Sub

Sub MsgTest()

MsgBox CDate(Range("F2")) - CDate(Left(Range("D2").value, 10))

End Sub
Sub TestTeamStats()

Dim callResult As Long
Dim url As String
Dim l As Integer
Dim boardJql As String


'Fetch the base url
url = PublicVariables.JiraBaseUrl
boardJql = "Team = 81 AND CATEGORY = calm AND NOT issuetype in (Initiative) ORDER BY Rank ASC"

' Allow the user to confirm the base url before executing the code
If MsgBox("You are using this Jira url: " & url, vbOKCancel) = vbCancel Then
    msgCaption = "Cancelled action"
    Exit Sub
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

'Fetch Data from Api
'callResult = TeamStats.funcGetDoneJiras(encodedAuth, url, boardJql, "In Progress", "Done", 0, 2)
callResult = TeamStats.funcGetIncompleteJiras(encodedAuth, url, boardJql, 0, 2)

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

Sub TestDictionary()

 Dim coll As New Collection 'Issue Key
 Dim dict As Dictionary 'Resource Name
 Dim c As Range
 Dim r As Integer
       
 
 With ws_LeadTimeData
    For r = 2 To 30
       Set dict = New Dictionary
       For Each c In .Range("I" & r & ":L" & r)
            dict(.Cells(1, c.Column).value) = c.value
       Next c
       coll.Add dict, .Cells(r, 2).value
    Next
 End With
  
    Debug.Print (Now() & " | Result: " & coll("CALM60691-43")("sanchit.gupta"))

End Sub


Sub Test()

Dim callResult As Long
Dim url As String
Dim l As Integer

'Fetch the base url
url = PublicVariables.JiraBaseUrl

' Allow the user to confirm the base url before executing the code
If MsgBox("You are using this Jira url: " & url, vbOKCancel) = vbCancel Then
    msgCaption = "Cancelled action"
    Exit Sub
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

'Fetch Data from Api
callResult = RestApiCalls.GetEpicsForBoard(encodedAuth, url, ws_BoardData.Cells(5, 2).value, ws_EpicsData, 10)

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

Sub Test2()
Dim apicall As String
Dim boardJql As String
Dim jql As String

'' The following can be added to speed up the api query but may make older release distribution less accurate
'' "fixversion changed after -78w AND "

boardJql = "Team = 81 AND CATEGORY = calm AND NOT issuetype in (Initiative) ORDER BY Rank ASC"

jql = "fixVersion is not EMPTY AND " & _
        "Sprint is not EMPTY AND " & _
        "NOT issuetype in (Initiative,Epic,Test,subTaskIssueTypes()) AND " & _
        "statusCategory in (Done) AND " & _
        boardJql

apicall = _
    "http://jira.unix.lch.com:8080/rest/api/latest/search?jql=" _
    & jql _
    & "&fields=" _
        & "key," _
        & "issuetype," _
        & "fixVersions," _
        & sprints & "," _
        & "created," _
        & "changelog" _
        & "&startAt=" _
        & 0 _
    & "&maxResults=1000&expand=changelog"


'        & "project," _
'        & parentlink & "," _
'        & epiclink & "," _
'        & epiccolour & "," _
'        & "id," _
'        & "summary," _
'        & "assignee," _
'        & "creator," _
'        & "status," _
'        & Team & "," _
'        & storypoints & "," _
'
'        & "created," _
'        & "aggregatetimespent," _


Debug.Print apicall

End Sub

