Attribute VB_Name = "Jira"
''
' Jira v1.0
' (c) Ben Doughton - https://github.com/bdoughton/vba-jira
'
' @Includes:
'               CheckJiraBaseUrl         - validates the format of the url passed to it
'               JiraBaseUrl                   - returns the stored base url or requests one if not stored in properties
'               UpdateBaseUrl             - updates a new base url from a Ribbon X control
'               CheckJRapidBoard         - validates the board Id passed to it exists
'               rapidViewId                   - returns the stored rapidViewId or requests one if not stored in properties
'               UpdateBoardId             - updates the BoardId from a Ribbon X control
'               GetJiraBoardResponse    - calls the Jira get agile/1.0/board api used to validate the rapidViewId
'               teamId                         - returns the stored teamId or requests one if not stored in properties # placeholder to be updated
'               boardJql                       - returns the stored boardJql or requests one if not stored in properties # placeholder to be updated
'               IsLoggedIn                   - checks if a username and password are stored in properties
'               GetJiraLoginResponse  - get request to Jira to validate user login redentials and stores them in properties if valid
'               UpdateUser                  - run the LoginUser from a Ribbon X control
'               LoginUser                    - requests login details and calls the GetJiraLoginResponse
'               date2epoch                 - converts a date to an epoch value (as used by Jira)
'               epoch2date                 - converts an epoch value (as used by Jira) to a date
'
' @Dependencies:
'               Worksheet - vbaJiraProperties
'               Mod - InputPassword
'               Mod - WebHelpers
'               Class - WebClient
'               Class - WebRequest
'               Class - WebResponse
'               Class - HttpBasicAuthenticator
'               RibbonX - Jira
'
' @module Jira
' @author bdoughton@me.com
' @license GNU General Public License v3.0 (https://opensource.org/licenses/GPL-3.0)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

Option Explicit

''Jira CustomFields
Public Const ExternalIssueID As String = "customfield_10116"
Public Const OriginalDueDate As String = "customfield_11604"
Public Const AccountableDepartment As String = "customfield_11700"
Public Const sprints As String = "customfield_10004"
Public Const parentlink As String = "customfield_11801"
Public Const epiclink As String = "customfield_10006"
Public Const Team As String = "customfield_11800"
Public Const storypoints As String = "customfield_10002"
Public Const epiccolour As String = "customfield_10008"

Public userName As String
Public userPassword As String

Function CheckJiraBaseUrl(ByVal url As String) As Boolean

    If StrPtr(url) = 0 Then 'the user pressed "cancel"
        Exit Function
    End If

    If url Like "http*://*/rest/" Then
        CheckJiraBaseUrl = True
        vbaJiraProperties.Range("A1:B1").Value = Array("JiraBaseUrl:", url)
    Else
        CheckJiraBaseUrl = False
        MsgBox ("Supplied url appears invalid! Must be in the format: http*://*/rest/")
    End If

End Function
Function JiraBaseUrl() As String
    
' If the JiraBaseUrl property is "" then request the user to enter a valid base url

    If vbaJiraProperties.Range("B1").Value = "" Then
        If Not CheckJiraBaseUrl(InputBox("Please enter the Jira Base Url:", "Jira", "http://localhost:8080/rest/")) Then
            Exit Function
        End If
    End If
    
    JiraBaseUrl = vbaJiraProperties.Range("B1").Value
    
End Function
Sub UpdateBaseUrl(control As IRibbonControl)
    Dim defaultUrl As String
    If vbaJiraProperties.Range("B1").Value = "" Then
        defaultUrl = "http://localhost:8080/rest/"
    Else
        defaultUrl = vbaJiraProperties.Range("B1").Value
    End If
    CheckJiraBaseUrl (InputBox("Please enter the new Jira Base Url:", "Jira", defaultUrl))
End Sub
Function CheckRapidBoard(ByVal Id As String) As Boolean

    If StrPtr(Id) = 0 Then 'the user pressed "cancel"
        Exit Function
    End If

    If GetJiraBoardResponse(Id).StatusCode = Ok Then
        CheckRapidBoard = True
        vbaJiraProperties.Range("A2:B2").Value = Array("RapidBoardId:", Id)
        MsgBox ("Successfully found board : " & Id & " - " & GetJiraBoardResponse(Id).StatusCode)
    Else
        CheckRapidBoard = False
        MsgBox ("Could not find board : " & GetJiraBoardResponse(Id).StatusCode)
    End If

End Function
Function rapidViewId() As String

    If vbaJiraProperties.Range("B2").Value = "" Then
        If Not CheckRapidBoard(InputBox("Please enter the RapidBoardId:", "Jira")) Then
            Exit Function
        End If
    End If
    
    rapidViewId = vbaJiraProperties.Range("B2").Value
End Function
Sub UpdateBoardId(control As IRibbonControl)
    CheckRapidBoard (InputBox("Please enter the RapidBoardId:", "Jira", vbaJiraProperties.Range("B2").Value))
End Sub
Function GetJiraBoardResponse(ByVal Id As String) As WebResponse
         
    ' Create a WebRequest to get the logged in user's details
    Dim BoardRequest As New WebRequest
    BoardRequest.Resource = "agile/1.0/board/{boardId}"
    ' Replace {boardId} segment
    BoardRequest.AddUrlSegment "boardId", Id
    BoardRequest.Method = WebMethod.HttpGet

    Dim JiraBoardResponse As New JiraResponse
    ' Execute the request and work with the response
    Set GetJiraBoardResponse = JiraBoardResponse.JiraCall(BoardRequest)
    
End Function

Public Function teamId() As String
''Placeholder to define other values
    teamId = "81"
End Function
Public Function boardJql() As String
''Placeholder to define other values
    boardJql = "Team = 81 AND CATEGORY = calm AND NOT issuetype in (Initiative) ORDER BY Rank ASC"
End Function

Function IsLoggedIn() As Boolean

    If userName <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
        Exit Function
    End If
    If userPassword = "" Then
        IsLoggedIn = False
    End If

End Function
Function GetJiraLoginResponse(ByVal name As String, ByVal password As String) As WebResponse
   
    Dim JiraClient As New WebClient
    JiraClient.baseUrl = JiraBaseUrl
    
    'Setup Authentication
    Dim JiraAuth As New HttpBasicAuthenticator
    JiraAuth.Setup _
        userName:=name, _
        password:=password
    
    Set JiraClient.Authenticator = JiraAuth
    
    ' Create a WebRequest to get the logged in user's details
    Dim LoginRequest As New WebRequest
    LoginRequest.Resource = "api/2/myself"
    LoginRequest.Method = WebMethod.HttpGet

    ' Set the request format
    ' -> Sets content-type and accept headers and parses the response
    LoginRequest.ContentType = "application/json;charset=UTF-8"

    ' Execute the request and work with the response
    Set GetJiraLoginResponse = JiraClient.Execute(LoginRequest)
    
    If GetJiraLoginResponse.StatusCode = Ok Then
        userName = name
        userPassword = password
        MsgBox ("Successfully logged in : " & userName)
    Else
        userName = ""
        userPassword = ""
        MsgBox ("Could not login : " & GetJiraLoginResponse.StatusCode)
    End If
    
End Function
Sub UpdateUser(control As IRibbonControl)
    LoginUser
End Sub
Function LoginUser() As Boolean

        'Request login details
        userName = InputBox("Please enter your Jira user name:", "User Name")
        If userName = "" Then
            LoginUser = False
            Exit Function
        End If
        
        #If Mac Then
            userPassword = InputBox("Mac Password:", "Password")
        #Else
            userPassword = InputBoxDK("Please enter your Jira password:", "Password")
        #End If
        
        If userPassword = "" Then
            LoginUser = False
            Exit Function
        End If
        
        'attempt to a login with the credentials provided
        If GetJiraLoginResponse(userName, userPassword).StatusCode = Ok Then
            LoginUser = True
        Else
            LoginUser = False
        End If
    
End Function
Function jiratime(ByVal timeseries As Double) As String

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
' result is GMT/UTC
Function epoch2date(myEpoch)
epoch2date = DateAdd("s", myEpoch / 1000, "01/01/1970 00:00:00")
End Function
Function date2epoch(myDate)
date2epoch = DateDiff("s", "01/01/1970 00:00:00", myDate) * 1000
End Function

