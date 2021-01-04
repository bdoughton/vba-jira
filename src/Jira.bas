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

''Commenting these 3 out for now as I don't believe they are used anymore : 19-Dec-2020
'Public Const ExternalIssueID As String = "customfield_10116"
'Public Const OriginalDueDate As String = "customfield_11604"
'Public Const AccountableDepartment As String = "customfield_11700"

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
    
    Dim oldStatusBar As Boolean
    oldStatusBar = Application.DisplayStatusBar

    Dim defaultUrl As String
    If vbaJiraProperties.Range("B1").Value = "" Then
        defaultUrl = "http://localhost:8080/rest/"
    Else
        defaultUrl = vbaJiraProperties.Range("B1").Value
    End If
    CheckJiraBaseUrl (InputBox("Please enter the new Jira Base Url:", "Jira", defaultUrl))
    
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    
End Sub
Function CheckRapidBoard(ByVal Id As String) As Boolean

    If StrPtr(Id) = 0 Then 'the user pressed "cancel"
        Exit Function
    End If

    If GetJiraBoardResponse(Id).StatusCode = Ok Then
        CheckRapidBoard = True
        MsgBox ("Successfully found board : " & Id & " - " & GetJiraBoardResponse(Id).StatusCode) ' Duplicate call remove this
    Else
        CheckRapidBoard = False
        MsgBox ("Could not find board : " & GetJiraBoardResponse(Id).StatusCode) ' Duplicate call remove this
    End If

End Function
Function rapidViewId() As String

    If vbaJiraProperties.Range("E1").Value = "" Then
        If Not CheckRapidBoard(InputBox("Please enter the RapidBoardId:", "Jira")) Then
            Exit Function
        End If
    End If
    rapidViewId = vbaJiraProperties.Range("E1").Value
End Function
Sub UpdateBoardId(control As IRibbonControl)

    Dim oldStatusBar As Boolean
    oldStatusBar = Application.DisplayStatusBar

    CheckRapidBoard (InputBox("Please enter the RapidBoardId:", "Jira", vbaJiraProperties.Range("E1").Value))

    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar

End Sub
Sub UpdateCustomFields(control As IRibbonControl)

    Dim oldStatusBar As Boolean
    oldStatusBar = Application.DisplayStatusBar
    
    ' First check that a RapidBoard exists - used as a proxy to ensure that JiraAgile is installed
    If vbaJiraProperties.Range("E1").Value = "" Then
        If Not CheckRapidBoard(InputBox("Please enter the RapidBoardId:", "Jira")) Then
            Exit Sub
        End If
    End If
    
    ' Store the CustomFields from Jira Agile
    If GetJiraCustomFieldResponse("JiraAgile").StatusCode = Ok Then
        MsgBox ("Successfully Updated Jira Agile Custom Fields")
    Else
        MsgBox ("Error retreiving Jira Agile Custom Fields")
    End If
    
    ' Store the CustomFields from Advanced Roadmaps
    If True Then ' To be updated with a check to see if the Advanced Roadmaps plugin is installed
        If GetJiraCustomFieldResponse("AdvancedRoadmaps").StatusCode = Ok Then
            MsgBox ("Successfully Updated Jira Advanced Roadmaps Custom Fields")
        Else
            MsgBox ("Error retreiving Jira Advanced Roadmaps Custom Fields")
        End If
    End If
    
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar
    
End Sub
Function GetJiraBoardResponse(ByVal Id As String) As WebResponse
         
    ' Create a WebRequest to get the logged in user's details
    Dim BoardRequest As New WebRequest

    BoardRequest.Resource = "greenhopper/1.0/rapidviewconfig/editmodel.json"
    BoardRequest.AddQuerystringParam "rapidViewId", Id
    BoardRequest.Method = WebMethod.HttpGet

    Dim JiraBoardResponse As New JiraResponse
    ' Execute the request and work with the response
    Set GetJiraBoardResponse = JiraBoardResponse.JiraCall(BoardRequest)
    
    If GetJiraBoardResponse.StatusCode = Ok Then
        vbaJiraProperties.Range("D1:E1").Value = Array("boardId:", GetJiraBoardResponse.Data("id"))
        vbaJiraProperties.Range("D2:E2").Value = Array("boardName:", GetJiraBoardResponse.Data("name"))
        vbaJiraProperties.Range("D3:E3").Value = Array("boardJql:", GetJiraBoardResponse.Data("filterConfig")("query"))
        vbaJiraProperties.Range("D4:E4").Value = Array("boardTimeTrackingEnabled:", GetJiraBoardResponse.Data("estimationStatisticConfig")("currentTrackingStatistic")("isEnabled"))
    End If
End Function
Function GetJiraCustomFieldResponse(ByVal PlugIn As String) As WebResponse
         
    ' Create a WebRequest to get the logged in user's details
    Dim CustomFieldRequest As New WebRequest
    CustomFieldRequest.Resource = "api/2/field"
    CustomFieldRequest.Method = WebMethod.HttpGet

    Dim JiraCustomFieldResponse As New JiraResponse
    ' Execute the request and work with the response
    Set GetJiraCustomFieldResponse = JiraCustomFieldResponse.JiraCall(CustomFieldRequest)
    
    If GetJiraCustomFieldResponse.StatusCode = Ok Then
    Select Case PlugIn
            Case "JiraAgile"
                vbaJiraProperties.Range("G1:H1").Value = Array("sprints:", find_customfield_id("Sprint", "com.pyxis.greenhopper.jira:gh-sprint", GetJiraCustomFieldResponse.Data))
                vbaJiraProperties.Range("G2:H2").Value = Array("epiclink:", find_customfield_id("Epic Link", "com.pyxis.greenhopper.jira:gh-epic-link", GetJiraCustomFieldResponse.Data))
                vbaJiraProperties.Range("G3:H3").Value = Array("storypoints:", find_customfield_id("Story Points", "com.atlassian.jira.plugin.system.customfieldtypes:float", GetJiraCustomFieldResponse.Data))
                vbaJiraProperties.Range("G4:H4").Value = Array("epiccolour:", find_customfield_id("Epic Colour", "com.pyxis.greenhopper.jira:gh-epic-color", GetJiraCustomFieldResponse.Data))
            Case "AdvancedRoadmaps"
                vbaJiraProperties.Range("G5:H5").Value = Array("parentlink:", find_customfield_id("Parent Link", "com.atlassian.jpo:jpo-custom-field-parent", GetJiraCustomFieldResponse.Data))
                vbaJiraProperties.Range("G6:H6").Value = Array("Team:", find_customfield_id("Team", "com.atlassian.teams:rm-teams-custom-field-team", GetJiraCustomFieldResponse.Data))
        End Select
    End If
End Function

Function find_customfield_id(ByVal name As String, ByVal custom As String, ByVal Data As Object) As String
'' Search through the Data object and find the customfield_id that corresponds to the correct custom field
    Dim field As Object
    Dim i As Integer
    i = 1
    For Each field In Data
        If Data(i)("name") = name Then
            If Data(i)("schema")("custom") = custom Then ' then validate the custom schema type to ensure accuracy in case two or more fields have the same name
                find_customfield_id = Data(i)("id")
                Exit Function
            End If
        End If
        i = i + 1
    Next field
End Function

' CustomFields from Jira Agile
Public Function sprints() As String
    sprints = vbaJiraProperties.Range("H1").Value
End Function

Public Function epiclink() As String
    epiclink = vbaJiraProperties.Range("H2").Value
End Function

Public Function storypoints() As String
    storypoints = vbaJiraProperties.Range("H3").Value
End Function

Public Function epiccolour() As String
    epiccolour = vbaJiraProperties.Range("H4").Value
End Function

' CustomFields from Advanced Roadmaps (formely Portfolio)
Public Function parentlink() As String
    ''parentlink = "customfield_11801"
    parentlink = vbaJiraProperties.Range("H5").Value
End Function

Public Function Team() As String 'Custom fieldname for Team
    'i.e. Team = "customfield_11800"
    Team = vbaJiraProperties.Range("H6").Value
End Function

' Values from Jira Agile

Public Function boardJql() As String
    'i.e. boardJql = "Team = 81 AND CATEGORY = calm AND NOT issuetype in (Initiative) ORDER BY Rank ASC"
    boardJql = vbaJiraProperties.Range("E3").Value
End Function

Public Function bln_TimeTrackingEnabled() As Boolean
    bln_TimeTrackingEnabled = vbaJiraProperties.Range("E4").Value
End Function

' Values from Advanced Roadmaps (formely Portfolio)

Public Function teamId() As String 'Field value for Team
    If vbaJiraProperties.Range("N1").Value = "" Then
        teamId = InputBox("Please enter the Advanced Roadmaps teamId:", "Jira", vbaJiraProperties.Range("N1").Value)
    Else
        teamId = vbaJiraProperties.Range("N1").Value
    End If
End Function

Sub UpdateTeamId(control As IRibbonControl)

    If GetJiraTeamResponse(InputBox("Please enter the Advanced Roadmaps teamId:", "Jira", vbaJiraProperties.Range("N1").Value)).StatusCode = Ok Then
        MsgBox ("Successfully Found Team")
    Else
        MsgBox ("Error Finding Team")
    End If

End Sub

Function GetJiraTeamResponse(ByVal Id As String) As WebResponse
         
    ' Create a WebRequest to get the logged in user's details
    Dim TeamRequest As New WebRequest

    TeamRequest.Resource = "teams-api/1.0/team/{id}"
    TeamRequest.AddUrlSegment "id", Id
    TeamRequest.Method = WebMethod.HttpGet

    Dim JiraTeamResponse As New JiraResponse
    ' Execute the request and work with the response
    Set GetJiraTeamResponse = JiraTeamResponse.JiraCall(TeamRequest)
    
    If GetJiraTeamResponse.StatusCode = Ok Then
        vbaJiraProperties.Range("M1:N1").Value = Array("teamId:", GetJiraTeamResponse.Data("id"))
        vbaJiraProperties.Range("M2:N2").Value = Array("teamName:", GetJiraTeamResponse.Data("title"))
        If Not GetJiraTeamResponse.Data("shareable") Then
            MsgBox ("Please ensure you share the team: " & GetJiraTeamResponse.Data("title") & " before you run Get Stats From Jira")
        End If
    End If
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

    WebHelpers.EnableLogging = True

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
Sub ClearProperties(control As IRibbonControl)
    vbaJiraProperties.Cells.ClearContents
End Sub
Sub ExportProperties(control As IRibbonControl)
    vbaJiraProperties.Copy
End Sub
Sub UpdateUser(control As IRibbonControl)

    Dim oldStatusBar As Boolean
    oldStatusBar = Application.DisplayStatusBar

    LoginUser
    
    Application.StatusBar = False
    Application.DisplayStatusBar = oldStatusBar

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


