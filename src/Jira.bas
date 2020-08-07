Attribute VB_Name = "Jira"
''
' Jira v0.1
' (c) Ben Doughton - https://github.com/bdoughton/vba-jira
'
' Includes:
'
' - CheckJiraBaseUrl - validates the format of the url passed to it
' - JiraBaseUrl - returns the stored base url or requests one if not stored in properties
' - UpdateBaseUrl - updates a new base url from a Ribbon X control
' - IsLoggedIn - checks if a username and password are stored in properties
' - GetJiraLoginResponse - get request to Jira to validate user login redentials and stores them in properties if valid
' - UpdateUser - run the LoginUser from a Ribbon X control
' - LoginUser - requests login details and calls the GetJiraLoginResponse
'
' @module Jira
' @author bdoughton@me.com
' @license GNU General Public License v3.0 (https://opensource.org/licenses/GPL-3.0)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Public userName As String
Public userPassword As String
Function CheckJiraBaseUrl(ByVal url As String) As Boolean

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
        If Not CheckJiraBaseUrl(InputBox("Please enter the Jira Base Url:")) Then
            Exit Function
        End If
    End If
    
    JiraBaseUrl = vbaJiraProperties.Range("B1").Value
    
End Function
Sub UpdateBaseUrl(control As IRibbonControl)
    CheckJiraBaseUrl (InputBox("Please enter the new Jira Base Url:"))
End Sub
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
        userName = InputBox("User Name:")
        If userName = "" Then
            LoginUser = False
            Exit Function
        End If
        
        userPassword = InputBox("Password:")
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

