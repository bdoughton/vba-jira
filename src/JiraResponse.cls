VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JiraResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' JiraResponse v1
' (c) Ben Doughton - https://github.com/bdoughton/vba-jira
'
' ```
'
'
' @class JiraResponse
' @author bdoughton@me.com
' @license GNU General Public License v3.0 (https://opensource.org/licenses/GPL-3.0)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' ============================================= '
' Public Methods
' ============================================= '

''
' @method GetJira
' @param Updated {WebResponse} Updated `WebResponse` to pull updated values from
''
Public Function GetJira(ByVal Resource As String) As WebResponse
       
    Dim JiraClient As New WebClient
    JiraClient.baseUrl = JiraBaseUrl
    
    'Setup Authentication
    Dim JiraAuth As New HttpBasicAuthenticator
    JiraAuth.Setup _
        userName:=userName, _
        password:=userPassword
    
    Set JiraClient.Authenticator = JiraAuth
    
    ' Create a WebRequest for getting the board details
    Dim GetJiraRequest As New WebRequest
    GetJiraRequest.Resource = Resource
    GetJiraRequest.Method = WebMethod.HttpGet

    ' Set the request format
    ' -> Sets content-type and accept headers and parses the response
    GetJiraRequest.ContentType = "application/json;charset=UTF-8"

    ' Execute the request and return the response
    Set GetJira = JiraClient.Execute(GetJiraRequest)
    
End Function

Public Function JiraCall(ByVal JiraRequest As WebRequest) As WebResponse
       
    Dim JiraClient As New WebClient
    JiraClient.baseUrl = JiraBaseUrl
    
    'Setup Authentication
    Dim JiraAuth As New HttpBasicAuthenticator
    JiraAuth.Setup _
        userName:=userName, _
        password:=userPassword
    
    Set JiraClient.Authenticator = JiraAuth
    
    ' Set the request format
    ' -> Sets content-type and accept headers and parses the response
    JiraRequest.ContentType = "application/json;charset=UTF-8"

    ' Execute the request and return the response
    Set JiraCall = JiraClient.Execute(JiraRequest)
    
End Function

