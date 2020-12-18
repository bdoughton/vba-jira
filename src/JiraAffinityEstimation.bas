Attribute VB_Name = "JiraAffinityEstimation"
''
' JiraAffinityEstimation v1.0
' (c) Ben Doughton - https://github.com/bdoughton/vba-jira
'
' Contains modules for Importing Jira tickets for Affinity Esimation and then updating Jira Story Points. Includes:
'
' - AffinityEstimationHelp
' - GetJiraSearchResults
' - UpdateStoryPoints
'
'
' @module JiraAffinityEstimation
' @author bdoughton@me.com
' @license GNU General Public License v3.0 (https://opensource.org/licenses/GPL-3.0)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

Sub AffinityEstimationHelp(control As IRibbonControl)
    Dim message As String
    message = "(1) Enter the JQL for the issues you want to estimate (the query will only return the first 50 issues - max)" & vbCrLf & _
                      "(2) Enter the story point scale you want to use in the first row of cells on the sheet" & vbCrLf & _
                      "(3) Estimate the stories by move the cards to the desired column (you may wish to resize the columns)" & vbCrLf & _
                      "(4) When ready choose the Save Estimates option (the top left corner of the card is used to determin the correct colum)"
    MsgBox (message)
End Sub

Sub GetJiraSearchResults(control As IRibbonControl)
    
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

    Dim JQL As String
    JQL = InputBox("Enter a JQL query that returns the Issues you want to estimate:")

    Dim apiFields As String
    apiFields = "key," _
            & "summary"
     
    'Define the new JQLRequest
    Dim JQLsearchRequest As New WebRequest
    With JQLsearchRequest
        .Resource = "api/2/search"
        .Method = WebMethod.HttpGet
        .AddQuerystringParam "jql", JQL
        .AddQuerystringParam "fields", apiFields
        .AddQuerystringParam "startAt", 0
        .AddQuerystringParam "maxResults", "50"
    End With
               
    Dim JQLsearchResponse As New JiraResponse
    Dim searchResponse As New WebResponse
     
    Set searchResponse = JQLsearchResponse.JiraCall(JQLsearchRequest)

If searchResponse.StatusCode = WebStatusCode.Ok Then
    Dim sh As Shape
    Dim n As Integer
    Dim Item As Object
    n = 0
    For Each Item In searchResponse.Data("issues")
        n = n + 1
        Set sh = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50 * n, 200, 45)
        sh.name = searchResponse.Data("issues")(i)("key")
        sh.TextFrame2.TextRange.Characters.Text = "Key: " & _
            searchResponse.Data("issues")(i)("key") & vbCrLf & _
            "Summary: " & searchResponse.Data("issues")(i)("summary")
    Next Item
End If
   
'Reverse the opening statements that paused calculations and screen updating
'hide worksheets that should not be edited by the user
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub


Sub UpdateStoryPoints(control As IRibbonControl)

    Dim s As Shape, message As String, sp As Integer
    For Each s In ActiveSheet.Shapes
        sp = Sheet1.Cells(1, s.TopLeftCell.Column).Value
        message = message & vbCrLf & s.name & "---" & sp
    Next s
    MsgBox message
End Sub

