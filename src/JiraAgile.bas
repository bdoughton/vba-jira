Attribute VB_Name = "JiraAgile"
''
' JiraAgile v1.0
' (c) Ben Doughton - https://github.com/bdoughton/vba-jira
'
' Contains modules for Importing Jira Agile Boards. Includes:
'
' - GetJiraAgileBoard
' - ProcessAgileBoardView
'
'
' @module JiraAgile
' @author bdoughton@me.com
' @license GNU General Public License v3.0 (https://opensource.org/licenses/GPL-3.0)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit


Sub GetJiraAgileBoard(control As IRibbonControl)

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
    
    If MsgBox("This will replace the data in the current sheet!", vbOKCancel, "Import Jira Board") = vbOK Then
        ActiveSheet.Cells.Clear
        Dim Resource As String
        Dim rapidViewId As String
        rapidViewId = InputBox("rapidViewId?")
        Resource = "greenhopper/1.0/rapidviewconfig/editmodel.json?rapidViewId=" & rapidViewId
        Dim agileBoard As New JiraResponse
        ProcessAgileBoardView agileBoard.GetJira(Resource), "work" 'ProcessAgileBoardView does not yet support plan mode
    End If
   
'Reverse the opening statements that paused calculations and screen updating
'hide worksheets that should not be edited by the user
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
   
End Sub

Sub ProcessAgileBoardView(ByVal agileBoardResponse As WebResponse, ByVal boardMode As String)

If agileBoardResponse.StatusCode = WebStatusCode.Ok Then

    '' Sub re-creates the Jira Agile board on a spreadsheet
    '' It works by:
    '' (1) identifying the columns and inserting them as a header row
    '' (2a) identifying the number of swimlanes and looping through each one
    '' (2b) for each swimlane a seperate JQL api search is called and issues added to the correct row and column
    
    Dim c%, s%, q%, i As Integer
    Dim countColumns As Integer
    
    ' Put the column headings into the active sheet
    ''Currently adds Kanaban backlog-- even if not in use
    countColumns = agileBoardResponse.Data("rapidListConfig")("mappedColumns").Count
    For c = 1 To countColumns
         ActiveSheet.Cells(1, c).Value = agileBoardResponse.Data("rapidListConfig")("mappedColumns")(c)("name")
         ActiveSheet.Columns(c).ColumnWidth = 60
    Next c
    
    ' Identify any user configure fields that have been added to the board
    If agileBoardResponse.Data("cardLayoutConfig")("currentFields").Count > 0 Then
        Dim extraFields As String
        For c = 1 To agileBoardResponse.Data("cardLayoutConfig")("currentFields").Count
            If agileBoardResponse.Data("cardLayoutConfig")("currentFields")(c)("mode") = "workMode" Then
                extraFields = extraFields & "," & agileBoardResponse.Data("cardLayoutConfig")("currentFields")(c)("fieldId")
            End If
        Next c
    End If
    
    ' Define all the fields to show, the out of the box fields plus the extra fields identified above
    Dim apiFields As String
    apiFields = "key," _
            & "summary," _
            & "issuetype," _
            & "assignee," _
            & "priority," _
            & "status" _
            & extraFields
    
    ' Identify the board_type
    Dim board_type As String
    If agileBoardResponse.Data("isSprintSupportEnabled") Then
        board_type = "Scrum"
        MsgBox ("Swimlanes for Scrum Boards may not be represented correctly!")
    Else
        board_type = "Kanban"
    End If
    
    ' Define the sub_query based on the board_type
    Dim sub_query As String
    If board_type = "Scrum" Then
        If boardMode = "plan" Then
            sub_query = " AND issuetype not in (Epic) AND (((Sprint is EMPTY OR Sprint in closedSprints()) AND " & _
                                    "statusCategory not in (Done)) OR Sprint in (openSprints(),futureSprints()))"
        ElseIf boardMode = "work" Then
            sub_query = " AND Sprint in (openSprints()) AND issuetype not in (Epic) AND (((Sprint is EMPTY OR Sprint in closedSprints()) AND " & _
                                    "statusCategory not in (Done)) OR Sprint in (openSprints(),futureSprints()))"
        End If
    ElseIf board_type = "Kanban" Then
        If boardMode = "plan" Then
            '' sub_query needs updating to support plan mode and needs to check if board supports it
            sub_query = " AND (" & agileBoardResponse.Data("subqueryConfig")("subqueries")(1)("query") & _
                                ") AND (statusCategory not in (Done) OR updated>=" & _
                                agileBoardResponse.Data("oldDoneIssuesCutoff") & ")"
        ElseIf boardMode = "work" Then
            sub_query = " AND (" & agileBoardResponse.Data("subqueryConfig")("subqueries")(1)("query") & _
                                ") AND (statusCategory not in (Done) OR updated>=" & _
                                agileBoardResponse.Data("oldDoneIssuesCutoff") & ")"
        End If
    End If
          
    ' Define the sub queries to be used for each swimlane
    Dim swimlaneQueries() As String
    Dim swimlaneCount As Integer
    
    swimlaneCount = agileBoardResponse.Data("swimlanesConfig")("swimlanes").Count
    
    ReDim swimlaneQueries(1 To swimlaneCount - 1) ' Causes an error if there is only one swimlane
    
    If swimlaneCount > 1 Then
        For c = 1 To swimlaneCount - 1 '' -1: don't include the default query
            swimlaneQueries(c) = agileBoardResponse.Data("swimlanesConfig")("swimlanes")(c)("query")
        Next c
        ' Define the default swimlane query
        Dim defaultswimlaneQuery As String
        defaultswimlaneQuery = "NOT " & Join(swimlaneQueries, " AND ")
    Else
        defaultswimlaneQuery = agileBoardResponse.Data("swimlanesConfig")("swimlanes")(1)("query")
    End If
    
    ' Define the first row for each query on the Active Sheet that the Jira issues will be added to this need to be per column
    Dim outputRow() As Long
    Dim swimlaneRow As Long
    ReDim outputRow(1 To countColumns)
    swimlaneRow = 2 ' set initally to 2
           
    Dim swimlaneQuery As String
    Dim FullJQL As String
    Dim JQLSearch() As New JiraResponse
    Dim JQLSeachResponse As New WebResponse
    Dim Resource As String
    Dim Item As Object
    Dim IssueState As String
    Dim IssueColumn As Integer
    Dim issueCard As String
    Dim assignee As String
    Dim priority As String
    
    ReDim JQLSearch(1 To swimlaneCount)
      
    For q = 1 To swimlaneCount
        If q < swimlaneCount Then
            swimlaneQuery = swimlaneQueries(q)
        Else
            swimlaneQuery = defaultswimlaneQuery
        End If
'        If swimlaneQuery <> "" Then swimlaneQuery = " AND " & swimlaneQuery '' not sure why this is here? But the final query is not working so could be related
        
        FullJQL = "Filter=" & agileBoardResponse.Data("filterConfig")("id") & sub_query & " AND " & swimlaneQuery & " ORDER BY Rank"
        
        'Define the new JQLRequest
        Dim JQLRequest As New WebRequest
        JQLRequest.Resource = "api/2/search"
        JQLRequest.Method = WebMethod.HttpGet
        JQLRequest.AddQuerystringParam "jql", FullJQL
        JQLRequest.AddQuerystringParam "fields", apiFields
        JQLRequest.AddQuerystringParam "maxResults", "1000"
                
        Set JQLSeachResponse = JQLSearch(q).JiraCall(JQLRequest)
        
        If JQLSeachResponse.StatusCode = WebStatusCode.Ok Then
            With ActiveSheet
                ' Name the Swimlane
                .Cells(swimlaneRow, 1).Value = agileBoardResponse.Data("swimlanesConfig")("swimlanes")(q)("name")
                For c = 1 To countColumns
                    outputRow(c) = swimlaneRow + 1 'increment the row for each column so that the issues remain in their swimlane
                Next c
                i = 1
                For Each Item In JQLSeachResponse.Data("issues")
                    IssueState = JQLSeachResponse.Data("issues")(i)("fields")("status")("id")
                    c = 1
                    s = 1
                    ' Map the issue to the correct column (c)
                    Do Until c > countColumns
                        s = 1
                        Do Until s > agileBoardResponse.Data("rapidListConfig")("mappedColumns")(c)("mappedStatuses").Count
                            If IssueState = agileBoardResponse.Data("rapidListConfig")("mappedColumns")(c)("mappedStatuses")(s)("id") Then
                                If IsNull(JQLSeachResponse.Data("issues")(i)("fields")("assignee")) Then
                                    assignee = ""
                                Else
                                    assignee = JQLSeachResponse.Data("issues")(i)("fields")("assignee")("name")
                                End If
                                If IsNull(JQLSeachResponse.Data("issues")(i)("fields")("priority")) Then
                                    priority = ""
                                Else
                                    priority = JQLSeachResponse.Data("issues")(i)("fields")("priority")("name")
                                End If
                                issueCard = JQLSeachResponse.Data("issues")(i)("key") & vbNewLine & _
                                                    JQLSeachResponse.Data("issues")(i)("fields")("summary") & vbNewLine & vbNewLine & _
                                                    JQLSeachResponse.Data("issues")(i)("fields")("issuetype")("name") & vbNewLine & _
                                                    priority & vbNewLine & assignee
                                With .Cells(outputRow(c), c)
                                    .Value = issueCard
                                    .HorizontalAlignment = xlGeneral
                                    .VerticalAlignment = xlTop
                                    .WrapText = True
                                End With
                                outputRow(c) = outputRow(c) + 1 'increment the row for each column after inputting the card
                            End If
                            s = s + 1
                        Loop
                        c = c + 1
                    Loop
                    i = i + 1
                Next Item
                swimlaneRow = Application.WorksheetFunction.Max(outputRow)
            End With
        End If
        'Reset the JQLRequest
        Set JQLRequest = Nothing
    Next q
    MsgBox ("Success")
Else
    MsgBox ("Error")
End If

End Sub

