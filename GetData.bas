Attribute VB_Name = "GetData"
'' -------------------------------------------------------------------
'' INFO
'' -------------------------------------------------------------------
'' This module is where all the macros for sourcing/refreshing data
'' from Jira are held i.e. Step 1 on the instructions tab
'' -------------------------------------------------------------------
Public encodedAuth As String

Option Explicit

Public Sub SourceStandardIssuesandSubTasks()

'' --------------------------------------------------------------------
'' This macro is called by the user and gets the latest issues from the
'' filter in cell C10 on the Instructions tab. It saves the issues to a
'' temporary table and then locates all the SubTasks for those issues
'' combining the StandardIssueTypes and SubTaskIssueTypes Data in a Matrix
'' --------------------------------------------------------------------

Dim http, JSON, Item As Object
Dim p, i, l, r, c As Integer
Dim url, apiFields As String
Dim callResult As Long

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
             
'Pause calculations and screen updating and make read-only worksheets visible
'These actions are reversed at the end of the macro
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual


'Rest the StandardIssueTypes worksheet
With ws_StandardIssueTypesData
        .Cells.ClearContents
        .Cells(1, 1).value = "Project"
        .Cells(1, 2).value = "Issue Type"
        .Cells(1, 3).value = "Key"
        .Cells(1, 4).value = "External Issue ID"
        .Cells(1, 5).value = "Summary"
        .Cells(1, 6).value = "Status"
        .Cells(1, 7).value = "Updated"
        .Cells(1, 8).value = "Assignee"
        .Cells(1, 9).value = "Labels"
        .Cells(1, 10).value = "Due Date"
        .Cells(1, 11).value = "Original Due Date"
        .Cells(1, 12).value = "Accountable Department"
        .Cells(1, 13).value = "Latest Comment"
End With

'Fetch Filter Data from Api
callResult = RestApiCalls.GetJQLForFilter(encodedAuth, url, ws_Sheet1.Range("filter").value)

'Output error if call was not successful
If callResult <> 200 Then
    ws_Sheet1.Range("filter").value = "Error - " & callResult
    ws_ProjectData.Visible = xlSheetVeryHidden
    ws_FixVersionsData.Visible = xlSheetVeryHidden
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
End If

apiFields = "&fields=" _
        & "project," _
        & "issuetype," _
        & "key," _
        & ExternalIssueID & "," _
        & "summary," _
        & "status," _
        & "updated," _
        & "assignee," _
        & "labels," _
        & "duedate," _
        & OriginalDueDate & "," _
        & AccountableDepartment & "," _
        & "comment" _
        & "&maxResults=1000"


'Fetch StandardIssue Data from Api
callResult = RestApiCalls.GetAuditIssues(encodedAuth, ws_Sheet1.Range("JQL").value & apiFields, ws_StandardIssueTypesData, 2)

'Output error if call was not successful
If callResult <> 200 Then
    ws_StandardIssueTypesData.Cells(2, 1).value = "Error - " & callResult
    ws_ProjectData.Visible = xlSheetVeryHidden
    ws_FixVersionsData.Visible = xlSheetVeryHidden
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
End If

'Find the number of rows imported
p = ws_StandardIssueTypesData.Cells(ws_StandardIssueTypesData.Rows.Count, "A").End(xlUp).Row

'Reset the All_Issues worksheet
With ws_AllIssues
        .Cells.ClearContents
        .Cells(1, 1).value = "Project"
        .Cells(1, 2).value = "Issue Type"
        .Cells(1, 3).value = "Key"
        .Cells(1, 4).value = "External Issue ID"
        .Cells(1, 5).value = "Summary"
        .Cells(1, 6).value = "Status"
        .Cells(1, 7).value = "Updated"
        .Cells(1, 8).value = "Assignee"
        .Cells(1, 9).value = "Labels"
        .Cells(1, 10).value = "Due Date"
        .Cells(1, 11).value = "Original Due Date"
        .Cells(1, 12).value = "Accountable Department"
        .Cells(1, 13).value = "Latest Comment"
End With

'This section inserts the subtasks
r = 2
For i = 2 To p
    For c = 1 To 13
        ws_AllIssues.Cells(r, c).value = ws_StandardIssueTypesData.Cells(i, c)
    Next c
    If RestApiCalls.GetAuditIssues(encodedAuth, _
        url & "rest/api/2/search?jql=parent%20%3D%20" & ws_AllIssues.Cells(r, 3) & "%20ORDER%20BY%20Rank" _
        & apiFields, ws_AllIssues, r + 1) = 200 Then
            r = ws_AllIssues.Cells(ws_AllIssues.Rows.Count, "B").End(xlUp).Row + 1
        Else
            r = r + 1
    End If
Next i


'' --------------------------------------------------------------------
'' Final section to tidy up and present the results back to the user
'' --------------------------------------------------------------------

'Reverse the opening statements that paused calculations and screen updating
'hide worksheets that should not be edited by the user
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

ws_AllIssues.Activate
ws_AllIssues.Range(Cells(1, 1), Cells(r, 13)).AutoFilter
ws_AllIssues.Cells(2, 1).Select

'Comfirm success to the user
'' ---------------------------------------------------------------------
''This needs updating to cater for api errors - to do later
MsgBox ("Complete") '& Chr(13) _
'' ---------------------------------------------------------------------

End Sub

Sub SourcePortfolioData()

'' --------------------------------------------------------------------
'' This macro is called by the user and gets the latest portfolio data
'' based on a planId and scenarioId stored in the Instructions sheet
'' --------------------------------------------------------------------

Dim http, JSON, Item As Object
Dim l As Integer
Dim url As String
Dim callResult As Long
Dim msgCaption As String

Dim cel As Range
Dim issueidRange As Range
Dim Order As String
Dim ParentId As Long



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
             
'Pause calculations and screen updating and make read-only worksheets visible
'These actions are reversed at the end of the macro
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

'Fetch Roadmap and Sprint Data from Backlog Api
callResult = RestApiCalls.PostBacklog(encodedAuth, url, 1, planId, scenarioId, TeamId)

'Output error if call was not successful
If callResult <> 200 Then
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox ("Error")
    Exit Sub
End If

MsgBox ("Success")

With ws_Roadmap
    .Activate
    If .FilterMode Then ActiveSheet.ShowAllData
    .Cells.EntireColumn.Hidden = False
    .Cells.EntireRow.Hidden = False
    .Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
End With

'With ws_Roadmap.Range("A3")
'    .FormulaR1C1 = _
'        "=IF(RC8="""",RC4,IF(RC5=""Epic"",INDEX(C4,MATCH(RC8,C3,0))&RC4,INDEX(C4,MATCH(INDEX(C8,MATCH(RC8,C3,0)),C3,0))&INDEX(C4,MATCH(RC8,C3,0))&RC4))"
'    .AutoFill Destination:=Range("A3:A" & ws_Roadmap.Range("C3").End(xlDown).row)
'End With

With ws_Roadmap.Range("B3")
    .FormulaR1C1 = _
        "=INDEX(IssueTypes!C[1],MATCH(RC[3],IssueTypes!C,0))"
    .AutoFill Destination:=Range("B3:B" & ws_Roadmap.Range("C3").End(xlDown).Row)
End With

With ws_Roadmap
    .Range("P3").FormulaR1C1 = _
        "=IF(AND(R1C>=RC14,R1C<=RC15),1,0)"
    .Range("P3").AutoFill Destination:=Range("P3:AN3"), Type:=xlFillValues
    .Range("P3:AN3").AutoFill Destination:=Range("P3:AN" & ws_Roadmap.Range("C3").End(xlDown).Row), Type:=xlFillValues

End With

Set issueidRange = ws_Roadmap.Range("C3:C" & ws_Roadmap.Range("C3").End(xlDown).Row)
For Each cel In issueidRange
    Order = ""
    ParentId = cel.Offset(0, 5).value
    Do Until ParentId = 0
        With ws_Roadmap.Range("C:C").Find(ParentId)
            Order = .Offset(0, 1).value & Order
            ParentId = .Offset(0, 5).value
        End With
    Loop
    cel.Offset(0, -2).value = Order & cel.Offset(0, 1).value
Next cel

'Resume calculations and screen updating
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

ws_Roadmap.AutoFilter.Sort.SortFields.Clear
ws_Roadmap.AutoFilter.Sort.SortFields.Add2 key:= _
    Range("A2:A" & ws_Roadmap.Range("C3").End(xlDown).Row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
    :=xlSortNormal
With ws_Roadmap.AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
 
End Sub
Sub funcFindParentRank()

Dim issueidRange As Range
Dim Order As String
Dim ParentId As Long

Set issueidRange = Selection

Order = ""
ParentId = issueidRange.Offset(0, 5).value

Do Until ParentId = 0
    With ws_Roadmap.Range("C:C").Find(ParentId)
        Order = .Offset(0, 1).value & Order
        ParentId = .Offset(0, 5).value
    End With
Loop

Order = Order & issueidRange.Offset(0, 1).value
Debug.Print Order
End Sub

Sub UpdateBaseData()

'' --------------------------------------------------------------------
'' This macro is optionally called by the user and gets the latest base
'' Project Data, Status Data and Team Data
'' --------------------------------------------------------------------

Dim callResult(1 To 5) As Long
Dim url As String
Dim l As Integer
Dim i As Integer
Dim msgCaption As String

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

'Fetch Filter Data from Api
callResult(1) = RestApiCalls.GetAllProjects(encodedAuth, url)
callResult(2) = RestApiCalls.GetAllStatuses(encodedAuth, url)
callResult(3) = RestApiCalls.PostTeamsFind(encodedAuth, url)
callResult(4) = RestApiCalls.GetRapidBoard(RapidBoardId, encodedAuth, url)

'Reset FixVersion sheet
ws_FixVersionsData.Visible = xlSheetVisible
ws_FixVersionsData.Activate
ws_FixVersionsData.Range(Cells(2, 1), Cells(ws_FixVersionsData.Range("A2").End(xlDown).Row, 6)).ClearContents

callResult(5) = RestApiCalls.GetProjectFixVersions(encodedAuth, url, 14000, "CMS56205", 2)
ws_FixVersionsData.Visible = xlSheetHidden

'Output error if call was not successful
For i = 1 To 5
    If callResult(i) <> 200 Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox ("Error - with call no: " & i)
        Exit Sub
    End If
Next i

MsgBox ("Success")


End Sub

Sub UpdateEpicBurnUpData()

'' --------------------------------------------------------------------
'' This macro is called by the user and updates the Epic BurnUp Data
'' Based on the Epic Burn Down chart data
'' --------------------------------------------------------------------

Dim callResult As Long
Dim url As String
Dim l As Integer
Dim epic As String
Dim c As Range
Dim msgCaption As String

'Fetch the base url and board id
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

''Update the Epics list and rank
callResult = RestApiCalls.GetEpicsForBoard(encodedAuth, url, ws_BoardData.Cells(5, 2).value, ws_EpicsData, 10)

Application.Calculate

ws_BurnUpData.Visible = xlSheetVisible
ws_BurnUpData.Activate

For Each c In ws_BurnUpData.Range("Epics")
    epic = c.value
    'Fetch Filter Data from Api
    callResult = RestApiCalls.GetEpicBurndown(RapidBoardId, epic, encodedAuth, url)
Next c

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

Sub ReleaseDatesToRoadmap()
'
' This macro adds the Release Dates to the Roadmap
' Still very much WIP
'

'
''1030 start

''48 per sprint
'' 3.4 per day

    
    ws_Roadmap.Shapes.AddShape(msoShapeDiamond, 2123, 16, 12, 12).Select
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset14
    
End Sub

Sub Build5YRoadmap()

'This does not yet clear previous data and requires that column m is visible on the 1 year roadmap

    Dim countTheme(1 To 12) As Long
    Dim rangeTheme() As Range
    Dim c As Range
    Dim EpicCell As Range
    Dim n As Integer 'n is the theme number
    Dim i As Long 'i is the initiatives in the theme
    Dim r As Long 'r is the total number of initatives (number of ranges)
    Dim totalCountOfThemes As Long 'running total of n
    Dim lastRow As Long
    Dim FiveYrRow As Long
    
    ws_Roadmap.Activate
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    lastRow = ws_Roadmap.Range("B1048576").End(xlUp).Row
    FiveYrRow = 2
    
    r = 0
    totalCountOfThemes = 0
    For n = 1 To UBound(countTheme)
        countTheme(n) = funcCountTheme(n)
        totalCountOfThemes = totalCountOfThemes + countTheme(n)
        If countTheme(n) > 0 Then
            For i = 1 To countTheme(n)
                If i = 1 Then Set c = ws_Roadmap.Range("M1")
                If n = 1 Then
                    ReDim rangeTheme(0 To countTheme(n) - 1)
                Else
                    ReDim Preserve rangeTheme(0 To totalCountOfThemes - 1)
                End If
                Set rangeTheme(r) = funcFindThemeRange(n, c)
                Set c = rangeTheme(r)
                If rangeTheme(r).End(xlDown).Row > lastRow Then
                    ' extend the range to the last row
                    Set rangeTheme(r) = Range(rangeTheme(r), Cells(lastRow, rangeTheme(r).Column))
                Else
                    If rangeTheme(r).Offset(1, 0).value <> "" Then
                        'Leave the range as it is
                    Else
                        'Extend the range to the next value in column m
                        Set rangeTheme(r) = Range(rangeTheme(r), rangeTheme(r).End(xlDown).Offset(-1, 0))
                    End If
                End If
                Debug.Print (rangeTheme(r).Address)
                For Each EpicCell In rangeTheme(r).Offset(0, -8)
                    Debug.Print (EpicCell.Address)
                    If EpicCell.value = "Epic" Or EpicCell.value = "Initiative" Then
                        With ws_FiveYrRoadmap
                            .Cells(FiveYrRow, 1).value = funcThemeName(n) 'Theme
                            .Cells(FiveYrRow, 2).value = EpicCell.Offset(0, 2) 'Key
                            .Cells(FiveYrRow, 3).value = EpicCell.Offset(0, 4) 'Summary
                            .Cells(FiveYrRow, 4).value = EpicCell 'IssueType
                            .Cells(FiveYrRow, 5).value = EpicCell.Offset(0, 9) 'StartDate
                            .Cells(FiveYrRow, 6).value = EpicCell.Offset(0, 10) 'EndDate
                        End With
                        FiveYrRow = FiveYrRow + 1
                    End If
                Next EpicCell
                
                Debug.Print (rangeTheme(r).Address) 'The range of cells in column M related to each theme in theme order
                r = r + 1
            Next i
        End If
        'Print the array countTheme
        Debug.Print (n & " = " & countTheme(n)) ' The number of initiatives in each theme
    Next n
    
    For r = 0 To UBound(rangeTheme)
        'Print the array rangeTheme
'        Debug.Print (r & " = " & rangeTheme(r).Address) 'The range of cells in column M related to each theme in theme order
    Next r
    
    With ws_FiveYrRoadmap
        .Activate
        .Range("G2").FormulaR1C1 = _
                "=IF(AND(RC5<=EOMONTH(DATE(20&RIGHT(R1C,2),3*MID(R1C,2,1),1),0),RC6>=DATE(20&RIGHT(R1C,2),3*MID(R1C,2,1)-2,1)),1,0)"
        .Range("G2").AutoFill Destination:=Range("G2:Z2") '& ws_FiveYrRoadmap.Range("A1").End(xlDown).row)
        .Range("G2:Z2").AutoFill Destination:=Range("G2:Z" & .Range("A1").End(xlDown).Row)
    End With
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

Function funcFindThemeRange(ByVal n As Integer, ByVal c As Range) As Range
    If n > 9 Then
        Set funcFindThemeRange = ws_Roadmap.Range("M:M").Find(n & ".", After:=c, LookIn:=xlValues, LookAt:=xlPart)
    Else
        Set funcFindThemeRange = ws_Roadmap.Range("M:M").Find("0" & n & ".", After:=c, LookIn:=xlValues, LookAt:=xlPart)
    End If
    Debug.Print (funcFindThemeRange.Address)
End Function
Function funcCountTheme(ByVal n As Integer) As Long
    If n > 9 Then
        funcCountTheme = Excel.WorksheetFunction.CountIf(ws_Roadmap.Range("M:M"), "*" & n & "." & "*")
    Else
        funcCountTheme = Excel.WorksheetFunction.CountIf(ws_Roadmap.Range("M:M"), "*0" & n & "." & "*")
    End If
End Function
Function funcThemeName(ByVal n As Integer) As String
    Select Case n
        Case 1
            funcThemeName = "01.Eligibility and Limits"
        Case 2
            funcThemeName = "02.Pre-Trade Analytics"
        Case 3
            funcThemeName = "03.Collateral Sourcing"
        Case 4
            funcThemeName = "04.Pricing and Market Data"
        Case 5
            funcThemeName = "05.Exposure Calculation"
        Case 6
            funcThemeName = "06.Optimisation"
        Case 7
            funcThemeName = "07.Counterparty Risk & Default"
        Case 8
            funcThemeName = "08.Settlements & Payments"
        Case 9
            funcThemeName = "09.Safekeeping"
        Case 10
            funcThemeName = "10.Reporting"
        Case 11
            funcThemeName = "11.Interest & Fees"
        Case 12
            funcThemeName = "12.Service Improvement & Availability"
    End Select
End Function

