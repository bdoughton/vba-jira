Attribute VB_Name = "RestApiCalls"
'' -------------------------------------------------------------------
'' INFO
'' -------------------------------------------------------------------
'' This module is where all the apicalls are stored, most of which
'' are public functions called from other modules
'' -------------------------------------------------------------------

Public str_displayName As String
Public str_avatarUrl48 As String
Option Explicit

Public Function MyCredentials(ByVal auth As String, ByVal baseUrl As String) As Long

'' --------------------------------------------------------------------
'' This function validates that a user is logged in and sets the public
'' variables str_displayName and str_avatarUrl48
'' It requires an encoded user id/password and baseUrl
'' --------------------------------------------------------------------

Dim http, JSON As Object
Dim apicall As String

apicall = baseUrl & "rest/api/latest/myself"

Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

If http.Status = 200 Then
    Set JSON = ParseJson(http.responseText)
    str_displayName = JSON("displayName")
    str_avatarUrl48 = JSON("avatarUrls")("48x48")
End If

MyCredentials = http.Status

Debug.Print (strDebug("MyCredentials", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetJQLForFilter(ByVal auth As String, ByVal url As String, ByVal filter As Integer) As Long

Dim http, JSON, Item As Object
          
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", url & "rest/api/2/filter/" & filter, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & encodedAuth
http.Send

GetJQLForFilter = http.Status
'MsgBox (http.Status & Chr(13) & http.responseText & Chr(13) & http.getAllResponseHeaders())

If http.Status = 200 Then
    Set JSON = ParseJson(http.responseText)
        With ws_Sheet1
            .Range("JQL").value = JSON("searchUrl")
        End With
End If

Debug.Print (strDebug("GetJQLForFilter", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetAuditIssues(ByVal auth As String, ByVal apicall As String, ws As Worksheet, r As Integer) As Long

Dim http, JSON, Item As Object
Dim i, l, c As Integer
Dim str_Comment As String
             
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & encodedAuth
http.Send

GetAuditIssues = http.Status
'MsgBox (http.Status & Chr(13) & http.responseText & Chr(13) & http.getAllResponseHeaders())

If http.Status = 200 Then
    'If saveJsonToFile("M:\documents\Downloads\jql.json", http.responseText) Then ' download the json file for debugging
        Set JSON = ParseJson(http.responseText)
        i = 1
        For Each Item In JSON("issues")
            l = 1
            With ws
                If Not JSON("issues")(i)("fields")("issuetype")("subtask") Then
                '' if the issue type is not a subtask then populate the project name
                    .Cells(r, 1).value = JSON("issues")(i)("fields")("project")("name")
                End If
                .Cells(r, 2).value = JSON("issues")(i)("fields")("issuetype")("name")
                .Cells(r, 3).value = JSON("issues")(i)("key")
                .Cells(r, 4).value = JSON("issues")(i)("fields")(ExternalIssueID)
                .Cells(r, 5).value = JSON("issues")(i)("fields")("summary")
                .Cells(r, 6).value = JSON("issues")(i)("fields")("status")("name")
                .Cells(r, 7).value = Left(JSON("issues")(i)("fields")("updated"), 10) & " " & _
                        Mid(JSON("issues")(i)("fields")("updated"), 12, 5)
                .Cells(r, 8).value = JSON("issues")(i)("fields")("assignee")("name")
                If JSON("issues")(i)("fields")("labels").Count > 0 Then
                    For l = 1 To JSON("issues")(i)("fields")("labels").Count
                        If l = 1 Then
                            .Cells(r, 9).value = JSON("issues")(i)("fields")("labels")(l)
                        Else
                            .Cells(r, 9).value = .Cells(r, 9).value & ", " & JSON("issues")(i)("fields")("labels")(l)
                        End If
                    Next l
                End If
                .Cells(r, 10).value = JSON("issues")(i)("fields")("duedate")
                .Cells(r, 11).value = JSON("issues")(i)("fields")(OriginalDueDate)
                On Error Resume Next 'Error handling to deal with Null AccountableDepartments
                .Cells(r, 12).value = JSON("issues")(i)("fields")(AccountableDepartment)("value")
                On Error GoTo 0
                If JSON("issues")(i)("fields")("comment")("comments").Count > 0 Then
                    c = JSON("issues")(i)("fields")("comment")("comments").Count ' find the last comment
                    ' Build the comment from the date it was upated, the user's display name and the comment body
                    str_Comment = Left(JSON("issues")(i)("fields")("comment")("comments")(c)("updated"), 10) _
                        & " - " & JSON("issues")(i)("fields")("comment")("comments")(c)("author")("displayName") _
                        & " : " & JSON("issues")(i)("fields")("comment")("comments")(c)("body")
                    'insert the comment
                    .Cells(r, 13).value = str_Comment
                End If
            End With
            i = i + 1 'increment the issue
            r = r + 1 'increment the row
        Next
    'End If
End If

Debug.Print (strDebug("GetAuditIssues", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetEpicsForBoard(ByVal auth As String, ByVal baseUrl As String, ByVal filter_query As String, ws As Worksheet, r As Integer) As Long

Dim http, JSON, Item As Object
Dim i, c As Integer
Dim str_Comment As String
Dim apiFields As String
             
Set http = CreateObject("MSXML2.XMLHTTP")

'' Add an additional filter to the board's filter
'Commented out until the Jira resolved date is corrected
'filter_query = "issuetype = Epic AND (resolved is EMPTY or resolved >= startOfYear()) AND " & filter_query

'' This works better! cf[10007] is Epic Status. Meaning that this matches up with the Team Board
filter_query = "cf[10007] not in (Done) AND issuetype = Epic AND " & filter_query

'' Update the order by in the filter to ensure resolved issues appear first
'Commented out until the Jira resolved date is corrected
'filter_query = Replace(filter_query, "ORDER BY Rank ASC", "ORDER BY resolved asc, Rank ASC")

'' Specifcy the fields that are required
apiFields = "&fields=" _
        & "project," _
        & "issuetype," _
        & "key," _
        & "summary," _
        & "status," _
        & storypoints & "," _
        & "resolutiondate" _
        & "&maxResults=1000"

http.Open "GET", baseUrl & "rest/api/2/search?jql=" & filter_query & apiFields, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & encodedAuth
http.Send
 Debug.Print baseUrl & "rest/api/2/search?jql=" & filter_query & apiFields
GetEpicsForBoard = http.Status

If http.Status = 200 Then
    ws.Cells.ClearContents
    Set JSON = ParseJson(http.responseText)
    'Call PrintDictionary("Board Epics", JSON)
    i = 1
    c = 1
    For Each Item In JSON("issues")
        With ws
            .Cells(1, c).value = JSON("issues")(i)("key")
            .Cells(2, c).value = JSON("issues")(i)("fields")("summary")
            .Cells(r, 1).value = JSON("issues")(i)("key")
            .Cells(r, 2).value = JSON("issues")(i)("fields")("project")("name")
            .Cells(r, 3).value = JSON("issues")(i)("fields")("issuetype")("name")
            .Cells(r, 4).value = JSON("issues")(i)("fields")("summary")
            .Cells(r, 5).value = JSON("issues")(i)("fields")("status")("name")
            .Cells(r, 6).value = JSON("issues")(i)("fields")(storypoints)
            If Not IsNull(JSON("issues")(i)("fields")("resolutiondate")) Then
                .Cells(r, 7).value = DateValue(Left(JSON("issues")(i)("fields")("resolutiondate"), 10))
            End If
        End With
        i = i + 1 'increment the issue
        c = c + 1 'increment the column
        r = r + 1
    Next
End If

Debug.Print (strDebug("GetEpicsForBoard", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetAllProjects(ByVal auth As String, ByVal baseUrl As String) As Long

Dim http, JSON, Item As Object
Dim p, r As Integer
          
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", baseUrl & "rest/api/2/project", False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & encodedAuth
http.Send

GetAllProjects = http.Status
'MsgBox (http.Status & Chr(13) & http.responseText & Chr(13) & http.getAllResponseHeaders())

If http.Status = 200 Then
    ws_ProjectData.Visible = xlSheetVisible
    ws_ProjectData.Activate
    Set JSON = ParseJson(http.responseText)
    p = 1 'reset the project to 1
    r = 2
    With ws_ProjectData
        .Range(Cells(2, 1), Cells(.Range("A1048576").End(xlUp).Row, 4)).ClearContents ' clear existing data
        For Each Item In JSON
            .Cells(r, 1).value = JSON(p)("id")
            .Cells(r, 2).value = JSON(p)("key")
            .Cells(r, 3).value = JSON(p)("name")
            On Error Resume Next
                .Cells(r, 4).value = JSON(p)("projectCategory")("name")
            On Error GoTo 0
            p = p + 1 'increment the project
            r = r + 1 'increment the row
        Next
    End With
    ws_ProjectData.Visible = xlSheetHidden
    ws_Sheet1.Activate
End If

Debug.Print (strDebug("GetAllProjects", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetProjectFixVersions(ByVal auth As String, ByVal baseUrl As String, _
    ByVal pId As Long, pKey As String, r As Long) As Long

Dim http, JSON, Item As Object
Dim v As Integer
             
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", baseUrl & "rest/api/latest/project/" & pId & "/versions", False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

GetProjectFixVersions = http.Status
'MsgBox (http.Status & Chr(13) & http.responseText & Chr(13) & http.getAllResponseHeaders())

If http.Status = 200 Then
    Set JSON = ParseJson(http.responseText)
    v = 1 'reset the fixVersion to 1
    On Error Resume Next
    For Each Item In JSON
        With ws_FixVersionsData
            .Cells(r, 1).value = pId
            .Cells(r, 2).value = pKey
            .Cells(r, 3).value = JSON(v)("id")
            .Cells(r, 4).value = JSON(v)("name")
            .Cells(r, 5).value = JSON(v)("released")
            .Cells(r, 6).value = JSON(v)("releaseDate")
        End With
        r = r + 1 'increment the row
        v = v + 1 'increment the project
    Next
End If

Debug.Print (strDebug("GetProjectFixVersions", pId & "[" & pKey & "]", http.Status, http.getResponseHeader("X-AUSERNAME")))
'Debug.Print (http.getAllResponseHeaders())

End Function

Public Function UpdateVersion(ByVal auth As String, ByVal vId As Long, ByVal ReleaseDate As String, ByVal Release As Boolean, ByVal apicall As String)

Dim http, JSON, Item As Object
Dim JsonPut As String
Dim rngVersion As Range
Dim vName As String
Dim pKey As String
Dim Release_Json As String
Dim ReleaseDate_Json As String

If Release Then
    Release_Json = "true"
Else
    Release_Json = "false"
End If

If ReleaseDate = "Blank" Then
    ReleaseDate_Json = "null"
Else
    ReleaseDate_Json = Chr(34) & ReleaseDate & Chr(34)
End If

JsonPut = "{" & Chr(34) & "released" & Chr(34) & ":" & Release_Json & Chr(44) _
            & Chr(34) & "releaseDate" & Chr(34) & ":" & ReleaseDate_Json & _
            "}"
           
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "PUT", apicall & vId, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send (JsonPut)

UpdateVersion = http.Status

Set rngVersion = ws_FixVersionsData.Range("C:C").Find(vId)
With rngVersion
    vName = .Offset(0, 1).value
    pKey = .Offset(0, -1).value
End With

Debug.Print (strDebug("UpdateVersion", vId & "[" & pKey & "|" & vName & "]", http.Status, http.getResponseHeader("X-AUSERNAME")))
'Debug.Print (http.Status & Chr(13) & http.responseText & Chr(13) & http.getAllResponseHeaders())

End Function

Public Function CreateVersion(ByVal auth As String, ByVal vName As String, ByVal ReleaseDate As String, _
    ByVal ProjectId As Long, ByVal apicall As String)

Dim http, JSON, Item As Object
Dim JsonPost As String
Dim rngProject As Range
Dim pKey As String

JsonPost = "{" & _
    Chr(34) & "name" & Chr(34) & ":" & Chr(34) & vName & Chr(34) & Chr(44) & _
    Chr(34) & "releaseDate" & Chr(34) & ":" & Chr(34) & ReleaseDate & Chr(34) & Chr(44) & _
    Chr(34) & "projectId" & Chr(34) & ":" & Chr(34) & ProjectId & Chr(34) _
    & "}"
             
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "POST", apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send (JsonPost)

CreateVersion = http.Status

Set rngProject = ws_ProjectData.Range("A:A").Find(ProjectId)
pKey = rngProject.Offset(0, 1).value

Debug.Print (strDebug("CreateVersion", vName & "|" & pKey, http.Status, http.getResponseHeader("X-AUSERNAME")))
'Debug.Print  (http.Status & Chr(13) & http.responseText & Chr(13) & http.getAllResponseHeaders())

End Function

Private Function strDebug(ByVal api As String, s As String, r As Integer, userName As String) As String

'' This function records a log of all the api calls to a Log tab

    Dim strLog As String
    Dim c As Range
    
    strLog = Now() & ": " & api & ": " & userName & ": " & s & " - " & r
    Set c = ws_Log.Cells(ws_ProjectData.Rows.Count, "A").End(xlUp)
    c.Offset(1, 0).value = strLog
    strDebug = strLog
End Function

Private Function saveJsonToFile(ByVal strPath As String, JsonTxt As String) As Boolean
' Use this function to Save the Json to File
' Sometimes the Json is to large and this does not work, so
' I have used a function to return the success or failure of the code

'Example strPath = "M:\documents\Downloads\jql.json"

On Error GoTo ErrorHandler

Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
Set oFile = FSO.CreateTextFile(strPath)
oFile.WriteLine JsonTxt
oFile.Close
Set FSO = Nothing
Set oFile = Nothing
saveJsonToFile = True

Exit Function

ErrorHandler:
saveJsonToFile = False

End Function
Public Function PostBacklog(ByVal auth As String, baseUrl As String, _
    r As Integer, planId As Integer, scenarioId As Integer, TeamId As String) As Long

'' I have commented out the code to source the epic list from the backlog as this
'' won't return closed items more than 30 days old. A new api is required for this


Dim http, JSON, Item As Object
Dim JsonPost As String
Dim i%, l As Integer
'Dim epics As Scripting.Dictionary
Dim str_Comment As String
Dim str_Labels As String
                     
JsonPost = "{" & _
    Chr(34) & "planId" & Chr(34) & ":" & planId & Chr(44) & _
    Chr(34) & "scenarioId" & Chr(34) & ":" & scenarioId & Chr(44) & _
    Chr(34) & "filter" & Chr(34) & ":" & "{" & _
    Chr(34) & "includeCompleted" & Chr(34) & ":" & "true" & Chr(44) & _
    Chr(34) & "includeCompletedSince" & Chr(34) & ":" & date2epoch(Now() - 50) & Chr(44) & _
    Chr(34) & "performDependencyCompletion" & Chr(34) & ":" & "false" & Chr(44) & _
    Chr(34) & "includeIssueLinks" & Chr(34) & ":" & "true" & _
    "}" & "}"

'Set epics = New Scripting.Dictionary
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "POST", baseUrl & "rest/jpo/1.0/backlog", False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send (JsonPost)

PostBacklog = http.Status

Debug.Print (baseUrl & "rest/jpo/1.0/backlog")
Debug.Print (JsonPost)

Debug.Print (strDebug("PostBacklog", "", http.Status, http.getResponseHeader("X-AUSERNAME")))


If http.Status = 200 Then
    'Reset the worksheet contents and formats
    ws_Roadmap.Range("A3:O" & ws_Roadmap.Range("A3").End(xlDown).Row).ClearContents
    
    Set JSON = ParseJson(http.responseText)
    r = 3
    i = 1
    For Each Item In JSON("issues")
        str_Labels = ""
        With ws_Roadmap
            .Cells(r, 3).value = JSON("issues")(i)("id")
            .Cells(r, 4).value = JSON("issues")(i)("values")("lexoRank")
            .Cells(r, 5).value = JiraLookUp(ws_IssueTypes.Range("A2").CurrentRegion, JSON("issues")(i)("values")("type"), 1)
            .Cells(r, 6).value = JiraLookUp(ws_ProjectData.Range("A:A"), JSON("issues")(i)("values")("project"), 2)
            .Cells(r, 7).value = JiraLookUp(ws_ProjectData.Range("A:A"), JSON("issues")(i)("values")("project"), 1) & "-" & JSON("issues")(i)("issueKey")
            .Cells(r, 8).value = JSON("issues")(i)("values")("parent")
            .Cells(r, 9).value = JSON("issues")(i)("values")("summary")
            .Cells(r, 10).value = JiraLookUp(ws_StatusData.Range("A:A"), JSON("issues")(i)("values")("status"), 1)
            .Cells(r, 11).value = JiraLookUp(ws_TeamsData.Range("A:A"), JSON("issues")(i)("values")("team"), 1)
            .Cells(r, 12).value = JSON("issues")(i)("values")("storyPoints")
            If .Cells(r, 5).value = "Initiative" Then
                If JSON("issues")(i)("values").Exists("labels") Then
                    For l = 1 To JSON("issues")(i)("values")("labels").Count
                        If str_Labels = "" Then
                            str_Labels = JSON("issues")(i)("values")("labels")(l)
                        Else
                            str_Labels = str_Labels & ", " & JSON("issues")(i)("values")("labels")(l)
                        End If
                    Next
                    .Cells(r, 13).value = str_Labels
                End If
            End If
            If JSON("issues")(i)("values")("baselineStart") <> "" Then
                .Cells(r, 14).value = epoch2date(JSON("issues")(i)("values")("baselineStart"))
            End If
            If JSON("issues")(i)("values")("baselineEnd") <> "" Then
                .Cells(r, 15).value = epoch2date(JSON("issues")(i)("values")("baselineEnd"))
            End If
        End With
        'If JSON("issues")(i)("values")("type") = 10000 Then ' If the issue is an Epic
        '    epics.Add JiraLookUp(ws_ProjectData.Range("A:A"), JSON("issues")(i)("values")("project"), 1) & "-" & JSON("issues")(i)("issueKey"), JSON("issues")(i)("values")("lexoRank")
        'End If
        i = i + 1 'increment the issue
        r = r + 1 'increment the row
    Next
    r = 3
    i = 1
    
    '' This next section updates the Sprints Worksheet
    
    '' Reset the worksheet contents and formats
    ws_Sprints.Range("A2:D" & ws_Sprints.Range("B2").End(xlDown).Row).ClearContents
    
    For Each Item In JSON("calculationResult")("solution")("teams")(TeamId)("intervals")
        With ws_Sprints
            If JSON("calculationResult")("solution")("teams")(TeamId)("intervals")(i)("sprintId") <> "" Then
                .Cells(r, 1).value = JSON("calculationResult")("solution")("teams")(TeamId)("intervals")(i)("sprintId")
            Else
                .Cells(r, 1).value = ""
            End If
            If JSON("calculationResult")("solution")("teams")(TeamId)("intervals")(i)("title") <> "" Then
                .Cells(r, 2).value = JSON("calculationResult")("solution")("teams")(TeamId)("intervals")(i)("title")
            Else
                .Cells(r, 2).value = "Proposed Sprint"
            End If
            .Cells(r, 3).value = epoch2date(JSON("calculationResult")("solution")("teams")(TeamId)("intervals")(i)("start"))
            .Cells(r, 4).value = epoch2date(JSON("calculationResult")("solution")("teams")(TeamId)("intervals")(i)("end"))
        End With
        i = i + 1 'increment the issue
        r = r + 1 'increment the row
    Next
    
    ' Sort the Epics in Ascending order and enter them into the Epics_Data tab
    'Set epics = SortDictionaryByValue(epics)
    'On Error Resume Next
    'For i = 0 To 26
    '    ws_EpicsData.Range("A1").Offset(0, i).value = epics.Keys(i)
    '    ws_EpicsData.Range("A2").Offset(0, i).value = epics.Items(i)
    'Next i
    
End If

End Function
Public Function GetAllStatuses(ByVal auth As String, ByVal baseUrl As String) As Long

Dim http, JSON, Item As Object
Dim s As Integer, r As Integer
          
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", baseUrl & "rest/api/2/status", False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

GetAllStatuses = http.Status
'MsgBox (http.Status & Chr(13) & http.responseText & Chr(13) & http.getAllResponseHeaders())

If http.Status = 200 Then
    ws_StatusData.Visible = xlSheetVisible
    ws_StatusData.Activate
    Set JSON = ParseJson(http.responseText)
    s = 1 'reset the status to 1
    r = 2
    With ws_StatusData
        .Range(Cells(2, 1), Cells(.Range("A1048576").End(xlUp).Row, 5)).ClearContents ' clear existing data
        For Each Item In JSON
            .Cells(r, 1).value = JSON(s)("id")
            .Cells(r, 2).value = JSON(s)("name")
            .Cells(r, 3).value = JSON(s)("statusCategory")("name")
            .Cells(r, 4).value = JSON(s)("statusCategory")("colorName")
            .Cells(r, 5).value = JSON(s)("statusCategory")("key")
            s = s + 1 'increment the status
            r = r + 1 'increment the row
        Next
    End With
    ws_StatusData.Visible = xlSheetHidden
    ws_Sheet1.Activate
End If

Debug.Print (strDebug("GetAllStatuses", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function
Public Function PostTeamsFind(ByVal auth As String, ByVal baseUrl As String) As Long

Dim http, JSON, Team, Resource, Person As Object
Dim t, p, r, l As Integer
Dim JsonPost As String

JsonPost = "{" & Chr(34) & "maxResults" & Chr(34) & ":50}"

Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "POST", baseUrl & "rest/teams/1.0/teams/find", False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send (JsonPost)

PostTeamsFind = http.Status

If http.Status = 200 Then
    ws_TeamsData.Visible = xlSheetVisible
    ws_TeamsData.Activate
    Set JSON = ParseJson(http.responseText)
    t = 1 'reset the teams to 1
    r = 2
    With ws_TeamsData
        .Range(Cells(2, 1), Cells(.Range("A1048576").End(xlUp).Row, 5)).ClearContents ' clear existing data
        For Each Team In JSON("teams")
            p = 1
            For Each Resource In JSON("teams")(t)("resources")
                .Cells(r, 1).value = JSON("teams")(t)("id")
                .Cells(r, 2).value = JSON("teams")(t)("title")
                .Cells(r, 3).value = JSON("teams")(t)("resources")(p)("id")
                .Cells(r, 4).value = JSON("teams")(t)("resources")(p)("personId")
                p = p + 1
                r = r + 1
            Next Resource
            t = t + 1
        Next Team
        r = 2
        p = 1
        For Each Person In JSON("persons")
            .Cells(r, 5).value = JSON("persons")(p)("personId")
            .Cells(r, 6).value = JSON("persons")(p)("jiraUser")("jiraUsername")
            p = p + 1
            r = r + 1
        Next Person
    End With
    ws_TeamsData.Visible = xlSheetHidden
    ws_Sheet1.Activate
End If

Debug.Print (strDebug("PostTeamsFind", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetVelocity(ByVal rapidViewId, ByVal TeamId As String, _
                            ByVal auth As String, ByVal baseUrl As String) As Long

Dim http, JSON, Item As Object
Dim apicall As String
Dim r, s As Integer

apicall = "rest/greenhopper/latest/rapid/charts/velocity.json?rapidViewId="
             
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", baseUrl & apicall & rapidViewId, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

GetVelocity = http.Status

If http.Status = 200 Then
    
    Set JSON = ParseJson(http.responseText)
    Debug.Print http.responseText
    r = 2
    s = 1
    For Each Item In JSON("sprints")
        With ws_VelocityData
            .Cells(r, 1).value = TeamId
            .Cells(r, 2).value = JSON("sprints")(s)("id") ' SprintId
            .Cells(r, 3).value = JSON("sprints")(s)("name") 'SprintName
            .Cells(r, 4).value = JSON("sprints")(s)("sequence") 'SprintSequence
            .Cells(r, 5).value = JSON("velocityStatEntries")(CStr(.Cells(r, 2).value))("estimated")("value") 'Comitted
            .Cells(r, 6).value = JSON("velocityStatEntries")(CStr(.Cells(r, 2).value))("completed")("value") 'Completed
        End With
        r = r + 1
        s = s + 1
    Next
End If

Debug.Print (strDebug("GetVelocity", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetEpicBurndown(ByVal rapidViewId, ByVal EpicKey As String, _
                            ByVal auth As String, ByVal baseUrl As String) As Long

Dim http, JSON, time, change As Object
Dim apicall As String
Dim r, c, s As Integer
Dim JiraStories As New Dictionary
Dim oStory As clsJiraStory, storyKey As String
Dim TotalStoryPoints As Long
Dim key As Variant
Dim EpicColumn As Integer
Dim JiraSprints As New Dictionary
Dim oSprint As clsJiraSprint, sprintKey As Long

apicall = "rest/greenhopper/latest/rapid/charts/epicburndownchart.json?rapidViewId=" & rapidViewId & "&epicKey=" & EpicKey
             
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", baseUrl & apicall, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

GetEpicBurndown = http.Status

If http.Status = 200 Then
    
    Set JSON = ParseJson(http.responseText)
    
    ''Store the Epic's Stories in a new Dictionary
    For Each time In JSON("changes")
        For c = 1 To JSON("changes")(time).Count
            storyKey = JSON("changes")(time)(c)("key")
            If JiraStories.Exists(storyKey) Then
                Set oStory = JiraStories(storyKey)
            Else
                Set oStory = New clsJiraStory
                JiraStories.Add storyKey, oStory
            End If
            If JSON("changes")(time)(c).Exists("statC") Then
                If JSON("changes")(time)(c)("statC").Exists("newValue") Then
                    oStory.storypoints = JSON("changes")(time)(c)("statC")("newValue")
                End If
            End If
            If JSON("changes")(time)(c).Exists("added") Then
                oStory.StoryLinkedToEpic = JSON("changes")(time)(c)("added")
            End If
            If JSON("changes")(time)(c).Exists("column") Then
                If JSON("changes")(time)(c)("column").Exists("done") Then
                    oStory.StoryDone = JSON("changes")(time)(c)("column")("done")
                    If oStory.StoryDone Then
                        oStory.StorySprintCompleted = time 'store the epoch time the story was done
                    End If
                End If
            End If
        Next c
    Next
        
    ''Store the Sprints in a new Dictionary
    For s = 1 To JSON("sprints").Count
        sprintKey = JSON("sprints")(s)("id")
        Set oSprint = New clsJiraSprint
        JiraSprints.Add sprintKey, oSprint
        oSprint.jiraSprintName = JSON("sprints")(s)("name")
        oSprint.jiraSprintState = JSON("sprints")(s)("state")
        oSprint.jiraSprintStartTime = JSON("sprints")(s)("startTime")
        oSprint.jiraSprintEndTime = JSON("sprints")(s)("endTime")
    Next s
    
    ''Only update sprints for the first Epic
    '' This won't work if the Epic is done
    If EpicKey = ws_BurnUpData.Range("X1").value Then
        s = 1
        ''Find final sprint -- assumes that the first sprint id has been entered manually
        Do Until JSON("sprints")(s)("id") = ws_BurnUpData.Range("A28").End(xlUp).value
            s = s + 1
            If s > JSON("sprints").Count Then 'Safety net
                Exit Do
            End If
        Loop
        s = s + 1 'next sprint
        ''Add New Sprints
        If s <= JSON("sprints").Count Then
            r = ws_BurnUpData.Range("A28").End(xlUp).Row + 1
            Do While s <= JSON("sprints").Count
                ws_BurnUpData.Cells(r, 1).value = JSON("sprints")(s)("id")
                ws_BurnUpData.Cells(r, 2).value = JSON("sprints")(s)("name")
                r = r + 1
                s = s + 1
                If r >= 28 Then 'Safety net
                    Exit Do
                End If
            Loop
        End If
        '' Update Sprint State
        For r = 3 To 27
            If ws_BurnUpData.Cells(r, 1).value <> "" Then
                '' Check if the Sprint Exists in the Epic Burndown Report
                If JiraSprints.Exists(ws_BurnUpData.Cells(r, 1).value) Then
                    Set oSprint = JiraSprints(ws_BurnUpData.Cells(r, 1).value)
                    Select Case oSprint.jiraSprintState
                        Case "CLOSED"
                            ws_BurnUpData.Cells(r, 4) = "Complete"
                        Case "ACTIVE"
                            ws_BurnUpData.Cells(r, 4) = "Active"
                        Case Else
                            ws_BurnUpData.Cells(r, 4) = "Projected"
                    End Select
                End If
            Else
                ws_BurnUpData.Cells(r, 4) = "Projected"
            End If
        Next r
    End If
    
    ''Locate the EpicColumn
    EpicColumn = ws_BurnUpData.Range("1:1").Find(EpicKey, LookIn:=xlValues, LookAt:=xlWhole).Column
    ''Reset the previous StoryPoints in this column
    ws_BurnUpData.Range(Cells(3, EpicColumn), Cells(27, EpicColumn)).ClearContents
    ''Add up Story Points Delivered for Epic in each Sprint
    For r = 3 To 27
        If ws_BurnUpData.Cells(r, 1).value <> "" Then
            '' Check if the Sprint Exists in the Epic Burndown Report
            If JiraSprints.Exists(ws_BurnUpData.Cells(r, 1).value) Then
                Set oSprint = JiraSprints(ws_BurnUpData.Cells(r, 1).value)
                For Each key In JiraStories.Keys
                    Set oStory = JiraStories(key)
                    If oStory.StoryDone And oStory.StoryLinkedToEpic Then
                        If Val(oStory.StorySprintCompleted) >= oSprint.jiraSprintStartTime _
                            And Val(oStory.StorySprintCompleted) <= oSprint.jiraSprintEndTime Then
                                ws_BurnUpData.Cells(r, EpicColumn) = ws_BurnUpData.Cells(r, EpicColumn) + oStory.storypoints
                        End If
                    End If
                Next key
            End If
        End If
    Next r
    
    TotalStoryPoints = 0
    Set oSprint = JiraSprints(ws_BurnUpData.Cells(3, 1).value)
    For Each key In JiraStories.Keys
        Set oStory = JiraStories(key)
        If oStory.StoryDone And oStory.StoryLinkedToEpic Then
            If Val(oStory.StorySprintCompleted) >= oSprint.jiraSprintStartTime Then
                TotalStoryPoints = TotalStoryPoints + oStory.storypoints
            End If
        ElseIf oStory.StoryDone = False And oStory.StoryLinkedToEpic Then
            TotalStoryPoints = TotalStoryPoints + oStory.storypoints
        End If
    Next key
    '' Check if the aggregate of all Stories is great than or equal to the Epic's Story Points
    '' and enter the larger value as the total for the burnup
    If TotalStoryPoints >= ws_BurnUpData.Cells(29, EpicColumn + 27) Then
        ws_BurnUpData.Cells(28, EpicColumn + 27) = TotalStoryPoints
    Else
        ws_BurnUpData.Cells(28, EpicColumn + 27) = ws_BurnUpData.Cells(29, EpicColumn + 27)
    End If
        
End If

Debug.Print (strDebug("GetEpicBurndown", EpicKey, http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetRapidBoard(ByVal rapidViewId, ByVal auth As String, ByVal baseUrl As String) As Long

'' this data is never reset, just overridden

Dim http, JSON, Item As Object
Dim apicall As String
Dim f%, c%, s%, r%, l As Integer

apicall = "rest/greenhopper/1.0/rapidviewconfig/editmodel.json?rapidViewId="

'http://jira.unix.lch.com:8080/rest/greenhopper/1.0/rapidviewconfig/editmodel.json?rapidViewId=6533
             
Set http = CreateObject("MSXML2.XMLHTTP")

http.Open "GET", baseUrl & apicall & rapidViewId, False
http.setRequestHeader "Content-Type", "application/json"
http.setRequestHeader "X-Atlassian-Token", "no-check"
http.setRequestHeader "Authorization", "Basic " & auth
http.Send

GetRapidBoard = http.Status

If http.Status = 200 Then
    
    Set JSON = ParseJson(http.responseText)
    Debug.Print http.responseText
    With ws_BoardData
        .Cells(1, 2).value = rapidViewId ' id
        .Cells(2, 2).value = JSON("name") ' name
        .Cells(3, 2).value = JSON("filterConfig")("id") 'filter_id
        .Cells(4, 2).value = JSON("filterConfig")("name") ' filter_name
        .Cells(5, 2).value = JSON("filterConfig")("query") ' filter_query
        .Cells(6, 2).value = JSON("subqueryConfig")("subqueries")(1)("query") ' sub_query
        .Cells(7, 2).value = JSON("oldOneIssuesCutoff") ' old_issue_cut-off
        If JSON("isSprintSupportEnabled") Then 'Identify a Scrum Board
            .Cells(8, 2).value = "Scrum"
        Else
            .Cells(8, 2).value = "Kanban"
        End If
        If JSON("isKanPlanEnabled") Then .Cells(9, 2).value = "Kanban Backlog" 'Identify a Kanban Backlog is enabled
       
        If JSON("swimlanesConfig")("swimlanes").Count > 0 Then
            For l = 1 To JSON("swimlanesConfig")("swimlanes").Count
                .Cells(l + 1, 4).value = JSON("swimlanesConfig")("swimlanes")(l)("name") 'swimlaneName
                .Cells(l + 1, 5).value = JSON("swimlanesConfig")("swimlanes")(l)("query") 'swimlaneQuery
                .Cells(l + 1, 6).value = JSON("swimlanesConfig")("swimlanes")(l)("isDefault") 'swimlaneisDefault
            Next l
        End If
        
        If JSON("cardLayoutConfig")("currentFields").Count > 0 Then
            For f = 1 To JSON("cardLayoutConfig")("currentFields").Count
                .Cells(f + 1, 8).value = JSON("cardLayoutConfig")("currentFields")(f)("fieldId") ' extra_fieldId
                .Cells(f + 1, 9).value = JSON("cardLayoutConfig")("currentFields")(f)("name") ' extra_fieldName
                .Cells(f + 1, 10).value = JSON("cardLayoutConfig")("currentFields")(f)("mode") 'extra_fieldMode
            Next f
        End If
        
        r = 2
        If JSON("rapidListConfig")("mappedColumns").Count > 0 Then
            For c = 1 To JSON("rapidListConfig")("mappedColumns").Count
                For s = 1 To JSON("rapidListConfig")("mappedColumns")(c)("mappedStatuses").Count
                    .Cells(r, 12).value = JSON("rapidListConfig")("mappedColumns")(c)("name") ' columnName
                    .Cells(r, 13).value = JSON("rapidListConfig")("mappedColumns")(c)("mappedStatuses")(s)("id") ' statusId
                    r = r + 1
                Next s
            Next c
        End If
        
        
    End With
End If

Debug.Print (strDebug("GetRapidBoard", "", http.Status, http.getResponseHeader("X-AUSERNAME")))

End Function

Public Function GetAgileBoardView(ByVal filterId As Long, ByVal boardType As String, ByVal extraFields As String, _
                                    ByVal SubQuery As String, ByVal oldissueCutoff As String, _
                                    ByVal swimlaneRange As Range, ByVal columnRange As Range, _
                                    ByVal auth As String, ByVal baseUrl As String) As Long

'' function re-creates the Jira Agile board on a spreadsheet
'' It works by:
'' (1) identifying the columns and inserting them as a header row
'' (2a) identifying the number of swimlanes and looping through each one
'' (2b) for each swimlane a seperate JQL api search is called and issues added to the correct row and column

Dim http, JSON, Item As Object
Dim apicall As String
Dim apiFields As String
Dim swimlaneQuery As String
Dim defaultswimlaneQuery As String
Dim c%, r%, i As Integer
Dim outputRow As Long
Dim columnHeadings() As String
Dim swimlaneQueries() As String
Dim dict As Scripting.Dictionary

Set dict = New Scripting.Dictionary

'' The following gets the unique column headings from the columnRange supplied, by:
'  (a) storing the original values from column 1 of the range as an array of strings
'  (b) for each value in the range updating a dictionary item to a vlaue of empty
'  (c) this automatically and adds the dictionary item if if doesn't already exist
''
columnHeadings = columnRange.Columns(1).value
For c = 1 To UBound(columnHeadings)
    dict(columnHeadings(r)) = Empty
Next c

For c = 1 To dict.Count
    ws_Board.Cells(1, c).value = dict.Keys(c)
Next c

apiFields = "&fields=" _
        & "key," _
        & "summary," _
        & "issuetype," _
        & "assignee," _
        & "priority" _
        & extraFields

If boardType = "Scrum" Then
    SubQuery = " AND issuetype not in (Epic) AND (((Sprint is EMPTY OR Sprint in closedSprints())AND statusCategory not in (Done)) OR Sprint in (openSprints(),futureSprints()))"
ElseIf boardType = "Kanban" Then
    SubQuery = " AND " & SubQuery & " and resolutionDate=" & oldissueCutoff
End If
             
swimlaneQueries = swimlaneRange.Range(Cells(1, 2), Cells(swimlaneRange.Rows - 1, 2)).value
defaultswimlaneQuery = "NOT " & Join(swimlaneQueries, " AND ")
outputRow = 2
       
For r = 1 To swimlaneRange.Rows
    swimlaneQuery = swimlaneRange.Cells(r, 2).value
    If swimlaneQuery = "" Then swimlaneQuery = defaultswimlaneQuery
    apicall = baseUrl & "rest/api/2/search?jql=Filter=" & filterId & SubQuery & _
            " AND " & swimlaneQuery & " ORDER BY Rank" & apiFields & "&maxResults=1000"
    
    '' Need to test the api call value to ensure it has been built correctly
    Debug.Print apicall
    
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", apicall, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "X-Atlassian-Token", "no-check"
    http.setRequestHeader "Authorization", "Basic " & auth
    http.Send
    
    GetAgileBoardView = http.Status
    'MsgBox (http.Status & Chr(13) & http.responseText & Chr(13) & http.getAllResponseHeaders())
    
    If http.Status = 200 Then
        With ws_Board
            .Cells(outputRow, 1).value = swimlaneRange.Cells(r, 1)
            outputRow = outputRow + 1
            Set JSON = ParseJson(http.responseText)
            i = 1
            For Each Item In JSON("issues")
                .Cells(outputRow, 1).value = JSON("issues")(i)("key")
                '' This just puts the key in the first column
                ' need to calculate the column to put the key in the right one
                ' then need to expand the data to more than just the key
                outputRow = outputRow + 1
            Next Item
        End With
    End If
Next r
             
End Function
