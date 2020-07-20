Attribute VB_Name = "PublicVariables"
'' CMS ''
Public Const RapidBoardId As Long = 6533
Public Const planId As Integer = 522
Public Const scenarioId As Integer = 602
Public Const TeamId As String = "81"

'' Team Two ''
'Public Const RapidBoardId As Long = 6553
'Public Const planId As Integer = 741
'Public Const scenarioId As Integer = 841
'Public Const TeamId As String = "41"

'' Team Three ''
'Public Const RapidBoardId As Long = 6534
'Public Const planId As Integer = 721
'Public Const scenarioId As Integer = 821
'Public Const TeamId As String = "42"

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


          


Option Explicit

Function JiraBaseUrl()
    Select Case ws_Sheet1.Range("Jira_Environment")
    Case "Test"
        JiraBaseUrl = ws_Sheet1.Range("Jira_Base_Urls").Cells(2, 1)
    Case "Production"
        JiraBaseUrl = ws_Sheet1.Range("Jira_Base_Urls").Cells(1, 1)
    Case Else
        JiraBaseUrl = ws_Sheet1.Range("Jira_Base_Urls").Cells(2, 1)
    End Select
End Function

Sub Revive()

Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub

Function saveJsonToFile(ByVal strPath As String, JsonTxt As String) As Boolean
' Use this function to Save the Json to File
' Sometimes the Json is to large and this does not work, so
' I have used a function to return the success or failure of the code

'Example strPath = "M:\documents\Downloads\jql.jason"

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
' result is GMT/UTC
Function epoch2date(myEpoch)
epoch2date = DateAdd("s", myEpoch / 1000, "01/01/1970 00:00:00")
End Function
Function date2epoch(myDate)
date2epoch = DateDiff("s", "01/01/1970 00:00:00", myDate) * 1000
End Function
Function JiraLookUp(ByVal LookupRange As Range, LookupData As Variant, OffsetColumns As Integer)
    If LookupRange.Find(LookupData, LookAt:=xlWhole) Is Nothing Then
        JiraLookUp = "Unknown (" & LookupData & ")"
    Else
        JiraLookUp = LookupRange.Find(LookupData, LookAt:=xlWhole).Offset(0, OffsetColumns)
    End If
End Function

' https://excelmacromastery.com/
Public Function SortDictionaryByValue(dict As Object _
                    , Optional sortorder As XlSortOrder = xlAscending) As Object
    
    On Error GoTo eh
    
    Dim arrayList As Object
    Set arrayList = CreateObject("System.Collections.ArrayList")
    
    Dim dictTemp As Object
    Set dictTemp = CreateObject("Scripting.Dictionary")
   
    ' Put values in ArrayList and sort
    ' Store values in tempDict with their keys as a collection
    Dim key As Variant, value As Variant, coll As Collection
    For Each key In dict
    
        value = dict(key)
        
        ' if the value doesn't exist in dict then add
        If dictTemp.Exists(value) = False Then
            ' create collection to hold keys
            ' - needed for duplicate values
            Set coll = New Collection
            dictTemp.Add value, coll
            
            ' Add the value
            arrayList.Add value
            
        End If
        
        ' Add the current key to the collection
        dictTemp(value).Add key
    
    Next key
    
    ' Sort the value
    arrayList.Sort
    
    ' Reverse if descending
    If sortorder = xlDescending Then
        arrayList.Reverse
    End If
    
    dict.RemoveAll
    
    ' Read through the ArrayList and add the values and corresponding
    ' keys from the dictTemp
    Dim Item As Variant
    For Each value In arrayList
        Set coll = dictTemp(value)
        For Each Item In coll
            dict.Add Item, value
        Next Item
    Next value
    
    Set arrayList = Nothing
    
    ' Return the new dictionary
    Set SortDictionaryByValue = dict
        
Done:
    Exit Function
eh:
    If Err.Number = 450 Then
        Err.Raise vbObjectError + 100, "SortDictionaryByValue" _
                , "Cannot sort the dictionary if the value is an object"
    End If
End Function

Public Sub PrintDictionary(ByVal sText As String, dict As Object)
    
    Debug.Print vbCrLf & sText & vbCrLf & String(Len(sText), "=")
    
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print key, dict(key)
    Next key
    
End Sub

Public Sub CleanUpDataLabelsOnBurnDownChart()

' This sub changes with width of the Epic datalabels on the BurnDown Chart to 1cm
' Which realigns then to the right of the final target column

' This setting does not get saved with the workbook so need to be updated each time
' the workbook is opened

Dim bd_Series As Integer

For bd_Series = 4 To 30
    ch_BurnDown.FullSeriesCollection(bd_Series).Points(26).DataLabel.Width = 28.3464566929134
Next bd_Series

End Sub

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
