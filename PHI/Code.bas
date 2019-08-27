Attribute VB_Name = "Module2"
'Code to identify Presence of Protected Health Information on System

Public Sub PHISCANNER()
    Dim fso, oFolder, oSubfolder, oFile, queue As Collection
    Dim k As String
    Dim s As String
    Dim count As Integer
    Dim newRow As Long
    Dim a As Integer
    a = MsgBox("Do you want to highlight the cells with possibility of PHI?", 4, "Choose Options")
        
' creating new worksheet for log
    For Each WS In Worksheets
    If WS.Name = "PHI_Found" Then
    Sheets("PHI_Found").Delete
    End If
    Next
    Sheets.Add.Name = "PHI_Found"
    Cells(1, 1) = "File Path"
    Cells(1, 2) = "File Name"
    Cells(1, 3) = "Sheet Name"
'Activating the worksheet with code
    
    
'Going through subfolders
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set queue = New Collection
  
   Worksheets("Code").Activate
    Range("J2").Select
    queue.Add fso.GetFolder(ActiveCell.Value)
    Do While queue.count > 0
        Set oFolder = queue(1)
        queue.Remove 1 'dequeue
        For Each oSubfolder In oFolder.subfolders
        queue.Add oSubfolder 'enqueue
        Next oSubfolder
        For Each oFile In oFolder.Files
        If Right(oFile, 5) = ".xlsx" Or Right(oFile, 4) = ".csv" Then
        Workbooks.Open oFile

'Going through all sheets
Dim wcount As Integer
Dim z As Integer

wcount = ActiveWorkbook.Worksheets.count
For z = 0 To wcount - 1
If (z < wcount) Then
Worksheets(z + 1).Activate
End If
R = Cells(Rows.count, 1).End(xlUp).Row
C = Cells(1, Columns.count).End(xlToLeft).Column
p = ActiveWorkbook.Name
count = 0
x = 0
'Checking for Column Names
For j = 1 To C
If InStr(1, Cells(1, j), "MRN") <> 0 Or InStr(1, Cells(1, j), "Fin") <> 0 Then
If a = 6 Then
Cells(1, j).Interior.ColorIndex = 38
End If
x = x + 1
End If
Next j
'the loop for cell check starts here
If x > 0 Then

For i = 2 To R
For j = 1 To C

If IsNumeric(Cells(i, j)) Then
If Cells(i, j) = Round(Cells(i, j)) Then
If Len(Cells(i, j)) = 8 Then
If a = 6 Then
Cells(i, j).Interior.ColorIndex = 34
End If
count = count + 1

Else
If Len(Cells(i, j)) = 12 Then
If a = 6 Then
Cells(i, j).Interior.ColorIndex = 36
End If
count = count + 1

End If
End If
End If
End If
Next j
Next i
End If
'Adding the location of PHI to log file
If (x > 0) Then
k = ActiveSheet.Name
Workbooks("Code3.xlsm").Activate
Worksheets("PHI_Found").Activate
newRow = Cells(Rows.count, 1).End(xlUp).Row + 1
Cells(newRow, 1) = oFile.Path
Cells(newRow, 2) = oFile.Name
Cells(newRow, 3) = k
Workbooks(p).Activate
End If
Next z
ActiveWorkbook.Close True
End If
Next oFile
Loop
End Sub
  
