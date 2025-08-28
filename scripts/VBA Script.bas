Attribute VB_Name = "Module1"
'Button VBA
Sub RefreshWithVBA()
'prompts the updateFiles VBA
    Call updateFiles
    ThisWorkbook.RefreshAll 'runs the Powerquery script after data extract by the VBA
End Sub

'updateFiles VBA
Sub updateFiles()
    Dim folderPath As String
    Dim wb As Workbook, over As Worksheet, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim WO As String, Month As String
    Dim carMakeValue As String, carMakeCell As Range
    Dim latestFile As String, latestDate As Date, tempFile As String, tempDate As Date
 
'get the folder path
onedrive = Environ("OneDrive") 'just makes a dynamic one drive path
folderPath = onedrive & "\Desktop\tryout\"

'Loop through the files in the folder and find the latest one
latestFile = ""
latestDate = 0
tempFile = Dir(folderPath & "*.xlsx*")

Do While tempFile <> ""
    tempDate = FileDateTime(folderPath & tempFile) ' this now becomes our latest modified date.
        If tempDate > latestDate Then
            latestDate = tempDate
            latestFile = tempFile
        End If
        tempFile = Dir
Loop

'Disable alerts,just for good practices!
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'open the latest file
    Set wb = Workbooks.Open(folderPath & latestFile)
    Set over = Nothing
    
    On Error Resume Next ' just for good practices!
    Set over = wb.Sheets("OVER") 'locate the summary sheet with the workorder numbers in each workbook
    On Error GoTo 0
    
    If Not over Is Nothing Then 'if the over sheet is available in the file
        lastRow = over.Cells(over.Rows.Count, "B").End(xlUp).Row 'find the last in column B that contains a value and return the number of that row
        
        'Create the car make column and Month In columns
        If over.Cells(1, 11).Value = "" Then over.Cells(1, 11).Value = "Make"
        If over.Cells(1, 12).Value = "" Then over.Cells(1, 12).Value = "Month In"

        'loop through each workorder in the over sheet (summary sheet)
        For i = 2 To lastRow
            WO = over.Cells(i, 2).Value 'locate the work order name in the current cell on the over sheet
            
            On Error Resume Next
            Set ws = wb.Sheets(WO) 'since the WO# in the WorkOrder columns on the over sheet is the same as the WO sheets in the file,then use it to locate those sheets.
            On Error GoTo 0
        
            If Not ws Is Nothing Then 'if the worksheet exists find the cell with car make year
                Set carMakeCell = ws.UsedRange.Find(What:="Car Model.", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
                
                If Not carMakeCell Is Nothing Then 'if the cell containing the make is found, then get the value adjustcent to it
                carMakeValue = carMakeCell.Offset(0, 1).Value
            
                'write car make value in the make column
                over.Cells(i, 11).Value = carMakeValue
                'write month in value in the month in column
                over.Cells(i, 12).Value = Mid(latestFile, 3, 3)
                End If
            Else
            'over.Cells(i, 11).Value = "Car Make Not found"
            End If
        Next i
        'save the update
        wb.Save
    End If
    wb.Close SaveChanges:=False
    
'restore alerts
Application.DisplayAlerts = True
Application.ScreenUpdating = True

MsgBox "Update complete for the latest file: " & latestFile
End Sub

