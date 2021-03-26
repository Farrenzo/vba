Public Sub Run()

Application.ScreenUpdating = False
Application.Calculation = xlManual

    'delete everything in blank sheet
    Application.Run ("ClearData")
    Application.Run ("FileLookUp")

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic

End Sub

Public Sub Export()

Application.ScreenUpdating = False
Application.Calculation = xlManual

    Application.Run ("SaveMergedDataToNewWorkbook")

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic

End Sub

Sub ClearData()

    Dim lR As String
    Sheets("sample").Select
    lR = Cells(Rows.Count, "F").End(xlUp).Row
    Range("A2", "F" & lR).Select
    Selection.Delete Shift:=xlToLeft

    Sheets("blank").Select
    Columns("A:Z").Select
    Selection.Delete Shift:=xlToLeft
    Range("A1").Select

End Sub

Sub FileLookUp()

    Dim i As Integer
    Dim lR As Long
    Dim rng1, rng2 As Range
    Dim rdFilePath As String
    Dim wb1, wb2 As Workbook
    Dim oFSO, oFolder, oFile As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    rdFilePath = Application.ActiveWorkbook.Path + "\rd\"
    Set oFolder = oFSO.GetFolder(rdFilePath)
    Set wb1 = ActiveWorkbook
    
        For Each oFile In oFolder.Files
            rdPathNames = rdFilePath + oFile.Name
            wb1.Sheets("blank").Cells(i + 1, 1) = rdPathNames
            
            Workbooks.Open Filename:=rdPathNames, ReadOnly:=True
            Set wb2 = ActiveWorkbook
            Application.Run ("GrabData")
            
            wb1.Activate
            Sheets("Sample").Select
            lR = Cells(Rows.Count, "A").End(xlUp).Row
            Cells(lR + 1, 1).Select
            ActiveSheet.Paste
            Application.CutCopyMode = False
            wb2.Close
            i = i + 1
        Next oFile
    
    Application.Run ("DuplicateRowDelete")

End Sub

Sub GrabData()

    ActiveSheet.Select
    Range("A2:F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
End Sub

Sub DuplicateRowDelete()

    Sheets("sample").Select
    Range("A1:F1").Select
    Range(Selection, Selection.End(xlDown)).RemoveDuplicates Columns:=Array(1, 2, 3, 4, 5, 6), Header:=xlYes
    
End Sub

Sub SaveMergedDataToNewWorkbook()

    Dim fName As String
    fName = Application.ActiveWorkbook.Path & "\" & Worksheets("Run").Cells(9, 7).Value
    
    ActiveWorkbook.Save
    
    Sheets(Array("Sample")).Copy
    Sheets(Array("Sample")).Select
    
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("A1").Select

    ActiveWorkbook.SaveAs Filename:=fName, FileFormat:=51, CreateBackup:=False
    ActiveWindow.Close SaveChanges:=False
    
End Sub