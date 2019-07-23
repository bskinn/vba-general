Attribute VB_Name = "ColsToMatrix"
Option Explicit

Public Sub DoColumnsToMatrix()
Attribute DoColumnsToMatrix.VB_Description = "Run columns-to-matrix conversion tool."
Attribute DoColumnsToMatrix.VB_ProcData.VB_Invoke_Func = "M\n14"
    
    Dim refCel As Range, origCel As Range, workRg As Range
    Dim resDel As VbMsgBoxResult
    Dim resAutoFind As VbMsgBoxResult
    'Dim firstChunk As Boolean
    
    resAutoFind = MsgBox("Autofind starting point?", vbYesNoCancel + vbQuestion)
    If resAutoFind = vbCancel Then Exit Sub
    
    resDel = MsgBox("Delete x-coordinate data?", vbYesNoCancel + vbQuestion)
    If resDel = vbCancel Then Exit Sub
    
    If resAutoFind = vbYes Then
        Set origCel = ActiveSheet.UsedRange.Resize(1, 1).Offset(1, 0)
        origCel.Select
    Else
        Set origCel = ActiveCell
    End If
    
    Set refCel = origCel
    
    Do While Not Intersect(ActiveCell, ActiveSheet.UsedRange) Is Nothing
        ActiveCell.Offset(1, 0).Select
        If ActiveCell.Value < ActiveCell.Offset(-1, 0).Value Then
            ' x cycled around, move data
            ' Just hack off everything below, including some blank space
            ActiveCell.Resize(ActiveSheet.UsedRange.Rows.Count, 2).Select
            Set workRg = Selection
            workRg.Copy
            
            ' Need to paste two columns over if either retaining the x
            ' data each time, or if it's the first chunk moved when
            ' not retaining the x data
            refCel.Offset(0, 2).PasteSpecial xlPasteValues
            
            ' Update the location of the ref cell
            Set refCel = refCel.Offset(0, 2)
            
            refCel.Select
            
            workRg.Clear
            
        End If
        
    Loop
    
    ' Cull extra stuff
    If resDel = vbYes Then
        origCel.Offset(0, 2).Select
        Do While Not Intersect(ActiveCell, ActiveSheet.UsedRange) Is Nothing
            ActiveCell.EntireColumn.Delete
            ActiveCell.Offset(0, 1).Select
        Loop
    End If
    
End Sub

