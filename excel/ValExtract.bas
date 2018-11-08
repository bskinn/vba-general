Attribute VB_Name = "ValExtract"
Option Explicit
Const FLAG_DEV As Boolean = False

Sub DoValExtract()
Attribute DoValExtract.VB_ProcData.VB_Invoke_Func = "Q\n14"
    Dim rg As Range, cl As Range, cl2 As Range
    Dim mx As Double, mn As Double, avg As Double, std As Double
    Dim flr As Double, cei As Double, str As String
    Dim val As Long
    
    ' Query for the desired range of values
    str = InputBox("Input the range of values to include." & Chr(10) & Chr(10) & _
                "Use the format ""MIN=MAX"", where MIN or MAX (but not the ""="") can be omitted " & _
                "to give an open-ended range." & Chr(10) & Chr(10) & _
                "Entering just an equals sign will include all data." & Chr(10) & Chr(10) & _
                "A blank input, or entering letters, symbols, or spaces, will cancel the operation." _
                , "Enter Filter Range", "=")
                'Chr (10) & Chr(10) & "Be sure the a cell within the desired data subset as output " & _
                '"by the FlukeView ""Copy Data"" command is currently selected; otherwise, cancel and retry." _

    ' If there's no "=" or if str is empty, just drop out of the Sub
    If str = "" Or InStr(str, "=") < 1 Then Exit Sub
    
    ' Set default threshold values
    flr = -100000#
    cei = 100000#
    
    ' Attempt to assign the threshold values from the input string
    ' Begin in-place error handling
    On Error Resume Next
    
    ' Attempt to assign the floor bound if the "=" is not the first character
    If InStr(str, "=") > 1 Then flr = CDbl(Left(str, InStr(str, "=") - 1))
    
    ' Attempt to assign the ceiling bound if the "=" is not the last character
    If InStr(str, "=") < Len(str) Then cei = CDbl(Right(str, Len(str) - InStr(str, "=")))
    
    ' Record the error status, clear the error status, and restore system error handling
    val = Err.Number
    Err.Clear
    On Error GoTo 0
    
    ' Handle depending on the stored error number
    Select Case val
    Case 0      ' Do nothing, no error
    Case 13     ' Non-numeric input, complain and exit
        Call MsgBox("Invalid input!", vbOKOnly + vbExclamation, "Error")
        Exit Sub
    Case Else ' Some other error, re-raise
        Call Err.Raise(val)
    End Select
    
    ' Call Application.InputBox to select a cell within the desired range
    ' Start in-place error handling
    On Error Resume Next
    
    ' Query the user for the desired range
    Set rg = Application.InputBox("Select a cell within the desired data set: ", "Select data set", _
                        ActiveCell, , , , , 8)
    
    ' Store the error number, clear the error status, and restore system error handling
    val = Err.Number
    Err.Clear
    On Error GoTo 0
    
    ' Handle depending on the error number
    Select Case val
    Case 0      ' No error, proceed
    Case 424    ' User cancelled, exit sub
        Exit Sub
    Case Else   ' Unhandled error
        Err.Raise val
    End Select
    
    ' Rather than throw any errors in the following, just cancel out if they occur
    On Error GoTo ErrCatch
    
    ' Halt screen updating to speed code
    Application.ScreenUpdating = False
    
    ' If multiple cells were selected, search to find the first non-empty cell
    For Each cl In rg.Cells
        If cl.Formula <> "" Then
            Set rg = cl
            Exit For
        End If
    Next cl
    
    ' If rg.Formula now is empty, there's a problem; error and exit
    If rg.Formula = "" Then
        Call MsgBox("No data found in selected range!", vbOKOnly + vbExclamation, _
                    "Error")
        Exit Sub
    End If
    
    ' Expand rg to CurrentRegion and unset bold
    Set rg = rg.CurrentRegion
    rg.Font.Bold = False
    
    ' Filter for nonnumeric values; subst zeroes for any such
    For Each cl In rg.Cells
        If Not IsNumeric(cl.Value) Then cl.Value = 0#
    Next cl
    
    ' The following depends on the particular formatting of the FlukeView data output, and is
    '  handled differently depending on whether the data was output as double-value or single-value
    If rg.Columns.Count = 3 Then        ' Double-value data detected
        ' Insert two columns to allow for calculating the means and for the final value
        Call rg.Offset(0, 3).Resize(, 2).Insert(xlShiftToRight)
        
        ' Compute and store the mean of each data pair in the first of the new columns
        For Each cl In rg.Columns(3).Cells
            cl.Offset(0, 1).Value = CDbl((cl.Value + cl.Offset(0, -1).Value)) / 2#
        Next cl
        
        ' Reset rg to the newly-calculated column of means
        Set rg = rg.Offset(0, 3).Resize(rg.Rows.Count, 1)
    Else        ' Single-value data detected
        ' Add two columns to the right of rg, for a parsing column and the final value
        '  (If the parsing column is not added, the original data is destroyed by the filtering process)
        Call rg.Offset(0, 2).Insert(xlShiftToRight)
        
        ' Straight-copy of the data values to the new column
        For Each cl In rg.Columns(2).Cells
            cl.Offset(0, 1).Value = CDbl(cl.Value)
        Next cl
        
        ' Reset rg to the column of newly-copied means
        Set rg = rg.Offset(0, 2).Resize(rg.Rows.Count, 1)
    End If
    
    ' Set cl2 as a calculation cell
    Set cl2 = rg.Offset(0, 1).Resize(1, 1)
    
    ' Filter the cells in rg based on the input floor and ceiling
    '  The defaults for flr and cei should be large enough not to interfere with most data
    For Each cl In rg.Cells
        cl.Activate
        If cl.Value < flr Or cl.Value > cei Then
            cl.Clear
        End If
    Next cl
    
    ' Only do the following if FLAG_DEV flag active
    If FLAG_DEV Then
        ' Calculate and store the average and stdev of the filtered data
        cl2.Formula = "=AVERAGE(" & rg.Address & ")"
        avg = cl2.Value
        cl2.Formula = "=STDEV(" & rg.Address & ")"
        std = cl2.Value
        
        ' Calculate the max and min thresholds for the filtering process at mean +/- 2*stdev
        mx = avg + 2 * std
        mn = avg - 2 * std
        
        ' Clear any value in the filtering column that exceeds the bounds
        For Each cl In rg
            If cl.Value > mx Or cl.Value < mn Then cl.Clear
        Next cl
    End If ' FLAG_DEV
    
    ' Reset the calculation cell to the average of the filtered data and format appropriately
    With cl2
        .Formula = "=AVERAGE(" & rg.Address & ")"
        .NumberFormat = "General"
        .Font.Bold = True
    End With
    With cl2.Offset(0, 1)
        .Formula = "Mean"
        .Font.Bold = True
    End With
    cl2.Offset(1, 0).Formula = flr
    cl2.Offset(1, 1).Formula = "Floor"
    cl2.Offset(2, 0).Formula = cei
    cl2.Offset(2, 1).Formula = "Ceiling"
    
    
    ' Restore screen updating and exit
    Application.ScreenUpdating = True
    Exit Sub
    
ErrCatch:
    ' Restore screen updating, notify, and exit
    Application.ScreenUpdating = True
    Call MsgBox("Unhandled error occurred", vbOKOnly, "Notification")
    Exit Sub
    
End Sub
