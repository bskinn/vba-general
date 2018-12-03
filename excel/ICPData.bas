Attribute VB_Name = "ICPData"
Option Explicit

' Column positions and offsets as constants; will break if structure of Plasma96 data export changes
Const nameCol As Long = 1, typeCol As Long = 2, lineCol As Long = 3, concCol As Long = 7
Const intCol As Long = 9, timeCol As Long = 11, intRSDCol As Long = 15
Const factorOffset As Long = 1, adjConcOffset As Long = 2, IDOffset As Long = 3
Const sortIDOffset As Long = 5, sortOrigOffset As Long = 6, sortAdjOffset As Long = 7
Const sortIntRSDOffset As Long = 8

' These are offsets relative to the sorted range, not the QueryTable
Const summIDOffset As Long = 5, summConcOffset As Long = 6, summVarOffset As Long = 7
Const summRSDOffset As Long = 8, summUncOffset As Long = 9

' Long constants for the colors to use to mark standards, samples, and std replicates
Const stdColor As Long = 14277081, splColor As Long = 14610923, repColor As Long = 14281213


Public Sub HandleICPDataImport()
Attribute HandleICPDataImport.VB_Description = "Automatic import and background drift processing of exported data from the Plasma96 software controlling Varian Liberty II radial ICP-OES."
Attribute HandleICPDataImport.VB_ProcData.VB_Invoke_Func = "I\n14"
' Wrapper sub for the data import and formatting process
    Dim destCell As Range, filePath As String
    Dim qt As QueryTable
    Dim cel As Range, lineSpan As Range
    
    ' Get target cell
    Set destCell = GetTgtCell
    If destCell Is Nothing Then Exit Sub
    
    ' Get desired file
    filePath = GetDataFile
    If filePath = "" Then Exit Sub
    
    ' Perform data import
    Set qt = doDataImport(destCell, filePath)
    If qt Is Nothing Then ' some sort of error; notify and exit
        Call MsgBox("Retrieval of file data failed!  Canceling...", vbOKOnly + vbCritical, "Error")
        Exit Sub
    End If
    
    ' Perform sorting on data import range
    With qt.ResultRange
        Call .Sort(Header:=xlYes, key1:=.Cells(1, lineCol), Key2:=.Cells(1, timeCol), Order1:=xlAscending, _
                        Order2:=xlAscending, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal)
    End With
        
    ' Under-border divisions between element lines
    For Each cel In qt.ResultRange.Columns(lineCol).Cells
        If cel.Formula <> cel.Offset(1, 0).Formula Then
            With cel.EntireRow.Borders(xlEdgeBottom)
                .Color = RGB(0, 0, 0)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        End If
    Next cel
    
    ' Bold Label, Corr Conc, and Int columns
    With qt.ResultRange.Rows(1)
        .Cells(1, nameCol).EntireColumn.Font.Bold = True ' Sample Label
        .Cells(1, concCol).EntireColumn.Font.Bold = True ' Corr Conc
        .Cells(1, intCol).EntireColumn.Font.Bold = True ' Int
    End With
    
    ' Set freeze panes (to right of 'Int' for now; later just to right of 'Element'
'    qt.ResultRange.Worksheet.Activate
'    qt.ResultRange.Cells(2, intCol + 1).Select
'    ActiveWindow.FreezePanes = True
'    With ActiveWindow
'        .Split = True
'        .SplitColumn = 9
'        .SplitRow = 1
'        .FreezePanes = True
'    End With
    
'    Exit Sub
    
    ' Identify blocks of result range that are data from a single element line and add
    '  drift correction
    ' Add headers for drift-corrected and sorted data
    With qt.ResultRange.Offset(0, qt.ResultRange.Columns.Count).Resize(1, 1)
        .Offset(0, factorOffset).Formula = "Factor"
        .Offset(0, adjConcOffset).Formula = "Adj Conc"
        .Offset(0, IDOffset).Formula = "Trimmed ID"
        .Offset(0, sortIDOffset).Formula = "Sorted IDs"
        .Offset(0, sortOrigOffset).Formula = "Orig Conc"
        .Offset(0, sortAdjOffset).Formula = "Adj Conc"
        .Offset(0, sortIntRSDOffset).Formula = "Int RSD"
    End With
    
    ' Set initial position of lineSpan to 'Element' header
    Set lineSpan = qt.ResultRange.Cells(1, lineCol)
    
    ' Repeat until running off the bottom of the result range
    Do Until Intersect(qt.ResultRange, lineSpan.Offset(lineSpan.Rows.Count + 1, 0)) Is Nothing
        ' Bump lineSpan down to the row following its current extent and resize
        '  to a single cell
        Set lineSpan = lineSpan.Offset(lineSpan.Rows.Count, 0).Resize(1, 1)
        ' Extend lineSpan until it encounters the end of the data from the current element line
        Do Until lineSpan.Cells(1, 1).Formula <> lineSpan.Cells(lineSpan.Rows.Count + 1, 1).Formula
            Set lineSpan = lineSpan.Resize(lineSpan.Rows.Count + 1, 1): lineSpan.Select
        Loop
        
        ' Pass result range and lineSpan into Sub for adding the drift corrections
        Call insertDriftCorr(qt.ResultRange, lineSpan)
    Loop
    
    ' Freeze panes to top row only; activate sorted results summary
    With qt.ResultRange
        .Worksheet.Rows(2).Select
        .Application.ActiveWindow.FreezePanes = True
        .Offset(0, .Columns.Count + 2).Resize(1, 1).Select
    End With
    
End Sub

Public Sub insertDriftCorr(queryRg As Range, lineRg As Range)
' Takes in Range covering full query (queryRg = qt.ResultRange) and Range
'  situated on the 'Element' column and spanning the full number of rows
'  of data for a single element line.

    Dim splTypeRg As Range, splNameRg As Range, timeRg As Range
    Dim intRg As Range, concRg As Range, intRSDRg As Range
    Dim factorRg As Range, adjConcRg As Range, IDRg As Range
    Dim sortIDRg As Range, sortOrigRg As Range, sortAdjRg As Range
    Dim sortIntRSDRg As Range
    Dim iter As Long, iter2 As Long, proofStr As String
    Dim wsf As WorksheetFunction, cel As Range
    Dim stdNames As Variant
    
    ' Attach the WorksheetFunction object
    Set wsf = Application.WorksheetFunction
    
    ' Null-initialize the standard names array
    stdNames = Array("")
    
    ' Bind information ranges
    Set splNameRg = Intersect(queryRg.Columns(nameCol), lineRg.EntireRow)
    Set splTypeRg = Intersect(queryRg.Columns(typeCol), lineRg.EntireRow)
    Set timeRg = Intersect(queryRg.Columns(timeCol), lineRg.EntireRow)
    Set intRg = Intersect(queryRg.Columns(intCol), lineRg.EntireRow)
    Set concRg = Intersect(queryRg.Columns(concCol), lineRg.EntireRow)
    Set intRSDRg = Intersect(queryRg.Columns(intRSDCol), lineRg.EntireRow)
    
    ' Bind factor, adjusted conc, ID; ID sort, sorted orig conc and sorted adj conc ranges
    With queryRg.Columns
        Set factorRg = Intersect(.Item(factorOffset).Offset(0, .Count + 1), lineRg.EntireRow)
        Set adjConcRg = Intersect(.Item(adjConcOffset).Offset(0, .Count + 1), lineRg.EntireRow)
        Set IDRg = Intersect(.Item(IDOffset).Offset(0, .Count + 1), lineRg.EntireRow)
        Set sortIDRg = Intersect(.Item(sortIDOffset).Offset(0, .Count + 1), lineRg.EntireRow)
        Set sortOrigRg = Intersect(.Item(sortOrigOffset).Offset(0, .Count + 1), lineRg.EntireRow)
        Set sortAdjRg = Intersect(.Item(sortAdjOffset).Offset(0, .Count + 1), lineRg.EntireRow)
        Set sortIntRSDRg = Intersect(.Item(sortIntRSDOffset).Offset(0, .Count + 1), lineRg.EntireRow)
    End With
    
    ' Proofread names to be sure all are search-unique save for " R#" replicates; report and exit if not
    proofStr = proofSplNames(splNameRg)
    If Not proofStr = "" Then
        Call MsgBox("Sample names are improperly defined for automatic drift correction " & _
                    "and data workup of element line " & lineRg.Cells(1, 1).Text & ":" & _
                    Chr(10) & Chr(10) & proofStr & Chr(10) & Chr(10) & "Skipping...", _
                    vbOKOnly + vbInformation, "Skipping drift correction")
        Exit Sub
    End If
    
    ' Apply highlighting to factor range and store the names of the calib standards.
    ' Insert formulas for replicate drift factors.
    ' Must insert interpolation formulas in a separate pass since all replicates must be
    '  marked before interpolation formula detection can succeed.
    For iter = 1 To lineRg.Rows.Count
        Select Case UCase(wsf.Trim(splTypeRg.Cells(iter, 1)))
        Case "STD", "BLK" ' Calib standard or blank
            ' Apply 'standard' fill color
            factorRg.Cells(iter, 1).Interior.Color = stdColor
            ' Store the standard name for later use
            If stdNames(UBound(stdNames)) <> "" Then
                ' Only expand the names array if it already has had a name assigned to it
                ReDim Preserve stdNames(LBound(stdNames) To UBound(stdNames) + 1)
            End If
            ' Store the trimmed standard name
            stdNames(UBound(stdNames)) = wsf.Trim(splNameRg.Cells(iter, 1))
            ' Drop a unity value into the factor cell
            factorRg.Cells(iter, 1) = 1
        Case Else ' Sample or duplicate, primarily?
            ' Must check whether it's a replicate of a standard or a sample. Since names have already
            '  been proofed for uniqueness, it should be sufficient just to check an InStr call
            For iter2 = LBound(stdNames) To UBound(stdNames)
                If stdNames(UBound(stdNames)) <> "" And _
                        InStr(wsf.Trim(splNameRg.Cells(iter, 1)), stdNames(iter2)) > 0 Then
                    ' Standard name found; apply standard replicate color & insert corr factor formula
                    With factorRg.Cells(iter, 1)
                        .Interior.Color = repColor
                        .Formula = "=" & intRg.Cells(iter, 1).Address(False, False) & "/" & _
                                    Intersect(splNameRg.Find( _
                                                        What:=stdNames(iter2), _
                                                        After:=splNameRg.Cells(splNameRg.Rows.Count, 1), _
                                                        LookIn:=xlFormulas, _
                                                        LookAt:=xlPart, _
                                                        SearchOrder:=xlByRows, _
                                                        SearchDirection:=xlNext, _
                                                        MatchCase:=False _
                                                            ).EntireRow, intRg).Address(False, False)
                    End With
                End If
            Next iter2
            If Not factorRg.Cells(iter, 1).Interior.Color = repColor Then
                ' Sample appears not to be a calib standard replicate; apply sample color
                factorRg.Cells(iter, 1).Interior.Color = splColor
            End If
        End Select
    Next iter
    
    ' Add sample interpolation formulas, adjusted concentration formulas, trimmed name formulas
    For iter = 1 To lineRg.Rows.Count
        If factorRg.Cells(iter, 1).Interior.Color = splColor Then
            Call insertInterpFormula(factorRg.Cells(iter, 1), timeRg)
            adjConcRg.Cells(iter, 1).Formula = "=" & concRg.Cells(iter, 1).Address(False, False) & _
                            "/" & factorRg.Cells(iter, 1).Address(False, False)
            IDRg.Cells(iter, 1).Formula = "=TRIM(" & splNameRg.Cells(iter, 1).Address(False, False) & ")"
        End If
    Next iter
    
    ' Retrieve and sort nonblank/nonempty IDs
    ' Drop out of the sub if IDRg is all blank (no samples actually measured)
    If wsf.CountBlank(IDRg) = IDRg.Rows.Count Then Exit Sub
    
    ' Resize the sorted data ranges
    Set sortIDRg = sortIDRg.Resize(IDRg.Rows.Count - wsf.CountBlank(IDRg), 1)
    Set sortOrigRg = sortOrigRg.Resize(IDRg.Rows.Count - wsf.CountBlank(IDRg), 1)
    Set sortAdjRg = sortAdjRg.Resize(IDRg.Rows.Count - wsf.CountBlank(IDRg), 1)
    Set sortIntRSDRg = sortIntRSDRg.Resize(IDRg.Rows.Count - wsf.CountBlank(IDRg), 1)
    
    ' Shift sorted ranges down to start at first nonstandard/nonreplicate row
    ' Find offset
    iter = 1
    Do Until IDRg.Cells(iter, 1).Text <> ""
        iter = iter + 1
    Loop
    iter = iter - 1 ' Correction for offset overcount
    
    ' Apply offset to the three sorted ranges
    Set sortIDRg = sortIDRg.Offset(iter, 0)
    Set sortOrigRg = sortOrigRg.Offset(iter, 0)
    Set sortAdjRg = sortAdjRg.Offset(iter, 0)
    Set sortIntRSDRg = sortIntRSDRg.Offset(iter, 0)
    
    ' Coerce sortIDRg to text-only
    sortIDRg.NumberFormat = "@"
    
    ' Text-copy the trimmed IDs to sortIDRg for later sorting; insert INDEX/MATCH formulas
    '  for original and adjusted concentrations
    iter = 0 ' Placeholder iterator for location in IDRg
    For iter2 = 1 To sortIDRg.Rows.Count
        ' Scan down to the next nonblank entry in IDRg
        Do
            iter = iter + 1
        Loop Until IDRg.Cells(iter, 1).Text <> ""  ' Might be better to use .Formula ?
        
        ' Copy the TEXT VALUE into the sort range
        sortIDRg.Cells(iter2, 1).Formula = IDRg.Cells(iter, 1).Text
        
        ' Insert the INDEX/MATCH formulas for original and adjusted concentrations
        sortOrigRg.Cells(iter2, 1).Formula = "=INDEX(" & concRg.Address(True, True) & ",MATCH(" & _
                sortIDRg.Cells(iter2, 1).Address(False, True) & "," & IDRg.Address(True, True) & ",0))"
        sortAdjRg.Cells(iter2, 1).Formula = "=INDEX(" & adjConcRg.Address(True, True) & ",MATCH(" & _
                sortIDRg.Cells(iter2, 1).Address(False, True) & "," & IDRg.Address(True, True) & ",0))"
        ' Int RSD values all divided by 100 to convert to decimal from (implied) percent format
        sortIntRSDRg.Cells(iter2, 1).Formula = "=INDEX(" & intRSDRg.Address(True, True) & ",MATCH(" & _
                sortIDRg.Cells(iter2, 1).Address(False, True) & "," & IDRg.Address(True, True) & ",0))/100"
    Next iter2 ' No check for overrunning IDRg, because sortIDRg should be trimmed to the right size
    
    ' Sort sortIDRg
    Call sortIDRg.Sort(key1:=sortIDRg, Order1:=xlAscending)
    
    ' Expand ID column widths to fit contents
    IDRg.EntireColumn.AutoFit
    sortIDRg.EntireColumn.AutoFit
    
    ' Insert line name above sorted summary & coerce to text format (just in case)
    With sortOrigRg.Offset(-1, 0).Resize(1, 1)
        .NumberFormat = "@"
        .Formula = lineRg.Cells(1, 1).Text
        .Font.Bold = True
    End With
    
    ' Pass the sorted ID range, the avg conc range, and the intRSD range to the summary
    '  creation subroutine
    Call insertDataSummary(sortIDRg, sortAdjRg, sortIntRSDRg)
    
End Sub

Public Sub insertDataSummary(sortIDRg As Range, sortConcRg As Range, sortRSDRg As Range)
    Dim summIDRg As Range, summConcRg As Range, summRSDRg As Range, summUncRg As Range
    Dim summVarRg As Range, parseRg As Range, cel As Range
    Dim rx As New RegExp, mtchs As MatchCollection
    Dim repRows As Range
    Dim iter1 As Long, iter2 As Long
    Dim st1 As String, st2 As String, formStr As String
    Dim localConcOffset As Long, localRSDOffset As Long
    
    ' Calculate the local offsets
    localConcOffset = sortAdjOffset - sortIDOffset
    localRSDOffset = sortIntRSDOffset - sortIDOffset
    
    ' Bind the top cells of the summary ranges
    Set summIDRg = sortIDRg.Offset(0, summIDOffset).Resize(1, 1)
    Set summConcRg = sortIDRg.Offset(0, summConcOffset).Resize(1, 1)
    Set summVarRg = sortIDRg.Offset(0, summVarOffset).Resize(1, 1)
    Set summRSDRg = sortIDRg.Offset(0, summRSDOffset).Resize(1, 1)
    Set summUncRg = sortIDRg.Offset(0, summUncOffset).Resize(1, 1)
    
    ' Add titles to the summary ranges
    summIDRg.Offset(-1, 0).Formula = "Sample"
    summConcRg.Offset(-1, 0).Formula = "Avg Conc"
    summVarRg.Offset(-1, 0).Formula = "Conc S^2"
    summRSDRg.Offset(-1, 0).Formula = "RSD"
    summUncRg.Offset(-1, 0).Formula = "Conc Unc"
    
    ' Construct RegExp to test for whether a sample name is a replicate or not
    With rx
        .Global = False
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "(.*[^ ])([ ]+R[0-9]+)$" ' End of sample name to be one or more spaces, 'R' or 'r', then one or more digits
                     ' Pattern set up so that sample name and replicate substring can be retrieved separately if needed
    End With
    
    ' Loop over all of the rows of sample data, searching for baseline sample names to process
    For iter1 = 1 To sortIDRg.Rows.Count
        st1 = sortIDRg.Cells(iter1, 1).Text  ' Store temp string
        If Not rx.Test(st1) Then
            ' Baseline sample name found!
            ' Add name to summary ID column
            summIDRg.Cells(summIDRg.Rows.Count, 1).Formula = st1
            
            ' Re-initalize the replicate range to just the baseline sample
            Set repRows = sortIDRg.Cells(iter1, 1)
            
            ' If not the last row of sortIDRg, begin search for associated replicates
            If Not iter1 = sortIDRg.Rows.Count Then
                ' Re-loop over the rows of sample data, searching for replicates
                For iter2 = 1 To sortIDRg.Rows.Count
                    st2 = sortIDRg.Cells(iter2, 1).Text
                    If rx.Test(st2) Then
                        ' This is a replicate; now test against the base string
                        Set mtchs = rx.Execute(st2)
                        
                        ' Only add as replicate if there's exactly one match
                        If mtchs.Count = 1 Then
                            ' And only add as replicate if the first substring of the one
                            '  match (zero-indexed!) is identical to the basename string
                            If mtchs(0).SubMatches(0) = st1 Then
                                Set repRows = Union(repRows, sortIDRg.Cells(iter2, 1))
                            End If
                        End If 'mtchs.Count = 1
                    End If ' rx.Test(st2)
                Next iter2 ' 1 to sortIDRg.Rows.Count
            End If ' Not iter1 = sortIDRg.Rows.Count
            
            ' All replicates found at this point; add the relevant formulas to the summary ranges
            ' === AVERAGE CONC ===
            summConcRg.Cells(summConcRg.Rows.Count, 1).Formula = _
                    "=AVERAGE(" & repRows.Offset(0, localConcOffset).Address(False, False) & ")"
            
            ' === CONC VARIANCE ===
            summVarRg.Cells(summVarRg.Rows.Count, 1).Formula = _
                    "=VAR.S(" & repRows.Offset(0, localConcOffset).Address(False, False) & ")"
            
            ' ==== RSS RSD ====
            summRSDRg.Cells(summRSDRg.Rows.Count, 1).Formula = _
                    "=SQRT(SUMSQ(" & repRows.Offset(0, localRSDOffset).Address(False, False) & "))"
            
            ' ==== UNCERTAINTY CALC ====
            With repRows.Areas
                ' Initialize formula string with SQRT and concentration variance value
                formStr = "=SQRT(" & summVarRg.Cells(summVarRg.Rows.Count, 1).Address(False, False) & "+"
                
                ' Add new terms for the uncertainty in each value
                For iter2 = 1 To .Count
                    formStr = formStr & "SUMSQ(" & .Item(1).Offset(0, localConcOffset).Address(False, False) & _
                                    "*" & .Item(1).Offset(0, localRSDOffset).Address(False, False) & ")"
                    If Not iter2 = .Count Then
                        ' Need to add a plus sign connecting the various terms
                        formStr = formStr & "+"
                    End If
                Next iter2 ' 1 To .Count
                formStr = formStr & ")"
            End With ' repRows.Areas
            
            ' Assign formula to appropriate cell
            summUncRg.Cells(summUncRg.Rows.Count, 1).FormulaArray = formStr
        
            ' Expand summary ranges downward to point to the next open row
            If Not iter1 = sortIDRg.Rows.Count Then
                Set summIDRg = summIDRg.Resize(summIDRg.Rows.Count + 1, 1)
                Set summConcRg = summConcRg.Resize(summConcRg.Rows.Count + 1, 1)
                Set summVarRg = summVarRg.Resize(summVarRg.Rows.Count + 1, 1)
                Set summRSDRg = summRSDRg.Resize(summRSDRg.Rows.Count + 1, 1)
                Set summUncRg = summUncRg.Resize(summUncRg.Rows.Count + 1, 1)
            End If
            
        End If  ' Not rx.Test(st)
            
    Next iter1 ' 1 To sortIDRg.Rows.Count
    
    ' Apply number formats to the summary range cells
    summConcRg.NumberFormat = "0"
    summRSDRg.NumberFormat = "0.0%"
    With Application.WorksheetFunction
        ' summUncRg
        Set parseRg = Nothing
        For Each cel In summUncRg.Cells
            ' Only examine cells that are not errors
            If Not .IsError(cel) Then
                If parseRg Is Nothing Then
                    Set parseRg = cel
                Else
                    Set parseRg = Union(parseRg, cel)
                End If
            End If
        Next cel
        
        ' Set format (if there's a non-error cell to base it on)
        If Not parseRg Is Nothing Then
            If parseRg.Cells.Count > 0 Then
                If .Average(parseRg) <= 200 Then
                    summUncRg.NumberFormat = "0.0"
                Else
                    summUncRg.NumberFormat = "0"
                End If
            End If
        End If
        
        ' summVarRg
        Set parseRg = Nothing
        For Each cel In summVarRg.Cells
            ' Only examine cells that are not errors
            If Not .IsError(cel) Then
                If parseRg Is Nothing Then
                    Set parseRg = cel
                Else
                    Set parseRg = Union(parseRg, cel)
                End If
            End If
        Next cel
        
        ' Set format (if there's a non-error cell to base it on)
        If Not parseRg Is Nothing Then
            If parseRg.Cells.Count > 0 Then
                If .Average(parseRg) <= 200 Then
                    summVarRg.NumberFormat = "0.0"
                Else
                    summVarRg.NumberFormat = "0"
                End If
            End If
        End If
    End With 'Application.WorksheetFunction
    
    ' Auto-size the summary range columns
    Union(summIDRg, summConcRg, summVarRg, summRSDRg, summUncRg).EntireColumn.AutoFit

End Sub

Public Function proofSplNames(nameRg As Range) As String
' Checks to be sure that sample naming is robust for searching purposes: no sample name
'  is a substring of any other sample name, save for an appended " R#" string indicating
'  a replicate.
' Function also checks for orphaned replicates (possibly a typo in a base name)
' Zero-length string indicates sample names are all unique-search-robust
' Nonzero-length string indicates sample names are NOT unique-search-robust;
'  function returns description of problem

    Dim rx As New RegExp, tgtCel As Range, srchCel As Range, tgtSt As String
    Dim rxOrph As New RegExp, baseFound  As Boolean
    Dim srchSt As String, remSt As String
    Dim val As Variant
    Dim wsf As WorksheetFunction
    Set wsf = Application.WorksheetFunction
    
    ' Initialize function return as a 'proof error' notice
    proofSplNames = "##PROOF ERROR##"
    
    ' Populate settings for the replicate testing RegExp
    With rx
        .Global = False
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "^[ ]+R[0-9]+$"  ' Remnant characters must be one or more spaces, an 'r' (case insensitive),
    End With                        '  and a series of one or more digits
    
    With rxOrph
        .Global = False
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "^(.*[^ ])([ ]+R[0-9]+)$" ' End of sample name to be one or more spaces, 'R' or 'r', then one or more digits
                     ' Pattern set up so that sample name and replicate substring can be retrieved separately if needed
    End With
    
    ' Iterate over each cell as the start-search (want to proof names from 'both directions')
    ' This block is to check uniqueness of naming
    For Each tgtCel In nameRg.Cells
        ' Store trimmed cell text
        tgtSt = wsf.Trim(tgtCel.Text)
        ' Iterate over all other cells, first checking for strict non-containment
        For Each srchCel In nameRg.Cells
            If tgtCel.Address <> srchCel.Address Then ' Not the same cell; check uniqueness
                srchSt = wsf.Trim(srchCel.Text) ' Store trimmed search string
                Select Case InStr(srchSt, tgtSt)
                Case 0: ' tgtCel not found in srchCel; no uniqueness problems; do nothing
                Case 1: ' tgtCel string found at start of srchCel; check for replicate name construct
                    remSt = Right(srchSt, Len(srchSt) - Len(tgtSt)) ' Retrieve remainder of search string
                    If Not rx.Test(remSt) Then
                        proofSplNames = """" & tgtSt & """ found in """ & srchSt & """; replicate syntax invalid"
                        Exit Function ' Exit function, reporting string found
                    End If
                Case Else:
                    If IsNull(InStr(srchSt, tgtSt)) Then
                        ' InStr or source/target string error
                        proofSplNames = "InStr() or string error"
                    Else
                        ' tgtCel string found elsewhere in srchCel or InStr error; return report
                        proofSplNames = """" & tgtSt & """ found in """ & srchSt & """"
                    End If
                    Exit Function
                End Select
            End If
        Next srchCel
    Next tgtCel
    
    ' This block is to check for orphaned replicates
    For Each tgtCel In nameRg.Cells
        ' Store trimmed cell text
        tgtSt = wsf.Trim(tgtCel.Text)
        
        ' Initialize the Boolean indicating whether the base name was found
        baseFound = False
        
        ' Check if it's a replicate; if not, it can't be an orphaned replicate; if it is, test it
        If rxOrph.Test(tgtSt) Then
            ' Retrieve the basename of the replicate -- FRAGILE if not a replicate, but
            '  fail case should never occur...
            srchSt = rxOrph.Execute(tgtSt).Item(0).SubMatches(0)
            
            ' Scan the trimmed strings of all of the rest of the names to confirm
            '  that the srchSt is present as a lone base name
            For Each srchCel In nameRg.Cells
                ' If the base name is found, set the Boolean to True and Exit For
                If wsf.Trim(srchCel.Text) = srchSt Then
                    baseFound = True
                    Exit For
                End If
            Next srchCel
            
            ' If baseFound remains False, complain of the orphaned replicate
            If Not baseFound Then
                proofSplNames = "Orphaned replicate found: """ & tgtSt & """"
                Exit Function
            End If   ' Not baseFound
        End If   ' rx.Test(tgtSt)
    Next tgtCel
    
    ' Made it all the way through - set a successful zero-length string return value
    proofSplNames = ""
    
End Function

Public Function doDataImport(destCell As Range, filePath As String) As QueryTable
' Import of data into the target cell from the chosen file
    
    Dim rx As New RegExp
    Dim qName As String
    Dim num As Long, so As String, desc As String, helpC As String, helpF As String
        
    ' Set up and execute regexp to find filename
    With rx
        .Global = False
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "\\([^\\]*)\."
        qName = .Execute(filePath).Item(0).SubMatches(0)
    End With
    
    ' Attach data file to QueryTable
    Set doDataImport = destCell.Worksheet.QueryTables.Add("TEXT;" & filePath, destCell)
    
    ' Configure the QueryTable
    With doDataImport
        .Name = qName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
    End With
    
    ' Broad error trap on QueryTable refresh
    On Error Resume Next
    doDataImport.Refresh BackgroundQuery:=False
    
    ' Collect info out of Err; may be absent if no error (Err.Number = 0)
    num = Err.Number
    so = Err.Source
    desc = Err.Description
    helpC = Err.HelpContext
    helpF = Err.HelpContext
    
    Err.Clear
    On Error GoTo 0
    
    Select Case num
    Case 0 ' Do nothing; successful QuertyTable refresh
    Case Else: Set doDataImport = Nothing ' File not found, or some other error; graceful exit
    End Select
    
End Function

Public Function GetTgtCell() As Range
' Retrieval of target cell
    
    Dim num As Long, so As String, desc As String, helpC As String, helpF As String
    
    ' Request target cell from user, with focused error trapping for Cancel of InputBox
    On Error Resume Next
    Set GetTgtCell = Application.InputBox(Prompt:="Please select target cell for data import.", _
                        Title:="Select target cell", Type:=8, Default:="=" & ActiveCell.Address(True, True))
    
    ' Collect info out of Err; may be absent if no error (Err.Number = 0)
    num = Err.Number
    so = Err.Source
    desc = Err.Description
    helpC = Err.HelpContext
    helpF = Err.HelpContext
    
    Err.Clear
    On Error GoTo 0
    
    Select Case num
    Case 0 ' Do nothing; successful range assignment
    Case 424: Set GetTgtCell = Nothing ' Probably, user cancelled
    Case Else: ' Re-raise prior error
        Call Err.Raise(num, so, desc, helpC, helpF)
    End Select
    
End Function

Public Function GetDataFile() As String
' Returns explicit full path string to selected data file, or empty string if canceled
    Dim fd As FileDialog
    Dim val As Long
    
    ' Attach file dialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Import"
        .InitialView = msoFileDialogViewList
        .Title = "Select data file to import"
        Call .Filters.Add("Text Files", "*.txt; *.csv")
        Call .Filters.Add("All Files", "*.*")
        val = .Show
        If val = 0 Then
            GetDataFile = ""
        Else
            GetDataFile = .SelectedItems(1)
        End If
    End With
    
End Function

' NO LONGER USED -- OBSOLETE FROM WHEN FORMULA INSERTION WAS NOT YET AUTOMATED
'Public Sub InsertAvgFormulaInRange()
'' Assumes properly configured worksheet with properly highlighted desired values
'' Extrapolation functions could go either way - will need stronger logic to implement decently
'
'    Dim WorkRange As Range, cel As Range
'
'        Set WorkRange = Application.Selection
'
'    For Each cel In WorkRange.Cells
'        If cel.Interior.Color = 14610923 Then
'            cel.Select
'            Call insertInterpFormula
'        End If
'    Next cel
'
'End Sub

Public Sub insertInterpFormula(WorkCell As Range, timeCol As Range)
Attribute insertInterpFormula.VB_ProcData.VB_Invoke_Func = " \n14"
' Assumes generally properly constructed worksheet;

'    Dim WorkCell As Range, timeCol As Range, valCol As Range
    Dim valCol As Range, PriorStdRow As Range, NextStdRow As Range, SplRow As Range
    
    ' Orient from active cell
'    Set WorkCell = ActiveCell
    Set SplRow = WorkCell.EntireRow
    Set valCol = WorkCell.EntireColumn
    
    ' Locate times column - just assume it's called 'Time', which should hold for all Plasma96 imported data
'    Set timeCol = WorkCell.EntireColumn
'    With Application.WorksheetFunction
'        Do Until UCase(.Trim(timeCol.Cells(1, 1).Text)) = "TIME"
'            Set timeCol = timeCol.Offset(0, -1).EntireColumn
'        Loop
'    End With
    
    ' Locate prior standard recheck
    Set PriorStdRow = WorkCell
    Do Until PriorStdRow.Interior.Color = repColor
        Set PriorStdRow = PriorStdRow.Offset(-1, 0)
        If Intersect(PriorStdRow.EntireRow, timeCol) Is Nothing Then
            Call MsgBox("Backward extrapolation not implemented for cell " & WorkCell.Address(False, False), _
                        vbOKOnly + vbInformation, "Extrapolation not implemented")
            Exit Sub
        End If
    Loop
    Set PriorStdRow = PriorStdRow.EntireRow
    
    ' Locate next standard recheck (will include trap to avoid infinite loop if there is no next)
    Set NextStdRow = WorkCell
    Do Until NextStdRow.Interior.Color = repColor
        Set NextStdRow = NextStdRow.Offset(1, 0)
        If Intersect(NextStdRow.EntireRow, timeCol) Is Nothing Then
            Call MsgBox("Forward extrapolation not implemented for cell " & WorkCell.Address(False, False), _
                    vbOKOnly + vbInformation, "Extrapolation not implemented")
            Exit Sub
        End If
    Loop
    Set NextStdRow = NextStdRow.EntireRow
    
    ' Construct formula
    WorkCell.Formula = "=(" & _
                                Intersect(timeCol, SplRow).Address(False, False) & _
                            "-" & _
                                Intersect(timeCol, PriorStdRow).Address(False, False) & _
                        ")/(" & _
                                Intersect(timeCol, NextStdRow).Address(False, False) & _
                            "-" & _
                                Intersect(timeCol, PriorStdRow).Address(False, False) & _
                        ")*(" & _
                                Intersect(valCol, NextStdRow).Address(False, False) & _
                            "-" & _
                                Intersect(valCol, PriorStdRow).Address(False, False) & _
                        ")+" & _
                            Intersect(valCol, PriorStdRow).Address(False, False)
    
    ' Should do it

End Sub
