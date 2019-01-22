Attribute VB_Name = "MyCode"
Option Explicit

Const FineZoom As Double = 5
Const CoarseZoom As Double = 25
Const MaxZoom As Double = 500
Const MinZoom As Double = 10
Const NoGroupErr As Long = -2147024891

Sub ZoomInFine()
    With ActiveWindow.View
        If .Zoom <= (MaxZoom - FineZoom) Then .Zoom = .Zoom + FineZoom
    End With
End Sub

Sub ZoomOutFine()
    With ActiveWindow.View
        If .Zoom >= (MinZoom + FineZoom) Then .Zoom = .Zoom - FineZoom
    End With
End Sub

Sub ZoomInCoarse()
    With ActiveWindow.View
        If .Zoom <= (MaxZoom - CoarseZoom) Then .Zoom = .Zoom + CoarseZoom
    End With
End Sub

Sub ZoomOutCoarse()
    With ActiveWindow.View
        If .Zoom >= (MinZoom + CoarseZoom) Then .Zoom = .Zoom - CoarseZoom
    End With
End Sub

Sub ZoomTo125()
    ActiveWindow.View.Zoom = 125
End Sub

Sub ImageFormat()
    
End Sub

Sub URLPull()
    Dim fs As FileSystemObject, fld As Folder, fl As File
    Dim wsh As WshShell, shct As WshShortcut
    Dim tStrm As TextStream, flTxt As String
    Dim rx As New RegExp
    Dim storeDoc As Document
    Dim fd As FileDialog
    Const storeName As String = "Links.docx"
    
    ' Set up regex
    With rx
        .Global = False
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "^URL=(.*)$"
    End With
    
    ' Bind filesystem, file dialog and wshell
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    Set wsh = CreateObject("WScript.Shell")
    
    ' Get the folder to process
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .Title = "Choose folder for URL collection"
        If InStr(ActiveDocument.Name, storeName) > 0 Then
            .InitialFileName = ActiveDocument.Path
        Else
            If InStr(UCase(.InitialFileName), "SYSTEM32") Or _
                        .InitialFileName = "" Then
                .InitialFileName = Environ("userprofile") & "\Documents"
            End If
        End If
        If .Show = 0 Then Exit Sub
        Set fld = fs.GetFolder(.SelectedItems(1))
    End With
    
    ' Check for an existing links doc; open if there, create
    '  new if not.
    ' Create a new doc and save in the working folder
    If fs.FileExists(fs.BuildPath(fld.Path, storeName)) Then
        Set storeDoc = Documents.Open(fs.BuildPath(fld.Path, storeName))
    Else
        Set storeDoc = Documents.Add
        storeDoc.SaveAs2 fs.BuildPath(fld.Path, storeName)
    End If
    
    ' Loop over all the files in the folder and check for .url
    For Each fl In fld.Files
        If UCase(Right(fl.Name, 3)) = "URL" Then
            ' Pull stream
            Set tStrm = fl.OpenAsTextStream(ForReading)
            flTxt = tStrm.ReadAll
            tStrm.Close
            
            ' Check for regex find
            If rx.Test(flTxt) Then
            
                ' Pull the URL and insert into storage doc
                ' Underline the filename for visibility
                storeDoc.Range.InsertAfter fl.Name & ": " & vbCrLf
                storeDoc.Range.InsertAfter _
                        rx.Execute(flTxt)(0).SubMatches(0) & vbCrLf
                storeDoc.Paragraphs(storeDoc.Paragraphs.Count - 3).Range _
                        .Font.Underline = wdUnderlineSingle
                ' Delete the URL file
                fl.Delete True
            Else
                ' Notify that URL info not found; do not delete
                storeDoc.Range.InsertAfter _
                        fl.Name & ": URL NOT FOUND IN FILE!" & vbCrLf
            End If
        ElseIf UCase(Right(fl.Name, 3)) = "LNK" Then
            ' Have to treat as a LNK object
            Set shct = wsh.CreateShortcut(fl.Path)
            
            ' Should always be there per the DOM
            storeDoc.Range.InsertAfter fl.Name & ": " & vbCrLf
            storeDoc.Range.InsertAfter _
                    shct.FullName & vbCrLf & vbCrLf
            storeDoc.Paragraphs(storeDoc.Paragraphs.Count - 3).Range _
                    .Font.Underline = wdUnderlineSingle
            
            ' Delete the LNK file
            Set shct = Nothing
            fl.Delete True
        End If
        
        ' Save the storage doc every loop -- conservative
        storeDoc.Save
        
        ' Underline the next to last paragraph
        'storeDoc.Paragraphs(storeDoc.Paragraphs.Count - 2).Range _
                .Font.Underline = wdUnderlineSingle
    Next fl
    
End Sub

Sub FieldInsertJustNumberFlag()

    Selection.Fields.ToggleShowCodes
    Selection.TypeText Text:="\t "
    Selection.Fields.Update

End Sub

Sub SuperscriptEndnoteNumbers()
    Dim en As Endnote, setting As Boolean
    Dim unit As WdUnits
    
    If ActiveDocument.Endnotes.Count < 1 Then Exit Sub
    
    With ActiveDocument.Endnotes
        setting = Not .Item(1).Range.Previous(wdCharacter, 1).Font.Superscript
        Select Case CStr(.Item(1).Range.Characters(1))
        Case ".", ")"
            unit = wdCharacter
        Case Else
            unit = wdWord
        End Select
        
        For Each en In .Parent.Endnotes
            en.Range.Previous(unit, 1).Font.Superscript = setting
        Next en
        
    End With
    
End Sub

Sub ScrubHyperlinks()
    Dim en As Endnote, fn As Footnote
    
'    For Each hl In ActiveDocument.Hyperlinks
'        doHLScrub hl
'    Next hl
'
'    For Each en In ActiveDocument.Endnotes
'        For Each hl In en.Range.Hyperlinks
'            doHLScrub hl
'        Next hl
'    Next en
'
'    For Each fn In ActiveDocument.Footnotes
'        For Each hl In fn.Range.Hyperlinks
'            doHLScrub hl
'        Next hl
'    Next fn
    
    With ActiveDocument.Hyperlinks
        Do While .Count > 0
            doHLScrub .Item(1)
        Loop
    End With
    
    For Each en In ActiveDocument.Endnotes
        With en.Range.Hyperlinks
            Do While .Count > 0
                doHLScrub .Item(1)
            Loop
        End With
    Next en
    
    For Each fn In ActiveDocument.Footnotes
        With fn.Range.Hyperlinks
            Do While .Count > 0
                doHLScrub .Item(1)
            Loop
        End With
    Next fn
    
End Sub

Sub doHLScrub(hl As Hyperlink)
    If hl.TextToDisplay = hl.Address Then
        hl.Delete
    Else
        hl.Range = hl.TextToDisplay & " [" & hl.Address & "] "
    End If
End Sub

Sub RotateHighlight()
'
' Rotates highlighting of selected text, including removal of the highlight.
'
' A side-effect of this construction: if the selection includes text whose
'  highlighting is not uniform, all highlighting is removed from the selection.
'

    Select Case Selection.Range.HighlightColorIndex
    Case wdYellow
        Selection.Range.HighlightColorIndex = wdBrightGreen
    Case wdBrightGreen
        Selection.Range.HighlightColorIndex = wdTurquoise
    Case wdTurquoise
        Selection.Range.HighlightColorIndex = wdRed
    Case wdNoHighlight
        Selection.Range.HighlightColorIndex = wdYellow
    Case wdRed
        Selection.Range.HighlightColorIndex = wdPink
    Case wdPink
        Selection.Range.HighlightColorIndex = wdGray25
    Case Else
        Selection.Range.HighlightColorIndex = wdNoHighlight
    End Select
    
End Sub


Sub OpenMLEqn()
Attribute OpenMLEqn.VB_Description = "Macro recorded 6/14/2010 by Brian Skinn"
Attribute OpenMLEqn.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.temp"
'
' Macro recorded 6/14/2010 by Brian Skinn, subsequently edited
'  Macro is for assignment to shortcut key for easy opening
'
    If Not Selection.Type = wdSelectionInlineShape Then Exit Sub
    Selection.InlineShapes(1).OLEFormat.DoVerb VerbIndex:=wdOLEVerbPrimary
    
End Sub


Public Sub VarItal()
    ' Function meant to be triggered by keyboard shortcut
    ' Function starts at the current cursor position and seeks backward one
    '  character at a time, italicizing all alphabetic or Greek characters.
    ' Function continues running backward until encountering the vertical pipe.
    
    Dim StopFlag As Boolean, PrevLoc As Range, ItalFlag As Boolean
    Dim PauseFlag As Boolean
    Dim mBoxResult As VbMsgBoxResult
    StopFlag = False
    PauseFlag = False
    Set PrevLoc = Selection.Range
    
    ' If the Selection falls within a table, or contains any tables,
    '  EXIT IMMEDIATELY!
    If Selection.Tables.Count > 0 Then
        Call MsgBox("VarItal cannot operate on tables", vbOKOnly, "Error - Table")
        Exit Sub
    End If
    
    ' Set the current selection to the end of any highlighting
    ' Detect whether italic will have to be de-set at the end
    If Selection.Start = Selection.End Then
        ItalFlag = True
        ' No collapse necessary
    Else
        Call Selection.Collapse(wdCollapseEnd)
        ItalFlag = False
    End If
    
    
    ' Check if vertical pipe is present in last 100 characters
    Call Selection.MoveLeft(wdCharacter, 100, True)
    If InStr(Selection.Text, "|") = 0 Then
        mBoxResult = MsgBox("Pipe character not found in previous 100 characters" & _
                Chr(10) & Chr(10) & _
                "Continue anyways?", _
                vbYesNo, "Pipe Not Found!")
        If mBoxResult <> vbYes Then
            Call Selection.Collapse(wdCollapseEnd)
            Exit Sub
        End If
    End If
    Call Selection.Collapse(wdCollapseEnd)
    
    Do While Not StopFlag
        ' Select the immediately previous character
        Call Selection.MoveLeft(wdCharacter, 1, True)
        ' If it should be italicized, set italics
        If Not PauseFlag Then
            ' Pause flag set means not to process content
            If isCharToItal(Selection) Then Selection.Font.Italic = True
        End If
        ' If it's the vertical pipe, set the stop flag and delete the pipe
        If AscW(Selection) = 124 Then
            StopFlag = True
            Call Selection.Delete(wdCharacter, 1)
        End If
        ' If it's an exclamation point, toggle the pause flag and delete the !
        If AscW(Selection) = 33 Then
            PauseFlag = Not PauseFlag
            Call Selection.Delete(wdCharacter, 1)
            ' Must bump insertion point one char to the right
            Call Selection.MoveRight(wdCharacter, 1, False)
        End If
        ' Set the insertion point to the left of the just-processed character
        Call Selection.MoveLeft(wdCharacter, 1, False)
    Loop
    
    ' Restore the pre-call selection
    PrevLoc.Select
    
    ' If deactivating italics mode is needed, do so
    If ItalFlag Then Selection.Range.Italic = False
    
End Sub

Public Sub VarBold()
    ' Function meant to be triggered by keyboard shortcut
    ' Function starts at the current cursor position and seeks backward one
    '  character at a time, bolding all alphabetic or Greek characters.
    ' Function continues running backward until encountering the vertical pipe.
    
    Dim StopFlag As Boolean, PrevLoc As Range, BoldFlag As Boolean
    Dim PauseFlag As Boolean
    Dim mBoxResult As VbMsgBoxResult
    StopFlag = False
    PauseFlag = False
    Set PrevLoc = Selection.Range
    
    ' If the Selection falls within a table, or contains any tables,
    '  EXIT IMMEDIATELY!
    If Selection.Tables.Count > 0 Then
        Call MsgBox("VarBold cannot operate on tables", vbOKOnly, "Error - Table")
        Exit Sub
    End If
    
    ' Set the current selection to the end of any highlighting
    ' Detect whether italic will have to be de-set at the end
    If Selection.Start = Selection.End Then
        BoldFlag = True
        ' No collapse necessary
    Else
        Call Selection.Collapse(wdCollapseEnd)
        BoldFlag = False
    End If
    
    
    ' Check if vertical pipe is present in last 100 characters
    Call Selection.MoveLeft(wdCharacter, 100, True)
    If InStr(Selection.Text, "|") = 0 Then
        mBoxResult = MsgBox("Pipe character not found in previous 100 characters" & _
                Chr(10) & Chr(10) & _
                "Continue anyways?", _
                vbYesNo, "Pipe Not Found!")
        If mBoxResult <> vbYes Then
            Call Selection.Collapse(wdCollapseEnd)
            Exit Sub
        End If
    End If
    Call Selection.Collapse(wdCollapseEnd)
    
    Do While Not StopFlag
        ' Select the immediately previous character
        Call Selection.MoveLeft(wdCharacter, 1, True)
        ' If it should be italicized, set bold
        If Not PauseFlag Then
            ' Pause flag set means not to process content
            If isCharToItal(Selection) Then Selection.Font.Bold = True
        End If
        ' If it's the vertical pipe, set the stop flag and delete the pipe
        If AscW(Selection) = 124 Then
            StopFlag = True
            Call Selection.Delete(wdCharacter, 1)
        End If
        ' If it's an exclamation point, toggle the pause flag and delete the !
        If AscW(Selection) = 33 Then
            PauseFlag = Not PauseFlag
            Call Selection.Delete(wdCharacter, 1)
            ' Must bump insertion point one char to the right
            Call Selection.MoveRight(wdCharacter, 1, False)
        End If
        ' Set the insertion point to the left of the just-processed character
        Call Selection.MoveLeft(wdCharacter, 1, False)
    Loop
    
    ' Restore the pre-call selection
    PrevLoc.Select
    
    ' If deactivating italics mode is needed, do so
    If BoldFlag Then Selection.Range.Bold = False
    
End Sub

Public Sub PullTitles()
    Dim pg As Paragraph
    
    For Each pg In ActiveDocument.Paragraphs
        pg.Range.Select
        'If Not InStr(pg.Range.Text, "TITLE:") > 0 Then
        'If Not InStr(pg.Range.Text, "13A-T") > 0 Then
        If pg.Range.Characters.Count < 5 Then
            pg.Range.Delete
        Else
            If Not (pg.Range.Characters(1).Underline <> wdUnderlineNone Or _
                    pg.Range.Characters(pg.Range.Characters.Count - 4).Underline <> wdUnderlineNone) Then
                pg.Range.Delete
            End If
        End If
    Next pg
    
End Sub

Public Sub PullTableTitles()
    
    Dim srcD As Document, tgtD As Document, t As Table
    
    Set srcD = ActiveDocument
    Set tgtD = Application.Documents.Add
    
    Application.ScreenUpdating = False
    For Each t In srcD.Tables
        If InStr(UCase(t.Range), "TPOC-") < 1 Then
            'srcD.Activate
            t.Select
            Selection.Copy
            
            'tgtD.Activate
            
            tgtD.Characters.Last.Select
            Selection.Collapse wdCollapseEnd
            Selection.Paste
        
'        If t.Rows.Count > 1 Then
'            tgtD.Characters.Last.Select
'            Selection.Collapse wdCollapseEnd
'            Selection.InsertAfter vbCrLf & "--------"
'        End If
        End If
        
    Next t
    
    Application.ScreenUpdating = True
    
End Sub

Public Sub DeleteTPOCTables()

    Dim t As Table, c As Cell
    
    For Each t In ActiveDocument.Tables
        t.Select
        If InStr(UCase(t.Range), "TPOC") Then
            t.Delete
        End If
    Next t
End Sub

Public Sub MergeRefFieldFormat()

    Dim sel As Range, workFld As Field, workChr As Range
    Dim rx As New RegExp
    
    ' Configure the regex
    With rx
        .Pattern = "[a-z0-9]"
        .Global = False
        .IgnoreCase = True
        .MultiLine = False
    End With
    
    ' Grab the search range
    Set sel = Selection.Range
    
    ' Iterate over all the fields
    For Each workFld In sel.Fields
        If InStr(UCase(workFld.Code), "REF") > 0 Then
            ' Only work with REF fields
            
            ' Add the MERGEFORMAT flag if it doesn't have it
            If Not InStr(UCase(workFld.Code), "MERGEFORMAT") > 0 Then
                workFld.Code.InsertAfter "\* MERGEFORMAT "
            End If
            
            ' Find the nearest previous letter/number character
            Set workChr = workFld.Result.Previous(wdCharacter, 2)
            Do Until rx.Test(workChr)
                Set workChr = workChr.Previous(wdCharacter, 1)
            Loop
            
            ' Bring the formatting across.  For now, just underline, bold, italic.
            ' Also font size
            workFld.Result.Font.Bold = workChr.Font.Bold
            workFld.Result.Font.Underline = workChr.Font.Underline
            workFld.Result.Font.Italic = workChr.Font.Italic
            workFld.Result.Font.Size = workChr.Font.Size
            
            ' Refresh the field code to ensure it's working
            workFld.Update
            
        End If
    Next workFld

End Sub

Public Function isCharToItal(c As String) As Boolean
    ' Indicates whether a character should be italicized
    Select Case AscW(c)
    '     UC Greek    LC Greek   UC Latin   LC Latin
    Case 913 To 937, 945 To 969, 65 To 90, 97 To 122
        isCharToItal = True
    Case Else
        isCharToItal = False
    End Select
End Function

Public Sub UpdateAllFields()
    Dim wkDoc As Document
    Dim fi As Field, shp As Shape
    
    ' Assign the document
    Set wkDoc = ActiveDocument

    ' Loop over all shapes
    For Each shp In wkDoc.Shapes
        ' Check if a shape is a group
        UpdateGroupFields shp
    Next shp
    
    ' Update all the base-level fields
    wkDoc.Fields.Update
    
End Sub

Public Sub UpdateGroupFields(inShp As Shape)
    Dim grp As GroupShapes, sh As Shape
    Dim errNum As Long, errDesc As String
    Dim iter As Long
    
    ' Check if group
    On Error Resume Next
    Set grp = inShp.GroupItems
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    
    ' Check status
    Select Case errNum
    Case 0
        ' Retrieved fine, process each subshape
        For iter = 1 To grp.Count
            Set sh = grp(iter)
            UpdateGroupFields sh
        Next iter
    Case NoGroupErr
        ' Not a group; process
        With inShp.TextFrame
            If .HasText Then
                .TextRange.Fields.Update
            End If
        End With
    Case Else
        ' Something else; reraise
        Err.Raise errNum, , errDesc
    End Select
    
    

End Sub

Sub PopupBookmarkName()
    
    If Selection.Bookmarks.Count > 0 Then
        InputBox Prompt:="Bookmark name:", _
                 Title:="Bookmark", _
                 Default:=Selection.Bookmarks(1).Name
    Else
        MsgBox "No bookmark here."
    End If
    
End Sub

Sub BookmarkNameToClipboard()

    If Selection.Bookmarks.Count > 0 Then
    
    Else
        MsgBox "No bookmark here."
    End If

End Sub

Sub ToggleBookmarkAnnotations()

    Dim bkm As Bookmark, cmt As Comment
    Dim tfm As TextFrame
    Const flag As String = "###BOOKMARK### "
    Dim check As Boolean
    
    ' First check all comments to see if any lead with the flag;
    check = False
    For Each cmt In ActiveDocument.Comments
        If TBA_FlagCheck(cmt, flag) Then
            check = True
            Exit For
        End If
    Next cmt
    
    ' If so, delete all comments that lead with the flag and exit
    If check Then
        For Each cmt In ActiveDocument.Comments
            If TBA_FlagCheck(cmt, flag) Then
                cmt.Delete
            End If
        Next cmt
        
        Exit Sub
    End If
    
    ' If not, add comments for all bookmarks.
    For Each bkm In ActiveDocument.Bookmarks
        Select Case bkm.StoryType
        Case wdMainTextStory
            ActiveDocument.Comments.Add Range:=bkm.Range, _
                                        Text:=flag & bkm.Name
        Case wdTextFrameStory
            ' Haven't figured this out yet.
        End Select
    Next bkm

End Sub

Sub ToggleShapeAnchorLock()
    doAnchorLockSwap Selection.ShapeRange
End Sub

Private Sub doAnchorLockSwap(obj As Object)
    On Error Resume Next
        obj.LockAnchor = Not obj.LockAnchor
    Err.Clear: On Error GoTo 0
End Sub

Private Function TBA_FlagCheck(cmt As Comment, flag As String) As Boolean
    TBA_FlagCheck = Left(cmt.Range.Text, Len(flag)) = flag
End Function

Sub CycleNumberingFormat()
    Dim oldRg As Range
    Dim wkStr As String, newNumType As String
    Dim newCodeStr As String
    Dim rx As New RegExp, mch As Match
    
    If Selection.Characters.Count > 1 And Selection.Fields.Count > 1 Then Exit Sub
    
    Set oldRg = Selection.Range
    
    Selection.Expand wdWord
    
    If Selection.Fields.Count <> 1 Then GoTo RestoreExit
    
    wkStr = Selection.Fields(1).Code
    
    With rx
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^(.+)\\[*]\s+(\w+)(.*)$"
        
        If Not .Test(wkStr) Then GoTo RestoreExit
        Set mch = .Execute(wkStr)(0)
    End With
    
    Select Case mch.SubMatches(1)
    Case "ARABIC", "Arabic", "arabic"
        newNumType = "ROMAN"
    Case "ROMAN"
        newNumType = "roman"
    Case "roman"
        newNumType = "ALPHABETIC"
    Case "ALPHABETIC"
        newNumType = "alphabetic"
    Case Else
        newNumType = "ARABIC"
    End Select
    
    newCodeStr = mch.SubMatches(0) & "\* " & _
                newNumType & mch.SubMatches(2)
    
    With Selection.Fields(1)
        .Code.Text = newCodeStr
        .Update
    End With
    
RestoreExit:
    oldRg.Select
    
End Sub

Sub GroupShapes()
    On Error Resume Next
        Selection.ShapeRange.Group
    Err.Clear: On Error GoTo 0
End Sub

Sub UnGroupShapes()
    On Error Resume Next
        Selection.ShapeRange.Ungroup
    Err.Clear: On Error GoTo 0
End Sub

Sub WrapSquare()
    On Error Resume Next
        Selection.ShapeRange.WrapFormat = wdWrapSquare
    Err.Clear: On Error GoTo 0
End Sub

Sub WrapTopBottom()
    On Error Resume Next
        Selection.ShapeRange.WrapFormat = wdWrapTopBottom
    Err.Clear: On Error GoTo 0
End Sub

Sub WrapInline()
    On Error Resume Next
        Selection.ShapeRange.WrapFormat = wdWrapInline
    Err.Clear: On Error GoTo 0
End Sub

Sub WrapToggle()
    On Error Resume Next
        With Selection.ShapeRange.WrapFormat
            Select Case .Type
            Case wdWrapSquare
                .Type = wdWrapTopBottom
            Case wdWrapTopBottom
                .Type = wdWrapInline
            Case Else
                .Type = wdWrapSquare
            End Select
        End With
    Err.Clear: On Error GoTo 0
End Sub

Sub VariableFormat()
    Dim sel As Range ', length As Long

    On Error Resume Next
        Set sel = Selection
        
        Selection.Expand wdWord
        Selection.Characters(1).Italic = True
        If Selection.Characters.Count > 1 Then
            Selection.MoveStart wdCharacter, 1
            Selection.Font.Subscript = msoTrue
        End If
        
        sel.Select
        
    Err.Clear: On Error GoTo 0
End Sub
