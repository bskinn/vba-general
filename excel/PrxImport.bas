Attribute VB_Name = "PrxImport"
Option Explicit

Sub PullPrxData()
Attribute PullPrxData.VB_ProcData.VB_Invoke_Func = "U\n14"

    Dim fd As FileDialog, fs As New FileSystemObject
    Dim fldPath As String, savePath As String
    Dim fld As Folder, fl As File
    Dim tStrm As TextStream
    Dim wbNew As Workbook, wkSht As Worksheet
    Dim dataStr As String
    Dim rxDate As New RegExp, rxValues As New RegExp
    Dim mColl As MatchCollection
    Dim mchDate As Match, mchValues As Match
    Dim wkCel As Range
    Dim iter As Long

    ' Regex for the particular date format in the export files
    ' Date: Tue Jan 09 12:01:16 EST 2018
    ' In order for Excel to parse it as a datetime, the day of the week
    ' and the time zone need to be stripped
    With rxDate
        .MultiLine = True
        .Global = False
        .IgnoreCase = True
        '                  DOW    mmm  dd    hh:mm:ss     TZ   yyyy
        .Pattern = "^Date: +\w+ +(\w+ +\d+ +\d+:\d+:\d+) +\w+ +(\d+).*$"
    End With
    
    ' Regex for the target 240nm wavelength. Finds the first line starting
    ' with "24" that has a line starting with "23" preceding it. The wavelength
    ' and absorbance are tab-separated (of course), and there's some oddball
    ' character (the \r of the \r\n Win newline?) before EOL that has to be
    ' matched
    With rxValues
        .MultiLine = False
        .Global = False
        .IgnoreCase = True
        '                                              WL value        Abs value
        .Pattern = "\n23\d[.]\d+[ \t]+[0-9.-]+[^\n]*\n(24\d\.\d+)[ \t]+([0-9.-]+)"
    End With

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Select"
        ' Only reset the initial view location if it's the default, in SYS32
        If InStr(UCase(.InitialFileName), "SYSTEM32") > 0 Then
            .InitialFileName = "%USERPROFILE%\Documents"
        End If
        .Title = "Select Folder for GNASAIIH2O2 UV-Vis Parsing"
        If .Show Then
            fldPath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Set fld = fs.GetFolder(fldPath)
    
    Set wbNew = Workbooks.Add
    
    Do Until wbNew.Worksheets.Count < 2
        wbNew.Sheets(1).Delete
    Loop
    
    Set wkSht = wbNew.Sheets(1)
    wkSht.Name = fs.GetFileName(fldPath)
    
    Set wkCel = wkSht.Range("A1")
    wkCel.Value = "Filename"
    wkCel.Offset(0, 1).Value = "Date/Time"
    wkCel.Offset(0, 2).Value = "Wavelength"
    wkCel.Offset(0, 3).Value = "Absorbance"
    
    Set wkCel = wkCel.Offset(1, 0)
    
    For Each fl In fld.Files
        If Right(fl.Name, 4) = ".txt" Then
            ' Pull data file
            Set tStrm = fl.OpenAsTextStream
            dataStr = tStrm.ReadAll
            tStrm.Close
            
            ' Extract data. Will blow up if either regex doesn't match.
            Set mchDate = rxDate.Execute(dataStr).Item(0)
            Set mchValues = rxValues.Execute(dataStr).Item(0)
            
            wkCel.Value = fl.Name
            wkCel.Offset(0, 1).Formula = "=VALUE(""" & _
                    mchDate.SubMatches(0) & " " & mchDate.SubMatches(1) & _
                    """)"
            wkCel.Offset(0, 2).Value = mchValues.SubMatches(0)
            wkCel.Offset(0, 3).Value = mchValues.SubMatches(1)
        End If
        
        Set wkCel = wkCel.Offset(1, 0)
        
    Next fl
    
    wkCel.Resize(1, 4).EntireColumn.AutoFit
    
    wkCel.EntireColumn.NumberFormat = "@"
    wkCel.Offset(0, 1).EntireColumn.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wkCel.Offset(0, 2).EntireColumn.NumberFormat = "0.000"
    wkCel.Offset(0, 3).EntireColumn.NumberFormat = "0.00"
    
    wkSht.UsedRange.Sort Key1:=wkSht.Range("B:B"), Header:=xlYes
    
    savePath = fs.BuildPath(fld.path, wkSht.Name & ".xlsx")
    
    If fs.FileExists(savePath) Then
        iter = 1
        Do Until Not fs.FileExists(dupePath(savePath, iter, fs))
            iter = iter + 1
        Loop
    End If
    
    If iter > 0 Then
        wbNew.SaveAs dupePath(savePath, iter, fs)
    Else
        wbNew.SaveAs savePath
    End If

End Sub

Function dupePath(path As String, num As Long, fs As FileSystemObject) As String
    dupePath = fs.GetParentFolderName(path) & "\" & fs.GetBaseName(path) & " (" & CStr(num) & ").xlsx"
End Function
