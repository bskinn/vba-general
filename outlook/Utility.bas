Attribute VB_Name = "Utility"
Option Explicit

Sub viewToInbox()
    Set ActiveExplorer.CurrentFolder = Session.GetDefaultFolder(olFolderInbox)
End Sub

Sub viewToSentMail()
    Set ActiveExplorer.CurrentFolder = Session.GetDefaultFolder(olFolderSentMail)
End Sub

Sub viewToToFile()
    Set ActiveExplorer.CurrentFolder = Session.GetFolderFromID("0000000078B8223B3673A24BA53C0F0317A6E87BA2A50000")
End Sub

Sub deleteUnread()
    Dim ob As Object, insp As Inspector, sel As Selection, cl As OlObjectClass
    Dim ht As Long, wd As Long, ns As NameSpace
    
    ' This registry key name is specific to Outlook 2010
    Const readReceiptRegKeyName As String = "HKCU\Software\Microsoft\Office\14.0\Outlook\Options\Mail\Receipt Response"
    Dim readReceiptRegSetting As Long, wshSh As WshShell
    
    ' NOTE: Despite its name, this method *WILL* trigger a read receipt if .ReadReceiptRequested = True
    '  on a received MailItem.  Read receipt distribution is triggered any time .UnRead is changed to False,
    '  which means that it is impossible to avoid distro of read receipt using standard object model.
    '
    '  Possible approach is to tweak Outlook registry entry for the 'read receipt behavior' setting
    
    ' Assign the Selection
    Set sel = ActiveExplorer.Selection
    
    ' If nothing selected, exit
    If sel.Count = 0 Then Exit Sub
    
'    ' Subject for utility message to clear mail received icon
'    st = "Icon Clearer"
    
'    ' Bind the NameSpace and create a new mail item
'    Set ns = Application.Session
'    Set mi = Application.CreateItem(olMailItem)
    
'    ' Assign the subject, save, and move the item to the Inbox
'    With mi
'        .Subject = st
'        .Save
'        Call .Move(ns.GetDefaultFolder(olFolderInbox))
'    End With
'    Set mi = Nothing
    
    ' Bind the WScript Shell
    Set wshSh = CreateObject("WScript.Shell")
    
    ' Store the current state of the Outlook read receipt behavior registry value, then
    '  set to 'never send receipts'
    readReceiptRegSetting = wshSh.RegRead(readReceiptRegKeyName)
    Call wshSh.RegWrite(readReceiptRegKeyName, &H1, "REG_DWORD")
    
    ' Generate an Inspector for the utility object and immediately close, to clear the 'mail received' icon
    Set insp = sel.Item(1).GetInspector
    With insp
        ' Store dimensions
        ht = insp.Height
        wd = insp.Width
        ' Zero dimensions so window does not appear
        .Height = 0
        .Width = 0
        ' "Show" the message to trigger removal of the 'new mail' icon
        .Display
        ' Bring the explorer window to the front
        ActiveExplorer.Activate
        ' Restore the dimensions of the Inspector to avoid munging future-opened Inspectors
        .Height = ht
        .Width = wd
        ' Close the Inspector
        Call .Close(olDiscard)
    End With
'
'    ' Mark utility item as read, and permanently delete
'    With ns.GetDefaultFolder(olFolderInbox).Items.Find("[Subject] = " & st)
'        .UnRead = False
'        .Delete
'    End With
'    ns.GetDefaultFolder(olFolderDeletedItems).Items.Find("[Subject] = " & st).Delete
    
    For Each ob In sel
        ' Mark as read (if can be marked as read)
        cl = ob.Class
        If cl = olAppointment Or cl = olContact Or cl = olMail Or cl = olMeetingCancellation Or _
                cl = olMeetingForwardNotification Or cl = olMeetingRequest Or _
                cl = olMeetingResponseNegative Or cl = olMeetingResponsePositive Or _
                cl = olMeetingResponseTentative Or cl = olPost Or cl = olReport Or cl = olTask Or _
                cl = olTaskRequest Or cl = olTaskRequestAccept Or cl = olTaskRequestDecline Or _
                cl = olTaskRequestUpdate Then
            ob.UnRead = False
        End If
        
        ' Delete object regardless of its type
        ob.delete
    Next ob
    
    ' Restore the prior-set read receipt behavior and dereference the WScript shell
    Call wshSh.RegWrite(readReceiptRegKeyName, readReceiptRegSetting, "REG_DWORD")
    Set wshSh = Nothing
    
End Sub

Public Sub tagTextIntoEmail2(itm As Object, atchName As String, saveName As String)
    ' Simple function thus far; will be good to add:
    '   Custom # of lines inserted, more general text addition, flexibility to font colors
    Dim timeRef As Double
    
    ' Re-display the item to establish a consistent state
    With itm
        Call .Close(olSave)
        .Display
    End With
    
    ' Activate editing mode
    Call SendKeys("%(H)AE", True)
    ' Insert two leadlines
    Call SendKeys("{ENTER 2}{UP 2}", True)
    ' Insert notice text
    Call SendKeys("Attachment '" & atchName & "' detached to " & saveName, True)
    ' Select notice text
    Call SendKeys("^+{UP}", True)
    ' Change font to red
    Call SendKeys("%", True)
    timeRef = Timer: Do While Timer <= timeRef + 0.1: DoEvents: Loop
    Call SendKeys("OFC{DOWN 7}{LEFT 4}{ENTER}", True)
    
End Sub

Public Sub highlightStrike()
    ' Takes item in active inspector and applies highlighting to any struck text
    Dim it As Object, wd As Word.Document, rg As Word.Range
    Dim cl As OlObjectClass
    
    ' Don't want to continue if no item is open!
    If ActiveInspector Is Nothing Then Exit Sub
    
    ' Believe that no Inspector should have no CurrentItem, but check in case
    If ActiveInspector.CurrentItem Is Nothing Then Exit Sub
    
    ' Bind the item
    Set it = ActiveInspector.CurrentItem
    
    ' Activate edit mode on the item if it's possibly needed
    cl = it.Class
    If cl = olMail Or cl = olMeetingCancellation Or cl = olMeetingRequest Or _
                    cl = olMeetingResponseNegative Or cl = olMeetingResponsePositive Or _
                    cl = olMeetingResponseTentative Or cl = olTaskRequest Or cl = olTaskRequestAccept Or _
                    cl = olTaskRequestDecline Or cl = olTaskRequestUpdate Then
        Call it.GetInspector.CommandBars.ExecuteMso("EditMessage")
    End If
    
    ' Bind the WordEditor  FAILS -- TaskItems apparently do not expose a WordEditor
    Set wd = it.WordEditor
    
    ' Iterate
    For Each rg In wd.Characters
        If rg.Font.Strikethrough Or rg.Font.DoubleStrikeThrough Then
            rg.HighlightColorIndex = wdYellow
        End If
    Next rg
    
End Sub

Sub dblStrike()
    With ActiveInspector.WordEditor.Application.Selection.Range.Font
        .DoubleStrikeThrough = Not .DoubleStrikeThrough
    End With
End Sub

Sub dupeItems()
    
    Dim it As Object

    For Each it In ActiveExplorer.Selection
        it.Copy
    Next it
End Sub

Sub deleteScrubAtchs()

    Dim mi As MailItem, resp As VbMsgBoxResult
    
    If Not TypeOf ActiveExplorer.Selection.Item(1) Is MailItem Then Exit Sub
    
    If MsgBox("Delete selected with attachment scrubbing?", vbOKCancel, _
                        "Confirm Scrubbed Delete") = vbCancel Then Exit Sub
    
    For Each mi In ActiveExplorer.Selection
        Do While mi.Attachments.Count > 0
            mi.Attachments.Item(1).delete
        Loop
        
        mi.Save
        mi.UnRead = False
        mi.delete
    Next mi
    

End Sub

Sub addAtchToAllSelected()

    Dim mi As MailItem
    Dim fd As FileDialog
    Dim fName As String
    Dim doSend As VbMsgBoxResult
    
    Set fd = CreateObject("Word.Application").FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Attach"
        .Filters.Add "All Files", "*.*"
        .Title = "Choose File to Attach"
        If .Show = 0 Then
            Exit Sub
        End If
        fName = .SelectedItems(1)
    End With
    
    doSend = MsgBox("Queue items to send after attaching?", _
                    vbYesNoCancel, "Queue items for send?")
    
    If doSend = vbCancel Then
        Exit Sub
    End If
    
    For Each mi In ActiveExplorer.Selection
        mi.Attachments.Add fName
        mi.Save
        If doSend = vbYes Then
            mi.Send
        End If
    Next mi

End Sub

Sub calExport()

    Dim ol As Outlook.Application
    Dim cal As Folder
    Dim exporter As CalendarSharing
    Dim mi As MailItem
    Dim insp As Inspector
    Dim shl As New Shell
    
    Set ol = Application
    Set cal = ol.Session.GetDefaultFolder(olFolderCalendar)
    Set exporter = cal.GetCalendarExporter
    
    With exporter
        .CalendarDetail = olFullDetails
        .IncludeAttachments = False
        .IncludePrivateDetails = False
        .RestrictToWorkingHours = False
        .IncludeWholeCalendar = True
        .SaveAsICal "C:\Users\Brian\Documents\btscal.ics"
'        Set mi = .ForwardAsICal(olCalendarMailFormatEventList)
    End With
    
    shl.ShellExecute "C:\Users\Brian\usrbin\calput.bat"
    
'    Set insp = ol.Inspectors.Add(mi)
'    insp.Activate
'

End Sub
