Attribute VB_Name = "CategoryHelper"
Option Explicit

Sub ToggleCategory(it As Object, cat As String)

    Dim cats As String
    Dim loc As Long
    
    cats = it.Categories
    loc = InStr(cats, cat)

    If loc > 0 Then
        ' Remove
        If cats = cat Then
            ' Was the only category
            it.Categories = ""
        ElseIf Right(cats, Len(cat)) = cat Then
            ' Was at end of the cat list
            it.Categories = Left(cats, loc - 3)
        Else
            ' Was elsewhere in the cats list of >1 items
            it.Categories = Left(cats, loc - 1) & _
                    Mid(cats, loc + Len(cat))
        End If
    Else
        ' Add
        If Len(cats) = 0 Then
            it.Categories = cat
        Else
            it.Categories = cats & ", " & cat
        End If
    End If
    
End Sub

Public Sub doToggle(cat As String)
    Dim it As Object, insp As Inspector
    Dim toggleTo As VbTriState
    Dim loc As Long
    
    toggleTo = vbUseDefault
    
    If TypeOf ActiveWindow Is Inspector Then
        ToggleCategory ActiveInspector.CurrentItem, cat
        ActiveInspector.CurrentItem.Save
    Else
        For Each it In ActiveExplorer.Selection
            ' Helper storing if/where the category is in the list
            loc = InStr(it.Categories, cat)
            
            If toggleTo = vbUseDefault Then
                ' Set which way to toggle
                If loc > 0 Then
                    toggleTo = vbFalse
                Else
                    toggleTo = vbTrue
                End If
            End If
            
            If (loc > 0 And toggleTo = vbFalse) Or _
                        (loc < 1 And toggleTo = vbTrue) Then
                Set insp = Application.Inspectors.Add(it)
                ToggleCategory insp.CurrentItem, cat
                insp.Close olSave
            End If
        Next it
    End If
End Sub

Sub ToggleSocieties()
    doToggle "Societies"
End Sub

Sub TogglePSI()
    doToggle "PSI"
End Sub

Sub ToggleContractors()
    doToggle "Contractors"
End Sub

Sub ToggleBnP()
    doToggle "B&P"
End Sub

Sub ToggleVendors()
    doToggle "Vendors"
End Sub

Sub ToggleFaraday()
    doToggle "Faraday"
End Sub

Sub ToggleCollabs()
    doToggle "Collaborators"
End Sub

Sub ToggleClients()
    doToggle "Clients"
End Sub

Sub ToggleFunders()
    doToggle "Funding Agencies"
End Sub
