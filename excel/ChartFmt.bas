Attribute VB_Name = "ChartFmt"
Option Explicit

Sub FormatInlineChart()
Attribute FormatInlineChart.VB_ProcData.VB_Invoke_Func = "Z\n14"
    
    Dim ca As ChartArea, co As ChartObject, pa As PlotArea, cht As Chart
    Dim shp As Shape, ax As Axis
    
    Set cht = ActiveChart
    If cht Is Nothing Then
        MsgBox "No chart selected!"
        Exit Sub
    End If
    
    Set ca = cht.ChartArea
    Set pa = cht.PlotArea
    Set co = cht.Parent
    Set shp = co.ShapeRange(1)
    
    With ca
        .Format.Line.Visible = msoFalse
    End With
    
    With co
        .Width = 300
        .Height = 250
    End With
    
    With cht
        On Error Resume Next
            .Legend.Delete
        Err.Clear: On Error GoTo 0
        
        .HasTitle = False
        
        Set ax = .Axes(xlCategory)
        With ax
            If Not .HasTitle Then .HasTitle = True
            If .AxisTitle.Text = "Axis Title" Then .AxisTitle.Text = "[X-Axis Text]"
            .AxisTitle.Characters.Font.Size = 14
            .TickLabels.Font.Size = 12
            .TickLabelPosition = xlTickLabelPositionLow
        End With
        
        Set ax = .Axes(xlValue)
        With ax
            If Not .HasTitle Then .HasTitle = True
            If .AxisTitle.Text = "Axis Title" Then .AxisTitle.Text = "[Y-Axis Text]"
            .AxisTitle.Orientation = xlUpward
            .AxisTitle.Characters.Font.Size = 14
            .TickLabels.Font.Size = 12
            .TickLabelPosition = xlTickLabelPositionLow
            With .MajorGridlines.Format.Line
                .Weight = 0.75
                .ForeColor.RGB = RGB(210, 210, 210)
                .DashStyle = msoLineDash
            End With
        End With
    End With
    
    With pa
        With .Format.Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(127, 127, 127)
        End With
    End With
    
End Sub
