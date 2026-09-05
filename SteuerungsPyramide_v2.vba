Sub CreateSteuerungsPyramide()
    Dim ws As Worksheet
    Dim wsExists As Boolean
    Dim i As Integer
    
    wsExists = False
    
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "Führungsinstrumente" Then
            Set ws = ThisWorkbook.Sheets(i)
            wsExists = True
            Exit For
        End If
    Next i
    
    If Not wsExists Then
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Führungsinstrumente"
    End If
    
    ws.Cells.Clear
    
    With ws
        .PageSetup.Orientation = xlLandscape
        .PageSetup.PaperSize = xlPaperA4
    End With
    
    Dim pyramidTop As Integer, pyramidLeft As Integer
    Dim pyramidWidth As Double
    
    pyramidTop = 100
    pyramidLeft = 150
    pyramidWidth = 400
    
    Dim colorTop As Long, colorMiddle As Long, colorBottom As Long
    Dim colorText As Long, colorBorder As Long, colorArrow As Long
    
    colorTop = RGB(41, 84, 209)
    colorMiddle = RGB(68, 114, 196)
    colorBottom = RGB(112, 173, 71)
    colorText = RGB(255, 255, 255)
    colorBorder = RGB(31, 31, 31)
    colorArrow = RGB(41, 84, 209)
    
    Call CreatePyramidLayer(ws, pyramidLeft, pyramidTop, 0, pyramidWidth, 90, colorTop, colorText, colorBorder, "MONATLICH" & vbNewLine & "Monthly Planning" & vbNewLine & "Strategic Direction" & vbNewLine & "Budget Review" & vbNewLine & "Performance Analysis")
    
    Call CreatePyramidLayer(ws, pyramidLeft + 50, pyramidTop + 90, 1, pyramidWidth - 100, 90, colorMiddle, colorText, colorBorder, "WÖCHENTLICH" & vbNewLine & "Weekly Meetings" & vbNewLine & "Progress Tracking" & vbNewLine & "Resource Allocation" & vbNewLine & "Risk Assessment")
    
    Call CreatePyramidLayer(ws, pyramidLeft + 100, pyramidTop + 180, 2, pyramidWidth - 200, 100, colorBottom, colorText, colorBorder, "TÄGLICH" & vbNewLine & "Daily Operations" & vbNewLine & "Task Execution" & vbNewLine & "Problem Resolution" & vbNewLine & "Quality Assurance" & vbNewLine & "Team Coordination")
    
    Call AddSeparatingLine(ws, pyramidLeft, pyramidTop + 90, pyramidWidth, colorBorder)
    Call AddSeparatingLine(ws, pyramidLeft + 50, pyramidTop + 180, pyramidWidth - 100, colorBorder)
    
    Call AddResponsibilitySection(ws, pyramidLeft + pyramidWidth + 100, pyramidTop, colorArrow)
    
    Call AddTitle(ws, pyramidLeft + pyramidWidth / 2 - 80, pyramidTop - 60, "Steuerungs- und Kommunikationspyramide", colorTop)
    
    Call AddLegend(ws, pyramidLeft, pyramidTop + 320, colorTop, colorMiddle, colorBottom)
    
    ActiveWindow.Zoom = 120
    
    MsgBox "Steuerungs- und Kommunikationspyramide erfolgreich erstellt!", vbInformation, "Erfolg"
    
End Sub

Sub CreatePyramidLayer(ws As Worksheet, leftPos As Integer, topPos As Integer, layerIndex As Integer, width As Double, height As Double, fillColor As Long, textColor As Long, borderColor As Long, pyramidText As String)
    
    Dim shape As Shape
    
    Set shape = ws.Shapes.AddShape(msoShapeTrapezoid, leftPos, topPos, width, height)
    
    With shape
        .Fill.ForeColor.RGB = fillColor
        .Fill.Transparency = 0
        .Line.Color.RGB = borderColor
        .Line.Weight = 2.5
    End With
    
    With shape.TextFrame
        .Clear
        .Word Wrap = True
        .MarginBottom = 8
        .MarginTop = 8
        .MarginLeft = 10
        .MarginRight = 10
        .VerticalAlignment = msoAnchorMiddle
        .HorizontalAlignment = msoAnchorCenter
        
        Dim textRange As TextRange
        Set textRange = .Characters
        textRange.Text = pyramidText
        
        With textRange.Font
            .Name = "Segoe UI"
            .Size = IIf(layerIndex = 0, 12, IIf(layerIndex = 1, 11, 10))
            .Color.RGB = textColor
            .Bold = True
        End With
    End With
    
    If layerIndex = 0 Then
        shape.Adjustments(1) = 0.15
    ElseIf layerIndex = 1 Then
        shape.Adjustments(1) = 0.25
    Else
        shape.Adjustments(1) = 0.35
    End If
    
End Sub

Sub AddSeparatingLine(ws As Worksheet, leftPos As Double, topPos As Double, lineWidth As Double, lineColor As Long)
    
    Dim line As Shape
    Set line = ws.Shapes.AddConnector(msoConnectorStraight, leftPos, topPos, leftPos + lineWidth, topPos)
    
    With line.Line
        .Color.RGB = lineColor
        .Weight = 2
        .DashStyle = msoLineSolid
    End With
    
End Sub

Sub AddResponsibilitySection(ws As Worksheet, leftPos As Double, topPos As Double, arrowColor As Long)
    
    Dim titleBox As Shape
    Set titleBox = ws.Shapes.AddTextbox(msoAnchorNone, leftPos, topPos - 40, 150, 30)
    With titleBox.TextFrame
        .Clear
        .Word Wrap = True
        .VerticalAlignment = msoAnchorMiddle
        With .Characters.Font
            .Name = "Segoe UI"
            .Size = 11
            .Bold = True
            .Color.RGB = RGB(41, 84, 209)
        End With
        .Characters.Text = "Verantwortung"
    End With
    
    Call AddResponsibilityRow(ws, leftPos, topPos + 20, "Executive" & vbNewLine & "Management", arrowColor)
    Call AddResponsibilityRow(ws, leftPos, topPos + 100, "Team Leaders" & vbNewLine & "Supervisors", arrowColor)
    Call AddResponsibilityRow(ws, leftPos, topPos + 190, "Team Members" & vbNewLine & "Employees", arrowColor)
    
End Sub

Sub AddResponsibilityRow(ws As Worksheet, leftPos As Double, topPos As Double, roleText As String, arrowColor As Long)
    
    Dim shape As Shape
    Dim textBox As Shape
    
    Set shape = ws.Shapes.AddShape(msoShapeRightArrow, leftPos, topPos, 30, 25)
    With shape
        .Fill.ForeColor.RGB = arrowColor
        .Fill.Transparency = 0.2
        .Line.Color.RGB = arrowColor
        .Line.Weight = 1.5
    End With
    
    Set shape = ws.Shapes.AddShape(msoShapeOval, leftPos + 35, topPos, 30, 30)
    With shape
        .Fill.ForeColor.RGB = RGB(200, 200, 200)
        .Fill.Transparency = 0.3
        .Line.Color.RGB = arrowColor
        .Line.Weight = 2
    End With
    
    Set shape = ws.Shapes.AddShape(msoShapeOval, leftPos + 41, topPos + 2, 18, 10)
    With shape
        .Fill.ForeColor.RGB = arrowColor
        .Fill.Transparency = 0
    End With
    
    Set shape = ws.Shapes.AddShape(msoShapeLeftArrow, leftPos + 70, topPos, 30, 25)
    With shape
        .Fill.ForeColor.RGB = arrowColor
        .Fill.Transparency = 0.2
        .Line.Color.RGB = arrowColor
        .Line.Weight = 1.5
    End With
    
    Set textBox = ws.Shapes.AddTextbox(msoAnchorNone, leftPos + 105, topPos + 2, 100, 26)
    With textBox.TextFrame
        .Clear
        .Word Wrap = True
        .VerticalAlignment = msoAnchorMiddle
        With .Characters.Font
            .Name = "Segoe UI"
            .Size = 10
            .Bold = True
            .Color.RGB = RGB(41, 84, 209)
        End With
        .Characters.Text = roleText
    End With
    
End Sub

Sub AddTitle(ws As Worksheet, leftPos As Double, topPos As Double, titleText As String, titleColor As Long)
    
    Dim textBox As Shape
    Set textBox = ws.Shapes.AddTextbox(msoAnchorNone, leftPos - 100, topPos, 300, 40)
    
    With textBox.TextFrame
        .Clear
        .Word Wrap = True
        .VerticalAlignment = msoAnchorMiddle
        .HorizontalAlignment = msoAnchorCenter
        
        With .Characters.Font
            .Name = "Segoe UI"
            .Size = 16
            .Bold = True
            .Color.RGB = titleColor
        End With
        .Characters.Text = titleText
    End With
    
End Sub

Sub AddLegend(ws As Worksheet, leftPos As Double, topPos As Double, color1 As Long, color2 As Long, color3 As Long)
    
    Dim legendBox As Shape
    Dim legendBg As Shape
    
    Set legendBg = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, 350, 80)
    With legendBg
        .Fill.ForeColor.RGB = RGB(245, 245, 245)
        .Line.Color.RGB = RGB(100, 100, 100)
        .Line.Weight = 1
        .Line.DashStyle = msoLineDash
    End With
    
    Dim legendText As String
    legendText = "Zeitleisten: Monatlich (Strategisch) → Wöchentlich (Taktisch) → Täglich (Operativ)"
    
    Set legendBox = ws.Shapes.AddTextbox(msoAnchorNone, leftPos + 10, topPos + 5, 330, 70)
    With legendBox.TextFrame
        .Clear
        .Word Wrap = True
        .VerticalAlignment = msoAnchorTop
        
        With .Characters.Font
            .Name = "Segoe UI"
            .Size = 9
            .Color.RGB = RGB(64, 64, 64)
        End With
        .Characters.Text = legendText
    End With
    
End Sub

Sub ClearPyramide()
    
    Dim ws As Worksheet
    Dim shape As Shape
    Dim i As Integer
    
    For i = 1 To ThisWorkbook.Sheets.Count
        If ThisWorkbook.Sheets(i).Name = "Führungsinstrumente" Then
            Set ws = ThisWorkbook.Sheets(i)
            Exit For
        End If
    Next i
    
    If Not ws Is Nothing Then
        For Each shape In ws.Shapes
            shape.Delete
        Next shape
        MsgBox "Pyramide gelöscht", vbInformation
    End If
    
End Sub
