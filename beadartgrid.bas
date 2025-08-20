Option Explicit

' Converts grid unit row/col to Excel row index
Private Function ToExcelRow(unitRow As Long, unitCol As Long) As Long
    ToExcelRow = (unitRow - 1) * 2 + IIf(unitCol Mod 2 = 1, 1, 2)
End Function

Sub CreateBeadArtGrid()
    ' --- Constants ---
    Const GRID_HEIGHT_UNITS As Long = 80  ' Grid height (units)
    Const GRID_WIDTH_UNITS As Long = 130  ' Grid width (units)
    Const DESIGN_BLOCK_HEIGHT As Long = 8 ' Design input area height (units)

    Dim DESIGN_BG_COLOR: DESIGN_BG_COLOR = RGB(255, 255, 224) ' Design input area background color
    Dim ConditionalRules() As Variant: ConditionalRules = Array( _
        Array("2", RGB(192, 192, 192)), _
        Array("4", RGB(224, 224, 224)), _
        Array("3", RGB(244, 244, 192)), _
        Array("6", RGB(244, 244, 192)) _
    )

    ' --- Initialization ---
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet: Set ws = ThisWorkbook.ActiveSheet
    With ws
        .Unprotect
        .Cells.ClearContents
        .Cells.FormatConditions.Delete
        .Cells.Interior.ColorIndex = xlNone
        .Cells.Borders.LineStyle = xlNone
        .Cells.RowHeight = 10     ' Set row height (20px)
        .Cells.ColumnWidth = 2.75 ' Set column width (40px)
        .Cells.Font.Size = 8
        .Cells.Locked = True
    End With

    ' --- Grid setup ---
    Dim i As Long, j As Long, k As Long
    For i = 1 To GRID_WIDTH_UNITS
        For j = 1 To GRID_HEIGHT_UNITS
            ' Merge two rows for each grid unit, set borders and alignment
            With Range(ws.Cells(ToExcelRow(j, i), i), ws.Cells(ToExcelRow(j, i) + 1, i))
                .Merge
                .Borders.Weight = xlThin
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next j
    Next i

    For i = 1 To GRID_WIDTH_UNITS
        ' Calculate shifted units for design area
        Dim shiftUnit As Long: shiftUnit = -Int(-(i - 1) / 2)
        Dim designBottomUnit As Long: designBottomUnit = GRID_HEIGHT_UNITS - shiftUnit
        Dim designTopUnit As Long: designTopUnit = designBottomUnit - DESIGN_BLOCK_HEIGHT + 1

        Dim shiftUnitLeft As Long: shiftUnitLeft = -Int(-((i - 1) - 1) / 2)
        Dim designBottomUnitLeft As Long: designBottomUnitLeft = GRID_HEIGHT_UNITS - shiftUnitLeft
        Dim designTopUnitLeft As Long: designTopUnitLeft = designBottomUnitLeft - DESIGN_BLOCK_HEIGHT + 1

        Dim shiftUnitRight As Long: shiftUnitRight = -Int(-((i - 1) + 1) / 2)
        Dim designBottomUnitRight As Long: designBottomUnitRight = GRID_HEIGHT_UNITS - shiftUnitRight
        Dim designTopUnitRight As Long: designTopUnitRight = designBottomUnitRight - DESIGN_BLOCK_HEIGHT + 1

        ' a) Fill design input area with yellow
        For j = IIf(designTopUnit > 0, designTopUnit, 1) To designBottomUnit
            Dim isBaseLineUnitCol As Boolean, isBaseLineUnitRow As Boolean, remainder As Long
            remainder = (i + 13) Mod 17
            isBaseLineUnitCol = remainder Mod 2 = 0 And remainder <= 14
            isBaseLineUnitRow = (j - designTopUnit - Int(remainder / 2)) = 0
            With ws.Cells(ToExcelRow(j, i), i)
                .Interior.Color = DESIGN_BG_COLOR
                .MergeArea.Locked = False
                .Formula = IIf(isBaseLineUnitCol And isBaseLineUnitRow, "3", "")
            End With
        Next j

        ' b) Set formulas above design input area
        If i > 1 Then
            For j = IIf((designTopUnit - DESIGN_BLOCK_HEIGHT) > 0, designTopUnit - DESIGN_BLOCK_HEIGHT, 1) To designTopUnit - 1
                Dim targetTopUnit As Long: targetTopUnit = designTopUnitLeft + (j - (designTopUnit - DESIGN_BLOCK_HEIGHT))
                ws.Cells(ToExcelRow(j, i), i).Formula = "=2*" & ws.Cells(ToExcelRow(targetTopUnit, i - 1), i - 1).Address(False, False)
            Next j
        End If

        ' c) Set formulas below design input area
        If i < GRID_WIDTH_UNITS And designTopUnitRight > 0 Then
            For j = designBottomUnit + 1 To IIf((designBottomUnit + DESIGN_BLOCK_HEIGHT) < GRID_HEIGHT_UNITS, designBottomUnit + DESIGN_BLOCK_HEIGHT, GRID_HEIGHT_UNITS)
                Dim targetBottomUnit As Long: targetBottomUnit = designTopUnitRight + (j - (designBottomUnit + 1))
                ws.Cells(ToExcelRow(j, i), i).Formula = "=2*" & ws.Cells(ToExcelRow(targetBottomUnit, i + 1), i + 1).Address(False, False)
            Next j
        End If
    Next i

    ' --- Conditional formatting ---
    Dim targetRange As Range: Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(GRID_HEIGHT_UNITS * 2 + 1, GRID_WIDTH_UNITS))
    Dim cf As Variant
    For Each cf In ConditionalRules
        ' Add cell value-based conditional formatting
        With targetRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=cf(0))
            .Interior.Color = cf(1)
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.TintAndShade = 0
        End With
    Next cf

    ' --- Finalization ---
    ws.Cells(1, 1).Select
    ws.Protect
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Bead art grid created." & vbCrLf & "Yellow area is for design input.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error occurred: " & Err.Description, vbCritical
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
