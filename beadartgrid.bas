Option Explicit

' Convert grid unit row/col to Excel row index
Private Function ToExcelRow(unitRow As Long, unitCol As Long) As Long
    ToExcelRow = (unitRow - 1) * 2 + IIf(unitCol Mod 2 = 1, 1, 2)
End Function

' Main procedure to create bead art grid
Sub CreateBeadArtGrid()
    On Error GoTo ErrorHandler

    ' --- Read config ---
    Dim cs As Worksheet: Set cs = ThisWorkbook.Sheets("__CONFIG__")
    Dim GRID_WIDTH_UNITS As Long: GRID_WIDTH_UNITS = Int(cs.Cells(1, 3).Value)
    Dim GRID_HEIGHT_UNITS As Long: GRID_HEIGHT_UNITS = Int(cs.Cells(2, 3).Value)
    Dim CELL_WIDTH As Double: CELL_WIDTH = CDbl(cs.Cells(3, 3).Value)
    Dim CELL_HEIGHT As Double: CELL_HEIGHT = CDbl(cs.Cells(4, 3).Value)
    Dim FONT_SIZE As Long: FONT_SIZE = Int(cs.Cells(5, 3).Value)
    Dim DESIGN_BLOCK_HEIGHT As Long: DESIGN_BLOCK_HEIGHT = Int(cs.Cells(6, 3).Value)
    Dim DESIGN_BG_COLOR As Long: DESIGN_BG_COLOR = cs.Cells(7, 3).Interior.Color

    Const NUMBER_FOR_DESIGN_BASELINE = 1
    Dim ConditionalRules As Collection: Set ConditionalRules = New Collection
    Dim k As Long
    For k = 1 To 9
        ConditionalRules.Add Array(Str(k), cs.Cells(8 + k, 2).Interior.Color)
        ConditionalRules.Add Array(Str(k + 0.01), cs.Cells(8 + k, 3).Interior.Color)
    Next k

    ' --- Init worksheet ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets.Add(After:=cs)
    With ws
        .Unprotect
        .Cells.ClearContents
        .Cells.FormatConditions.Delete
        .Cells.Interior.ColorIndex = xlNone
        .Cells.Borders.LineStyle = xlNone
        .Cells.RowHeight = CELL_HEIGHT
        .Cells.ColumnWidth = CELL_WIDTH
        .Cells.Font.Size = FONT_SIZE
        .Cells.Locked = True
    End With

    ' --- Create grid ---
    Dim i As Long, j As Long
    For i = 1 To GRID_WIDTH_UNITS
        For j = 1 To GRID_HEIGHT_UNITS
            With Range(ws.Cells(ToExcelRow(j, i), i), ws.Cells(ToExcelRow(j, i) + 1, i))
                .Merge
                .Borders.Weight = xlThin
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
        Next j
    Next i

    ' --- Fill design area and set formulas ---
    For i = 1 To GRID_WIDTH_UNITS
        Dim shiftUnit As Long: shiftUnit = -Int(-(i - 1) / 2)
        Dim designBottomUnit As Long: designBottomUnit = GRID_HEIGHT_UNITS - shiftUnit
        Dim designTopUnit As Long: designTopUnit = designBottomUnit - DESIGN_BLOCK_HEIGHT + 1

        Dim shiftUnitLeft As Long: shiftUnitLeft = -Int(-((i - 1) - 1) / 2)
        Dim designTopUnitLeft As Long: designTopUnitLeft = GRID_HEIGHT_UNITS - shiftUnitLeft - DESIGN_BLOCK_HEIGHT + 1

        Dim shiftUnitRight As Long: shiftUnitRight = -Int(-((i - 1) + 1) / 2)
        Dim designTopUnitRight As Long: designTopUnitRight = GRID_HEIGHT_UNITS - shiftUnitRight - DESIGN_BLOCK_HEIGHT + 1

        ' Fill design input area
        For j = IIf(designTopUnit > 0, designTopUnit, 1) To designBottomUnit
            Dim isBaseLineUnitCol As Boolean, isBaseLineUnitRow As Boolean, remainder As Long
            remainder = (i + 2 * DESIGN_BLOCK_HEIGHT - 3) Mod (2 * DESIGN_BLOCK_HEIGHT + 1)
            isBaseLineUnitCol = remainder Mod 2 = 0 And remainder < (2 * DESIGN_BLOCK_HEIGHT - 1)
            isBaseLineUnitRow = (j - designTopUnit - Int(remainder / 2)) = 0
            With ws.Cells(ToExcelRow(j, i), i)
                .Interior.Color = DESIGN_BG_COLOR
                .MergeArea.Locked = False
                .Formula = IIf(isBaseLineUnitCol And isBaseLineUnitRow, Str(NUMBER_FOR_DESIGN_BASELINE), "")
            End With
        Next j

        ' Set formulas above design area
        If i > 1 Then
            For j = IIf((designTopUnit - DESIGN_BLOCK_HEIGHT) > 0, designTopUnit - DESIGN_BLOCK_HEIGHT, 1) To designTopUnit - 1
                Dim targetTopUnit As Long: targetTopUnit = designTopUnitLeft + (j - (designTopUnit - DESIGN_BLOCK_HEIGHT))
                ws.Cells(ToExcelRow(j, i), i).Formula = "=0.01+" & ws.Cells(ToExcelRow(targetTopUnit, i - 1), i - 1).Address(False, False)
            Next j
        End If

        ' Set formulas below design area
        If i < GRID_WIDTH_UNITS And designTopUnitRight > 0 Then
            For j = designBottomUnit + 1 To IIf((designBottomUnit + DESIGN_BLOCK_HEIGHT) < GRID_HEIGHT_UNITS, designBottomUnit + DESIGN_BLOCK_HEIGHT, GRID_HEIGHT_UNITS)
                Dim targetBottomUnit As Long: targetBottomUnit = designTopUnitRight + (j - (designBottomUnit + 1))
                ws.Cells(ToExcelRow(j, i), i).Formula = "=0.01+" & ws.Cells(ToExcelRow(targetBottomUnit, i + 1), i + 1).Address(False, False)
            Next j
        End If
    Next i

    ' --- Conditional formatting ---
    Dim targetRange As Range: Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(GRID_HEIGHT_UNITS * 2 + 1, GRID_WIDTH_UNITS))
    Dim cf As Variant
    For Each cf In ConditionalRules
        With targetRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:=cf(0))
            .Interior.Color = cf(1)
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.TintAndShade = 0
        End With
    Next cf

    ' --- Finalize ---
    ws.Cells(ToExcelRow(GRID_HEIGHT_UNITS - DESIGN_BLOCK_HEIGHT + 1, 1), 1).Select
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

' Create config sheet for grid parameters and colors
Sub CreateConfigSheet()
    Const ConfigSheetName As String = "__CONFIG__"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(ConfigSheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Set ws = ThisWorkbook.Worksheets.Add
    With ws
        .Name = ConfigSheetName
        .Unprotect
        .Range("A1:C17").Locked = True
    End With

    Dim configValues As Variant
    configValues = Array( _
        Array("Units", "Width", 130), _
        Array("Units", "Height", 80), _
        Array("Cells", "Width", 2.75), _
        Array("Cells", "Height", 10), _
        Array("Cells", "Fontsize", 8), _
        Array("Design", "Units", 8) _
    )
    Dim i As Integer
    For i = 0 To UBound(configValues)
        ws.Cells(i + 1, 1).Value = configValues(i)(0)
        ws.Cells(i + 1, 2).Value = configValues(i)(1)
        ws.Cells(i + 1, 3).Value = configValues(i)(2)
        ws.Cells(i + 1, 3).Locked = False
    Next i

    ws.Cells(7, 1).Value = "Design": ws.Cells(7, 2).Value = "Color"
    ws.Cells(7, 3).Interior.Color = RGB(255, 255, 224): ws.Cells(7, 3).Locked = False ' Light yellow

    ws.Cells(8, 1).Value = "Colors": ws.Cells(8, 2).Value = "Color": ws.Cells(8, 3).Value = "Shadow"

    ws.Cells(9, 1).Value = "Baseline"
    ws.Cells(9, 2).Interior.Color = RGB(255, 255, 153): ws.Cells(9, 2).Locked = False ' Light yellow
    ws.Cells(9, 3).Interior.Color = RGB(255, 255, 153): ws.Cells(9, 3).Locked = False ' Light yellow

    Dim beadColors As Variant
    beadColors = Array( _
        Array("2", RGB(0, 0, 0), RGB(128, 128, 128)), _
        Array("3", RGB(128, 64, 0), RGB(128, 128, 128)), _
        Array("4", RGB(0, 112, 192), RGB(128, 128, 128)), _
        Array("5", RGB(255, 0, 0), RGB(128, 128, 128)), _
        Array("6", RGB(0, 176, 80), RGB(128, 128, 128)), _
        Array("7", RGB(255, 255, 0), RGB(128, 128, 128)), _
        Array("8", RGB(0, 176, 240), RGB(128, 128, 128)), _
        Array("9", RGB(128, 0, 128), RGB(128, 128, 128)) _
    )
    For i = 0 To UBound(beadColors)
        ws.Cells(10 + i, 1).Value = beadColors(i)(0)
        ws.Cells(10 + i, 2).Interior.Color = beadColors(i)(1): ws.Cells(10 + i, 2).Locked = False
        ws.Cells(10 + i, 3).Interior.Color = beadColors(i)(2): ws.Cells(10 + i, 3).Locked = False
    Next i

    With ws.Buttons.Add(ws.Cells(1, 5).Left, ws.Cells(1, 5).Top, ws.Range(ws.Cells(1, 5), ws.Cells(17, 7)).Width, ws.Range(ws.Cells(1, 5), ws.Cells(17, 7)).Height)
        .OnAction = "CreateBeadArtGrid"
        .Caption = "Create Bead Art Grid"
        .Name = "btnCreateBeadArtGrid"
    End With
    ws.Protect
End Sub
