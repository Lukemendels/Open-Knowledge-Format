Attribute VB_Name = "OKFDashboard"
' =====================================================================
'  OKF Dashboard  — one-time setup for a macro-button control panel
'
'  Run CreateOKFDashboard once to build the "OKF Dashboard" sheet.
'  Re-run it at any time to reset the sheet (e.g. after importing
'  into a new workbook).  The sheet itself has no persistent state —
'  it is a pure UI layer over the three macros.
'
'  Requires the other three modules in the same workbook:
'    OKFWriteApply   → ApplyOKFWrite
'    OKFIndexGenerator → GenerateOKFIndexes
'    OKFLint         → RunOKFLint
' =====================================================================

Option Explicit

Sub CreateOKFDashboard()
    Const SHEET_NAME As String = "OKF Dashboard"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SHEET_NAME
    Else
        ws.Cells.Clear
        Dim shp As Shape
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If

    ws.Activate
    ws.DisplayGridlines = False

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("C").ColumnWidth = 2
    ws.Columns("D").ColumnWidth = 36
    ws.Columns("E").ColumnWidth = 2

    ws.Rows("1").RowHeight  = 10
    ws.Rows("2").RowHeight  = 32
    ws.Rows("3").RowHeight  = 16
    ws.Rows("4").RowHeight  = 16
    ws.Rows("5").RowHeight  = 46
    ws.Rows("6").RowHeight  = 34
    ws.Rows("7").RowHeight  = 10
    ws.Rows("8").RowHeight  = 46
    ws.Rows("9").RowHeight  = 34
    ws.Rows("10").RowHeight = 10
    ws.Rows("11").RowHeight = 46
    ws.Rows("12").RowHeight = 34
    ws.Rows("13").RowHeight = 10
    ws.Rows("14").RowHeight = 18

    ws.Cells.Interior.Color = RGB(248, 250, 252)

    Dim r As Range
    Set r = ws.Range("B2:D2"): r.Merge
    r.Value = "OKF Build Portfolio"
    r.Font.Size = 17: r.Font.Bold = True
    r.Font.Color = RGB(15, 23, 42)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    Set r = ws.Range("B3:D3"): r.Merge
    r.Value = "Workflow:  write change  " & Chr(8594) & "  apply  " & _
              Chr(8594) & "  regenerate  " & Chr(8594) & "  lint"
    r.Font.Size = 9: r.Font.Color = RGB(100, 116, 139)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    Set r = ws.Range("B3:D3")
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).Color     = RGB(226, 232, 240)
    r.Borders(xlEdgeBottom).Weight    = xlThin

    MakeButton ws, 5, _
        "1 — Apply Write Envelope", "ApplyOKFWrite", RGB(37, 99, 235), _
        "Paste the Portfolio Writer's <VBA_WRITE> output to the clipboard, then click." & Chr(10) & _
        "New builds are written directly; edits land as .proposed sidecars for review."

    MakeButton ws, 8, _
        "2 — Regenerate Index", "GenerateOKFIndexes", RGB(22, 163, 74), _
        "Rebuilds index.md from every .md build in the folder." & Chr(10) & _
        "Run after renaming a .proposed sidecar to .md to confirm an edit."

    MakeButton ws, 11, _
        "3 — Run Linter", "RunOKFLint", RGB(220, 85, 10), _
        "Scans the bundle: broken links, WIP violations, stalls, pending .proposed." & Chr(10) & _
        "Findings appear colour-coded in the " & Chr(34) & "OKF Lint Report" & Chr(34) & " sheet."

    Set r = ws.Range("B14:D14"): r.Merge
    r.Value = "Before first use: set BUNDLE_ROOT in OKFWriteApply.bas, " & _
              "OKFIndexGenerator.bas, and OKFLint.bas."
    r.Font.Size = 8: r.Font.Color = RGB(148, 163, 184)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    ws.Range("A1").Select
    MsgBox "Dashboard ready. Click any button to run its macro.", _
           vbInformation, "OKF Dashboard"
End Sub


Private Sub MakeButton(ByVal ws As Worksheet, ByVal btnRow As Long, _
                        ByVal caption As String, ByVal macroName As String, _
                        ByVal btnColor As Long, ByVal descText As String)

    Const INSET As Double = 4

    Dim cell As Range: Set cell = ws.Cells(btnRow, 2)

    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                cell.Left  + INSET, _
                                cell.Top   + INSET, _
                                cell.Width - INSET * 2, _
                                cell.Height - INSET * 2)
    With s
        .Name    = "btn_" & macroName
        .OnAction = macroName

        .Fill.ForeColor.RGB = btnColor
        .Fill.Solid
        .Line.Visible = msoFalse

        On Error Resume Next
        .Adjustments(1) = 0.18
        On Error GoTo 0

        With .Shadow
            .Visible     = msoTrue
            .OffsetX     = 0
            .OffsetY     = 2
            .Transparency = 0.75
            .Size        = 102
            .ForeColor.RGB = RGB(0, 0, 0)
        End With

        With .TextFrame2
            .TextRange.Text = caption
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange.Font
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Size  = 11
                .Bold  = msoTrue
            End With
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With

    Dim descRange As Range
    Set descRange = ws.Range(ws.Cells(btnRow + 1, 2), ws.Cells(btnRow + 1, 4))
    descRange.Merge
    descRange.Value              = descText
    descRange.Font.Size          = 8.5
    descRange.Font.Color         = RGB(51, 65, 85)
    descRange.VerticalAlignment  = xlVAlignCenter
    descRange.HorizontalAlignment = xlLeft
    descRange.WrapText           = True
    descRange.Interior.Color     = RGB(241, 245, 249)

    With descRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color     = RGB(226, 232, 240)
        .Weight    = xlThin
    End With
    With descRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color     = btnColor
        .Weight    = xlMedium
    End With
End Sub
