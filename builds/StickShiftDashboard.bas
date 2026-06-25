Attribute VB_Name = "StickShiftDashboard"
' =====================================================================
'  StickShift Dashboard -- cockpit-style control panel
'
'  Run CreateStickShiftDashboard once to build the "StickShift" sheet.
'  Re-run at any time to reset the sheet.
'
'  Public surface:
'    CONTEXT_CELL          -- cell address where the current context path lives
'    RefreshContextDisplay -- rewrites CONTEXT_CELL from BundleRootRaw()
'
'  Requires:
'    StickShiftConfig         -> BundleRootRaw, SetBundleRoot
'    StickShiftWriteApply     -> ApplyStickShiftWrite
'    StickShiftIndexGenerator -> GenerateStickShiftIndexes
'    StickShiftLint           -> RunStickShiftLint
'    StickShiftContextBundle  -> BuildContextBundle
'    StickShiftBootstrap      -> BootstrapBundle
' =====================================================================

Option Explicit

Public Const CONTEXT_CELL As String = "B6"

Private Const SHEET_NAME As String = "StickShift"

' -- Constant shims: keep this module portable under pure late binding --
Private Const xlContinuous As Long = 1
Private Const xlThin As Long = 2
Private Const xlMedium As Long = -4138
Private Const xlEdgeLeft As Long = 7
Private Const xlEdgeTop As Long = 8
Private Const xlEdgeBottom As Long = 9
Private Const xlVAlignCenter As Long = -4108
Private Const xlCenter As Long = -4108
Private Const xlLeft As Long = -4131
Private Const msoFalse As Long = 0
Private Const msoTrue As Long = -1
Private Const msoShapeRoundedRectangle As Long = 5
Private Const msoAnchorMiddle As Long = 3
Private Const msoAlignCenter As Long = 2


Sub CreateStickShiftDashboard()
    ' -- Remove legacy sheet if present --
    Dim oldWs As Object
    On Error Resume Next
    Set oldWs = ThisWorkbook.Sheets("OKF Dashboard")
    On Error GoTo 0
    If Not oldWs Is Nothing Then
        Application.DisplayAlerts = False
        oldWs.Delete
        Application.DisplayAlerts = True
    End If

    ' -- Get or create the sheet --
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        If ThisWorkbook.ProtectStructure Then
            MsgBox "The workbook structure is protected." & vbCrLf & vbCrLf & _
                   "Go to Review > Unprotect Workbook, then run this again.", _
                   vbExclamation, "StickShift"
            Exit Sub
        End If
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.name = SHEET_NAME
    Else
        ws.Cells.Clear
        Dim shp As Object
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' -- Column widths (character units) --
    ws.Columns("A").ColumnWidth = 2      ' left margin
    ws.Columns("B").ColumnWidth = 22     ' left button
    ws.Columns("C").ColumnWidth = 4      ' inner gap
    ws.Columns("D").ColumnWidth = 22     ' right button
    ws.Columns("E").ColumnWidth = 2      ' right margin

    ' -- Row heights (points) --
    ws.Rows("1").RowHeight = 8       ' top padding
    ws.Rows("2").RowHeight = 36      ' title
    ws.Rows("3").RowHeight = 20      ' tagline
    ws.Rows("4").RowHeight = 10      ' gap
    ws.Rows("5").RowHeight = 16      ' context bar label
    ws.Rows("6").RowHeight = 22      ' context value (CONTEXT_CELL)
    ws.Rows("7").RowHeight = 10      ' gap before grid
    ws.Rows("8").RowHeight = 46      ' grid row 1 buttons
    ws.Rows("9").RowHeight = 60      ' grid row 1 descriptions
    ws.Rows("10").RowHeight = 46     ' grid row 2 (Switch Context centered)
    ws.Rows("11").RowHeight = 38     ' grid row 2 description
    ws.Rows("12").RowHeight = 46     ' grid row 3 buttons
    ws.Rows("13").RowHeight = 60     ' grid row 3 descriptions
    ws.Rows("14").RowHeight = 14     ' divider
    ws.Rows("15").RowHeight = 38     ' Install HTML Tool (demoted)
    ws.Rows("16").RowHeight = 28     ' Install HTML Tool label
    ws.Rows("17").RowHeight = 38     ' Open HTML Tool (demoted)
    ws.Rows("18").RowHeight = 28     ' Open HTML Tool label
    ws.Rows("19").RowHeight = 38     ' Initialize Context (demoted)
    ws.Rows("20").RowHeight = 28     ' Initialize Context label
    ws.Rows("21").RowHeight = 38     ' Restart Setup Wizard (demoted)
    ws.Rows("22").RowHeight = 28     ' Restart Setup Wizard label
    ws.Rows("23").RowHeight = 16     ' footer
    ws.Rows("24").RowHeight = 8      ' bottom padding

    ' -- Background --
    ws.Cells.Interior.Color = RGB(248, 250, 252)

    ' -- Title --
    Dim r As Object
    Set r = ws.Range("B2:D2"): r.Merge
    r.Value = "StickShift"
    r.Font.Size = 17: r.Font.Bold = True
    r.Font.Color = RGB(15, 23, 42)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    ' -- Tagline --
    Set r = ws.Range("B3:D3"): r.Merge
    r.Value = "Everything automatic AI promises - except your hand's on the StickShift."
    r.Font.Size = 9: r.Font.Color = RGB(100, 116, 139)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    r.Borders(xlEdgeBottom).Weight = xlThin

    ' -- Context bar label --
    Set r = ws.Range("B5:D5"): r.Merge
    r.Value = "Current context:"
    r.Font.Size = 8: r.Font.Bold = True
    r.Font.Color = RGB(71, 85, 105)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(241, 245, 249)
    r.Borders(xlEdgeTop).LineStyle = xlContinuous
    r.Borders(xlEdgeTop).Color = RGB(203, 213, 225)
    r.Borders(xlEdgeTop).Weight = xlThin

    ' -- Context value (CONTEXT_CELL = B6) -- display only --
    Dim rootVal As String
    rootVal = StickShiftConfig.BundleRootRaw()
    If rootVal = "" Then rootVal = "(none set - click Switch Context)"

    Set r = ws.Range("B6:D6"): r.Merge
    r.Value = rootVal
    r.Font.Size = 9
    r.Font.Color = IIf(StickShiftConfig.BundleRootRaw() = "", RGB(148, 163, 184), RGB(15, 23, 42))
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(241, 245, 249)
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).Color = RGB(226, 232, 240)
    r.Borders(xlEdgeBottom).Weight = xlThin
    r.Borders(xlEdgeLeft).LineStyle = xlContinuous
    r.Borders(xlEdgeLeft).Color = RGB(109, 40, 217)
    r.Borders(xlEdgeLeft).Weight = xlMedium

    ' -- Grid row 1: [Build Context Bundle] | [Apply Write Envelope] --
    MakeSideBySideButtons ws, 8, _
        "Build Context Bundle", "BuildContextBundle", RGB(109, 40, 217), _
        "Copy a <CONTEXT_REQUEST> from the chat, then click. Assembles" & vbLf & _
        "StickShift-context.md in the -dist folder to paste back into the chat.", _
        "Apply Write Envelope", "ApplyStickShiftWrite", RGB(37, 99, 235), _
        "Copy a <VBA_WRITE> block from the chat, then click. Writes the files" & vbLf & _
        "into your context. Every write is recorded in log.md."

    ' -- Grid row 2: [Switch Context] centered --
    MakeCenteredButton ws, 10, _
        "Switch Context", "SetBundleRoot", RGB(71, 85, 105), _
        "Pick or change the folder StickShift is pointed at. Switch anytime -" & vbLf & _
        "personal, shared, or per-project. Remembered per machine."

    ' -- Grid row 3: [Regenerate Index] | [Run Linter] --
    MakeSideBySideButtons ws, 12, _
        "Regenerate Index", "GenerateStickShiftIndexes", RGB(22, 163, 74), _
        "Rebuilds the listings. Runs automatically after every write; this is the" & vbLf & _
        "manual version.", _
        "Run Linter", "RunStickShiftLint", RGB(220, 85, 10), _
        "Scans the context: broken links, WIP violations, stalls," & vbLf & _
        "active-to-archived. Findings appear in the ""StickShift Lint Report"" sheet."

    ' -- Divider --
    Set r = ws.Range("B14:D14"): r.Merge
    r.Interior.Color = RGB(248, 250, 252)
    r.Borders(xlEdgeBottom).LineStyle = xlContinuous
    r.Borders(xlEdgeBottom).Color = RGB(203, 213, 225)
    r.Borders(xlEdgeBottom).Weight = xlThin

    ' -- Install HTML Tool (demoted) --
    MakeDemotedButton ws, 15, _
        "Install HTML Tool", "InstallHtmlTool", RGB(71, 85, 105), _
        "Install a new HTML tool (.html) and extract its companion skill."

    ' -- Open HTML Tool (demoted) --
    MakeDemotedButton ws, 17, _
        "Open HTML Tool", "OpenHtmlTool", RGB(109, 40, 217), _
        "Read <HTML_OPEN> block from clipboard and launch the tool."

    ' -- Initialize Context (demoted) --
    MakeDemotedButton ws, 19, _
        "Initialize Context", "BootstrapBundle", RGB(5, 150, 105), _
        "First-time setup: seeds a new, empty context." & vbLf & _
        "(Moves into the setup wizard later.)"

    ' -- Restart Setup Wizard (demoted) --
    MakeDemotedButton ws, 21, _
        "Restart Setup Wizard", "ReenableWizard", RGB(71, 85, 105), _
        "Reset and rerun the guided setup wizard from Step 1."

    ' -- Footer --
    Set r = ws.Range("B23:D23"): r.Merge
    r.Value = "Part of StickShift"
    r.Font.Size = 8: r.Font.Color = RGB(148, 163, 184)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    ws.Range("A1").Select
    MsgBox "StickShift dashboard ready.", vbInformation, "StickShift"
End Sub



Public Sub RefreshContextDisplay()
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim rootVal As String
    rootVal = StickShiftConfig.BundleRootRaw()

    Dim r As Object
    Set r = ws.Range(CONTEXT_CELL)
    If rootVal = "" Then
        r.Value = "(none set - click Switch Context)"
        r.Font.Color = RGB(148, 163, 184)
    Else
        r.Value = rootVal
        r.Font.Color = RGB(15, 23, 42)
    End If
End Sub


Private Sub MakeSideBySideButtons(ByVal ws As Object, ByVal btnRow As Long, _
    ByVal cap1 As String, ByVal macro1 As String, ByVal color1 As Long, ByVal desc1 As String, _
    ByVal cap2 As String, ByVal macro2 As String, ByVal color2 As Long, ByVal desc2 As String)

    Const inset As Double = 4

    Dim cellL As Object: Set cellL = ws.Cells(btnRow, 2)
    Dim sL As Object
    Set sL = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 cellL.Left + inset, cellL.Top + inset, _
                                 cellL.Width - inset * 2, cellL.Height - inset * 2)
    SetupButton sL, "btn_" & macro1, macro1, color1, cap1

    Dim cellR As Object: Set cellR = ws.Cells(btnRow, 4)
    Dim sR As Object
    Set sR = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 cellR.Left + inset, cellR.Top + inset, _
                                 cellR.Width - inset * 2, cellR.Height - inset * 2)
    SetupButton sR, "btn_" & macro2, macro2, color2, cap2

    StyleDesc ws.Cells(btnRow + 1, 2), color1, desc1
    StyleDesc ws.Cells(btnRow + 1, 4), color2, desc2
End Sub


Private Sub MakeCenteredButton(ByVal ws As Object, ByVal btnRow As Long, _
    ByVal cap As String, ByVal macroName As String, ByVal btnColor As Long, _
    ByVal descText As String)

    Const inset As Double = 4

    Dim cellL As Object: Set cellL = ws.Cells(btnRow, 2)
    Dim cellR As Object: Set cellR = ws.Cells(btnRow, 4)
    Dim totalWidth As Double
    totalWidth = cellR.Left + cellR.Width - cellL.Left

    Dim s As Object
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                cellL.Left + inset, cellL.Top + inset, _
                                totalWidth - inset * 2, cellL.Height - inset * 2)
    SetupButton s, "btn_" & macroName, macroName, btnColor, cap

    Dim descRange As Object
    Set descRange = ws.Range(ws.Cells(btnRow + 1, 2), ws.Cells(btnRow + 1, 4))
    descRange.Merge
    descRange.Value = descText
    descRange.Font.Size = 8.5
    descRange.Font.Color = RGB(51, 65, 85)
    descRange.VerticalAlignment = xlVAlignCenter
    descRange.HorizontalAlignment = xlLeft
    descRange.WrapText = True
    descRange.Interior.Color = RGB(241, 245, 249)
    With descRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(226, 232, 240)
        .Weight = xlThin
    End With
    With descRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = btnColor
        .Weight = xlMedium
    End With
End Sub


Private Sub MakeDemotedButton(ByVal ws As Object, ByVal btnRow As Long, _
    ByVal cap As String, ByVal macroName As String, ByVal btnColor As Long, _
    ByVal labelText As String)

    Const inset As Double = 4
    Const DEMOTE_INSET As Double = 30

    Dim cellL As Object: Set cellL = ws.Cells(btnRow, 2)
    Dim cellR As Object: Set cellR = ws.Cells(btnRow, 4)
    Dim totalWidth As Double
    totalWidth = cellR.Left + cellR.Width - cellL.Left

    Dim s As Object
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                cellL.Left + DEMOTE_INSET, cellL.Top + inset, _
                                totalWidth - DEMOTE_INSET * 2, cellL.Height - inset * 2)
    SetupButton s, "btn_" & macroName, macroName, btnColor, cap
    s.TextFrame2.TextRange.Font.Size = 9

    Dim lbl As Object
    Set lbl = ws.Range(ws.Cells(btnRow + 1, 2), ws.Cells(btnRow + 1, 4))
    lbl.Merge
    lbl.Value = labelText
    lbl.Font.Size = 8
    lbl.Font.Color = RGB(148, 163, 184)
    lbl.Font.Italic = True
    lbl.VerticalAlignment = xlVAlignCenter
    lbl.HorizontalAlignment = xlCenter
    lbl.WrapText = True
    lbl.Interior.Color = RGB(248, 250, 252)
End Sub


Private Sub SetupButton(ByVal s As Object, ByVal shapeName As String, _
                         ByVal macroName As String, ByVal btnColor As Long, _
                         ByVal cap As String)
    With s
        .name = shapeName
        .OnAction = macroName
        .Fill.ForeColor.RGB = btnColor
        .Fill.Solid
        .line.Visible = msoFalse
        On Error Resume Next
        .Adjustments(1) = 0.18
        On Error GoTo 0
        With .Shadow
            .Visible = msoTrue
            .OffsetX = 0
            .OffsetY = 2
            .Transparency = 0.75
            .Size = 102
            .ForeColor.RGB = RGB(0, 0, 0)
        End With
        With .TextFrame2
            .TextRange.Text = cap
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange.Font
                .Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Size = 11
                .Bold = msoTrue
            End With
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub


Private Sub StyleDesc(ByVal cell As Object, ByVal btnColor As Long, ByVal descText As String)
    cell.Value = descText
    cell.Font.Size = 8.5
    cell.Font.Color = RGB(51, 65, 85)
    cell.VerticalAlignment = xlVAlignCenter
    cell.HorizontalAlignment = xlLeft
    cell.WrapText = True
    cell.Interior.Color = RGB(241, 245, 249)
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(226, 232, 240)
        .Weight = xlThin
    End With
    With cell.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = btnColor
        .Weight = xlMedium
    End With
End Sub
