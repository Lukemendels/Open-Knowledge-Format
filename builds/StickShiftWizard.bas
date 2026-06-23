Attribute VB_Name = "StickShiftWizard"
' =====================================================================
'  StickShift Wizard -- guided first-run setup tutorial
'
'  Walks the user through the real StickShift loop one button at a time,
'  using the seeded setup-interview skill as the first round-trip, then
'  hands off to the canonical steady-state dashboard.
'
'  State persists in workbook Name "SS_WizardState":
'    missing / "step:N" -> show wizard at step N (default 1)
'    "done"             -> skip straight to canonical dashboard
'    "skipped"          -> skip straight to canonical dashboard
'
'  Steps:
'    1  Switch Context       (Wiz_SwitchContext)
'    2  Initialize Context   (Wiz_InitializeContext)
'    3  Build Context Bundle (Wiz_BuildBundle)
'    4  Apply Write Envelope (Wiz_ApplyWrite)
'    5  Finish               (Wiz_Finish)
'
'  Requires: StickShiftConfig, StickShiftBootstrap, StickShiftContextBundle,
'            StickShiftWriteApply, StickShiftDashboard.
' =====================================================================

Option Explicit

Private Const SHEET_NAME       As String = "StickShift"
Private Const STATE_NAME       As String = "SS_WizardState"
Private Const TOTAL_STEPS      As Long = 5

' -- Constant shims: keep module portable under pure late binding --
Private Const xlContinuous              As Long = 1
Private Const xlThin                    As Long = 2
Private Const xlMedium                  As Long = -4138
Private Const xlEdgeLeft                As Long = 7
Private Const xlEdgeTop                 As Long = 8
Private Const xlEdgeBottom              As Long = 9
Private Const xlVAlignCenter            As Long = -4108
Private Const xlCenter                  As Long = -4108
Private Const xlLeft                    As Long = -4131
Private Const msoFalse                  As Long = 0
Private Const msoTrue                   As Long = -1
Private Const msoShapeRoundedRectangle  As Long = 5
Private Const msoAnchorMiddle           As Long = 3
Private Const msoAlignCenter            As Long = 2


' =====================================================================
'  Entry / routing
' =====================================================================

Public Sub Auto_Open()
    Dim stateVal As String
    Dim step As Long

    stateVal = GetWizardState()

    If stateVal = "done" Or stateVal = "skipped" Then
        StickShiftDashboard.CreateStickShiftDashboard
        Exit Sub
    End If

    If Left(stateVal, 5) = "step:" Then
        step = CLng(Mid(stateVal, 6))
        If step < 1 Or step > TOTAL_STEPS Then step = 1
    Else
        step = 1
    End If

    SetWizardState "step:" & CStr(step)
    RenderWizard step
End Sub


Public Sub StartStickShiftSetup()
    SetWizardState "step:1"
    RenderWizard 1
End Sub


' =====================================================================
'  State persistence via workbook defined Name
' =====================================================================

Public Function GetWizardState() As String
    Dim nm As Object
    Dim raw As String

    On Error Resume Next
    Set nm = ThisWorkbook.Names(STATE_NAME)
    On Error GoTo 0

    If nm Is Nothing Then
        GetWizardState = ""
        Exit Function
    End If

    raw = nm.RefersTo
    ' RefersTo for a string constant Name looks like ="value"
    If Left(raw, 2) = "=""" And Right(raw, 1) = """" Then
        GetWizardState = Mid(raw, 3, Len(raw) - 3)
    Else
        GetWizardState = ""
    End If
End Function


Public Sub SetWizardState(ByVal stateVal As String)
    On Error Resume Next
    ThisWorkbook.Names(STATE_NAME).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=STATE_NAME, RefersTo:="=""" & stateVal & """"
End Sub


' =====================================================================
'  Wizard rendering
' =====================================================================

Private Sub RenderWizard(ByVal step As Long)
    Dim ws As Object
    Dim r As Object
    Dim s As Object
    Dim skipBtn As Object
    Dim cellL As Object
    Dim cellR As Object
    Dim totalWidth As Double
    Dim inset As Double
    Dim i As Long
    Dim btnBaseRow As Long
    Dim footerRow As Long
    Dim stepLabel As String
    Dim bRow As Long

    If step < 1 Then step = 1
    If step > TOTAL_STEPS Then step = TOTAL_STEPS

    ' -- Get or create the sheet --
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        If ThisWorkbook.ProtectStructure Then
            MsgBox "The workbook structure is protected." & vbLf & _
                   "Go to Review > Unprotect Workbook, then run again.", _
                   vbExclamation, "StickShift"
            Exit Sub
        End If
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SHEET_NAME
    Else
        ws.Cells.Clear
        Do While ws.Shapes.Count > 0
            ws.Shapes(1).Delete
        Loop
    End If

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    ' -- Column widths (character units) --
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 22
    ws.Columns("C").ColumnWidth = 4
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 2

    ' -- Row heights (points) --
    ws.Rows("1").RowHeight = 8        ' top pad
    ws.Rows("2").RowHeight = 30       ' wizard title
    ws.Rows("3").RowHeight = 20       ' step indicator
    ws.Rows("4").RowHeight = 8        ' inner gap
    ws.Rows("5").RowHeight = 72       ' coach text
    ws.Rows("6").RowHeight = 12       ' gap before buttons

    btnBaseRow = 7
    For i = 0 To TOTAL_STEPS - 1
        ws.Rows(CStr(btnBaseRow + i)).RowHeight = 46
    Next i
    footerRow = btnBaseRow + TOTAL_STEPS
    ws.Rows(CStr(footerRow)).RowHeight = 16

    ' -- Background --
    ws.Cells.Interior.Color = RGB(248, 250, 252)

    ' -- Wizard title (row 2) --
    Set r = ws.Range("B2:D2"): r.Merge
    r.Value = "StickShift Setup"
    r.Font.Size = 17
    r.Font.Bold = True
    r.Font.Color = RGB(15, 23, 42)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    ' -- Step indicator (row 3) --
    stepLabel = "Step " & CStr(step) & " of " & CStr(TOTAL_STEPS) & _
                "  --  " & StepTitle(step)
    Set r = ws.Range("B3:D3"): r.Merge
    r.Value = stepLabel
    r.Font.Size = 9
    r.Font.Color = RGB(100, 116, 139)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)
    With r.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(226, 232, 240)
        .Weight = xlThin
    End With

    ' -- Coach text (row 5) --
    Set r = ws.Range("B5:D5"): r.Merge
    r.Value = CoachText(step)
    r.Font.Size = 9
    r.Font.Color = RGB(15, 23, 42)
    r.VerticalAlignment = xlVAlignCenter
    r.HorizontalAlignment = xlLeft
    r.WrapText = True
    r.Interior.Color = RGB(241, 245, 249)
    With r.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = RGB(203, 213, 225)
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(203, 213, 225)
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = StepColor(step)
        .Weight = xlMedium
    End With

    ' -- Skip Setup button: top-right corner of coach panel, neutral gray --
    Set cellR = ws.Cells(2, 4)
    Set skipBtn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                     cellR.Left + 4, cellR.Top + 4, _
                                     cellR.Width - 8, cellR.Height - 8)
    AddWizardButton skipBtn, "btn_SkipSetup", "SkipSetup", RGB(100, 116, 139), "Skip Setup"
    skipBtn.TextFrame2.TextRange.Font.Size = 9

    ' -- Stacked buttons: one per step, unlocked so far --
    inset = 6
    Set cellL = ws.Cells(btnBaseRow, 2)
    Set cellR = ws.Cells(btnBaseRow, 4)
    totalWidth = cellR.Left + cellR.Width - cellL.Left

    For i = 1 To step
        bRow = btnBaseRow + (i - 1)
        Set cellL = ws.Cells(bRow, 2)
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                    cellL.Left + inset, cellL.Top + inset, _
                                    totalWidth - inset * 2, cellL.Height - inset * 2)
        If i < TOTAL_STEPS Then
            AddWizardButton s, "wiz_btn_" & CStr(i), StepMacro(i), StepColor(i), StepTitle(i)
        Else
            AddWizardButton s, "wiz_btn_finish", "Wiz_Finish", RGB(22, 163, 74), "Finish"
        End If
    Next i

    ' -- Footer --
    Set r = ws.Range(ws.Cells(footerRow, 2), ws.Cells(footerRow, 4)): r.Merge
    r.Value = "Part of StickShift"
    r.Font.Size = 8
    r.Font.Color = RGB(148, 163, 184)
    r.VerticalAlignment = xlVAlignCenter
    r.Interior.Color = RGB(248, 250, 252)

    ws.Range("A1").Select
End Sub


Private Function StepTitle(ByVal step As Long) As String
    Select Case step
        Case 1:     StepTitle = "Switch Context"
        Case 2:     StepTitle = "Initialize Context"
        Case 3:     StepTitle = "Build Context Bundle"
        Case 4:     StepTitle = "Apply Write Envelope"
        Case 5:     StepTitle = "Finish"
        Case Else:  StepTitle = ""
    End Select
End Function


Private Function StepMacro(ByVal step As Long) As String
    Select Case step
        Case 1:     StepMacro = "Wiz_SwitchContext"
        Case 2:     StepMacro = "Wiz_InitializeContext"
        Case 3:     StepMacro = "Wiz_BuildBundle"
        Case 4:     StepMacro = "Wiz_ApplyWrite"
        Case 5:     StepMacro = "Wiz_Finish"
        Case Else:  StepMacro = ""
    End Select
End Function


Private Function StepColor(ByVal step As Long) As Long
    Select Case step
        Case 1:     StepColor = RGB(71, 85, 105)    ' Switch Context slate
        Case 2:     StepColor = RGB(5, 150, 105)    ' Initialize Context teal
        Case 3:     StepColor = RGB(109, 40, 217)   ' Build Context Bundle purple
        Case 4:     StepColor = RGB(37, 99, 235)    ' Apply Write Envelope blue
        Case 5:     StepColor = RGB(22, 163, 74)    ' Finish green
        Case Else:  StepColor = RGB(71, 85, 105)
    End Select
End Function


Private Function CoachText(ByVal step As Long) As String
    Select Case step
        Case 1
            CoachText = "Point StickShift at a folder. Click Switch Context and pick or make one, e.g. C:\StickShift."
        Case 2
            CoachText = "Seed this folder with starter files - including two skills the agent can use. Click Initialize Context."
        Case 3
            CoachText = "Now the loop. In your AI chat ask: 'Run the setup interview skill.' The AI replies with a <CONTEXT_REQUEST>. Copy the whole block, then click Build Context Bundle and paste the result back into the chat."
        Case 4
            CoachText = "The AI interviews you, then returns a <VBA_WRITE> block with your profile. Copy it, click Apply Write Envelope, and it writes _foundation/00-operating-profile.md."
        Case 5
            CoachText = "That was a skill - a written procedure the agent followed. From now on you can ask it to write a skill for anything you do repeatedly, and it will reuse it next time. Setup is done."
        Case Else
            CoachText = ""
    End Select
End Function


' =====================================================================
'  Button factory (private; mirrors dashboard style without importing it)
' =====================================================================

Private Sub AddWizardButton(ByVal s As Object, ByVal shapeName As String, _
                              ByVal macroName As String, ByVal btnColor As Long, _
                              ByVal cap As String)
    With s
        .Name = shapeName
        .OnAction = macroName
        .Fill.ForeColor.RGB = btnColor
        .Fill.Solid
        .Line.Visible = msoFalse
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


' =====================================================================
'  Wizard wrappers: call the real macro, gate on world state, advance
' =====================================================================

Public Sub Wiz_SwitchContext()
    StickShiftConfig.SetBundleRoot
    If StickShiftConfig.BundleRootRaw() <> "" Then
        SetWizardState "step:2"
        RenderWizard 2
    End If
End Sub


Public Sub Wiz_InitializeContext()
    Dim root As String
    Dim fso As Object

    StickShiftBootstrap.BootstrapBundle

    root = StickShiftConfig.BundleRootRaw()
    If root = "" Then Exit Sub
    If Right(root, 1) <> "\" Then root = root & "\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(root & "_foundation\00-operating-profile.md") Then
        SetWizardState "step:3"
        RenderWizard 3
    End If
End Sub


Public Sub Wiz_BuildBundle()
    StickShiftContextBundle.BuildContextBundle
    SetWizardState "step:4"
    RenderWizard 4
End Sub


Public Sub Wiz_ApplyWrite()
    StickShiftWriteApply.ApplyStickShiftWrite
    SetWizardState "step:5"
    RenderWizard 5
End Sub


Public Sub Wiz_Finish()
    SetWizardState "done"
    StickShiftDashboard.CreateStickShiftDashboard
End Sub


Public Sub SkipSetup()
    SetWizardState "skipped"
    StickShiftDashboard.CreateStickShiftDashboard
End Sub
