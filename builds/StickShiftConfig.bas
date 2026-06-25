Attribute VB_Name = "StickShiftConfig"
' =====================================================================
'  StickShift Config -- OKF-compliant
'  Single source of truth for the context root path.
'
'  Stores the root in the Windows registry via VBA's SaveSetting /
'  GetSetting API (HKCU\Software\VB and VBA Program Settings\StickShift\Config).
'
'  Public surface:
'    BundleRoot()    - get the root (prompts picker if not yet set).
'    SetBundleRoot() - folder-picker UI; saves and refreshes dashboard.
'    BundleRootRaw() - non-prompting read; safe for dashboard display.
'    DistDir()       - the -dist sibling of BundleRoot (created if absent).
'
'  Registry note: renaming REG_APP from "OKF" to "StickShift" orphans any
'  root saved under the old key. Zero impact for a fresh user; existing
'  users re-pick their context once after upgrading.
' =====================================================================

Option Explicit

Private Const REG_APP          As String = "StickShift"
Private Const REG_SECTION      As String = "Config"
Private Const REG_KEY          As String = "BundleRoot"

' -- Constant shims: keep this module portable under pure late binding --
Private Const msoFileDialogFolderPicker As Long = 4


Public Function BundleRoot() As String
    Dim root As String
    root = GetSetting(REG_APP, REG_SECTION, REG_KEY, "")
    If root = "" Then
        SetBundleRoot
        root = GetSetting(REG_APP, REG_SECTION, REG_KEY, "")
    End If
    If root = "" Then BundleRoot = "": Exit Function
    If Right(root, 1) <> "\" Then root = root & "\"
    BundleRoot = root
End Function


Public Sub SetBundleRoot()
    Dim dlg As Object
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Select StickShift Context Folder"
    dlg.AllowMultiSelect = False

    If dlg.Show <> -1 Then Exit Sub    ' user cancelled

    Dim root As String
    root = dlg.SelectedItems(1)
    If Right(root, 1) <> "\" Then root = root & "\"

    SaveSetting REG_APP, REG_SECTION, REG_KEY, root

    ' Refresh context display on dashboard if the sheet is present.
    On Error Resume Next
    StickShiftDashboard.RefreshContextDisplay
    On Error GoTo 0

    MsgBox "Context set to:" & vbLf & root, vbInformation, "StickShift"
End Sub


Public Function BundleRootRaw() As String
    BundleRootRaw = GetSetting(REG_APP, REG_SECTION, REG_KEY, "")
End Function


Public Function DistDir() As String
    Dim root As String
    root = BundleRoot()
    If root = "" Then DistDir = "": Exit Function

    ' Strip the trailing backslash, append -dist\.
    Dim stripped As String: stripped = Left(root, Len(root) - 1)
    Dim dist As String:     dist = stripped & "-dist\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dist) Then
        On Error Resume Next
        fso.CreateFolder dist
        On Error GoTo 0
    End If

    DistDir = dist
End Function


Public Function HtmlDir() As String
    Dim root As String
    root = BundleRoot()
    If root = "" Then HtmlDir = "": Exit Function

    ' Strip the trailing backslash, append -html\.
    Dim stripped As String: stripped = Left(root, Len(root) - 1)
    Dim h As String:        h = stripped & "-html\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(h) Then
        On Error Resume Next
        fso.CreateFolder h
        On Error GoTo 0
    End If

    HtmlDir = h
End Function
