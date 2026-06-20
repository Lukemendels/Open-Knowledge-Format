Attribute VB_Name = "OKFConfig"
' =====================================================================
'  OKF Config  — single source of truth for the bundle root path
'
'  Stores the root in the Windows registry via VBA's SaveSetting /
'  GetSetting API (HKCU\Software\VB and VBA Program Settings\OKF\Config).
'
'  Public surface:
'    BundleRoot()    — get the root (prompts picker if not yet set).
'    SetBundleRoot() — folder-picker UI; saves and refreshes dashboard.
'    BundleRootRaw() — non-prompting read; safe for dashboard display.
'    DistDir()       — the -dist sibling of BundleRoot (created if absent).
' =====================================================================

Option Explicit

Private Const DASH_SHEET       As String = "OKF Dashboard"
Private Const DASH_ROOT_CELL   As String = "B21"
Private Const REG_APP          As String = "OKF"
Private Const REG_SECTION      As String = "Config"
Private Const REG_KEY          As String = "BundleRoot"


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
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Select OKF Bundle Root Folder"
    dlg.AllowMultiSelect = False

    If dlg.Show <> -1 Then Exit Sub    ' user cancelled

    Dim root As String
    root = dlg.SelectedItems(1)
    If Right(root, 1) <> "\" Then root = root & "\"

    SaveSetting REG_APP, REG_SECTION, REG_KEY, root

    ' Refresh dashboard display if the sheet is open.
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(DASH_SHEET)
    If Not ws Is Nothing Then ws.Range(DASH_ROOT_CELL).Value = root
    On Error GoTo 0

    MsgBox "Bundle root set to:" & vbLf & root, vbInformation, "OKF Config"
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
