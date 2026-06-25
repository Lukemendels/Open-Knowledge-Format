Attribute VB_Name = "StickShiftHtmlTools"
' =====================================================================
'  StickShift HTML Tools  (conformant HTML Tool Installer)
'
'  Allows installing an HTML tool from a local file, copying it to
'  the context's -html sibling directory, and extracting its companion
'  skill to the skills/ folder.
'
'  Requires: StickShiftConfig (BundleRoot, HtmlDir),
'            StickShiftWriteApply (ApplyWriteEnvelopeText),
'            StickShiftIndexGenerator (GenerateStickShiftIndexes).
'  Microsoft ActiveX Data Objects 2.x (ADODB.Stream for UTF-8 I/O).
' =====================================================================

Option Explicit

Private Const msoFileDialogFilePicker As Long = 3


Public Sub InstallHtmlTool()
    Dim root As String
    root = StickShiftConfig.BundleRoot()
    If root = "" Then
        MsgBox "Bundle root not set - click Switch Context.", vbExclamation, "StickShift"
        Exit Sub
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(root) Then
        MsgBox "Bundle root not found: " & root, vbCritical, "StickShift"
        Exit Sub
    End If

    ' --- File Picker for HTML ---
    Dim dlg As Object
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    dlg.Title = "Select HTML Tool file"
    dlg.AllowMultiSelect = False
    dlg.Filters.Clear
    dlg.Filters.Add "HTML Files", "*.html"

    If dlg.Show <> -1 Then Exit Sub ' user cancelled

    Dim srcHtml As String: srcHtml = dlg.SelectedItems(1)
    Dim htmlDir As String: htmlDir = StickShiftConfig.HtmlDir()
    If htmlDir = "" Then Exit Sub

    Dim leafName As String: leafName = fso.GetFileName(srcHtml)
    fso.CopyFile srcHtml, htmlDir & leafName, True

    ' --- Extract embedded skill from the copied HTML ---
    Dim htmlContent As String
    htmlContent = ReadUtf8(htmlDir & leafName)

    Dim startScriptPos As Long
    startScriptPos = InStr(1, htmlContent, "<script", vbTextCompare)
    
    Dim scriptBlock As String: scriptBlock = ""
    Dim skillSlug As String: skillSlug = ""
    Dim skillMarkdown As String: skillMarkdown = ""
    
    Do While startScriptPos > 0
        Dim endScriptPos As Long
        endScriptPos = InStr(startScriptPos, htmlContent, "</script>", vbTextCompare)
        If endScriptPos = 0 Then Exit Do
        
        Dim tagContent As String
        tagContent = Mid(htmlContent, startScriptPos, endScriptPos - startScriptPos)
        
        If InStr(1, tagContent, "id=""stickshift-skill""", vbTextCompare) > 0 Or InStr(1, tagContent, "id='stickshift-skill'", vbTextCompare) > 0 Then
            scriptBlock = tagContent
            
            ' Extract data-skill-slug="..."
            Dim slugPos As Long
            slugPos = InStr(1, tagContent, "data-skill-slug=", vbTextCompare)
            If slugPos > 0 Then
                Dim valStartChar As String
                valStartChar = Mid(tagContent, slugPos + Len("data-skill-slug="), 1)
                Dim slugEndPos As Long
                If valStartChar = """" Or valStartChar = "'" Then
                    slugEndPos = InStr(slugPos + Len("data-skill-slug=") + 1, tagContent, valStartChar)
                    If slugEndPos > 0 Then
                        skillSlug = Mid(tagContent, slugPos + Len("data-skill-slug=") + 1, slugEndPos - (slugPos + Len("data-skill-slug=") + 1))
                    End If
                Else
                    slugEndPos = InStr(slugPos + Len("data-skill-slug="), tagContent, " ")
                    If slugEndPos > 0 Then
                        skillSlug = Mid(tagContent, slugPos + Len("data-skill-slug="), slugEndPos - (slugPos + Len("data-skill-slug=")))
                    End If
                End If
            End If
            
            ' Extract markdown inside the script tag
            Dim closeTagPos As Long
            closeTagPos = InStr(1, tagContent, ">")
            If closeTagPos > 0 Then
                skillMarkdown = Mid(tagContent, closeTagPos + 1)
            End If
            Exit Do
        End If
        
        startScriptPos = InStr(endScriptPos + 9, htmlContent, "<script", vbTextCompare)
    Loop

    skillSlug = Trim(skillSlug)
    skillMarkdown = Trim(skillMarkdown)

    If skillSlug = "" Or skillMarkdown = "" Then
        MsgBox "Failed to install: not a StickShift-compliant tool." & vbLf & _
               "The file must contain a script tag with id=""stickshift-skill"" and a data-skill-slug attribute.", _
               vbCritical, "StickShift"
        Exit Sub
    End If

    ' --- Rewrite tool: line for rename-safety ---
    Dim toolPos As Long
    toolPos = InStr(1, skillMarkdown, "tool:", vbTextCompare)
    If toolPos > 0 Then
        Dim eolPos As Long
        eolPos = InStr(toolPos, skillMarkdown, vbLf)
        If eolPos > 0 Then
            Dim prefix As String
            prefix = Left(skillMarkdown, toolPos - 1)
            Dim suffix As String
            suffix = Mid(skillMarkdown, eolPos)
            skillMarkdown = prefix & "tool: " & leafName & suffix
        End If
    End If

    ' --- Construct and apply VBA_WRITE envelope ---
    Dim envelope As String
    envelope = "<VBA_WRITE>" & vbLf & _
               "### FILE: skills/" & skillSlug & ".md" & vbLf & _
               skillMarkdown & vbLf & _
               "### END FILE" & vbLf & _
               "</VBA_WRITE>"

    Dim wCount As Long, sCount As Long
    Dim success As Boolean
    success = StickShiftWriteApply.ApplyWriteEnvelopeText(envelope, wCount, sCount)

    If success Then
        StickShiftIndexGenerator.GenerateStickShiftIndexes
        MsgBox "Installed tool: " & leafName & vbLf & _
               "Skill written to: skills/" & skillSlug & ".md", _
               vbInformation, "StickShift"
    Else
        MsgBox "Failed to write companion skill to skills/" & skillSlug & ".md", _
               vbCritical, "StickShift"
    End If
End Sub


Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function
