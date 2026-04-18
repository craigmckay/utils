Attribute VB_Name = "UtilExport"
Option Explicit

' Export every VBA module in the current project to C:\Dev\utils\Word\Normal
' as individual source files. Standard modules -> .bas, class modules -> .cls,
' UserForms -> .frm (+ .frx). Document-code modules (ThisDocument etc.) are
' skipped because they cannot be cleanly re-imported.
'
' Requires: Word > File > Options > Trust Center > Trust Center Settings >
'           Macro Settings > "Trust access to the VBA project object model"

Public Sub ExportNormalModules()
    Const ExportFolder As String = "C:\Dev\utils\Word\Normal\"

    Dim vbProj As Object
    Dim vbComp As Object
    Dim ext As String
    Dim filePath As String
    Dim exported As Long
    Dim skipped As Long

    ' Make sure the destination folder exists (creates the final folder only;
    ' parent path must already exist).
    If Len(Dir(ExportFolder, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir ExportFolder
        On Error GoTo 0
    End If

    ' Use the project that contains this macro.
    On Error Resume Next
    Set vbProj = ThisDocument.VBProject
    On Error GoTo 0

    If vbProj Is Nothing Then
        MsgBox "Could not access the current VBProject." & vbCrLf & vbCrLf & _
               "Enable: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
               "Macro Settings > 'Trust access to the VBA project object model'.", _
               vbExclamation, "Export Modules"
        Exit Sub
    End If

    exported = 0
    skipped = 0

    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case 1   ' vbext_ct_StdModule
                ext = ".bas"
            Case 2   ' vbext_ct_ClassModule
                ext = ".cls"
            Case 3   ' vbext_ct_MSForm
                ext = ".frm"
            Case Else
                ' 100 = vbext_ct_Document (ThisDocument, etc.) - skip
                ext = ""
        End Select

        If Len(ext) > 0 Then
            filePath = ExportFolder & vbComp.Name & ext
            ' Overwrite any previous export
            If Len(Dir(filePath)) > 0 Then Kill filePath
            vbComp.Export filePath
            exported = exported + 1
        Else
            skipped = skipped + 1
        End If
    Next vbComp

    MsgBox "Exported " & exported & " module(s) to " & ExportFolder & vbCrLf & _
           "Skipped " & skipped & " document-code module(s).", _
           vbInformation, "Export Modules"
End Sub

