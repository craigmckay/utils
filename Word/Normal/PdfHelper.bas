Attribute VB_Name = "PdfHelper"
Private Sub ConvertAllDocxToPDF_CAUTION()
  Dim sRootFolder As String
  
  ' Set your root folder here
  'sRootFolder = "C:\Users\craig\Calibre Library\Various\"
  
  ProcessFolder sRootFolder
  
  MsgBox "Done! All DOCX files converted to PDF."
End Sub

Private Sub ProcessFolder(sFolderPath As String)
  Dim oDoc As Document
  Dim sFile As String
  Dim sPDFPath As String
  Dim oSubFolders As Object
  Dim oFolder As Object
  Dim oFSO As Object
  
  Set oFSO = CreateObject("Scripting.FileSystemObject")
  
  If Not oFSO.FolderExists(sFolderPath) Then Exit Sub
  
  ' Process all DOCX files in this folder
  sFile = Dir(sFolderPath & "*Notes*.docx")
  Do While sFile <> ""
    Dim sFullPath As String
    sFullPath = sFolderPath & sFile
    sPDFPath = Left(sFullPath, Len(sFullPath) - 5) & ".pdf"
    
    ' Open the document
    Set oDoc = Documents.Open(sFullPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
    
    ' Export as PDF with heading bookmarks
    oDoc.ExportAsFixedFormat OutputFileName:=sPDFPath, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, CreateBookmarks:=wdExportCreateHeadingBookmarks, DocStructureTags:=True
    
    oDoc.Close SaveChanges:=False
    
    sFile = Dir()
  Loop
  
  ' Recurse through subfolders
  Set oFolder = oFSO.GetFolder(sFolderPath)
  For Each oSubFolders In oFolder.SubFolders
      ProcessFolder oSubFolders.Path & "\"
  Next oSubFolders
  
  Set oFSO = Nothing
End Sub

