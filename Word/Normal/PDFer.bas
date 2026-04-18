Attribute VB_Name = "PDFer"
Public Const NomTitle = "NOTES OF MINISTRY"

Sub ConvertCalibreDocxToPDF()
    Dim sDBPath As String
    
    ' Path to your Calibre database
    sDBPath = "C:\Users\craig\Calibre Library\metadata.db"
    
    ProcessCalibreDocx sDBPath
    
    MsgBox "Done! All DOCX files converted to PDF."
End Sub

Sub ProcessCalibreDocx(sDBPath As String)
    Dim oConn As Object
    Dim oRS As Object
    Dim oDoc As Document
    Dim sDocxPath As String
    Dim sPDFPath As String
    Dim sTitle As String
    Dim sSubTitle As String
    Dim sPubDate As String
    
    ' Connect to Calibre SQLite database
    Set oConn = CreateObject("ADODB.Connection")
    oConn.Open "Driver={SQLite3 ODBC Driver};Database=" & sDBPath & ";"
    
    ' Run the query to get DOCX files
    Dim sSQL As String
    sSQL = "SELECT " & _
           "    b.id, " & _
           "    b.title, " & _
           "    b.series_index, " & _
           "    b.pubdate, " & _
           "    'C:\Users\craig\Calibre Library\' || replace(b.path, '/', '\') || '\' || d.name || '.docx' AS docx_full_path, " & _
           "    'C:\Users\craig\Calibre Library\' || replace(b.path, '/', '\') || '\' || d.name || '.pdf' AS pdf_full_path " & _
           "FROM books b " & _
           "JOIN books_series_link bsl ON b.id = bsl.book " & _
           "JOIN series s ON bsl.series = s.id " & _
           "LEFT JOIN books_publishers_link bpl ON b.id = bpl.book " & _
           "LEFT JOIN publishers p ON bpl.publisher = p.id " & _
           "JOIN data d ON b.id = d.book AND d.format = 'DOCX' " & _
           "WHERE s.name = 'Notes of Ministry' " & _
           "AND b.title NOT LIKE '%SPLIT%' " & _
           "AND b.series_index BETWEEN 410 AND 421 " & _
           "AND (p.name NOT LIKE 'Brown%' OR p.name IS NULL) " & _
           "ORDER BY b.series_index"
    
           '"AND b.series_index=1 " & _

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    
    Set oRS = oConn.Execute(sSQL)
    
    ' Loop through results
    Do While Not oRS.EOF
        sDocxPath = oRS.Fields("docx_full_path").Value
        sPDFPath = oRS.Fields("pdf_full_path").Value
        sTitle = oRS.Fields("title").Value
        
        ' Check file exists before opening
        If Dir(sDocxPath) <> "" Then
            'On Error GoTo NextFile
            
            ' Open the document
            Set oDoc = Documents.Open(sDocxPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)
            
            ' Ensure title property is set
            oDoc.BuiltInDocumentProperties(wdPropertyTitle).Value = sTitle
                      
            ' --- Build front page ---
            Dim oRange As Range
            Dim oTOC As TableOfContents
            
            ' Move to start of document
            Set oRange = oDoc.Range(0, 0)
            
            ' Insert page break at start to push existing content down
            oRange.InsertBreak Type:=wdPageBreak
                    
            ' Move back to start
            Set oRange = oDoc.Range(0, 0)
            
            ' Insert NomTitle as Heading 1
            oRange.InsertBefore NomTitle & vbCr
            Set oRange = oDoc.Range(0, Len(NomTitle))
            oRange.Style = oDoc.Styles("Title")
            oRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
            oRange.Font.Bold = True
            
            ' Move to next line and insert title as Subtitle
            sSubTitle = "No. " & CStr(oRS.Fields("series_index").Value) & vbTab & Format(CDate(oRS.Fields("pubdate").Value), "MMMM yyyy")
            Set oRange = oDoc.Range(Len(NomTitle) + 1, Len(NomTitle) + 1)
            oRange.InsertBefore sSubTitle & vbCr
            Set oRange = oDoc.Range(Len(NomTitle) + 1, Len(NomTitle) + 1 + Len(sSubTitle))
            oRange.Style = oDoc.Styles("Subtitle")
            oRange.ParagraphFormat.SpaceAfter = 28
            
            ' Move to after title line and insert TOC
            Dim iTitleEnd As Integer
            iTitleEnd = Len(NomTitle) + 1 + Len(sSubTitle) + 1
            Set oRange = oDoc.Range(iTitleEnd, iTitleEnd)
            
            oDoc.TablesOfContents.Add _
                Range:=oRange, _
                UseHeadingStyles:=True, _
                UpperHeadingLevel:=1, _
                LowerHeadingLevel:=3, _
                IncludePageNumbers:=True, _
                AddedStyles:="", _
                UseHyperlinks:=True, _
                HidePageNumbersInWeb:=False, _
                UseOutlineLevels:=True
            
            ' Update TOC to get correct page numbers
            oDoc.TablesOfContents(1).Update
            
            ' --- End front page ---
                        
            ' --- FOOTER  ---
            Dim oSection As Section
            Dim oFooter As HeaderFooter
            Dim oFooterRange As Range
            
            For Each oSection In oDoc.Sections
                Set oFooter = oSection.Footers(wdHeaderFooterPrimary)
                
                If Len(Trim(oFooter.Range.Text)) <= 1 Then  ' <= 1 accounts for the paragraph mark
                    Set oFooterRange = oFooter.Range
                    oFooterRange.Collapse wdCollapseStart
                    
                    ' Set font
                    oFooterRange.Font.Size = 10
                    
                    ' Insert page number field
                    oFooterRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    oDoc.Fields.Add Range:=oFooterRange, Type:=wdFieldPage
                End If
            Next oSection

            ' Export as PDF with heading bookmarks
            oDoc.ExportAsFixedFormat _
                OutputFileName:=sPDFPath, _
                ExportFormat:=wdExportFormatPDF, _
                OpenAfterExport:=False, _
                OptimizeFor:=wdExportOptimizeForPrint, _
                CreateBookmarks:=wdExportCreateHeadingBookmarks, _
                DocStructureTags:=True, _
                IncludeDocProps:=True
            
            oDoc.Close SaveChanges:=False
            oShell.Run "cmd.exe /c nompdf.cmd """ & sPDFPath & """", 0, True
            
            Debug.Print "Converted: " & sDocxPath
        Else
            Debug.Print "File not found: " & sDocxPath
        End If
        
NextFile:
        On Error GoTo 0
        oRS.MoveNext
    Loop
    
    oRS.Close
    oConn.Close

    Set oShell = Nothing
    Set oRS = Nothing
    Set oConn = Nothing
End Sub

