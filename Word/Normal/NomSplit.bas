Attribute VB_Name = "NomSplit"
Sub SplitMagazineByMonth()
    Dim srcDoc As Document
    Dim newDoc As Document
    Dim para As Paragraph
    Dim monthCount As Integer
    Dim splitPoints() As Long
    Dim splitCount As Integer
    Dim i As Long, j As Long
    Dim startPara As Long, endPara As Long
    Dim totalParas As Long
    Dim baseName As String
    Dim savePath As String
    
    Set srcDoc = ActiveDocument
    totalParas = srcDoc.Paragraphs.count
    
    ' Get save path from source document location
    savePath = srcDoc.Path & "\"
    baseName = Left(srcDoc.Name, InStrRev(srcDoc.Name, ".") - 1)
    
    ' Check template exists
    Dim templatePath As String
    templatePath = savePath & "..\" & "NOM Template.docx"
    If Dir(templatePath) = "" Then
        MsgBox "Template file not found: " & vbCrLf & templatePath, vbExclamation
        Exit Sub
    End If
    
    ' --- Pass 1: Find all split points (Heading 1 paragraphs containing a dash) ---
    splitCount = 0
    ReDim splitPoints(1 To 50)
    
    For i = 1 To totalParas
        Dim p As Paragraph
        Set p = srcDoc.Paragraphs(i)
        If p.Style = srcDoc.Styles("Heading 1") Then
            Dim txt As String
            txt = p.Range.Text
            If InStr(txt, ChrW(8212)) > 0 Then
                splitCount = splitCount + 1
                splitPoints(splitCount) = i
            End If
        End If
    Next i
    
    If splitCount = 0 Then
      MsgBox "No month-separator Heading 1 paragraphs found. Make sure they contain an em-dash (—).", vbExclamation, "Error"
      Exit Sub
    End If
    
    If splitCount <> 12 Then
      If MsgBox(splitCount & " month headings found (expected 12). Continue?", vbExclamation + vbYesNo, "Warning") = vbNo Then Exit Sub
    End If
    
    ' --- Pass 2: Extract each month ---
    Application.ScreenUpdating = False
    
    For i = 1 To splitCount
      ' Determine paragraph range for this month
      startPara = splitPoints(i)
      If i < splitCount Then
          endPara = splitPoints(i + 1) - 1
      Else
          endPara = totalParas
      End If
      
      ' Select the range for this month and copy it
      Dim rng As Range
      Set rng = srcDoc.Range( _
          srcDoc.Paragraphs(startPara + 1).Range.Start, _
          srcDoc.Paragraphs(endPara).Range.End)
      rng.Copy
      
      ' Build filename from heading text
      Dim headingText As String
      headingText = srcDoc.Paragraphs(startPara).Range.Text
      headingText = Trim(Replace(headingText, Chr(13), " "))
      headingText = Trim(Replace(headingText, Chr(10), " "))
      headingText = Trim(Replace(headingText, ChrW(8212), " - "))
      headingText = Trim(Replace(headingText, ChrW(8211), " - "))
      headingText = Trim(Replace(headingText, "NOTES OF MINISTRY", " "))
      headingText = Trim(Replace(headingText, "No.", " "))
      
      Do While InStr(headingText, "  ") > 0
        headingText = Replace(headingText, "  ", " ")
      Loop
      headingText = "Notes of Ministry - " & headingText
      
      Dim fileName As String
      fileName = savePath & headingText & ".docx"
      
      ' Copy template to new filename, then open it
      FileCopy templatePath, fileName
      Set newDoc = Documents.Open(fileName, , False, False)
      
      ' Move to end of document and paste
      newDoc.Range(newDoc.Content.End - 1, newDoc.Content.End - 1).Select
      'Selection.PasteAndFormat (wdFormatOriginalFormatting)
      Selection.Paste
      
      ' Remove trailing empty paragraph if present
      With newDoc
          Dim lastP As Paragraph
          Set lastP = .Paragraphs(.Paragraphs.count)
          If Len(Trim(lastP.Range.Text)) <= 1 And .Paragraphs.count > 1 Then
              lastP.Range.Delete
          End If
      End With
      
      ' Set document title to match filename (without path or extension)
      newDoc.BuiltInDocumentProperties(wdPropertyTitle).Value = headingText
      
      newDoc.Save
      'newDoc.ExportAsFixedFormat savePath & headingText & ".pdf", wdExportFormatPDF, OpenAfterExport:=False, _
      '      OptimizeFor:=wdExportOptimizeForPrint, CreateBookmarks:=wdExportCreateHeadingBookmarks
      newDoc.Close False
    Next i
    
    Application.ScreenUpdating = True
    MsgBox splitCount & " files saved to:" & vbCrLf & savePath, vbInformation
End Sub

