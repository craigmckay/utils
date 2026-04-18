Attribute VB_Name = "NomContents"
Sub BuildCalibreContents()
    Dim sDBPath As String
    
    ' Path to your Calibre database
    sDBPath = "C:\Users\craig\Calibre Library\metadata.db"
    
    ParseMagazineContents sDBPath
    
    MsgBox "Done! All DOCX files scanned and contents built."
End Sub

Private Sub ParseMagazineContents(sDBPath As String)

    Dim oConn       As Object
    Dim oRS         As Object
    Dim oDoc        As Document
    Dim sDocxPath   As String
    Dim sTitle      As String
    Dim sHTML       As String
    Dim sSQL        As String

    ' ---------------------------------------------------------------
    ' Connect to Calibre SQLite database
    ' ---------------------------------------------------------------
    Set oConn = CreateObject("ADODB.Connection")
    oConn.Open "Driver={SQLite3 ODBC Driver};Database=" & sDBPath & ";"

    ' ---------------------------------------------------------------
    ' Ensure ministrycontents table exists, then wipe it for a clean rebuild
    ' ---------------------------------------------------------------
    oConn.Execute "CREATE TABLE IF NOT EXISTS ministrycontents (" & _
                  "  bookid         INTEGER NOT NULL, " & _
                  "  article_number INTEGER NOT NULL, " & _
                  "  title          TEXT, " & _
                  "  author         TEXT, " & _
                  "  place          TEXT, " & _
                  "  date           TEXT, " & _
                  "  PRIMARY KEY (bookid, article_number)" & _
                  ")"
    oConn.Execute "DELETE FROM ministrycontents"

    ' ---------------------------------------------------------------
    ' Fetch every book in the series that has a DOCX file
    ' Adjust the WHERE clause (series name, index range, etc.) as needed
    ' ---------------------------------------------------------------
    sSQL = "SELECT " & _
           "    b.id, " & _
           "    b.title, " & _
           "    b.series_index, " & _
           "    'C:\Users\craig\Calibre Library\' || replace(b.path, '/', '\') || '\' || d.name || '.docx' AS docx_full_path " & _
           "FROM books b " & _
           "JOIN books_series_link bsl ON b.id = bsl.book " & _
           "JOIN series s ON bsl.series = s.id " & _
           "JOIN data d ON b.id = d.book AND d.format = 'DOCX' " & _
           "WHERE s.name = 'Notes of Ministry' " & _
           "AND b.title NOT LIKE '%SPLIT%' " & _
           "AND b.series_index BETWEEN 299 AND 457 " & _
           "ORDER BY b.series_index"

    Set oRS = oConn.Execute(sSQL)

    ' ---------------------------------------------------------------
    ' Loop through each book
    ' ---------------------------------------------------------------
    Do While Not oRS.EOF

        sDocxPath = oRS.Fields("docx_full_path").Value
        sTitle = Trim(oRS.Fields("title").Value)

        If Dir(sDocxPath) <> "" Then

            On Error GoTo NextFile

            Set oDoc = Documents.Open(sDocxPath, ReadOnly:=True, Visible:=False, AddToRecentFiles:=False)

            Dim sBookID As String
            sBookID = CStr(oRS.Fields("id").Value)

            ' Parse the document — returns HTML and populates ministrycontents rows
            sHTML = BuildContentsHTML(oDoc, sTitle, sBookID, oConn)

            oDoc.Close SaveChanges:=False

            ' Write HTML to Calibre comments field
            oConn.Execute "INSERT OR REPLACE INTO comments (book, text) VALUES (" & _
                          sBookID & ", '" & EscapeSQL(sHTML) & "')"

            Debug.Print "Updated: " & sTitle

        Else
            Debug.Print "File not found: " & sDocxPath
        End If

NextFile:
        On Error GoTo 0
        oRS.MoveNext
    Loop

    oRS.Close
    oConn.Close
    Set oRS = Nothing
    Set oConn = Nothing
End Sub


' ===================================================================
' BuildContentsHTML
' Walks the document paragraphs, finds each Heading 1 article block,
' extracts title/author/place/date, writes rows to ministrycontents,
' and returns an HTML string for the Calibre comments field.
' ===================================================================
Private Function BuildContentsHTML(oDoc As Document, sBookTitle As String, _
                                    sBookID As String, oConn As Object) As String

    Dim oParas      As Paragraphs
    Dim oPara       As Paragraph
    Dim i           As Long
    Dim n           As Long

    Dim sParaText   As String
    Dim sStyle      As String

    Dim sArtTitle   As String
    Dim sAuthor     As String
    Dim sPlace      As String
    Dim sDate       As String

    Dim bInArticle  As Boolean
    Dim bInExtract  As Boolean

    Dim iArticle    As Long     ' running article number within this book

    Dim html        As String
    Dim rows        As String

    Set oParas = oDoc.Paragraphs
    n = oParas.count

    html = ""
    rows = ""
    bInArticle = False
    bInExtract = False
    iArticle = 0

    i = 1
    Do While i <= n

        Set oPara = oParas(i)
        sParaText = Trim(oPara.Range.Text)

        ' Strip the trailing paragraph mark Word appends
        If Len(sParaText) > 0 Then
            If Asc(Right(sParaText, 1)) = 13 Or Asc(Right(sParaText, 1)) = 7 Then
                sParaText = Left(sParaText, Len(sParaText) - 1)
            End If
        End If

        sStyle = oPara.Style.NameLocal

        ' -----------------------------------------------------------
        ' Detect Heading 1 — start of a new article (or EXTRACT)
        ' -----------------------------------------------------------
        If InStr(1, sStyle, "Heading 1", vbTextCompare) > 0 Or _
           sStyle = "1" Or _
           UCase(sParaText) = "KEY TO INITIALS" Or _
           UCase(sParaText) = "LIST OF INITIALS" Then

            ' Before starting new article, save the previous one
            If bInArticle And Not bInExtract Then
                iArticle = iArticle + 1
                
                'If sAuthor = "Verse" And Len(sPlace) > 0 Then
                '  sAuthor = sPlace
                '  sPlace = ""
                'End If
                rows = rows & FormatArticleRow(sArtTitle, sAuthor, sPlace, sDate)
                Call InsertArticleRow(oConn, sBookID, iArticle, sArtTitle, sAuthor, sPlace, sDate)
            End If

            ' Reset for new article
            sArtTitle = sParaText
            sAuthor = ""
            sPlace = ""
            sDate = ""
            bInArticle = True

            ' Flag EXTRACT / EXTRACTS sections so we skip them
            bInExtract = (UCase(Trim(sParaText)) = "EXTRACT" Or _
                          UCase(Trim(sParaText)) = "EXTRACTS" Or _
                          UCase(Trim(sParaText)) = "MUSINGS")
                          
            i = i + 1

        ' -----------------------------------------------------------
        ' We are inside an article block — gather author from line 2
        ' -----------------------------------------------------------
        ElseIf bInArticle And Not bInExtract Then

            ' Line immediately after the heading = potential author
            If sAuthor = "" And sPlace = "" And sDate = "" Then

                If Len(sParaText) > 0 Then
                    If InStr(1, sStyle, "Scriptures", vbTextCompare) > 0 Then
                      ' Scriptures no named author
                      sAuthor = "Reading"
                    ElseIf InStr(1, sStyle, "Verse", vbTextCompare) > 0 Then
                      ' Verse - author is probably the next non-verse line
                      sAuthor = "Verse"
                        
                    ElseIf Len(sParaText) < 40 Then ' Longer than this, not an author name
                        sAuthor = sParaText
                    End If
                End If

                i = i + 1

            Else
                ' We are past the author line.
                ' Scan forward to the next heading, picking up the place/date
                ' from whichever paragraph contains a soft return Chr(11).

                Dim iNextHeading As Long
                iNextHeading = 0
                Dim k As Long
                For k = i To n
                    Dim sLookStyle As String
                    sLookStyle = oParas(k).Style.NameLocal
                    Dim sLookText As String
                    sLookText = UCase(Trim(oParas(k).Range.Text))
                    If InStr(1, sLookStyle, "Heading 1", vbTextCompare) > 0 Or sLookStyle = "1" Or _
                       sLookText = "KEY TO INITIALS" Or sLookText = "LIST OF INITIALS" Then
                        iNextHeading = k
                        Exit For
                    End If
                    ' Place/date is the paragraph that contains a soft return
                    If InStr(oParas(k).Range.Text, Chr(11)) > 0 Then
                        Call ExtractPlaceDate(oParas(k), sPlace, sDate)
                    End If
                Next k

                If iNextHeading = 0 Then iNextHeading = n + 1

                ' Jump straight to the next heading
                i = iNextHeading

            End If

        Else
            i = i + 1
        End If

    Loop

    ' Save the final article
    If bInArticle And Not bInExtract Then
        iArticle = iArticle + 1
        rows = rows & FormatArticleRow(sArtTitle, sAuthor, sPlace, sDate)
        Call InsertArticleRow(oConn, sBookID, iArticle, sArtTitle, sAuthor, sPlace, sDate)
    End If

    ' ---------------------------------------------------------------
    ' Wrap rows in a clean HTML table
    ' ---------------------------------------------------------------
    If Len(rows) = 0 Then
        BuildContentsHTML = "<p><em>No articles found.</em></p>"
    Else
        html = "<h2>" & HTMLEncode(sBookTitle) & "</h2>" & vbCrLf
        html = html & "<table border=""0"" cellpadding=""4"" cellspacing=""0"" style=""font-family:Arial,sans-serif;font-size:0.9em;width:100%"">" & vbCrLf
        html = html & "<tr style=""background:#4472C4;color:#fff;"">" & vbCrLf
        html = html & "  <th align=""left"">Title</th>" & vbCrLf
        html = html & "  <th align=""left"">Author</th>" & vbCrLf
        html = html & "  <th align=""left"">Place</th>" & vbCrLf
        html = html & "  <th align=""left"">Date</th>" & vbCrLf
        html = html & "</tr>" & vbCrLf
        html = html & rows
        html = html & "</table>"
        BuildContentsHTML = html
    End If

End Function


' ===================================================================
' InsertArticleRow
' Inserts one parsed article into the ministrycontents table.
' ===================================================================
Private Sub InsertArticleRow(oConn As Object, sBookID As String, iArticle As Long, _
                              sTitle As String, sAuthor As String, _
                              sPlace As String, sDate As String)

    Dim sSQL As String
    sSQL = "INSERT INTO ministrycontents (bookid, article_number, title, author, place, date) VALUES (" & _
           sBookID & ", " & _
           iArticle & ", " & _
           "'" & EscapeSQL(sTitle) & "', " & _
           "'" & EscapeSQL(sAuthor) & "', " & _
           "'" & EscapeSQL(sPlace) & "', " & _
           "'" & EscapeSQL(sDate) & "')"
    oConn.Execute sSQL

End Sub


' ===================================================================
' ExtractPlaceDate
' Splits a paragraph on Chr(11) (soft return) into place and date.
' If no soft return is present the whole text is treated as the date.
' ===================================================================
Private Sub ExtractPlaceDate(oPara As Paragraph, ByRef sPlace As String, ByRef sDate As String)

    Dim sRaw As String
    sRaw = oPara.Range.Text

    ' Strip trailing paragraph/section marks
    Do While Len(sRaw) > 0
        Dim c As Integer
        c = Asc(Right(sRaw, 1))
        If c = 13 Or c = 7 Or c = 11 Then
            sRaw = Left(sRaw, Len(sRaw) - 1)
        Else
            Exit Do
        End If
    Loop

    Dim iSoft As Long
    iSoft = InStr(sRaw, Chr(11))

    If iSoft > 0 Then
        sPlace = Trim(Left(sRaw, iSoft - 1))
        If Right(sPlace, 1) = "," Then sPlace = Left(sPlace, Len(sPlace) - 1) ' Remove any trailing comma
        sDate = Trim(Mid(sRaw, iSoft + 1))
    Else
        ' No soft return — treat the whole thing as the date
        sPlace = ""
        sDate = Trim(sRaw)
    End If

End Sub


' ===================================================================
' IsItalic
' Returns True if the majority of the paragraph's text is italic.
' ===================================================================
Private Function IsItalic(oPara As Paragraph) As Boolean

    Dim oRange  As Range
    Dim oChar   As Range
    Dim nItalic As Long
    Dim nTotal  As Long

    Set oRange = oPara.Range

    Dim iMax As Long
    iMax = oRange.Characters.count
    If iMax > 200 Then iMax = 200

    Dim idx As Long
    For idx = 1 To iMax
        Set oChar = oRange.Characters(idx)
        If oChar.Font.Italic = True Then nItalic = nItalic + 1
        nTotal = nTotal + 1
    Next idx

    If nTotal = 0 Then
        IsItalic = False
    Else
        IsItalic = (nItalic / nTotal) > 0.5
    End If

End Function


' ===================================================================
' FormatArticleRow
' Returns one <tr> HTML row. Alternates row shading.
' ===================================================================
Private Function FormatArticleRow(sTitle As String, sAuthor As String, _
                                   sPlace As String, sDate As String) As String

    Static iRow As Long
    iRow = iRow + 1

    Dim sBg As String
    If iRow Mod 2 = 0 Then
        sBg = "#D9E1F2"
    Else
        sBg = "#FFFFFF"
    End If

    FormatArticleRow = "<tr style=""background:" & sBg & """>" & vbCrLf & _
                       "  <td>" & HTMLEncode(sTitle) & "</td>" & vbCrLf & _
                       "  <td>" & HTMLEncode(sAuthor) & "</td>" & vbCrLf & _
                       "  <td>" & HTMLEncode(sPlace) & "</td>" & vbCrLf & _
                       "  <td>" & HTMLEncode(sDate) & "</td>" & vbCrLf & _
                       "</tr>" & vbCrLf

End Function


' ===================================================================
' HTMLEncode  —  escapes the five XML/HTML special characters
' ===================================================================
Private Function HTMLEncode(s As String) As String
    Dim r As String
    r = s
    r = Replace(r, "&", "&amp;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, """", "&quot;")
    r = Replace(r, "'", "&#39;")
    HTMLEncode = r
End Function


' ===================================================================
' EscapeSQL  —  escapes single quotes for SQL string literals
' ===================================================================
Private Function EscapeSQL(s As String) As String
    EscapeSQL = Replace(s, "'", "''")
End Function


