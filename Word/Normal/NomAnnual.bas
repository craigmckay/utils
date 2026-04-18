Attribute VB_Name = "NomAnnual"
Private Const DeleteMe = "DELETEME"
Private Const PubYear = 2007

Public Sub TidyNOM()
  Dim r As Long
  Dim pattern As String
  Dim hits As Long

  RemoveLegacyPageNumbers
  
  For r = 200 To 1 Step -1
    pattern = CStr(r) & " ^p"
    hits = CountOccurrences(pattern)

    If hits = 1 Then
      With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = pattern
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
      End With
    ElseIf hits > 1 Then
      Selection.EndKey Unit:=wdStory
      Selection.TypeText Text:=pattern
      Selection.TypeParagraph
    End If
  Next
  
  RemoveHeaderVolume
End Sub

Private Sub RemoveHeaderVolume()
  Dim dIssue As Date
  Dim sIssueShort As String
  
  dIssue = DateSerial(1972, 12, 1)
  For r = 1 To 457
    If Year(dIssue) = PubYear Then
      sIssueShort = "No. " & CStr(r) & ", " & Format(dIssue, "mmm YYYY")
    
      With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = sIssueShort & " ^p"
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
      
        .Text = sIssueShort
        .Replacement.Text = " "
        .Execute Replace:=wdReplaceAll
      End With
    End If
    
    dIssue = DateAdd("m", 1, dIssue)
  Next

End Sub

Private Function CountOccurrences(pattern As String) As Long
    Dim rng As Range
    Dim count As Long

    Set rng = ActiveDocument.Content
    With rng.Find
        .ClearFormatting
        .Text = pattern
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With

    Do While rng.Find.Execute
        count = count + 1
        rng.Collapse wdCollapseEnd
    Loop

    CountOccurrences = count
End Function

Private Sub RemoveLegacyPageNumbers()
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = "\[[0-9]{1,3}\] "
    .Replacement.Text = DeleteMe
    .Forward = True
    .Wrap = wdFindAsk
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchAllWordForms = False
    .MatchSoundsLike = False
    .MatchWildcards = True
  End With
  Selection.Find.Execute Replace:=wdReplaceAll
  
  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = DeleteMe & "^p"
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindAsk
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
  End With
  Selection.Find.Execute Replace:=wdReplaceAll

  Selection.Find.ClearFormatting
  Selection.Find.Replacement.ClearFormatting
  With Selection.Find
    .Text = DeleteMe
    .Replacement.Text = ""
    .Forward = True
    .Wrap = wdFindAsk
    .Format = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
  End With
  Selection.Find.Execute Replace:=wdReplaceAll
End Sub

