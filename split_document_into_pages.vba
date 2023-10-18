Sub PodzielNaStrony()
    Dim docOriginal As Document
    Dim docNew As Document
    Dim rngPage As Range
    Dim strPath As String
    Dim intPageCount As Integer
    Dim intCurrentPage As Integer

    Set docOriginal = ActiveDocument
    intPageCount = docOriginal.ComputeStatistics(wdStatisticPages)

    ' Ścieżka do zapisywania dokumentów
    strPath = docOriginal.Path & "\"
    
    ' Iteracja przez wszystkie strony
    For intCurrentPage = 1 To intPageCount
        Set rngPage = docOriginal.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=intCurrentPage)
        Set rngPage = rngPage.GoTo(What:=wdGoToBookmark, Name:="\page")
        
        Set docNew = Documents.Add
        docNew.Content.FormattedText = rngPage.FormattedText
        docNew.SaveAs strPath & "Strona_" & intCurrentPage & ".docx"
        docNew.Close
    Next intCurrentPage
    
    ' Czyszczenie
    Set docOriginal = Nothing
    Set docNew = Nothing
    Set rngPage = Nothing
End Sub
