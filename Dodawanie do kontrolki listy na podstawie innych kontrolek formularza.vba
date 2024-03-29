    ' Pobierz wartości z TextBoxów (lub innych kontrolek) na formularzu
    Dim iloscSztuk As String
    Dim kodWady As String
    
      iloscSztuk = Me.TbIloscScrap.Value
      kodWady = Me.TbkodWady.Value
  
      ' Dodaj nowy wiersz do ListBoxa
      Me.ListaScrap.AddItem iloscSztuk & vbTab & kodWady
  
      ' Wyczyść TextBoxy po dodaniu do ListBoxa
      Me.TbIloscScrap.Value = ""
      CbKodWady.Value = ""

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Dodawanie rekordów z listu do akrusza excel

    Dim i As Integer
    Dim Material As String    'dodatkowe wyliczane wartości dostosować lub usunąć
    Dim Wartosc As Double    'dodatkowe wyliczane wartości dostosować lub usunąć
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TabelaScrap") ' Zmień "NazwaTwojegoArkusza" na właściwą nazwę arkusza
    i = ws.Range("A1").CurrentRegion.Rows.Count
    Material = WorksheetFunction.VLookup(CbListaMat, Sheets("Lista materialow").Range("$A$2:$B$200"), 2, 0)
    On Error Resume Next
    
    Wartosc = WorksheetFunction.VLookup(Material, Sheets("Lista materialow").Range("$B$2:$C$200"), 2, 0)
    
    On Error GoTo 0
    
    If IsEmpty(Wartosc) Then
        MsgBox "Wystąpił błąd podczas pobierania wartości scrap poinformuj o tym administratora!"
        Wartosc = 0
    End If
    
    ' Przejdź przez elementy w ListBoxie i zapisz do arkusza
    For il = 0 To Me.ListaScrap.ListCount - 1
        ws.Cells(i + il + 1, 1).Value = i 
        ws.Cells(i + il + 1, 2).Value = Now()
        ws.Cells(i + il + 1, 3).Value = Material

    Next il

    ' Wyczyść ListBox po zapisaniu danych
    Me.ListaScrap.Clear
