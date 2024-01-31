    Dim i As Integer
    Dim Material As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tabela") 'Nazwa twojego arkusza
    
    i = ws.Range("A1").CurrentRegion.Rows.Count  'ile wypełnionych wierszy w arkuszu
    
    ws.Cells(i + 1, 1).Value = i 'Numer ID jeśli potrzebny 
    ws.Cells(i + 1, 2).Value = Now()  'Dzisiejsza data
    ws.Cells(i + 1, 3).Value = KontrolkaFormularza  'Nazwa kontrolki w formularzu
