'Códigos VBA para o Excel.
'
'
Sub NovaLinha_Clique()

'Torna ativa a planilha que tem que estar ativa.
    Worksheets("Planilha1").Activate

'Inserir nova linha no final da planilha
    Range("Tabela1").Select
    Selection.End(xlDown).Select
    Selection.EntireRow.Insert
    
'Ajustar linhas automaticamente.
    Cells.Select
    Cells.EntireRow.AutoFit

'Salvar o arquivo.
    ActiveWorkbook.Save


End Sub