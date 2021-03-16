Sub BuscarDadosNoBanco()
Application.ScreenUpdating = False

    Dim conexao As New ADODB.Connection
    Dim rec     As New ADODB.Recordset
    Dim banco   As String
    Dim msg     As String
    Dim linha   As Long
                
    banco = "G:\Programa Excel_Games\VBA com ACCESS\bd_DataBase.mdb"
    
        With conexao
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .CursorLocation = adUseClient
            .ConnectionString = banco
            .Open
        End With
        
    rec.Open "select * from produtos ", conexao
    
    rec.MoveFirst
    linha = 2
        Do
            Planilha1.Cells(linha, 1).Value = rec("codigo_produto")
            Planilha1.Cells(linha, 2).Value = rec("descrição")
            Planilha1.Cells(linha, 3).Value = rec("valor_produto")
            
        linha = linha + 1
                                       
        rec.MoveNext
        
        Loop Until rec.EOF 'E.O.F = End of Files (Fim dos Arquivos)
            If rec.EOF = True And _
                Planilha1.Cells(linha, 1).Value < 0 And _
                Planilha1.Cells(linha, 1).Value = "" Then
                    Planilha1.Cells(linha, 1).Value = " - "
                    Planilha1.Cells(linha, 2).Value = " - "
                    Planilha1.Cells(linha, 3).Value = " - "
            End If
            
                        ActiveWorkbook.Worksheets("Planilha de Conexão").Sort.SortFields.Clear
                            ActiveWorkbook.Worksheets("Planilha de Conexão").Sort.SortFields.Add2 _
                                Key:=Range("a:a"), _
                                SortOn:=xlSortOnValues, _
                                Order:=xlAscending, _
                                DataOption:=xlSortNormal
                                    With ActiveWorkbook.Worksheets("Planilha de Conexão").Sort
                                        .SetRange Range("A:C")
                                        .Header = xlYes
                                        .MatchCase = False
                                        .Orientation = xlTopToBottom
                                        .SortMethod = xlPinYin
                                        .Apply
                                    End With
    rec.Close
    conexao.Close
    Application.ScreenUpdating = True
End Sub