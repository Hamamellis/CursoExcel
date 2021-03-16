
Option Explicit
    'Abaixo vamos declarar as variáveis Globais para a "Sub AnalisarLinha" e "Sub PegarValor"
Dim Valor As Double
Dim Doc As String
Dim Data As Date

Sub AtualizarCompilado()
    
    'Para a conta 1
        'Entrar na ABA conta 1
            'Se a movimentação foi finalizada:
                'Pega os valores das variáveis;
                'Vai pra ABA de Consolidação;
                'Se Existirem dados da Mesma Conta, Mesma Data e Mesmo Tipo;
                'ACRESCENTA
                'Senão
                    'Registra uma Nova Linha;
        'Lopping deste processo até o Fim da conta 1
        
     'Repete o processo na Conta 2
        'Colocar em Ordem crescente de data (Opcional)
         
    Consolidar ("Banco Itaú")
    Consolidar ("Banco do Brasil")
    Consolidar ("Banco Caixa Econômica")

End Sub

Sub Consolidar(nome_aba As String)

Dim range1, cell As Range

Sheets(nome_aba).Activate 'Temos que ativar a ABA que estamos trabalhando...

Set range1 = Range("A2:A300")

    'Vamos definir o intervalo de DADOS
    'Para cada Linha
    For Each cell In range1
        If AnalisarLinha(cell) Then
            RegistraLinha
        End If
    Next
        Sheets("Consolidação de Contas").Select
        Range("A1:H1").Select
        ActiveWorkbook.Save
End Sub

Function AnalisarLinha(cell As Range) As Boolean
    If cell.Offset(0, 5).Value = "Finalizado" Then
    
    Data = cell.Value
    Doc = cell.Offset(0, 4).Value
    Call PegarValor(cell)  'Função PegarValor... Utilizando o Call para chamar a Sub PegarValor
        AnalisarLinha = True
    Else
        AnalisarLinha = False
    End If
    'MsgBox (cell.Row)
    'MsgBox (AnalisarLinha)
End Function

Sub PegarValor(cell As Range)
    If cell.Offset(0, 2).Value = "Entrada" Then
        Valor = cell.Offset(0, 3).Value
        ElseIf cell.Offset(0, 2).Value = "Saída" Then
            Valor = -cell.Offset(0, 3).Value '-cell.Offset... Se negativo
            Else
            Valor = 0
                cell.Offset(0, 8).Value = "Não Compilado"
        End If
End Sub

'Registrar na Aba de Consolidação de Contas as Informações
'Logo após, retornar para a ABA da Conta
Sub RegistraLinha()

    Dim nome_aba_conta, nome_aba_consolidacao As String
    Dim range_consolidado, cell As Range
    
    nome_aba_conta = ActiveSheet.Name
    nome_aba_consolidacao = "Consolidação de Contas"
    
    Sheets(nome_aba_consolidacao).Activate
    
    'Atribuindo o range consolidado para receber o Intervalo até o final... é o Mesmo que Ctrl + Seta para Baixo + 1 linha
    Set range_consolidado = Range("A1", Range("A1").End(xlDown)).Offset(2, 0)
        
        For Each cell In range_consolidado
        If cell.Value = "" Then
            cell.Value = Data
            cell.Offset(0, 5).Value = Doc
            cell.Offset(0, 4).Value = nome_aba_conta
            
            If Valor < 0 Then    'Se o valor da celula for menor que "Zero"
                    cell.Offset(0, 3).Value = Valor 'A valor da célula da linha 0 e na Coluna 3 vai receber o Valor Negativo
                    cell.Offset(0, 1).Value = "Saída" 'A valor da célula da linha 0 e na Coluna 1 vai receber Saída
                    
                    Selection.NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
                    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                    With Selection.FormatConditions(1).Font
                        .Color = -16776961
                    End With
                    Selection.FormatConditions(1).StopIfTrue = False
                    
                ElseIf Valor >= 0 Then ' Senão se o Valor for positivo...
                    cell.Offset(0, 2).Value = Valor
                    cell.Offset(0, 1).Value = "Entrada"
                    
                    Selection.NumberFormat = "$ #,##0.00"
                    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
                    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                    With Selection.FormatConditions(1).Font
                        .TintAndShade = -0.249946592608417
                    End With
                    Selection.FormatConditions(1).StopIfTrue = False
            End If
            Exit For
            
        End If
        
        If cell.Value = Data Then ' se as datas forem iguais...
            If cell.Offset(0, 1).Value = "Entrada" And Valor >= 0 And cell.Offset(0, 4).Value = nome_aba_conta Then
               
                cell.Offset(0, 2).Value = Valor + cell.Offset(0, 2).Value
                cell.Offset(0, 5).Value = cell.Offset(0, 5).Value & "; " & "/ " & Doc
               Exit For
               
               ElseIf cell.Offset(0, 1).Value = "Saída" And Valor < 0 And cell.Offset(0, 4).Value = nome_aba_conta Then
               
                cell.Offset(0, 3).Value = Valor + cell.Offset(0, 3).Value
                cell.Offset(0, 5).Value = cell.Offset(0, 5).Value & "; " & "/" & Doc
               Exit For
               
            End If
        End If
        
        Next
    
        Sheets(nome_aba_conta).Activate
        
End Sub
