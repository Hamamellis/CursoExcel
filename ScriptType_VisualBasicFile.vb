
Private Sub FormularioRastreamento_Activate()
    FormularioRastreamento.Show
End Sub

Private Sub FormularioRastreamento_Initialize()
    FormularioRastreamento.Show
End Sub

Private Sub CadFormulario_Click()
    
    Dim Objeto As Control

    'Cadastrar as informações dos Motorista em cada Linha da planilha...
        Call Cadastrar
    'Fechar(Esconder) o Formulário...
        FormularioRastreamento.Hide
    'Limpar o Formulário...
    
        For Each Objeto In FormularioRastreamento.Controls
            On Error Resume Next
                ' usamos on Error, pois pedimos para colocar como vazio todos os objeto de Controls,
                ' porém as labels não podem ter o nome alterado para vazio... então o código faz ignorar...
            Objeto.Value = ""
        Next
    
End Sub

-------------------------------------------------------------------------------------------------------------------------------

Sub Cadastrar()
    Dim range1, cell As Range

    ' Escolher o local do Cadastro
    ' Colocar informação na Coluna respectiva...
        If Range("A2").Value = "" Then
            Set range1 = Range("A2")
            Else
                Set range1 = Range("A1").End(xlDown).Offset(1, 0)
        End If
        range1.Value = FormularioRastreamento.MultiPage1.Pages(0).TextBox1.Value
        range1.Offset(0, 1).Value = FormularioRastreamento.MultiPage1.Pages(0).TextBox45.Value
        range1.Offset(0, 2).Value = FormularioRastreamento.MultiPage1.Pages(0).TextBox29.Value
        range1.Offset(0, 3).Value = FormularioRastreamento.MultiPage1.Pages(0).TextBox2.Value
        range1.Offset(0, 4).Value = FormularioRastreamento.MultiPage1.Pages(0).TextBox3.Value
        range1.Offset(0, 5).Value = FormularioRastreamento.MultiPage1.Pages(0).TextBox4.Value
        
        If FormularioRastreamento.MultiPage1.Pages(0).OptionButton1.Value = True Then
            range1.Offset(0, 6).Value = "FROTA"
            Else
            range1.Offset(0, 6).Value = "AGREGADO"
        End If
        'Abaixo temos: TextBox14 = Destino 1º - TextBox48 = Estado - TextBox24 = Peso - TextBox19 Tipo
        range1.Offset(0, 7).Value = _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox14.Value & " (" & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox48.Value & ") " & " Peso: " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox24.Value & " " & " " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox19.Value & "; " & " - " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox15.Value & " (" & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox49.Value & ") " & " Peso: " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox25.Value & " " & " " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox20.Value & "; " & " - " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox16.Value & " (" & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox50.Value & ") " & " Peso: " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox26.Value & " " & " " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox21.Value & "; " & " - " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox17.Value & " (" & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox51.Value & ") " & " Peso: " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox27.Value & " " & " " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox22.Value & "; " & " - " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox18.Value & " (" & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox52.Value & ") " & " Peso: " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox28.Value & " " & " " & _
        FormularioRastreamento.MultiPage1.Pages(2).TextBox23.Value & "; " & " - " & " + Cubagem: " & _
        FormularioRastreamento.MultiPage1.Pages(1).TextBox13.Value & " " & _
        FormularioRastreamento.MultiPage1.Pages(1).TextBox53.Value
        range1.Offset(0, 7).End(xlToRight).Columns.EntireColumn.AutoFit 'AO CONCATENAR, VAI AJUSTAR A LARGURA DA COLUNA
        
                
        range1.Offset(0, 8).Value = FormularioRastreamento.MultiPage1.Pages(1).TextBox6.Value
        range1.Offset(0, 9).Value = FormularioRastreamento.MultiPage1.Pages(1).TextBox7.Value
        range1.Offset(0, 10).Value = FormularioRastreamento.MultiPage1.Pages(1).TextBox9.Value
        range1.Offset(0, 11).Value = FormularioRastreamento.MultiPage1.Pages(1).TextBox10.Value
        range1.Offset(0, 12).Value = FormularioRastreamento.MultiPage1.Pages(1).TextBox11.Value
        range1.Offset(0, 13).Value = FormularioRastreamento.MultiPage1.Pages(1).TextBox8.Value
        
        range1.Offset(0, 14).Value = FormularioRastreamento.MultiPage1.Pages(1).TextBox12.Value
        range1.Offset(0, 15).Value = FormularioRastreamento.MultiPage1.Pages(1).ListBox1.Value

End Sub

-------------------------------------------------------------------------------------------------------------------------------

Private Sub CommandButton3_Click()

Dim varTexto As String
'Dim NameNewPlan As String
Dim resultado As VbMsgBoxResult

        Sheets("NovaPlanilha").Visible = True
        Sheets("NovaPlanilha").Select
        Sheets("NovaPlanilha").Copy After:=Sheets(10)
        Sheets("NovaPlanilha").Visible = False
        On Error Resume Next
        Sheets("NovaPlanilha (2)").Name = "RENOMEAR"
    
        resultado = MsgBox("VOCÊ CRIOU UMA NOVA GUIA NA PLANILHA!!!", vbYesNo, "NOVA PLANILHA")
     
     If resultado = vbYes Then
        FormularioRastreamento.MultiPage1.Pages(4).TextBox54.Value = _
        " Você criou uma nova planilha, - Por favor Renomeá-la com a Data do período!!!"
        Else
             MsgBox "Você escolheu 'NÂO' - A Planilha será Deletada"
             ActiveSheet.Delete
             FormularioRastreamento.MultiPage1.Pages(4).TextBox54.Value = " Você escolheu deletar a Nova Planilha!!!"
        End If
End Sub

-------------------------------------------------------------------------------------------------------------------------------

Private Sub CommandButton4_Click()
    FormularioRastreamento.Hide
End Sub

Private Sub PosicaoAtual_Click()
    Dim Objeto1 As Control
    
        Call Cadastrar1
    'Fechar(Esconder) o Formulário...
        'Se desejar use o FormularioRastreamento.Hide para esconder
    'Abaixo vamos Limpar todos os Objetosde Controle do Formulário...
    
        For Each Objeto1 In FormularioRastreamento.MultiPage1.Pages(3).Controls
            On Error Resume Next
                ' usamos on Error, pois pedimos para colocar como vazio todos os objeto de Controls,
                ' porém as labels não podem ter o nome alterado para vazio... então o código faz ignorar...
            Objeto1.Value = ""
        Next
End Sub

-------------------------------------------------------------------------------------------------------------------------------

Sub Cadastrar1()
    Dim range2, cell2 As Range
        
    Set range2 = Range("A2:A50")
        
    For Each cell2 In range2
            If cell2.Offset(0, 3).Value = FormularioRastreamento.MultiPage1.Pages(3).TextBox46.Value And _
            FormularioRastreamento.MultiPage1.Pages(3).OptionButton4 = True Then 'Se OptionButton4 NÃO estiver acionado!
                cell2.End(xlToRight).Offset(0, 1).Value = _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox33.Value & " (" & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox34.Value & ") (" & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox35.Value & " " & " Km de " & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox40.Value & ") " & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox36.Value & " (" & " Km de " & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox41.Value & ") (" & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox37.Value & " " & " Km de " & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox42.Value & ") " & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox38.Value & " (" & " Km de " & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox43.Value & ") (" & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox39.Value & " " & " Km de " & _
                    FormularioRastreamento.MultiPage1.Pages(3).TextBox44.Value & ") "
                cell2.End(xlToRight).Columns.EntireColumn.AutoFit 'AO CONCATENAR, VAI AJUSTAR A LARGURA DA COLUNA
            End If
            
            If cell2.Offset(0, 3).Value = FormularioRastreamento.MultiPage1.Pages(3).TextBox46.Value And _
                    FormularioRastreamento.MultiPage1.Pages(3).OptionButton3 = True Then
                    
                        cell2.End(xlToRight).Offset(0, 1).Value = _
                            FormularioRastreamento.MultiPage1.Pages(3).TextBox33.Value & " " & _
                            FormularioRastreamento.MultiPage1.Pages(3).TextBox34.Value & " - " & "FINALIZADO!"
                        cell2.End(xlToRight).Columns.EntireColumn.AutoFit
            End If
Next
End Sub

-------------------------------------------------------------------------------------------------------------------------------

'1ª FORMA DE ENVIAR E-MAIL

Sub ENVIAR_EMAIL_ADD_PLANILHA()

Dim MyOlapp, MeuItem As Object

Set MyOlapp = CreateObject("Outlook.Application")
Set MeuItem = MyOlapp.CreateItem(olMailItem)

With MeuItem

    .To = ("exemplo@email.com;exemplo@email.com")
    .Subject = "RELATORIO: PAGAMENTOS DE JANEIRO/2020"
    .Body = "Bom dia Sr." & Plan1.[d1].Value & vbCrLf & _
           "Anexo estamos lhe enviando a planilha Relatório" & vbCrLf & _
           "Janeiro/2020 " & _
           "Saudações " & vbCrLf & _
           Plan1.[D2].Value
           
     'troque o diretorio do documento que queira enviar 'add' anexo.
    .Attachments.Add "C:\SaberExcel\decompor_treina_email.xlsm"
    .Display
End With
End Sub

'2ª FORMA DE ENVIAR E-MAIL

Sub Email()

  Dim rng As Range
  Dim OutApp As Object
  Dim OutMail As Object
  
  Para = "wagnerhamamellis@gmail.com"
  File = "C:\Temp\Excel.xlsm"
  
  Set rng = Nothing
    On Error Resume Next
    Set rng = Range("A1:B3").SpecialCells(xlCellTypeVisible)
  
  On Error GoTo 0
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
  
  On Error Resume Next
    With OutMail
      .To = Para
      .Subject = "Assunto"
      .HTMLBody = RangetoHTML(rng)
      .Attachments.Add File
      .Display
    End With
  
  On Error GoTo 0
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
  Set OutMail = Nothing
  Set OutApp = Nothing
  
End Sub

'3ª FORMA DE ENVIAR E-MAIL

Sub MandaEmail()
    
    Dim EnviarPara As String
    Dim Mensagem As String
    For I = 1 To 10
        EnviarPara = ThisWorkbook.Sheets(1).Cells(I, 1)
        If EnviarPara <> "" Then
            Mensagem = ThisWorkbook.Sheets(1).Cells(I, 3)
            Envia_Emails EnviarPara, Mensagem
        End If
    Next I
End Sub

Sub Envia_Emails(EnviarPara As String, Mensagem As String)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    With OutlookMail
        .To = EnviarPara
        .CC = ""
        .BCC = ""
        .Subject = "Pedido enviado"
        .Body = Mensagem
        .Display ' para envia o email diretamente defina o código  .Send
    End With
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub

-------------------------------------------------------------------------------------------------------------------------------



