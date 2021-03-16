
Sub ExercicioReforco()

' Variáveis cell e range1 do tipo Range
Dim cell, range1 As Range

' Atribuímos um intervalo onde faremos a busca inicial
Set range1 = Range("A2:A30")

' Selecionamos em qual Guia (Sheet) iremos fazer a busca
Sheets("09.11 à 15.11").Select

    ' Examinaremos em cada cell dentro do range1 - E veremos se atende as condições
         For Each cell In range1
                    
                If cell.Offset(0, 7).Value = "FROTA" Then
                cell.Offset(0, 7).Select
                    With Selection.Font
                    .Color = rgbBlack
                    .Size = 11
                    .Bold = False
                    .Name = "Bookman Old Style"
                    End With
                Else
                    cell.Offset(0, 7).Select
                    With Selection.Font
                    .Color = rgbBlack
                    .Size = 11
                    .Bold = False
                    .Name = "Courier New"
                    End With
                End If
        Next
    Range("A2").Select
        
End Sub

-------------------------------------------------------------------------------------------------------------------------

Sub WorksheetLoop()
'Verfica quantas guias existem vai mostrar o nome de cada uma delas

         Dim Sheets_Count As Integer ' Aqui recebe o numero total de Guias...
         Dim Iteracao As Integer ' Aqui faz a iteração de cada Loop... ** Repetição **

         ' Set Sheets_Count equal to the number of worksheets in the active
         ' workbook.
         ' Aqui atribui à variável a quantidade de guias na pasta de trabalho ativa...
         Sheets_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For Iteracao = 1 To Sheets_Count
                              
            ' Insert your code here.
            ' The following line shows how to reference a sheet within
            ' the loop by displaying the worksheet name in a dialog box.
            MsgBox ActiveWorkbook.Worksheets(Iteracao).Name
            
         Next Iteracao

      End Sub
	  
-------------------------------------------------------------------------------------------------------------------------
-------------------------------------------------------------------------------------------------------------------------
