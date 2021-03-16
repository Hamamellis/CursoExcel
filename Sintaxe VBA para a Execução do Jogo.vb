
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

Dim lin As Integer, col As Integer

Dim linSup As Integer, linInf As Integer
Dim colEsq As Integer, colDir As Integer


linSup = Range("jogo").Cells(1).Row
linInf = Range("jogo").Cells(16).Row
colEsq = Range("jogo").Cells(1).Column
colDir = Range("jogo").Cells(16).Column

    If Not Intersect(Target, Range("jogo")) Is Nothing Then
    
    lin = Target.Row
    col = Target.Column
            
        If Cells(lin - 1, col) = "" And lin > linSup Then
            Cells(lin - 1, col) = Cells(lin, col)
            Cells(lin, col) = ""
        End If
        
        If Cells(lin + 1, col) = "" And lin < linInf Then
            Cells(lin + 1, col) = Cells(lin, col)
            Cells(lin, col) = ""
        End If
        
        If Cells(lin, col - 1) = "" And col > colEsq Then
            Cells(lin, col - 1) = Cells(lin, col)
            Cells(lin, col) = ""
        End If
        
        
        If Cells(lin, col + 1) = "" And col < colDir Then
            Cells(lin, col + 1) = Cells(lin, col)
            Cells(lin, col) = ""
        End If
        
        
        'Diagonal Esquerda Superior - Uma linha Acima e Uma coluna a Esquerda
        If Cells(lin - 1, col - 1) = "" And col > colEsq Then
            Cells(lin - 1, col - 1) = Cells(lin, col)
           Cells(lin, col) = ""
        End If
        
        'Diagonal Direita Superior - Uma linha Acima e Uma coluna a Direita
        If Cells(lin - 1, col + 1) = "" And col < colDir Then
            Cells(lin - 1, col + 1) = Cells(lin, col)
            Cells(lin, col) = ""
        End If
        
        'Diagonal Esquerda Inferior- Uma linha Acima e Uma coluna a Esquerda
        If Cells(lin + 1, col - 1) = "" And col > colEsq Then
            Cells(lin + 1, col - 1) = Cells(lin, col)
           Cells(lin, col) = ""
        End If
    
        'Diagonal Direita Inferior - Uma linha Acima e Uma coluna a Direita
        If Cells(lin + 1, col + 1) = "" And col < colDir Then
            Cells(lin + 1, col + 1) = Cells(lin, col)
            Cells(lin, col) = ""
        End If
        
    End If

End Sub

-------------------------------------------------------------------------------------------------------


Sub Sortear()

Application.ScreenUpdating = False

Dim lista(15) As Integer
Dim n As Integer

Range("jogo").ClearContents

        For a = 1 To 15
        lista(a) = a
        Next

    For a = 1 To 15
        n = Application.WorksheetFunction.RandBetween(a, 15)
        Range("jogo").Cells(a) = lista(n)
        lista(n) = lista(a)
    Next
    
    Application.ScreenUpdating = True

End Sub



