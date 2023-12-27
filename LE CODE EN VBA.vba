Dim numero1 As Double
Dim numero2 As Double
Dim operation As String
Dim operateur As Boolean

Private Sub btn_percent_Click()
    numero2 = Val(lbl_resultat.Caption)
    If (operation = "*" And operateur = True) Then
        If (numero2 = 0) Then
            lbl_resultat.Caption = "1"
            numero2 = 1
        Else
            lbl_resultat.Caption = (numero2 / 100)
            numero2 = (numero2 / 100)
        End If
    End If
End Sub

Private Sub btn_egale_Click()
    numero2 = Val(lbl_resultat.Caption)
    If (operateur = True) Then
        Select Case operation
            Case "+"
                lbl_resultat.Caption = numero1 + numero2
            Case "-"
                lbl_resultat.Caption = numero1 - numero2
            Case "*"
                lbl_resultat.Caption = numero1 * numero2
            Case "/"
                If numero2 <> 0 Then
                    lbl_resultat.Caption = numero1 / numero2
                Else
                    lbl_resultat.Caption = "Désolé... Nous ne pouvons pas diviser par zéro"
                End If
        End Select
    End If
End Sub

Private Sub btn_0_Click()
    If (lbl_resultat.Caption <> "0") Then
        lbl_resultat.Caption = lbl_resultat.Caption & "0"
    End If
End Sub

Private Sub btn_c_Click()
    lbl_resultat.Caption = "0"
    operateur = False
End Sub

Private Sub btn_moin_Click()
    operationn ("-")
End Sub

Private Sub btn_plus_Click()
    operationn ("+")
End Sub
Private Sub btn_mult_Click()
    operationn ("*")
End Sub

Private Sub btn_div_Click()
    operationn ("/")
End Sub

Private Sub btn_cinq_Click()
    numero ("5")
End Sub

Private Sub btn_deux_Click()
    numero ("2")
End Sub

Private Sub btn_huit_Click()
    numero ("8")
End Sub

Private Sub btn_neuf_Click()
    numero ("9")
End Sub

Private Sub btn_quatre_Click()
    numero ("4")
End Sub

Private Sub btn_sept_Click()
    numero ("7")
End Sub

Private Sub btn_six_Click()
    numero ("6")
End Sub

Private Sub btn_trois_Click()
    numero ("3")
End Sub

Private Sub btn_un_Click()
    numero ("1")
End Sub

Private Sub operationn(op As String)
    numero1 = Val(lbl_resultat.Caption)
    lbl_resultat.Caption = "0"
    operateur = True
    operation = op
End Sub

Private Sub numero(num As String)
    If (lbl_resultat.Caption = "0") Then
        lbl_resultat.Caption = num
    Else
        lbl_resultat.Caption = lbl_resultat.Caption & num
    End If
End Sub

