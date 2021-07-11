Private Sub CommandButton1_Click()

' Rotina para exibir erro de campo em branco

Dim flag As Integer
flag = 0

If TextBox_nome.Value = "" Then
    MsgBox "O campo NOME não pode ficar vazio", 48, "Campo em branco"
    flag = 1
End If

If flag = 0 Then

' imprime relatório

Me.Hide

Início.Hide

Sheets(TextBox_nome.Value).Select

Sheets(TextBox_nome.Value).PrintPreview

Range("B3:G500").Select

Me.Show

Início.Show

End If

End Sub