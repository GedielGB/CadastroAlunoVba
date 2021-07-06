VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CadastroAluno 
   Caption         =   "CadastroAluno"
   ClientHeight    =   6910
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   13110
   OleObjectBlob   =   "CadastroAluno.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CadastroAluno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Salvar_Click()

' Rotina de erro para campo em branco

Dim flag As Integer
flag = 0

If (TextBox_nome.Value = "") Or (TextBox_endereco.Value = "") Or (TextBox_idade.Value = "") Or (c) Or (TextBox_comum.Value = "") Then
    MsgBox "O campo NOME não pode ficar vazio", 48, "Campo em branco"
    flag = 1

Else

Worksheets("MODELO").Copy after:=Sheets(Sheets.Count)

ActiveSheet.Name = TextBox_nome

'Insere os dados nos campos da planilha

Cells(4, 4) = TextBox_nome

Cells(6, 5) = TextBox_nome

Cells(7, 5) = TextBox_endereco

Cells(8, 5) = TextBox_idade

Cells(9, 5) = TextBox_comum

End If

End Sub

Private Sub CommandButton2_Click()

' Fecha o formulário CadastroAluno

Unload CadastroAluno

End Sub

Private Sub Label7_Click()

End Sub

Private Sub TextBox_horini_Change()

End Sub


Private Sub UserForm_Click()

End Sub
