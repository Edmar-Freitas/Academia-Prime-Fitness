VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrincipal 
   Caption         =   "Prime Fitness - MENU PRINCIPAL"
   ClientHeight    =   9204
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13332
   OleObjectBlob   =   "frmPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAvaliacaoFisica_Click()
    frmAvaliacao.Show
End Sub

Private Sub cmdCadastroAlunos_Click()
    frmAluno.Show
End Sub

Private Sub cmdCadastroFuncionarios_Click()
    frmFuncionario.Show
End Sub

Private Sub cmdSair_Click()

    EncerrarSessao

    Unload Me

    frmAutenticacao.txtLogin.Value = vbNullString
    frmAutenticacao.txtSenha.Value = vbNullString

    frmAutenticacao.Show

End Sub

Private Sub CommandButton3_Click()
    frmAluno.Show

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, _
                                CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then

        EncerrarSessao

    End If

End Sub

