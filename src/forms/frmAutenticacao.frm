VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAutenticacao 
   Caption         =   "Academia Prime Fitness - AUTENTICAÇÃO"
   ClientHeight    =   4524
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3864
   OleObjectBlob   =   "frmAutenticacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAutenticacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEntrar_Click()

    If Autenticar( _
            Me.txtLogin.Value, _
            Me.txtSenha.Value) Then

        Me.Hide
        frmPrincipal.Show

    Else

        MsgBox "Login ou senha inválidos.", _
               vbExclamation, _
               "Academia Prime Fitness"

        Me.txtSenha.Value = vbNullString
        Me.txtSenha.SetFocus

    End If

End Sub


Private Sub txtSenha_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                             ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then

        KeyCode = 0

        cmdEntrar_Click

    End If

End Sub

Private Sub UserForm_Initialize()

    Me.txtLogin.Value = vbNullString
    Me.txtSenha.Value = vbNullString

    Me.txtSenha.PasswordChar = "*"

    Me.txtLogin.SetFocus

End Sub
