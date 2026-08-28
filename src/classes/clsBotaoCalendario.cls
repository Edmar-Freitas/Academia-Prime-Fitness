VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBotaoCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents Botao As MSForms.CommandButton
Attribute Botao.VB_VarHelpID = -1

Public Formulario As frmCalendario

Private Sub Botao_Click()

    Formulario.DiaSelecionado CInt(Botao.Caption)

End Sub
