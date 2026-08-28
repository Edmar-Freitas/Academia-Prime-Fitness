VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendario 
   Caption         =   "Academia Prime Fitness - CALENDÁRIO"
   ClientHeight    =   4716
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4356
   OleObjectBlob   =   "frmCalendario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Private ColecaoBotoes As Collection
    Private DataAtual As Date
    Private MesAtual As Integer
    Private AnoAtual As Integer

Private Sub cmdAnterior_Click()
    MesAtual = MesAtual - 1

    If MesAtual = 0 Then
        MesAtual = 12
        AnoAtual = AnoAtual - 1
    End If

    MontarCalendario
End Sub



Private Sub cmdProximo_Click()
    MesAtual = MesAtual + 1

    If MesAtual = 13 Then
        MesAtual = 1
        AnoAtual = AnoAtual + 1
    End If

    MontarCalendario
End Sub

Private Sub lblMesAno_Click()

End Sub

Private Sub spbtAno_Change()
    AnoAtual = spbtAno.Value
    lblMesAno.Caption = _
        Format(DateSerial(AnoAtual, MesAtual, 1), "mmmm yyyy")
End Sub


Private Sub UserForm_Initialize()
    DataAtual = Date

    MesAtual = Month(DataAtual)
    AnoAtual = Year(DataAtual)
    spbtAno.Value = Year(DataAtual)
    
    Dim i As Integer

    Dim Evento As clsBotaoCalendario

    Set ColecaoBotoes = New Collection

    For i = 1 To 42

        Set Evento = New clsBotaoCalendario

        Set Evento.Botao = Me.Controls("cmdDia" & i)

        Set Evento.Formulario = Me

        ColecaoBotoes.Add Evento

    Next i

    MontarCalendario

End Sub


Public Sub MontarCalendario()

    Dim PrimeiroDia As Date
    Dim PrimeiroDiaSemana As Integer
    Dim UltimoDia As Integer
    Dim Dia As Integer
    Dim Posicao As Integer
    Dim i As Integer

    PrimeiroDia = DateSerial(AnoAtual, MesAtual, 1)

    PrimeiroDiaSemana = Weekday(PrimeiroDia)

    UltimoDia = Day(DateSerial(AnoAtual, MesAtual + 1, 0))

    'Limpa os 42 botões
    For i = 1 To 42

        With Controls("cmdDia" & i)
            .Caption = ""
            .Visible = False
            .Enabled = False
        End With

    Next i

    Posicao = PrimeiroDiaSemana

    For Dia = 1 To UltimoDia

        With Controls("cmdDia" & Posicao)

            .Caption = CStr(Dia)
            .Visible = True
            .Enabled = True

        End With

        Posicao = Posicao + 1

    Next Dia

    lblMesAno.Caption = _
        Format(DateSerial(AnoAtual, MesAtual, 1), "mmmm yyyy")

End Sub

Public Sub DiaSelecionado(ByVal Dia As Integer)

    MsgBox Dia, , "Dia Escolhido"
    
    ControleDestino.Value = _
    Format(DateSerial(AnoAtual, MesAtual, Dia), "dd/mm/yyyy")

    Unload Me

End Sub

