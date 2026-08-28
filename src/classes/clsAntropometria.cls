VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAntropometria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigoAvaliacao As Long
Private pPeso As Double
Private pAltura As Double

'==================================================
' CODIGO DA AVALIACAO
'==================================================

Public Property Get codigoAvaliacao() As Long
    codigoAvaliacao = pCodigoAvaliacao
End Property

Public Property Let codigoAvaliacao(ByVal valor As Long)
    pCodigoAvaliacao = valor
End Property

'==================================================
' PESO
'==================================================

Public Property Get peso() As Double
    peso = pPeso
End Property

Public Property Let peso(ByVal valor As Double)
    pPeso = valor
End Property

'==================================================
' ALTURA
'==================================================

Public Property Get altura() As Double
    altura = pAltura
End Property

Public Property Let altura(ByVal valor As Double)
    pAltura = valor
End Property
