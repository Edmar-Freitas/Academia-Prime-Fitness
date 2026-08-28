VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDensidadeOssea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigoAvaliacao As Long
Private pBiestiloideRadioUlnar As Double
Private pBiepicondilianoFemural As Double

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
' BIESTILOIDE RADIO ULNAR
'==================================================

Public Property Get biestiloideRadioUlnar() As Double
    biestiloideRadioUlnar = pBiestiloideRadioUlnar
End Property

Public Property Let biestiloideRadioUlnar(ByVal valor As Double)
    pBiestiloideRadioUlnar = valor
End Property

'==================================================
' BIEPICONDILIANO FEMURAL
'==================================================

Public Property Get biepicondilianoFemural() As Double
    biepicondilianoFemural = pBiepicondilianoFemural
End Property

Public Property Let biepicondilianoFemural(ByVal valor As Double)
    pBiepicondilianoFemural = valor
End Property
