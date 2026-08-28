VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDobrasCultaneas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigoAvaliacao As Long
Private pCodigoProtocolo As Integer
Private pPanturrilhaMedial As Double
Private pCoxa As Double
Private pAbdominal As Double
Private pSuprailiaca As Double
Private pAxilarMedia As Double
Private pToracica As Double
Private pBiceps As Double
Private pTriceps As Double
Private pSubescapular As Double

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
' CODIGO DO PROTOCOLO
'==================================================

Public Property Get codigoProtocolo() As Integer
    codigoProtocolo = pCodigoProtocolo
End Property

Public Property Let codigoProtocolo(ByVal valor As Integer)
    pCodigoProtocolo = valor
End Property


'==================================================
' PANTURRILHA MEDIAL
'==================================================

Public Property Get panturrilhaMedial() As Double
    panturrilhaMedial = pPanturrilhaMedial
End Property

Public Property Let panturrilhaMedial(ByVal valor As Double)
    pPanturrilhaMedial = valor
End Property

'==================================================
' COXA
'==================================================

Public Property Get coxa() As Double
    coxa = pCoxa
End Property

Public Property Let coxa(ByVal valor As Double)
    pCoxa = valor
End Property

'==================================================
' ABDOMINAL
'==================================================

Public Property Get abdominal() As Double
    abdominal = pAbdominal
End Property

Public Property Let abdominal(ByVal valor As Double)
    pAbdominal = valor
End Property

'==================================================
' SUPRAILIACA
'==================================================

Public Property Get suprailiaca() As Double
    suprailiaca = pSuprailiaca
End Property

Public Property Let suprailiaca(ByVal valor As Double)
    pSuprailiaca = valor
End Property

'==================================================
' AXILAR MEDIA
'==================================================

Public Property Get axilarMedia() As Double
    axilarMedia = pAxilarMedia
End Property

Public Property Let axilarMedia(ByVal valor As Double)
    pAxilarMedia = valor
End Property

'==================================================
' TORACICA
'==================================================

Public Property Get toracica() As Double
    toracica = pToracica
End Property

Public Property Let toracica(ByVal valor As Double)
    pToracica = valor
End Property

'==================================================
' BICEPS
'==================================================

Public Property Get biceps() As Double
    biceps = pBiceps
End Property

Public Property Let biceps(ByVal valor As Double)
    pBiceps = valor
End Property

'==================================================
' TRICEPS
'==================================================

Public Property Get triceps() As Double
    triceps = pTriceps
End Property

Public Property Let triceps(ByVal valor As Double)
    pTriceps = valor
End Property

'==================================================
' SUBESCAPULAR
'==================================================

Public Property Get subescapular() As Double
    subescapular = pSubescapular
End Property

Public Property Let subescapular(ByVal valor As Double)
    pSubescapular = valor
End Property
