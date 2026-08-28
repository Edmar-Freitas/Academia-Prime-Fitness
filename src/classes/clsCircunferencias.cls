VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCircunferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pCodigoAvaliacao As Long

Private pTorax As Double

Private pBracoDireito As Double
Private pBracoEsquerdo As Double

Private pAntebracoDireito As Double
Private pAntebracoEsquerdo As Double

Private pAbdominal As Double
Private pCintura As Double
Private pQuadril As Double

Private pCoxaDireita As Double
Private pCoxaEsquerda As Double

Private pPernaDireita As Double
Private pPernaEsquerda As Double


'========================================
' Código da Avaliação
'========================================

Public Property Get codigoAvaliacao() As Long
    codigoAvaliacao = pCodigoAvaliacao
End Property

Public Property Let codigoAvaliacao(ByVal valor As Long)
    pCodigoAvaliacao = valor
End Property


'========================================
' Tórax
'========================================

Public Property Get torax() As Double
    torax = pTorax
End Property

Public Property Let torax(ByVal valor As Double)
    pTorax = valor
End Property


'========================================
' Braço Direito
'========================================

Public Property Get bracoDireito() As Double
    bracoDireito = pBracoDireito
End Property

Public Property Let bracoDireito(ByVal valor As Double)
    pBracoDireito = valor
End Property


'========================================
' Braço Esquerdo
'========================================

Public Property Get bracoEsquerdo() As Double
    bracoEsquerdo = pBracoEsquerdo
End Property

Public Property Let bracoEsquerdo(ByVal valor As Double)
    pBracoEsquerdo = valor
End Property


'========================================
' Antebraço Direito
'========================================

Public Property Get antebracoDireito() As Double
    antebracoDireito = pAntebracoDireito
End Property

Public Property Let antebracoDireito(ByVal valor As Double)
    pAntebracoDireito = valor
End Property


'========================================
' Antebraço Esquerdo
'========================================

Public Property Get antebracoEsquerdo() As Double
    antebracoEsquerdo = pAntebracoEsquerdo
End Property

Public Property Let antebracoEsquerdo(ByVal valor As Double)
    pAntebracoEsquerdo = valor
End Property


'========================================
' Abdominal
'========================================

Public Property Get abdominal() As Double
    abdominal = pAbdominal
End Property

Public Property Let abdominal(ByVal valor As Double)
    pAbdominal = valor
End Property


'========================================
' Cintura
'========================================

Public Property Get cintura() As Double
    cintura = pCintura
End Property

Public Property Let cintura(ByVal valor As Double)
    pCintura = valor
End Property


'========================================
' Quadril
'========================================

Public Property Get quadril() As Double
    quadril = pQuadril
End Property

Public Property Let quadril(ByVal valor As Double)
    pQuadril = valor
End Property


'========================================
' Coxa Direita
'========================================

Public Property Get coxaDireita() As Double
    coxaDireita = pCoxaDireita
End Property

Public Property Let coxaDireita(ByVal valor As Double)
    pCoxaDireita = valor
End Property


'========================================
' Coxa Esquerda
'========================================

Public Property Get coxaEsquerda() As Double
    coxaEsquerda = pCoxaEsquerda
End Property

Public Property Let coxaEsquerda(ByVal valor As Double)
    pCoxaEsquerda = valor
End Property


'========================================
' Perna Direita
'========================================

Public Property Get pernaDireita() As Double
    pernaDireita = pPernaDireita
End Property

Public Property Let pernaDireita(ByVal valor As Double)
    pPernaDireita = valor
End Property


'========================================
' Perna Esquerda
'========================================

Public Property Get pernaEsquerda() As Double
    pernaEsquerda = pPernaEsquerda
End Property

Public Property Let pernaEsquerda(ByVal valor As Double)
    pPernaEsquerda = valor
End Property
