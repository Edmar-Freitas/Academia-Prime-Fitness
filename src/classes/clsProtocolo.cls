VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProtocolo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigo As Long
Private pNome As String
Private pDescricao As String


'==================================================
' CODIGO
'==================================================

Public Property Get codigo() As Long
    codigo = pCodigo
End Property

Public Property Let codigo(ByVal valor As Long)
    pCodigo = valor
End Property

'==================================================
' NOME
'==================================================

Public Property Get nome() As String
    nome = pNome
End Property

Public Property Let nome(ByVal valor As String)
    pNome = valor
End Property

'==================================================
' DESCRICAO
'==================================================

Public Property Get descricao() As String
    descricao = pDescricao
End Property

Public Property Let descricao(ByVal valor As String)
    pDescricao = valor
End Property
