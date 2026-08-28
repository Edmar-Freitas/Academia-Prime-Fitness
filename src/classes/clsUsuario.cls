VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigo As Long
Private pLogin As String
Private pSenhaHash As String
Private pSalt As String
Private pNome As String
Private pPerfil As String
Private pAtivo As Boolean

Public Property Get codigo() As Long
    codigo = pCodigo
End Property

Public Property Let codigo(ByVal valor As Long)
    pCodigo = valor
End Property

Public Property Get login() As String
    login = pLogin
End Property

Public Property Let login(ByVal valor As String)
    pLogin = UCase$(Trim$(valor))
End Property

Public Property Get senhaHash() As String
    senhaHash = pSenhaHash
End Property

Public Property Let senhaHash(ByVal valor As String)
    pSenhaHash = valor
End Property

Public Property Get salt() As String
    salt = pSalt
End Property

Public Property Let salt(ByVal valor As String)
    pSalt = valor
End Property

Public Property Get nome() As String
    nome = pNome
End Property

Public Property Let nome(ByVal valor As String)
    pNome = valor
End Property

Public Property Get perfil() As String
    perfil = pPerfil
End Property

Public Property Let perfil(ByVal valor As String)
    pPerfil = UCase$(Trim$(valor))
End Property

Public Property Get ativo() As Boolean
    ativo = pAtivo
End Property

Public Property Let ativo(ByVal valor As Boolean)
    pAtivo = valor
End Property
