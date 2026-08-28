VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigo As Long
Private pNome As String
Private pDataNascimento As Date
Private pIdade As Integer
Private pGenero As String
Private pTelefone As String
Private pTelefoneEmergencial As String
Private pEmail As String
Private pCidade As String
Private pBairro As String
Private pLogradouro As String
Private pNumero As String
Private pCEP As String
Private pRG As String
Private pCPF As String
Private pCREF As String
Private pFotoCaminho As String
Private pObservacao As String
Private pFuncao As String
Private pSituacao As String

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
' DATA DE NASCIMENTO
'==================================================

Public Property Get DataNascimento() As Date
    DataNascimento = pDataNascimento
End Property

Public Property Let DataNascimento(ByVal valor As Date)
    pDataNascimento = valor
End Property


'==================================================
' IDADE
'==================================================

Public Property Get Idade() As Integer
  
    If pDataNascimento > 0 Then
        Idade = Year(Date) - Year(pDataNascimento)

        If Date < DateSerial(Year(Date), _
                             Month(pDataNascimento), _
                             Day(pDataNascimento)) Then
            Idade = Idade - 1
        End If
    Else
        Idade = 0
    End If
End Property



'==================================================
' GENERO
'==================================================

Public Property Get Genero() As String
    Genero = pGenero
End Property

Public Property Let Genero(ByVal valor As String)
    pGenero = valor
End Property


'==================================================
' TELEFONE
'==================================================

Public Property Get Telefone() As String
    Telefone = pTelefone
End Property

Public Property Let Telefone(ByVal valor As String)
    pTelefone = valor
End Property


'==================================================
' TELEFONE EMERGENCIAL
'==================================================

Public Property Get TelefoneEmergencial() As String
    TelefoneEmergencial = pTelefoneEmergencial
End Property

Public Property Let TelefoneEmergencial(ByVal valor As String)
    pTelefoneEmergencial = valor
End Property


'==================================================
' EMAIL
'==================================================

Public Property Get Email() As String
    Email = pEmail
End Property

Public Property Let Email(ByVal valor As String)
    pEmail = valor
End Property


'==================================================
' CIDADE
'==================================================

Public Property Get Cidade() As String
    Cidade = pCidade
End Property

Public Property Let Cidade(ByVal valor As String)
    pCidade = valor
End Property


'==================================================
' BAIRRO
'==================================================

Public Property Get Bairro() As String
    Bairro = pBairro
End Property

Public Property Let Bairro(ByVal valor As String)
    pBairro = valor
End Property


'==================================================
' LOGRADOURO
'==================================================

Public Property Get Logradouro() As String
    Logradouro = pLogradouro
End Property

Public Property Let Logradouro(ByVal valor As String)
    pLogradouro = valor
End Property


'==================================================
' NUMERO
'==================================================

Public Property Get Numero() As String
    Numero = pNumero
End Property

Public Property Let Numero(ByVal valor As String)
    pNumero = valor
End Property


'==================================================
' CEP
'==================================================

Public Property Get CEP() As String
    CEP = pCEP
End Property

Public Property Let CEP(ByVal valor As String)
    pCEP = valor
End Property


'==================================================
' RG
'==================================================

Public Property Get RG() As String
    RG = pRG
End Property

Public Property Let RG(ByVal valor As String)
    pRG = valor
End Property


'==================================================
' CPF
'==================================================

Public Property Get CPF() As String
    CPF = pCPF
End Property

Public Property Let CPF(ByVal valor As String)
    pCPF = valor
End Property

'==================================================
' CREF
'==================================================

Public Property Get CREF() As String
    CREF = pCREF
End Property

Public Property Let CREF(ByVal valor As String)
    pCREF = valor
End Property


'==================================================
' FOTOCAMINHO
'==================================================

Public Property Get FotoCaminho() As String
    FotoCaminho = pFotoCaminho
End Property

Public Property Let FotoCaminho(ByVal valor As String)
    pFotoCaminho = valor
End Property


'==================================================
' OBSERVACAO
'==================================================

Public Property Get Observacao() As String
    Observacao = pObservacao
End Property

Public Property Let Observacao(ByVal valor As String)
    pObservacao = valor
End Property

'==================================================
' FUNCAO
'==================================================

Public Property Get Funcao() As String
    Funcao = pFuncao
End Property

Public Property Let Funcao(ByVal valor As String)
    pFuncao = valor
End Property

'==================================================
' SITUACAO
'==================================================

Public Property Get Situacao() As String
    Situacao = pSituacao
End Property

Public Property Let Situacao(ByVal valor As String)
    pSituacao = valor
End Property


