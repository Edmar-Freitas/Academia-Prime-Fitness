VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAvaliacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pCodigo As Long
Private pCodigoAluno As Long
Private pNomeAluno As String
Private pCodigoAvaliador As Long
Private pNomeAvaliador As Long
Private pDataAvaliacao As Date
Private pStatus As String

Private pAluno As clsAluno
Private pAvaliador As clsFuncionario
Private pAntropometria As clsAntropometria
Private pCircunferencias As clsCircunferencias
Private pDobrasCultaneas As clsDobrasCultaneas
Private pCardiorrespiratoria As clsCardiorrespiratoria
Private pDensidadeOssea As clsDensidadeOssea

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
' CODIGO DO ALUNO
'==================================================

Public Property Get codigoAluno() As Long
    codigoAluno = pCodigoAluno
End Property

Public Property Let codigoAluno(ByVal valor As Long)
    pCodigoAluno = valor
End Property

'==================================================
' NOME DO ALUNO
'==================================================

Public Property Get nomeAluno() As String
    nomeAluno = pNomeAluno
End Property

Public Property Let nomeAluno(ByVal valor As String)
    pNomeAluno = valor
End Property

'==================================================
' CODIGO DO AVALIADOR
'==================================================

Public Property Get codigoAvaliador() As Long
    codigoAvaliador = pCodigoAvaliador
End Property

Public Property Let codigoAvaliador(ByVal valor As Long)
    pCodigoAvaliador = valor
End Property

'==================================================
' NOME DO AVALIADOR
'==================================================

Public Property Get nomeAvaliador() As String
    nomeAvaliador = pNomeAvaliador
End Property

Public Property Let nomeAvaliador(ByVal valor As String)
    pNomeAvaliador = valor
End Property

'==================================================
' DATA DA AVALIACAO
'==================================================

Public Property Get dataAvaliacao() As Date
    dataAvaliacao = pDataAvaliacao
End Property

Public Property Let dataAvaliacao(ByVal valor As Date)
    pDataAvaliacao = valor
End Property

'==================================================
' STATUS
'==================================================

Public Property Get status() As String
    status = pStatus
End Property

Public Property Let status(ByVal valor As String)
    pStatus = valor
End Property



'==================================================
' ANTROPOMETRIA
'==================================================

Public Property Get antropometria() As clsAntropometria
    Set antropometria = pAntropometria
End Property

Public Property Set antropometria(ByVal valor As clsAntropometria)
    Set pAntropometria = valor
End Property

'==================================================
' CIRCUNFERENCIAS
'==================================================

Public Property Get circunferencias() As clsCircunferencias
    Set circunferencias = pCircunferencias
End Property

Public Property Set circunferencias(ByVal valor As clsCircunferencias)
    Set pCircunferencias = valor
End Property

'==================================================
' DOBRAS CULTANEAS
'==================================================

Public Property Get dobrasCultaneas() As clsDobrasCultaneas
    Set dobrasCultaneas = pDobrasCultaneas
End Property

Public Property Set dobrasCultaneas(ByVal valor As clsDobrasCultaneas)
    Set pDobrasCultaneas = valor
End Property

'==================================================
' CARDIORRESPIRATORIA
'==================================================

Public Property Get cardiorrespiratoria() As clsCardiorrespiratoria
    Set cardiorrespiratoria = pCardiorrespiratoria
End Property

Public Property Set cardiorrespiratoria(ByVal valor As clsCardiorrespiratoria)
    Set pCardiorrespiratoria = valor
End Property

'==================================================
' DENSIDADE OSSEA
'==================================================

Public Property Get densidadeOssea() As clsDensidadeOssea
    Set densidadeOssea = pDensidadeOssea
End Property

Public Property Set densidadeOssea(ByVal valor As clsDensidadeOssea)
    Set pDensidadeOssea = valor
End Property

'==================================================
' ALUNO
'==================================================

Public Property Get aluno() As clsAluno
    Set aluno = pAluno
End Property

Public Property Set aluno(ByVal valor As clsAluno)
    Set pAluno = valor
End Property

'==================================================
' AVALIADOR
'==================================================

Public Property Get avaliador() As clsFuncionario
    Set avaliador = pAvaliador
End Property

Public Property Set avaliador(ByVal valor As clsFuncionario)
    Set pAvaliador = valor
End Property
