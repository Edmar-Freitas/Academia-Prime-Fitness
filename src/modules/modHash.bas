Attribute VB_Name = "modHash"
Option Explicit

'==========================================================
' CONSTANTES DA API CNG
'==========================================================

Private Const BCRYPT_SHA256_ALGORITHM As String = "SHA256"
Private Const BCRYPT_OBJECT_LENGTH As String = "ObjectLength"
Private Const BCRYPT_HASH_LENGTH As String = "HashDigestLength"

Private Const STATUS_SUCCESS As Long = 0&


'==========================================================
' DECLARAÇÕES DA API DO WINDOWS
' Compatível com VBA7 / Office 32 bits / Office 64 bits
'==========================================================

#If VBA7 Then

    Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
        ByRef phAlgorithm As LongPtr, _
        ByVal pszAlgId As LongPtr, _
        ByVal pszImplementation As LongPtr, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" ( _
        ByVal hObject As LongPtr, _
        ByVal pszProperty As LongPtr, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByRef pcbResult As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, _
        ByRef phHash As LongPtr, _
        ByVal pbHashObject As LongPtr, _
        ByVal cbHashObject As Long, _
        ByVal pbSecret As LongPtr, _
        ByVal cbSecret As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr, _
        ByRef pbInput As Any, _
        ByVal cbInput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" ( _
        ByVal hHash As LongPtr) As Long
        
    Private Declare PtrSafe Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As LongPtr, _
        ByRef pbBuffer As Any, _
        ByVal cbBuffer As Long, _
        ByVal dwFlags As Long) As Long

#Else

    'Compatibilidade com versões antigas do VBA
    'que não possuem VBA7.

    Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" ( _
        ByRef phAlgorithm As Long, _
        ByVal pszAlgId As Long, _
        ByVal pszImplementation As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptGetProperty Lib "bcrypt.dll" ( _
        ByVal hObject As Long, _
        ByVal pszProperty As Long, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByRef pcbResult As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptCreateHash Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, _
        ByRef phHash As Long, _
        ByVal pbHashObject As Long, _
        ByVal cbHashObject As Long, _
        ByVal pbSecret As Long, _
        ByVal cbSecret As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptHashData Lib "bcrypt.dll" ( _
        ByVal hHash As Long, _
        ByRef pbInput As Any, _
        ByVal cbInput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptFinishHash Lib "bcrypt.dll" ( _
        ByVal hHash As Long, _
        ByRef pbOutput As Any, _
        ByVal cbOutput As Long, _
        ByVal dwFlags As Long) As Long

    Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" ( _
        ByVal hHash As Long) As Long
    
    Private Declare Function BCryptGenRandom Lib "bcrypt.dll" ( _
        ByVal hAlgorithm As Long, _
        ByRef pbBuffer As Any, _
        ByVal cbBuffer As Long, _
        ByVal dwFlags As Long) As Long

#End If


'==========================================================
' FUNÇÃO PRINCIPAL
'
' Recebe:
'     texto
'
' Retorna:
'     SHA-256 em hexadecimal
'
' Exemplo:
'
'     GerarSHA256("123456")
'
'==========================================================

Public Function GerarSHA256(ByVal texto As String) As String

#If VBA7 Then

    Dim hAlgorithm As LongPtr
    Dim hHash As LongPtr

#Else

    Dim hAlgorithm As Long
    Dim hHash As Long

#End If

    Dim hashLength As Long
    Dim objectLength As Long
    Dim resultLength As Long

    Dim hashObject() As Byte
    Dim hashBytes() As Byte

    Dim dados() As Byte

    Dim propriedade As String
    Dim retorno As Long

    Dim i As Long
    Dim resultado As String


    On Error GoTo TrataErro


    '------------------------------------------------------
    ' Converte o texto para bytes UTF-8
    '------------------------------------------------------

    dados = TextoParaUTF8(texto)


    '------------------------------------------------------
    ' Abre o provedor SHA-256
    '------------------------------------------------------

    retorno = BCryptOpenAlgorithmProvider( _
        hAlgorithm, _
        StrPtr(BCRYPT_SHA256_ALGORITHM), _
        0, _
        0)

    If retorno <> STATUS_SUCCESS Then
        Err.Raise vbObjectError + 1000, _
                  "GerarSHA256", _
                  "Não foi possível abrir o provedor SHA-256."
    End If


    '------------------------------------------------------
    ' Obtém o tamanho do objeto interno do hash
    '------------------------------------------------------

    propriedade = BCRYPT_OBJECT_LENGTH

    retorno = BCryptGetProperty( _
        hAlgorithm, _
        StrPtr(propriedade), _
        objectLength, _
        4, _
        resultLength, _
        0)

    If retorno <> STATUS_SUCCESS Then
        Err.Raise vbObjectError + 1001, _
                  "GerarSHA256", _
                  "Não foi possível obter o tamanho do objeto SHA-256."
    End If


    ReDim hashObject(0 To objectLength - 1)


    '------------------------------------------------------
    ' Cria o objeto hash
    '------------------------------------------------------

    retorno = BCryptCreateHash( _
        hAlgorithm, _
        hHash, _
        VarPtr(hashObject(0)), _
        objectLength, _
        0, _
        0, _
        0)

    If retorno <> STATUS_SUCCESS Then
        Err.Raise vbObjectError + 1002, _
                  "GerarSHA256", _
                  "Não foi possível criar o objeto SHA-256."
    End If


    '------------------------------------------------------
    ' Envia os dados para o SHA-256
    '------------------------------------------------------

    If UBound(dados) >= 0 Then

        retorno = BCryptHashData( _
            hHash, _
            dados(0), _
            UBound(dados) + 1, _
            0)

        If retorno <> STATUS_SUCCESS Then
            Err.Raise vbObjectError + 1003, _
                      "GerarSHA256", _
                      "Não foi possível processar os dados."

        End If

    End If


    '------------------------------------------------------
    ' Obtém o tamanho do hash
    '------------------------------------------------------

    propriedade = BCRYPT_HASH_LENGTH

    retorno = BCryptGetProperty( _
        hAlgorithm, _
        StrPtr(propriedade), _
        hashLength, _
        4, _
        resultLength, _
        0)

    If retorno <> STATUS_SUCCESS Then
        Err.Raise vbObjectError + 1004, _
                  "GerarSHA256", _
                  "Não foi possível obter o tamanho do hash."
    End If


    ReDim hashBytes(0 To hashLength - 1)


    '------------------------------------------------------
    ' Finaliza o hash
    '------------------------------------------------------

    retorno = BCryptFinishHash( _
        hHash, _
        hashBytes(0), _
        hashLength, _
        0)

    If retorno <> STATUS_SUCCESS Then
        Err.Raise vbObjectError + 1005, _
                  "GerarSHA256", _
                  "Não foi possível finalizar o SHA-256."
    End If


    '------------------------------------------------------
    ' Converte os bytes para hexadecimal
    '------------------------------------------------------

    For i = 0 To hashLength - 1

        resultado = resultado & _
                    LCase$(Right$("0" & _
                    Hex$(hashBytes(i)), 2))

    Next i


    GerarSHA256 = resultado


Saida:

    '------------------------------------------------------
    ' Libera os recursos da API
    '------------------------------------------------------

    If hHash <> 0 Then
        BCryptDestroyHash hHash
    End If

    If hAlgorithm <> 0 Then
        BCryptCloseAlgorithmProvider hAlgorithm, 0
    End If

    Exit Function


TrataErro:

    Debug.Print "ERRO:"
    Debug.Print "Número: " & Err.Number
    Debug.Print "Descrição: " & Err.Description

    GerarSHA256 = vbNullString

    Resume Saida

End Function


'==========================================================
' CONVERTE STRING VBA PARA UTF-8
'==========================================================

Private Function TextoParaUTF8(ByVal texto As String) As Byte()

    Dim bytes() As Byte

    Dim i As Long
    Dim codigo As Long
    Dim pos As Long

    Dim resultado() As Byte


    'Reserva espaço inicial
    ReDim resultado(0 To 0)

    pos = -1


    For i = 1 To Len(texto)

        codigo = AscW(Mid$(texto, i, 1))

        If codigo < 0 Then
            codigo = codigo + 65536
        End If


        '--------------------------------------------------
        ' UTF-8: 1 byte
        '--------------------------------------------------

        If codigo <= &H7F Then

            pos = pos + 1
            ReDim Preserve resultado(0 To pos)

            resultado(pos) = codigo


        '--------------------------------------------------
        ' UTF-8: 2 bytes
        '--------------------------------------------------

        ElseIf codigo <= &H7FF Then

            pos = pos + 1
            ReDim Preserve resultado(0 To pos)
            resultado(pos) = &HC0 Or (codigo \ 64)

            pos = pos + 1
            ReDim Preserve resultado(0 To pos)
            resultado(pos) = &H80 Or (codigo And &H3F)


        '--------------------------------------------------
        ' UTF-8: 3 bytes
        '--------------------------------------------------

        Else

            pos = pos + 1
            ReDim Preserve resultado(0 To pos)
            resultado(pos) = &HE0 Or (codigo \ 4096)

            pos = pos + 1
            ReDim Preserve resultado(0 To pos)
            resultado(pos) = &H80 Or ((codigo \ 64) And &H3F)

            pos = pos + 1
            ReDim Preserve resultado(0 To pos)
            resultado(pos) = &H80 Or (codigo And &H3F)

        End If

    Next i


    If pos = -1 Then
        ReDim resultado(0 To 0)
        resultado(0) = 0
    End If


    TextoParaUTF8 = resultado

End Function

Public Function GerarSalt(Optional ByVal quantidadeBytes As Long = 16) As String

    Dim bytesSalt() As Byte
    Dim i As Long
    Dim resultado As String

    If quantidadeBytes <= 0 Then
        Err.Raise vbObjectError + 1100, _
                  "GerarSalt", _
                  "A quantidade de bytes deve ser maior que zero."
    End If

    ReDim bytesSalt(0 To quantidadeBytes - 1)

    If BCryptGenRandom( _
        0, _
        bytesSalt(0), _
        quantidadeBytes, _
        2) <> STATUS_SUCCESS Then

        Err.Raise vbObjectError + 1101, _
                  "GerarSalt", _
                  "Não foi possível gerar o Salt."
    End If

    For i = 0 To quantidadeBytes - 1

        resultado = resultado & _
                    LCase$(Right$("0" & _
                    Hex$(bytesSalt(i)), 2))

    Next i

    GerarSalt = resultado

End Function

Public Function GerarHashSenha(ByVal senha As String, _
                               ByVal salt As String) As String

    GerarHashSenha = GerarSHA256(salt & senha)

End Function


Public Sub TestarBuscarUsuario()

    Dim usuario As clsUsuario

    Set usuario = BuscarUsuario("ADMIN")

    If usuario Is Nothing Then

        Debug.Print "Usuário não encontrado."
        Exit Sub

    End If

    Debug.Print "===== USUÁRIO ====="
    Debug.Print "Código: " & usuario.codigo
    Debug.Print "Login: " & usuario.login
    Debug.Print "Salt: " & usuario.salt
    Debug.Print "SenhaHash: " & usuario.senhaHash
    Debug.Print "Nome: " & usuario.nome
    Debug.Print "Perfil: " & usuario.perfil
    Debug.Print "Ativo: " & usuario.ativo

End Sub

Public Sub GerarDadosUsuarioTeste()

    Dim salt As String
    Dim senhaHash As String

    salt = GerarSalt()

    senhaHash = GerarHashSenha( _
                    "123456", _
                    salt)

    Debug.Print "SALT:"
    Debug.Print salt

    Debug.Print "SENHA_HASH:"
    Debug.Print senhaHash

End Sub


'Public Sub TestarHashSenha()

'    Dim senha As String
'    Dim salt1 As String
'    Dim salt2 As String

'    Dim hash1 As String
'    Dim hash2 As String
'    Dim hash3 As String

'    senha = "123456"

'    salt1 = GerarSalt()
'    salt2 = GerarSalt()

'    hash1 = GerarHashSenha(senha, salt1)
'    hash2 = GerarHashSenha(senha, salt1)
'    hash3 = GerarHashSenha(senha, salt2)

'    Debug.Print "Senha: " & senha

'    Debug.Print "Salt 1: " & salt1
'    Debug.Print "Salt 2: " & salt2

'    Debug.Print "Hash 1: " & hash1
'    Debug.Print "Hash 2: " & hash2
'    Debug.Print "Hash 3: " & hash3

'    Debug.Print "Tamanho Hash 1: " & Len(hash1)
'    Debug.Print "Hash 1 = Hash 2: " & (hash1 = hash2)
'    Debug.Print "Hash 1 = Hash 3: " & (hash1 = hash3)

'End Sub

'Public Sub TestarSalt()

'    Dim salt1 As String
'    Dim salt2 As String

'    salt1 = GerarSalt()
'    salt2 = GerarSalt()

'    Debug.Print "Salt 1: " & salt1
'    Debug.Print "Salt 2: " & salt2

'    Debug.Print "Tamanho Salt 1: " & Len(salt1)
'    Debug.Print "Tamanho Salt 2: " & Len(salt2)

'    Debug.Print "Salt 1 = Salt 2: " & (salt1 = salt2)

'End Sub


'Public Sub TestarHash()

'    Dim hash1 As String
'   Dim hash2 As String
'    Dim hash3 As String

'    hash1 = GerarSHA256("123456")
'    hash2 = GerarSHA256("123456")
'    hash3 = GerarSHA256("123457")

'    Debug.Print "Hash 1: [" & hash1 & "]"
'    Debug.Print "Hash 2: [" & hash2 & "]"
'    Debug.Print "Hash 3: [" & hash3 & "]"

'    Debug.Print "Tamanho: " & Len(hash1)
'    Debug.Print "Hash 1 = Hash 2: " & (hash1 = hash2)
'    Debug.Print "Hash 1 = Hash 3: " & (hash1 = hash3)

'End Sub
