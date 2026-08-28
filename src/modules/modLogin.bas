Attribute VB_Name = "modLogin"
Option Explicit

'==========================================================
' USUÁRIO ATUALMENTE AUTENTICADO
'==========================================================

Public UsuarioLogado As clsUsuario


'==========================================================
' AUTENTICA USUÁRIO
'==========================================================

Public Function Autenticar(ByVal login As String, _
                           ByVal senha As String) As Boolean

    Dim usuario As clsUsuario
    Dim senhaHash As String

    '---------------------------------------------
    ' Estado inicial
    '---------------------------------------------

    Autenticar = False


    '---------------------------------------------
    ' 1. VALIDA LOGIN
    '---------------------------------------------

    login = UCase$(Trim$(login))

    If login = vbNullString Then Exit Function


    '---------------------------------------------
    ' 2. VALIDA SENHA
    '---------------------------------------------

    If senha = vbNullString Then Exit Function


    '---------------------------------------------
    ' 3. BUSCA USUÁRIO
    '---------------------------------------------

    Set usuario = BuscarUsuario(login)

    If usuario Is Nothing Then Exit Function


    '---------------------------------------------
    ' 4. VERIFICA SE ESTÁ ATIVO
    '---------------------------------------------

    If Not usuario.ativo Then Exit Function


    '---------------------------------------------
    ' 5. GERA HASH DA SENHA INFORMADA
    '---------------------------------------------

    senhaHash = GerarHashSenha( _
                    senha, _
                    usuario.salt)


    '---------------------------------------------
    ' 6. COMPARA HASH
    '---------------------------------------------

    If StrComp( _
            senhaHash, _
            usuario.senhaHash, _
            vbBinaryCompare) <> 0 Then

        Exit Function

    End If


    '---------------------------------------------
    ' 7. REGISTRA USUÁRIO AUTENTICADO
    '---------------------------------------------

    Set UsuarioLogado = usuario


    '---------------------------------------------
    ' 8. AUTENTICAÇÃO BEM-SUCEDIDA
    '---------------------------------------------

    Autenticar = True

End Function


'==========================================================
' ENCERRA A SESSÃO
'==========================================================

Public Sub EncerrarSessao()

    Set UsuarioLogado = Nothing

End Sub


'==========================================================
' TESTE
'==========================================================

Public Sub TestarAutenticacao()

    Dim resultado As Boolean

    resultado = Autenticar("ADMI", "123456")

    Debug.Print "Resultado da autenticação: " & resultado

    If resultado Then

        Debug.Print "Usuário autenticado:"
        Debug.Print "Código: " & UsuarioLogado.codigo
        Debug.Print "Login: " & UsuarioLogado.login
        Debug.Print "Nome: " & UsuarioLogado.nome
        Debug.Print "Perfil: " & UsuarioLogado.perfil

    End If

End Sub

