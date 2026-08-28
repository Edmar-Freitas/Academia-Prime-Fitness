Attribute VB_Name = "modGit"
Option Explicit

'========================================================
' EXECUTA UM COMANDO DO WINDOWS
'========================================================

Public Function ExecutarGit( _
    ByVal comando As String, _
    Optional ByRef codigoSaida As Long = 0) As String

    Dim shell As Object
    Dim exec As Object

    Dim saida As String
    Dim erro As String

    Set shell = CreateObject("WScript.Shell")

    Set exec = shell.exec(comando)

    Do While exec.status = 0
        DoEvents
    Loop

    saida = exec.StdOut.ReadAll
    erro = exec.StdErr.ReadAll

    codigoSaida = exec.ExitCode

    If Len(erro) > 0 Then

        If Len(saida) > 0 Then
            saida = saida & vbCrLf
        End If

        saida = saida & erro

    End If

    ExecutarGit = Trim$(saida)

End Function


'========================================================
' RETORNA O CAMINHO DO REPOSITÓRIO
'========================================================

Private Function CaminhoRepositorio() As String

    CaminhoRepositorio = ThisWorkbook.Path

End Function


'========================================================
' STATUS DO REPOSITÓRIO
'========================================================

Public Function GitStatus( _
    Optional ByRef codigoSaida As Long = 0) As String

    Dim caminhoProjeto As String

    caminhoProjeto = CaminhoRepositorio()

    GitStatus = ExecutarGit( _
        "cmd /c cd /d """ & _
        caminhoProjeto & _
        """ && git status --short", _
        codigoSaida _
    )

End Function


'========================================================
' VERIFICA SE EXISTEM ALTERAÇÕES
'========================================================

Public Function GitTemAlteracoes() As Boolean

    Dim resultado As String
    Dim codigo As Long

    resultado = GitStatus(codigo)

    If codigo <> 0 Then

        MsgBox _
            "Não foi possível consultar o Git." & _
            vbCrLf & vbCrLf & _
            resultado, _
            vbCritical, _
            "Git - Erro"

        Exit Function

    End If

    GitTemAlteracoes = _
        Len(Trim$(resultado)) > 0

End Function


'========================================================
' ADICIONAR ARQUIVOS
'========================================================

Public Function GitAdd( _
    Optional ByRef codigoSaida As Long = 0) As String

    Dim caminhoProjeto As String

    caminhoProjeto = CaminhoRepositorio()

    GitAdd = ExecutarGit( _
        "cmd /c cd /d """ & _
        caminhoProjeto & _
        """ && git add .", _
        codigoSaida _
    )

End Function


'========================================================
' COMMIT
'========================================================

Public Function GitCommit( _
    ByVal mensagem As String, _
    Optional ByRef codigoSaida As Long = 0) As String

    Dim caminhoProjeto As String

    caminhoProjeto = CaminhoRepositorio()

    GitCommit = ExecutarGit( _
        "cmd /c cd /d """ & _
        caminhoProjeto & _
        """ && git commit -m """ & _
        mensagem & _
        """", _
        codigoSaida _
    )

End Function


'========================================================
' PUSH
'========================================================

Public Function GitPush( _
    Optional ByRef codigoSaida As Long = 0) As String

    Dim caminhoProjeto As String

    caminhoProjeto = CaminhoRepositorio()

    GitPush = ExecutarGit( _
        "cmd /c cd /d """ & _
        caminhoProjeto & _
        """ && git push origin main", _
        codigoSaida _
    )

End Function


'========================================================
' TESTAR STATUS
'========================================================

Public Sub TestarGitStatus()

    Dim resultado As String
    Dim codigo As Long

    resultado = GitStatus(codigo)

    If codigo <> 0 Then

        MsgBox _
            "Erro ao consultar o Git." & _
            vbCrLf & vbCrLf & _
            resultado, _
            vbCritical, _
            "Git - Status"

        Exit Sub

    End If

    If Len(Trim$(resultado)) = 0 Then

        resultado = _
            "Nenhuma alteração detectada." & _
            vbCrLf & vbCrLf & _
            "Working tree limpa."

    End If

    MsgBox _
        resultado, _
        vbInformation, _
        "Git - Status"

End Sub


'========================================================
' TESTAR GIT ADD
'========================================================

Public Sub TestarGitAdd()

    Dim resultado As String
    Dim codigo As Long

    resultado = GitAdd(codigo)

    If codigo <> 0 Then

        MsgBox _
            "Erro ao executar git add." & _
            vbCrLf & vbCrLf & _
            resultado, _
            vbCritical, _
            "Git - Add"

        Exit Sub

    End If

    MsgBox _
        "Git add executado com sucesso.", _
        vbInformation, _
        "Git - Add"

End Sub


'========================================================
' TESTE
'========================================================

Public Sub TesteSincronizacao()

    MsgBox _
        "Infraestrutura Git funcionando.", _
        vbInformation, _
        "Prime Fitness"

End Sub


Public Sub TesteAlteracaoGit()

    MsgBox "Alteração detectada pelo Git.", _
           vbInformation, _
           "Teste"

End Sub
