Attribute VB_Name = "modGit"
Option Explicit

'========================================================
' EXECUTA UM COMANDO DO WINDOWS
'========================================================

Public Function ExecutarGit(ByVal comando As String) As String

    Dim shell As Object
    Dim exec As Object
    Dim resultado As String
    
    Set shell = CreateObject("WScript.Shell")
    
    Set exec = shell.exec(comando)
    
    Do While exec.status = 0
        DoEvents
    Loop
    
    resultado = exec.StdOut.ReadAll
    
    If Not exec.StdErr.AtEndOfStream Then
        resultado = resultado & vbCrLf & exec.StdErr.ReadAll
    End If
    
    ExecutarGit = resultado

End Function


'========================================================
' STATUS DO REPOSITÓRIO
'========================================================

Public Function GitStatus() As String

    Dim caminhoProjeto As String
    
    caminhoProjeto = ThisWorkbook.Path
    
    GitStatus = ExecutarGit( _
        "cmd /c cd /d """ & caminhoProjeto & _
        """ && git status --short" _
    )

End Function


'========================================================
' ADICIONAR ARQUIVOS
'========================================================

Public Function GitAdd() As String

    Dim caminhoProjeto As String
    
    caminhoProjeto = ThisWorkbook.Path
    
    GitAdd = ExecutarGit( _
        "cmd /c cd /d """ & caminhoProjeto & _
        """ && git add ." _
    )

End Function


'========================================================
' COMMIT
'========================================================

Public Function GitCommit(ByVal mensagem As String) As String

    Dim caminhoProjeto As String
    
    caminhoProjeto = ThisWorkbook.Path
    
    GitCommit = ExecutarGit( _
        "cmd /c cd /d """ & caminhoProjeto & _
        """ && git commit -m """ & mensagem & """" _
    )

End Function


'========================================================
' PUSH
'========================================================

Public Function GitPush() As String

    Dim caminhoProjeto As String
    
    caminhoProjeto = ThisWorkbook.Path
    
    GitPush = ExecutarGit( _
        "cmd /c cd /d """ & caminhoProjeto & _
        """ && git push origin main" _
    )

End Function


Public Sub TestarGitStatus()

    Dim resultado As String
    
    resultado = GitStatus()
    
    If Trim$(resultado) = vbNullString Then
        resultado = "Nenhuma alteração detectada."
    End If
    
    MsgBox resultado, _
           vbInformation, _
           "Git - Status"

End Sub

Public Sub TestarGitAdd()

    Dim resultado As String
    
    resultado = GitAdd()
    
    If Trim$(resultado) = vbNullString Then
        resultado = "Git add executado com sucesso."
    End If
    
    MsgBox resultado, _
           vbInformation, _
           "Git - Add"

End Sub

Public Sub TesteSincronizacao()

    MsgBox "Teste de sincronização do Git.", _
           vbInformation, _
           "Prime Fitness"

End Sub
