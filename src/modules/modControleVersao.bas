Attribute VB_Name = "modControleVersao"
Option Explicit

'========================================================
' PREPARAR E VERSIONAR O PROJETO
'========================================================

Public Sub SincronizarLocal()

    Dim resultado As String
    Dim codigo As Long
    Dim mensagem As String

    '----------------------------------------------------
    ' 1. Salvar o arquivo Excel
    '----------------------------------------------------

    ThisWorkbook.Save

    '----------------------------------------------------
    ' 2. Exportar o projeto VBA
    '----------------------------------------------------

    ExportarProjetoVBA

    '----------------------------------------------------
    ' 3. Verificar alterações
    '----------------------------------------------------

    resultado = GitStatus(codigo)

    If codigo <> 0 Then

        MsgBox _
            "Erro ao consultar o Git." & _
            vbCrLf & vbCrLf & _
            resultado, _
            vbCritical, _
            "Prime Fitness - Git"

        Exit Sub

    End If

    '----------------------------------------------------
    ' 4. Nenhuma alteração
    '----------------------------------------------------

    If Len(Trim$(resultado)) = 0 Then

        MsgBox _
            "Nenhuma alteração foi detectada." & _
            vbCrLf & vbCrLf & _
            "O repositório já está atualizado.", _
            vbInformation, _
            "Prime Fitness - Git"

        Exit Sub

    End If

    '----------------------------------------------------
    ' 5. Adicionar alterações
    '----------------------------------------------------

    resultado = GitAdd(codigo)

    If codigo <> 0 Then

        MsgBox _
            "Erro ao executar Git Add." & _
            vbCrLf & vbCrLf & _
            resultado, _
            vbCritical, _
            "Prime Fitness - Git"

        Exit Sub

    End If

    '----------------------------------------------------
    ' 6. Mensagem do commit
    '----------------------------------------------------

    mensagem = _
        "chore: atualizar projeto VBA"

    '----------------------------------------------------
    ' 7. Criar commit
    '----------------------------------------------------

    resultado = GitCommit(mensagem, codigo)

    If codigo <> 0 Then

        MsgBox _
            "Erro ao criar o commit." & _
            vbCrLf & vbCrLf & _
            resultado, _
            vbCritical, _
            "Prime Fitness - Git"

        Exit Sub

    End If

    '----------------------------------------------------
    ' 8. Sucesso
    '----------------------------------------------------

    MsgBox _
        "Alterações versionadas com sucesso." & _
        vbCrLf & vbCrLf & _
        "Commit:" & vbCrLf & _
        mensagem & _
        vbCrLf & vbCrLf & _
        "O GitHub ainda não foi atualizado.", _
        vbInformation, _
        "Prime Fitness - Git"

End Sub
