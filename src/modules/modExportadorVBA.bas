Attribute VB_Name = "modExportadorVBA"
Option Explicit

Public Sub ExportarProjetoVBA()

    Dim projeto As Object
    Dim componente As Object

    Dim pastaBase As String
    Dim pastaClasses As String
    Dim pastaModulos As String
    Dim pastaForms As String

    Dim caminhoArquivo As String
    Dim caminhoFrx As String
    Dim caminhoFrxBackup As String

    Dim quantidade As Long
    Dim quantidadeClasses As Long
    Dim quantidadeModulos As Long
    Dim quantidadeForms As Long

    '--------------------------------------------------
    ' Validação
    '--------------------------------------------------

    If ThisWorkbook.Path = vbNullString Then

        MsgBox _
            "O arquivo precisa estar salvo antes da exportação.", _
            vbExclamation, _
            "Exportador VBA"

        Exit Sub

    End If

    '--------------------------------------------------
    ' Diretórios
    '--------------------------------------------------

    pastaBase = ThisWorkbook.Path

    pastaClasses = pastaBase & "\src\classes"
    pastaModulos = pastaBase & "\src\modules"
    pastaForms = pastaBase & "\src\forms"

    '--------------------------------------------------
    ' Criar diretórios
    '--------------------------------------------------

    CriarPasta pastaBase & "\src"
    CriarPasta pastaClasses
    CriarPasta pastaModulos
    CriarPasta pastaForms

    '--------------------------------------------------
    ' Acesso ao projeto VBA
    '--------------------------------------------------

    On Error GoTo ErroAcesso

    Set projeto = ThisWorkbook.VBProject

    On Error GoTo ErroExportacao

    '--------------------------------------------------
    ' Percorrer componentes
    '--------------------------------------------------

    For Each componente In projeto.VBComponents

        Select Case componente.Type

            '==========================================
            ' MÓDULO PADRÃO
            '==========================================

            Case 1

                caminhoArquivo = _
                    pastaModulos & "\" & _
                    componente.Name & ".bas"

                ExportarComponente _
                    componente, _
                    caminhoArquivo

                quantidadeModulos = _
                    quantidadeModulos + 1


            '==========================================
            ' CLASSE
            '==========================================

            Case 2

                caminhoArquivo = _
                    pastaClasses & "\" & _
                    componente.Name & ".cls"

                ExportarComponente _
                    componente, _
                    caminhoArquivo

                quantidadeClasses = _
                    quantidadeClasses + 1


            '==========================================
            ' USERFORM
            '==========================================

            Case 3

                caminhoArquivo = _
                    pastaForms & "\" & _
                    componente.Name & ".frm"

                caminhoFrx = _
                    pastaForms & "\" & _
                    componente.Name & ".frx"

                caminhoFrxBackup = _
                    pastaForms & "\" & _
                    componente.Name & ".frx.backup"

                '--------------------------------------
                ' Fazer backup do .frx existente
                '--------------------------------------

                BackupFrx _
                    caminhoFrx, _
                    caminhoFrxBackup

                '--------------------------------------
                ' Exportar UserForm
                '--------------------------------------

                ExportarComponente _
                    componente, _
                    caminhoArquivo

                '--------------------------------------
                ' Restaurar .frx original
                '--------------------------------------

                RestaurarFrx _
                    caminhoFrx, _
                    caminhoFrxBackup

                quantidadeForms = _
                    quantidadeForms + 1

        End Select

    Next componente

    '--------------------------------------------------
    ' Quantidade total
    '--------------------------------------------------

    quantidade = _
        quantidadeClasses + _
        quantidadeModulos + _
        quantidadeForms

    '--------------------------------------------------
    ' Finalização
    '--------------------------------------------------

    MsgBox _
        "Exportação VBA concluída." & vbCrLf & vbCrLf & _
        "Classes: " & quantidadeClasses & vbCrLf & _
        "Módulos: " & quantidadeModulos & vbCrLf & _
        "UserForms: " & quantidadeForms & vbCrLf & _
        "Total: " & quantidade, _
        vbInformation, _
        "Exportador VBA"

    Exit Sub


'======================================================
' TRATAMENTO DE ERROS
'======================================================

ErroAcesso:

    MsgBox _
        "O Excel não permitiu acesso ao projeto VBA." & _
        vbCrLf & vbCrLf & _
        "Verifique se a opção:" & vbCrLf & _
        "'Confiar no acesso ao modelo de objeto do projeto do VBA'" & _
        " está habilitada.", _
        vbCritical, _
        "Erro - Exportador VBA"

    Exit Sub


ErroExportacao:

    MsgBox _
        "Ocorreu um erro durante a exportação." & _
        vbCrLf & vbCrLf & _
        "Erro: " & Err.Number & vbCrLf & _
        "Descrição: " & Err.Description, _
        vbCritical, _
        "Erro - Exportador VBA"

End Sub


'======================================================
' EXPORTAR COMPONENTE
'======================================================

Private Sub ExportarComponente( _
    ByVal componente As Object, _
    ByVal caminhoArquivo As String)

    componente.Export caminhoArquivo

End Sub


'======================================================
' BACKUP DO FRX
'======================================================

Private Sub BackupFrx( _
    ByVal caminhoFrx As String, _
    ByVal caminhoBackup As String)

    'Remove backup antigo

    If Dir(caminhoBackup) <> vbNullString Then
        Kill caminhoBackup
    End If

    'Se existir .frx, copia para backup

    If Dir(caminhoFrx) <> vbNullString Then

        FileCopy _
            caminhoFrx, _
            caminhoBackup

    End If

End Sub


'======================================================
' RESTAURAR FRX
'======================================================

Private Sub RestaurarFrx( _
    ByVal caminhoFrx As String, _
    ByVal caminhoBackup As String)

    'Remove o .frx gerado pelo Excel

    If Dir(caminhoFrx) <> vbNullString Then
        Kill caminhoFrx
    End If

    'Se existe backup, restaura

    If Dir(caminhoBackup) <> vbNullString Then

        FileCopy _
            caminhoBackup, _
            caminhoFrx

        Kill caminhoBackup

    End If

End Sub


'======================================================
' CRIAR PASTA
'======================================================

Private Sub CriarPasta(ByVal caminho As String)

    If Dir(caminho, vbDirectory) = vbNullString Then
        MkDir caminho
    End If

End Sub

