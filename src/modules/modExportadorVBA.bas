Attribute VB_Name = "modExportadorVBA"
Option Explicit

Public Sub ExportarProjetoVBA()

    Dim projeto As Object
    Dim componente As Object
    
    Dim pastaBase As String
    Dim pastaClasses As String
    Dim pastaModulos As String
    Dim pastaForms As String
    
    Dim extensao As String
    Dim caminhoArquivo As String
    
    Dim quantidade As Long
    
    '--------------------------------------------------
    ' Validação
    '--------------------------------------------------
    
    If ThisWorkbook.Path = vbNullString Then
        MsgBox "O arquivo precisa estar salvo antes da exportação.", _
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
    ' Criar diretórios caso não existam
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
        
            '------------------------------------------
            ' Módulos padrão
            '------------------------------------------
            
            Case 1
                
                extensao = ".bas"
                caminhoArquivo = pastaModulos & "\" & _
                                 componente.Name & extensao
                
                ExcluirArquivoSeExistir caminhoArquivo
                
                componente.Export caminhoArquivo
                
                quantidade = quantidade + 1
            
            '------------------------------------------
            ' Classes
            '------------------------------------------
            
            Case 2
                
                extensao = ".cls"
                caminhoArquivo = pastaClasses & "\" & _
                                 componente.Name & extensao
                
                ExcluirArquivoSeExistir caminhoArquivo
                
                componente.Export caminhoArquivo
                
                quantidade = quantidade + 1
            
            '------------------------------------------
            ' UserForms
            '------------------------------------------
            
            Case 3
                
                extensao = ".frm"
                caminhoArquivo = pastaForms & "\" & _
                                 componente.Name & extensao
                
                ExcluirArquivoSeExistir caminhoArquivo
                
                ExcluirArquivoSeExistir _
                    pastaForms & "\" & componente.Name & ".frx"
                
                componente.Export caminhoArquivo
                
                quantidade = quantidade + 1
            
        End Select
    
    Next componente
    
    '--------------------------------------------------
    ' Finalização
    '--------------------------------------------------
    
    MsgBox quantidade & " componentes VBA exportados com sucesso.", _
           vbInformation, _
           "Exportador VBA"
    
    Exit Sub


'======================================================
' TRATAMENTO DE ERROS
'======================================================

ErroAcesso:

    MsgBox "O Excel não permitiu acesso ao projeto VBA." & vbCrLf & vbCrLf & _
           "Verifique se a opção:" & vbCrLf & _
           "'Confiar no acesso ao modelo de objeto do projeto do VBA'" & _
           " está habilitada.", _
           vbCritical, _
           "Erro - Exportador VBA"
    
    Exit Sub


ErroExportacao:

    MsgBox "Ocorreu um erro durante a exportação." & vbCrLf & vbCrLf & _
           "Erro: " & Err.Number & vbCrLf & _
           "Descrição: " & Err.Description, _
           vbCritical, _
           "Erro - Exportador VBA"

End Sub


Private Sub CriarPasta(ByVal caminho As String)

    If Dir(caminho, vbDirectory) = vbNullString Then
        MkDir caminho
    End If

End Sub


Private Sub ExcluirArquivoSeExistir(ByVal caminho As String)

    If Dir(caminho) <> vbNullString Then
        Kill caminho
    End If

End Sub

