Attribute VB_Name = "modAluno"
Public Sub NovoAluno()

   frmAluno.txtCodigo = ProximoCodigoAluno

End Sub

Public Sub CadastrarAluno(ByVal aluno As clsAluno)
    
    SalvarAluno aluno

End Sub

Public Function CarregarAluno(ByVal codigo As Long) As clsAluno

    Set CarregarAluno = modAlunoDAO.BuscarPorCodigo(codigo)

End Function

Public Sub ExcluirAluno(ByVal codigo As Long)

    If codigo <= 0 Then
        MsgBox "Código de aluno inválido.", vbExclamation
        Exit Sub
    End If

    modAlunoDAO.Excluir codigo

End Sub

Public Function AtualizarAluno(ByVal aluno As clsAluno) As Boolean

    AtualizarAluno = modAlunoDAO.Atualizar(aluno)

End Function
