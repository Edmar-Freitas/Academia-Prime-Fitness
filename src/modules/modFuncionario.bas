Attribute VB_Name = "modFuncionario"
Public Sub NovoFuncionario()

   frmFuncionario.txtCodigo = ProximoCodigoFuncionario

End Sub

Public Sub CadastrarFuncionario(ByVal funcionario As clsFuncionario)
    
    SalvarFuncionario funcionario

End Sub

Public Function CarregarFuncionario(ByVal codigo As Long) As clsFuncionario

    Set CarregarFuncionario = modFuncionarioDAO.BuscarPorCodigo(codigo)

End Function

Public Sub ExcluirFuncionario(ByVal codigo As Long)

    If codigo <= 0 Then
        MsgBox "Código de funcionário inválido.", vbExclamation
        Exit Sub
    End If

    modFuncionarioDAO.Excluir codigo

End Sub

Public Function AtualizarFuncionario(ByVal funcionario As clsFuncionario) As Boolean

    AtualizarFuncionario = modFuncionarioDAO.Atualizar(funcionario)

End Function
