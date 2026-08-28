Attribute VB_Name = "modFuncionarioDAO"
Public Function ProximoCodigoFuncionario() As Long

    Dim ultimaLinha As Long
    
    Set repoFuncionarios = ThisWorkbook.Worksheets("repoFuncionarios")

    ultimaLinha = repoFuncionarios.Cells(repoFuncionarios.Rows.Count, 1).End(xlUp).Row

    If ultimaLinha < 2 Then
        ProximoCodigoFuncionario = 1
    Else
        ProximoCodigoFuncionario = repoFuncionarios.Cells(ultimaLinha, 1).Value + 1
    End If

End Function

Public Sub SalvarFuncionario(ByVal funcionario As clsFuncionario)

    Dim repoFuncionarios As Worksheet
    Dim linha As Long
       
    Set repoFuncionarios = ThisWorkbook.Worksheets("repoFuncionarios")
   
    linha = repoFuncionarios.Cells(repoFuncionarios.Rows.Count, 1).End(xlUp).Row
    
    If repoFuncionarios.Cells(2, 1) = "" Then
        linha = 2
    Else
        linha = linha + 1
    End If

    repoFuncionarios.Cells(linha, 1).Value = funcionario.codigo
    repoFuncionarios.Cells(linha, 2).Value = funcionario.nome
    repoFuncionarios.Cells(linha, 3).Value = funcionario.DataNascimento
    repoFuncionarios.Cells(linha, 4).Value = funcionario.Genero
    repoFuncionarios.Cells(linha, 5).Value = funcionario.Telefone
    repoFuncionarios.Cells(linha, 6).Value = funcionario.TelefoneEmergencial
    repoFuncionarios.Cells(linha, 7).Value = funcionario.Email
    repoFuncionarios.Cells(linha, 8).Value = funcionario.Cidade
    repoFuncionarios.Cells(linha, 9).Value = funcionario.Bairro
    repoFuncionarios.Cells(linha, 10).Value = funcionario.Logradouro
    repoFuncionarios.Cells(linha, 11).Value = funcionario.Numero
    repoFuncionarios.Cells(linha, 12).Value = funcionario.CEP
    repoFuncionarios.Cells(linha, 13).Value = funcionario.RG
    repoFuncionarios.Cells(linha, 14).Value = funcionario.CPF
    repoFuncionarios.Cells(linha, 15).Value = funcionario.CREF
    repoFuncionarios.Cells(linha, 16).Value = funcionario.FotoCaminho
    repoFuncionarios.Cells(linha, 17).Value = funcionario.Observacao
    repoFuncionarios.Cells(linha, 18).Value = funcionario.Funcao
    repoFuncionarios.Cells(linha, 19).Value = funcionario.Situacao
 
End Sub

Public Sub Excluir(ByVal codigo As Long)

    Dim repoFuncionarios As Worksheet
    Dim cel As Range

    Set repoFuncionarios = ThisWorkbook.Worksheets("repoFuncionarios")

    Set cel = repoFuncionarios.Columns(1).Find( _
        What:=codigo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If cel Is Nothing Then
        Exit Sub
    End If

    repoFuncionarios.Rows(cel.Row).Delete

End Sub

Public Function Atualizar(ByVal funcionario As clsFuncionario) As Boolean

    Dim repoFuncionarios As Worksheet
    Dim cel As Range
    Dim linha As Long

    Set repoFuncionarios = ThisWorkbook.Worksheets("repoFuncionarios")

    Set cel = repoFuncionarios.Columns(1).Find( _
        What:=funcionario.codigo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If cel Is Nothing Then
        Atualizar = False
        Exit Function
    End If

    linha = cel.Row

    repoFuncionarios.Cells(linha, 1).Value = funcionario.codigo
    repoFuncionarios.Cells(linha, 2).Value = funcionario.nome
    repoFuncionarios.Cells(linha, 3).Value = funcionario.DataNascimento
    repoFuncionarios.Cells(linha, 4).Value = funcionario.Genero
    repoFuncionarios.Cells(linha, 5).Value = funcionario.Telefone
    repoFuncionarios.Cells(linha, 6).Value = funcionario.TelefoneEmergencial
    repoFuncionarios.Cells(linha, 7).Value = funcionario.Email
    repoFuncionarios.Cells(linha, 8).Value = funcionario.Cidade
    repoFuncionarios.Cells(linha, 9).Value = funcionario.Bairro
    repoFuncionarios.Cells(linha, 10).Value = funcionario.Logradouro
    repoFuncionarios.Cells(linha, 11).Value = funcionario.Numero
    repoFuncionarios.Cells(linha, 12).Value = funcionario.CEP
    repoFuncionarios.Cells(linha, 13).Value = funcionario.RG
    repoFuncionarios.Cells(linha, 14).Value = funcionario.CPF
    repoFuncionarios.Cells(linha, 15).Value = funcionario.CREF
    repoFuncionarios.Cells(linha, 16).Value = funcionario.FotoCaminho
    repoFuncionarios.Cells(linha, 17).Value = funcionario.Observacao
    repoFuncionarios.Cells(linha, 18).Value = funcionario.Funcao
    repoFuncionarios.Cells(linha, 19).Value = funcionario.Situacao

    Atualizar = True

End Function


Public Function BuscarPorCodigo(ByVal codigo As Long) As clsFuncionario

    Dim repoFuncionarios As Worksheet
    Dim cel As Range
    Dim funcionario As clsFuncionario

    Set repoFuncionarios = ThisWorkbook.Worksheets("repoFuncionarios")

    Set cel = repoFuncionarios.Columns(1).Find( _
        What:=codigo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If cel Is Nothing Then
        Set BuscarPorCodigo = Nothing
        Exit Function
    End If

    Set funcionario = New clsFuncionario

    With funcionario

        .codigo = repoFuncionarios.Cells(cel.Row, 1).Value
        .nome = repoFuncionarios.Cells(cel.Row, 2).Value
        .DataNascimento = repoFuncionarios.Cells(cel.Row, 3).Value
        .Genero = repoFuncionarios.Cells(cel.Row, 4).Value
        .Telefone = repoFuncionarios.Cells(cel.Row, 5).Value
        .TelefoneEmergencial = repoFuncionarios.Cells(cel.Row, 6).Value
        .Email = repoFuncionarios.Cells(cel.Row, 7).Value
        .Cidade = repoFuncionarios.Cells(cel.Row, 8).Value
        .Bairro = repoFuncionarios.Cells(cel.Row, 9).Value
        .Logradouro = repoFuncionarios.Cells(cel.Row, 10).Value
        .Numero = repoFuncionarios.Cells(cel.Row, 11).Value
        .CEP = repoFuncionarios.Cells(cel.Row, 12).Value
        .RG = repoFuncionarios.Cells(cel.Row, 13).Value
        .CPF = repoFuncionarios.Cells(cel.Row, 14).Value
        .CREF = repoFuncionarios.Cells(cel.Row, 15).Value
        .FotoCaminho = repoFuncionarios.Cells(cel.Row, 16).Value
        .Observacao = repoFuncionarios.Cells(cel.Row, 17).Value
        .Funcao = repoFuncionarios.Cells(cel.Row, 18).Value
        .Situacao = repoFuncionarios.Cells(cel.Row, 19).Value

    End With

    Set BuscarPorCodigo = funcionario

End Function



