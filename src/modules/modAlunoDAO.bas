Attribute VB_Name = "modAlunoDAO"
Public Function ProximoCodigoAluno() As Long

    Dim ultimaLinha As Long
    
    Set repoAlunos = ThisWorkbook.Worksheets("repoAlunos")

    ultimaLinha = repoAlunos.Cells(repoAlunos.Rows.Count, 1).End(xlUp).Row

    If ultimaLinha < 2 Then
        ProximoCodigoAluno = 1
    Else
        ProximoCodigoAluno = repoAlunos.Cells(ultimaLinha, 1).Value + 1
    End If

End Function

Public Sub SalvarAluno(ByVal aluno As clsAluno)

    Dim repoAlunos As Worksheet
    Dim linha As Long
       
    Set repoAlunos = ThisWorkbook.Worksheets("repoAlunos")
   
    linha = repoAlunos.Cells(repoAlunos.Rows.Count, 1).End(xlUp).Row
    
    If repoAlunos.Cells(2, 1) = "" Then
        linha = 2
    Else
        linha = linha + 1
    End If

    repoAlunos.Cells(linha, 1).Value = aluno.codigo
    repoAlunos.Cells(linha, 2).Value = aluno.nome
    repoAlunos.Cells(linha, 3).Value = aluno.DataNascimento
    repoAlunos.Cells(linha, 4).Value = aluno.Genero
    repoAlunos.Cells(linha, 5).Value = aluno.Telefone
    repoAlunos.Cells(linha, 6).Value = aluno.TelefoneEmergencial
    repoAlunos.Cells(linha, 7).Value = aluno.Email
    repoAlunos.Cells(linha, 8).Value = aluno.Cidade
    repoAlunos.Cells(linha, 9).Value = aluno.Bairro
    repoAlunos.Cells(linha, 10).Value = aluno.Logradouro
    repoAlunos.Cells(linha, 11).Value = aluno.Numero
    repoAlunos.Cells(linha, 12).Value = aluno.CEP
    repoAlunos.Cells(linha, 13).Value = aluno.RG
    repoAlunos.Cells(linha, 14).Value = aluno.CPF
    repoAlunos.Cells(linha, 15).Value = aluno.FotoCaminho
    repoAlunos.Cells(linha, 16).Value = aluno.Observacao
    repoAlunos.Cells(linha, 17).Value = aluno.Situacao
 
End Sub

Public Sub Excluir(ByVal codigo As Long)

    Dim repoAlunos As Worksheet
    Dim cel As Range

    Set repoAlunos = ThisWorkbook.Worksheets("repoAlunos")

    Set cel = repoAlunos.Columns(1).Find( _
        What:=codigo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If cel Is Nothing Then
        Exit Sub
    End If

    repoAlunos.Rows(cel.Row).Delete

End Sub

Public Function Atualizar(ByVal aluno As clsAluno) As Boolean

    Dim repoAlunos As Worksheet
    Dim cel As Range
    Dim linha As Long

    Set repoAlunos = ThisWorkbook.Worksheets("repoAlunos")

    Set cel = repoAlunos.Columns(1).Find( _
        What:=aluno.codigo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If cel Is Nothing Then
        Atualizar = False
        Exit Function
    End If

    linha = cel.Row

    repoAlunos.Cells(linha, 1).Value = aluno.codigo
    repoAlunos.Cells(linha, 2).Value = aluno.nome
    repoAlunos.Cells(linha, 3).Value = aluno.DataNascimento
    repoAlunos.Cells(linha, 4).Value = aluno.Genero
    repoAlunos.Cells(linha, 5).Value = aluno.Telefone
    repoAlunos.Cells(linha, 6).Value = aluno.TelefoneEmergencial
    repoAlunos.Cells(linha, 7).Value = aluno.Email
    repoAlunos.Cells(linha, 8).Value = aluno.Cidade
    repoAlunos.Cells(linha, 9).Value = aluno.Bairro
    repoAlunos.Cells(linha, 10).Value = aluno.Logradouro
    repoAlunos.Cells(linha, 11).Value = aluno.Numero
    repoAlunos.Cells(linha, 12).Value = aluno.CEP
    repoAlunos.Cells(linha, 13).Value = aluno.RG
    repoAlunos.Cells(linha, 14).Value = aluno.CPF
    repoAlunos.Cells(linha, 15).Value = aluno.FotoCaminho
    repoAlunos.Cells(linha, 16).Value = aluno.Observacao
    repoAlunos.Cells(linha, 17).Value = aluno.Situacao

    Atualizar = True

End Function


Public Function BuscarPorCodigo(ByVal codigo As Long) As clsAluno

    Dim repoAlunos As Worksheet
    Dim cel As Range
    Dim aluno As clsAluno

    Set repoAlunos = ThisWorkbook.Worksheets("repoAlunos")

    Set cel = repoAlunos.Columns(1).Find( _
        What:=codigo, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    If cel Is Nothing Then
        Set BuscarPorCodigo = Nothing
        Exit Function
    End If

    Set aluno = New clsAluno

    With aluno

        .codigo = repoAlunos.Cells(cel.Row, 1).Value
        .nome = repoAlunos.Cells(cel.Row, 2).Value
        .DataNascimento = repoAlunos.Cells(cel.Row, 3).Value
        .Genero = repoAlunos.Cells(cel.Row, 4).Value
        .Telefone = repoAlunos.Cells(cel.Row, 5).Value
        .TelefoneEmergencial = repoAlunos.Cells(cel.Row, 6).Value
        .Email = repoAlunos.Cells(cel.Row, 7).Value
        .Cidade = repoAlunos.Cells(cel.Row, 8).Value
        .Bairro = repoAlunos.Cells(cel.Row, 9).Value
        .Logradouro = repoAlunos.Cells(cel.Row, 10).Value
        .Numero = repoAlunos.Cells(cel.Row, 11).Value
        .CEP = repoAlunos.Cells(cel.Row, 12).Value
        .RG = repoAlunos.Cells(cel.Row, 13).Value
        .CPF = repoAlunos.Cells(cel.Row, 14).Value
        .FotoCaminho = repoAlunos.Cells(cel.Row, 15).Value
        .Observacao = repoAlunos.Cells(cel.Row, 16).Value
        .Situacao = repoAlunos.Cells(cel.Row, 17).Value

    End With

    Set BuscarPorCodigo = aluno

End Function

