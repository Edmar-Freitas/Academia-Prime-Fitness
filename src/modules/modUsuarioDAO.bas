Attribute VB_Name = "modUsuarioDAO"
Public Function BuscarUsuario(ByVal login As String) As clsUsuario

    Dim repoUsuarios As Worksheet
    Dim cel As Range
    Dim usuario As clsUsuario

    Set repoUsuarios = ThisWorkbook.Worksheets("repoUsuarios")

    Set cel = repoUsuarios.Columns(2).Find( _
        What:=login, _
        LookIn:=xlValues, _
        LookAt:=xlWhole)

    '---------------------------------------------
    ' Usuário não encontrado
    '---------------------------------------------

    If cel Is Nothing Then

        Set BuscarUsuario = Nothing
        Exit Function

    End If


    '---------------------------------------------
    ' Cria objeto usuário
    '---------------------------------------------

    Set usuario = New clsUsuario


    '---------------------------------------------
    ' Preenche objeto
    '---------------------------------------------

    With usuario

        .codigo = repoUsuarios.Cells(cel.Row, 1).Value

        .login = repoUsuarios.Cells(cel.Row, 2).Value

        .salt = repoUsuarios.Cells(cel.Row, 3).Value

        .senhaHash = repoUsuarios.Cells(cel.Row, 4).Value

        .nome = repoUsuarios.Cells(cel.Row, 5).Value

        .perfil = repoUsuarios.Cells(cel.Row, 6).Value

        .ativo = (UCase$(Trim$( _
                    repoUsuarios.Cells(cel.Row, 7).Value)) = "SIM")

    End With


    '---------------------------------------------
    ' Retorna objeto
    '---------------------------------------------

    Set BuscarUsuario = usuario

End Function

Public Sub SalvarUsuario(ByVal usuario As clsUsuario)

End Sub

Public Sub AtualizarUsuario(ByVal usuario As clsUsuario)

End Sub

Public Sub ExcluirUsuario(ByVal id As Long)

End Sub
