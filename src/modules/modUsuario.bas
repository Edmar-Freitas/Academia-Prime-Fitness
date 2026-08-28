Attribute VB_Name = "modUsuario"
Public Sub DefinirSenha(ByVal usuario As clsUsuario, _
                        ByVal senha As String)

    If senha = vbNullString Then
        Err.Raise vbObjectError + 1200, _
                  "DefinirSenha", _
                  "A senha não pode estar vazia."
    End If

    usuario.salt = GerarSalt()

    usuario.senhaHash = GerarHashSenha( _
                            senha, _
                            usuario.salt)

End Sub

