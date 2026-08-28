VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFuncionario 
   Caption         =   "Prime Fitness - CADASTRO DE FUNCIONÁRIOS"
   ClientHeight    =   10320
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13272
   OleObjectBlob   =   "frmFuncionario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualizar_Click()

    Dim funcionario As clsFuncionario

    Set funcionario = New clsFuncionario
    
    If Me.txtCodigo.Value = "" Then
        MsgBox "Escolha um registro para Atualizar!", vbInformation
    Else
        With funcionario
            .codigo = CLng(Me.txtCodigo.Value)
            .nome = Me.txtNomeCompleto.Value
            .DataNascimento = CDate(Me.txtDataNascimento.Value)
            .Genero = Me.cmbGenero.Value
            .Funcao = Me.cmbFuncao.Value
            .Telefone = Me.txtTelefone.Value
            .TelefoneEmergencial = Me.txtTelefoneEmergencial.Value
            .Email = Me.txtEmail.Value
            .Cidade = Me.txtCidade.Value
            .Bairro = Me.txtBairro.Value
            .Logradouro = Me.txtLogradouro.Value
            .Numero = Me.txtNumero.Value
            .CEP = Me.txtCEP.Value
            .RG = Me.txtRG.Value
            .CPF = Me.txtCPF.Value
            .CREF = Me.txtCREF.Value
            .FotoCaminho = Me.txtFotoCaminho.Value
            .Observacao = Me.txtObservacao.Value
            .Situacao = Me.cmbSituacao.Value
            
            If modFuncionario.AtualizarFuncionario(funcionario) Then
                CarregarListaFuncionarios
                MsgBox "Funcionario atualizado com sucesso.", vbInformation
            Else
                MsgBox "Não foi possível atualizar o Funcionario!", vbExclamation
            End If
    
        End With
    End If
    
End Sub

Private Sub cmdAtualizar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdAtualizar.ForeColor = &H80000012
End Sub

Private Sub cmdAvaliacaoFisica_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdAvaliacaoFisica.ForeColor = &H80000012
End Sub

Private Sub cmdCalendario_Click()
    Set FormDestino = Me
    Set ControleDestino = Me.txtDataNascimento
    frmCalendario.Show
End Sub

Private Sub cmdExcluir_Click()

    Dim codigo As Long

    If Me.lstFuncionarios.ListIndex = -1 Then
        MsgBox "Selecione um Funcionário.", vbExclamation
        Exit Sub
    End If

    codigo = CLng(Me.lstFuncionarios.List(Me.lstFuncionarios.ListIndex, 0))

    If MsgBox("Deseja realmente excluir este Funcionário?", _
              vbQuestion + vbYesNo, _
              "Excluir Funcionário") = vbNo Then
        Exit Sub
    End If

    ExcluirFuncionario codigo

    CarregarListaFuncionarios
    
    LimparFormulario

End Sub

Private Sub cmdExcluir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
     cmdExcluir.ForeColor = &H80000012
End Sub


Private Sub cmdFoto_Click()
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd

        .Title = "Selecione uma foto"

        .Filters.Clear
        .Filters.Add "Imagens", "*.jpg;*.jpeg;*.png;*.bmp"

        If .Show = -1 Then

            frmFuncionario.txtFotoCaminho.Value = .SelectedItems(1)

            frmFuncionario.imgFoto.Picture = LoadPicture(txtFotoCaminho.Value)

        End If

    End With
End Sub

Private Sub cmdNovo_Click()
    LimparFormulario
    NovoFuncionario
    txtNomeCompleto.SetFocus
End Sub

Private Sub cmdNovo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovo.ForeColor = &H80000012
End Sub


Private Sub cmdPesquisar_Click()

    Dim codigo As Long
    Dim i As Long

    If Trim(Me.txtCodigo.Value) = "" Then
        MsgBox "Informe o código do Funcionario.", vbExclamation
        Me.txtCodigo.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(Me.txtCodigo.Value) Then
        MsgBox "O código deve ser numérico!", vbExclamation
        Me.txtCodigo.SetFocus
        Exit Sub
    End If

    codigo = CLng(Me.txtCodigo.Value)

    'Procura o código na ListBox
    For i = 0 To Me.lstFuncionarios.ListCount - 1

        If CLng(Me.lstFuncionarios.List(i, 0)) = codigo Then

            'Seleciona o registro
            Me.lstFuncionarios.ListIndex = i

            Exit Sub

        End If

    Next i

    MsgBox "Funcionário não encontrado!", vbInformation

End Sub

Private Sub cmdPesquisar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdPesquisar.ForeColor = &H80000012
End Sub


Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub cmdSair_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdSair.ForeColor = &H80000012
End Sub

Private Sub cmdSalvar_Click()
    Dim funcionario As New clsFuncionario
    
    If Me.txtCodigo.Value = "" Then
        MsgBox "Click em *Novo* para gerar um código!", vbInformation
    Else
        funcionario.codigo = CLng(Me.txtCodigo.Value)
        funcionario.nome = Me.txtNomeCompleto.Value
        funcionario.DataNascimento = CDate(Me.txtDataNascimento.Value)
        funcionario.Genero = Me.cmbGenero.Value
        funcionario.Funcao = Me.cmbFuncao.Value
        funcionario.RG = Me.txtRG.Value
        funcionario.CPF = Me.txtCPF.Value
        funcionario.CREF = Me.txtCREF.Value
        funcionario.Cidade = Me.txtCidade.Value
        funcionario.Bairro = Me.txtBairro.Value
        funcionario.Logradouro = Me.txtLogradouro.Value
        funcionario.Numero = Me.txtNumero.Value
        funcionario.CEP = Me.txtCEP.Value
        funcionario.Telefone = Me.txtTelefone.Value
        funcionario.TelefoneEmergencial = Me.txtTelefoneEmergencial.Value
        funcionario.Email = Me.txtEmail.Value
        funcionario.Observacao = Me.txtObservacao.Value
        funcionario.Funcao = Me.cmbFuncao.Value
        funcionario.Situacao = Me.cmbSituacao.Value
        funcionario.FotoCaminho = Me.txtFotoCaminho.Value
        
        CadastrarFuncionario funcionario
        CarregarListaFuncionarios
        LimparFormulario
    End If

End Sub

Private Sub cmdSalvar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdSalvar.ForeColor = &H80000012
End Sub

Private Sub fmeControle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovo.ForeColor = &H8000&
    cmdSalvar.ForeColor = &H8000&
    cmdExcluir.ForeColor = &H8000&
    cmdAtualizar.ForeColor = &H8000&
    cmdPesquisar.ForeColor = &H8000&
    cmdSair.ForeColor = &H8000&
    cmdAvaliacaoFisica.ForeColor = &H8000&
End Sub

Private Sub LimparFormulario()
    With frmFuncionario
        .txtCodigo = ""
        .txtNomeCompleto = ""
        .cmbGenero = ""
        .cmbFuncao = ""
        .txtDataNascimento = ""
        .txtFotoCaminho = ""
        .cmbSituacao.ListIndex = 0
        .txtRG = ""
        .txtCPF = ""
        .txtCREF = ""
        .txtCidade = ""
        .txtBairro = ""
        .txtLogradouro = ""
        .txtNumero = ""
        .txtCEP = ""
        .txtTelefone = ""
        .txtTelefoneEmergencial = ""
        .txtEmail = ""
        .txtObservacao = ""
    End With
End Sub


Private Sub lstFuncionarios_Change()

    Dim codigo As Long
    Dim funcionario As clsFuncionario

    If Me.lstFuncionarios.ListIndex = -1 Then Exit Sub

    codigo = CLng(Me.lstFuncionarios.List(Me.lstFuncionarios.ListIndex, 0))

    Set funcionario = modFuncionario.CarregarFuncionario(codigo)

    ExibirFuncionario funcionario

End Sub


Private Sub txtDataNascimento_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Set FormDestino = Me
    Set ControleDestino = Me.txtDataNascimento

    frmCalendario.Show
End Sub

Private Sub UserForm_Activate()
    LimparFormulario
    txtFotoCaminho.Visible = False
    CarregarListaFuncionarios
End Sub

Private Sub UserForm_Initialize()
    With Me.cmbGenero
        .Clear
        .AddItem "Masculino"
        .AddItem "Feminino"
        .AddItem "Outro"
    End With
    
    With Me.cmbSituacao
        .Clear
        .AddItem "Ativo"
        .AddItem "Inativo"
        .ListIndex = 0
    End With
    
    With Me.cmbFuncao
        .Clear
        .AddItem "Administrativa"
        .AddItem "Instrutor(a)"
        .AddItem "Serviços Gerais"
        .AddItem "Recepcionista"
        .ListIndex = 0
    End With
    CarregarListaFuncionarios
End Sub

Private Sub CarregarListaFuncionarios()

    Dim repoFuncionarios As Worksheet
    Dim ultimaLinha As Long
    Dim dados As Variant

    Set repoFuncionarios = ThisWorkbook.Worksheets("repoFuncionarios")

    ultimaLinha = repoFuncionarios.Cells(repoFuncionarios.Rows.Count, 1).End(xlUp).Row

    If ultimaLinha < 2 Then
        Me.lstFuncionarios.Clear
        Exit Sub
    End If

    dados = repoFuncionarios.Range("A2:S" & ultimaLinha).Value

    Me.lstFuncionarios.Clear

    With Me.lstFuncionarios

        .ColumnCount = 4
        .ColumnWidths = "42 pt; 306 pt; 84 pt; 102 pt"

        Dim i As Long

        For i = 1 To UBound(dados, 1)

            .AddItem dados(i, 1)                    'Código
            .List(.ListCount - 1, 1) = dados(i, 2)  'Nome
            .List(.ListCount - 1, 2) = dados(i, 5)  'Telefone
            .List(.ListCount - 1, 3) = dados(i, 18) 'Função

        Next i

    End With

End Sub

Private Sub ExibirFuncionario(ByVal funcionario As clsFuncionario)

    If funcionario Is Nothing Then
        MsgBox "Funcionário não encontrado!", vbExclamation
        Exit Sub
    End If

    With Me

        .txtCodigo.Value = funcionario.codigo
        .txtNomeCompleto.Value = funcionario.nome
        .txtDataNascimento.Value = funcionario.DataNascimento
        .cmbGenero.Value = funcionario.Genero
        .cmbFuncao.Value = funcionario.Funcao
        .txtTelefone.Value = funcionario.Telefone
        .txtTelefoneEmergencial.Value = funcionario.TelefoneEmergencial
        .txtEmail.Value = funcionario.Email
        .txtCidade.Value = funcionario.Cidade
        .txtBairro.Value = funcionario.Bairro
        .txtLogradouro.Value = funcionario.Logradouro
        .txtNumero.Value = funcionario.Numero
        .txtCEP.Value = funcionario.CEP
        .txtRG.Value = funcionario.RG
        .txtCPF.Value = funcionario.CPF
        .txtCREF.Value = funcionario.CREF
        .txtFotoCaminho.Value = funcionario.FotoCaminho
        .txtObservacao.Value = funcionario.Observacao
        .cmbSituacao.Value = funcionario.Situacao

    End With

    CarregarFoto funcionario.FotoCaminho

End Sub

Private Sub CarregarFoto(ByVal caminho As String)

    If Len(Trim(caminho)) = 0 Then
        Set Me.imgFoto.Picture = Nothing
        Exit Sub
    End If

    If Dir(caminho) = "" Then
        Set Me.imgFoto.Picture = Nothing
        Exit Sub
    End If

    Me.imgFoto.Picture = LoadPicture(caminho)

End Sub


