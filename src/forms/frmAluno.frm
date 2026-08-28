VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAluno 
   Caption         =   "Prime Fitness - CADASTRO DE ALUNOS"
   ClientHeight    =   10272
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13248
   OleObjectBlob   =   "frmAluno.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAluno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAtualizar_Click()

    Dim aluno As clsAluno

    Set aluno = New clsAluno
    
    If Me.txtCodigo.Value = "" Then
        MsgBox "Escolha um registro para Atualizar!", vbInformation
    Else
        With aluno
            .codigo = CLng(Me.txtCodigo.Value)
            .nome = Me.txtNomeCompleto.Value
            .DataNascimento = CDate(Me.txtDataNascimento.Value)
            .Genero = Me.cmbGenero.Value
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
            .FotoCaminho = Me.txtFotoCaminho.Value
            .Observacao = Me.txtObservacao.Value
            .Situacao = Me.cmbSituacao.Value
            
            If modAluno.AtualizarAluno(aluno) Then
                CarregarListaAlunos
                MsgBox "Aluno atualizado com sucesso.", vbInformation
            Else
                MsgBox "Não foi possível atualizar o aluno!", vbExclamation
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

    If Me.lstAlunos.ListIndex = -1 Then
        MsgBox "Selecione um aluno.", vbExclamation
        Exit Sub
    End If

    codigo = CLng(Me.lstAlunos.List(Me.lstAlunos.ListIndex, 0))

    If MsgBox("Deseja realmente excluir este aluno?", _
              vbQuestion + vbYesNo, _
              "Excluir aluno") = vbNo Then
        Exit Sub
    End If

    ExcluirAluno codigo

    CarregarListaAlunos
    
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

            frmAluno.txtFotoCaminho.Value = .SelectedItems(1)

            frmAluno.imgFoto.Picture = LoadPicture(txtFotoCaminho.Value)

        End If

    End With
End Sub

Private Sub cmdNovo_Click()
    LimparFormulario
    NovoAluno
    txtNomeCompleto.SetFocus
End Sub

Private Sub cmdNovo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovo.ForeColor = &H80000012
End Sub


Private Sub cmdPesquisar_Click()

    Dim codigo As Long
    Dim i As Long

    If Trim(Me.txtCodigo.Value) = "" Then
        MsgBox "Informe o código do aluno.", vbExclamation
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
    For i = 0 To Me.lstAlunos.ListCount - 1

        If CLng(Me.lstAlunos.List(i, 0)) = codigo Then

            'Seleciona o registro
            Me.lstAlunos.ListIndex = i

            Exit Sub

        End If

    Next i

    MsgBox "Aluno não encontrado!", vbInformation

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
    Dim aluno As New clsAluno
    
    If Me.txtCodigo.Value = "" Then
        MsgBox "Click em *Novo* para gerar um código!", vbInformation
    Else
        aluno.codigo = CLng(Me.txtCodigo.Value)
        aluno.nome = Me.txtNomeCompleto.Value
        aluno.DataNascimento = CDate(Me.txtDataNascimento.Value)
        aluno.Genero = Me.cmbGenero.Value
        aluno.RG = Me.txtRG.Value
        aluno.CPF = Me.txtCPF.Value
        aluno.Cidade = Me.txtCidade.Value
        aluno.Bairro = Me.txtBairro.Value
        aluno.Logradouro = Me.txtLogradouro.Value
        aluno.Numero = Me.txtNumero.Value
        aluno.CEP = Me.txtCEP.Value
        aluno.Telefone = Me.txtTelefone.Value
        aluno.TelefoneEmergencial = Me.txtTelefoneEmergencial.Value
        aluno.Email = Me.txtEmail.Value
        aluno.Observacao = Me.txtObservacao.Value
        aluno.Situacao = Me.cmbSituacao.Value
        aluno.FotoCaminho = Me.txtFotoCaminho.Value
        
        CadastrarAluno aluno
        CarregarListaAlunos
        LimparFormulario
    End If

End Sub

Private Sub cmdSalvar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdSalvar.ForeColor = &H80000012
End Sub

Private Sub fmeControle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    cmdNovo.ForeColor = &H8000000D
    cmdSalvar.ForeColor = &H8000000D
    cmdExcluir.ForeColor = &H8000000D
    cmdAtualizar.ForeColor = &H8000000D
    cmdPesquisar.ForeColor = &H8000000D
    cmdSair.ForeColor = &H8000000D
    cmdAvaliacaoFisica.ForeColor = &H8000000D
End Sub

Private Sub LimparFormulario()
    With frmAluno
        .txtCodigo = ""
        .txtNomeCompleto = ""
        .cmbGenero = ""
        .txtDataNascimento = ""
        .txtFotoCaminho = ""
        .cmbSituacao.ListIndex = 0
        .txtRG = ""
        .txtCPF = ""
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


Private Sub lstAlunos_Change()

    Dim codigo As Long
    Dim aluno As clsAluno

    If Me.lstAlunos.ListIndex = -1 Then Exit Sub

    codigo = CLng(Me.lstAlunos.List(Me.lstAlunos.ListIndex, 0))

    Set aluno = modAluno.CarregarAluno(codigo)

    ExibirAluno aluno

End Sub


Private Sub txtDataNascimento_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Set FormDestino = Me
    Set ControleDestino = Me.txtDataNascimento

    frmCalendario.Show
End Sub

Private Sub UserForm_Activate()
    LimparFormulario
    txtFotoCaminho.Visible = False
    CarregarListaAlunos
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
    CarregarListaAlunos
End Sub

Private Sub CarregarListaAlunos()

    Dim repoAlunos As Worksheet
    Dim ultimaLinha As Long
    Dim dados As Variant

    Set repoAlunos = ThisWorkbook.Worksheets("repoAlunos")

    ultimaLinha = repoAlunos.Cells(repoAlunos.Rows.Count, 1).End(xlUp).Row

    If ultimaLinha < 2 Then
        Me.lstAlunos.Clear
        Exit Sub
    End If

    dados = repoAlunos.Range("A2:R" & ultimaLinha).Value

    Me.lstAlunos.Clear

    With Me.lstAlunos

        .ColumnCount = 4
        .ColumnWidths = "42 pt; 306 pt; 84 pt; 102 pt"

        Dim i As Long

        For i = 1 To UBound(dados, 1)

            .AddItem dados(i, 1)                    'Código
            .List(.ListCount - 1, 1) = dados(i, 2)  'Nome
            .List(.ListCount - 1, 2) = dados(i, 5)  'Telefone
            .List(.ListCount - 1, 3) = dados(i, 17) 'Situação

        Next i

    End With

End Sub

Private Sub ExibirAluno(ByVal aluno As clsAluno)

    If aluno Is Nothing Then
        MsgBox "Aluno não encontrado!", vbExclamation
        Exit Sub
    End If

    With Me

        .txtCodigo.Value = aluno.codigo
        .txtNomeCompleto.Value = aluno.nome
        .txtDataNascimento.Value = aluno.DataNascimento
        .cmbGenero.Value = aluno.Genero
        .txtTelefone.Value = aluno.Telefone
        .txtTelefoneEmergencial.Value = aluno.TelefoneEmergencial
        .txtEmail.Value = aluno.Email
        .txtCidade.Value = aluno.Cidade
        .txtBairro.Value = aluno.Bairro
        .txtLogradouro.Value = aluno.Logradouro
        .txtNumero.Value = aluno.Numero
        .txtCEP.Value = aluno.CEP
        .txtRG.Value = aluno.RG
        .txtCPF.Value = aluno.CPF
        .txtFotoCaminho.Value = aluno.FotoCaminho
        .txtObservacao.Value = aluno.Observacao
        .cmbSituacao.Value = aluno.Situacao

    End With

    CarregarFoto aluno.FotoCaminho

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
