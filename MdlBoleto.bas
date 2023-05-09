Attribute VB_Name = "MdlBoleto"
Public ArquivoLicensa       As String
Public Email_Contato        As String
Public EmailCopia           As String
Public Tipo_Documento       As String

Public Titulosselecionados  As Integer
Public Diretorio As String 'OK
Public Arquivo As String 'OK
Public Layout As String 'OK
Public Agencia As String 'OK
'Public ContaCorrente As String 'OK
Public NomeAgencia As String 'OK
Public OutrosDadosConfiguracao1 As String 'OK
Public OutrosDadosConfiguracao2 As String 'OK
Public Instrucoes As String 'OK
Public Remessa As Boolean 'OK
Public Enviar_Email As Boolean 'OK
Public Seq As Long 'OK
Public Especie As String 'OK

Public CobreBemX As CobreBemX.ContaCorrente 'OK
Public CobreBemX1 As New ContaCorrente

Public Sub ProcGravarNumeroBoleto(IDConta As Long, IDnota As Long)
On Error GoTo tratar_erro

If Chk_novo.Value = 1 Or Chk_atualizar.Value = 1 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from tbl_Detalhes_Recebimento_Nboletos where IDContaReceber = " & IDConta & " and Nosso_numero = '" & Txt_nosso_numero & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = True Then TBAbrir.AddNew
    TBAbrir!data = Date
    TBAbrir!Responsavel = pubUsuario
    TBAbrir!IdContaReceber = IDConta
    TBAbrir!Nosso_Numero = Txt_nosso_numero
    TBAbrir!ID_nota = IDnota
    TBAbrir.Update
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaComboEmpresaBoleto()
On Error GoTo tratar_erro

With frm_Instituicoes.cmbempresa
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Empresa order by Razao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then

        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Razao) = False And TBCarregarCombo!Razao <> "" Then
                .AddItem TBCarregarCombo!Razao
                .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
End With
If TBCarregarCombo.RecordCount = 1 Then frm_Instituicoes.cmbempresa.Text = TBCarregarCombo!Razao
TBCarregarCombo.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaComboEmpresaConciliacao()
On Error GoTo tratar_erro

With Frm_InstituicoesConciliacao.cmbempresa
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from Empresa order by Razao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Razao) = False And TBCarregarCombo!Razao <> "" Then
                .AddItem TBCarregarCombo!Razao
                .ItemData(.NewIndex) = TBCarregarCombo!CODIGO
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
End With
If TBCarregarCombo.RecordCount = 1 Then
    Frm_InstituicoesConciliacao.cmbempresa.Text = TBCarregarCombo!Razao
    Frm_InstituicoesConciliacao.txtCodcedente = TBCarregarCombo!CODIGO
End If

TBCarregarCombo.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaComboCliente()
On Error GoTo tratar_erro
Dim NomeRazao As String

StrSql = "select Nome_Razao , FormaBaixa from tbl_contas_receber where LogSit = 'N' and formabaixa ='BOLETO' group by Nome_Razao, FormaBaixa ORDER BY Nome_Razao"

With frm_Instituicoes.cmbCliente
    .Clear
    .AddItem ""
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If NomeRazao <> TBCarregarCombo!Nome_Razao Then
                .AddItem TBCarregarCombo!Nome_Razao
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaComboFornecedor()
On Error GoTo tratar_erro
Dim NomeRazao As String

StrSql = "SELECT tbl_Detalhes_Recebimento.IDContaReceber, tbl_contas_receber.Nome_Razao " _
& "FROM tbl_contas_receber INNER JOIN" _
& " tbl_Detalhes_Recebimento ON tbl_contas_receber.IDIntconta = tbl_Detalhes_Recebimento.IDContaReceber" _
& " ORDER BY tbl_contas_receber.Nome_Razao"

With frm_Instituicoes.cmbCliente
    .Clear
    .AddItem ""
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Nome_Razao) = False And TBCarregarCombo!Nome_Razao <> "" Then
            If NomeRazao <> TBCarregarCombo!Nome_Razao Then
                .AddItem TBCarregarCombo!Nome_Razao
                .ItemData(.NewIndex) = TBCarregarCombo!IdContaReceber
                NomeRazao = TBCarregarCombo!Nome_Razao
            End If
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
    End If
    TBCarregarCombo.Close
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregacomboCarteira()
On Error GoTo tratar_erro

With frm_Instituicoes.cmbCarteira
    .Clear
        
Select Case frm_Instituicoes.txtNBanco
    Case "341": 'Itaú
           .AddItem "109 - Direta Eletrônica Sem Emissão - Simples":
           .AddItem "112 - Escritual Eletrônica - simples / contratual":
           .AddItem "175 - Sem Registro Sem Emissão":
           .Text = "109 - Direta Eletrônica Sem Emissão - Simples":
    Case "001": 'Banco do brasil
            .AddItem "11 - Simples - Com Registro":
            .AddItem "11 - Vinculada - Com Registro":
            .AddItem "17 - Direta Especial - Com Registro":
            .AddItem "17Simples - Direta Especial Simples - Com Registro":
            .AddItem "17-7 - Direta Especial - Com Registro Convênio 7 dígitos":
            .AddItem "18 - Simples - Sem Registro":
            .AddItem "18-7 - Simples - Sem Registro - Convênio 7 dígitos":
            .Text = "11 - Vinculada - Com Registro":
    Case "033": 'Santander
            .AddItem "COB - Cobrança Simples":
            .AddItem "COBR - Cobrança Simples - Rápida Com Registro":
            .AddItem "COBR-Nova - Cobrança Simples - Rápida Com Registro"
            .AddItem "CSR - Cobrança Simples Sem Registro":
            .AddItem "ECR - Cobrança Simples Com Registro":
            .Text = "COBR - Cobrança Simples - Rápida Com Registro":
    Case "104": 'Caixa
            .AddItem "CR - Cobrança Rápida":
            .AddItem "SR - Cobrança Sem Registro":
            .AddItem "SIG14 - SIG Com Registro - Emissão pelo Cedente":
            .Text = "SIG14 - SIG Com Registro - Emissão pelo Cedente":
    Case "237": 'Bradesco
    Case "356": 'ABN e Real
            .AddItem "20 - Cobrança Simples":
            .Text = "20 - Cobrança Simples":
    Case "399": 'HSBC
            .AddItem "CNR - Sem Registro":
            .Text = "CNR - Sem Registro":
    Case "409": 'Unibanco
            .AddItem "Especial":
            .Text = "Especial":
        End Select
End With
frm_Instituicoes.ProcBuscaArquivolicenca

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaInstrucoesBoleto()
On Error GoTo tratar_erro

With frm_Instituicoes
            .Txtpercentual_juros = ""
            .Txtpercentual_desconto = ""
            .Txtpercentual_multa = ""
            .Txtdias_protesto = ""
            .Txtinstrucoes = ""
            .txtAssunto = ""

Set TBBoleto = CreateObject("adodb.recordset")
    TBBoleto.Open "Select * from tbl_Instituicoes_Instrucoes_Boleto where ID_Instituicao = " & .txtid & "", Conexao, adOpenKeyset, adLockOptimistic
        If TBBoleto.EOF = False Then
            .Txtpercentual_juros = TBBoleto!Juros
            .Txtpercentual_desconto = TBBoleto!Desconto
            .Txtpercentual_multa = TBBoleto!Multa
            .Txtdias_protesto = TBBoleto!Dias_Protesto
            .Txtinstrucoes = TBBoleto!Instrucoes_protesto
            .txtAssunto = TBBoleto!AssuntoEmail
    TBBoleto.Close
        End If
End With
  
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub

End Sub


Public Sub ProcCarregaInstituicaoBoleto()
On Error GoTo tratar_erro

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select * from Tbl_Instituicoes where txt_descricao = '" & frm_Instituicoes.txtDescricao.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
With frm_Instituicoes
    .txtConta = IIf(IsNull(TBFIltro!txt_conta) = False, TBFIltro!txt_conta, "")
    .txtid = TBFIltro!ID
    If DS.DS_ArquivoExiste(Localrel & "\Imagens\Bancos\" & TBFIltro!int_NBanco & ".jpg") = True Then
    .Logo_Banco.Picture = LoadPicture(Localrel & "\Imagens\Bancos\" & TBFIltro!int_NBanco & ".jpg")
    End If
    .txtAgencia = IIf(IsNull(TBFIltro!txt_Agencia) = False, TBFIltro!txt_Agencia, "")
    .txtNBanco = IIf(IsNull(TBFIltro!int_NBanco) = False, TBFIltro!int_NBanco, "")
    .Txt_codigo_cedente1 = IIf(IsNull(TBFIltro!Codigo_cedente_registrado) = False, TBFIltro!Codigo_cedente_registrado, "")
    .txtcarteiraconf = ArquivoLicensa
    .Txt_nome_agencia = IIf(IsNull(TBFIltro!Nome_agencia) = False, TBFIltro!Nome_agencia, "")
    .Txtlocal = Localrel & "\Boletos\Arquivos remessa\" & frm_Instituicoes.txtDescricao.Text
End With
End If


TBFIltro.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaComboBancoConciliacao()
On Error GoTo tratar_erro

With Frm_InstituicoesConciliacao.cmbBanco
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from tbl_Instituicoes order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Txt_descricao) = False And TBCarregarCombo!Txt_descricao <> "" Then
                .AddItem TBCarregarCombo!Txt_descricao
                .ItemData(.NewIndex) = TBCarregarCombo!ID
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
        Frm_InstituicoesConciliacao.cmbBanco.Text = TBCarregarCombo!Txt_descricao
    End If
End With
TBCarregarCombo.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcCarregaComboBancoBoleto()
On Error GoTo tratar_erro

With frm_Instituicoes.cmbBanco
    .Clear
    Set TBCarregarCombo = CreateObject("adodb.recordset")
    TBCarregarCombo.Open "Select * from tbl_Instituicoes order by txt_Descricao", Conexao, adOpenKeyset, adLockOptimistic
    If TBCarregarCombo.EOF = False Then
        
        Do While TBCarregarCombo.EOF = False
            If IsNull(TBCarregarCombo!Txt_descricao) = False And TBCarregarCombo!Txt_descricao <> "" Then
                .AddItem TBCarregarCombo!Txt_descricao
                .ItemData(.NewIndex) = TBCarregarCombo!ID
            End If
            TBCarregarCombo.MoveNext
        Loop
        TBCarregarCombo.MoveFirst
        frm_Instituicoes.cmbBanco.Text = TBCarregarCombo!Txt_descricao
        ProcCarregacomboCarteira
    End If
End With
TBCarregarCombo.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaListaDuplicatas()
On Error GoTo tratar_erro
Dim SQLBusca As String

Init = 0
Sit_REG = 0
valor = 0

frm_Instituicoes.lst_Duplicata.ListItems.Clear

Set TBLISTA = CreateObject("adodb.recordset")
Debug.Print StrSql

TBLISTA.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
        With frm_Instituicoes.lst_Duplicata.ListItems
            .Add , , TBLISTA!ID
            .Item(.Count).SubItems(1) = TBLISTA!int_NotaFiscal
            .Item(.Count).SubItems(2) = TBLISTA!Nome_Razao
            .Item(.Count).SubItems(3) = Format(TBLISTA!dt_Vencimento, "dd/mm/yyyy")
            .Item(.Count).SubItems(4) = TBLISTA!txt_Parcela
            .Item(.Count).SubItems(5) = Format(TBLISTA!dbl_Valor, "###,##0.00")
            .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Nosso_Numero), "", TBLISTA!Nosso_Numero)
            
            If IsNull(TBLISTA!Seq_remessa) = False And TBLISTA!Seq_remessa <> "" And TBLISTA!txt_Portador_Banco <> "" Then
                ProcPeganumeroremessa
                .Item(.Count).SubItems(7) = Arquivo
            End If
            
            If TBLISTA!IdContaReceber = "" Then
            .Item(.Count).SubItems(8) = "Não"
            Else
            .Item(.Count).SubItems(8) = "Sim"
            End If
            
            If TBLISTA!Enviado = True Then
            .Item(.Count).SubItems(9) = "Sim"
            Else
            .Item(.Count).SubItems(9) = "Não"
            End If
            
            
            Init = Init + 1
            .Item(Init).Checked = False
        End With
                
        TBLISTA.MoveNext
        Contador = Contador + 1
    Loop

End If
TBLISTA.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcPeganumeroremessa()
On Error GoTo tratar_erro

'Verifica o último sequencial no banco para gerar o próximo
If IsNull(TBLISTA!data_envio) = True Then Exit Sub

Dia = Day(TBLISTA!data_envio)
If Len(Dia) = 1 Then Dia = "0" & Dia
Mes = Month(TBLISTA!data_envio)
If Len(Mes) = 1 Then Mes = "0" & Mes
ano = Year(TBLISTA!data_envio)

    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Seq_remessa from tbl_Detalhes_Recebimento where IDContaReceber = '" & TBLISTA!IdContaReceber & "' order by Seq_remessa desc", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        If IsNull(TBAbrir!Seq_remessa) = False And TBAbrir!Seq_remessa <> "" Then Seq = TBAbrir!Seq_remessa Else Seq = 1
    End If
    TBAbrir.Close
    
If frm_Instituicoes.txtNBanco.Text = "341" Then 'Itau then
    'O nome do arquivo remessa do Itaú só aceita no máximo 8 caracteres
    'seqremessa = Seq
    If Seq < 10 Then SeqRemessa = "0" & Seq & ".txt" Else SeqRemessa = Seq & ".txt"
    SeqRemessaTexto = Left(SeqRemessa, Len(SeqRemessa) - 4)
    Select Case Len(SeqRemessaTexto)
        Case 1: RemessaTexto = "0" & Right(SeqRemessaTexto, 1)
        Case 2: RemessaTexto = SeqRemessaTexto
        Case Is >= 3: RemessaTexto = Right(SeqRemessaTexto, 2)
    End Select
    Arquivo = Dia & Mes & Right(ano, 2) & RemessaTexto & ".txt"
   ' Layout = "CNAB400"
    'CobreBemX1.ArquivoRemessa.Sequencia = Left(SeqRemessa, Len(SeqRemessa) - 4)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

'Public Function FunTamanhoTextoZeroEsq(Texto As Variant, Tamanho As Integer) As String
'On Error GoTo tratar_erro
'Dim QuantZeroEsq As Double 'OK
'
'Texto1 = ""
'QuantZeroEsq = Tamanho - Len(Texto)
'If QuantZeroEsq > 0 Then
'    Do While QuantZeroEsq > 0
'        If Texto1 = "" Then Texto1 = "0" Else Texto1 = Texto1 & "0"
'        QuantZeroEsq = QuantZeroEsq - 1
'    Loop
'    FunTamanhoTextoZeroEsq = Texto1 & Texto
'Else
'    FunTamanhoTextoZeroEsq = Texto
'End If
'
'Exit Function
'tratar_erro:
'    MsgBox ("Descrição do erro : " + Error()), vbCritical
'    Exit Function
'End Function
'
'Public Function FunTamanhoTextoZeroDir(Texto As Variant, Tamanho As Integer) As String
'On Error GoTo tratar_erro
'Dim QuantZeroDir As Double 'OK
'
'Texto1 = ""
'QuantZeroDir = Tamanho - Len(Texto)
'If QuantZeroDir > 0 Then
'    Do While QuantZeroDir > 0
'        If Texto1 = "" Then Texto1 = "0" Else Texto1 = Texto1 & "0"
'        QuantZeroDir = QuantZeroDir - 1
'    Loop
'    FunTamanhoTextoZeroDir = Texto & Texto1
'Else
'    FunTamanhoTextoZeroDir = Texto
'End If
'
'Exit Function
'tratar_erro:
'    MsgBox ("Descrição do erro : " + Error()), vbCritical
'End Function
'
'
'Public Function FunAbreBD() As Boolean
'On Error GoTo tratar_erro
'
'Abrir = True
'FunAbreBD = True
'
'NomeCampo = "Caprind"
'Set Conexao = New ADODB.Connection
'With Conexao
'    .Provider = "SQLOLEDB"
'    .Properties("Data Source").Value = NomeServidor
'    .Properties("Initial catalog").Value = Nome_banco
'    .Properties("User ID").Value = IIf(Usuario_banco = "", "Procam", Usuario_banco)
'    .Properties("Password").Value = IIf(Senha_banco = "", "PRO0902loc$?", Senha_banco)
'    .Properties("Persist Security Info") = "False"
'    .Open
'End With
'
'
'
'Exit Function
'tratar_erro:
'    If Err.Number = "-2147467259" Then
'        Abrir = False
'        FunAbreBD = False
'        Exit Function
'    End If
'    MsgBox ("Descrição do erro : " + Error()), vbCritical
'    Exit Function
'End Function

'Sub ProcCarregaBancoDados()
'On Error GoTo tratar_erro
'
'NomeServidor = GetSetting("Procam", "CaprindSQL", "NomeServidor")
'Localrel = GetSetting("Procam", "CaprindSQL", "Localrel")
'Nome_banco = GetSetting("Procam", "CaprindSQL", "Nome_banco")
'Usuario_banco = GetSetting("Procam", "CaprindSQL", "Usuario_banco")
'Senha_banco = GetSetting("Procam", "CaprindSQL", "Senha_banco")
'
'Exit Sub
'tratar_erro:
'    MsgBox ("Descrição do erro : " + Error()), vbCritical
'    Exit Sub
'End Sub

'Sub ProcOrdenaListView(ByVal lvw As MSComctlLib.ListView, ByVal Coluna_Cabecalho As MSComctlLib.ColumnHeader)
'On Error GoTo tratar_erro
'
'If Coluna_Cabecalho.Tag = "N" Then
'    ProcSortListView lvw, Coluna_Cabecalho.Index, "ldtNumber", OrdAsc
'ElseIf Coluna_Cabecalho.Tag = "T" Then
'    ProcSortListView lvw, Coluna_Cabecalho.Index, "ldtString", OrdAsc
'ElseIf Coluna_Cabecalho.Tag = "D" Then
'        ProcSortListView lvw, Coluna_Cabecalho.Index, "ldtDateTime", OrdAsc
'End If
'OrdAsc = Not OrdAsc
'
'Exit Sub
'tratar_erro:
'    MsgBox ("Descrição do erro : " + Error()), vbCritical
'    Exit Sub
'End Sub

'Sub ProcSortListView(ListView As ListView, ByVal Index As Integer, ByVal DataType As String, ByVal Ascending As Boolean)
'On Error GoTo tratar_erro
'Dim i As Integer
'Dim l As Long
'Dim strFormat As String
'Dim lngCursor As Long
'Dim blnRestoreFromTag As Boolean
'Dim dte As Date
'
'Permitido = False
'
'lngCursor = ListView.MousePointer
'ListView.MousePointer = vbHourglass
'LockWindowUpdate ListView.hWnd
'
'Select Case DataType
'    Case "ldtString": blnRestoreFromTag = False
'    Case "ldtNumber":
'        strFormat = String$(20, "0") & "." & String$(10, "0")
'        With ListView.ListItems
'            If (Index = 1) Then
'                For l = 1 To .Count
'                    With .Item(l)
'                        If IsNumeric(.Text) Or Right(.Text, 1) = "%" Then
'                            If Right(.Text, 1) = "%" Then
'                                valor = Len(.Text) - 1
'                                Familiatext = Mid(.Text, 1, valor)
'                                Permitido = True
'                            Else
'                                Familiatext = .Text
'                            End If
'
'                            .Tag = Familiatext & Chr$(0) & .Tag
'                            If CDbl(Familiatext) >= 0 Then .Text = Format(CDbl(Familiatext), strFormat) Else .Text = "&" & Format(0 - CDbl(Familiatext), strFormat)
'                        Else
'                            .Tag = .Text & Chr$(0) & .Tag
'                            .Text = ""
'                        End If
'                    End With
'                Next l
'            Else
'                For l = 1 To .Count
'                    With .Item(l).ListSubItems(Index - 1)
'                        If IsNumeric(.Text) Or Right(.Text, 1) = "%" Then
'                            If Right(.Text, 1) = "%" Then
'                                valor = Len(.Text) - 1
'                                Familiatext = Mid(.Text, 1, valor)
'                                Permitido = True
'                            Else
'                                Familiatext = .Text
'                            End If
'
'                            .Tag = Familiatext & Chr$(0) & .Tag
'                            If CDbl(Familiatext) >= 0 Then .Text = Format(CDbl(Familiatext), strFormat) Else .Text = "&" & Format(0 - CDbl(Familiatext), strFormat)
'                        Else
'                            .Tag = .Text & Chr$(0) & .Tag
'                            .Text = ""
'                        End If
'                    End With
'                Next l
'            End If
'        End With
'        blnRestoreFromTag = True
'    Case "ldtDateTime":
'        strFormat = "YYYYMMDDHhNnSs"
'        With ListView.ListItems
'            If (Index = 1) Then
'                For l = 1 To .Count
'                    With .Item(l)
'                        If .Text <> "" Then
'                            .Tag = .Text & Chr$(0) & .Tag
'                            dte = (.Text)
'                            .Text = Format$(dte, strFormat)
'                        End If
'                    End With
'                Next l
'            Else
'                For l = 1 To .Count
'                    With .Item(l).ListSubItems(Index - 1)
'                        If .Text <> "" Then
'                            .Tag = .Text & Chr$(0) & .Tag
'                            dte = (.Text)
'                            .Text = Format$(dte, strFormat)
'                        End If
'                    End With
'                Next l
'            End If
'        End With
''        blnRestoreFromTag = True
''End Select
''
'ListView.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
'ListView.SortKey = Index - 1
'ListView.Sorted = True
'
'If blnRestoreFromTag Then
'    With ListView.ListItems
'        If (Index = 1) Then
'            For l = 1 To .Count
'                With .Item(l)
'                    If .Tag <> "" Then
'                        i = InStr(.Tag, Chr$(0))
'                        .Text = Left$(.Tag, i - 1)
'                        .Tag = Mid$(.Tag, i + 1)
'                    End If
'                End With
'            Next l
'        Else
'            For l = 1 To .Count
'                With .Item(l).ListSubItems(Index - 1)
'                    If .Tag <> "" Then
'                        i = InStr(.Tag, Chr$(0))
'                        .Text = Left$(.Tag, i - 1)
'                        .Tag = Mid$(.Tag, i + 1)
'                    End If
'                    If Permitido = True Then .Text = .Text & "%"
'                End With
'            Next l
'        End If
'    End With
'End If
'
'LockWindowUpdate 0&
'ListView.MousePointer = lngCursor
'ListView.Sorted = False
'
'Exit Sub
'tratar_erro:
'    MsgBox ("Descrição do erro : " + Error()), vbCritical
'    Exit Sub
'End Sub

