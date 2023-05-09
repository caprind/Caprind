VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmEstoque_Recebimento_Menu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Estoque | Recebimento | Menu"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   3390
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   714
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   1650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2355
      Left            =   240
      TabIndex        =   1
      Top             =   660
      Width           =   4065
      Begin DrawSuite2022.USButton btnPedido 
         Height          =   855
         Left            =   270
         TabIndex        =   2
         ToolTipText     =   "Criar nota fiscal com importação do pedido de compras"
         Top             =   330
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1508
         DibPicture      =   "frmEstoque_Recebimento_Menu.frx":0000
         Caption         =   "Importar nota do pedido de compras"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   4960354
         BorderColorDisabled=   13160660
         BorderColorDown =   4210752
         BorderColorOver =   49152
         GradientColor1  =   4960354
         GradientColor2  =   4960354
         GradientColor3  =   4960354
         GradientColor4  =   4960354
         GradientColorDisabled1=   14215660
         GradientColorDisabled2=   14215660
         GradientColorDisabled3=   14215660
         GradientColorDisabled4=   14215660
         GradientColorOver1=   49152
         GradientColorOver2=   49152
         GradientColorOver3=   49152
         GradientColorOver4=   49152
         GradientColorDown1=   32768
         GradientColorDown2=   32768
         GradientColorDown3=   32768
         GradientColorDown4=   32768
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   3
      End
      Begin DrawSuite2022.USButton btnXML 
         Height          =   855
         Left            =   270
         TabIndex        =   3
         ToolTipText     =   "Criar nota fiscal por importação do XML"
         Top             =   1260
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   1508
         DibPicture      =   "frmEstoque_Recebimento_Menu.frx":7180
         Caption         =   "Importar nota de arquivo XML"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderColor     =   5263559
         BorderColorDisabled=   13160660
         BorderColorDown =   4013465
         BorderColorOver =   4408288
         GradientColor1  =   5263559
         GradientColor2  =   5263559
         GradientColor3  =   5263559
         GradientColor4  =   5263559
         GradientColorDisabled1=   13160660
         GradientColorDisabled2=   13160660
         GradientColorDisabled3=   13160660
         GradientColorDisabled4=   13160660
         GradientColorOver1=   4408288
         GradientColorOver2=   4408288
         GradientColorOver3=   4408288
         GradientColorOver4=   4408288
         GradientColorDown1=   4013465
         GradientColorDown2=   4013465
         GradientColorDown3=   4013465
         GradientColorDown4=   4013465
         PicAlign        =   7
         PicSize         =   3
         PicSizeH        =   32
         PicSizeW        =   32
         ShowFocusRect   =   0   'False
         Theme           =   4
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   767
      DibPicture      =   "frmEstoque_Recebimento_Menu.frx":A568
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmEstoque_Recebimento_Menu.frx":14015
   End
End
Attribute VB_Name = "frmEstoque_Recebimento_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BtnPedido_Click()
On Error GoTo tratar_erro

'If ImportarXML = False Then

If USMsgBox("Deseja realmente criar a nota por importação do Pedido de compras ?", vbYesNo, "CAPRIND v5.0") = vbNo Then
 Exit Sub
End If

Moeda = frmEstoque_Recebimento.txtMoeda.Text
ValorMoeda = frmEstoque_Recebimento.txtvlrMoeda.Text

With frmEstoque_Recebimento
'==================================================================
'Cria a nota fiscal
'==================================================================
' Criar nota de itens importados
'==================================================================
If .txtuf = "EX" Then
 If .txtnotafiscal = "" Then
 Set TBAbrir = CreateObject("adodb.recordset")
'==================================================================
' Verifica se existe cadastro dessa nota
'==================================================================
 TBAbrir.Open "Select CAST(int_NotaFiscal AS int) AS NF, Serie,Modelo  FROM tbl_Dados_Nota_Fiscal where dt_DataEmissao = '" & DataEmissao & "' AND tipoNF = '" & TipoNF & "' and Aplicacao = 'P' and ID_empresa = " & .txtID_empresa & " and int_NotaFiscal IS NOT NULL order by dt_DataEmissao desc, NF desc", Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
   QuantsolicitadoN1 = TBAbrir!NF + 1
   FamiliaAntiga = QuantsolicitadoN1
   Familiatext = FunTamanhoTextoZeroEsq(FamiliaAntiga, 9)
   SerieNF = IIf(IsNull(TBAbrir!Serie), 1, TBAbrir!Serie)
  Else
   Familiatext = "000000001"
   SerieNF = 1
  End If
 .txtDataemissao = Format(Date, "dd/mm/yyyy")
 .txtnotafiscal = FunVerifExisteNumNF(TipoNF, .txtID_empresa, Familiatext, SerieNF, TBAbrir!Modelo)
 .txtSerie = SerieNF
'=============================================================================
' Atualiza os dados da nota no estoque controle recebimento e na movimentação
'=============================================================================
 Conexao.Execute "Update ECR set ECR.Nota_fiscal = '" & .txtnotafiscal & "', ECR.Serie = '" & txtSerie & "', ECR.Data_emissao =  '" & Format(txtDataemissao, "Short Date") & "' from Estoque_controle_recebimento ECR INNER JOIN Compras_pedido_lista CPL ON CPL.IDPedido = ECR.IDPedido and CPL.IDLista = ECR.IDLista and CPL.Desenho = ECR.Desenho where ECR.IDPedido = " & IIf(Txt_ID_pedido = "", 0, Txt_ID_pedido) & " and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " and (ECR.Nota_fiscal IS NULL or ECR.Nota_fiscal = N'')"
 StrSql = "Update EM set EM.Documento = '" & .txtnotafiscal & "' from Estoque_movimentacao EM INNER JOIN Estoque_controle_recebimento ECR ON ECR.Id = EM.IDEstoque_recebimento INNER JOIN Compras_pedido_lista CPL ON CPL.IDPedido = ECR.IDPedido and CPL.IDLista = ECR.IDLista and CPL.Desenho = ECR.Desenho where ECR.IDPedido = " & IIf(Txt_ID_pedido = "", 0, Txt_ID_pedido) & " and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " and ECR.Nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Data_emissao = '" & Format(txtDataemissao, "Short Date") & "'"
 'Debug.print StrSql
 
 Conexao.Execute "Update EM set EM.Documento = '" & .txtnotafiscal & "' from Estoque_movimentacao EM INNER JOIN Estoque_controle_recebimento ECR ON ECR.Id = EM.IDEstoque_recebimento INNER JOIN Compras_pedido_lista CPL ON CPL.IDPedido = ECR.IDPedido and CPL.IDLista = ECR.IDLista and CPL.Desenho = ECR.Desenho where ECR.IDPedido = " & IIf(Txt_ID_pedido = "", 0, Txt_ID_pedido) & " and ECR.Programacao = 'False' and ECR.id_empresa = " & txtID_empresa & " and ECR.Nota_fiscal = '" & txtnotafiscal & "' and ECR.Serie = '" & txtSerie & "' and ECR.Data_emissao = '" & Format(txtDataemissao, "Short Date") & "'"
 End If
End If

'==============================================================================
' Verifica se existe cadastro da nota fiscal
'==============================================================================
Set TBGravar = CreateObject("adodb.recordset")
StrSql = "Select * from tbl_Dados_Nota_Fiscal where dt_DataEmissao = '" & DataEmissao & "'  and ID_empresa = " & .txtID_empresa & " and Id_Int_Cliente = " & .Txt_ID_forn & " and int_NotaFiscal = '" & .txtnotafiscal & "' and Serie = '" & .txtSerie & "' and int_TipoNota = 2 and TipoNF = '" & TipoNF & "'"
'Debug.print StrSql


TBGravar.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then

'===================================================================
' Busca os dados do pedido de compra
'===================================================================

    Set TBFornecedor = CreateObject("adodb.recordset")
    TBFornecedor.Open "Select CF.*, CP.ID_empresa, CC.Moeda FROM (Compras_fornecedores CF INNER JOIN Compras_pedido CP ON CF.IDCliente = CP.idfornecedor) LEFT JOIN Compras_comercial CC ON CC.IDpedido = CP.IDpedido where CP.Pedido = '" & .txtProg_pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFornecedor.EOF = False Then
        

'        If IsNull(TBFornecedor!Moeda) = False And TBFornecedor!Moeda <> "" And TBFornecedor!Moeda <> "REAL" Then
'            Moeda = TBFornecedor!Moeda
'Mensagem:
'            Dolar = InputBox("Favor informar o valor do " & Moeda & ".")
'            If Dolar = "" Then Exit Sub
'            If IsNumeric(Dolar) = False Then
'                USMsgBox ("Só é permitido número neste campo."), vbExclamation, "CAPRIND v5.0"
'                GoTo Mensagem
'            End If
'            ValorMoeda = Dolar
'        End If
        
        TBGravar.AddNew
        TBGravar!TabelaSN = 0
        TBGravar!Regime = FunVerifRegimeEmpresa(TBFornecedor!ID_empresa)
        TBGravar!pedido_interno = False
        TBGravar!DtValidacaoOF = Now
        TBGravar!RespValidacaoOF = pubUsuario
        TBGravar!int_NotaFiscal = .txtnotafiscal
        TBGravar!Serie = .txtSerie
        If .txtuf = "EX" Then TBGravar!Aplicacao = "P" Else TBGravar!Aplicacao = "T"
        TBGravar!int_TipoNota = "2"
        TBGravar!dt_DataEmissao = .txtDataemissao
        TBGravar!txt_Hora_Saida = Format(Now, "hh:mm:ss")
        TBGravar!TipoNF = TipoNF
        TBGravar!txt_Razao_Nome = .txtFornecedor
        TBGravar!Moeda = Moeda
        TBGravar!ValorMoeda = ValorMoeda
        TBGravar!ID_empresa = TBFornecedor!ID_empresa
        TBGravar!Id_Int_Cliente = TBFornecedor!IDCliente
        TBGravar!txt_Endereco = IIf(IsNull(TBFornecedor!Endereco), "", TBFornecedor!Endereco)
        TBGravar!Numero = IIf(IsNull(TBFornecedor!Numero), "", TBFornecedor!Numero)
        TBGravar!txt_Bairro = IIf(IsNull(TBFornecedor!Bairro), "", TBFornecedor!Bairro)
        TBGravar!Txt_CEP = IIf(IsNull(TBFornecedor!CEP), "", TBFornecedor!CEP)
        TBGravar!txt_Municipio = IIf(IsNull(TBFornecedor!Cidade), "", TBFornecedor!Cidade)
        TBGravar!txt_Fone_Fax = IIf(IsNull(TBFornecedor!Telefones), "", TBFornecedor!Telefones)
        
        TBGravar!txt_UF = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
        UF = IIf(IsNull(TBFornecedor!Estado), "", TBFornecedor!Estado)
        
        If TBFornecedor!idTipoEmpresa = 1 Then TBGravar!txt_CNPJ_CPF = IIf(IsNull(TBFornecedor!CPF_CNPJ), "", TBFornecedor!CPF_CNPJ)
        TBGravar!txt_IE_Cliente = IIf(IsNull(TBFornecedor!RG_IE), "", TBFornecedor!RG_IE)
        If TBFornecedor!Pessoa = "JURÍDICA" Then TBGravar!txt_tipocliente = "J" Else TBGravar!txt_tipocliente = "F"
    End If
    TBGravar!Int_status = "1"
    TBGravar.Update
    ID_nota = TBGravar!ID
Else
    ID_nota = TBGravar!ID
    ValorMoeda = TBGravar!ValorMoeda
    
    'Verifica se a NF já foi validada e não permite alteração
    If IsNull(TBGravar!DtValidacao) = False Then
        USMsgBox ("Esta nota fiscal não será alterada, pois a mesma já foi validada."), vbInformation, "CAPRIND v5.0"
        TBGravar.Close
        GoTo Validada
    End If
End If
TBGravar.Close

'=============================================================================================================
' ***** Cria ou altera a lista dos produtos
'=============================================================================================================
Desenho = ""
OrdemTexto = ""
valor = 0
ValorTotal = 0
OF = 0
NovoValor = ""

Set TBReceber = CreateObject("adodb.recordset")
StrSql = "Select ECR.*, CP.idfornecedor, CP.pedido from (Estoque_Controle_recebimento ECR INNER JOIN Compras_pedido CP ON ECR.idpedido = CP.idpedido) INNER JOIN Compras_fornecedores CF ON CF.IDCliente = CP.idfornecedor where CP.Pedido = '" & .txtProg_pedido & "' and  CP.idfornecedor = " & .Txt_ID_forn & " and ECR.nota_fiscal = '" & .txtnotafiscal & "' and ECR.Serie = '" & .txtSerie & "' and ECR.Programacao = 'False' and ECR.id_empresa = " & .txtID_empresa & " order by ECR.Desenho"
'Debug.print StrSql

TBReceber.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBReceber.EOF = False Then
Do While TBReceber.EOF = False

 If TBReceber!Desenho <> Desenho Then
  OF = 0
 End If

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from compras_pedido_lista where idlista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then

 If TBPedido!Tipo = "P" Then
  Prod = True
  Prodpedido = True
  ServPedido = False
 Else
  Prodpedido = False
  ServPedido = True
 End If
 
 If TBPedido!Tipo = "S" Then
  Servicos = True
 End If
 
 If Desenho <> TBReceber!Desenho Or Desenho = TBReceber!Desenho And valor <> IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto) Or OrdemTexto <> IIf(IsNull(TBPedido!Ordem), 0, TBPedido!Ordem) Then
  ValorTotal = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
  OF = IIf(IsNull(TBPedido!Ordem), 0, TBPedido!Ordem)
  GoTo Prosseguir
 Else
  GoTo Proximo
 End If

End If
TBPedido.Close
Prosseguir:

If OF = 0 Then
TextoFiltro = "(Ordem = 0 or Ordem is null)"
Else
TextoFiltro = "Ordem = '" & OF & "'"
End If


qt = 0

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from compras_pedido_lista where idlista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
NovoValor1 = Replace(ValorTotal, ",", ".")
Set TBFI = CreateObject("adodb.recordset")
StrSql = "Select Sum(ECR.recebido) as qt from compras_pedido_lista CPL inner join estoque_controle_recebimento ECR on CPL.idlista = ECR.idlista where CPL.preco_unitario_desconto = " & IIf(NovoValor1 = "", 0, NovoValor1) & " and " & TextoFiltro & " and ECR.nota_fiscal = '" & .txtnotafiscal & "' and ECR.Serie = '" & .txtSerie & "' and ECR.Desenho = '" & TBReceber!Desenho & "' and ECR.id_empresa = " & .txtID_empresa
'Debug.print StrSql

TBFI.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

If TBFI.EOF = False Then
qt = Format(IIf(IsNull(TBFI!qt), 0, TBFI!qt), "###,##0.0000")
If TBPedido!Un <> TBPedido!Unidade_com Then

If FunVerifUNConversao(TBPedido!Un, TBPedido!Unidade_com) = True Then
    qt = FunConverteUN(TBPedido!Unidade_com, TBPedido!Un, qt, TBReceber!Desenho)
Else
    qt = qt * FunVerificaTabelaConversaoUnidade(TBPedido!Un, TBPedido!Unidade_com)
End If

End If

End If

TBFI.Close
End If
TBPedido.Close
'===================================================================
' Multiplica o valor total dos produtos pelo valor do Dolar
'===================================================================

If Moeda <> "REAL" Then
ValorTotal = Format(ValorTotal * ValorMoeda, "###,##0.0000000000")
End If

NovoValor = Replace(ValorTotal, ",", ".")
'======================================================================================
' Começa a gravar os itens na lista da nota
'======================================================================================
If Prodpedido = True Then
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Detalhes_Nota where int_Cod_Produto = '" & TBReceber!Desenho & "' and id_nota = " & ID_nota & " and dbl_valorunitario = " & IIf(NovoValor = "", 0, NovoValor) & " and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = True Then
TBAbrir.AddNew
End If

TBAbrir!Tipo = "P"
TBAbrir!int_Cod_Produto = TBReceber!Desenho
TBAbrir!int_Qtd = IIf(qt <> 0, qt, frmEstoque_Recebimento.txtQuantidade)
TBAbrir!Saldo = IIf(qt <> 0, qt, frmEstoque_Recebimento.txtQuantidade)
TBAbrir!int_NotaFiscal = .txtnotafiscal
TBAbrir!ID_nota = ID_nota
'================================================================================================
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from compras_pedido_lista where idpedido = " & TBReceber!IDpedido & " and IDLista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic

If TBPedido.EOF = False Then
TBAbrir!Txt_descricao = IIf(IsNull(TBPedido!Descricao), "", TBPedido!Descricao)
TBAbrir!Codproduto = IIf(IsNull(TBPedido!Codproduto), "", TBPedido!Codproduto)
IDlista = IIf(IsNull(TBPedido!IDlista), "", TBPedido!IDlista)
TBAbrir!txt_Unid = IIf(IsNull(TBPedido!Un), "", TBPedido!Un)
TBAbrir!Unidade_com = IIf(IsNull(TBPedido!Unidade_com), "", TBPedido!Unidade_com)
TBAbrir!Familia = IIf(IsNull(TBPedido!Familia), "", TBPedido!Familia)
TBAbrir!N_referencia = IIf(IsNull(TBPedido!N_referencia), "", TBPedido!N_referencia)
TBAbrir!Ordem = TBPedido!Ordem

If TBPedido!Remessa = True Then
TBAbrir!retorno = True
End If

If IsNull(TBPedido!ID_CFOP) = False And TBPedido!ID_CFOP <> "" Then
TBAbrir!ID_CFOP = TBPedido!ID_CFOP
End If

If IsNull(TBPedido!ID_CF) = False And TBPedido!ID_CF <> "" Then
TBAbrir!ID_CF = TBPedido!ID_CF
End If

If IsNull(TBPedido!CST) = False And TBPedido!CST <> "" Then
TBAbrir!txt_CST = TBPedido!CST
End If

End If
TBPedido.Close
'====================================================================================================

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select ID_CFOP, ID_CF from projproduto where desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then

If IsNull(TBAbrir!ID_CFOP) = True Or TBAbrir!ID_CFOP = "" Then
TBAbrir!ID_CFOP = IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP)
End If

If IsNull(TBAbrir!ID_CF) = True Or TBAbrir!ID_CF = "" Then
TBAbrir!ID_CF = IIf(IsNull(TBItem!ID_CF), 0, TBItem!ID_CF)
End If
TBItem.Close
End If
'======================================================================================================

If IsNull(TBAbrir!ID_CFOP) = False And TBAbrir!ID_CFOP <> "" Then
 Set TBItem = CreateObject("adodb.recordset")
 TBItem.Open "Select * from tbl_NaturezaOperacao_CST where ID_CFOP = " & TBAbrir!ID_CFOP, Conexao, adOpenKeyset, adLockOptimistic
 If TBItem.EOF = False Then
  If TBItem.RecordCount = 1 Then
   If IsNull(TBAbrir!txt_CST) = True Or TBAbrir!txt_CST = "" Then
    TBAbrir!txt_CST = TBItem!CST_ICMS
   End If
   TBAbrir!CST_IPI = TBItem!CST_IPI
   TBAbrir!CST_PIS = TBItem!CST_PIS
   TBAbrir!CST_Cofins = TBItem!CST_Cofins
  End If
  End If
End If
TBItem.Close


Set TBPI_Lista_produto = CreateObject("adodb.recordset")
StrSql = "Select CPL.* from compras_pedido_lista CPL INNER JOIN compras_pedido CP ON CPL.idpedido = CP.idpedido where CPL.idlista = " & TBReceber!IDlista & " and CP.idpedido = " & TBReceber!IDpedido
'Debug.print StrSql

TBPI_Lista_produto.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
If TBPI_Lista_produto.EOF = False Then

If Moeda <> "REAL" Then
TBAbrir!dbl_ValorUnitario = Format(TBPI_Lista_produto!preco_unitario_desconto * ValorMoeda, "###,##0.0000000000")
Else
TBAbrir!dbl_ValorUnitario = Format(TBPI_Lista_produto!preco_unitario_desconto, "###,##0.0000000000")
End If

TBAbrir!int_ICMS = IIf(IsNull(TBPI_Lista_produto!ICMS), 0, TBPI_Lista_produto!ICMS)
TBAbrir!int_IPI = IIf(IsNull(TBPI_Lista_produto!IPI), 0, TBPI_Lista_produto!IPI)
TBAbrir!dbl_valoripi = Format(((TBAbrir!dbl_ValorUnitario * qt) * IIf(IsNull(TBPI_Lista_produto!IPI), 0, TBPI_Lista_produto!IPI)) / 100, "###,##0.00")
TBAbrir!dbl_ValorTotal = Format(TBAbrir!dbl_ValorUnitario * qt, "###,##0.00")
TBAbrir!Valor_frete = Format(IIf(IsNull(TBPI_Lista_produto!Frete), 0, TBPI_Lista_produto!Frete) * ValorMoeda, "###,##0.00")
TBAbrir!Valor_seguro = Format(IIf(IsNull(TBPI_Lista_produto!Seguro), 0, TBPI_Lista_produto!Seguro) * ValorMoeda, "###,##0.00")
TBAbrir!Valor_acessorias = Format(IIf(IsNull(TBPI_Lista_produto!Acessorias), 0, TBPI_Lista_produto!Acessorias) * ValorMoeda, "###,##0.00")
TBAbrir!Tem_IPI_frete = TBPI_Lista_produto!Frete_IPI

If IsNull(TBPI_Lista_produto!OS) = False And TBPI_Lista_produto!OS <> "" Then
ProcAtualizaCTTEROrdem TBPI_Lista_produto!OS
End If

'==================================================================
' Acerta valor unitario e valor total em estoque movimentação
'==================================================================
StrSql = "update Estoque_movimentacao set VlrUnit = '" & Replace(TBAbrir!dbl_ValorUnitario, ",", ".") & "', VlrTotal = '" & Replace(TBAbrir!dbl_ValorTotal, ",", ".") & "' From Estoque_Movimentacao EM Where EM.IDLista_recebimento = '" & TBPI_Lista_produto!IDlista & "'"
'Debug.print StrSql

Conexao.Execute StrSql

'==================================================================
' Acerta valor unitario e valor total em estoque controle
'==================================================================
'StrSql = "update Estoque_controle set Valor_Unitario = '" & Replace(TBAbrir!dbl_ValorUnitario, ",", ".") & "', Valor_Total = '" & Replace(TBAbrir!dbl_ValorTotal, ",", ".") & "' From Estoque_controle EC Where EC.Lote = '" & .txtProg_pedido.Text & "'"
'Conexao.Execute StrSql
'==================================================================
' Grava o id da lista da nota em estoque movimentação
'==================================================================
StrSql = "update Estoque_movimentacao set ID_prod_NF = '" & TBAbrir!Int_codigo & "' From Estoque_Movimentacao EM Where EM.IDLista_recebimento = '" & TBPI_Lista_produto!IDlista & "'"
'Debug.print StrSql

Conexao.Execute StrSql
End If

TBAbrir.Update


TBPI_Lista_produto.Close

'Salvar CST
'ProcSalvarCSTLista
Else
'==================================================================
' Nota de serviços
'==================================================================
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from tbl_Detalhes_Nota where int_Cod_Produto = '" & TBReceber!Desenho & "' and id_nota = " & ID_nota & " and dbl_ValorUnitario = " & IIf(NovoValor = "", 0, NovoValor) & " and Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
TBAbrir!Tipo = "S"
TBAbrir!int_Cod_Produto = TBReceber!Desenho
TBAbrir!int_Qtd = IIf(qt <> 0, qt, frmEstoque_Recebimento.txtQuantidade)
TBAbrir!int_NotaFiscal = txtnotafiscal
TBAbrir!ID_nota = ID_nota
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from compras_pedido_lista where idpedido = " & TBReceber!IDpedido & " and IDLista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
TBAbrir!Txt_descricao = IIf(IsNull(TBPedido!Descricao), "", TBPedido!Descricao)
TBAbrir!Codproduto = IIf(IsNull(TBPedido!Codproduto), "", TBPedido!Codproduto)
IDlista = IIf(IsNull(TBPedido!IDlista), "", TBPedido!IDlista)
TBAbrir!txt_Unid = IIf(IsNull(TBPedido!Un), "", TBPedido!Un)
TBAbrir!Unidade_com = IIf(IsNull(TBPedido!Unidade_com), "", TBPedido!Unidade_com)
TBAbrir!Familia = IIf(IsNull(TBPedido!Familia), "", TBPedido!Familia)
TBAbrir!N_referencia = IIf(IsNull(TBPedido!N_referencia), "", TBPedido!N_referencia)
TBAbrir!Ordem = TBPedido!Ordem

If IsNull(TBPedido!ID_CFOP) = False And TBPedido!ID_CFOP <> "" Then TBAbrir!ID_CFOP = TBPedido!ID_CFOP
End If
TBPedido.Close

If IsNull(TBAbrir!ID_CFOP) = True Or TBAbrir!ID_CFOP = "" Then
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & TBReceber!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
TBAbrir!ID_CFOP = IIf(IsNull(TBItem!ID_CFOP), 0, TBItem!ID_CFOP)
End If
End If

Set TBPI_Lista_produto = CreateObject("adodb.recordset")
TBPI_Lista_produto.Open "Select CPL.* from compras_pedido_lista CPL INNER JOIN compras_pedido CP ON CPL.idpedido = CP.idpedido where CPL.idlista = " & TBReceber!IDlista & " and CP.idpedido = " & TBReceber!IDpedido, Conexao, adOpenKeyset, adLockOptimistic
If TBPI_Lista_produto.EOF = False Then
TBAbrir!dbl_ValorUnitario = Format(TBPI_Lista_produto!preco_unitario_desconto * ValorMoeda, "###,##0.0000000000")
TBAbrir!ISS = IIf(IsNull(TBPI_Lista_produto!ISSQN), 0, TBPI_Lista_produto!ISSQN)
TBAbrir!VlrISS = Format(((TBAbrir!dbl_ValorUnitario * qt) * IIf(IsNull(TBPI_Lista_produto!ISSQN), 0, TBPI_Lista_produto!ISSQN)) / 100, "###,##0.00")
TBAbrir!dbl_ValorTotal = Format(TBAbrir!dbl_ValorUnitario * IIf(qt <> 0, qt, frmEstoque_Recebimento.txtQuantidade), "###,##0.00")

If IsNull(TBPI_Lista_produto!OS) = False And TBPI_Lista_produto!OS <> "" Then ProcAtualizaCTTEROrdem TBPI_Lista_produto!OS
End If
TBPI_Lista_produto.Close
TBAbrir.Update
End If

Set TBFIltro = CreateObject("adodb.recordset")
TBFIltro.Open "Select ECR.IDlista, ECR.Recebido from Estoque_Controle_recebimento ECR inner join Compras_pedido CP on ECR.idpedido = CP.idpedido where CP.idfornecedor = " & TBReceber!IDFornecedor & " and ECR.nota_fiscal = '" & .txtnotafiscal & "' and ECR.Serie = '" & .txtSerie & "' and ECR.Programacao = 'False' and ECR.id_empresa = " & .txtID_empresa & " and ECR.Desenho = '" & TBAbrir!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBFIltro.EOF = False Then
Do While TBFIltro.EOF = False
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select idlista from compras_pedido_lista where idlista = " & TBFIltro!IDlista & " and preco_unitario_desconto = " & IIf(NovoValor1 = "", 0, NovoValor1) & " and (Ordem = " & OF & " or Ordem IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * FROM tbl_Detalhes_Nota_pedidos where ID_nota = " & ID_nota & " and ID_prod_NF = " & TBAbrir!Int_codigo & " and ID_carteira = " & TBFIltro!IDlista & " and Codinterno = '" & TBAbrir!int_Cod_Produto & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBGravar.EOF = True Then TBGravar.AddNew
TBGravar!ID_nota = ID_nota
TBGravar!ID_prod_NF = TBAbrir!Int_codigo
TBGravar!ID_carteira = TBFIltro!IDlista
TBGravar!Codinterno = TBAbrir!int_Cod_Produto
TBGravar!quantidade = TBFIltro!Recebido
TBGravar.Update
TBGravar.Close
End If
TBPedido.Close
TBFIltro.MoveNext
Loop
End If
TBFIltro.Close

TBAbrir.Close
Proximo:
Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select * from tbl_proposta_nota where id_nota = " & ID_nota & " and proposta = '" & TBReceber!Pedido & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = True Then
TBPedido.AddNew
TBPedido!Proposta = TBReceber!Pedido
TBPedido!NF = txtnotafiscal
TBPedido!ID_nota = ID_nota
TBPedido.Update
End If

Set TBPedido = CreateObject("adodb.recordset")
TBPedido.Open "Select preco_unitario_desconto, Ordem from compras_pedido_lista where idlista = " & TBReceber!IDlista, Conexao, adOpenKeyset, adLockOptimistic
If TBPedido.EOF = False Then
valor = IIf(IsNull(TBPedido!preco_unitario_desconto), 0, TBPedido!preco_unitario_desconto)
OrdemTexto = IIf(IsNull(TBPedido!Ordem), "", TBPedido!Ordem)
End If
TBPedido.Close
Desenho = TBReceber!Desenho
TBReceber.MoveNext
Loop
Else
USMsgBox ("Não há produto recebido para a nota " & txtnotafiscal & "."), vbExclamation, "CAPRIND v5.0"
TBFornecedor.Close
Exit Sub
End If
If Prod = True And Servicos = True Then
TipoNF = "M1SA"
ElseIf Prod = True And Servicos = False Then
TipoNF = "M1"
Else
TipoNF = "SA"
End If
Conexao.Execute "Update tbl_Dados_Nota_Fiscal Set TipoNF = '" & TipoNF & "' where ID = " & ID_nota

Validada:
    If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then Unload frmFaturamento_Prod_Serv
    If txtuf = "EX" Then
        Faturamento_NF_Saida = True
        Formulario = "Faturamento/Nota fiscal/Própria"
    Else
        Faturamento_NF_Saida = False
        Formulario = "Estoque/Nota fiscal"
    End If
    
    With frmFaturamento_Prod_Serv
    NF_Recebimento = True
        .Novo_Nota = False
        .Faturamento_Vendas_PI = False
        .txtId.Text = ID_nota
        .txtNFiscal.Text = frmEstoque_Recebimento.txtnotafiscal
        .ProcCarregaDadosNota .txtId.Text
        .ProcCarregaLista
        .ProcCarregaListaServicos
        .ProcGravarTotaisNota
        .ProcCarregaDadosTransporte
        .ProcCarregaDuplicatas
        .txt_DtEmissao.Text = Format(frmEstoque_Recebimento.txtDataemissao, "dd/mm/yyyy")
        .txtSerie.Locked = False
        .txtSerie.TabStop = True
        
        CamposFiltro = "NF.ID, NF.dt_DataEmissao, NF.dt_Saida_Entrada, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, TN.dbl_Valor_Total_Nota, NF.txt_Razao_Nome, NF.Int_status, NF.Imprimir, NF.ID_empresa, NF.Aplicacao, NF.DtValidacaoOF, NF.DtValidacao"
        .Strsql_Faturamento = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .txtId
        .Strsql_FaturamentoTotal = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor1, Sum(TN.Valor_Total_Receber_Pagar) as Valor2 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .txtId & " and NF.Int_status = 1"
        .Strsql_FaturamentoTotalCanc = "Select Sum(TN.dbl_Valor_Total_Nota) as Valor3 from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.ID = " & .txtId & " and NF.Int_status = 2"
        .Strsql_FaturamentoNFe = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.TipoNF <> 'SA' AND NF.ID = " & .txtId
        .Strsql_FaturamentoNFSe = "Select " & CamposFiltro & " from tbl_Dados_Nota_Fiscal NF INNER JOIN tbl_Totais_Nota TN ON NF.ID = TN.ID_Nota where NF.TipoNF = 'SA' AND NF.ID = " & .txtId
        .ProcCarregaListaNota (1)
        
        If USMsgBox("Deseja prosseguir com o preenchimento dos dados da nota fiscal?", vbYesNo, "CAPRIND v5.0") = vbNo Then
        Unload frmFaturamento_Prod_Serv
        Else
        Unload frmEstoque_Recebimento_Menu
        End If
        
    End With

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub btnXML_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente criar a nota fiscal por importação de arquivo XML ?", vbYesNo, "CAPRIND v5.0") = vbYes Then
NFentrada = True
NFSaida = False
TPNota = "T"
ProcImportarXML

If ImportarXML = True Then
USMsgBox "importação realizada com sucesso, verifique o cadastro da nota fiscal!", vbInformation, "CAPRIND v5.0"
Else
USMsgBox "Importação não realizada, verifique o cadastro da nota fiscal!", vbInformation, "CAPRIND v5.0"
End If

End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub
