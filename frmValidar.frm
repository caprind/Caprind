VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmValidar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "CAPRIND V5.0 | Validação de registro"
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   5205
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton btnValidar 
      Height          =   885
      Left            =   390
      TabIndex        =   2
      Top             =   2400
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1561
      DibPicture      =   "frmValidar.frx":0000
      Caption         =   "Validar registro(s)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   767
      DibPicture      =   "frmValidar.frx":23FD
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
      Icon            =   "frmValidar.frx":47FA
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   2760
      Top             =   150
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmValidar.frx":4B14
      Count           =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   390
      TabIndex        =   3
      Top             =   870
      Width           =   4335
      Begin VB.TextBox txtUsuario 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1020
         TabIndex        =   0
         ToolTipText     =   "Usuário."
         Top             =   330
         Width           =   2865
      End
      Begin VB.TextBox txtSenha 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1020
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Senha."
         Top             =   690
         Width           =   2865
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         Caption         =   "Nº:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   -11580
         TabIndex        =   6
         Top             =   4200
         Width           =   270
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuário :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   300
         TabIndex        =   5
         Top             =   330
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Senha :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   390
         TabIndex        =   4
         Top             =   690
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmValidar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TextoFiltroFin As String
Dim EmpenhoVerificar As Boolean

Private Sub btnValidar_Click()
On Error GoTo tratar_erro

procValidar (True)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: procValidar (True)
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select A.* from usuarios U INNER JOIN acessos A ON A.IDUsuario = U.IDUsuario where U.usuario = '" & pubUsuario & "' and A.Acesso = '" & Formulario & "' and A.Validacao = 'True' ", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    procValidar (False)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtUsuario_GotFocus()
On Error GoTo tratar_erro

'ProcCarregaToolBar1 Me, 3975, 4, True
Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select A.* from usuarios U INNER JOIN acessos A ON A.IDUsuario = U.IDUsuario where U.usuario = '" & pubUsuario & "' and A.Acesso = '" & Formulario & "' and A.Validacao = 'True' ", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = False Then
    Unload Me
Else
    FunGotFocus txtUsuario
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub procValidar(Validacao_obrig As Boolean)
On Error GoTo tratar_erro

Select Case Formulario
    Case "Vendas/Clientes": ProcValidar1 frmVendas_cliente.Lista, "Clientes", "IDCliente", Formulario, Validacao_obrig
    Case "Compras/Famílias": ProcValidar1 frmproj_familia.Lista, "projfamilia", "Codigo", Formulario, Validacao_obrig
    Case "Compras/Fornecedores": ProcValidar1 frmCompras_fornecedores.Lista, "Compras_Fornecedores", "IDCliente", Formulario, Validacao_obrig
    Case "Compras/Pedido": ProcValidar1 frmCompras_Pedido.listapedido, "compras_pedido", "IDpedido", Formulario, Validacao_obrig
    Case "Compras/Pedido/Aprovar": ProcValidar1 frmCompras_Pedido.listapedido, "compras_pedido", "IDpedido", Formulario, Validacao_obrig
    Case "Compras/Produtos e serviços": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Compras/Produtos e serviços/Validar estrutura": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Compras/Produtos e serviços/Validar plano de inspeção": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Engenharia/Famílias": ProcValidar1 frmproj_familia.Lista, "projfamilia", "Codigo", Formulario, Validacao_obrig
    Case "Engenharia/Produtos e serviços": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Engenharia/Produtos e serviços/Validar estrutura": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Engenharia/Produtos e serviços/Validar plano de inspeção": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Engenharia/Processos": ProcValidar1 frmProcessos.ListaProcessos, "Processos", "IDprocesso", Formulario, Validacao_obrig
    Case "Engenharia/Conjuntos": ProcValidarEngenhariaConjunto Formulario, Validacao_obrig
    Case "Engenharia/Estrutura/Detalhada": ProcValidarEngenhariaEstrutura Formulario, Validacao_obrig
    Case "Engenharia/Estrutura/Resumida": ProcValidarEngenhariaEstruturaResumida Formulario, Validacao_obrig
    Case "Faturamento/Fiscal/Classificação fiscal": ProcValidar1 frm_Classificacao_Fiscal.Lista, "tbl_ClassificacaoFiscal", "Idclass", Formulario, Validacao_obrig
    Case "Faturamento/Fiscal/Natureza de operação": ProcValidar1 frm_Natureza_OP.lst_NatOp, "tbl_NaturezaOperacao", "IDcountCFOP", Formulario, Validacao_obrig
    Case "Faturamento/Nota fiscal/Terceiros": ProcValidar1 frmFaturamento_Prod_Serv.ListaNota, "tbl_Dados_Nota_Fiscal", "ID", Formulario, Validacao_obrig
    Case "Faturamento/Nota fiscal/Própria": ProcValidar1 frmFaturamento_Prod_Serv.ListaNota, "tbl_Dados_Nota_Fiscal", "ID", Formulario, Validacao_obrig
    Case "Estoque/Local de armazenamento": ProcValidar1 frmEstoque_Localarmaz.Lista_locarmazenamento, "Estoque_Localarmazenamento_criar", "ID", Formulario, Validacao_obrig
    Case "Estoque/Ordem de faturamento": ProcValidar1 frmEstoque_Ordem_Faturamento.ListaNota, "tbl_Dados_Nota_Fiscal", "ID", Formulario, Validacao_obrig
    Case "Estoque/Nota fiscal": ProcValidar1 frmFaturamento_Prod_Serv.ListaNota, "tbl_Dados_Nota_Fiscal", "ID", Formulario, Validacao_obrig
    Case "Estoque/Inventário": ProcValidar1 frmestoque_fisico.Lista, "Estoque_fisico", "ID", Formulario, Validacao_obrig
    Case "Outros/Solicitação": ProcValidar1 frmCompras_Requisicao.Lista_req, "Compras_requisicao", "ID_Requisicao", Formulario, Validacao_obrig
    Case "Outros/Solicitação/Autorizar solicitação": ProcValidar1 frmCompras_Requisicao.Lista_req, "Compras_requisicao", "ID_Requisicao", Formulario, Validacao_obrig
    Case "Qualidade/Famílias": ProcValidar1 frmproj_familia.Lista, "projfamilia", "Codigo", Formulario, Validacao_obrig
    Case "Qualidade/Inspeção de recebimento": ProcValidar1 frmCompras_recebimento.ListProdReceb, "Compras_recebimento", "ID", Formulario, Validacao_obrig
    Case "Qualidade/PPAP/PSW": ProcValidar1 frmQualidadePPAP.Lista, "QualidadePPAP", "IdPPAP", Formulario, Validacao_obrig
    Case "Qualidade/Não conformidade/Descrição da não conformidade": ProcValidar1 frmcqnc_descricaoNC.Lista, "CQ_NC_FABRICA_causa", "ID", Formulario, Validacao_obrig
    Case "Vendas/Famílias": ProcValidar1 frmproj_familia.Lista, "projfamilia", "Codigo", Formulario, Validacao_obrig
    Case "Vendas/Produtos e serviços": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Vendas/Produtos e serviços/Validar estrutura": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Vendas/Produtos e serviços/Validar plano inspeção": ProcValidar1 frmproj_produto.Lista, "projproduto", "codproduto", Formulario, Validacao_obrig
    Case "Vendas/Proposta comercial": ProcValidar1 frmVendas_proposta.Lista, "vendas_proposta", "Cotacao", Formulario, Validacao_obrig
    Case "Vendas/Pedido interno": ProcValidar1 frmVendas_PI.Lista, "vendas_proposta", "Cotacao", Formulario, Validacao_obrig
    Case "Vendas/Vendedores": ProcValidar1 frmVendas_Vendedores.Lista, "Vendas_Vendedores", "ID", Formulario, Validacao_obrig
    Case "PCP/Gerenciamento de ordem": ProcValidar1 frmprod.Lista, "Producao", "Ordem", Formulario, Validacao_obrig
    Case "PCP/Gerenciamento de ordem/Validar resultados": ProcValidarPCP_Custo Formulario, Validacao_obrig
    Case "PCP/Plano da produção": ProcValidar1 frmPlano_producao.Lista, "ProducaoFases_OS", "ID", Formulario, Validacao_obrig
    Case "Engenharia/Normas": ProcValidar1 frmNorma.Lista, "Norma", "ID", Formulario, Validacao_obrig
    Case "RH/Funcionários": ProcValidar1 frmRH_Funcionarios.Lista, "Funcionarios", "ID", Formulario, Validacao_obrig
    Case "Financeiro/Instituições": ProcValidar1 frm_Instituicoes.lst_Instituicoes, "tbl_Instituicoes", "ID", Formulario, Validacao_obrig
    Case "Qualidade/Plano de inspeção":
        If frmPlanoinspecao_validacao.Opt_fase = True Then
            ProcValidar1 frmPlanoinspecao.Lista, "Plano", "IdPlano", Formulario, Validacao_obrig
        Else
            ProcValidarProdPlano Formulario, Validacao_obrig
        End If
    Case "Estoque/Requisição de materiais": ProcValidar1 frmRequisicao_materiais.Lista_req, "Requisicao_materiais", "ID", Formulario, Validacao_obrig
    Case "Qualidade/Solicitação de ação": ProcValidar1 frmCQ_SA.ListView1, "CQ_SA", "Id", Formulario, Validacao_obrig
    Case "Outros/Análise crítica/Engenharia": ProcValidarAnalise Formulario, Validacao_obrig
    Case "Outros/Análise crítica/Processos": ProcValidarAnalise Formulario, Validacao_obrig
    Case "Outros/Análise crítica/Qualidade": ProcValidarAnalise Formulario, Validacao_obrig
    Case "Outros/Análise crítica/Pcp": ProcValidarAnalise Formulario, Validacao_obrig
    Case "Outros/Análise crítica/Compras": ProcValidarAnalise Formulario, Validacao_obrig
    Case "Outros/Solicitação de produção": ProcValidar1 frmOutros_Solicitacao_PCP.Lista_solicitacao, "Outros_SolicitacaoPCP", "ID", Formulario, Validacao_obrig
    Case "Outros/Solicitação de produção/Autorizar solicitação": ProcValidar1 frmOutros_Solicitacao_PCP.Lista_solicitacao, "Outros_SolicitacaoPCP", "ID", Formulario, Validacao_obrig
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidar1(Lista As ListView, NomeTabelaFiltro As String, NomeCampoFiltro As String, Modulo1 As String, Validacao_obrig As Boolean)
On Error GoTo tratar_erro

If Validacao_obrig = True Then
    Acao = "validar/cancelar validação"
    If txtUsuario.Text = "" Then
        NomeCampo = "o usuário"
        ProcVerificaAcao
        txtUsuario.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = "" Then
        NomeCampo = "a senha"
        ProcVerificaAcao
        txtSenha.SetFocus
        Exit Sub
    End If
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select IDUsuario, Usuario, Setor from usuarios where usuario = '" & txtUsuario.Text & "' and Senha = '" & txtSenha.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "select * from acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = '" & Modulo1 & "' and Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = True Then
            USMsgBox ("Atenção usuário " & txtUsuario & ", você não tem autorização para este recurso."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
        txtUsuario = TBUsuarios!Usuario
        SetorTexto = TBUsuarios!Setor
    Else
        USMsgBox ("Verifique se a senha e o usuário estão corretos."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBUsuarios.Close
Else
    txtUsuario = pubUsuario
    SetorTexto = pubSetor
End If

TextoFiltroFin = ""
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            If TextoFiltroFin = "" Then TextoFiltroFin = "ID = " & .ListItems.Item(InitFor) Else TextoFiltroFin = TextoFiltroFin & " or ID = " & .ListItems.Item(InitFor)
            
            Set TBVendas = CreateObject("adodb.recordset")
            TBVendas.Open "Select * from " & NomeTabelaFiltro & " where " & NomeCampoFiltro & " = " & .ListItems.Item(InitFor), Conexao, adOpenKeyset, adLockOptimistic
            If TBVendas.EOF = False Then
                If Modulo1 = "Estoque/Ordem de faturamento" Then
                    If IsNull(TBVendas!DtValidacaoOF) = True Then
                        TBVendas!DtValidacaoOF = Now
                        TBVendas!RespValidacaoOF = txtUsuario
                        Evento = "Validar"
                        If TBVendas!txt_UF <> "EX" Then frmEstoque_Ordem_Faturamento.ProcCorrigeValorImpostosSN TBVendas!ID
                    Else
                        TBVendas!RespValidacaoOF = Null
                        TBVendas!DtValidacaoOF = Null
                        Evento = "Cancelar validação"
                    End If
                ElseIf Modulo1 = "Outros/Solicitação/Autorizar solicitação" Then
                        If IsNull(TBVendas!Data_autorizacao) = True Then
                            TBVendas!Data_autorizacao = Now
                            TBVendas!Autorizado = txtUsuario
                            TBVendas!setorautor = IIf(SetorTexto = "", Null, SetorTexto)
                            TBVendas!status = "LIBERADA"
                            Evento = "Aprovar"
                        Else
                            TBVendas!Autorizado = Null
                            TBVendas!Data_autorizacao = Null
                            TBVendas!setorautor = Null
                            TBVendas!status = "ABERTA"
                            Evento = "Cancelar aprovação"
                        End If
                    ElseIf Modulo1 = "Vendas/Pedido interno" Then
                            If IsNull(TBVendas!DtValidacaoPI) = True Then
                                TBVendas!DtValidacaoPI = Now
                                TBVendas!RespValidacaoPI = txtUsuario
                                Evento = "Validar"
                                ProcEmpenharProdServEst
                                ProcGerarNecessidadePI
                            Else
                                TBVendas!RespValidacaoPI = Null
                                TBVendas!DtValidacaoPI = Null
                                Evento = "Cancelar validação"
                                Conexao.Execute "DELETE from ECEV from Estoque_Controle_Empenho_Vendas ECEV INNER JOIN vendas_carteira VC ON VC.Codigo = ECEV.ID_carteira where VC.Cotacao = " & TBVendas!Cotacao
                                Conexao.Execute "DELETE from PM from Producaomaterial PM INNER JOIN vendas_carteira VC ON VC.Codigo = PM.ID_carteira where VC.Cotacao = " & TBVendas!Cotacao
                            End If
                        ElseIf Modulo1 = "Compras/Pedido/Aprovar" Then
                                If IsNull(TBVendas!Data_aprovado) = True Then
                                    TBVendas!Data_aprovado = Now
                                    TBVendas!Resp_aprovado = txtUsuario
                                    
                                    If FunVerifStatusAprovadoPC(TBVendas!ID_empresa) = True Then
                                        StatusPC = "APROVADO"
                                        StatusProd = "APROVADO"
                                    Else
                                        StatusPC = "ABERTO"
                                        StatusProd = "N_RECEBIDO"
                                    End If
                                    
                                    TBVendas!Status_pedido = StatusPC
                                    Evento = "Aprovar pedido de compra"
                                    Conexao.Execute "UPDATE Compras_pedido_lista Set Status_Item = '" & StatusProd & "' where IDpedido = " & .ListItems(InitFor)
                                    
                                    FunAlterarProdSimiliarOrdemPC frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex), frmCompras_Pedido.txtIDPedido
                                    ProcCriarRMOrdemPC .ListItems(InitFor), frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex)
                                Else
                                    TBVendas!Data_aprovado = Null
                                    TBVendas!Resp_aprovado = Null
                                    TBVendas!Status_pedido = "AGUARDANDO APROVAÇÃO"
                                    Evento = "Cancelar aprovação do pedido de compra"
                                    Conexao.Execute "UPDATE Compras_pedido_lista Set Status_Item = 'AGUARDANDO APROVAÇÃO' where IDpedido = " & .ListItems(InitFor)
                                    
                                    ProcExcluirRMOrdemPC .ListItems(InitFor), frmCompras_Pedido.Cmb_empresa.ItemData(frmCompras_Pedido.Cmb_empresa.ListIndex)
                                End If
                            ElseIf Modulo1 = "Outros/Solicitação de produção/Autorizar solicitação" Then
                                    If IsNull(TBVendas!Data_autorizacao) = True Then
                                        TBVendas!Data_autorizacao = Now
                                        TBVendas!Autorizado = txtUsuario
                                        TBVendas!status = "LIBERADA"
                                        Evento = "Aprovar"
                                    Else
                                        TBVendas!Autorizado = Null
                                        TBVendas!Data_autorizacao = Null
                                        TBVendas!status = "ABERTA"
                                        Evento = "Cancelar aprovação"
                                    End If
                                ElseIf Right(Modulo1, 37) = "Produtos e serviços/Validar estrutura" Then
                                        If IsNull(TBVendas!DtValidacaoConj) = True Then
                                            TBVendas!DtValidacaoConj = Now
                                            TBVendas!RespValidacaoConj = txtUsuario
                                            Evento = "Validar"
                                        Else
                                            TBVendas!DtValidacaoConj = Null
                                            TBVendas!RespValidacaoConj = Null
                                            Evento = "Cancelar validação"
                                        End If
                                    ElseIf Right(Modulo1, 45) = "Produtos e serviços/Validar plano de inspeção" Then
                                            If IsNull(TBVendas!DtValidacaoPlano) = True Then
                                                TBVendas!DtValidacaoPlano = Now
                                                TBVendas!RespValidacaoPlano = txtUsuario
                                                Evento = "Validar"
                                                ProcValidarPlano TBVendas!Desenho, True
                                            Else
                                                TBVendas!DtValidacaoPlano = Null
                                                TBVendas!RespValidacaoPlano = Null
                                                Evento = "Cancelar validação"
                                                ProcValidarPlano TBVendas!Desenho, False
                                            End If
                                        Else
                                            If IsNull(TBVendas!DtValidacao) = True Then
                                                TBVendas!DtValidacao = Now
                                                TBVendas!RespValidacao = txtUsuario
                                                Evento = "Validar"
                                                If Modulo1 = "Estoque/Inventário" Then frmestoque_fisico.procValidar True
                                                If Modulo1 = "Engenharia/Processos" Then ProcCriarOS .ListItems.Item(InitFor).ListSubItems(6), TBVendas!IDPROCESSO
                                                If Modulo1 = "Qualidade/Plano de inspeção" Then If IsNull(TBVendas!IDFase) = False And TBVendas!IDFase <> "" Then ProcPlano .ListItems.Item(InitFor).ListSubItems(4), .ListItems(InitFor), TBVendas!IDFase
                                                
                                                If Modulo1 = "Faturamento/Nota fiscal/Própria" Then
                                                    procNotaFiscal (Left(frmFaturamento_Prod_Serv.Cmb_modelo, 2)) 'Gera um numero para a nota fiscal
                                                    'If TBVendas!txt_UF <> "EX" Then frmFaturamento_Prod_Serv.ProcCorrigeValorImpostosSN TBVendas!ID
                                                End If
                                                If Modulo1 = "Faturamento/Nota fiscal/Terceiros" Or Modulo1 = "Estoque/Nota fiscal" Then
                                                    ProcAtualizaVlrEntradaEstoque False
                                                ElseIf Modulo1 = "Faturamento/Nota fiscal/Própria" Then
                                                        If TBVendas!txt_UF = "EX" Then ProcAtualizaVlrEntradaEstoque True
                                                End If
                                            Else
                                                TBVendas!RespValidacao = Null
                                                TBVendas!DtValidacao = Null
                                                'TBVendas!Novo_lote = True
                                                
                                                Evento = "Cancelar validação"
                                                If Modulo1 = "Estoque/Inventário" Then frmestoque_fisico.procValidar False
                                                If Modulo1 = "Engenharia/Processos" Then ProcExcluirOS .ListItems.Item(InitFor).ListSubItems(6), TBVendas!IDPROCESSO
                                                If Modulo1 = "Qualidade/Plano de inspeção" Then If IsNull(TBVendas!IDFase) = False And TBVendas!IDFase <> "" Then ProcPlano .ListItems.Item(InitFor).ListSubItems(4), 0, TBVendas!IDFase
                                                
                                                If Left(Modulo1, 23) = "Faturamento/Nota fiscal" Or Modulo1 = "Estoque/Ordem de faturamento" Or Modulo1 = "Estoque/Nota fiscal" Then
                                                    'Conta gerada pela nota
                                                    Conexao.Execute "DELETE from CC from CC_realizado CC INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = CC.ID_financeiro Where CR.ID_Nota = " & TBVendas!ID & " and CC.Operacao = 'Crédito'"
                                                    Conexao.Execute "DELETE from FF from Familia_financeiro FF INNER JOIN tbl_contas_receber CR ON CR.IDIntconta = FF.IDconta Where CR.ID_Nota = " & TBVendas!ID & " and FF.Tipoconta = 'R' and (CR.Proposta IS NULL or CR.Proposta = N'')"
                                                    Conexao.Execute "DELETE from FC from tbl_Fluxo_de_caixa FC INNER JOIN tbl_contas_receber CR ON CR.IDFluxo = FC.IDFluxo Where CR.ID_Nota = " & TBVendas!ID & " and (CR.Proposta IS NULL or CR.Proposta = N'')"
                                                    Conexao.Execute "DELETE from tbl_contas_receber where ID_Nota = " & TBVendas!ID & " and (Proposta IS NULL or Proposta = N'')"
                                                    'Conta gerada pelo pedido
                                                    Conexao.Execute "Update FC set FC.int_NotaFiscal = NULL from tbl_Fluxo_de_caixa FC INNER JOIN tbl_contas_receber CR ON CR.IDFluxo = FC.IDFluxo Where CR.ID_Nota = " & TBVendas!ID & " and CR.Proposta IS NOT NULL"
                                                    Conexao.Execute "Update tbl_contas_receber Set ID_nota = NULL, NFiscal = NULL where ID_Nota = " & TBVendas!ID & " and Proposta IS NOT NULL"
                                                    
                                                    'Conta gerada pela nota
                                                    Conexao.Execute "DELETE from FF from Familia_financeiro FF INNER JOIN tbl_ContasPagar CP ON CP.IDIntconta = FF.IDconta Where CP.ID_Nota = " & TBVendas!ID & " and FF.Tipoconta = 'P' and (CP.txt_pedido IS NULL or CP.txt_pedido = N'')"
                                                    Conexao.Execute "DELETE from FC from tbl_Fluxo_de_caixa FC INNER JOIN tbl_ContasPagar CP ON CP.IDFluxo = FC.IDFluxo Where CP.ID_Nota = " & TBVendas!ID & " and (CP.txt_pedido IS NULL or CP.txt_pedido = N'')"
                                                    Conexao.Execute "DELETE from tbl_ContasPagar where ID_Nota = " & TBVendas!ID & "  and (txt_pedido IS NULL or txt_pedido = N'')"
                                                    'Conta gerada pelo pedido
                                                    Conexao.Execute "Update FC set FC.int_NotaFiscal = NULL from tbl_Fluxo_de_caixa FC INNER JOIN tbl_ContasPagar CP ON CP.IDFluxo = FC.IDFluxo Where CP.ID_Nota = " & TBVendas!ID & " and CP.txt_pedido IS NOT NULL"
                                                    Conexao.Execute "Update tbl_ContasPagar Set ID_nota = NULL, txt_ndocumento = NULL where ID_Nota = " & TBVendas!ID & " and txt_pedido IS NOT NULL"
                                                    
                                                    Conexao.Execute "Update CC set CC.ID_Financeiro = 0 from CC_realizado CC INNER JOIN tbl_Detalhes_Recebimento TBL on CC.ID_duplicata = TBL.ID where TBL.ID_nota = " & TBVendas!ID
                                                    
'                                                    If Modulo1 = "Faturamento/Nota fiscal/Própria" Then
'                                                        Set TBCodigoDesc = CreateObject("adodb.recordset")
'                                                        TBCodigoDesc.Open "Select Codigo from Empresa where Codigo = " & TBVendas!ID_empresa & " and Baixa_Auto_Estoque_NF = 'True'", Conexao, adOpenKeyset, adLockOptimistic
'                                                        If TBCodigoDesc.EOF = False And TBVendas!int_TipoNota = 1 Then
'                                                            procExcluirMovimentacao_NFe
'                                                        End If
'                                                        TBCodigoDesc.Close
'                                                    End If
                                                End If
                                            End If
                End If
                TBVendas.Update
                
                If Modulo1 = "Faturamento/Nota fiscal/Própria" Then ProcAlteraStatusPI
            End If
            
            Modulo = Modulo1
            ID_documento = .ListItems.Item(InitFor)
            Select Case Modulo1
                Case "Vendas/Clientes": Documento = "Cliente: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "Compras/Famílias": Documento = "Código: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "Compras/Fornecedores": Documento = "Fornecedor: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "Compras/Pedido": Documento = "Nº pedido: " & .ListItems.Item(InitFor).ListSubItems(2)
                Case "Compras/Pedido/Aprovar": Documento = "Nº pedido: " & .ListItems.Item(InitFor).ListSubItems(2)
                Case "Compras/Produtos e serviços": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Vendas/Famílias": Documento = "Código: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "Vendas/Produtos e serviços": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Engenharia/Famílias": Documento = "Código: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "Engenharia/Produtos e serviços": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Engenharia/Processos": Documento = "Processo: " & .ListItems.Item(InitFor).ListSubItems(1) & " - Rev.: " & .ListItems.Item(InitFor).ListSubItems(2) & " - Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(6) & " - Rev.: " & .ListItems.Item(InitFor).ListSubItems(7)
                Case "Faturamento/Fiscal/Classificação fiscal": Documento = "Classificação: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Faturamento/Fiscal/Natureza de operação": Documento = "CFOP: " & .ListItems.Item(InitFor).ListSubItems(1) & " - Descrição.: " & .ListItems.Item(InitFor).ListSubItems(2)
                Case "Faturamento/Nota fiscal/Terceiros": Documento = "Nº nota: " & .ListItems.Item(InitFor).ListSubItems(3) & " - Tipo: " & .ListItems.Item(InitFor).ListSubItems(4) & " - Série: " & .ListItems.Item(InitFor).ListSubItems(5)
                Case "Faturamento/Nota fiscal/Própria": Documento = "Nº nota: " & .ListItems.Item(InitFor).ListSubItems(3) & " - Tipo: " & .ListItems.Item(InitFor).ListSubItems(4) & " - Série: " & .ListItems.Item(InitFor).ListSubItems(5)
                Case "Estoque/Local de armazenamento": Documento = "Local de armazenamento: " & .ListItems.Item(InitFor).ListSubItems(4)
                Case "Estoque/Ordem de faturamento": Documento = "Nº ordem: " & .ListItems.Item(InitFor).ListSubItems(2) & " - Tipo: " & .ListItems.Item(InitFor).ListSubItems(4)
                Case "Estoque/Nota fiscal": Documento = "Nº nota: " & .ListItems.Item(InitFor).ListSubItems(3) & " - Tipo: " & .ListItems.Item(InitFor).ListSubItems(4) & " - Série: " & .ListItems.Item(InitFor).ListSubItems(5)
                Case "Estoque/Inventário": Documento = "Data: " & .ListItems.Item(InitFor).ListSubItems(2) & " - Código interno: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "PCP/Gerenciamento de ordem": Documento = "Ordem: " & .ListItems.Item(InitFor).ListSubItems(1) & " - Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(4)
                Case "PCP/Plano da produção": Documento = "Nº plano: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Engenharia/Normas": Documento = "Norma: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Vendas/Proposta comercial": Documento = "Nº proposta: " & .ListItems.Item(InitFor).ListSubItems(2) & " - Rev.: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "Vendas/Pedido interno": Documento = "Nº pedido: " & .ListItems.Item(InitFor).ListSubItems(2) & " - Rev.: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "RH/Funcionários": Documento = "Código: " & .ListItems.Item(InitFor).ListSubItems(1) & " - Funcionário: " & .ListItems.Item(InitFor).ListSubItems(2)
                Case "Financeiro/Instituições": Documento = "Instituição bancária: " & .ListItems.Item(InitFor).ListSubItems(7)
                Case "Qualidade/Famílias": Documento = "Código: " & .ListItems.Item(InitFor).ListSubItems(3)
                Case "Qualidade/Inspeção de recebimento": Documento = "Lote: " & frmCompras_recebimento.Txt_lote & " - Cliente/fornecedor: " & frmCompras_recebimento.Txt_cliente_forn & " - Cód. interno: " & .ListItems(InitFor).ListSubItems(1)
                Case "Qualidade/Plano de inspeção": Documento = "Plano de inspeção: " & .ListItems(InitFor)
                Case "Qualidade/Não conformidade/Descrição da não conformidade": Documento = "Descrição: " & .ListItems(InitFor).ListSubItems(1)
                Case "Outros/Solicitação": Documento = "Nº solicitação: " & .ListItems(InitFor).ListSubItems(3)
                Case "Outros/Solicitação/Autorizar solicitação": Documento = "Nº solicitação: " & .ListItems(InitFor).ListSubItems(3)
                Case "Estoque/Requisição de materiais": Documento = "Nº requisição: " & .ListItems(InitFor).ListSubItems(2)
                Case "Qualidade/Solicitação de ação": Documento = "Nº SA: " & .ListItems(InitFor)
                Case "Compras/Produtos e serviços/Validar estrutura": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Vendas/Produtos e serviços/Validar estrutura": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Engenharia/Produtos e serviços/Validar estrutura": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Compras/Produtos e serviços/Validar plano de inspeção": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Vendas/Produtos e serviços/Validar plano de inspeção": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
                Case "Engenharia/Produtos e serviços/Validar plano de inspeção": Documento = "Cód. interno: " & .ListItems.Item(InitFor).ListSubItems(1)
'                Case "Outros/Solicitação de produção": Documento = "Nº solicitação: " & .ListItems(InitFor).ListSubItems(3)
'                Case "Outros/Solicitação de produção/Autorizar solicitação": Documento = "Nº solicitação: " & .ListItems(InitFor).ListSubItems(3)
            End Select
            '==================================
            Documento1 = ""
            ProcGravaEvento
            '==================================
        End If
    Next InitFor
End With

If Modulo1 = "Outros/Solicitação/Autorizar solicitação" Or Modulo1 = "Compras/Pedido/Aprovar" Then TextoMsg = "aprovação" Else TextoMsg = "validação"
USMsgBox "Operação de " & TextoMsg & " realizada com sucesso.", vbInformation, "CAPRIND v5.0"
ProcCarregaLista Modulo1
If Validacao_obrig = True Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidarPCP_Custo(Modulo1 As String, Validacao_obrig As Boolean)
On Error GoTo tratar_erro

If Validacao_obrig = True Then
    Acao = "validar/cancelar validação"
    If txtUsuario.Text = "" Then
        NomeCampo = "o usuário"
        ProcVerificaAcao
        txtUsuario.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = "" Then
        NomeCampo = "a senha"
        ProcVerificaAcao
        txtSenha.SetFocus
        Exit Sub
    End If
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select * from usuarios where usuario = '" & txtUsuario.Text & "' and Senha = '" & txtSenha.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "select * from acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = '" & Modulo1 & "' and Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = True Then
            USMsgBox ("Atenção usuário " & txtUsuario & ", você não tem autorização para este recurso."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
        txtUsuario = TBUsuarios!Usuario
    Else
        USMsgBox ("Verifique se a senha e o usuário estão corretos."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBUsuarios.Close
Else
    txtUsuario = pubUsuario
End If

With frmprod
    If .SSTab1.Tab = 1 Then
        With .Lista
            For InitFor = 1 To .ListItems.Count
                If .ListItems.Item(InitFor).Checked = True Then ProcValidaPCP_Custo1 .ListItems.Item(InitFor), .ListItems.Item(InitFor).ListSubItems(4)
            Next InitFor
        End With
    Else
        ProcValidaPCP_Custo1 .txtof, .txtdesenho
    End If
    
    .atualiza_lista_ordens (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
    Ordem = IIf(.txtof = "", 0, .txtof)
    .ProcCarregaOrdem
    .ProcCarregaResultados
End With
USMsgBox ("Operação de validação realizada com sucesso."), vbInformation, "CAPRIND v5.0"
If Validacao_obrig = True Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidaPCP_Custo1(Ordem As Long, Codinterno As String)
On Error GoTo tratar_erro

Set TBVendas = CreateObject("adodb.recordset")
TBVendas.Open "Select * from producao where ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
If TBVendas.EOF = False Then
    If IsNull(TBVendas!DtValidacao_custo) = True Then
        TBVendas!DtValidacao_custo = Now
        TBVendas!RespValidacao_Custo = txtUsuario
        Evento = "Validar resultado"
        TBVendas.Update
        
        'Corrige empenhos
        Conexao.Execute "DELETE from Producao_NF_Consignada where Ordem = " & Ordem & " and Qtde_saida = 0"
        Conexao.Execute "UPDATE Producao_NF_Consignada Set Quantidade = Qtde_saida, Quantidade_PC = Qtde_saida_PC where Ordem = " & Ordem
                
        ProcBackup_Apontamento True
    Else
        TBVendas!DtValidacao_custo = Null
        TBVendas!RespValidacao_Custo = Null
        Evento = "Cancelar validação do resutltado"
        TBVendas.Update
        ProcBackup_Apontamento False
    End If
End If
'==================================
Modulo = "PCP/Gerenciamento de ordem"
ID_documento = Ordem
Documento = "Ordem: " & Ordem & " - Cód. interno: " & Codinterno
Documento1 = ""
ProcGravaEvento
'==================================

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidarProdPlano(Modulo1 As String, Validacao_obrig As Boolean)
On Error GoTo tratar_erro

If Validacao_obrig = True Then
    Acao = "validar/cancelar validação"
    If txtUsuario.Text = "" Then
        NomeCampo = "o usuário"
        ProcVerificaAcao
        txtUsuario.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = "" Then
        NomeCampo = "a senha"
        ProcVerificaAcao
        txtSenha.SetFocus
        Exit Sub
    End If
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select * from usuarios where usuario = '" & txtUsuario.Text & "' and Senha = '" & txtSenha.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "select * from acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = '" & Modulo1 & "' and Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = True Then
            USMsgBox ("Atenção usuário " & txtUsuario & ", você não tem autorização para este recurso."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
        txtUsuario = TBUsuarios!Usuario
    Else
        USMsgBox ("Verifique se a senha e o usuário estão corretos."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBUsuarios.Close
Else
    txtUsuario = pubUsuario
End If

With frmPlanoinspecao
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select Codproduto, DtValidacaoPlano, RespValidacaoPlano from Projproduto where Desenho = '" & .txtdesenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        If IsNull(TBVendas!DtValidacaoPlano) = True Then
            TBVendas!DtValidacaoPlano = Now
            TBVendas!RespValidacaoPlano = txtUsuario
            Evento = "Validar plano de inspeção no produto"
        Else
            TBVendas!DtValidacaoPlano = Null
            TBVendas!RespValidacaoPlano = Null
            Evento = "Cancelar validação do plano de inspeção no produto"
        End If
        TBVendas.Update
        '==================================
        Modulo = Modulo1
        ID_documento = TBVendas!Codproduto
        Documento = "Cód. interno: " & .txtdesenho
        Documento1 = ""
        ProcGravaEvento
        '==================================
    End If
    TBVendas.Close
    
    Set TBplano = CreateObject("adodb.recordset")
    TBplano.Open "Select * from Plano where IDPlano = " & .txtPI, Conexao, adOpenKeyset, adLockOptimistic
    If TBplano.EOF = False Then
        .ProcLimpar
        .ProcCarregaDados
    End If
End With
USMsgBox ("Operação de validação realizada com sucesso."), vbInformation, "CAPRIND v5.0"
If Validacao_obrig = True Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidarPlano(Desenho_plano As String, Valida_plano As Boolean)
On Error GoTo tratar_erro

Set TBplano = CreateObject("adodb.recordset")
TBplano.Open "Select DtValidacao, RespValidacao from plano where Desenho = '" & Desenho_plano & "'", Conexao, adOpenKeyset, adLockOptimistic
Do While TBplano.EOF = False
    If Valida_plano = True Then
        TBplano!DtValidacao = Now
        TBplano!RespValidacao = txtUsuario
    Else
        TBplano!DtValidacao = Null
        TBplano!RespValidacao = Null
    End If
    TBplano.Update
    TBplano.MoveNext
Loop
TBplano.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidarEngenhariaConjunto(Modulo1 As String, Validacao_obrig As Boolean)
On Error GoTo tratar_erro

If Validacao_obrig = True Then
    Acao = "validar/cancelar validação"
    If txtUsuario.Text = "" Then
        NomeCampo = "o usuário"
        ProcVerificaAcao
        txtUsuario.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = "" Then
        NomeCampo = "a senha"
        ProcVerificaAcao
        txtSenha.SetFocus
        Exit Sub
    End If
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select * from usuarios where usuario = '" & txtUsuario.Text & "' and Senha = '" & txtSenha.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "select * from acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = '" & Modulo1 & "' and Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = True Then
            USMsgBox ("Atenção usuário " & txtUsuario & ", você não tem autorização para este recurso."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
        txtUsuario = TBUsuarios!Usuario
    Else
        USMsgBox ("Verifique se a senha e o usuário estão corretos."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBUsuarios.Close
Else
    txtUsuario = pubUsuario
End If

With frmproj_conjunto
    Set TBProduto = CreateObject("adodb.recordset")
    
    TBProduto.Open "Select * from Projconjunto_desc_versao where Codproduto = " & .Txt_cod_produto & " and Versao = '" & .cmbVersao_pesquisar & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = True Then
        TBProduto.AddNew
    End If
    
    TBProduto!Codproduto = .Txt_cod_produto
    TBProduto!versao = .cmbVersao_pesquisar
    If IsNull(TBProduto!DtValidacao) = True Or TBProduto!DtValidacao = "" Then
        TBProduto!DtValidacao = Now
        TBProduto!RespValidacao = txtUsuario
        Evento = "Validar"
    Else
        TBProduto!DtValidacao = Null
        TBProduto!RespValidacao = Null
        Evento = "Cancelar validação"
    End If
    TBProduto.Update
    .txtDtValidacao = IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao)
    .txtRespValidacao = IIf(IsNull(TBProduto!RespValidacao), "", TBProduto!RespValidacao)
    ProcValidarEngenhariaEstruturaProd TBProduto!Codproduto
    '==================================
    Modulo = "Engenharia/Conjuntos"
    ID_documento = TBProduto!Codproduto
    Documento = "Cód. interno: " & .txtdesenhoproduto & " - Versão: " & .cmbVersao_pesquisar
    Documento1 = ""
    ProcGravaEvento
    '==================================
End With
USMsgBox ("Operação de validação realizada com sucesso."), vbInformation, "CAPRIND v5.0"
If Validacao_obrig = True Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidarEngenhariaEstrutura(Modulo1 As String, Validacao_obrig As Boolean)
On Error GoTo tratar_erro

If Validacao_obrig = True Then
    Acao = "validar/cancelar validação"
    If txtUsuario.Text = "" Then
        NomeCampo = "o usuário"
        ProcVerificaAcao
        txtUsuario.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = "" Then
        NomeCampo = "a senha"
        ProcVerificaAcao
        txtSenha.SetFocus
        Exit Sub
    End If
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select * from usuarios where usuario = '" & txtUsuario.Text & "' and Senha = '" & txtSenha.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "select * from acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = '" & Modulo1 & "' and Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = True Then
            USMsgBox ("Atenção usuário " & txtUsuario & ", você não tem autorização para este recurso."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
        txtUsuario = TBUsuarios!Usuario
    Else
        USMsgBox ("Verifique se a senha e o usuário estão corretos."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBUsuarios.Close
Else
    txtUsuario = pubUsuario
End If

With frmproj_produto_estrutura
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from Projconjunto_desc_versao where ID = " & .Txt_ID_desc_versao, Conexao, adOpenKeyset, adLockOptimistic
    If TBProduto.EOF = True Then TBProduto.AddNew
    TBProduto!Codproduto = .IDProduto
    TBProduto!versao = .cmbVersao
    If IsNull(TBProduto!DtValidacao) = True Then
        TBProduto!DtValidacao = Now
        TBProduto!RespValidacao = txtUsuario
        Evento = "Validar"
    Else
        TBProduto!DtValidacao = Null
        TBProduto!RespValidacao = Null
        Evento = "Cancelar validação"
    End If
    TBProduto.Update
    ProcValidarEngenhariaEstruturaProd TBProduto!Codproduto
    FunValidarEstrutura TBProduto!Codproduto, IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao), IIf(IsNull(TBProduto!RespValidacao), "", TBProduto!RespValidacao), .cmbVersao
    '==================================
    Modulo = "Engenharia/Estrutura"
    ID_documento = TBProduto!Codproduto
    Documento = "Cód. interno: " & Desenho & " - Versão: " & .cmbVersao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    .ProcCarregaLista
End With
USMsgBox ("Operação de validação realizada com sucesso."), vbInformation, "CAPRIND v5.0"
If Validacao_obrig = True Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidarEngenhariaEstruturaResumida(Modulo1 As String, Validacao_obrig As Boolean)
On Error GoTo tratar_erro

If Validacao_obrig = True Then
    Acao = "validar/cancelar validação"
    If txtUsuario.Text = "" Then
        NomeCampo = "o usuário"
        ProcVerificaAcao
        txtUsuario.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = "" Then
        NomeCampo = "a senha"
        ProcVerificaAcao
        txtSenha.SetFocus
        Exit Sub
    End If
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select * from usuarios where usuario = '" & txtUsuario.Text & "' and Senha = '" & txtSenha.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "select * from acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = '" & Modulo1 & "' and Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = True Then
            USMsgBox ("Atenção usuário " & txtUsuario & ", você não tem autorização para este recurso."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
        txtUsuario = TBUsuarios!Usuario
    Else
        USMsgBox ("Verifique se a senha e o usuário estão corretos."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBUsuarios.Close
Else
    txtUsuario = pubUsuario
End If

With frmproj_produto_estrutura_Resumida
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from Projconjunto_desc_versao where ID = " & .Txt_ID_desc_versao, Conexao, adOpenKeyset, adLockOptimistic
'====================================================================
' Se não existir cria já validado
'====================================================================
    If TBProduto.EOF = True Then
        TBProduto.AddNew
            TBProduto!Codproduto = .IDProduto
            TBProduto!versao = .cmbVersao
            TBProduto!DtValidacao = Now
            TBProduto!RespValidacao = txtUsuario
            Evento = "Validar"
            TBProduto.Update
            GoTo Feito
    End If
    
'====================================================================
' Se existir validada ou não
'====================================================================
        If IsNull(TBProduto!DtValidacao) = True Then
            TBProduto!DtValidacao = Now
            TBProduto!RespValidacao = txtUsuario
            Evento = "Validar"
            TBProduto.Update
        Else
        If TBProduto.RecordCount > 1 Then
           Conexao.Execute "DELETE from Projconjunto_desc_versao where ID = " & .Txt_ID_desc_versao
        Else
            TBProduto!DtValidacao = Null
            TBProduto!RespValidacao = Null
            Evento = "Cancelar validação"
            TBProduto.Update
        End If
        End If

Feito:
    ProcValidarEngenhariaEstruturaProd TBProduto!Codproduto
    FunValidarEstrutura TBProduto!Codproduto, IIf(IsNull(TBProduto!DtValidacao), "", TBProduto!DtValidacao), IIf(IsNull(TBProduto!RespValidacao), "", TBProduto!RespValidacao), .cmbVersao
    '==================================
    Modulo = "Engenharia/Estrutura"
    ID_documento = TBProduto!Codproduto
    Documento = "Cód. interno: " & Desenho & " - Versão: " & .cmbVersao
    Documento1 = ""
    ProcGravaEvento
    '==================================
    .ProcCarregaLista
End With
USMsgBox ("Operação de validação realizada com sucesso."), vbInformation, "CAPRIND v5.0"
If Validacao_obrig = True Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcValidarAnalise(Modulo1 As String, Validacao_obrig As Boolean)
On Error GoTo tratar_erro

If Validacao_obrig = True Then
    Acao = "validar/cancelar validação"
    If txtUsuario.Text = "" Then
        NomeCampo = "o usuário"
        ProcVerificaAcao
        txtUsuario.SetFocus
        Exit Sub
    End If
    If txtSenha.Text = "" Then
        NomeCampo = "a senha"
        ProcVerificaAcao
        txtSenha.SetFocus
        Exit Sub
    End If
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select * from usuarios where usuario = '" & txtUsuario & "' and Senha = '" & txtSenha & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "select * from acessos where IDUsuario = " & TBUsuarios!IDUsuario & " and Acesso = '" & Modulo1 & "' and Validacao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = True Then
            USMsgBox ("Atenção usuário " & txtUsuario & ", você não tem autorização para este recurso."), vbExclamation, "CAPRIND v5.0"
            TBAcessos.Close
            Exit Sub
        End If
        TBAcessos.Close
        txtUsuario = TBUsuarios!Usuario
    Else
        USMsgBox ("Verifique se a senha e o usuário estão corretos."), vbExclamation, "CAPRIND v5.0"
        Exit Sub
    End If
    TBUsuarios.Close
Else
    txtUsuario = pubUsuario
End If

With frmVendas_analise
    Set TBVendas = CreateObject("adodb.recordset")
    TBVendas.Open "Select * from Vendas_analise where ID = " & .txtId, Conexao, adOpenKeyset, adLockOptimistic
    If TBVendas.EOF = False Then
        Select Case Formulario
            Case "Outros/Análise crítica/Engenharia":
                If IsNull(TBVendas!DtValidacao_Engenharia) = True Then
                    TBVendas!DtValidacao_Engenharia = Now
                    TBVendas!RespValidacao_Engenharia = txtUsuario
                    .txtDtValidacao_Engenharia = Now
                    .txtRespValidacao_Engenharia = txtUsuario
                    Evento = "Validar"
                Else
                    TBVendas!DtValidacao_Engenharia = Null
                    TBVendas!RespValidacao_Engenharia = Null
                    .txtDtValidacao_Engenharia = ""
                    .txtRespValidacao_Engenharia = ""
                    Evento = "Cancelar validação"
                End If
            Case "Outros/Análise crítica/Processos":
                If IsNull(TBVendas!DtValidacao_Processo) = True Then
                    TBVendas!DtValidacao_Processo = Now
                    TBVendas!RespValidacao_Processo = txtUsuario
                    .txtDtValidacao_processo = Now
                    .txtRespValidacao_processo = txtUsuario
                    Evento = "Validar"
                Else
                    TBVendas!DtValidacao_Processo = Null
                    TBVendas!RespValidacao_Processo = Null
                    .txtDtValidacao_processo = ""
                    .txtRespValidacao_processo = ""
                    Evento = "Cancelar validação"
                End If
            Case "Outros/Análise crítica/Qualidade":
                If IsNull(TBVendas!DtValidacao_Qualidade) = True Then
                    TBVendas!DtValidacao_Qualidade = Now
                    TBVendas!RespValidacao_Qualidade = txtUsuario
                    .txtDtValidacao_Qualidade = Now
                    .txtRespValidacao_Qualidade = txtUsuario
                    Evento = "Validar"
                Else
                    TBVendas!DtValidacao_Qualidade = Null
                    TBVendas!RespValidacao_Qualidade = Null
                    .txtDtValidacao_Qualidade = ""
                    .txtRespValidacao_Qualidade = ""
                    Evento = "Cancelar validação"
                End If
            Case "Outros/Análise crítica/Pcp":
                If IsNull(TBVendas!DtValidacao_Pcp) = True Then
                    TBVendas!DtValidacao_Pcp = Now
                    TBVendas!RespValidacao_Pcp = txtUsuario
                    .txtDtValidacao_PCP = Now
                    .txtRespValidacao_PCP = txtUsuario
                    Evento = "Validar"
                Else
                    TBVendas!DtValidacao_Pcp = Null
                    TBVendas!RespValidacao_Pcp = Null
                    .txtDtValidacao_PCP = ""
                    .txtRespValidacao_PCP = ""
                    Evento = "Cancelar validação"
                End If
            Case "Outros/Análise crítica/Compras":
                If IsNull(TBVendas!DtValidacao_Compras) = True Then
                    TBVendas!DtValidacao_Compras = Now
                    TBVendas!RespValidacao_Compras = txtUsuario
                    .txtDtValidacao_Compras = Now
                    .txtRespValidacao_Compras = txtUsuario
                    Evento = "Validar"
                Else
                    TBVendas!DtValidacao_Compras = Null
                    TBVendas!RespValidacao_Compras = Null
                    .txtDtValidacao_Compras = ""
                    .txtRespValidacao_Compras = ""
                    Evento = "Cancelar validação"
                End If
        End Select
        TBVendas.Update
    End If
    '==================================
    Modulo = Formulario
    ID_documento = .txtId
    Documento = "Nº análise: " & .Txt_analise & " - Rev.: " & .Txt_rev_analise
    Documento1 = ""
    ProcGravaEvento
    '==================================
End With
USMsgBox ("Operação de validação realizada com sucesso."), vbInformation, "CAPRIND v5.0"
If Validacao_obrig = True Then Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcBackup_Apontamento(Backup As Boolean)
On Error GoTo tratar_erro

If Backup = True Then
    'Salva dados na tabela ProducaoFases_backup
    Conexao.Execute "INSERT INTO ProducaoFases_Backup (Ordem, IDFASE, CodigoDesc, quantidade, Descricao, Fase, maquina, Usuario, TempoInicio, TempoFinal, TempoTotal, Pronto, Dias, Preparacao, Execucao, Data, Quant, Reprovada, OS, Turno, TempoTotalSeg, QTCD) Select Ordem, IDFASE, CodigoDesc, quantidade, Descricao, Fase, maquina, Usuario, TempoInicio, TempoFinal, TempoTotal, Pronto, Dias, Preparacao, Execucao, Data, Quant, Reprovada, OS, Turno, TempoTotalSeg, QTCD from ProducaoFases where Ordem = " & TBVendas!Ordem & " order by Tempoinicio"
    
    'Altera a ordem para backup
    Conexao.Execute "Update Producao Set Ap_backup = 'True' where Ordem = " & TBVendas!Ordem
    
    'Altera ID da produção na NC
    Conexao.Execute "Update CQNCF set CQNCF.idproducao = PFB.IDproducao from (CQ_NC_FABRICA CQNCF INNER JOIN ProducaoFases PF on CQNCF.IDproducao = PF.IDproducao) INNER JOIN ProducaoFases_Backup PFB ON PFB.OS = PF.OS and PFB.Tempoinicio = PF.Tempoinicio where PF.Ordem = " & TBVendas!Ordem & " and PF.Reprovada <> 0"
    
    'Altera ID da produção na manutenção
    Conexao.Execute "Update MD set MD.idproducao2 = PFB.IDproducao from (Manutencao_data MD INNER JOIN ProducaoFases PF on MD.IDproducao2 = PF.IDproducao) INNER JOIN ProducaoFases_Backup PFB ON PFB.OS = PF.OS and PFB.Tempoinicio = PF.Tempoinicio where PF.Ordem = " & TBVendas!Ordem
    
    'Apaga dados da tabela ProducaoFases
    Conexao.Execute "DELETE from ProducaoFases where Ordem = " & TBVendas!Ordem
    
    'Salva dados na tabela ProducaoFases_Totalizacao_backup
    Conexao.Execute "INSERT INTO ProducaoFases_Totalizacao_Backup (Ordem, OS, Fase, Data, Usuario, maquina, Turno, Pronto, Preparacao, Execucao, QTNC, QTOK, TPUTIL, TEUTIL, TETTUTIL, CRLOTE, CRPECA, CPLOTE, CPPECA, Eficiencia, Totalprod, Eficiencia_prep, Eficiencia_exec, Valor_hs_prep, Valor_hs_exec) Select Ordem, OS, Fase, Data, Usuario, maquina, Turno, Pronto, Preparacao, Execucao, QTNC, QTOK, TPUTIL, TEUTIL, TETTUTIL, CRLOTE, CRPECA, CPLOTE, CPPECA, Eficiencia, Totalprod, Eficiencia_prep, Eficiencia_exec, Valor_hs_prep, Valor_hs_exec from ProducaoFases_Totalizacao where Ordem = " & TBVendas!Ordem
        
    'Apaga dados da tabela ProducaoFases_Totalizacao
    Conexao.Execute "DELETE from ProducaoFases_Totalizacao where Ordem = " & TBVendas!Ordem
Else
    'Salva dados na tabela ProducaoFases ok
    Conexao.Execute "INSERT INTO ProducaoFases (Ordem, IDFASE, CodigoDesc, quantidade, Descricao, Fase, maquina, Usuario, TempoInicio, TempoFinal, TempoTotal, Pronto, Dias, Preparacao, Execucao, Data, Quant, Reprovada, OS, Turno, TempoTotalSeg) Select Ordem, IDFASE, CodigoDesc, quantidade, Descricao, Fase, maquina, Usuario, TempoInicio, TempoFinal, TempoTotal, Pronto, Dias, Preparacao, Execucao, Data, Quant, Reprovada, OS, Turno, TempoTotalSeg from ProducaoFases_Backup where Ordem = " & TBVendas!Ordem & " order by Tempoinicio"
    
    'Desmarca a ordem para backup ok
    Conexao.Execute "Update Producao Set Ap_backup = 'False' where Ordem = " & TBVendas!Ordem
    
    'Altera ID da produção na NC ok
    Conexao.Execute "Update CQNCF set CQNCF.idproducao = PF.IDproducao from (CQ_NC_FABRICA CQNCF INNER JOIN ProducaoFases_Backup PFB on CQNCF.IDproducao = PFB.IDproducao) INNER JOIN ProducaoFases PF ON PF.OS = PFB.OS and PF.Tempoinicio = PFB.Tempoinicio where PFB.Ordem = " & TBVendas!Ordem & " and PFB.Reprovada <> 0"
    
    'Altera ID da produção na manutenção ok
    Conexao.Execute "Update MD set MD.idproducao2 = PF.IDproducao from (Manutencao_data MD INNER JOIN ProducaoFases_Backup PFB on MD.IDproducao2 = PFB.IDproducao) INNER JOIN ProducaoFases PF ON PF.OS = PFB.OS and PF.Tempoinicio = PFB.Tempoinicio where PFB.Ordem = " & TBVendas!Ordem
    
    'Apaga dados da tabela ProducaoFases ok
    Conexao.Execute "DELETE from ProducaoFases_Backup where Ordem = " & TBVendas!Ordem
    
    'Salva dados na tabela ProducaoFases_Totalizacao
    Conexao.Execute "INSERT INTO ProducaoFases_Totalizacao (Ordem, OS, Fase, Data, Usuario, maquina, Turno, Pronto, Preparacao, Execucao, QTNC, QTOK, TPUTIL, TEUTIL, TETTUTIL, CRLOTE, CRPECA, CPLOTE, CPPECA, Eficiencia, Totalprod, Eficiencia_prep, Eficiencia_exec, Valor_hs_prep, Valor_hs_exec) Select Ordem, OS, Fase, Data, Usuario, maquina, Turno, Pronto, Preparacao, Execucao, QTNC, QTOK, TPUTIL, TEUTIL, TETTUTIL, CRLOTE, CRPECA, CPLOTE, CPPECA, Eficiencia, Totalprod, Eficiencia_prep, Eficiencia_exec, Valor_hs_prep, Valor_hs_exec from ProducaoFases_Totalizacao_Backup where Ordem = " & TBVendas!Ordem
        
    'Apaga dados da tabela ProducaoFases_Totalizacao_Backup
    Conexao.Execute "DELETE from ProducaoFases_Totalizacao_Backup where Ordem = " & TBVendas!Ordem
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCriarOS(Codinterno As String, Processo As Integer)
On Error GoTo tratar_erro

'pesquisa se existe alguma ordem com o mesmo desenho e sem ordem criada
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select P.IDPROCESSO, P.Quant, P.Desenho, P.Ordem, P.Prazoentrega from producao P LEFT JOIN ordemservico O on P.Ordem = O.Ordem where P.desenho = '" & Codinterno & "' and P.concluida = 'False' and O.IDproducao IS NULL and (P.Impof = 0 or P.Impof IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
Do While TBproducao.EOF = False
    TotalUtilizado = "00:00:00"
    QuantSolicitado = IIf(IsNull(TBproducao!Quant), 0, TBproducao!Quant)
    OF = IIf(IsNull(TBproducao!Ordem), 0, TBproducao!Ordem)
    
    'Salva ID do processo na ordem
    TBproducao!IDPROCESSO = Processo
    TBproducao.Update
    
    'Busca dados das fases do processo
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select * from fases where idprocesso = " & Processo & " AND versao = 'A' order by fase", Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        Do While TBFases.EOF = False
            'Cria OS
            Set TBProducaoFases = CreateObject("adodb.recordset")
            TBProducaoFases.Open "Select * from ordemservico", Conexao, adOpenKeyset, adLockOptimistic
            TBProducaoFases.AddNew
            If TBFases!Nao_aponta = True Then
                TBProducaoFases!Pronto = "SIM"
                TBProducaoFases!status = "Concluída"
                TBProducaoFases!DataConclusao = Date
            Else
                TBProducaoFases!Pronto = "NÃO"
                TBProducaoFases!status = "Aguardando"
            End If
            TBProducaoFases!Fase = TBFases!Fase
            TBProducaoFases!Rev_Fase = IIf(IsNull(TBFases!Revisao), 0, TBFases!Revisao)
            TBProducaoFases!Grupo_op = IIf(IsNull(TBFases!Grupo_op), "", TBFases!Grupo_op)
            TBProducaoFases!IDFase = TBFases!IDFase
            TBProducaoFases!maquina = TBFases!maquina
            TBProducaoFases!quantidade = QuantSolicitado
            
            DecimoSegundos = (IIf(IsNull(TBFases!TESegundos), 0, TBFases!TESegundos) * QuantSolicitado) + IIf(IsNull(TBFases!TPSegundos), 0, TBFases!TPSegundos)
            TBProducaoFases!TTLPREVS = DecimoSegundos 'Tempo total do lote previsto em segundos
            TBProducaoFases!TempoTotalLote = FormataTempo(DecimoSegundos) 'Tempo total do lote previsto
                    
            'Verifica se a maquina agrega custos/eficiencia na ordem
            Set TBMaquinas = CreateObject("adodb.recordset")
            TBMaquinas.Open "Select custos, PrecoHora_setup, PrecoHora from cadmaquinas where maquina = '" & TBFases!maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBMaquinas.EOF = False Then
                If TBMaquinas!custos = True Then TBProducaoFases!custos = True Else TBProducaoFases!custos = False
                If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then TBProducaoFases!Valor_hs_prep = TBMaquinas!PrecoHora_Setup Else TBProducaoFases!Valor_hs_prep = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
                TBProducaoFases!Valor_hs_exec = IIf(IsNull(TBMaquinas!PrecoHora), 0, TBMaquinas!PrecoHora)
            End If
            TBMaquinas.Close
            
            TBProducaoFases!IDPROCESSO = TBFases!IDPROCESSO
            TBProducaoFases!Ordem = OF
            
            TBProducaoFases!pecahora = TBFases!pecahora
            If TBFases!pecahora = True Then
                TBProducaoFases!Pcshora = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
            Else
                If IsNull(TBFases!Execucao) = False And TBFases!Execucao <> "00:00:00" Then
                    ElapsedTime (TBFases!Execucao)
                    TBProducaoFases!Pcshora = 3600 / s
                End If
            End If
            TBProducaoFases!pc_te = IIf(IsNull(TBFases!pc_te) = False, TBFases!pc_te, 1)
            TBProducaoFases!Preparacao = IIf(IsNull(TBFases!Preparacao), "00:00:00", TBFases!Preparacao)
            TBProducaoFases!Execucao = IIf(IsNull(TBFases!Execucao), "00:00:00", TBFases!Execucao)
            TBProducaoFases!TempoPreparacao = TBFases!TempoPreparacao
            TBProducaoFases!TempoExecucao = TBFases!TempoExecucao
            TBProducaoFases!descfase = TBFases!Descricao
            
            TBProducaoFases!TESegundos = FunCalculaSegPC(TBProducaoFases!Execucao, TBProducaoFases!pc_te)
            
            If OptOSControlada = 1 Then TBProducaoFases!OSControlada = True Else TBProducaoFases!OSControlada = False
            If Opt_processo_controlado = 1 Then TBProducaoFases!Processo_controlado = True Else TBProducaoFases!Processo_controlado = False
            TBProducaoFases!IDPlano = FunVerifIDPlano(TBFases!IDFase)
            TBProducaoFases.Update
            TBFases.MoveNext
        Loop
    End If
    
    ProcDefinirPrazosOS OF, TBproducao!PrazoEntrega, False
    ProcAcertaOS QuantSolicitado, False
    
    TBproducao.MoveNext
Loop
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcPlano(Codinterno As String, IDPlano As Long, IDFase As Long)
On Error GoTo tratar_erro

'pesquisa se existe alguma ordem com o mesmo desenho e sem apontamento
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select P.ordem from producao P LEFT JOIN producaofases PF on P.Ordem = PF.Ordem where P.desenho = '" & Codinterno & "' and P.concluida = 'False' and PF.IDproducao IS NULL and (P.Impof = 0 or P.Impof IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
Do While TBproducao.EOF = False
    Conexao.Execute "Update Ordemservico Set IDplano = " & IDPlano & " where IDfase = " & IDFase & " And Ordem = " & TBproducao!Ordem
    TBproducao.MoveNext
Loop
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcExcluirOS(Codinterno As String, Processo As Integer)
On Error GoTo tratar_erro

'pesquisa se existe a ordem com esse processo e se a ordem não estiver validada e não tiver apontamento ele exclui
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select P.IDPROCESSO, P.Ordem from producao P LEFT JOIN producaofases PF on P.Ordem = PF.Ordem where P.IDPROCESSO = " & Processo & " and P.concluida = 'False' and PF.IDproducao IS NULL and (P.Impof = 0 or P.Impof IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
Do While TBproducao.EOF = False
    TBproducao!IDPROCESSO = 0
    TBproducao.Update
    OF = TBproducao!Ordem
    Conexao.Execute "DELETE from Ordemservico where Ordem = " & OF
    'Variavel OF usa nesta proc
    ProcAcertaOS QuantSolicitado, False
    
    TBproducao.MoveNext
Loop
TBproducao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procNotaFiscal(Modelo As String)
On Error GoTo tratar_erro

If IsNull(TBVendas!int_NotaFiscal) = True Or TBVendas!int_NotaFiscal = "" Then
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSql = "Select CAST(int_NotaFiscal AS int) AS NF, Serie, Modelo FROM tbl_Dados_Nota_Fiscal where tipoNF = '" & TBVendas!TipoNF & "' and Aplicacao = 'P' and Serie = '" & NF_Serie & "' and Modelo = '" & Modelo & "' and ID_empresa = " & TBVendas!ID_empresa & " and int_NotaFiscal IS NOT NULL order by dt_DataEmissao desc, NF desc"
    'Debug.print StrSql
    
    TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        QuantsolicitadoN1 = TBAbrir!NF + 1
        FamiliaAntiga = QuantsolicitadoN1
        Familiatext = FunTamanhoTextoZeroEsq(FamiliaAntiga, 9)
        SerieNF = IIf(IsNull(TBAbrir!Serie), 1, TBAbrir!Serie)
    Else
        Familiatext = "000000001"
        SerieNF = 1
    End If
    TBAbrir.Close
    
    TBVendas!dt_DataEmissao = Date
    TBVendas!Hora_emissao = Format(Now, "hh:mm:ss")
    TBVendas!int_NotaFiscal = FunVerifExisteNumNF(TBVendas!TipoNF, TBVendas!ID_empresa, Familiatext, SerieNF, TBVendas!Modelo)
    TBVendas!Serie = SerieNF
    TBVendas!dt_Saida_Entrada = Date
    TBVendas!txt_Hora_Saida = Format(Now, "hh:mm:ss")
Else
    Familiatext = TBVendas!int_NotaFiscal
    SerieNF = TBVendas!Serie
End If

If TBVendas!TipoNF = "SA" And TBVendas!int_TipoNota = 1 And IsNull(SerieNF) = False And IsNull(TBVendas!RPS) = True Then
    Set TBCarteira = CreateObject("adodb.recordset")
    TBCarteira.Open "Select RPS from tbl_dados_nota_fiscal where Serie = '" & SerieNF & "' AND RPS IS NOT NULL AND ID_empresa = " & TBVendas!ID_empresa & " ORDER BY RPS DESC", Conexao, adOpenKeyset, adLockReadOnly
    If TBCarteira.EOF = False Then
        TBVendas!RPS = TBCarteira!RPS + 1
    Else
        TBVendas!RPS = 1
    End If
    TBCarteira.Close
End If

Conexao.Execute "Update tbl_Dados_Transp Set int_NotaFiscal = '" & TBVendas!int_NotaFiscal & "' where ID_Nota = " & TBVendas!ID
Conexao.Execute "Update Tbl_DadosAdicionais Set IDNotaFiscal = '" & TBVendas!int_NotaFiscal & "' where ID_Nota = " & TBVendas!ID
Conexao.Execute "Update tbl_Detalhes_Nota Set int_NotaFiscal = '" & TBVendas!int_NotaFiscal & "' where ID_Nota = " & TBVendas!ID
Conexao.Execute "Update tbl_Detalhes_Recebimento Set int_NotaFiscal = '" & TBVendas!int_NotaFiscal & "' where ID_Nota = " & TBVendas!ID
Conexao.Execute "Update tbl_proposta_nota Set NF = '" & TBVendas!int_NotaFiscal & "' where ID_Nota = " & TBVendas!ID
Conexao.Execute "Update tbl_Totais_Nota Set int_NotaFiscal = '" & TBVendas!int_NotaFiscal & "' where ID_Nota = " & TBVendas!ID

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAlteraStatusPI()
On Error GoTo tratar_erro

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select * from tbl_Detalhes_Nota_pedidos where ID_nota = " & TBVendas!ID & " and ID_carteira is not null and ID_carteira <> 0", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    Do While TBTempo.EOF = False
        frmFaturamento_Prod_Serv.ProcGravarQtdeFaturadaProdServ TBTempo!ID_carteira, TBTempo!Codinterno
        TBTempo.MoveNext
    Loop
End If
TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub ProcCarregaLista(Modulo1 As String)
On Error GoTo tratar_erro

If Left(Modulo1, 23) = "Faturamento/Nota fiscal" Or Modulo1 = "Estoque/Ordem de faturamento" Or Modulo1 = "Estoque/Nota fiscal" Then
    Permitido1 = False
    Permitido2 = False
    If Modulo1 <> "Estoque/Ordem de faturamento" Then
        With frmFaturamento_Prod_Serv
        
            Set TBVendas = CreateObject("adodb.recordset")
            TBVendas.Open "Select * from tbl_Dados_Nota_Fiscal where DtValidacao IS NOT NULL and (" & TextoFiltroFin & ")", Conexao, adOpenKeyset, adLockOptimistic
            If TBVendas.EOF = False Then
                Do While TBVendas.EOF = False
                    .ProcCarregaDadosNota TBVendas!ID
                    .ProcCarregaListaDuplicatas TBVendas!ID
                    .ProcCarregaTotaisNota TBVendas!ID
                    
                    'Se não tiver dados adicionais ele cria puxando da CFOP
                    Set TBCodigoDesc = CreateObject("adodb.recordset")
                    TBCodigoDesc.Open "Select * from tbl_dadosadicionais where id_nota = " & TBVendas!ID, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCodigoDesc.EOF = False And IsNull(TBCodigoDesc!mem_DadosAdicionais) = True Then
                        TBCodigoDesc!mem_DadosAdicionais = .FunCarregaCampoDACFOP(TBVendas!ID)
                        TBCodigoDesc.Update
                    ElseIf TBCodigoDesc.EOF = True Then
                        TBCodigoDesc.AddNew
                        TBCodigoDesc!IDNotaFiscal = TBVendas!int_NotaFiscal
                        TBCodigoDesc!ID_nota = TBVendas!ID
                        TBCodigoDesc!mem_DadosAdicionais = .FunCarregaCampoDACFOP(TBVendas!ID)
                        TBCodigoDesc.Update
                    End If
                    TBCodigoDesc.Close
                    
                    If .lst_Duplicata.ListItems.Count <> 0 And .txtIDcliente <> "0" Then
                        If Permitido1 = False Then If USMsgBox("Deseja enviar a(s) duplicata(s) para o financeiro agora?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido2 = True Else Permitido2 = False
                        If Permitido2 = True Then .ProcEnviarDupFinanceiro TBVendas!ID, False
                        Permitido1 = True
                    End If
                    TBVendas.MoveNext
                Loop
                If Permitido2 = True Then USMsgBox ("Nova(s) duplicata(s) enviada(s) para o financeiro com sucesso."), vbInformation, "CAPRIND v5.0"
            End If
            
            If .txtId <> "" Then .ProcCarregaDadosNota IIf(.txtId = "", 0, .txtId)
            .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
        End With
    Else
        With frmEstoque_Ordem_Faturamento
        
            Set TBVendas = CreateObject("adodb.recordset")
            TBVendas.Open "Select * from tbl_Dados_Nota_Fiscal where DtValidacao IS NOT NULL and (" & TextoFiltroFin & ")", Conexao, adOpenKeyset, adLockOptimistic
            If TBVendas.EOF = False Then
                Do While TBVendas.EOF = False
                    .ProcCarregaDadosNota TBVendas!ID
                    .ProcCarregaListaDuplicatas TBVendas!ID
                    .ProcCarregaTotaisNota TBVendas!ID
                    
                    'Se não tiver dados adicionais ele cria puxando da CFOP
                    Set TBCodigoDesc = CreateObject("adodb.recordset")
                    TBCodigoDesc.Open "Select * from tbl_dadosadicionais where id_nota = " & TBVendas!ID, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCodigoDesc.EOF = False And IsNull(TBCodigoDesc!mem_DadosAdicionais) = True Then
                        TBCodigoDesc!mem_DadosAdicionais = .FunCarregaCampoDACFOP(TBVendas!ID)
                        TBCodigoDesc.Update
                    ElseIf TBCodigoDesc.EOF = True Then
                        TBCodigoDesc.AddNew
                        TBCodigoDesc!IDNotaFiscal = TBVendas!int_NotaFiscal
                        TBCodigoDesc!ID_nota = TBVendas!ID
                        TBCodigoDesc!mem_DadosAdicionais = .FunCarregaCampoDACFOP(TBVendas!ID)
                        TBCodigoDesc.Update
                    End If
                    TBCodigoDesc.Close
                    
                    If .lst_Duplicata.ListItems.Count <> 0 Then
                        If Permitido1 = False Then If USMsgBox("Deseja enviar a(s) duplicata(s) para o financeiro agora?", vbYesNo, "CAPRIND v5.0") = vbYes Then Permitido2 = True Else Permitido2 = False
                        If Permitido2 = True Then .ProcEnviarDupFinanceiro TBVendas!ID, False
                        Permitido1 = True
                    End If
                    TBVendas.MoveNext
                Loop
                If Permitido2 = True Then USMsgBox ("Nova(s) duplicata(s) enviada(s) para o financeiro com sucesso."), vbInformation, "CAPRIND v5.0"
            End If
            
            If .txtId <> "" Then .ProcCarregaDadosNota IIf(.txtId = "", 0, .txtId)
            .ProcCarregaListaNota (IIf(ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(1).Caption, Len(.lblPaginas(1).Caption) - 5))))
        End With
    
     
    End If
    
ElseIf Right(Modulo1, 37) = "Produtos e serviços/Validar estrutura" Or Right(Modulo1, 45) = "Produtos e serviços/Validar plano de inspeção" Then
        With frmproj_produto
            .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select * From projproduto where codproduto = " & IIf(.txtcodproduto = "", 0, .txtcodproduto), Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                .ProcCarregaDados
            End If
        End With
    Else
        Select Case Modulo1
            Case "Vendas/Famílias":
                With frmproj_familia
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from projfamilia where Codigo = " & IIf(.txtid_familia = "", 0, .txtid_familia), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then
                        .ProcLimpaCampos
                        .ProcPuxaDados
                    End If
                End With
            Case "Vendas/Clientes":
                With frmVendas_cliente
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBClientes = CreateObject("adodb.recordset")
                    TBClientes.Open "Select * from clientes where idcliente = " & IIf(.txtIDcliente = "", 0, .txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
                    If TBClientes.EOF = False Then
                        .txtIDcliente = ""
                        .txtIDcliente = TBClientes!IDCliente
                    End If
                End With
            Case "Compras/Famílias":
                With frmproj_familia
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from projfamilia where Codigo = " & IIf(.txtid_familia = "", 0, .txtid_familia), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then
                        .ProcLimpaCampos
                        .ProcPuxaDados
                    End If
                End With
            Case "Compras/Fornecedores"
                With frmCompras_fornecedores
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBFornecedor = CreateObject("adodb.recordset")
                    TBFornecedor.Open "Select * from compras_fornecedores where idcliente = " & IIf(.txtIDcliente = "", 0, .txtIDcliente), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFornecedor.EOF = False Then
                        .txtIDcliente = ""
                        .txtIDcliente = TBFornecedor!IDCliente
                    End If
                End With
            Case "Compras/Pedido":
                With frmCompras_Pedido
                    .ProcAtualizalistapedido (IIf(ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5))))
                    Set TBCompras_Pedido = CreateObject("adodb.recordset")
                    TBCompras_Pedido.Open "Select * from compras_pedido where IDpedido = " & IIf(.txtIDPedido = "", 0, .txtIDPedido), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras_Pedido.EOF = False Then
                        .ProcPuxaDados
                    End If
                    TBCompras_Pedido.Close
                End With
            Case "Compras/Pedido/Aprovar":
                With frmCompras_Pedido
                    .ProcAtualizalistapedido (IIf(ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas(3).Caption, Len(.lblPaginas(3).Caption) - 5))))
                    Set TBCompras_Pedido = CreateObject("adodb.recordset")
                    TBCompras_Pedido.Open "Select * from compras_pedido where IDpedido = " & IIf(.txtIDPedido = "", 0, .txtIDPedido), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras_Pedido.EOF = False Then
                        .ProcPuxaDados
                    End If
                    TBCompras_Pedido.Close
                End With
            Case "Compras/Produtos e serviços":
                With frmproj_produto
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * From projproduto where codproduto = " & IIf(.txtcodproduto = "", 0, .txtcodproduto), Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        .ProcCarregaDados
                    End If
                End With
            Case "Faturamento/Fiscal/Natureza de operação":
                With frm_Natureza_OP
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBLISTA = CreateObject("adodb.recordset")
                    TBLISTA.Open "Select * From tbl_NaturezaOperacao where IDCountCfop = " & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBLISTA.EOF = False Then
                        .ProcCarregaDados
                    End If
                    TBLISTA.Close
                End With
            Case "Faturamento/Fiscal/Classificação fiscal":
                With frm_Classificacao_Fiscal
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBLISTA = CreateObject("adodb.recordset")
                    TBLISTA.Open "Select * From tbl_ClassificacaoFiscal where Idclass = " & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBLISTA.EOF = False Then
                        .ProcCarregaDados
                    End If
                    TBLISTA.Close
                End With
            Case "Vendas/Produtos e serviços":
                With frmproj_produto
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * From projproduto where codproduto = " & IIf(.txtcodproduto = "", 0, .txtcodproduto), Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        .ProcCarregaDados
                    End If
                End With
            Case "Engenharia/Famílias":
                With frmproj_familia
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from projfamilia where Codigo = " & IIf(.txtid_familia = "", 0, .txtid_familia), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then
                        .ProcLimpaCampos
                        .ProcPuxaDados
                    End If
                End With
            Case "Engenharia/Produtos e serviços":
                With frmproj_produto
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * From projproduto where codproduto = " & IIf(.txtcodproduto = "", 0, .txtcodproduto), Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        .ProcCarregaDados
                    End If
                End With
            Case "Engenharia/Processos":
                With frmProcessos
                    .ProcCarregaListaProcessos (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBProcessos = CreateObject("adodb.recordset")
                    TBProcessos.Open "Select * From processos where IDProcesso = " & IIf(.txtidprocesso = "", 0, .txtidprocesso), Conexao, adOpenKeyset, adLockOptimistic
                    If TBProcessos.EOF = False Then
                        .ProcPuxaDados
                    End If
                End With
            Case "Estoque/Local de armazenamento":
                With frmEstoque_Localarmaz
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select * from Estoque_Localarmazenamento_criar where ID = " & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        .ProcPuxaDados
                    End If
                End With
            Case "Estoque/Inventário":
                With frmestoque_fisico
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Estoque_fisico.*, projproduto.Desenho, projproduto.Descricao, projproduto.Unidade, projproduto.Classe from Estoque_fisico inner join projproduto on Estoque_fisico.Codproduto = projproduto.Codproduto where Estoque_fisico.ID = " & .Txt_ID, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        .ProcPuxaDados
                    End If
                End With
            Case "Outros/Solicitação":
                With frmCompras_Requisicao
                    .ProcCarregaLista_Req (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from compras_requisicao where ID_requisicao = " & IIf(.Txt_ID_req = "", 0, .Txt_ID_req), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then
                        .ProcAbrir
                    End If
                    TBCompras.Close
                End With
            Case "Outros/Solicitação/Autorizar solicitação":
                With frmCompras_Requisicao
                    .ProcCarregaLista_Req (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from compras_requisicao where ID_requisicao = " & IIf(.Txt_ID_req = "", 0, .Txt_ID_req), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then
                        .ProcAbrir
                    End If
                    TBCompras.Close
                End With
            Case "Qualidade/Famílias":
                With frmproj_familia
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBFamilia = CreateObject("adodb.recordset")
                    TBFamilia.Open "Select * from projfamilia where Codigo = " & IIf(.txtid_familia = "", 0, .txtid_familia), Conexao, adOpenKeyset, adLockOptimistic
                    If TBFamilia.EOF = False Then
                        .ProcLimpaCampos
                        .ProcPuxaDados
                    End If
                End With
            Case "Qualidade/Inspeção de recebimento":
                With frmCompras_recebimento
                    .ProcCarregaListaInspecionados
                    .ProcCarregaDadosInsp .txtId
                End With
            Case "Qualidade/Não conformidade/Descrição da não conformidade":
                With frmcqnc_descricaoNC
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBLISTA = CreateObject("adodb.recordset")
                    TBLISTA.Open "Select * from CQ_NC_FABRICA_causa where ID = " & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBLISTA.EOF = False Then
                        .ProcLimpaCampos
                        .ProcCarregaDados
                    End If
                End With
            Case "Qualidade/PPAP/PSW":
                With frmQualidadePPAP
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from QualidadePPAP where IDPPAP = " & IIf(.txtIDPPAP = "", 0, .txtIDPPAP), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .ProcPuxaDadosPPAP
                    End If
                    TBAbrir.Close
                End With
            Case "Vendas/Proposta comercial":
                With frmVendas_proposta
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBAbrir = CreateObject("adodb.recordset")
                     StrSql = "Select VP.*, CL.CPF_CNPJ as CNPJ_CPF, CL.RG_IE as RG_IE, CL.CEP as CEP from vendas_proposta VP inner join Clientes CL on VP.IDcliente = CL.IDCliente where cotacao ="
                     TBAbrir.Open StrSql & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .ProcPuxaDados
                    End If
                    TBAbrir.Close
                End With
            Case "Vendas/Pedido interno":
                With frmVendas_PI
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBAbrir = CreateObject("adodb.recordset")
                     StrSql = "Select VP.*, CL.CPF_CNPJ as CNPJ_CPF, CL.RG_IE as RG_IE, CL.CEP as CEP from vendas_proposta VP inner join Clientes CL on VP.IDcliente = CL.IDCliente where cotacao ="
                     TBAbrir.Open StrSql & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                  
                    If TBAbrir.EOF = False Then
                        .ProcPuxaDados
                    End If
                    TBAbrir.Close
                End With
            Case "PCP/Gerenciamento de ordem":
                With frmprod
                    .atualiza_lista_ordens (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Ordem = IIf(.txtof = "", 0, .txtof)
                    .ProcCarregaOrdem
                End With
            Case "PCP/Plano da produção":
                With frmPlano_producao
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from ProducaoFases_OS where ID = " & IIf(.Txt_ID = "", 0, .Txt_ID), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then .ProcCarregaDadosPlano
                    TBCompras.Close
                End With
            Case "Engenharia/Normas":
                With frmNorma
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Norma where id = " & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then .ProcPuxaDados
                    TBAbrir.Close
                End With
            Case "RH/Funcionários":
                With frmRH_Funcionarios
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Funcionarios where id = " & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then .ProcPuxaDados
                    TBAbrir.Close
                End With
            Case "Financeiro/Instituições":
                With frm_Instituicoes
                    .ProcCarregaLista
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from tbl_Instituicoes where id = " & IIf(.txtCodBanco = "", 0, .txtCodBanco), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then .ProcCarregaDados
                    TBAbrir.Close
                End With
            Case "Qualidade/Plano de inspeção":
                With frmPlanoinspecao
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBplano = CreateObject("adodb.recordset")
                    TBplano.Open "Select * from Plano where IDPlano = " & IIf(.txtPI = "", 0, .txtPI), Conexao, adOpenKeyset, adLockOptimistic
                    If TBplano.EOF = False Then .ProcCarregaDados
                    TBplano.Close
                End With
            Case "Estoque/Requisição de materiais":
                With frmRequisicao_materiais
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from Requisicao_materiais where id = " & .txtId, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then .ProcPuxaDados
                    TBAbrir.Close
                End With
            Case "Qualidade/Solicitação de ação":
                With frmCQ_SA
                    .ProcCarregaLista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from CQ_SA where id = " & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then .ProcPuxaDados
                    TBAbrir.Close
                End With
            Case "Outros/Solicitação de produção":
                With frmOutros_Solicitacao_PCP
                    .ProcCarregalista_Solicitacao (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from Outros_SolicitacaoPCP where ID = " & IIf(.Txt_ID_req = "", 0, .Txt_ID_req), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then .ProcAbrir
                    TBCompras.Close
                End With
            Case "Outros/Solicitação de produção/Autorizar solicitação":
                With frmOutros_Solicitacao_PCP
                    .ProcCarregalista_Solicitacao (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from Outros_SolicitacaoPCP where ID = " & IIf(.Txt_ID_req = "", 0, .Txt_ID_req), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then .ProcAbrir
                    TBCompras.Close
                End With
            Case "Vendas/Vendedores":
                With frmVendas_Vendedores
                    .ProcAtualizalista (IIf(ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5)) <= 1, 1, ReturnNumbersOnly(Left(.lblPaginas.Caption, Len(.lblPaginas.Caption) - 5))))
                    Set TBCompras = CreateObject("adodb.recordset")
                    TBCompras.Open "Select * from Vendas_Vendedores where ID = " & IIf(.txtId = "", 0, .txtId), Conexao, adOpenKeyset, adLockOptimistic
                    If TBCompras.EOF = False Then .ProcPuxaDados
                    TBCompras.Close
                End With
        End Select
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEmpenharProdServEst()
On Error GoTo tratar_erro


Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select Qtde_produzir, CODIGO, Desenho, Qtde_produzir, PrazoFinal from vendas_carteira where cotacao = " & TBVendas!Cotacao & " order by PrazoFinal", Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    Do While TBCotacao.EOF = False
      If TBCotacao!Qtde_produzir <> "" Or IsNull(TBCotacao!Qtde_produzir) = False Then
        QuantSolicitado = TBCotacao!Qtde_produzir
        ProcEmpenharProdEstoque TBVendas!ID_empresa, TBCotacao!CODIGO, TBCotacao!Desenho, True, False, TBCotacao!Qtde_produzir
        If QuantSolicitado > 0 Then ProcEmpenharProdProduzindo TBVendas!ID_empresa, TBCotacao!CODIGO, TBCotacao!Desenho, TBCotacao!PrazoFinal, True
     End If
     TBCotacao.MoveNext

    Loop
End If
TBCotacao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGerarNecessidadePI()
On Error GoTo tratar_erro

Set TBCotacao = CreateObject("adodb.recordset")
TBCotacao.Open "Select Qtde_produzir, CODIGO, Desenho, Qtde_produzir, PrazoFinal from vendas_carteira where cotacao = " & TBVendas!Cotacao & " order by PrazoFinal", Conexao, adOpenKeyset, adLockOptimistic
If TBCotacao.EOF = False Then
    VersaoEstrutura = ""
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select PCDV.Versao from Projproduto P INNER JOIN Projconjunto_desc_versao PCDV ON PCDV.Codproduto = P.Codproduto where P.Desenho = '" & TBCotacao!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        If TBItem.RecordCount = 1 Then VersaoEstrutura = TBItem!versao
    End If
    VersaoProcesso = ""
    Set TBItem = CreateObject("adodb.recordset")
    TBItem.Open "Select PRO.IDProcesso from Projproduto P INNER JOIN Processos PRO ON PRO.Codproduto = P.Codproduto where P.Desenho = '" & TBCotacao!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBItem.EOF = False Then
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select Versao, COUNT(Versao) as NV from Fases where IDprocesso = " & TBItem!IDPROCESSO & " group by Versao", Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            If TBFases!NV = 1 Then VersaoProcesso = TBFases!versao
        End If
        TBFases.Close
    End If
    TBItem.Close
    Conexao.Execute "Update Vendas_Carteira Set Versao_estrutura = '" & IIf(VersaoEstrutura = "", Null, VersaoEstrutura) & "', Versao_processo = '" & IIf(VersaoProcesso = "", Null, VersaoProcesso) & "' where Codigo = " & TBCotacao!CODIGO
    
    Conexao.Execute "DELETE from Producaomaterial where ID_carteira = " & TBCotacao!CODIGO
    If VersaoEstrutura <> "" Then
        'Verifica qtde empenhada
        QuantSolicitado = 0
        Set TBLISTA = CreateObject("adodb.recordset")
        TBLISTA.Open "Select VC.Qtde_produzir - ISNULL(ECEV.Qtde_empenhada, 0) AS Quantsolicitado from Vendas_carteira VC LEFT JOIN Estoque_Controle_Empenho_Vendas ECEV ON ECEV.ID_carteira = VC.Codigo where VC.Codigo = " & TBCotacao!CODIGO, Conexao, adOpenKeyset, adLockOptimistic
        If TBLISTA.EOF = False Then
            Familiatext = VersaoEstrutura
            ProcAcertaRequisicao TBCotacao!Desenho, 0, TBCotacao!CODIGO, Familiatext, IIf(IsNull(TBLISTA!QuantSolicitado), 0, TBLISTA!QuantSolicitado), True
        End If
        TBLISTA.Close
    End If
End If
TBCotacao.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procRetirarEstoqueNF()
On Error GoTo tratar_erro


'Verifica se movimenta estoque
Set TBNivel10 = CreateObject("adodb.recordset")
StrSql = "Select TBL.Int_codigo, TBL.int_Cod_Produto, TBL.int_Qtd, TBL.txt_Unid, TBL.Unidade_com, TBL.qtde_estoque, TBL.ID_nota, TBL.N_Referencia, TBL.Int_NotaFiscal, P.ID_PC, P.Codproduto from (tbl_detalhes_nota TBL INNER JOIN ProjProduto P ON TBL.Codproduto = P.Codproduto) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = TBL.ID_CFOP where TBL.ID_Nota = " & TBVendas!ID & " and P.Estoque = 'True' and (CFOP.Remessa IS NULL or CFOP.Remessa = 'False')and (CFOP.Retorno IS NULL or CFOP.Retorno = 'False')"
'Debug.print StrSql
'=================================================================
' Busca o item na lista da nota com CFOP que não seja remessa e nem retorno
'=================================================================
TBNivel10.Open "Select TBL.Int_codigo, TBL.int_Cod_Produto, TBL.int_Qtd, TBL.txt_Unid, TBL.Unidade_com, TBL.qtde_estoque, TBL.ID_nota, TBL.N_Referencia, TBL.Int_NotaFiscal, P.ID_PC, P.Codproduto from (tbl_detalhes_nota TBL INNER JOIN ProjProduto P ON TBL.Codproduto = P.Codproduto) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = TBL.ID_CFOP where TBL.ID_Nota = " & TBVendas!ID & " and P.Estoque = 'True' and (CFOP.Remessa IS NULL or CFOP.Remessa = 'False')and (CFOP.Retorno IS NULL or CFOP.Retorno = 'False')", Conexao, adOpenKeyset, adLockOptimistic
Do While TBNivel10.EOF = False

    If TBNivel10!txt_Unid <> TBNivel10!Unidade_com And IsNull(TBNivel10!Qtde_estoque) = False And TBNivel10!Qtde_estoque <> 0 Then
        qtdeliberada = IIf(IsNull(TBNivel10!Qtde_estoque), 0, TBNivel10!Qtde_estoque)
    Else
        qtdeliberada = IIf(IsNull(TBNivel10!int_Qtd), 0, TBNivel10!int_Qtd)
    End If
'======================================================================
' Só retira do estoque se estiver empenhado
'======================================================================
'Verifica se tem empenho no estoque
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select EC.*, EE.Qtde_empenhada - EE.Qtde_saida as qtdeEstoqueMin from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where NFPP.ID_prod_NF = " & TBNivel10!Int_codigo & " and EC.ID_empresa = " & frmFaturamento_Prod_Serv.txtIDEmpresa.Text & " and EC.desenho = '" & TBNivel10!int_Cod_Produto & "' and EC.Lote is not null and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') and EE.Qtde_empenhada - EE.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
    StrSql = "Select EC.*, EE.Qtde_empenhada - EE.Qtde_saida as qtdeEstoqueMin from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where NFPP.ID_prod_NF = " & TBNivel10!Int_codigo & " and EC.ID_empresa = " & frmFaturamento_Prod_Serv.txtIDEmpresa.Text & " and EC.desenho = '" & TBNivel10!int_Cod_Produto & "' and EC.Lote is not null and (Left(EC.status, 7) = 'ENTRADA' or EC.status = 'CONSIGNAÇÃO RECEBIDA') and EE.Qtde_empenhada - EE.Qtde_saida > 0"
    'Debug.print StrSql
    If TBEstoque.EOF = False Then
        EmpenhoVerificar = True
          If USMsgBox("Deseja baixar estoque com essa nota fiscal?", vbYesNo, "CAPRIND v5.0") = vbNo Then
          Exit Sub
          End If
        procRetirarEstoqueNF1
    Else
Estoque:
        EmpenhoVerificar = False
        Set TBCFOP = CreateObject("adodb.recordset")
        TBCFOP.Open "Select QPP.Ordem from (Qtde_produzindo_produto QPP INNER JOIN Producao_pedidos PP ON PP.Ordem = QPP.Ordem) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.IDcarteira where NFPP.ID_prod_NF = " & TBNivel10!Int_codigo & "", Conexao, adOpenKeyset, adLockOptimistic
        If TBCFOP.EOF = True Then
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select EP.* from Estoque_produtos EP INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz where EP.ID_empresa = " & TBVendas!ID_empresa & " and EP.Desenho = '" & TBNivel10!int_Cod_Produto & "' and EP.Estoque_real > 0 and EP.Liberado = 'SIM' and EP.Consignacao = 'False' AND EL.Estoque = 'False' order by Data, IdEstoque", Conexao, adOpenKeyset, adLockOptimistic
'            TBEstoque.Open "Select EP.* from Estoque_produtos EP where EP.ID_empresa = " & TBVendas!ID_empresa & " and EP.Desenho = '" & TBNivel10!int_Cod_Produto & "' and EP.Estoque_real > 0 and EP.Liberado = 'SIM' and EP.Consignacao = 'False' order by Data, IdEstoque", Conexao, adOpenKeyset, adLockOptimistic
            StrSql = "Select EP.* from Estoque_produtos EP INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz where EP.ID_empresa = " & TBVendas!ID_empresa & " and EP.Desenho = '" & TBNivel10!int_Cod_Produto & "' and EP.Estoque_real > 0 and EP.Liberado = 'SIM' and EP.Consignacao = 'False' AND EL.Estoque = 'False' order by Data, IdEstoque"
            'Debug.print StrSql
            
            If TBEstoque.EOF = False Then
                procRetirarEstoqueNF1
            Else
            USMsgBox "Não existe saldo disponivel para baixa no estoque do item código: " & TBNivel10!int_Cod_Produto, vbCritical, "CAPRIND v5.0"
            End If
        End If
        TBCFOP.Close
    End If
    
    'Se acabar os empenhos e ainda tiver saldo precisa ver do estoque normal
    If TBEstoque.EOF = True And qtdeliberada > 0 And EmpenhoVerificar = True Then GoTo Estoque
    
    TBEstoque.Close
    TBNivel10.MoveNext
Loop
TBNivel10.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub procRetirarEstoqueNF1()
On Error GoTo tratar_erro

CampoFiltro = "Saida"
'qtdeliberar = qtdeliberada
Qtd = 0
Do While TBEstoque.EOF = False And qtdeliberada > 0
    
    valor = 0
    'Verifica se este RE já está empenhado
    TextoFiltro = ""
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select ID_carteira from tbl_Detalhes_Nota_pedidos where ID_prod_NF = " & TBNivel10!Int_codigo & " and Codinterno = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        Do While TBFIltro.EOF = False
            If TextoFiltro = "" Then TextoFiltro = "and ID_carteira <> " & TBFIltro!ID_carteira Else TextoFiltro = TextoFiltro & " and ID_carteira <> " & TBFIltro!ID_carteira
            TBFIltro.MoveNext
        Loop
    End If
    TBFIltro.Close
            
    Qtd = TBEstoque!estoque_real
    
'===============================================================
'Verifica se este RE já está empenhado
'===============================================================
' Material consignado
'===============================================================
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(PNFC.Quantidade - ISNULL(PNFC.Qtde_saida, 0)) as Valor from (Producao_NF_Consignada PNFC INNER JOIN Producaomaterial PM ON PM.Ordem = PNFC.Ordem and PM.Codigo = PNFC.Codinterno) INNER JOIN Producao P ON P.Ordem = PNFC.Ordem where PNFC.IDestoque = " & TBEstoque!IDEstoque & " and PNFC.Quantidade - ISNULL(PNFC.Qtde_saida, 0) > 0 and P.Status <> 'Cancelada' and P.Concluida = 0 and (PM.Saida = 'NÃO' OR PM.Saida = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtd = Qtd - IIf(IsNull(TBFI!valor), 0, TBFI!valor)
    End If
'===============================================================
' Produto acabado (Vendas)
'===============================================================
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select Sum(Qtde_empenhada - ISNULL(Qtde_saida, 0)) as Valor from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBEstoque!IDEstoque & " " & TextoFiltro & " and Qtde_empenhada - ISNULL(Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = False Then
        Qtd = Qtd - IIf(IsNull(TBFI!valor), 0, TBFI!valor)
    End If
    TBFI.Close
    
    If Qtd <= 0 Then GoTo Proximo
    Permitido2 = True
            
    If EmpenhoVerificar = True Then
        QtdeSaida = TBEstoque!qtdeEstoqueMin
    Else
        If qtdeliberada >= Qtd Then QtdeSaida = Qtd Else QtdeSaida = qtdeliberada
    End If
    QtdeSaidaPC = FunCalculaQtdePCKG(TBEstoque!estoque_real, IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC), QtdeSaida, True)
    
    Set TBMateriaprima = CreateObject("adodb.recordset")
    TBMateriaprima.Open "Select * from tbl_Detalhes_Nota_pedidos where ID_prod_NF = " & TBNivel10!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
 
 '=========================================================================
 'Se existe vinculo da nota com pedido interno, atualiza saida no pedido
 '=========================================================================
    If TBMateriaprima.EOF = False Then
    ProcAtualizaQtdeExpProdPed TBMateriaprima!ID_prod_NF, TBMateriaprima!Codinterno, QtdeSaida, TBEstoque!LOTE, TBEstoque!IDEstoque, Date
    End If
    
    If EmpenhoVerificar = True Then
        QtdeSaida = TBEstoque!qtdeEstoqueMin
    Else
        If qtdeliberada >= Qtd Then QtdeSaida = Qtd Else QtdeSaida = qtdeliberada
    End If
    
    'If qtdeliberada >= Qtd Then QtdeSaida = Qtd Else QtdeSaida = qtdeliberada
    
    If FunVerifCodRefCliDANFE(frmFaturamento_Prod_Serv.txtIDEmpresa.Text) = True Then
        Conexao.Execute "UPDATE Estoque_controle Set REF = '" & TBNivel10!N_referencia & "' where IDestoque = " & TBEstoque!IDEstoque
    End If
    
 '===========================================================================
 'Acerta saldo e valores (Valor total) na RE
 '===========================================================================
  Set TBProduto = CreateObject("adodb.recordset")
  TBProduto.Open "Select * from projproduto where Desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
   If TBProduto.EOF = False Then
   
   QuantEmpenho = Format(TBEstoque!estoque_real - QtdeSaida, "###,##0.0000")
   QuantEmpenhoPC = IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC) - QtdeSaidaPC
   NovoValor = Replace(QuantEmpenho, ",", ".")
   NovoValor1 = Replace(QuantEmpenhoPC, ",", ".")
   
   Conexao.Execute "UPDATE Estoque_controle Set Estoque_real = " & NovoValor & ", Estoque_real_PC = " & NovoValor1 & ", Estoque_venda = " & NovoValor & ", peso_unit = '" & IIf(IsNull(TBProduto!peso_metro), "", TBProduto!peso_metro) & "', Pedido = '" & IIf(IsNull(TBProduto!Un_Kg), "", TBProduto!Un_Kg) & "' where IDestoque = " & TBEstoque!IDEstoque
   Conexao.Execute "UPDATE Estoque_controle Set Valor_Total = ROUND(valor_unitario * Estoque_real, 2) where IDestoque = " & TBEstoque!IDEstoque
   
   End If
  TBProduto.Close
 '===========================================================================
 'Cria a movimentação de saida do estoque
 '===========================================================================
    Set TBProduto = CreateObject("adodb.recordset")
    TBProduto.Open "Select * from Estoque_movimentacao", Conexao, adOpenKeyset, adLockOptimistic
    TBProduto.AddNew
    TBProduto!Destino = "Interno"
    TBProduto!Terceiros = False
    TBProduto!Documento = TBNivel10!int_NotaFiscal
    TBProduto!LOTE = TBEstoque!LOTE
    TBProduto!Desenho = TBEstoque!Desenho
    TBProduto!Data = Date
    TBProduto!Descricao = IIf(IsNull(TBEstoque!Descricao), "", TBEstoque!Descricao)
    TBProduto!Familia = IIf(IsNull(TBEstoque!Classe), "", TBEstoque!Classe)
    TBProduto!Requisitante = txtUsuario
    TBProduto!Responsavel = txtUsuario
    TBProduto!IDEstoque = TBEstoque!IDEstoque
    TBProduto!OE = TBNivel10!int_NotaFiscal
    
    TBProduto!ID_prod_NF = TBNivel10!Int_codigo
    If QtdeSaida >= qtdeliberada Then TBProduto!Operacao = "SAIDA_NOTA" Else TBProduto!Operacao = "SAIDA_NOTA_PARCIAL"
    
    If EmpenhoVerificar = True Then
        If qtdeliberada >= TBEstoque!qtdeEstoqueMin Then
            qtdeliberada = qtdeliberada - TBEstoque!qtdeEstoqueMin
            'qtdeliberar = qtdeliberar - TBEstoque!qtdeEstoqueMin
        Else
            qtdeliberada = 0
            'qtdeliberar = 0
        End If
    Else
        If qtdeliberada >= Qtd Then
            qtdeliberada = qtdeliberada - Qtd
            'qtdeliberar = qtdeliberar - Qtd
        Else
            qtdeliberada = 0
            'qtdeliberar = 0
        End If
    End If
    
    
    TBProduto!Saida = QtdeSaida
    TBProduto!Saida_PC = QtdeSaidaPC
    TBProduto!estoque_venda = QtdeSaida

    'Atualiza valor do material no estoque
    TBProduto!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "###,##0.0000000000"))
    TBProduto!vlrTotal = Format(QtdeSaida * TBProduto!VlrUnit, "###,##0.00")
    TBProduto.Update
    USMsgBox "Produto(s) da nota fiscal baixado(s) do estoque com sucesso!", vbInformation, "CAPRIND v5.0"
    '==================================
    Modulo = "Faturamento/Nota fiscal/Própria"
    Evento = "Retirar"
    ID_documento = TBNivel10!Int_codigo
    Documento = "Cód. interno: " & TBEstoque!Desenho & " - RE: " & TBEstoque!IDEstoque
    Documento1 = ""
    ProcGravaEvento
    '==================================
    Permitido1 = True
    
    'Centro de custo
    ProcCriaCreditoCCProdutoItemSelecionada TBNivel10!Codproduto, Date, frmFaturamento_Prod_Serv.txtIDEmpresa.Text, TBProduto!IDoperacao, TBProduto!vlrTotal

    TBProduto.Close
Proximo:
    TBEstoque.MoveNext
Loop

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Public Sub procExcluirMovimentacao_NFe()
On Error GoTo tratar_erro

Set TBNivel10 = CreateObject("adodb.recordset")
TBNivel10.Open "Select TBL.Int_codigo, TBL.int_Cod_Produto, TBL.int_Qtd, TBL.ID_nota, TBL.N_Referencia, TBL.Int_NotaFiscal, P.ID_PC, P.Codproduto from (tbl_detalhes_nota TBL INNER JOIN ProjProduto P ON TBL.Codproduto = P.Codproduto) INNER JOIN tbl_NaturezaOperacao CFOP ON CFOP.IDCountCfop = TBL.ID_CFOP where TBL.ID_Nota = " & TBVendas!ID & " and P.Estoque = 'True' and (CFOP.Remessa IS NULL or CFOP.Remessa = 'False')", Conexao, adOpenKeyset, adLockOptimistic
Do While TBNivel10.EOF = False
    quantidade = 0
    Qtde = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from estoque_movimentacao where ID_Prod_NF = " & TBNivel10!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Do While TBAbrir.EOF = False
            valor = IIf(IsNull(TBAbrir!VlrUnit), 0, TBAbrir!VlrUnit)
            'Atualiza qtde. expedida
            Qtd = IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
            
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select PP.Qtdeexpedida, PP.Dataexpedicao, NFPP.* from vendas_carteira PP INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = PP.Codigo and NFPP.Codinterno = PP.Desenho where NFPP.ID_prod_NF = " & TBAbrir!ID_prod_NF & " order by PP.PrazoFinal", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = False Then
                Do While TBGravar.EOF = False
                    If Qtd >= TBGravar!qtdeexpedida Then qt = TBGravar!qtdeexpedida Else qt = Qtd
                    TBGravar!qtdeexpedida = TBGravar!qtdeexpedida - qt
                    Qtd = Qtd - qt
                    
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select Data from Estoque_movimentacao where Idoperacao <> " & TBAbrir!IDoperacao & " and ID_prod_NF = " & IIf(IsNull(TBAbrir!ID_prod_NF), 0, TBAbrir!ID_prod_NF) & " and (Operacao = 'SAIDA_NOTA' or Operacao = 'SAIDA_NOTA_PARCIAL') order by Data desc", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = False Then
                        TBGravar!dataexpedicao = TBFI!Data
                    Else
                        TBGravar!dataexpedicao = Null
                    End If
                    TBFI.Close
                    TBGravar.Update
                    
                    'Desvincula pedido da ordem para estoque
                    If IsNumeric(TBAbrir!LOTE) = True Then
                        Set TBAliquota = CreateObject("adodb.recordset")
                        TBAliquota.Open "Select * from Producao_pedidos where IDcarteira = " & IIf(IsNull(TBGravar!ID_carteira), 0, TBGravar!ID_carteira) & " and Ordem = " & TBAbrir!LOTE & " and Expedicao = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAliquota.EOF = False Then
                            TBAliquota!Qtde_empenho = TBAliquota!Qtde_empenho - qt
                            TBAliquota!Qtde_entrada = TBAliquota!Qtde_empenho
                            TBAliquota.Update
                            
                            If TBAliquota!Qtde_empenho <= 0 Then Conexao.Execute "DELETE from Producao_pedidos where IDcarteira = " & IIf(IsNull(TBGravar!ID_carteira), 0, TBGravar!ID_carteira) & " and Ordem = " & TBAbrir!LOTE & " and Expedicao = 'True'"
                        End If
                        TBAliquota.Close
                    End If
                    
                    Do While qt > 0
                        'Atualiza qtde. de saída no empenho
                        Set TBAliquota = CreateObject("adodb.recordset")
                        TBAliquota.Open "Select EE.Qtde_saida from Estoque_Controle_Empenho_Vendas EE INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON EE.ID_carteira = NFPP.ID_carteira where NFPP.ID_prod_NF = " & TBAbrir!ID_prod_NF & " and EE.ID_estoque = " & TBAbrir!IDEstoque & " and EE.Qtde_saida > 0", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAliquota.EOF = False Then
                            If TBAliquota!Qtde_saida >= qt Then
                                TBAliquota!Qtde_saida = TBAliquota!Qtde_saida - qt
                                qt = 0
                            Else
                                qt = qt - TBAliquota!Qtde_saida
                                TBAliquota!Qtde_saida = 0
                            End If
                            TBAliquota.Update
                        Else
                            GoTo Prosseguir
                        End If
                        TBAliquota.Close
                    Loop
Prosseguir:
                    If Qtd <= 0 Then GoTo Prosseguir1
                    TBGravar.MoveNext
                Loop
            End If
    
Prosseguir1:
            IDEstoque = TBAbrir!IDEstoque
            quantidade = IIf(IsNull(TBAbrir!Saida), 0, TBAbrir!Saida)
            TBAbrir.Delete
            
            'Centro de custo
            Conexao.Execute "DELETE from CC_realizado where ID_estoque = " & IDEstoque
                    
            'Corrige estoque real
            Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select * from estoque_controle where idestoque = " & IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBEstoque.EOF = False Then
                '==================================
                Modulo = "Faturamento/Nota fiscal/Própria"
                Evento = "Excluir movimentação"
                ID_documento = IDEstoque
                Documento = "Cód. interno: " & TBEstoque!Desenho & " - Nº lote: " & TBEstoque!LOTE & " - Nº corrida: " & TBEstoque!Corrida & " - Nº certificado: " & TBEstoque!Certificado & " - Local armaz.: " & TBEstoque!local_armaz
                Documento1 = ""
                ProcGravaEvento
                '==================================
                Qtd = IIf(IsNull(TBEstoque!estoque_real), 0, TBEstoque!estoque_real)
                Qtde = Format(Qtd + quantidade, "###,##0.0000")
                TBEstoque!estoque_real = Format(Qtde, "###,##0.0000")
                TBEstoque!Qtde = Qtde
                TBEstoque!estoque_real_PC = 0
                                                
                'Atualiza valor do material no estoque
                TBEstoque!Valor_total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * Qtde, "###,##0.00")
                        
                TBEstoque.Update
                Set TBMaterial = CreateObject("adodb.recordset")
                TBMaterial.Open "Select * from Estoque_movimentacao where IDEstoque = " & TBEstoque!IDEstoque, Conexao, adOpenKeyset, adLockOptimistic
                If TBMaterial.EOF = True Then TBEstoque.Delete
                TBMaterial.Close
            End If
            TBEstoque.Close

            TBAbrir.MoveNext
        Loop
    End If
    TBAbrir.Close
    TBNivel10.MoveNext
Loop
TBNivel10.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
