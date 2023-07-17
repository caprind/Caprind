VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#19.3#0"; "Codejock.SkinFramework.v19.3.0.ocx"
Begin VB.MDIForm frmMDI 
   Appearance      =   0  'Flat
   BackColor       =   &H00815135&
   Caption         =   "Caprind v5.0. - Menu Principal"
   ClientHeight    =   10020
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   15240
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin DrawSuite2022.USMenu USMenu1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   635
   End
   Begin DrawSuite2022.USStatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   9615
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   714
      ShowLastSeperators=   -1  'True
   End
   Begin VB.Timer TimerProdutos 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5430
      Top             =   870
   End
   Begin VB.PictureBox picVBar 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   4200
      MousePointer    =   9  'Size W E
      Picture         =   "frmMDI.frx":0CCA
      ScaleHeight     =   9255
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   360
      Width           =   15
   End
   Begin VB.Timer Timer_logoff_automatico 
      Interval        =   60000
      Left            =   6360
      Top             =   390
   End
   Begin VB.Timer Timer_logon 
      Interval        =   10000
      Left            =   5430
      Top             =   390
   End
   Begin VB.Timer Timer_avisodiario 
      Interval        =   10000
      Left            =   5910
      Top             =   390
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7560
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6900
      Top             =   390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   10
      ImageHeight     =   9
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D290
            Key             =   "A"
            Object.Tag             =   "A"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D5FA
            Key             =   "Menu"
            Object.Tag             =   "Menu"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":D677
            Key             =   "B"
            Object.Tag             =   "B"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picMain 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      ScaleHeight     =   9255
      ScaleWidth      =   4200
      TabIndex        =   0
      Top             =   360
      Width           =   4200
      Begin VB.PictureBox picBot 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   11415
         Left            =   90
         ScaleHeight     =   11415
         ScaleWidth      =   4200
         TabIndex        =   2
         Top             =   270
         Width           =   4200
         Begin VB.OptionButton optTheme 
            Caption         =   "System"
            Height          =   255
            Index           =   60
            Left            =   5070
            TabIndex        =   5
            Top             =   8970
            Width           =   1335
         End
         Begin VB.OptionButton optTheme 
            Caption         =   "Windows 10"
            Height          =   255
            Index           =   6
            Left            =   5070
            TabIndex        =   4
            Top             =   8610
            Width           =   1335
         End
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   9285
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   16378
            _Version        =   393217
            Indentation     =   53
            LabelEdit       =   1
            Style           =   7
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ResizeFonts     =   0   'False
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10785
      FormWidthDT     =   15360
      FormScaleHeightDT=   10785
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4410
      Top             =   90
      _Version        =   1245187
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuSistema 
      Caption         =   "Configuração do sistema"
      Begin VB.Menu mnuOptGeral 
         Caption         =   "Opções gerais"
      End
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuários"
         Begin VB.Menu Mnu_usuarios 
            Caption         =   "Usuários"
         End
         Begin VB.Menu Mnu_usuarios_conectados 
            Caption         =   "Conectados"
         End
         Begin VB.Menu Mnu_usuarios_eventos 
            Caption         =   "Eventos realizados"
         End
      End
      Begin VB.Menu mnu_backup 
         Caption         =   "Criar backup"
         Begin VB.Menu mnu_backup_config 
            Caption         =   "Configurações"
         End
         Begin VB.Menu Mnu_bkp_apontamentos 
            Caption         =   "Apontamentos"
         End
         Begin VB.Menu Mnu_bkp_eventos 
            Caption         =   "Eventos realizados por usuário"
         End
      End
      Begin VB.Menu Mnu_reindexar_BD 
         Caption         =   "Reindexar BD"
         Begin VB.Menu Mnu_reindexar_BD_caprind 
            Caption         =   "Caprind e Gerprod"
         End
         Begin VB.Menu Mnu_reindexar_BD_GNFe 
            Caption         =   "GNFe"
         End
      End
   End
   Begin VB.Menu mnucomercial 
      Caption         =   "Administrativo"
      Begin VB.Menu mnurh 
         Caption         =   "RH"
         Begin VB.Menu mnufuncionarios 
            Caption         =   "Funcionários"
         End
         Begin VB.Menu Mnu_RH_relatorios 
            Caption         =   "Relatórios"
            Begin VB.Menu Mnu_RH_relatorios_desoneracao 
               Caption         =   "Desoneração da folha de pagamento"
            End
         End
      End
      Begin VB.Menu mnuCompras 
         Caption         =   "Compras"
         Begin VB.Menu mnucompras_familia 
            Caption         =   "Famílias"
         End
         Begin VB.Menu mnucompras_itens 
            Caption         =   "Produtos e serviços"
         End
         Begin VB.Menu mnufornecedor 
            Caption         =   "Fornecedores"
         End
         Begin VB.Menu mnuProgramacaodecompra 
            Caption         =   "Programação"
         End
         Begin VB.Menu mnucotacao 
            Caption         =   "Cotação"
         End
         Begin VB.Menu Mnu_PC 
            Caption         =   "Pedidos"
            Begin VB.Menu mnuPedidodecompra 
               Caption         =   "Gerenciar"
            End
            Begin VB.Menu Mnu_AprovarPedido 
               Caption         =   "Aprovar"
            End
         End
         Begin VB.Menu Mnu_necessidade_compras 
            Caption         =   "Necessidade"
         End
         Begin VB.Menu Mnu_compras_NC 
            Caption         =   "Não conformidade"
         End
         Begin VB.Menu Mnu_compras_atualizacao_valores 
            Caption         =   "Atualização de valores"
         End
         Begin VB.Menu mnucompras_relatorios 
            Caption         =   "Relatórios"
            Begin VB.Menu mnuCompras_Rel_Historico 
               Caption         =   "Histórico"
            End
            Begin VB.Menu mnuCompras_Rel_Indice_Atraso 
               Caption         =   "Índice de atraso"
            End
         End
      End
      Begin VB.Menu mnuVendas 
         Caption         =   "Vendas"
         Begin VB.Menu mnuvendas_familia 
            Caption         =   "Famílias"
         End
         Begin VB.Menu mnuvendas_itens 
            Caption         =   "Produtos e serviços"
         End
         Begin VB.Menu mnuclientes 
            Caption         =   "Clientes"
         End
         Begin VB.Menu mnuControledevendedores 
            Caption         =   "Vendedores"
         End
         Begin VB.Menu mnuTelemarketing 
            Caption         =   "Telemarketing"
         End
         Begin VB.Menu Mnu_vendas_simulacao 
            Caption         =   "Simulação"
         End
         Begin VB.Menu Mnu_vendas_empenho 
            Caption         =   "Empenho"
         End
         Begin VB.Menu Mnu_programacao_venda 
            Caption         =   "Programação"
         End
         Begin VB.Menu mnuPropostacomercial 
            Caption         =   "Proposta comercial"
         End
         Begin VB.Menu mnupedidointerno 
            Caption         =   "Pedido interno"
         End
         Begin VB.Menu mnuCarteiradepedidos 
            Caption         =   "Carteira de vendas"
         End
         Begin VB.Menu mnuSituacaodaproposta 
            Caption         =   "Situação da produção"
         End
         Begin VB.Menu mnuInformacoesdefaturamento 
            Caption         =   "Informações de faturamento"
         End
         Begin VB.Menu Mnu_vendas_atualizacao_valores 
            Caption         =   "Atualização de valores"
         End
         Begin VB.Menu mnuvendas_relatorios 
            Caption         =   "Relatórios"
            Begin VB.Menu mnuVendas_Rel_Historico 
               Caption         =   "Histórico"
            End
            Begin VB.Menu mnuDesempenho 
               Caption         =   "Desempenho"
            End
            Begin VB.Menu mnuVendas_Rel_indice_atrazo 
               Caption         =   "Índice de atraso"
            End
            Begin VB.Menu mnuVendas_Rel_comissao 
               Caption         =   "Comissão "
            End
            Begin VB.Menu mnumeta 
               Caption         =   "Comissões x meta"
            End
         End
      End
      Begin VB.Menu mnucontas 
         Caption         =   "Financeiro"
         Begin VB.Menu mnuFinanceiro_familias 
            Caption         =   "Plano de contas"
         End
         Begin VB.Menu mnuinstituicoes 
            Caption         =   "Instituições Financeiras"
         End
         Begin VB.Menu mnuapagar 
            Caption         =   "Contas a pagar"
         End
         Begin VB.Menu mnuucpagas 
            Caption         =   "Contas pagas"
         End
         Begin VB.Menu mnureceber 
            Caption         =   "Contas a receber"
         End
         Begin VB.Menu mnurecebidas 
            Caption         =   "Contas recebidas"
         End
         Begin VB.Menu Mnu_desconto_duplicata 
            Caption         =   "Desconto de duplicata"
         End
         Begin VB.Menu mnuFluxoMenu 
            Caption         =   "Fluxo de caixa"
            Begin VB.Menu mnuFluxoResumido 
               Caption         =   "Resumido Gráfico"
            End
            Begin VB.Menu mnuFluxo 
               Caption         =   "Detalhado"
            End
         End
         Begin VB.Menu Mnu_financeiro_relatorios 
            Caption         =   "Relatórios"
            Begin VB.Menu Mnu_financeiro_historico 
               Caption         =   "Histórico"
            End
            Begin VB.Menu Mnu_financeiro_relatorios_razao 
               Caption         =   "Razão"
               Begin VB.Menu mnuRazaoDetalhado 
                  Caption         =   "Detalhado"
               End
               Begin VB.Menu mnuRazaoResumido 
                  Caption         =   "Resumido"
               End
            End
         End
      End
      Begin VB.Menu mnuFaturamento 
         Caption         =   "Faturamento"
         Begin VB.Menu mnufiscal 
            Caption         =   "Fiscal"
            Begin VB.Menu mnucf 
               Caption         =   "Classificação fiscal"
            End
            Begin VB.Menu Mnunatureza 
               Caption         =   "Natureza de operação"
            End
            Begin VB.Menu mnuSub 
               Caption         =   "Regiões e Substituição tributária"
            End
         End
         Begin VB.Menu mnunotafiscal 
            Caption         =   "Nota fiscal"
            Begin VB.Menu mnuNotaEntrada 
               Caption         =   "Terceiros"
            End
            Begin VB.Menu mnuNotaSaida 
               Caption         =   "Própria"
            End
            Begin VB.Menu Mnu_SPED 
               Caption         =   "SPED"
            End
            Begin VB.Menu mnuCarteiraWEB 
               Caption         =   "Carteira WEB"
            End
            Begin VB.Menu Mnu_NF_exportar 
               Caption         =   "Exportar (.txt)"
            End
         End
         Begin VB.Menu mnucartacorrecao 
            Caption         =   "Carta de correção"
         End
         Begin VB.Menu mnuminuta 
            Caption         =   "Minuta de despacho"
         End
         Begin VB.Menu mnurelatoriofaturamento 
            Caption         =   "Relatórios"
            Begin VB.Menu mnuFaturamento_Rel_Historico 
               Caption         =   "Histórico"
            End
            Begin VB.Menu mnuFaturamento_Rel_Relacionamento 
               Caption         =   "Relacionamento de notas fiscais"
            End
            Begin VB.Menu mnuFaturamento_Rel_Impostos 
               Caption         =   "Impostos"
            End
            Begin VB.Menu mnuFaturamento_Rel_dozemeses 
               Caption         =   "Doze últimos meses"
            End
            Begin VB.Menu mnuMensal 
               Caption         =   "Mensal x CFOP"
            End
            Begin VB.Menu mnuMensalCST 
               Caption         =   "Mensal x CST"
            End
            Begin VB.Menu mnuNotas 
               Caption         =   "Notas fiscais no período"
            End
         End
      End
      Begin VB.Menu MnuCustos 
         Caption         =   "Custos"
         Begin VB.Menu Mnu_centro_de_custo 
            Caption         =   "Centro de custo"
         End
         Begin VB.Menu Mnu_relatorios_custos 
            Caption         =   "Relatórios"
            Begin VB.Menu MnuCustos_detalhado 
               Caption         =   "Detalhado"
            End
            Begin VB.Menu MnuCustos_resumido 
               Caption         =   "Resumido"
            End
            Begin VB.Menu Mnu_relatorios_custo_prev_real 
               Caption         =   "Previsto x Realizado"
            End
         End
      End
   End
   Begin VB.Menu mnuprojeto 
      Caption         =   "Engenharia"
      Begin VB.Menu mnuFamilia 
         Caption         =   "Famílias"
      End
      Begin VB.Menu mnuprodutos 
         Caption         =   "Produtos e serviços"
      End
      Begin VB.Menu mnuconjunto 
         Caption         =   "Conjuntos"
      End
      Begin VB.Menu mnuestrutura 
         Caption         =   "Estrutura"
         Begin VB.Menu mnuestComp 
            Caption         =   "Completa"
         End
         Begin VB.Menu mnuEstResum 
            Caption         =   "Resumida"
         End
      End
      Begin VB.Menu mnucontprojetos 
         Caption         =   "Controle de projetos"
      End
      Begin VB.Menu mnuProcessosPrinc 
         Caption         =   "Processos"
      End
      Begin VB.Menu mnuNormas 
         Caption         =   "Normas"
      End
   End
   Begin VB.Menu mnuControle 
      Caption         =   "PCP"
      Begin VB.Menu mnuCadastroFerramentas 
         Caption         =   "Postos de trabalho"
      End
      Begin VB.Menu mnucodigos 
         Caption         =   "Códigos de trabalho"
      End
      Begin VB.Menu mnuProdCargaMaq 
         Caption         =   "Carga de posto de trabalho"
      End
      Begin VB.Menu mnuordem 
         Caption         =   "Gerenciamento de ordem"
      End
      Begin VB.Menu mnuMonitorTrabalho 
         Caption         =   "Monitor de trabalho"
      End
      Begin VB.Menu mnuPecasemproducao 
         Caption         =   "Situação da produção"
      End
      Begin VB.Menu mnuprogramas 
         Caption         =   "Programas CNC"
      End
      Begin VB.Menu Mnu_necessidade_PCP 
         Caption         =   "Necessidade"
      End
      Begin VB.Menu Mnu_PCP_NC 
         Caption         =   "Não conformidade"
      End
      Begin VB.Menu Mnu_PCP_programacao 
         Caption         =   "Programação da produção"
      End
      Begin VB.Menu Mnu_plano_apontamento 
         Caption         =   "Plano da produção"
      End
      Begin VB.Menu Mnu_requisicao_ordem 
         Caption         =   "Requisição da ordem"
      End
      Begin VB.Menu mnuPCP_relatorios 
         Caption         =   "Relatórios"
         Begin VB.Menu mnuProdutividade 
            Caption         =   "Produtividade"
         End
         Begin VB.Menu mnunaoconformidade 
            Caption         =   "Não conformidade"
         End
         Begin VB.Menu mnugrafeven 
            Caption         =   "Monitor de eventos"
         End
         Begin VB.Menu mnuindiceatrazo 
            Caption         =   "Índice de atraso"
         End
         Begin VB.Menu mnuResultadosOrdem 
            Caption         =   "Resultados da ordem"
         End
         Begin VB.Menu mnuProducaoDia 
            Caption         =   "Produção Dia x MP"
         End
      End
   End
   Begin VB.Menu mnuinspecao 
      Caption         =   "Qualidade"
      Begin VB.Menu mnuqualidade_familia 
         Caption         =   "Famílias"
      End
      Begin VB.Menu mnuinstrumentos 
         Caption         =   "Instrumentos"
      End
      Begin VB.Menu Mnu_qualidade_almoxarifado 
         Caption         =   "Almoxarifado"
      End
      Begin VB.Menu mnuplanoinspecao 
         Caption         =   "Plano de inspeção"
      End
      Begin VB.Menu mnuplanomedicao 
         Caption         =   "Controle de medição"
      End
      Begin VB.Menu mnurecebimento 
         Caption         =   "Inspeção de recebimento"
      End
      Begin VB.Menu mnuEnsaios 
         Caption         =   "Ensaios"
         Begin VB.Menu mnuUltraSom 
            Caption         =   "Ultra-som (Em manutenção)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuLiquido 
            Caption         =   "Líquido penetrante (Em manutenção)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuEnsaio_ControleCertificados 
            Caption         =   "Controle de certificados (Em manutenção)"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuControleCertificados 
         Caption         =   "Controle de certificados"
      End
      Begin VB.Menu Mnu_CDD 
         Caption         =   "Controle de documentos e dados"
      End
      Begin VB.Menu mnucqnc 
         Caption         =   "Não conformidade"
         Begin VB.Menu mnucqnc_descricao 
            Caption         =   "Descrição da não conformidade"
         End
         Begin VB.Menu mnucqnc_NC 
            Caption         =   "Gerenciar"
         End
      End
      Begin VB.Menu MnuSA 
         Caption         =   "Solicitação de ação"
      End
      Begin VB.Menu MnuSD 
         Caption         =   "Solicitação de desvio"
      End
      Begin VB.Menu MnuRNC 
         Caption         =   "RNC"
      End
      Begin VB.Menu MnuPPAP 
         Caption         =   "PPAP"
         Begin VB.Menu MnuPSW 
            Caption         =   "PSW (Em manutenção)"
            Enabled         =   0   'False
         End
         Begin VB.Menu MnuFMEA 
            Caption         =   "FMEA (Em manutenção)"
            Enabled         =   0   'False
         End
         Begin VB.Menu MnuPlano_de_Controle 
            Caption         =   "Plano de controle (Em manutenção)"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu MnuRevisao_relatorios 
         Caption         =   "Histórico de revisão dos relatórios"
      End
      Begin VB.Menu Mnu_qualidade_relatorios 
         Caption         =   "Relatórios"
         Begin VB.Menu Mnu_qualidade_relatorios_NC 
            Caption         =   "Não conformidade"
         End
         Begin VB.Menu Mnu_qualidade_relatorios_historico 
            Caption         =   "Histórico"
         End
      End
   End
   Begin VB.Menu mnuestoque 
      Caption         =   "Estoque"
      Begin VB.Menu mnualmoxarifado 
         Caption         =   "Almoxarifado"
      End
      Begin VB.Menu mnulocalarmaz 
         Caption         =   "Local de armazenamento"
      End
      Begin VB.Menu Mnu_requisicao_materiais 
         Caption         =   "Requisição de materiais"
      End
      Begin VB.Menu mnumateriais 
         Caption         =   "Recebimento"
         Begin VB.Menu mnureceberpedidodecompra 
            Caption         =   "Pedido de compra"
         End
         Begin VB.Menu mnurecebernfconsignada 
            Caption         =   "Consignação"
         End
         Begin VB.Menu Mnu_importacao_XML 
            Caption         =   "Importar nota de terceiros"
         End
      End
      Begin VB.Menu Mnu_inventario 
         Caption         =   "Inventário"
      End
      Begin VB.Menu mnuferramenta 
         Caption         =   "Movimentação"
         Begin VB.Menu mnuEMResumido 
            Caption         =   "Resumido"
         End
         Begin VB.Menu mnuEMDetalhado 
            Caption         =   "Detalhado"
         End
      End
      Begin VB.Menu Mnu_necessidade_estoque 
         Caption         =   "Necessidade"
      End
      Begin VB.Menu Mnu_ordem_fat 
         Caption         =   "Ordem de faturamento"
      End
      Begin VB.Menu Mnu_estoque_nota 
         Caption         =   "Nota fiscal"
      End
   End
   Begin VB.Menu mnumanutenc 
      Caption         =   "Manutenção"
      Begin VB.Menu mnu_Planomanutencao 
         Caption         =   "Plano de manutenção"
      End
      Begin VB.Menu mnumanutencao 
         Caption         =   "Gerenciamento"
      End
      Begin VB.Menu mnurelatorio_manut 
         Caption         =   "Relatórios"
         Begin VB.Menu mnuhistorico_manut 
            Caption         =   "Histórico"
         End
      End
   End
   Begin VB.Menu Mnu_outros 
      Caption         =   "Outros"
      Begin VB.Menu mnuSolicitaçãodecompra 
         Caption         =   "Solicitação de compra"
      End
      Begin VB.Menu mnuSolicitacaoDeProducao 
         Caption         =   "Solicitação de produção"
      End
      Begin VB.Menu mnuListadeitens 
         Caption         =   "Follow up de compras"
      End
      Begin VB.Menu Mnu_validacao 
         Caption         =   "Validação de procedimentos"
      End
      Begin VB.Menu Mnu_analise_critica 
         Caption         =   "Análise crítica"
      End
      Begin VB.Menu mnu_orcamento 
         Caption         =   "Orçamento"
      End
      Begin VB.Menu Mnu_calculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu Mnu_downloads 
         Caption         =   "Downloads"
         Begin VB.Menu Mnu_downloads_NF 
            Caption         =   "Nota fiscal"
         End
         Begin VB.Menu Mnu_downloads_boleto 
            Caption         =   "Boleto"
         End
      End
      Begin VB.Menu mnuTema 
         Caption         =   "Temas"
         Begin VB.Menu mnu_nenhum 
            Caption         =   "Padrão"
         End
         Begin VB.Menu mnu_Luna 
            Caption         =   "Luna"
         End
         Begin VB.Menu mnu_2007 
            Caption         =   "Office 2007"
         End
         Begin VB.Menu mnu_Vista 
            Caption         =   "Vista"
         End
         Begin VB.Menu mnu_royale 
            Caption         =   "Royale"
         End
         Begin VB.Menu mnu_2010 
            Caption         =   "Office 2010"
         End
         Begin VB.Menu mnu_Windos8 
            Caption         =   "Windows 8"
         End
      End
   End
   Begin VB.Menu Mnu_suporte 
      Caption         =   "Suporte"
      Begin VB.Menu mnuSuportetecnico 
         Caption         =   "Chamado"
      End
      Begin VB.Menu Mnu_chat 
         Caption         =   "Chat (online)"
      End
      Begin VB.Menu Mnu_SAS 
         Caption         =   "Solicitação de atendimento"
      End
      Begin VB.Menu mnu_TeamViewer 
         Caption         =   "Download Team Viewer"
      End
      Begin VB.Menu Mnu_atualizacao 
         Caption         =   "Atualização"
         Begin VB.Menu Mnu_atualizacao_caprind 
            Caption         =   "Caprind e Gerprod"
         End
         Begin VB.Menu mnu_historico 
            Caption         =   "Histórico de atualizações"
         End
      End
   End
   Begin VB.Menu mnufinalizar 
      Caption         =   "Finalizar"
      Begin VB.Menu mnulogoff 
         Caption         =   "Efetuar logoff"
      End
      Begin VB.Menu mnusair 
         Caption         =   "Sair"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bDrag As Boolean    'True if the user has the mouse pressed while on the resize bars

Private Sub MDIForm_Load()
On Error GoTo tratar_erro

mnuSistema.Visible = False
mnucomercial.Visible = False
mnuprojeto.Visible = False
mnuControle.Visible = False
mnuinspecao.Visible = False
mnuestoque.Visible = False
mnumanutenc.Visible = False
Mnu_outros.Visible = False
Mnu_suporte.Visible = False
mnufinalizar.Visible = False


USMenu1.AddMenuObj "Configuração do sistema", mnuSistema, , vbWhite
USMenu1.AddMenuObj "Administrativo", mnucomercial, , vbWhite
USMenu1.AddMenuObj "Engenharia", mnuprojeto, , vbWhite
USMenu1.AddMenuObj "PCP", mnuControle, , vbWhite
USMenu1.AddMenuObj "Qualidade", mnuinspecao, , vbWhite
USMenu1.AddMenuObj "Estoque", mnuestoque, , vbWhite
USMenu1.AddMenuObj "Manutenção", mnumanutenc, , vbWhite
USMenu1.AddMenuObj "Outros", Mnu_outros, , vbWhite
USMenu1.AddMenuObj "Suporte", Mnu_suporte, , vbWhite
USMenu1.AddMenuObj "Finalizar", mnufinalizar, , vbWhite

'==================================================
' Desabilitar "X" do MDI (Fechar)
'==================================================
Dim hMenu As Long
hMenu = GetSystemMenu(hWnd, False)
DeleteMenu hMenu, 6, MF_BYPOSITION

'==================================================
'Muda a cor do fundo
'==================================================
Me.BackColor = RGB(53, 81, 129)
'==================================================
'Habilita Skyn para o formulário MDI
'==================================================
caminho = App.Path & "\Styles"

SkinFrameworkGlobalSettings.UseLegacyCore = False
TemaCaprind = ""
TemaINI = ""

'TemaCaprind = "\WinXP.Royale.cjstyles"
'TemaINI = "NormalRoyale.ini"

SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
SkinFramework1.ApplyWindow Me.hWnd

'=======================================================================
'frmMDI.SkinFramework1.ApplySkin Me.hWnd
'=======================================================================
If xPixels = 1024 And YPixels = 768 Then
    picMain.Width = 0
    WindowState = 2
    frmMenucaprind_menulateral.Show
Else
    Height = 11520
    Width = 19200 '15480
    frmMenucaprind_menulateral.Show
End If


Caption = VarE & " - CAPRIND - v" & App.Major & "." & App.Minor & "." & App.Revision & " - Menu Principal"
ProcMontaStatusBar
      
If Time$ > "19:00:00" Then
    USMsgBox "Boa noite " & pubUsuario & ", bem vindo ao Caprind.", vbInformation, "CAPRIND v5.0"
  ElseIf Time$ > "12:00:00" Then
    USMsgBox "Boa tarde " & pubUsuario & ", bem vindo ao Caprind.", vbInformation, "CAPRIND v5.0"
  ElseIf Time$ > "00:00:00" Then
    USMsgBox "Bom dia " & pubUsuario & ", bem vindo ao Caprind.", vbInformation, "CAPRIND v5.0"
End If

Timer_logon.Enabled = True

Logoff = False
ProcVerificaAvisoDiario
ProcVerificaLogoffAutomatico

ProcCarregaMenu TreeView1, ImageList1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MDIForm_Resize()
On Error GoTo tratar_erro

ProcVerificaAvisoDiario
ProcVerificaLogoffAutomatico
'frmMenucaprind_menulateral.Visible = True
'SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
'SkinFramework1.ApplyWindow Me.hWnd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo tratar_erro

ProcLogonOut
'If Time$ > "19:00:00" Then
    'usMsgbox "Boa noite " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind... !!!", _
            'vbInformation, "CAPRIND v5.0"
    'ElseIf Time$ > "12:00:00" And Time$ < "18:59:59" Then
        'usMsgbox "Boa Tarde " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind... !!!", _
                'vbInformation, "CAPRIND v5.0"
    'ElseIf Time$ > "00:00:00" And Time$ < "11:59:59" Then
        'usMsgbox "Bom dia " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind... !!!", _
            'vbInformation, "CAPRIND v5.0"
'End If
FunFechaBD
If Logoff = False Then End

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub mnu_2007_Click()
On Error GoTo tratar_erro

TemaCaprind = "\Office2007.cjstyles"
TemaINI = "NormalBlue.ini"

SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
SkinFramework1.ApplyWindow Me.hWnd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_2010_Click()
On Error GoTo tratar_erro

TemaCaprind = "\Office2010.cjstyles"
TemaINI = "NormalBlue.ini"

SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
SkinFramework1.ApplyWindow Me.hWnd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_analise_critica_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Outros/Análise crítica"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_analise.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_AprovarPedido_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Pedido/Aprovar"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_Aprovar_Pedido.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_atualizacao_caprind_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Suporte/Atualização/Caprind e Gerprod"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
If TemInternet = True And ErroDriverMYSQL = False Then
    If FunVerificaManutencaoAtiva = False Then Exit Sub
    If FunVerificaVersaoAtualizacaoCaprind = False Then Exit Sub
    If GerArqPastas.FileExists(Left(Localrel, 3) & "Caprind.zip") = True Then GerArqPastas.DeleteFile (Left(Localrel, 3) & "Caprind.zip")
    Atualizacao_GNFe = False
    Atualizacao_GMRE = False
    Atualizacao_versao = False
    Atualizacao_TeamViewer = False
    Frm_atualizacao_sistema.Show 1
Else
    If TemInternet = False Then MsgTexto = "não foi encontrado conexão com a internet" Else MsgTexto = "no momento estamos sem conexão com nosso servidor, favor tentar mais tarde"
    USMsgBox ("Não é permitido baixar a atualização, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub mnu_backup_config_Click()
On Error GoTo tratar_erro

Formulario = "Configuração do sistema/Criar backup/Configurações"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Frm_configuracao_backup.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_bkp_apontamentos_Click()
On Error GoTo tratar_erro

Formulario = "Configuração do sistema/Criar backup/Apontamentos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmBackup_apontamentos.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_bkp_eventos_Click()
On Error GoTo tratar_erro

Formulario = "Configuração do sistema/Criar backup/Eventos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
ProcCriarBackupEventos

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriarBackupEventos()
On Error GoTo tratar_erro
Dim TabelaTexto As String

If USMsgBox("Deseja gerar um backup dos eventos realizados por usuário?", vbYesNo, "Caprind") = vbYes Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Mascara", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        TabelaTexto = "Mascara_Backup_" & Format(Date, "dd_mm_yy")
        'Muda o nome da tabela
        Conexao.Execute "sp_rename 'dbo.Mascara', '" & TabelaTexto & "'"
        'Criar nova tabela
        Conexao.Execute "CREATE TABLE dbo.Mascara (idevento Int PRIMARY KEY IDENTITY,Modulo Nvarchar(255) null,Usuario Nvarchar(50) null, Operacao Nvarchar(100) null, Data Datetime null, Hora Datetime null, Documento Nvarchar(Max) null, Documento1 Nvarchar(Max) null, Id_documento int null)"
        Conexao.Execute "CREATE UNIQUE INDEX Mascara$idevento ON Mascara (idevento)"
        Conexao.Execute "CREATE INDEX Mascara$Modulo ON Mascara (Modulo)"
        Conexao.Execute "CREATE INDEX Mascara$Usuario ON Mascara (Usuario)"
        
        USMsgBox ("Backup efetuado com sucesso."), vbInformation, "CAPRIND v5.0"
        '==================================
        Modulo = Formulario
        Evento = "Criar backup eventos"
        ID_documento = 0
        Documento = "Nome do backup: " & TabelaTexto
        Documento1 = ""
        ProcGravaEvento
        '==================================
    Else
        USMsgBox ("Não foi encontrado nenhum registro para esta pesquisa."), vbExclamation, "CAPRIND v5.0"
    End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_calculadora_Click()
On Error GoTo tratar_erro

NomeCampo = "calc.exe na pasta C:\WINDOWS."
Shell "Calc.exe", vbNormalFocus
1:

Exit Sub
tratar_erro:
    If Err.Number = "53" Then
        USMsgBox ("Não será possível atualizar o sistema, pois não foi encontrado o arquivo " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Mnu_CDD_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Controle de documentos e dados"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCQ_sistema.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_centro_de_custo_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Custos/Centro de custo"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Frm_centro_de_custo.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_chat_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Suporte/Chat (online)"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
If TemInternet = True And ErroDriverMYSQL = False Then
    If FunVerifHorarioChat = False Then Exit Sub
    If FunVerificaManutencaoAtiva = False Then Exit Sub
    'Chat = True
    'Video_ajuda = False
'    With Frm_web
'        .WindowState = WindowState
'        .Web.Visible = False
'        .Web.Navigate "http://www.caprind.com.br"
'        .Show
'    End With
    
    'Set ie = New InternetExplorer
    'ie.Navigate "http://www.caprind.com.br/Suporte/chat.php"
    'ie.Visible = True
    
    Dim iret As Long
    iret = ShellExecute(Me.hWnd, vbNullString, "http://www.caprind.com.br/Suporte/chat.php", vbNullString, "c:\", SW_SHOWNORMAL)

    'Timer_chat.Enabled = True
Else
    If TemInternet = False Then MsgTexto = "não foi encontrado conexão com a internet" Else MsgTexto = "no momento estamos sem conexão com nosso servidor, favor tentar mais tarde"
    USMsgBox ("Não é permitido abrir este módulo, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_compras_atualizacao_valores_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Atualização de valores"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Vendas_Atualização_Valores = False
Frm_atualizacao_valores.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_compras_NC_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Não conformidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_NaoConformidade.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_desconto_duplicata_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Desconto de duplicata"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frm_trocaduplicata.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_downloads_boleto_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Outros/Downloads/Boleto"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
If TemInternet = True Then
    Downloads_NF = False
    frmMDI_downloads.Show 1
Else
    USMsgBox ("Não é permitido abrir este módulo, pois não foi encontrado conexão com a internet."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_downloads_NF_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Outros/Downloads/Nota fiscal"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
If TemInternet = True Then
    Downloads_NF = True
    frmMDI_downloads.Show 1
Else
    USMsgBox ("Não é permitido abrir este módulo, pois não foi encontrado conexão com a internet."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_Fluxo_Click()
On Error GoTo tratar_erro

frmFluxo_Caixa_Grafico.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_historico_Click()
On Error GoTo tratar_erro

frm_Atualizacoes.Show
'frmEstoque_consignado.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Mnu_importacao_XML_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Nota fiscal/Terceiros"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_nota = "Faturamento/Nota fiscal/Própria" Or Formulario_nota = "Estoque/Ordem de faturamento" Or Formulario_nota = "Estoque/Nota fiscal" Then
    If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
        If USMsgBox("O módulo " & Formulario_nota & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
    End If
End If

Formulario = "Faturamento/Nota fiscal/Terceiros"
Faturamento_NF_Saida = False
'======================================================
Faturamento_NF_Terceiro = True
Faturamento_NF_Propria = False
TPNota = "T"
'======================================================
frmFaturamento_Prod_Serv.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_Luna_Click()
On Error GoTo tratar_erro

TemaCaprind = "\WinXP.Luna.cjstyles"
TemaINI = "NormalBlue.ini"

Me.SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
Me.SkinFramework1.ApplyWindow Me.hWnd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_manifesto_Click()
On Error GoTo tratar_erro

'frmDDFeAPI.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_Mat_Terceiro_Click()
On Error GoTo tratar_erro

frmEstoque_consignado.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_nenhum_Click()
On Error GoTo tratar_erro

TemaCaprind = ""
TemaINI = ""

SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
SkinFramework1.ApplyWindow Me.hWnd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_orcamento_Click()
On Error GoTo tratar_erro

frm_orcamento.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_ordem_fat_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Ordem de faturamento"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_nota = "Faturamento/Nota fiscal/Terceiros" Or Formulario_nota = "Faturamento/Nota fiscal/Própria" Or Formulario_nota = "Estoque/Nota fiscal" Then
    If FunVerifFormAberto(frmEstoque_Ordem_Faturamento) = True Then
        If USMsgBox("O módulo " & Formulario_nota & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
    End If
End If

Formulario = "Estoque/Ordem de faturamento"
Faturamento_NF_Saida = True
frmEstoque_Ordem_Faturamento.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_inventario_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Inventário"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmestoque_fisico.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_necessidade_compras_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Necessidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_necessidade = "Compras/Necessidade" Or Formulario_necessidade = "PCP/Necessidade" Or Formulario_necessidade = "Estoque/Necessidade" Then
    If FunVerifFormAberto(Frm_necessidade) = True Then
        If USMsgBox("O módulo " & Formulario_necessidade & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload Frm_necessidade
    End If
End If

Compras_Necessidade = True
PCP_Necessidade = False
Frm_necessidade.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_estoque_nota_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Nota fiscal"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_nota = "Faturamento/Nota fiscal/Terceiros" Or Formulario_nota = "Faturamento/Nota fiscal/Própria" Or Formulario_nota = "Estoque/Ordem de faturamento" Then
    If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
        If USMsgBox("O módulo " & Formulario_nota & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
    End If
End If

Formulario = "Estoque/Nota fiscal"
Faturamento_NF_Saida = False
frmFaturamento_Prod_Serv.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_financeiro_historico_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Relatórios/Histórico"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFinanceiro_Relatorios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_financeiro_relatorios_razao_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
'
'Formulario = "Financeiro/Relatórios/Razão"
'ProcLiberaAcessos True
'If Acessos = False Then Exit Sub
'frmFinanceiro_Relatorios_Razao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_necessidade_estoque_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Necessidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_necessidade = "Compras/Necessidade" Or Formulario_necessidade = "PCP/Necessidade" Or Formulario_necessidade = "Estoque/Necessidade" Then
    If FunVerifFormAberto(Frm_necessidade) = True Then
        If USMsgBox("O módulo " & Formulario_necessidade & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload Frm_necessidade
    End If
End If

Compras_Necessidade = False
PCP_Necessidade = False
Frm_necessidade.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_NF_exportar_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Nota fiscal/Exportar"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFaturamento_Prod_serv_exportar.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_necessidade_PCP_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Necessidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_necessidade = "Compras/Necessidade" Or Formulario_necessidade = "PCP/Necessidade" Or Formulario_necessidade = "Estoque/Necessidade" Then
    If FunVerifFormAberto(Frm_necessidade) = True Then
        If USMsgBox("O módulo " & Formulario_necessidade & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload Frm_necessidade
    End If
End If

Compras_Necessidade = False
PCP_Necessidade = True
Frm_necessidade.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_PCP_NC_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Não conformidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
PCP_Ordem = True
frmcqnc.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_PCP_programacao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Programação da produção"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmProd_programacao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_plano_apontamento_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Plano da produção"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmPlano_producao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_Planomanutencao_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
'
'Formulario = "Vendas/Programação"
'ProcLiberaAcessos True
'If Acessos = False Then Exit Sub
frmMan_plano.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_programacao_venda_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Programação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_programacao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_qualidade_almoxarifado_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Almoxarifado"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Qualidade_Almox = True
frmCFI.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_qualidade_relatorios_historico_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Relatórios/Histórico"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmQualidade_Relatorios_historico.Show (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_qualidade_relatorios_NC_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Relatórios/Não conformidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Qualidade_NC = True
frmQualidade_Relatorios_NC.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_reindexar_BD_caprind_Click()
On Error GoTo tratar_erro

Formulario = "Configuração do sistema/Reindexar BD/Caprind e Gerprod"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
If USMsgBox("Deseja realmente reindexar o BD do Caprind e Gerprod?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerificaUsuariosConectados(pubUsuario) = False Then
        USMsgBox ("Não é permitido reindexar o BD, pois outros usuários estão utilizando o sistema."), vbExclamation, "CAPRIND v5.0"
    Else
        ProcReindexarBDCaprindeGerprod
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_reindexar_BD_GNFe_Click()
On Error GoTo tratar_erro

Formulario = "Configuração do sistema/Reindexar BD/GNFe"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
If USMsgBox("Deseja realmente reindexar o BD do GNFe?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    If FunVerificaUsuariosConectados(pubUsuario) = False Then
        USMsgBox ("Não é permitido reindexar o BD, pois outros usuários estão utilizando o sistema."), vbExclamation, "CAPRIND v5.0"
    Else
        ProcReindexarBDGNFe
    End If
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_Relatorios_Click()
On Error GoTo tratar_erro

frmestoque_item_imprimir.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_relatorios_custo_prev_real_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Custos/Relatórios/Previsto x Realizado"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmRelatorios_Custos_Prev_Real.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_requisicao_materiais_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Requisição de materiais"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmRequisicao_materiais.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_requisicao_ordem_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Requisição da ordem"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmprod_alterarRM.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_RH_relatorios_desoneracao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "RH/Relatórios/Desoneração da folha de pagamento"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
FrmRH_Relatorios_Desoneracao.Show (1)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_royale_Click()
On Error GoTo tratar_erro
            
    SkinFramework1.LoadSkin caminho & "\WinXP.Royale.cjstyles", "NormalRoyale.ini"
    SkinFramework1.ApplyWindow Me.hWnd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_SAS_Click()
On Error GoTo tratar_erro
Dim Comando As String
Dim appName As String


'================================================================
'Fecha Team Viewer se estiver aberto
'================================================================
appName = "TeamViewer.exe"
Comando = "TASKKILL -F -IM " & appName
Shell Comando
'================================================================

caminho = App.Path & "\TeamViewerQS.exe"

If USMsgBox("Deseja realmente solicitar uma conexão remota em sua máquina?", vbYesNo, "CAPRIND v5.0") = vbNo Then
    Exit Sub
End If

'================================================================
'Verifica se tem internet disponível
'================================================================
If IsInternetOnline = True Then
    If FileOrDirExists(caminho) = False Then
        Atualizacao_TeamViewerQS = True
        Frm_atualizacao_sistema.Show 1
    Else
        ProcAbrirArquivo (caminho)
    End If
Else
    If IsInternetOnline = False Then
        MsgTexto = "não foi encontrado conexão com a internet"
    Else
        MsgTexto = "no momento estamos sem conexão com nosso servidor, favor tentar mais tarde"
    End If
        
    USMsgBox ("Não é permitido abrir este módulo, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub Mnu_SPED_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Nota fiscal/SPED"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFaturamento_Prod_serv_SPED.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_TeamViewer_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
If DS.FileOrDirExists(App.Path & "\TeamViewer_Setup.exe") = True Then
    DS.FileExecute (App.Path & "\TeamViewer_Setup.exe")
    Exit Sub
End If
If TemInternet = True Then
    If FunVerificaManutencaoAtiva = False Then Exit Sub
    Atualizacao_versao = False
    Atualizacao_TeamViewer = True
    Frm_atualizacao_sistema.Show 1
Else
    USMsgBox ("Não é permitido baixar o Team Viewer, pois não foi encontrado conexão com a internet."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Mnu_usuarios_Click()
On Error GoTo tratar_erro
 
Formulario = "Configuração do sistema/Usuários"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmUsuarios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_usuarios_conectados_Click()
On Error GoTo tratar_erro
 
Formulario = "Configuração do sistema/Usuários/Conectados"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmOpcoes_Lista_usuarios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_usuarios_eventos_Click()
On Error GoTo tratar_erro
 
Formulario = "Configuração do sistema/Usuários/Eventos realizados"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmOpcoes_Lista_eventos.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_validacao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Outros/Validação de procedimentos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmValidacao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_vendas_atualizacao_valores_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Atualização de valores"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Vendas_Atualização_Valores = True
Frm_atualizacao_valores.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_vendas_empenho_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Empenho"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_Empenho.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnu_vendas_simulacao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

If FunVerifAtualizacaoObrigatoria(True, False) = False Then
    Formulario = "Vendas/Simulação"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    frmvendas_simulacao.Show
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_Vista_Click()
On Error GoTo tratar_erro

TemaCaprind = "\Vista.cjstyles"
TemaINI = "NormalBlack2.ini"

SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
SkinFramework1.ApplyWindow Me.hWnd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnu_Windos8_Click()
On Error GoTo tratar_erro

TemaCaprind = "\Windows8.cjstyles"
TemaINI = ""

SkinFramework1.LoadSkin caminho & TemaCaprind, TemaINI
SkinFramework1.ApplyWindow Me.hWnd

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnualmoxarifado_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Almoxarifado"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Qualidade_Almox = False
frmCFI.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuapagar_Click()
On Error GoTo tratar_erro
    
If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
    
Formulario = "Financeiro/Contas a pagar"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmContas_Pagar.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuCadastroFerramentas_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Postos de trabalho"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmGermaqfer.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucartacorrecao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
Formulario = "Faturamento/Carta de correção"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

frmFaturamento_CartaCorrecao_NS.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuCarteiraWEB_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub
'
'Formulario = "Faturamento/Fiscal/Classificação fiscal"
'ProcLiberaAcessos True
'If Acessos = False Then Exit Sub
frmVendas_Pedidos_WEB.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucf_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Fiscal/Classificação fiscal"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frm_Classificacao_Fiscal.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuclientes_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Clientes"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_cliente.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucodigos_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Códigos de trabalho"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCodigoDesc.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucompras_familia_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Famílias"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_familia = "Compras/Famílias" Or Formulario_familia = "Vendas/Famílias" Or Formulario_familia = "Engenharia/Famílias" Or Formulario_familia = "Qualidade/Famílias" Then
    If FunVerifFormAberto(frmproj_familia) = True Then
        If USMsgBox("O módulo " & Formulario_familia & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_familia
    End If
End If

Compras_Familia = True
Vendas_Familia = False
Qualidade_Familia = False
frmproj_familia.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucompras_itens_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Produtos e serviços"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_produto = "Compras/Produtos e serviços" Or Formulario_produto = "Vendas/Produtos e serviços" Or Formulario_produto = "Engenharia/Produtos e serviços" Then
    If FunVerifFormAberto(frmproj_produto) = True Then
        If USMsgBox("O módulo " & Formulario_produto & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_produto
    End If
End If

Engenharia_Produtos = False
Compras_Produtos = True
Vendas_Produtos = False
frmproj_produto.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuCompras_Rel_Historico_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Relatórios/Histórico"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_Relatorios_Historico2.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuCompras_Rel_Indice_Atraso_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Relatórios/Índice de atraso"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_Relatorios_Indice_Atraso.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuconjunto_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Engenharia/Conjuntos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmproj_conjunto.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucontprojetos_Click()
On Error GoTo tratar_erro
  
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
  
Formulario = "Engenharia/Controle de projetos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmControle_projetos.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucontrolecertificados_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Controle de certificados"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
'frmCertificado.Show
frmCQ_Certificado_Analise.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucotacao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Cotação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmcompras_reqcot.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucqnc_descricao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Não conformidade/Descrição da não conformidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmcqnc_descricaoNC.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnucqnc_NC_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Não conformidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
PCP_Ordem = False
frmcqnc.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuCustos_detalhado_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Custos/Relatórios/Detalhado"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_custos_relatorios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuCustos_resumido_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Custos/Relatórios/Resumido"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmRelatorios_Custos.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuDesempenho_Click()
On Error GoTo tratar_erro
'If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
'
'Formulario = "Financeiro/Contas pagas"
'ProcLiberaAcessos True
'If Acessos = False Then Exit Sub

frmVendas_Desempenho.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuestComp_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

If FunVerifAtualizacaoObrigatoria(True, False) = False Then
    Formulario = "Engenharia/Estrutura/Detalhada"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    frmproj_produto_estrutura.Show
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuEstResum_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

If FunVerifAtualizacaoObrigatoria(True, False) = False Then
    Formulario = "Engenharia/Estrutura/Resumida"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    frmproj_produto_estrutura_Resumida.Show
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuFluxoResumido_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Fluxo de caixa"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFluxo_Caixa_Grafico.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuMensalCST_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
'
'Formulario = "Financeiro/Relatórios/Razão"
'ProcLiberaAcessos True
'If Acessos = False Then Exit Sub

frmFaturamento_Impostos.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnumeta_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Relatórios/Comissão"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_Comissoes_Metas.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuRazaoDetalhado_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Relatórios/Razão"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFinanceiro_Relatorios_Razao_Detalhado.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuEMDetalhado_Click()
On Error GoTo tratar_erro
 

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Movimentação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

frmestoque_item.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuEMResumido_Click()
On Error GoTo tratar_erro
 

frmestoque_Movimentacao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuEnsaio_ControleCertificados_Click()
On Error GoTo tratar_erro
 
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
 
Formulario = "Qualidade/Ensaios/Controle de certificados"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCertificado_qualidade.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub mnuestrutura_Click()
'On Error GoTo tratar_erro
'
'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
'
'If FunVerifAtualizacaoObrigatoria(True, False) = False Then
'    Formulario = "Engenharia/Estrutura"
'    ProcLiberaAcessos True
'    If Acessos = False Then Exit Sub
'    frmproj_produto_estrutura.Show
'End If
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
End Sub

Private Sub mnufamilia_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Engenharia/Famílias"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_familia = "Compras/Famílias" Or Formulario_familia = "Vendas/Famílias" Or Formulario_familia = "Engenharia/Famílias" Or Formulario_familia = "Qualidade/Famílias" Then
    If FunVerifFormAberto(frmproj_familia) = True Then
        If USMsgBox("O módulo " & Formulario_familia & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_familia
    End If
End If

Compras_Familia = False
Vendas_Familia = False
Qualidade_Familia = False
frmproj_familia.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuFaturamento_Rel_dozemeses_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Relatórios/Doze últimos meses"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFaturamento_12ultimos_meses.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuFaturamento_Rel_Historico_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Relatórios/Histórico"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Faturamento = True
frmFaturamento_Relatorios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuFaturamento_Rel_Impostos_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Relatórios/Impostos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFaturamento_Relatorios_Impostos.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuFaturamento_Rel_Relacionamento_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Relatórios/Relacionamento de notas fiscais"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFaturamento_Relatorios_Relacionamento.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuferramenta_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
'
'Formulario = "Estoque/Movimentação"
'ProcLiberaAcessos True
'If Acessos = False Then Exit Sub
'frmestoque_item.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuFinanceiro_familias_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Plano de contas"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFinanceiro_familia.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnufluxo_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Fluxo de caixa"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFluxodecaixa.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuFMEA_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/PPAP/FMEA"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmQualidadePPAP_FMEA.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnufornecedor_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Fornecedores"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_fornecedores.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnufuncionarios_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "RH/Funcionários"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmRH_Funcionarios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnugrafeven_Click()
On Error GoTo tratar_erro
 
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
 
Formulario = "PCP/Relatórios/Monitor de eventos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmProd_Relatorios_Monitor_Eventos.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuhistorico_manut_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Manutenção/Relatórios/Histórico"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmManutencao_relatorios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuindiceatrazo_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Relatórios/Índice de atraso"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmRelatorios_indice_atraso.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuinstituicoes_Click()
On Error GoTo tratar_erro
    
If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Instituições"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frm_Instituicoes.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuinstrumentos_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Instrumentos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmInstrumentos.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuLiquido_Click()
On Error GoTo tratar_erro
 
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
 
Formulario = "Qualidade/Ensaios/Líquido penetrante"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmLiquido.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuListadeitens_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Outros/Follow up de compras"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_Requisicao_Lista.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnulocalarmaz_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Local de armazenamento"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmEstoque_Localarmaz.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnulogoff_Click()
On Error GoTo tratar_erro

If USMsgBox("Atenção " & pubUsuario & "!!!" & vbCrLf & "Você deseja fazer logoff do CAPRIND?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    ProcLogonOut
    frmabertura.Timer2.Enabled = True
    frmabertura.Show 1
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnumanutencao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Manutenção/Equipamentos"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmManutencao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuMensal_Click()
On Error GoTo tratar_erro

frm_Faturamento_Filtrar_Mes_CFOP.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuminuta_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Minuta de despacho"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmMinuta.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuMonitorTrabalho_Click()
On Error GoTo tratar_erro
  
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
  
Formulario = "PCP/Monitor de trabalho"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmMonitorTrab.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnunaoconformidade_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Relatórios/Não conformidade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmPCP_Relatorios_NC.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Mnunatureza_Click()
On Error GoTo tratar_erro
  
If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Fiscal/Natureza de operação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frm_Natureza_OP.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuNormas_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Engenharia/Normas"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmNorma.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuNotaEntrada_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Faturamento/Nota fiscal/Terceiros"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_nota = "Faturamento/Nota fiscal/Própria" Or Formulario_nota = "Estoque/Ordem de faturamento" Or Formulario_nota = "Estoque/Nota fiscal" Then
    If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
        If USMsgBox("O módulo " & Formulario_nota & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
    End If
End If

Formulario = "Faturamento/Nota fiscal/Terceiros"
Faturamento_NF_Saida = False
'======================================================
Faturamento_NF_Terceiro = True
Faturamento_NF_Propria = False
TPNota = "T"
'======================================================
frmFaturamento_Prod_Serv.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuNotas_Click()
On Error GoTo tratar_erro

frmFaturamento_FiltrarNotas.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuNotaSaida_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(True, False, True) = False Then
'Exit Sub
'End If

Formulario = "Faturamento/Nota fiscal/Própria"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_nota = "Faturamento/Nota fiscal/Terceiros" Or Formulario_nota = "Estoque/Ordem de faturamento" Or Formulario_nota = "Estoque/Nota fiscal" Then
    If FunVerifFormAberto(frmFaturamento_Prod_Serv) = True Then
        If USMsgBox("O módulo " & Formulario_nota & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmFaturamento_Prod_Serv
    End If
End If

Formulario = "Faturamento/Nota fiscal/Própria"
Faturamento_NF_Saida = True

'======================================================
Faturamento_NF_Terceiro = False
Faturamento_NF_Propria = True
TPNota = "P"
'======================================================

frmFaturamento_Prod_Serv.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuPlano_de_Controle_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/PPAP/Plano de controle"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmQualidadePPAP_PlanoControle.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuProcessosPrinc_Click()
On Error GoTo tratar_erro
 
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
 
If FunVerifAtualizacaoObrigatoria(True, False) = False Then
    Formulario = "Engenharia/Processos"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    frmProcessos.Show
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuProducaoDia_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
'
'Formulario = "PCP/Relatórios/Produtividade"
'ProcLiberaAcessos True
'If Acessos = False Then Exit Sub
frmProd_Producao_Dia.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuProdutividade_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Relatórios/Produtividade"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmProd_Relatorios_Produtividade.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuOptGeral_Click()
On Error GoTo tratar_erro

Dim Formulario1 As String
Dim Formulario2 As String
Dim Formulario3 As String
Dim Formulario4 As String
Dim Formulario5 As String

Formulario = "Configuração do sistema/Opções gerais/Configuração do sistema"
Formulario1 = "Configuração do sistema/Opções gerais/Cadastro de empresa"
Formulario2 = "Configuração do sistema/Opções gerais/Cadastro de moedas"
Formulario3 = "Configuração do sistema/Opções gerais/Cadastro de unidades"
Formulario4 = "Configuração do sistema/Opções gerais/Cadastro de condição de pagamento/recebimento"
Formulario5 = "Configuração do sistema/Opções gerais/Cadastro de feriados"

Set TBAcessos = CreateObject("adodb.recordset")
TBAcessos.Open "Select * FROM Acessos WHERE IDUsuario = " & pubIDUsuario & " AND (Acesso = '" & Formulario & "' or Acesso = '" & Formulario1 & "' or Acesso = '" & Formulario2 & "' or Acesso = '" & Formulario3 & "' or Acesso = '" & Formulario4 & "' or Acesso = '" & Formulario5 & "')", Conexao, adOpenKeyset, adLockOptimistic
If TBAcessos.EOF = True Then Exit Sub
TBAcessos.Close
frmOpcoesGeral.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuordem_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

If FunVerifAtualizacaoObrigatoria(True, False) = False Then
    Formulario = "PCP/Gerenciamento de ordem"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    frmprod.Show
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuPecasemproducao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Situação da produção"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Vendas = False
FrmSituacao_pedido_producao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnupedidointerno_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

If FunVerifAtualizacaoObrigatoria(True, False) = False Then
    Formulario = "Vendas/Pedido interno"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    Vendas_PI = True
    Vendas_Proposta = False
    frmVendas_PI.Show
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuplanoinspecao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Plano de inspeção"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmPlanoinspecao.Show
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuplanomedicao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Controle de medição"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmPlanomedicao.Show
      
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuProdCargaMaq_Click()
On Error GoTo tratar_erro
  
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Carga de posto de trabalho"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCargaMaq.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuprodutos_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Engenharia/Produtos e serviços"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_produto = "Compras/Produtos e serviços" Or Formulario_produto = "Vendas/Produtos e serviços" Or Formulario_produto = "Engenharia/Produtos e serviços" Then
    If FunVerifFormAberto(frmproj_produto) = True Then
        If USMsgBox("O módulo " & Formulario_produto & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_produto
    End If
End If

Engenharia_Produtos = True
Compras_Produtos = False
Vendas_Produtos = False
frmproj_produto.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuProgramacaodecompra_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Programação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Programacao = True
frmCompras_programacao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuprogramas_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Programas CNC"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmprogramacao.Show
  
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuPSW_Click()
On Error GoTo tratar_erro

Formulario = "Qualidade/PPAP/PSW"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmQualidadePPAP.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuqualidade_familia_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Famílias"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_familia = "Compras/Famílias" Or Formulario_familia = "Vendas/Famílias" Or Formulario_familia = "Engenharia/Famílias" Or Formulario_familia = "Qualidade/Famílias" Then
    If FunVerifFormAberto(frmproj_familia) = True Then
        If USMsgBox("O módulo " & Formulario_familia & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_familia
    End If
End If

Compras_Familia = False
Vendas_Familia = False
Qualidade_Familia = True
frmproj_familia.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub mnuRazaoResumido_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Relatórios/Razão"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmFinanceiro_Relatorios_Razao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnureceber_Click()
On Error GoTo tratar_erro
    
If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub
    
Formulario = "Financeiro/Contas a receber"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmContas_Receber.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnurecebernfconsignada_Click()
On Error GoTo tratar_erro
  
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
  
Formulario = "Estoque/Recebimento/Consignação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmEstoque_Recebimento_consignacao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnureceberpedidodecompra_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Estoque/Recebimento/Pedido de compra"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Programacao = False
frmEstoque_Recebimento.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnurecebidas_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Contas recebidas"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmContas_recebidas.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnurecebimento_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Inspeção de recebimento"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_recebimento.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuResultadosOrdem_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "PCP/Relatórios/Resultados da ordem"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmPCP_relatorios_resultados.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuRevisao_relatorios_Click()
On Error GoTo tratar_erro

'If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Histórico de revisão dos relatórios"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmQualidade_Revisao_Relatorios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuRNC_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/RNC"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
RNC_Inspecao_Recebimento = False
RNC_Controle_Medicao = False
RNC_Nao_Conformidade = False
RNC_Solicitacao_Desvio = False
frmQualidade_RNC.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuSA_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Solicitação de ação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCQ_SA.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnusair_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente abandonar o Sistema Caprind?", vbYesNo, "CAPRIND v5.0") = vbYes Then
    Call ProcLogonOut
    If Time$ > "19:00:00" Then
        USMsgBox "Boa noite " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind.", _
                vbInformation, "CAPRIND v5.0"
        ElseIf Time$ > "12:00:00" And Time$ < "18:59:59" Then
            USMsgBox "Boa Tarde " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind.", _
                    vbInformation, "CAPRIND v5.0"
        ElseIf Time$ > "00:00:00" And Time$ < "11:59:59" Then
            USMsgBox "Bom dia " & pubUsuario & ", Obrigado por utilizar o Sistema Caprind.", _
                vbInformation, "CAPRIND v5.0"
    End If
    FunFechaBD
    End
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub MnuSD_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub

Formulario = "Qualidade/Solicitação de desvio"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
RNC_Nao_Conformidade = False
frmCQ_SD.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuSituacaodaproposta_Click()
On Error GoTo tratar_erro
   
If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub
   
Formulario = "Vendas/Situação da produção"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Vendas = True
FrmSituacao_pedido_producao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuSolicitacaoDeProducao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Outros/Solicitação de produção"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmOutros_Solicitacao_PCP.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuSub_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

'Formulario = "Fiscal/Substituicao"
'ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmSubstituicao_Tributaria.Show 1


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuSuportetecnico_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Suporte/Chamado"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
If TemInternet = True And ErroDriverMYSQL = False Then
    If FunVerificaManutencaoAtiva = False Then Exit Sub
    frmSuporte.Show
Else
    If TemInternet = False Then MsgTexto = "não foi encontrado conexão com a internet" Else MsgTexto = "no momento estamos sem conexão com nosso servidor, favor tentar mais tarde"
    USMsgBox ("Não é permitido abrir este módulo, pois " & MsgTexto & "."), vbExclamation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuTelemarketing_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Telemarketing"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_Tele_Clientes.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuucpagas_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, False) = False Then Exit Sub

Formulario = "Financeiro/Contas pagas"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmContas_Pagas.Show
 
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuUltraSom_Click()
On Error GoTo tratar_erro
 
If FunVefificaModuloLocacao(False, True, False) = False Then Exit Sub
 
Formulario = "Qualidade/Ensaios/Ultra-som"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmUltraSom.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuPedidodecompra_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, False, True) = False Then Exit Sub

Formulario = "Compras/Pedido"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_Pedido.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuSolicitaçãodecompra_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Outros/Solicitação"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmCompras_Requisicao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuPropostacomercial_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

If FunVerifAtualizacaoObrigatoria(True, False) = False Then
    Formulario = "Vendas/Proposta comercial"
    ProcLiberaAcessos True
    If Acessos = False Then Exit Sub
    Vendas_PI = False
    Vendas_Proposta = True
    frmVendas_proposta.Show
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuCarteiradepedidos_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Follow up"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_carteira.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuControledevendedores_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Vendedores"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_Vendedores.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuInformacoesdefaturamento_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Informações faturamento"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
Faturamento = False
frmFaturamento_Relatorios.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuvendas_familia_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Famílias"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_familia = "Compras/Famílias" Or Formulario_familia = "Vendas/Famílias" Or Formulario_familia = "Engenharia/Famílias" Or Formulario_familia = "Qualidade/Famílias" Then
    If FunVerifFormAberto(frmproj_familia) = True Then
        If USMsgBox("O módulo " & Formulario_familia & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_familia
    End If
End If

Compras_Familia = False
Vendas_Familia = True
Qualidade_Familia = False
frmproj_familia.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuvendas_itens_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Produtos e serviços"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub

If Formulario_produto = "Compras/Produtos e serviços" Or Formulario_produto = "Vendas/Produtos e serviços" Or Formulario_produto = "Engenharia/Produtos e serviços" Then
    If FunVerifFormAberto(frmproj_produto) = True Then
        If USMsgBox("O módulo " & Formulario_produto & " está aberto, deseja fechá-lo para prosseguir?", vbYesNo, "CAPRIND v5.0") = vbNo Then Exit Sub Else Unload frmproj_produto
    End If
End If

Engenharia_Produtos = False
Compras_Produtos = False
Vendas_Produtos = True
frmproj_produto.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuVendas_Rel_comissao_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Relatórios/Comissão"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_comissao.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuVendas_Rel_Historico_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Relatórios/Histórico"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_Relatorios_Historico.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub mnuVendas_Rel_indice_atrazo_Click()
On Error GoTo tratar_erro

If FunVefificaModuloLocacao(True, True, True) = False Then Exit Sub

Formulario = "Vendas/Relatórios/Índice de atraso"
ProcLiberaAcessos True
If Acessos = False Then Exit Sub
frmVendas_Relatorios_Indice_Atraso.Show

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



'Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
'On Error GoTo tratar_erro
'
'If StatusBar1.Panels(4).Visible = True Then
''FrmMDI_AvisoDiario.Show
'End If
'
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

'Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
'On Error GoTo tratar_erro
'
'If StatusBar1.Panels(4).Visible = True Then
''FrmMDI_AvisoDiario.Show
'End If
'
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Private Sub Timer_avisodiario_Timer()
On Error GoTo tratar_erro
Dim TextoOr As String 'OK

Permitido = False

TextoOr = "Acesso Like 'Avisos diário%'"
'TextoOr = "(Acesso = 'Avisos diário/Estoque/Produtos á vencer' or Acesso = 'Avisos diário/Solicitação' or Acesso = 'Avisos diário/Contas a pagar'" _
& "or Acesso = 'Avisos diário/Contas a receber' or Acesso = 'Avisos diário/Manutenção' or Acesso = 'Avisos diário/Requisição de materiais' or Acesso = 'Avisos diário/Compras/Necessidade' or Acesso = 'Avisos diário/PCP/Necessidade' or Acesso = 'Avisos diário/Estoque/Necessidade' or Acesso = 'Avisos diário/Faturamento/Carteira de faturamento' or Acesso = 'Avisos diário/PCP/OSs em atraso' or Acesso = 'Avisos diário/Custos/Centro de custo' or Acesso = 'Avisos diário/Análise crítica/Engenharia' or Acesso = 'Avisos diário/Análise crítica/Processos' or Acesso = 'Avisos diário/Análise crítica/Pcp' or Acesso = 'Avisos diário/Análise crítica/Qualidade' or Acesso = 'Avisos diário/Análise crítica/Compras' or Acesso = 'Avisos diário/Terceiros' or Acesso = 'Avisos diário/Compras/Pedidos em atraso' or Acesso = 'Avisos diário/Qualidade/Calibração a vencer' or Acesso = 'Avisos diário/Qualidade/Não conformidades')"
Set TBLISTA = CreateObject("adodb.recordset")
StrSql = "Select Acesso from acessos where IDUsuario = " & pubIDUsuario & " and " & TextoOr
'Debug.print StrSql
'Debug.print TextoOr

TBLISTA.Open "Select IDAcesso, Acesso from acessos where IDUsuario = " & pubIDUsuario & " and " & TextoOr, Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Do While TBLISTA.EOF = False
    'Debug.print TBLISTA!IDAcesso
        Set TBAbrir = CreateObject("adodb.recordset")
        Select Case TBLISTA!Acesso
            Case "Avisos diário/Processos/Sugestões"
                TBAbrir.Open "Select * From Fases_Sugestao Where Status = '1'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
                TBAbrir.Close
            Case "Avisos diário/Estoque/Produtos á vencer"
                TBAbrir.Open "Select * From Estoque_Produtos_Vencer", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
                TBAbrir.Close
            Case "Avisos diário/Contas a pagar"
                TBAbrir.Open "Select IdIntConta from tbl_ContasPagar where dt_Pagamento <= '" & Format(Date, "Short Date") & "' and Logsit = 'N' and Bloqueado = 'False' and Status <> 'TÍTULO LIQUIDADO ANTECIPADO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
                TBAbrir.Close
            Case "Avisos diário/Contas a receber"
                TBAbrir.Open "Select IdIntConta from tbl_contas_receber where Vencimento <= '" & Format(Date, "Short Date") & "' and Logsit = 'N' and Bloqueado = 'False' and Status <> 'TÍTULO LIQUIDADO ANTECIPADO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
                TBAbrir.Close
            Case "Avisos diário/Solicitação"
                TBAbrir.Open "Select ID_Requisicao from Compras_requisicao where Data_Solicitacao <= '" & Format(Date, "Short Date") & "' and Status = 'ABERTA' and DtValidacao IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
                TBAbrir.Close
            Case "Avisos diário/Manutenção"
                TBAbrir.Open "Select M.codigo, M.Tipo, M.IDMaquina, MD.data, MD.Dias_proxima from Manutencao_data MD INNER JOIN Manutencao M on MD.idManutencao = M.codigo where (MD.Data <= '" & Format(Date, "Short Date") & "' or (MD.Data + MD.Dias_proxima) <= '" & Format(Date, "Short Date") & "') and MD.Status = 'ABERTA' and M.Tipo <> 'C'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
                TBAbrir.Close
            Case "Avisos diário/Requisição de materiais"
                TBAbrir.Open "Select RML.* from (Requisicao_materiais RM INNER JOIN Requisicao_materiais_lista RML ON RM.ID = RML.IDrequisicao) LEFT JOIN Usuarios_Setor_Responsavel ON Usuarios_Setor_Responsavel.ID_CC = RML.ID_CC where RML.Status = 'REQUISIT.' and RML.Data_autorizacao is null and (Usuarios_Setor_Responsavel.Responsavel_CC = '" & pubUsuario & "' or RML.ID_CC IS NULL)", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Compras/Necessidade"
                TBAbrir.Open "Select codproduto from Estoque_necessidade_resumido where Compras = 'True' and Necessidade > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/PCP/Necessidade"
                TBAbrir.Open "Select codproduto from Estoque_necessidade_resumido where Producao = 'True' and Necessidade > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Estoque/Necessidade"
                TBAbrir.Open "Select codproduto from Estoque_necessidade_resumido where Necessidade > 0", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Faturamento/Carteira de faturamento"
                TBAbrir.Open "Select NF.ID, E.Empresa, NF.dt_DataEmissao, NF.int_NotaFiscal, NF.TipoNF, NF.Serie, NF.Id_Int_Cliente, NF.txt_Razao_Nome from tbl_Dados_Nota_Fiscal NF INNER JOIN Empresa E ON NF.ID_empresa = E.Codigo where NF.Aplicacao = 'P' and NF.DtValidacaoOF IS NOT NULL and NF.int_NotaFiscal IS NULL order by NF.ID", Conexao, adOpenKeyset, adLockReadOnly
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/PCP/OSs em atraso"
                TBAbrir.Open "Select IDProducao from ordemservico where Prazofinal < '" & Format(Date, "Short Date") & "' and pronto = 'NÃO'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Custos/Centro de custo"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select IDAcesso from Acessos where IDUsuario = " & pubIDUsuario & " and Acesso = 'Custos/Centro de custo/Visualizar todos'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAcessos.EOF = False Then
                    TBAbrir.Open "Select ID from CC_realizado where Data = '" & Format(Date, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                Else
                    TBAbrir.Open "Select USR.ID from Usuarios_Setor_Responsavel USR INNER JOIN CC_realizado CC on USR.ID_CC = CC.ID_CC where USR.Responsavel_CC = '" & pubUsuario & "' and CC.Data = '" & Format(Date, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                End If
                TBAcessos.Close
                
                If TBAbrir.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Análise crítica/Engenharia"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select ID from Vendas_analise where DtValidacao_Engenharia IS NULL and Status = 'ABERTA EM ANALISE'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAcessos.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Análise crítica/Processos"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select ID from Vendas_analise where DtValidacao_PCP IS NULL and Status = 'ABERTA EM ANALISE'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAcessos.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Análise crítica/Pcp"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select ID from Vendas_analise where DtValidacao_PCP IS NULL and Status = 'ABERTA EM ANALISE'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAcessos.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Análise crítica/Qualidade"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select ID from Vendas_analise where DtValidacao_Qualidade IS NULL and Status = 'ABERTA EM ANALISE'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAcessos.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Análise crítica/Compras"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select ID from Vendas_analise where DtValidacao_Compras IS NULL and Status = 'ABERTA EM ANALISE'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAcessos.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Terceiros"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select A.Dias_Terceiros from Usuarios U INNER JOIN Acessos A on U.IDusuario = A.IDusuario where U.IDusuario = " & pubIDUsuario & " and A.Dias_Terceiros IS NOT NULL and A.Acesso = 'Avisos diário/Terceiros'", Conexao, adOpenKeyset, adLockReadOnly
                If TBAcessos.EOF = False Then
                    Dataini = Date
                    Dataini = Dataini + TBAcessos!Dias_Terceiros
                    Set TBAcessos = CreateObject("adodb.recordset")
                    TBAcessos.Open "SELECT CPL.IDlista FROM (Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CP.IDPedido = CPL.IDPedido) LEFT JOIN Estoque_controle_recebimento ECR ON ECR.IdLista = CPL.IdLista WHERE CPL.OS IS NOT NULL and (CPL.Status_Item = 'N_RECEBIDO' or CPL.Status_Item = 'PARCIAL') GROUP BY CP.IDPedido, CP.Pedido, CP.Fornecedor, CPL.IDlista, CPL.Desenho, CPL.Descricao, CPL.Prazo, CPL.Quant_Comp, CPL.Ordem, CPL.OS, CPL.Obs_AvisoDiario HAVING CPL.Prazo <= '" & Format(Dataini, "Short Date") & "' ORDER BY CP.IDPedido", Conexao, adOpenKeyset, adLockReadOnly
                    If TBAcessos.EOF = False Then
                        Permitido = True
                        GoTo Pular
                    End If
                End If
            Case "Avisos diário/Compras/Pedidos em atraso"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "SELECT CPL.IdLista FROM (Compras_pedido CP INNER JOIN Compras_pedido_lista CPL ON CP.IDPedido = CPL.IDPedido) LEFT JOIN Estoque_controle_recebimento ECR ON ECR.IdLista = CPL.IdLista WHERE CPL.Remessa = 'False' and CPL.OS IS NULL and (CPL.Status_Item = 'N_RECEBIDO' or CPL.Status_Item = 'PARCIAL') GROUP BY CP.IDPedido, CPL.IdLista, CPL.Prazo, CPL.Quant_Comp HAVING CPL.Prazo < '" & Format(Date, "Short Date") & "' ORDER BY CP.IDPedido", Conexao, adOpenKeyset, adLockReadOnly
                If TBAcessos.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Qualidade/Calibração a vencer"
                Dataini = Date + 5
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "Select I.CODIGO from (((Instrumentos I LEFT JOIN Estoque_controle EC ON EC.IDestoque = I.IDestoque) LEFT JOIN Projproduto P ON P.Desenho = I.Numero) LEFT JOIN item_aplicacoes IA ON IA.Codproduto = P.Codproduto) INNER JOIN Afericao A ON I.Codigo = A.ID_inst and I.ID_ultima_afericao = A.Codigo where A.proxima_afericao <= '" & Format(Dataini, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAcessos.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
            Case "Avisos diário/Qualidade/Não conformidades"
                Set TBAcessos = CreateObject("adodb.recordset")
                TBAcessos.Open "SELECT NC.CODIGO FROM cq_nc_fabrica NC LEFT JOIN ordemservico OS ON NC.OS = OS.idproducao WHERE NC.analizada = 'False'", Conexao, adOpenKeyset, adLockReadOnly
                If TBAcessos.EOF = False Then
                    Permitido = True
                    GoTo Pular
                End If
        End Select
        TBLISTA.MoveNext
    Loop
End If
Pular:
    TBLISTA.Close
    If Permitido = True Then
        With frmMenucaprind_menulateral
'            .lblAvisoDiário1.Caption = "Atenção " & pubUsuario & "! Você tem novo(s) aviso(s)."
'            .lblAvisoDiário1.ToolTipText = "Atenção " & pubUsuario & "! Você tem novo(s) aviso(s)."
'            .lblAvisodiario2.Caption = "Clique aqui para visualizar."
'            .lblAvisodiario2.ToolTipText = "Clique aqui para visualizar."
            
'            .Timer1.Enabled = True
            
            Set .USSysTray1.IconPicture = Me.Icon
            .USSysTray1.toolTip = "Aviso(s) disponível(eis)." & vbCrLf & "Clique no ícone para visualizar."
            .USSysTray1.SysTrayAddIcon
            .USSysTray1.BalloonTipShow "CAPRIND V5.0", "ATENÇÃO!" & vbCrLf & "Aviso(s) disponível(eis)." & vbCrLf & "Clique no ícone para visualizar.", ICON_INFO
            .USSysTray1.BlinkStart 1
            .USSysTray1.BalloonTipShowLast
        End With
    End If
    

Exit Sub
tratar_erro:
    If Err.Number = 401 Then
        Timer_avisodiario.Enabled = True
        Exit Sub
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaAvisoDiario()
On Error GoTo tratar_erro

Set TBTempo = CreateObject("adodb.recordset")
TBTempo.Open "Select * from Usuarios where IDUsuario = " & pubIDUsuario & " and Aviso_diario = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBTempo.EOF = False Then
    Timer_avisodiario.Enabled = True
Else
    Timer_avisodiario.Enabled = False
End If
TBTempo.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcVerificaLogoffAutomatico()
On Error GoTo tratar_erro

If pubUsuario <> "PROCAM" Then
    Set TBTempo = CreateObject("adodb.recordset")
    TBTempo.Open "Select * from Empresa where Verificar_desconectar_usuario = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBTempo.EOF = False Then
        Timer_logoff_automatico.Enabled = True
    Else
        Timer_logoff_automatico.Enabled = False
    End If
    TBTempo.Close
Else
    Timer_logoff_automatico.Enabled = False
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Timer_logoff_automatico_Timer()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Empresa where Verificar_desconectar_usuario = 'True'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Set TBFIltro = CreateObject("adodb.recordset")
    TBFIltro.Open "Select * from Logon where IDlogon = " & IDlogon, Conexao, adOpenKeyset, adLockOptimistic
    If TBFIltro.EOF = False Then
        ElapsedTime (TBFIltro!Hora_ultimo_evento)
        TPPSEG = TBAbrir!Minutos_desconectar * 60
        Dataini = FormataTempo(s + TPPSEG)
        DataFim = Time
        If DataFim >= Dataini Then
            ProcLogonOut1 IDlogon, pubUsuario, "C"
            Timer_logon.Enabled = False
            If USMsgBox("O usuário " & pubUsuario & " foi desconectado do sistema, deseja reconectar?", vbYesNo, "CAPRIND v5.0") = vbNo Then
                End
            Else
                ProcVerificaInternet False, False 'Verifica conexão com a internet
                If TemInternet = True And ErroDriverMYSQL = False Then FunValidarCliente Else FunValidarClienteSemInternet
                FunLogonIn False
                Timer_logon.Enabled = True
            End If
        End If
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Timer_logon_Timer()
On Error GoTo tratar_erro

If Err.Number = -2147467259 Or Err.Number = 3709 Then
    FunAbreBD
    GoTo 1
End If
1:
    Set TBLogon = CreateObject("adodb.recordset")
    TBLogon.Open "Select IDLogon from Logon where IDLogon = " & IDlogon, Conexao, adOpenKeyset, adLockOptimistic
    If TBLogon.EOF = True Then
        If pubUsuario <> "PROCAM" Then
            USMsgBox ("O usuário " & pubUsuario & " foi desconectado do sistema, o sistema Caprind será encerrado."), vbCritical, "CAPRIND v5.0"
            End
        End If
    End If
    TBLogon.Close
    
Exit Sub
tratar_erro:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        FunAbreBD
        GoTo 1
    End If
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub picBot_Resize()
On Error GoTo tratar_erro

'Arrange in case width change
If picBot.Width > 100 Then
    TreeView1.Width = picBot.Width - 100
End If

'Arrange in case height change
If picBot.Height - TreeView1.Top > 150 Then
    TreeView1.Height = picBot.Height - TreeView1.Top - 150
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub picHBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo tratar_erro
    
bDrag = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub picHBar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo tratar_erro

bDrag = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub picVBar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo tratar_erro

bDrag = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub picVBar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo tratar_erro
    
'Adjust the width if the user is holding down the mouse button
'If bDrag = True Then
    'picMain.Width = X + picMain.Width
'End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub picVBar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo tratar_erro

bDrag = False

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub RefControls()
On Error GoTo tratar_erro

 'Refresh all of the controls
' lblXTop.Refresh
' lblTasks.Refresh
' lstTasks.Refresh
' lblXBot.Refresh
' lblRes.Refresh
 TreeView1.Refresh
 picMain.Refresh
' picHBar.Refresh
 picVBar.Refresh
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TimerProdutos_Timer()
On Error GoTo tratar_erro
    
'procAtualizaProdutosWEB

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo tratar_erro

ProcAbreModuloMenuTreeView (Node.key)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


