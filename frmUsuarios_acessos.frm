VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmUsuarios_acessos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Configurações do sistema - Usuários - Definir acessos por módulo"
   ClientHeight    =   6960
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   7605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsuarios_acessos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   979
      DibPicture      =   "frmUsuarios_acessos.frx":000C
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmUsuarios_acessos.frx":09AA
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   6555
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   570
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Salvar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Salvar (F3)"
      ButtonKey1      =   "3"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   44
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   48
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "13"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   52
      ButtonTop3      =   2
      ButtonWidth3    =   41
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "14"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   95
      ButtonTop4      =   2
      ButtonWidth4    =   30
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "15"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   127
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   4260
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmUsuarios_acessos.frx":09C6
         Count           =   1
      End
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4425
      Left            =   180
      TabIndex        =   0
      Top             =   1680
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   7805
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Módulo"
         Object.Width           =   11289
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   6120
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   "Atualizando..."
      Value           =   0
   End
End
Attribute VB_Name = "frmUsuarios_acessos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcSalvar()
On Error GoTo tratar_erro

If Alterar = False Then
    USMsgBox ("Atenção usuário " & pubUsuario & ", voce não está autorizado a alterar neste formulário."), vbInformation, "CAPRIND v5.0"
    Exit Sub
End If
With Lista
    For InitFor = 1 To .ListItems.Count
        If .ListItems.Item(InitFor).Checked = True Then
            Set TBGravar = CreateObject("adodb.recordset")
            TBGravar.Open "Select * from acessos where IDUsuario = " & frmUsuarios.txtId & " and Acesso = '" & .ListItems(InitFor).ListSubItems(1) & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBGravar.EOF = True Then
                TBGravar.AddNew
                TBGravar!Data = Date
                TBGravar!Responsavel = pubUsuario
                TBGravar!Incluir = True
                TBGravar!Alterar = True
                TBGravar!Excluir = True
                TBGravar!Validacao = True
                TBGravar!IDUsuario = frmUsuarios.txtId
                TBGravar!Acesso = .ListItems(InitFor).ListSubItems(1)
                TBGravar.Update
                '==================================
                Modulo = "Configuração do sistema/Usuários"
                Evento = "Novo acesso"
                ID_documento = TBGravar!IDAcesso
                Documento = "Usuario: " & frmUsuarios.txtUsuario
                Documento1 = "Acesso: " & .ListItems(InitFor).ListSubItems(1)
                ProcGravaEvento
                '==================================
            End If
            TBGravar.Close
        Else
            Conexao.Execute "DELETE from acessos where IDUsuario = " & frmUsuarios.txtId & " and acesso = '" & .ListItems(InitFor).ListSubItems(1) & "'"
        End If
    Next InitFor
End With
USMsgBox ("Definição de acesso(s) cadastrado(s) com sucesso."), vbInformation, "CAPRIND v5.0"
Direitos
frmUsuarios.ProcCarregaLista_acesso
ProcRecarregaMenu
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcSalvar
    'Case vbKeyF1: cmdAjuda_Click
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

'contador = 0
'PBLista.Min = 0
'PBLista.Max = 206
'PBLista.Value = 1
'Do While contador <> 206
'    contador = contador + 1
'    PBLista.Value = contador
    
    With Lista.ListItems
    .Clear
        .Add , , "Configuração do sistema/Opções gerais"
        .Item(.Count).SubItems(1) = "Configuração do sistema/Opções gerais"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Opções gerais/Configuração do sistema"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Opções gerais/Cadastro de empresa"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Opções gerais/Cadastro de moedas"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Opções gerais/Cadastro de unidades"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Opções gerais/Cadastro de condição de pagamento/recebimento"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Opções gerais/Cadastro de feriados"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Usuários"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Usuários/Eventos realizados"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Usuários/Conectados"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Criar backup/Configurações"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Criar backup/Apontamentos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Criar backup/Eventos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Reindexar BD/Caprind e Gerprod"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Configuração do sistema/Reindexar BD/GNFe"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "RH/Funcionários"
        .Add , , ""
        .Item(.Count).SubItems(1) = "RH/Relatórios/Desoneração da folha de pagamento"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Famílias"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Produtos e serviços"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Produtos e serviços/Valores e descontos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Produtos e serviços/Clientes e fornecedores"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Produtos e serviços/Validar estrutura"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Produtos e serviços/Validar plano de inspeção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Fornecedores"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Programação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Cotação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Cotação/Liberar cotação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Pedido"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Pedido/Aprovar"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Necessidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Não conformidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Atualização de valores"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Relatórios/Histórico"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Relatórios/Índice de atraso"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Compras/Autorização de centro de custo sem previsão"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Famílias"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Produtos e serviços"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Produtos e serviços/Valores e descontos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Produtos e serviços/Valores e descontos/Visualizar valor de custo"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Produtos e serviços/Clientes e fornecedores"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Produtos e serviços/Validar estrutura"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Produtos e serviços/Validar plano de inspeção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Clientes"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Vendedores"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Simulação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Empenho"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Telemarketing"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Programação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Proposta comercial"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Pedido interno"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Follow up"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Situação da produção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Informações faturamento"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Pós-vendas/Assistência técnica"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Atualização de valores"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Relatórios/Histórico"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Relatórios/Índice de atraso"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Vendas/Relatórios/Comissão"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Plano de contas"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Instituições"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Contas a pagar"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Contas a pagar/Visualizar contas dos funcionários"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Contas pagas"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Contas pagas/Visualizar contas dos funcionários"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Contas a receber"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Contas recebidas"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Desconto de duplicata"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Fluxo de caixa"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Relatórios/Histórico"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Relatórios/Razão"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Financeiro/Autorização de centro de custo sem previsão"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Fiscal/Classificação fiscal"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Fiscal/Natureza de operação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Nota fiscal/Terceiros"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Nota fiscal/Própria"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Nota fiscal/Cancelar nota"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Nota fiscal/Excluir duplicatas"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Nota fiscal/SPED"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Nota fiscal/Exportar"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Carta de correção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Minuta de despacho"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Relatórios/Histórico"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Relatórios/Relacionamento de notas fiscais"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Relatórios/Impostos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Relatórios/Doze últimos meses"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Faturamento/Autorização de centro de custo sem previsão"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Custos/Centro de custo"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Custos/Centro de custo/Visualizar todos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Custos/Centro de custo/Visualizar lançamentos realizados"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Custos/Relatórios/Detalhado"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Custos/Relatórios/Resumido"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Custos/Relatórios/Previsto x Realizado"

        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Famílias"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Produtos e serviços"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Produtos e serviços/Validar estrutura"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Produtos e serviços/Validar plano de inspeção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Conjuntos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Estrutura/Detalhada"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Estrutura/Resumida"
        .Add , , ""
        
        .Item(.Count).SubItems(1) = "Engenharia/Estrutura/Visualizar valor de custo"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Controle de projetos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Processos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Processos/Histórico"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Engenharia/Normas"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Postos de trabalho"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Códigos de trabalho"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Carga de posto de trabalho"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Gerenciamento de ordem"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Gerenciamento de ordem/Validar resultados"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Monitor de trabalho"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Programas CNC"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Situação da produção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Necessidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Não conformidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Programação da produção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Plano da produção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Requisição da ordem"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Relatórios/Produtividade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Relatórios/Não conformidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Relatórios/Monitor de eventos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Relatórios/Índice de atraso"
        .Add , , ""
        .Item(.Count).SubItems(1) = "PCP/Relatórios/Resultados da ordem"

        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Famílias"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Instrumentos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Almoxarifado"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Plano de inspeção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Controle de medição"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Inspeção de recebimento"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Ensaios/Ultra-som"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Ensaios/Líquido penetrante"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Ensaios/Controle de certificados"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Controle de certificados"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Controle de documentos e dados"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Não conformidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Não conformidade/Descrição da não conformidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Não conformidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Solicitação de ação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Solicitação de desvio"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/RNC"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/PPAP"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/PPAP/PSW"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/PPAP/FMEA"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/PPAP/Plano de controle"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Histórico de revisão dos relatórios"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Relatórios/Não conformidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Qualidade/Relatórios/Histórico"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Almoxarifado"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Local de armazenamento"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Requisição de materiais"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Recebimento/Pedido de compra"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Recebimento/Consignação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Recebimento/Importar nota de terceiros"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Inventário"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Movimentação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Controle de certificado"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Não conformidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Necessidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Ordem de faturamento"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Nota fiscal"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Estoque/Autorização de centro de custo sem previsão"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Manutenção/Equipamentos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Manutenção/Equipamentos/Aprovar manutenção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Manutenção/Assistência técnica"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Manutenção/Relatórios/Histórico"
    
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Solicitação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Solicitação/Autorizar solicitação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Solicitação de produção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Solicitação de produção/Autorizar solicitação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Follow up de compras"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Validação de procedimentos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Análise crítica"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Análise crítica/Engenharia"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Análise crítica/Processos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Análise crítica/Pcp"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Análise crítica/Qualidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Análise crítica/Compras"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Análise crítica/Vendas"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Análise crítica/Documentos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Downloads/Nota fiscal"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Outros/Downloads/Boleto"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Suporte/Chamado"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Suporte/Chat (online)"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Suporte/Solicitação de atendimento"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Suporte/Atualização/Caprind e Gerprod"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Suporte/Atualização/GNFe"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Suporte/Atualização/GMRE (relatórios)"
        
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Solicitação"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Contas a pagar"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Contas a receber"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Manutenção"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Requisição de materiais"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Compras/Necessidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/PCP/Necessidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Estoque/Necessidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Faturamento/Carteira de faturamento"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/PCP/OSs em atraso"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Custos/Centro de custo"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Análise crítica/Engenharia"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Análise crítica/Processos"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Análise crítica/Pcp"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Análise crítica/Qualidade"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Análise crítica/Compras"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Terceiros"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Compras/Pedidos em atraso"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Qualidade/Calibração a vencer"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Qualidade/Não conformidades"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Estoque/Produtos á vencer"
        .Add , , ""
        .Item(.Count).SubItems(1) = "Avisos diário/Processos/Sugestões"
    End With
    
contador = Lista.ListItems.Count
Do While contador > 0
    With Lista.ListItems
    Modulos = Lista.ListItems.Item(contador).ListSubItems(1).Text
        Set TBAcessos = CreateObject("adodb.recordset")
        TBAcessos.Open "Select * from Acessos where IDUsuario = " & frmUsuarios.txtId & " and Acesso = '" & Modulos & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAcessos.EOF = False Then
            .Item(contador).Checked = True
        End If
        TBAcessos.Close
    End With
    contador = contador - 1
Loop


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
'End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 7575, 5, True
ProcCarregaLista

'contador = 0
'PBLista.Min = 0
'PBLista.Max = 206
'PBLista.Value = 1
'Do While contador <> 206
'    contador = contador + 1
'    PBLista.Value = contador
'    Select Case contador
'        Case 1: Modulos = "Configuração do sistema/Opções gerais"
'            Case 2: Modulos = "Configuração do sistema/Opções gerais/Configuração do sistema"
'            Case 3: Modulos = "Configuração do sistema/Opções gerais/Cadastro de empresa"
'            Case 4: Modulos = "Configuração do sistema/Opções gerais/Cadastro de moedas"
'            Case 5: Modulos = "Configuração do sistema/Opções gerais/Cadastro de unidades"
'            Case 6: Modulos = "Configuração do sistema/Opções gerais/Cadastro de condição de pagamento/recebimento"
'            Case 7: Modulos = "Configuração do sistema/Opções gerais/Cadastro de feriados"
'        Case 8: Modulos = "Configuração do sistema/Usuários"
'            Case 9: Modulos = "Configuração do sistema/Usuários/Eventos realizados"
'            Case 10: Modulos = "Configuração do sistema/Usuários/Conectados"
'        Case 11: Modulos = "Configuração do sistema/Criar backup/Configurações"
'        Case 12: Modulos = "Configuração do sistema/Criar backup/Apontamentos"
'        Case 13: Modulos = "Configuração do sistema/Criar backup/Eventos"
'        Case 14: Modulos = "Configuração do sistema/Reindexar BD/Caprind e Gerprod"
'        Case 15: Modulos = "Configuração do sistema/Reindexar BD/GNFe"
'
'        Case 16: Modulos = "RH/Funcionários"
'        Case 17: Modulos = "RH/Relatórios/Desoneração da folha de pagamento"
'
'        Case 18: Modulos = "Compras/Famílias"
'        Case 19: Modulos = "Compras/Produtos e serviços"
'            Case 20: Modulos = "Compras/Produtos e serviços/Valores e descontos"
'            Case 21: Modulos = "Compras/Produtos e serviços/Clientes e fornecedores"
'            Case 22: Modulos = "Compras/Produtos e serviços/Validar estrutura"
'            Case 23: Modulos = "Compras/Produtos e serviços/Validar plano de inspeção"
'        Case 24: Modulos = "Compras/Fornecedores"
'        Case 25: Modulos = "Compras/Programação"
'        Case 26: Modulos = "Compras/Cotação"
'            Case 27: Modulos = "Compras/Cotação/Liberar cotação"
'        Case 28: Modulos = "Compras/Pedido"
'        Case 29: Modulos = "Compras/Pedido/Aprovar"
'        Case 30: Modulos = "Compras/Necessidade"
'        Case 31: Modulos = "Compras/Não conformidade"
'        Case 32: Modulos = "Compras/Atualização de valores"
'        Case 33: Modulos = "Compras/Relatórios/Histórico"
'        Case 34: Modulos = "Compras/Relatórios/Índice de atraso"
'        Case 35: Modulos = "Compras/Autorização de centro de custo sem previsão"
'
'        Case 36: Modulos = "Vendas/Famílias"
'        Case 37: Modulos = "Vendas/Produtos e serviços"
'            Case 38: Modulos = "Vendas/Produtos e serviços/Valores e descontos"
'                Case 39: Modulos = "Vendas/Produtos e serviços/Valores e descontos/Visualizar valor de custo"
'            Case 40: Modulos = "Vendas/Produtos e serviços/Clientes e fornecedores"
'            Case 41: Modulos = "Vendas/Produtos e serviços/Validar estrutura"
'            Case 42: Modulos = "Vendas/Produtos e serviços/Validar plano de inspeção"
'        Case 43: Modulos = "Vendas/Clientes"
'        Case 44: Modulos = "Vendas/Vendedores"
'        Case 45: Modulos = "Vendas/Simulação"
'        Case 46: Modulos = "Vendas/Empenho"
'        Case 47: Modulos = "Vendas/Telemarketing"
'        Case 48: Modulos = "Vendas/Programação"
'        Case 49: Modulos = "Vendas/Proposta comercial"
'        Case 50: Modulos = "Vendas/Pedido interno"
'        Case 51: Modulos = "Vendas/Follow up"
'        Case 52: Modulos = "Vendas/Situação da produção"
'        Case 53: Modulos = "Vendas/Informações faturamento"
'        Case 54: Modulos = "Vendas/Pós-vendas/Assistência técnica"
'        Case 55: Modulos = "Vendas/Atualização de valores"
'        Case 56: Modulos = "Vendas/Relatórios/Histórico"
'        Case 57: Modulos = "Vendas/Relatórios/Índice de atraso"
'        Case 58: Modulos = "Vendas/Relatórios/Comissão"
'
'        Case 59: Modulos = "Financeiro/Plano de contas"
'        Case 60: Modulos = "Financeiro/Instituições"
'        Case 61: Modulos = "Financeiro/Contas a pagar"
'        Case 62: Modulos = "Financeiro/Contas a pagar/Visualizar contas dos funcionários"
'        Case 63: Modulos = "Financeiro/Contas pagas"
'        Case 64: Modulos = "Financeiro/Contas pagas/Visualizar contas dos funcionários"
'        Case 65: Modulos = "Financeiro/Contas a receber"
'        Case 66: Modulos = "Financeiro/Contas recebidas"
'        Case 67: Modulos = "Financeiro/Desconto de duplicata"
'        Case 68: Modulos = "Financeiro/Fluxo de caixa"
'        Case 69: Modulos = "Financeiro/Relatórios/Histórico"
'        Case 70: Modulos = "Financeiro/Relatórios/Razão"
'        Case 71: Modulos = "Financeiro/Autorização de centro de custo sem previsão"
'
'        Case 72: Modulos = "Faturamento/Fiscal/Classificação fiscal"
'        Case 73: Modulos = "Faturamento/Fiscal/Natureza de operação"
'        Case 74: Modulos = "Faturamento/Nota fiscal/Terceiros"
'        Case 75: Modulos = "Faturamento/Nota fiscal/Própria"
'        Case 76: Modulos = "Faturamento/Nota fiscal/Cancelar nota"
'        Case 77: Modulos = "Faturamento/Nota fiscal/Excluir duplicatas"
'        Case 78: Modulos = "Faturamento/Nota fiscal/SPED"
'        Case 79: Modulos = "Faturamento/Nota fiscal/Exportar"
'        Case 80: Modulos = "Faturamento/Carta de correção"
'        Case 81: Modulos = "Faturamento/Minuta de despacho"
'        Case 82: Modulos = "Faturamento/Relatórios/Histórico"
'        Case 83: Modulos = "Faturamento/Relatórios/Relacionamento de notas fiscais"
'        Case 84: Modulos = "Faturamento/Relatórios/Impostos"
'        Case 85: Modulos = "Faturamento/Relatórios/Doze últimos meses"
'        Case 86: Modulos = "Faturamento/Autorização de centro de custo sem previsão"
'
'        Case 87: Modulos = "Custos/Centro de custo"
'        Case 88: Modulos = "Custos/Centro de custo/Visualizar todos"
'        Case 89: Modulos = "Custos/Centro de custo/Visualizar lançamentos realizados"
'        Case 90: Modulos = "Custos/Relatórios/Detalhado"
'        Case 91: Modulos = "Custos/Relatórios/Resumido"
'        Case 92: Modulos = "Custos/Relatórios/Previsto x Realizado"
'
'        Case 93: Modulos = "Engenharia/Famílias"
'        Case 94: Modulos = "Engenharia/Produtos e serviços"
'            Case 95: Modulos = "Engenharia/Produtos e serviços/Validar estrutura"
'            Case 96: Modulos = "Engenharia/Produtos e serviços/Validar plano de inspeção"
'        Case 97: Modulos = "Engenharia/Conjuntos"
'        Case 98: Modulos = "Engenharia/Estrutura"
'            Case 99: Modulos = "Engenharia/Estrutura/Visualizar valor de custo"
'        Case 100: Modulos = "Engenharia/Controle de projetos"
'        Case 101: Modulos = "Engenharia/Processos"
'        Case 102: Modulos = "Engenharia/Processos/Histórico"
'        Case 103: Modulos = "Engenharia/Normas"
'
'        Case 104: Modulos = "PCP/Postos de trabalho"
'        Case 105: Modulos = "PCP/Códigos de trabalho"
'        Case 106: Modulos = "PCP/Carga de posto de trabalho"
'        Case 107: Modulos = "PCP/Gerenciamento de ordem"
'            Case 108: Modulos = "PCP/Gerenciamento de ordem/Validar resultados"
'        Case 109: Modulos = "PCP/Monitor de trabalho"
'        Case 110: Modulos = "PCP/Programas CNC"
'        Case 111: Modulos = "PCP/Situação da produção"
'        Case 112: Modulos = "PCP/Necessidade"
'        Case 113: Modulos = "PCP/Não conformidade"
'        Case 114: Modulos = "PCP/Programação da produção"
'        Case 115: Modulos = "PCP/Plano da produção"
'        Case 116: Modulos = "PCP/Requisição da ordem"
'        Case 117: Modulos = "PCP/Relatórios/Produtividade"
'        Case 118: Modulos = "PCP/Relatórios/Não conformidade"
'        Case 119: Modulos = "PCP/Relatórios/Monitor de eventos"
'        Case 120: Modulos = "PCP/Relatórios/Índice de atraso"
'        Case 121: Modulos = "PCP/Relatórios/Resultados da ordem"
'
'        Case 122: Modulos = "Qualidade/Famílias"
'        Case 123: Modulos = "Qualidade/Instrumentos"
'        Case 124: Modulos = "Qualidade/Almoxarifado"
'        Case 125: Modulos = "Qualidade/Plano de inspeção"
'        Case 126: Modulos = "Qualidade/Controle de medição"
'        Case 127: Modulos = "Qualidade/Inspeção de recebimento"
'        Case 128: Modulos = "Qualidade/Ensaios/Ultra-som"
'        Case 129: Modulos = "Qualidade/Ensaios/Líquido penetrante"
'        Case 130: Modulos = "Qualidade/Ensaios/Controle de certificados"
'        Case 131: Modulos = "Qualidade/Controle de certificados"
'        Case 132: Modulos = "Qualidade/Controle de documentos e dados"
'        Case 133: Modulos = "Qualidade/Não conformidade"
'            Case 134: Modulos = "Qualidade/Não conformidade/Descrição da não conformidade"
'            Case 135: Modulos = "Qualidade/Não conformidade"
'        Case 136: Modulos = "Qualidade/Solicitação de ação"
'        Case 137: Modulos = "Qualidade/Solicitação de desvio"
'        Case 138: Modulos = "Qualidade/RNC"
'        Case 139: Modulos = "Qualidade/PPAP"
'            Case 140: Modulos = "Qualidade/PPAP/PSW"
'            Case 141: Modulos = "Qualidade/PPAP/FMEA"
'            Case 142: Modulos = "Qualidade/PPAP/Plano de controle"
'        Case 143: Modulos = "Qualidade/Histórico de revisão dos relatórios"
'        Case 144: Modulos = "Qualidade/Relatórios/Não conformidade"
'        Case 145: Modulos = "Qualidade/Relatórios/Histórico"
'
'        Case 146: Modulos = "Estoque/Almoxarifado"
'        Case 147: Modulos = "Estoque/Local de armazenamento"
'        Case 148: Modulos = "Estoque/Requisição de materiais"
'        Case 149: Modulos = "Estoque/Recebimento/Pedido de compra"
'        Case 150: Modulos = "Estoque/Recebimento/Consignação"
'        Case 151: Modulos = "Estoque/Recebimento/Importar nota de terceiros"
'        Case 152: Modulos = "Estoque/Inventário"
'        Case 153: Modulos = "Estoque/Movimentação"
'        Case 154: Modulos = "Estoque/Controle de certificado"
'        Case 155: Modulos = "Estoque/Não conformidade"
'        Case 156: Modulos = "Estoque/Necessidade"
'        Case 157: Modulos = "Estoque/Ordem de faturamento"
'        Case 158: Modulos = "Estoque/Nota fiscal"
'        Case 159: Modulos = "Estoque/Autorização de centro de custo sem previsão"
'
'        Case 160: Modulos = "Manutenção/Equipamentos"
'            Case 161: Modulos = "Manutenção/Equipamentos/Aprovar manutenção"
'        Case 162: Modulos = "Manutenção/Assistência técnica"
'        Case 163: Modulos = "Manutenção/Relatórios/Histórico"
'
'        Case 164: Modulos = "Outros/Solicitação"
'            Case 165: Modulos = "Outros/Solicitação/Autorizar solicitação"
'        Case 166: Modulos = "Outros/Solicitação de produção"
'            Case 167: Modulos = "Outros/Solicitação de produção/Autorizar solicitação"
'        Case 168: Modulos = "Outros/Follow up de compras"
'        Case 169: Modulos = "Outros/Validação de procedimentos"
'        Case 170: Modulos = "Outros/Análise crítica"
'            Case 171: Modulos = "Outros/Análise crítica/Engenharia"
'            Case 172: Modulos = "Outros/Análise crítica/Processos"
'            Case 173: Modulos = "Outros/Análise crítica/Pcp"
'            Case 174: Modulos = "Outros/Análise crítica/Qualidade"
'            Case 175: Modulos = "Outros/Análise crítica/Compras"
'            Case 176: Modulos = "Outros/Análise crítica/Vendas"
'            Case 177: Modulos = "Outros/Análise crítica/Documentos"
'        Case 178: Modulos = "Outros/Downloads/Nota fiscal"
'        Case 179: Modulos = "Outros/Downloads/Boleto"
'
'        Case 180: Modulos = "Suporte/Chamado"
'        Case 181: Modulos = "Suporte/Chat (online)"
'        Case 182: Modulos = "Suporte/Solicitação de atendimento"
'            Case 183: Modulos = "Suporte/Atualização/Caprind e Gerprod"
'            Case 184: Modulos = "Suporte/Atualização/GNFe"
'            Case 185: Modulos = "Suporte/Atualização/GMRE (relatórios)"
'
'        Case 186: Modulos = "Avisos diário/Solicitação"
'        Case 187: Modulos = "Avisos diário/Contas a pagar"
'        Case 188: Modulos = "Avisos diário/Contas a receber"
'        Case 189: Modulos = "Avisos diário/Manutenção"
'        Case 190: Modulos = "Avisos diário/Requisição de materiais"
'        Case 191: Modulos = "Avisos diário/Compras/Necessidade"
'        Case 192: Modulos = "Avisos diário/PCP/Necessidade"
'        Case 193: Modulos = "Avisos diário/Estoque/Necessidade"
'        Case 194: Modulos = "Avisos diário/Faturamento/Carteira de faturamento"
'        Case 195: Modulos = "Avisos diário/PCP/OSs em atraso"
'        Case 196: Modulos = "Avisos diário/Custos/Centro de custo"
'        Case 197: Modulos = "Avisos diário/Análise crítica/Engenharia"
'        Case 198: Modulos = "Avisos diário/Análise crítica/Processos"
'        Case 199: Modulos = "Avisos diário/Análise crítica/Pcp"
'        Case 200: Modulos = "Avisos diário/Análise crítica/Qualidade"
'        Case 201: Modulos = "Avisos diário/Análise crítica/Compras"
'        Case 202: Modulos = "Avisos diário/Terceiros"
'        Case 203: Modulos = "Avisos diário/Compras/Pedidos em atraso"
'        Case 204: Modulos = "Avisos diário/Qualidade/Calibração a vencer"
'        Case 205: Modulos = "Avisos diário/Qualidade/Não conformidades"
'        Case 206: Modulos = "Avisos diário/Estoque/Produtos á vencer"
'
'    End Select
'    With Lista.ListItems
'        .Add , , ""
'        .Item(.Count).SubItems(1) = Modulos
'        Set TBAcessos = CreateObject("adodb.recordset")
'        TBAcessos.Open "Select * from Acessos where IDUsuario = " & frmUsuarios.txtID & " and Acesso = '" & Modulos & "'", Conexao, adOpenKeyset, adLockOptimistic
'        If TBAcessos.EOF = False Then
'            .Item(contador).Checked = True
'        End If
'        TBAcessos.Close
'    End With
'Loop


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

If ColumnHeader = "" Then
    With Lista
        For InitFor = 1 To .ListItems.Count
            If .ListItems.Item(InitFor).Checked = True Then .ListItems.Item(InitFor).Checked = False Else .ListItems.Item(InitFor).Checked = True
        Next InitFor
    End With
Else
    ProcOrdenaListView Lista, ColumnHeader
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcSalvar
    'Case 3: ProcAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
