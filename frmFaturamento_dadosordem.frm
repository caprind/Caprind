VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmFaturamento_dadosordem 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administrativo - Faturamento - Localizar ordem/lote no estoque"
   ClientHeight    =   7740
   ClientLeft      =   -15
   ClientTop       =   315
   ClientWidth     =   11745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFaturamento_dadosordem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   55
      TabIndex        =   3
      Top             =   6840
      Width           =   11655
      Begin VB.TextBox txtPagIr 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6390
         TabIndex        =   5
         ToolTipText     =   "Número da página."
         Top             =   180
         Width           =   555
      End
      Begin VB.TextBox txtNreg 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         TabIndex        =   4
         Text            =   "30"
         ToolTipText     =   "Número de registros por página."
         Top             =   180
         Width           =   555
      End
      Begin DrawSuite2022.USButton cmdPagProx 
         Height          =   315
         Left            =   8610
         TabIndex        =   6
         ToolTipText     =   "Próxima página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_dadosordem.frx":212A
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagAnt 
         Height          =   315
         Left            =   8070
         TabIndex        =   7
         ToolTipText     =   "Página anterior."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_dadosordem.frx":58CE
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagIr 
         Height          =   315
         Left            =   6960
         TabIndex        =   8
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   "Ir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagPrim 
         Height          =   315
         Left            =   7530
         TabIndex        =   9
         ToolTipText     =   "Primeira página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_dadosordem.frx":93D7
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin DrawSuite2022.USButton cmdPagUlt 
         Height          =   315
         Left            =   9150
         TabIndex        =   10
         ToolTipText     =   "Última página."
         Top             =   180
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   556
         DibPicture      =   "frmFaturamento_dadosordem.frx":D4C6
         BorderColor     =   14404026
         BorderColorDown =   11632444
         BorderColorOver =   11632444
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColor2  =   16777215
         GradientColor3  =   16777215
         GradientColorDown2=   16246986
         GradientColorDown3=   15189380
         GradientColorDown4=   14596208
         GradientColorOver1=   16643560
         GradientColorOver2=   16576988
         GradientColorOver3=   16441780
         GradientColorOver4=   16178091
         PicSizeH        =   19
         PicSizeW        =   19
      End
      Begin VB.Label lblRegistros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de registros: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label lblPaginas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Página: 0 de: 0"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9900
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carregar               registros por página"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   2760
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   55
      TabIndex        =   1
      Top             =   7455
      Width           =   11655
      _ExtentX        =   20558
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
      SearchText      =   ""
      Value           =   0
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   5850
      Left            =   55
      TabIndex        =   2
      Top             =   990
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10319
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Object.Tag             =   "D"
         Text            =   "Data"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Lote"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Local de armazenamento"
         Object.Width           =   6818
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Corrida"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Certificado"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Qtde. estoque"
         Object.Width           =   2293
      EndProperty
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   1720
      ButtonCount     =   3
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Ajuda"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Ajuda (F1)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   36
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonCaption2  =   "Sair"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Sair (Esc)"
      ButtonKey2      =   "2"
      ButtonAlignment2=   2
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   40
      ButtonTop2      =   2
      ButtonWidth2    =   26
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   5
      ButtonLeft3     =   68
      ButtonTop3      =   2
      ButtonWidth3    =   24
      ButtonHeight3   =   24
      ButtonUseMaskColor3=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   6240
         Top             =   210
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmFaturamento_dadosordem.frx":10D52
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmFaturamento_dadosordem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

lblRegistros.Caption = "Nº de registros: 0"
lblPaginas.Caption = "Página: 0 de: 0"
Lista.ListItems.Clear
If Formulario <> "Estoque/Ordem de faturamento" Then
With frmFaturamento_Prod_Serv
    Acao = "localizar ordem/lote"
    If Faturamento_ListaProdudos = True Then
        If .txtCod_Produto = "" Then
            NomeCampo = "o código interno"
            ProcVerificaAcao
            Exit Sub
        End If
        If .txtQTD = "" Then
            NomeCampo = "a quantidade"
            ProcVerificaAcao
            Exit Sub
        End If
        IDprodserv = .txtidproduto
        Desenho = .txtCod_Produto
        valor = .txtQTD
    Else
        If .txtcodServ = "" Then
            NomeCampo = "o código interno"
            ProcVerificaAcao
            Exit Sub
        End If
        If .txtqtdserv = "" Then
            NomeCampo = "a quantidade"
            ProcVerificaAcao
            Exit Sub
        End If
        IDprodserv = .txtidservico
        Desenho = .txtcodServ
        valor = .txtqtdserv
    End If
    If .txtUN <> "" And .Cmb_un_com <> "" Then valor = valor / FunVerificaTabelaConversaoUnidade(.txtUN, .Cmb_un_com)
    NovoValor = Replace(valor, ",", ".")
    
'    'Carrega empenhados se houver
'    'Empenhado para OF
'    Set TBEstoque = CreateObject("adodb.recordset")
'    TBEstoque.Open "Select ISNULL(SUM(ECEV.Qtde_empenhada), 0) as qt from (tbl_Detalhes_Nota_pedidos NFPP INNER JOIN Estoque_Controle_Empenho_Vendas ECEV ON ECEV.ID_carteira = NFPP.ID_carteira) INNER JOIN Estoque_controle EC ON EC.IDestoque = ECEV.ID_estoque where NFPP.ID_prod_NF = " & TBProduto!Int_codigo & " and EC.Desenho = '" & TBProduto!int_Cod_Produto & "' and ECEV.Qtde_empenhada - ISNULL(ECEV.Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
'    If TBEstoque.EOF = False Then
    
    
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select EC.*, (ISNULL(EE.Qtde_empenhada, 0) - ISNULL(EE.Qtde_saida, 0)) as Valor1 from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where EC.ID_empresa = " & .txtIDEmpresa & " and EC.Desenho = '" & Desenho & "' and (ISNULL(EE.Qtde_empenhada, 0) - ISNULL(EE.Qtde_saida, 0)) > 0 order by EC.Data, EC.LOTE, EC.local_armaz, EC.Corrida, EC.Certificado", Conexao, adOpenKeyset, adLockReadOnly
    If TBEstoque.EOF = False Then
        ProcExibePagina (1)
    Else
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select EC.*, EC.estoque_real - ((ISNULL(EE.Qtde_empenhada, 0) - ISNULL(EE.Qtde_saida, 0)) + (ISNULL(PNFC.Quantidade, 0) - ISNULL(PNFC.Qtde_saida, 0))) as Valor1 from (Estoque_controle EC LEFT JOIN Estoque_Controle_Empenho_Vendas EE ON EE.ID_estoque = EC.IDestoque) LEFT JOIN Producao_NF_Consignada PNFC ON PNFC.IDestoque = EC.IDestoque where EC.ID_empresa = " & .txtIDEmpresa & " and EC.Desenho = '" & Desenho & "' and EC.estoque_real - ((ISNULL(EE.Qtde_empenhada, 0) - ISNULL(EE.Qtde_saida, 0)) + (ISNULL(PNFC.Quantidade, 0) - ISNULL(PNFC.Qtde_saida, 0))) >= " & NovoValor & " order by EC.Data, EC.LOTE, EC.local_armaz, EC.Corrida, EC.Certificado", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then ProcExibePagina (1)
    End If
End With
Else

With frmEstoque_Ordem_Faturamento
    Acao = "localizar ordem/lote"
    If Faturamento_ListaProdudos = True Then
        If .txtCod_Produto = "" Then
            NomeCampo = "o código interno"
            ProcVerificaAcao
            Exit Sub
        End If
        If .txtQTD = "" Then
            NomeCampo = "a quantidade"
            ProcVerificaAcao
            Exit Sub
        End If
        IDprodserv = .txtidproduto
        Desenho = .txtCod_Produto
        valor = .txtQTD
    Else
        If .txtcodServ = "" Then
            NomeCampo = "o código interno"
            ProcVerificaAcao
            Exit Sub
        End If
        If .txtqtdserv = "" Then
            NomeCampo = "a quantidade"
            ProcVerificaAcao
            Exit Sub
        End If
        IDprodserv = .txtidservico
        Desenho = .txtcodServ
        valor = .txtqtdserv
    End If
    If .txtUN <> "" And .Cmb_un_com <> "" Then valor = valor / FunVerificaTabelaConversaoUnidade(.txtUN, .Cmb_un_com)
    NovoValor = Replace(valor, ",", ".")
    
'    'Carrega empenhados se houver
'    'Empenhado para OF
'    Set TBEstoque = CreateObject("adodb.recordset")
'    TBEstoque.Open "Select ISNULL(SUM(ECEV.Qtde_empenhada), 0) as qt from (tbl_Detalhes_Nota_pedidos NFPP INNER JOIN Estoque_Controle_Empenho_Vendas ECEV ON ECEV.ID_carteira = NFPP.ID_carteira) INNER JOIN Estoque_controle EC ON EC.IDestoque = ECEV.ID_estoque where NFPP.ID_prod_NF = " & TBProduto!Int_codigo & " and EC.Desenho = '" & TBProduto!int_Cod_Produto & "' and ECEV.Qtde_empenhada - ISNULL(ECEV.Qtde_saida, 0) > 0", Conexao, adOpenKeyset, adLockOptimistic
'    If TBEstoque.EOF = False Then
    
    
    Set TBEstoque = CreateObject("adodb.recordset")
    TBEstoque.Open "Select EC.*, (ISNULL(EE.Qtde_empenhada, 0) - ISNULL(EE.Qtde_saida, 0)) as Valor1 from (Estoque_Controle_Empenho_Vendas EE INNER JOIN estoque_controle EC ON EE.ID_estoque = EC.IDEstoque) INNER JOIN tbl_Detalhes_Nota_pedidos NFPP ON NFPP.ID_carteira = EE.ID_carteira and NFPP.Codinterno = EC.Desenho where EC.ID_empresa = " & .txtIDEmpresa & " and EC.Desenho = '" & Desenho & "' and (ISNULL(EE.Qtde_empenhada, 0) - ISNULL(EE.Qtde_saida, 0)) > 0 order by EC.Data, EC.LOTE, EC.local_armaz, EC.Corrida, EC.Certificado", Conexao, adOpenKeyset, adLockReadOnly
    If TBEstoque.EOF = False Then
        ProcExibePagina (1)
    Else
        Set TBEstoque = CreateObject("adodb.recordset")
        TBEstoque.Open "Select EC.*, EC.estoque_real - ((ISNULL(EE.Qtde_empenhada, 0) - ISNULL(EE.Qtde_saida, 0)) + (ISNULL(PNFC.Quantidade, 0) - ISNULL(PNFC.Qtde_saida, 0))) as Valor1 from (Estoque_controle EC LEFT JOIN Estoque_Controle_Empenho_Vendas EE ON EE.ID_estoque = EC.IDestoque) LEFT JOIN Producao_NF_Consignada PNFC ON PNFC.IDestoque = EC.IDestoque where EC.ID_empresa = " & .txtIDEmpresa & " and EC.Desenho = '" & Desenho & "' and EC.estoque_real - ((ISNULL(EE.Qtde_empenhada, 0) - ISNULL(EE.Qtde_saida, 0)) + (ISNULL(PNFC.Quantidade, 0) - ISNULL(PNFC.Qtde_saida, 0))) >= " & NovoValor & " order by EC.Data, EC.LOTE, EC.local_armaz, EC.Corrida, EC.Certificado", Conexao, adOpenKeyset, adLockOptimistic
        If TBEstoque.EOF = False Then ProcExibePagina (1)
    End If
End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcExibePagina(Pagina)
On Error GoTo tratar_erro

Lista.ListItems.Clear
TBEstoque.PageSize = IIf(txtNreg = "", 30, txtNreg)
TBEstoque.AbsolutePage = Pagina
TamanhoPagina = TBEstoque.PageSize
ContadorReg = 1

PBLista.Min = 0
PBLista.Max = FunVerifMaxPBListaPaginacao(TBEstoque.RecordCount - IIf(Pagina > 1, (TBEstoque.PageSize * (Pagina - 1)), 0), TBEstoque.PageSize)
PBLista.Value = 1
contador = 0
Do While TBEstoque.EOF = False And (ContadorReg <= TamanhoPagina)
    With Lista.ListItems
        .Add , , TBEstoque!IDEstoque
        .Item(.Count).SubItems(1) = IIf(IsNull(TBEstoque!Data), "", Format(TBEstoque!Data, "dd/mm/yy"))
        .Item(.Count).SubItems(2) = IIf(IsNull(TBEstoque!LOTE), "", TBEstoque!LOTE)
        .Item(.Count).SubItems(3) = IIf(IsNull(TBEstoque!local_armaz), "", TBEstoque!local_armaz)
        .Item(.Count).SubItems(4) = IIf(IsNull(TBEstoque!Corrida), "", TBEstoque!Corrida)
        .Item(.Count).SubItems(5) = IIf(IsNull(TBEstoque!Certificado), "", TBEstoque!Certificado)
        .Item(.Count).SubItems(6) = IIf(IsNull(TBEstoque!Valor1), "", Format(TBEstoque!Valor1, "###,##0.0000"))
    End With
    TBEstoque.MoveNext
    ContadorReg = ContadorReg + 1
    contador = contador + 1
    PBLista.Value = contador
Loop
lblRegistros.Caption = "Nº de registros: " & TBEstoque.RecordCount
If TBEstoque.AbsolutePage = adPosBOF Then
   lblPaginas.Caption = "Página: 1 de: " & TBEstoque.PageCount
ElseIf TBEstoque.AbsolutePage = adPosEOF Then
        lblPaginas.Caption = "Página: " & TBEstoque.PageCount & " de: " & TBEstoque.PageCount
    Else
        lblPaginas.Caption = "Página: " & TBEstoque.AbsolutePage - 1 & " de: " & TBEstoque.PageCount
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagAnt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBEstoque.AbsolutePage <> 2 Then
    If TBEstoque.AbsolutePage = -3 Then
        ProcExibePagina (TBEstoque.PageCount - 1)
    Else
        TBEstoque.AbsolutePage = TBEstoque.AbsolutePage - 2
        ProcExibePagina (TBEstoque.AbsolutePage)
    End If
Else
    ProcExibePagina (1)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagIr_Click()
On Error GoTo tratar_erro

If txtPagIr = "" Then Exit Sub
Quant = ReturnNumbersOnly(Right(lblPaginas.Caption, 4))
If Quant <= 1 Or txtPagIr > Quant Then Exit Sub
If txtPagIr.Text >= 1 And txtPagIr.Text <= Quant Then
    TBEstoque.AbsolutePage = txtPagIr.Text
    ProcExibePagina (TBEstoque.AbsolutePage)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagPrim_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBEstoque.AbsolutePage = 1
ProcExibePagina (TBEstoque.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagProx_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
If TBEstoque.AbsolutePage <> -3 Then
    If TBEstoque.AbsolutePage = 1 Then
        ProcExibePagina (2)
    Else
        ProcExibePagina (TBEstoque.AbsolutePage)
    End If
Else
    ProcExibePagina (TBEstoque.PageCount)
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdPagUlt_Click()
On Error GoTo tratar_erro

If ReturnNumbersOnly(Right(lblPaginas.Caption, 4)) <= 1 Then Exit Sub
TBEstoque.AbsolutePage = TBEstoque.PageCount
ProcExibePagina (TBEstoque.AbsolutePage)

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyReturn: Lista_DblClick
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 11655, 3, True

If Formulario = "Faturamento/Nota fiscal/Própria" Then
    Caption = "Administrativo - Faturamento - Nota fiscal - Própria - Localizar ordem/lote no estoque"
ElseIf Formulario = "Faturamento/Nota fiscal/Terceiros" Then
        Caption = "Administrativo - Faturamento - Nota fiscal - Terceiros - Localizar ordem/lote no estoque"
    ElseIf Formulario = "Estoque/Ordem de faturamento" Then
            Caption = "Estoque - Ordem de faturamento - Localizar ordem/lote no estoque"
        Else
            Caption = "Estoque - Nota fiscal - Localizar ordem/lote no estoque"
End If

ProcCarregaLista

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

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count = 0 Then Exit Sub

If Formulario <> "Estoque/Ordem de faturamento" Then
    With frmFaturamento_Prod_Serv
        If Faturamento_ListaProdudos = True Then
            .Txt_IDestoque = Lista.SelectedItem
            .txtof = Lista.SelectedItem.ListSubItems(2)
        Else
            .txtof_servico = Lista.SelectedItem.ListSubItems(2)
        End If
    End With

Else
    With frmEstoque_Ordem_Faturamento
        If Faturamento_ListaProdudos = True Then
            .Txt_IDestoque = Lista.SelectedItem
            .txtof = Lista.SelectedItem.ListSubItems(2)
        Else
            .txtof_servico = Lista.SelectedItem.ListSubItems(2)
        End If
    End With
End If
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNreg_Change()
On Error GoTo tratar_erro

If txtNreg <> "" Then
    VerifNumero = txtNreg
    ProcVerificaNumero
    If VerifNumero = False Then
        txtNreg = ""
        txtNreg.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtPagIr_Change()
On Error GoTo tratar_erro

If txtPagIr <> "" Then
    VerifNumero = txtPagIr
    ProcVerificaNumero
    If VerifNumero = False Then
        txtPagIr = ""
        txtPagIr.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    'Case 1: ProcAjuda
    Case 2: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
