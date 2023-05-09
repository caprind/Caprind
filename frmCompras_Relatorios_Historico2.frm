VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCompras_Relatorios_Historico2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Administrativo - Compras - Relatórios - Histórico"
   ClientHeight    =   10035
   ClientLeft      =   10350
   ClientTop       =   450
   ClientWidth     =   15360
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
   ForeColor       =   &H8000000D&
   Icon            =   "frmCompras_Relatorios_Historico2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   13590
      TabIndex        =   56
      Top             =   960
      Width           =   1755
      Begin VB.TextBox Txt_limite 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         ToolTipText     =   "Limite de registros para carregar na lista."
         Top             =   990
         Width           =   555
      End
      Begin VB.OptionButton Opt_valor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valor"
         Enabled         =   0   'False
         Height          =   195
         Left            =   540
         TabIndex        =   58
         Top             =   330
         Width           =   675
      End
      Begin VB.OptionButton Opt_quantidade 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quant."
         Enabled         =   0   'False
         Height          =   195
         Left            =   540
         TabIndex        =   57
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limitar em  "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   1080
         Width           =   810
      End
   End
   Begin MSComctlLib.ListView ListaDetalhada 
      Height          =   6435
      Left            =   0
      TabIndex        =   15
      Top             =   2430
      Width           =   15345
      _ExtentX        =   27067
      _ExtentY        =   11351
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   19
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "N"
         Text            =   "Pedido"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Fornecedor"
         Object.Width           =   7849
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "CNPJ | CPF"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Cód. de ref."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Object.Tag             =   "T"
         Text            =   "Família"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Qtde."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Valor unitário"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Valor desconto"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Valor IPI"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Valor ICMS ST"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Valor total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Object.Tag             =   "D"
         Text            =   "Moeda"
         Object.Width           =   1413
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Text            =   "Valor moeda"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   17
         Object.Tag             =   "T"
         Text            =   "Status"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Object.Tag             =   "T"
         Text            =   "Posto de trabalho"
         Object.Width           =   2646
      EndProperty
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10500
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15360
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
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
      Height          =   825
      Left            =   60
      TabIndex        =   33
      Top             =   9180
      Width           =   15195
      Begin VB.TextBox txtICMS_ST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9795
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do ICMS substituto."
         Top             =   375
         Width           =   1395
      End
      Begin VB.TextBox txtSubtotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6333
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Subtotal."
         Top             =   375
         Width           =   1395
      End
      Begin VB.TextBox txtDesconto_percentual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5010
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Percentual do desconto."
         Top             =   375
         Width           =   1005
      End
      Begin VB.TextBox txtDesconto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3642
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do(s) serviço(s)."
         Top             =   375
         Width           =   1335
      End
      Begin VB.TextBox txtTotal_geral 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   11526
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Valor total comprado."
         Top             =   375
         Width           =   1395
      End
      Begin VB.TextBox txtTotal_ipi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8064
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do IPI."
         Top             =   375
         Width           =   1395
      End
      Begin VB.TextBox txtTotal_produtos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   83
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do(s) produto(s)."
         Top             =   375
         Width           =   1395
      End
      Begin VB.TextBox Txt_total_servicos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1911
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Valor total do(s) serviço(s)."
         Top             =   375
         Width           =   1395
      End
      Begin VB.TextBox Txt_qtde_total_vendido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   13260
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade total comprado."
         Top             =   375
         Width           =   1755
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total ICMS ST"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   9990
         TabIndex        =   52
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6730
         TabIndex        =   51
         Top             =   180
         Width           =   600
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   195
         Left            =   9570
         TabIndex        =   50
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   195
         Left            =   6120
         TabIndex        =   49
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentual"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5130
         TabIndex        =   48
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total desconto"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3777
         TabIndex        =   47
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Left            =   3450
         TabIndex        =   46
         Top             =   435
         Width           =   60
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   195
         Left            =   1680
         TabIndex        =   41
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         Height          =   195
         Left            =   11295
         TabIndex        =   40
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total comprado"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   11668
         TabIndex        =   39
         Top             =   180
         Width           =   1110
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total produtos"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   255
         TabIndex        =   38
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total IPI"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8454
         TabIndex        =   37
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         Height          =   195
         Left            =   7830
         TabIndex        =   36
         Top             =   435
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total serviços"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2113
         TabIndex        =   35
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtd. total comprado"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   13410
         TabIndex        =   34
         Top             =   180
         Width           =   1455
      End
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   0
      TabIndex        =   43
      Top             =   8910
      Width           =   11475
      _ExtentX        =   20241
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
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   13920
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmCompras_Relatorios_Historico2.frx":0442
      Count           =   1
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relatório"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   30
      TabIndex        =   25
      Top             =   960
      Width           =   1365
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   570
         Width           =   1245
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   30
      TabIndex        =   45
      Top             =   1710
      Width           =   1365
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   420
         Width           =   1155
      End
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opções de filtro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1410
      TabIndex        =   26
      Top             =   960
      Width           =   10095
      Begin VB.Frame frameTexto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2910
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   7005
         Begin VB.TextBox txtTexto 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   30
            TabIndex        =   63
            ToolTipText     =   "Texto para pesquisa."
            Top             =   210
            Width           =   6975
         End
         Begin VB.OptionButton optIgual 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Igual"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   4110
            TabIndex        =   67
            Top             =   -30
            Width           =   705
         End
         Begin VB.OptionButton Optfim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fim"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   3510
            TabIndex        =   66
            Top             =   -30
            Width           =   555
         End
         Begin VB.OptionButton Optinicio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Início"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2100
            TabIndex        =   65
            Top             =   -30
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Optmeio 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Meio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   2820
            TabIndex        =   64
            Top             =   -30
            Width           =   675
         End
      End
      Begin VB.CheckBox chkTexto 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Digitar texto para pesquisa"
         Height          =   195
         Left            =   2340
         TabIndex        =   61
         Top             =   30
         Width           =   2265
      End
      Begin VB.ComboBox cmbFornecedor 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":3236
         Left            =   180
         List            =   "frmCompras_Relatorios_Historico2.frx":3238
         MouseIcon       =   "frmCompras_Relatorios_Historico2.frx":323A
         Sorted          =   -1  'True
         TabIndex        =   53
         ToolTipText     =   "Texto para pesquisa."
         Top             =   990
         Width           =   9765
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":3544
         Left            =   180
         List            =   "frmCompras_Relatorios_Historico2.frx":3546
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Opções para filtro."
         Top             =   450
         Width           =   2745
      End
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":3548
         Left            =   2940
         List            =   "frmCompras_Relatorios_Historico2.frx":354A
         MouseIcon       =   "frmCompras_Relatorios_Historico2.frx":354C
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Texto para pesquisa."
         Top             =   450
         Width           =   6975
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar histórico de compras  do Fornecedor"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3525
         TabIndex        =   54
         Top             =   810
         Width           =   3075
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         TabIndex        =   28
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5692
         TabIndex        =   27
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   11520
      TabIndex        =   29
      Top             =   960
      Width           =   2055
      Begin VB.ComboBox cmbPor 
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
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":3856
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico2.frx":385D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Por."
         Top             =   270
         Width           =   1305
      End
      Begin VB.ComboBox Cmb_ano_de 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":3866
         Left            =   1260
         List            =   "frmCompras_Relatorios_Historico2.frx":3868
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Ano de."
         Top             =   660
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox Cmb_ano_ate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":386A
         Left            =   1260
         List            =   "frmCompras_Relatorios_Historico2.frx":386C
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Ano até."
         Top             =   1020
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.ComboBox Cmb_ano_de1 
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
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":386E
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico2.frx":3870
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Ano de."
         Top             =   660
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ComboBox Cmb_ano_ate1 
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
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":3872
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico2.frx":3874
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Ano até."
         Top             =   1020
         Visible         =   0   'False
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   630
         TabIndex        =   7
         ToolTipText     =   "Data inicio."
         Top             =   660
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   199360513
         CurrentDate     =   39799
      End
      Begin VB.ComboBox Cmb_mes_de 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":3876
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico2.frx":389E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         ToolTipText     =   "Mês de."
         Top             =   660
         Visible         =   0   'False
         Width           =   645
      End
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   630
         TabIndex        =   11
         ToolTipText     =   "Data final."
         Top             =   1020
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   255
         Format          =   199360513
         CurrentDate     =   39799
      End
      Begin VB.ComboBox Cmb_mes_ate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         ItemData        =   "frmCompras_Relatorios_Historico2.frx":38DF
         Left            =   630
         List            =   "frmCompras_Relatorios_Historico2.frx":3907
         Style           =   2  'Dropdown List
         TabIndex        =   13
         ToolTipText     =   "Mês até."
         Top             =   1020
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Por :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   32
         Top             =   345
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   690
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   1050
         Width           =   360
      End
   End
   Begin MSComctlLib.ListView Lista1 
      Height          =   6375
      Left            =   30
      TabIndex        =   55
      Top             =   2430
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   11245
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   30
      TabIndex        =   44
      Top             =   0
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   1720
      ButtonCount     =   6
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Filtrar"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Filtrar (F2)"
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
      ButtonCaption2  =   "Relatório"
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonToolTipText2=   "Relatório (F5)"
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
      ButtonWidth2    =   51
      ButtonHeight2   =   21
      ButtonUseMaskColor2=   0   'False
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonAlignment3=   2
      ButtonType3     =   1
      ButtonStyle3    =   -1
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState3    =   -1
      ButtonLeft3     =   93
      ButtonTop3      =   4
      ButtonWidth3    =   2
      ButtonHeight3   =   54
      ButtonCaption4  =   "Ajuda"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Ajuda (F1)"
      ButtonKey4      =   "4"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   97
      ButtonTop4      =   2
      ButtonWidth4    =   36
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonCaption5  =   "Sair"
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonToolTipText5=   "Sair (Esc)"
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   135
      ButtonTop5      =   2
      ButtonWidth5    =   26
      ButtonHeight5   =   21
      ButtonUseMaskColor5=   0   'False
      ButtonEnabled6  =   0   'False
      ButtonIconSize6 =   32
      ButtonKey6      =   "6"
      ButtonAlignment6=   2
      BeginProperty ButtonFont6 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState6    =   5
      ButtonLeft6     =   163
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
      ButtonUseMaskColor6=   0   'False
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   12000
      TabIndex        =   42
      Top             =   8940
      Width           =   2895
   End
End
Attribute VB_Name = "frmCompras_Relatorios_Historico2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaCombo_Fornecedor()
On Error GoTo tratar_erro
cmbFornecedor.Clear

Select Case cmbPor.Text
        Case "Período":
                FiltroData = FiltroData & " (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            
        Case "Mês":
                qt = FunVerificaMes(Cmb_mes_de)
                Qtd = FunVerificaMes(Cmb_mes_ate)
        
                FiltroData = FiltroData & " Mes >= '" & qt & "' AND Mes <= '" & Qtd & "' And Year(Data) >= '" & Cmb_ano_de & "' And Ano <= '" & Cmb_ano_ate & "'"
            
        Case "Ano":
                If Cmb_ano_de1.Text = Cmb_ano_ate1 Then
                FiltroData = FiltroData & " Ano = '" & Cmb_ano_de1 & "'"
                Else
                FiltroData = FiltroData & " Ano >= '" & Cmb_ano_de1 & "' And Ano <= '" & Cmb_ano_ate1 & "'"
                End If
End Select


If cmbfiltrarpor.Text <> " " And FiltroData <> "" Then
StrSql = "select distinct fornecedor from Compras_relatorios_historico_detalhado where " & FiltroData & " Order By fornecedor"
'Debug.print StrSql

'StrSql = "select distinct fornecedor from Compras_relatorios_historico_detalhado where fornecedor is not NULL order by fornecedor"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
If TBAbrir.EOF = False Then
 cmbFornecedor.AddItem ""
    Do While TBAbrir.EOF = False
        vFornecedor = LTrim(TBAbrir!Fornecedor)
        cmbFornecedor.AddItem vFornecedor
        TBAbrir.MoveNext
    Loop
End If
End If

cmbFornecedor.Enabled = True

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcAjuda()
On Error GoTo tratar_erro

FunAbrirVideoWeb ("http://www.youtube.com/watch?v=UNB5MhQdTA0&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=34&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub chkTexto_Click()
On Error GoTo tratar_erro

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
frameTexto.Visible = False
'USMsgBox "Não utilizar texto para esse tipo de pesquisa", vbCritical, "CAPRIND v5.0"
chkTexto.Value = 0
Exit Sub
End If



frameTexto.Visible = chkTexto.Value

If frameTexto.Visible = True Then
    cmbFornecedor.Clear
    cmbFornecedor.Enabled = False
    cmbTexto.Clear
    cmbTexto.Enabled = False
    txtTexto.SetFocus
Else
    cmbTexto.Enabled = True
    cmbFornecedor.Enabled = True
    cmbfiltrarpor_Click
End If

txtTexto.Text = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub Cmb_ano_ate1_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Cmb_ano_de1 <> "" And Cmb_ano_ate1 <> "" Then
    qt = Cmb_ano_ate1
    Qtd = Cmb_ano_de1
    If qt < Qtd Then
        USMsgBox ("O ano final não pode ser menor que o ano inicial."), vbExclamation, "CAPRIND v5.0"
        Cmb_ano_ate1 = Cmb_ano_de1
    End If
End If

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
ProcCarregaCombo_Fornecedor
Else
ProcCarregaCombo_Texto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_ano_de_Click()
On Error GoTo tratar_erro

Cmb_ano_ate = Cmb_ano_de
ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
ProcCarregaCombo_Fornecedor
Else
ProcCarregaCombo_Texto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub Cmb_ano_de1_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Cmb_ano_de1 <> "" And Cmb_ano_ate1 <> "" Then
    qt = Cmb_ano_de1
    Qtd = Cmb_ano_ate1
    If qt > Qtd Then
        USMsgBox ("O ano inicial não pode ser maior que o ano final."), vbExclamation, "CAPRIND v5.0"
        Cmb_ano_de1 = Cmb_ano_ate1
    End If
End If

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
ProcCarregaCombo_Fornecedor
Else
ProcCarregaCombo_Texto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub



Private Sub Cmb_mes_ate_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Cmb_mes_de <> "" And Cmb_mes_ate <> "" Then
    qt = FunVerificaMes(Cmb_mes_de)
    Qtd = FunVerificaMes(Cmb_mes_ate)
    If Qtd < qt Then
        USMsgBox ("O mês final não pode ser menor que o mês inicial."), vbExclamation, "CAPRIND v5.0"
        Cmb_mes_ate = Cmb_mes_de
    End If
End If

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
ProcCarregaCombo_Fornecedor
Else
ProcCarregaCombo_Texto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_mes_de_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Cmb_mes_de <> "" And Cmb_mes_ate <> "" Then
    qt = FunVerificaMes(Cmb_mes_de)
    Qtd = FunVerificaMes(Cmb_mes_ate)
    If qt > Qtd Then
        USMsgBox ("O mês inicial não pode ser maior que o mês final."), vbExclamation, "CAPRIND v5.0"
        Cmb_mes_de = Cmb_mes_ate
    End If
End If

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
ProcCarregaCombo_Fornecedor
Else
ProcCarregaCombo_Texto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFornecedor_Change()
On Error GoTo tratar_erro
Dim CampoFiltro As String

Select Case cmbfiltrarpor.Text
        Case "Código interno x Fornecedor":  CampoFiltro = "Desenho"
        Case "Código de referência x Fornecedor": CampoFiltro = "n_referencia"
End Select

If cmbfiltrarpor.Text <> " " And CampoFiltro <> "" Then
StrSql = "select Distinct " & CampoFiltro & " as campoFiltro from Compras_relatorios_historico_detalhado where fornecedor = '" & cmbFornecedor & "' and " & CampoFiltro & " Is not Null ORDER BY " & CampoFiltro & ""
        
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
            If TBAbrir.EOF = False Then
            cmbTexto.Clear
            With cmbTexto
                Do While TBAbrir.EOF = False
                         If IsNull(TBAbrir!CampoFiltro) = False Then
                            .AddItem TBAbrir!CampoFiltro
                        End If
                
                    TBAbrir.MoveNext
                Loop
            End With
            End If

End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbFornecedor_Click()
On Error GoTo tratar_erro

cmbTexto.Clear
ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
ProcMostrarEsconderCombosData

ProcCarregaCombo_Texto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbPor_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
ProcMostrarEsconderCombosData

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
chkTexto.Value = False
ProcCarregaCombo_Fornecedor
Else
ProcCarregaCombo_Texto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'Private Sub ProcImprimir()
'On Error GoTo tratar_erro
'
'If Opt_individual.Value = True Then
'    If optDetalhado.Value = True Then
'        If ListaDetalhada.ListItems.Count = 0 Then Exit Sub
'    Else
'        If Lista1.ListItems.Count = 0 Then Exit Sub
'    End If
'Else
'    If Lista1.ListItems.Count = 0 Then Exit Sub
'End If
'frmCompras_Relatorios_Historico_MenuImpressao.Show 1
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal
Unload Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcFiltrar
 '   Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaDetalhada()
On Error GoTo tratar_erro
Dim Total_produtos As Double
Dim Total_Servicos As Double
Dim Total_IPI As Double
Dim Total_ICMS As Double
Dim Total_Desconto As Double
Dim Total_Comprado As Double
Dim Total_PDesconto As Double
Dim Total_Sem_Desconto As Double

Set TBListaDetalhada = CreateObject("adodb.recordset")
TBListaDetalhada.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
Contador1 = 1
ListaDetalhada.ListItems.Clear

If TBListaDetalhada.EOF = False Then
    TBListaDetalhada.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBListaDetalhada.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBListaDetalhada.MoveFirst
    Do While TBListaDetalhada.EOF = False
            With ListaDetalhada.ListItems
                    .Add , , TBListaDetalhada!IDlista
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBListaDetalhada!Data), "", Format(TBListaDetalhada!Data, "dd/mm/yy"))
                    .Item(.Count).SubItems(2) = IIf(IsNull(TBListaDetalhada!Pedido), "", TBListaDetalhada!Pedido)
                    .Item(.Count).SubItems(3) = IIf(IsNull(TBListaDetalhada!Fornecedor), "", TBListaDetalhada!Fornecedor)
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBListaDetalhada!CPF_CNPJ), "", TBListaDetalhada!CPF_CNPJ)
                    
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBListaDetalhada!Desenho), "", TBListaDetalhada!Desenho)
                    .Item(.Count).SubItems(6) = IIf(IsNull(TBListaDetalhada!N_referencia), "", TBListaDetalhada!N_referencia)
                    .Item(.Count).SubItems(7) = IIf(IsNull(TBListaDetalhada!Descricao), "", TBListaDetalhada!Descricao)
                    .Item(.Count).SubItems(8) = IIf(IsNull(TBListaDetalhada!Familia), "", TBListaDetalhada!Familia)
                    .Item(.Count).SubItems(9) = IIf(IsNull(TBListaDetalhada!Quant_Comp), "", Format(TBListaDetalhada!Quant_Comp, "###,##0.0000"))
                    .Item(.Count).SubItems(10) = IIf(IsNull(TBListaDetalhada!preco_unitario_desconto), "", (Format(TBListaDetalhada!preco_unitario_desconto, "###,##0.0000")))
                    .Item(.Count).SubItems(11) = IIf(IsNull(TBListaDetalhada!ValorDesconto), "", (Format(TBListaDetalhada!ValorDesconto, "###,##0.0000")))
                    .Item(.Count).SubItems(12) = IIf(IsNull(TBListaDetalhada!VlrIPI), "", (Format(TBListaDetalhada!VlrIPI, "###,##0.00")))
                    .Item(.Count).SubItems(13) = IIf(IsNull(TBListaDetalhada!Valor_ICMS_ST), "", (Format(TBListaDetalhada!Valor_ICMS_ST, "###,##0.00")))
                    .Item(.Count).SubItems(14) = Format(IIf(IsNull(TBListaDetalhada!preco_total), 0, TBListaDetalhada!preco_total) + IIf(IsNull(TBListaDetalhada!VlrIPI), 0, TBListaDetalhada!VlrIPI), "###,##0.00")
                    .Item(.Count).SubItems(15) = IIf(IsNull(TBListaDetalhada!Moeda), "", TBListaDetalhada!Moeda)
                    .Item(.Count).SubItems(16) = IIf(IsNull(TBListaDetalhada!Valor_moeda), "", Format(TBListaDetalhada!Valor_moeda, "###,##0.00"))
                    If IsNull(TBListaDetalhada!Status_Item) = False And TBListaDetalhada!Status_Item <> "" Then
                        If TBListaDetalhada!Status_Item = "N_RECEBIDO" Then .Item(.Count).SubItems(17) = "COMPRADO" Else .Item(.Count).SubItems(17) = TBListaDetalhada!Status_Item
                    End If
                    .Item(.Count).SubItems(18) = IIf(IsNull(TBListaDetalhada!maquina), "", TBListaDetalhada!maquina)
            End With
        
        If TBListaDetalhada!Tipo = "P" Then
         Total_produtos = Total_produtos + TBListaDetalhada!preco_total
        Else
         Total_Servicos = Total_Servicos + TBListaDetalhada!preco_total
        End If
        Total_Sem_Desconto = Total_Sem_Desconto + TBListaDetalhada!preco_total
              
        Total_Desconto = Total_Desconto + TBListaDetalhada!ValorDesconto
        Total_IPI = Total_IPI + TBListaDetalhada!VlrIPI
        Total_ICMS = Total_ICMS + TBListaDetalhada!Valor_ICMS_ST
        Total_Comprado = Total_Comprado + TBListaDetalhada!Quant_Comp
        
        TBListaDetalhada.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
 '          frmCompras_Relatorios_Historico.Refresh
    Loop
End If
TBListaDetalhada.Close

If Total_Desconto And Total_Sem_Desconto <> 0 Then
Total_PDesconto = (Total_Desconto / Total_Sem_Desconto) * 100
End If

txtTotal_produtos = Format(Total_produtos, "###,##0.0000")
Txt_total_servicos = Format(Total_Servicos, "###,##0.0000")
txtDesconto = Format(Total_Desconto, "###,##0.0000")
txtICMS_ST = Format(Total_ICMS, "###,##0.0000")
txtTotal_ipi = Format(Total_IPI, "###,##0.0000")

txtSubtotal = Format(Total_produtos + Total_Servicos - Total_Desconto, "###,##0.00")
txtDesconto_percentual = Format(Total_PDesconto, "###,##0.000") & "%"
txtTotal_geral = Format((Total_produtos + Total_Servicos + Total_IPI + Total_ICMS) - Total_Desconto, "###,##0.00")
Txt_qtde_total_vendido = Format(Total_Comprado, "###,##0.00")


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
txtTotal_produtos = ""
Txt_total_servicos = ""
txtDesconto = ""
txtDesconto_percentual = ""
txtSubtotal = ""
txtTotal_ipi = ""
txtICMS_ST = ""
txtTotal_geral = ""
Txt_qtde_total_vendido = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 15195, 5, True
Formulario = "Compras/Relatórios/Histórico"
Direitos
ProcLimpaVariaveisPrincipais
msk_fltInicio.Value = Date
msk_fltFim.Value = Date
ProcCarregaComboAno Cmb_ano_ate, "2005", 1
ProcCarregaComboAno Cmb_ano_ate1, "2005", 1
ProcCarregaComboAno Cmb_ano_de, "2005", 1
ProcCarregaComboAno Cmb_ano_de1, "2005", 1

    With cmbfiltrarpor
        .Clear
        .AddItem " "
        .AddItem "Pedido"
        .AddItem "Status"
        .AddItem "Fornecedor"
        .AddItem "Código interno"
        .AddItem "Descrição"
        .AddItem "Família"
        .AddItem "Grupo"
        .AddItem "Posto de trabalho"
        .AddItem "Centro de custo"
        .AddItem "Código de referência"
        .AddItem "Detalhe"
        .AddItem "Código interno x Fornecedor"
        .AddItem "Código de referência x Fornecedor"
        .AddItem "Família x Grupo"
        .Text = " "
    End With


    With cmbPor
        .Clear
        .AddItem "Período"
        .AddItem "Mês"
        .AddItem "Ano"
        .Text = "Período"
    End With
    ProcMostrarEsconderCombosData

cmbPor = "Período"

ProcRemoveObjetosResize Me

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Formulario = "Compras/Relatórios/Histórico"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaDetalhada_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaDetalhada, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaResumida_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView Lista1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro
ListaDetalhada.ListItems.Clear
cmbTexto.Enabled = True
        
If cmbfiltrarpor.Text <> "" And chkTexto.Value = 0 Then

    If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
        ProcCarregaCombo_Fornecedor
    Else
        cmbFornecedor.Enabled = False
        ProcCarregaCombo_Texto
    End If

    If optDetalhado.Value = True And cmbfiltrarpor = "Posto de trabalho" Then
        ListaDetalhada.ColumnHeaders(14).Width = 1500
    Else
        ListaDetalhada.ColumnHeaders(14).Width = 0
    End If
    
Else
    txtTexto.SetFocus
    txtTexto.Text = ""
    cmbFornecedor.Clear
    cmbTexto.Clear
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaCombo_Texto()
On Error GoTo tratar_erro
Dim CampoFiltro, FiltroData As String
cmbTexto.Clear

Select Case cmbfiltrarpor.Text
        Case "Status":  CampoFiltro = "Status_Item"
        Case "Pedido":  CampoFiltro = "Pedido"
        Case "Código interno":  CampoFiltro = "Desenho"
        Case "Fornecedor": CampoFiltro = "Fornecedor"
        Case "Código de referência": CampoFiltro = "n_referencia"
        Case "Descrição": CampoFiltro = "Descricao"
        Case "Família": CampoFiltro = "Familia"
        Case "Grupo":  CampoFiltro = "Grupo"
        Case "Posto de trabalho": CampoFiltro = "maquina"
        Case "Centro de custo": CampoFiltro = "Setor"
        Case "Detalhe": CampoFiltro = "Detalhe"
        Case "Código interno x Fornecedor":
        CampoFiltro = "Desenho"
        FiltroData = " Fornecedor = '" & cmbFornecedor.Text & "' and  "
        Case "Código de referência x Fornecedor": CampoFiltro = "n_referencia"
        Case "Família x Grupo":  CampoFiltro = "Grupo"
End Select

Select Case cmbPor.Text
        Case "Período":
                FiltroData = FiltroData & " (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            
        Case "Mês":
                qt = FunVerificaMes(Cmb_mes_de)
                Qtd = FunVerificaMes(Cmb_mes_ate)
        
                FiltroData = FiltroData & " Mes >= '" & qt & "' AND Mes <= '" & Qtd & "' And Year(Data) >= '" & Cmb_ano_de & "' And Ano <= '" & Cmb_ano_ate & "'"
            
        Case "Ano":
                If Cmb_ano_de1.Text = Cmb_ano_ate1 Then
                FiltroData = FiltroData & " Ano = '" & Cmb_ano_de1 & "'"
                Else
                FiltroData = FiltroData & " Ano >= '" & Cmb_ano_de1 & "' And Ano <= '" & Cmb_ano_ate1 & "'"
                End If
End Select


If cmbfiltrarpor.Text <> " " And CampoFiltro <> "" And FiltroData <> "" Then
StrSql = "select Distinct " & CampoFiltro & " as campoFiltro from Compras_relatorios_historico_detalhado where " & FiltroData & " Order By " & CampoFiltro
'Debug.print StrSql

            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
            If TBAbrir.EOF = False Then

            With cmbTexto
                Do While TBAbrir.EOF = False
                         If IsNull(TBAbrir!CampoFiltro) = False Then
                            .AddItem IIf(TBAbrir!CampoFiltro = "N_RECEBIDO", "COMPRADO", TBAbrir!CampoFiltro)
                        End If
                
                    TBAbrir.MoveNext
                Loop
            End With
            End If

End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

'
'Private Sub ProcCarregaComboTexto()
'On Error GoTo tratar_erro
'
'If cmbfiltrarpor.Text <> "Código interno x Fornecedor" And cmbfiltrarpor.Text <> "Código de referência x Fornecedor" Then
'    cmbFornecedor.Enabled = False
'    cmbFornecedor.Clear
'Else
'    cmbFornecedor.Enabled = True
'End If
'
'ListaDetalhada.ListItems.Clear
'Lista1.ListItems.Clear
'ProcLimpaCamposTotais
'With cmbTexto
'    .Clear
'    .AddItem ""
'    If Opt_individual.Value = True Or cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Or cmbfiltrarpor = "Família x Grupo" Then
'        If cmbfiltrarpor = "Fornecedor" Then
'            Set TBAbrir = CreateObject("adodb.recordset")
'            TBAbrir.Open "Select Fornecedor from Compras_pedido where Data_aprovado IS NOT NULL group by Fornecedor order by fornecedor", Conexao, adOpenKeyset, adLockReadOnly
'            If TBAbrir.EOF = False Then
'            cmbTexto.Clear
'            With cmbTexto
'                Do While TBAbrir.EOF = False
'                    .AddItem TBAbrir!Fornecedor
'                    TBAbrir.MoveNext
'                Loop
'            End With
'            End If
'            Ordenar = "desenho"
'        Else
'            If cmbfiltrarpor = "Código de referência" Then
'                Ordenar = "n_referencia"
'            ElseIf cmbfiltrarpor = "Família" Then
'                    Ordenar = "familia"
'                ElseIf cmbfiltrarpor = "Grupo" Then
'                        Ordenar = "Grupo"
'                    ElseIf cmbfiltrarpor = "Família x Grupo" Then
'                            Ordenar = "Grupo"
'                        ElseIf cmbfiltrarpor = "Descrição" Then
'                                Ordenar = "desenho, descricao"
'                                TextoFiltro = "descricao"
'                            ElseIf cmbfiltrarpor = "Posto de trabalho" Then
'                                    Ordenar = "maquina"
'                                Else
'                                    Ordenar = "desenho"
'            End If
'
'            If cmbFornecedor <> "" Then
'            textoFornecedor = "Fornecedor = '" & cmbFornecedor & "' and "
'            End If
'
'            Set TBAbrir = CreateObject("adodb.recordset")
'            If cmbfiltrarpor = "Descrição" Then
'                StrSql = "Select " & Ordenar & " as NomeCampo1 from Compras_relatorios_historico_detalhado where " & textoFornecedor & TextoFiltro & " <> 'Null' Group by " & Ordenar
'                'Debug.print StrSql
'
'                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
'            Else
'                StrSql = "Select " & Ordenar & " as NomeCampo1 from Compras_relatorios_historico_detalhado where " & textoFornecedor & Ordenar & " <> 'Null' Group by " & Ordenar
'                'Debug.print StrSql
'                TBAbrir.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
'            End If
'            If TBAbrir.EOF = False Then
'                Do While TBAbrir.EOF = False
'                    .AddItem TBAbrir!NomeCampo1
'                    TBAbrir.MoveNext
'                Loop
'            End If
'            TBAbrir.Close
'        End If
'    End If
'    If Opt_comparativo = True And optResumido.Value = True Then
'        If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Or cmbfiltrarpor = "Família x Grupo" Then .Enabled = True Else .Enabled = False
'    End If
'End With
'
'Exit Sub
'tratar_erro:
'    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
'    Exit Sub
'End Sub

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

Acao = "filtrar"

'Se for comparativo
If Opt_comparativo.Value = True And cmbfiltrarpor = "Código interno x Fornecedor" And cmbTexto = "" Or Opt_comparativo.Value = True And cmbfiltrarpor = "Código de referência x Fornecedor" And cmbTexto = "" Or cmbfiltrarpor = "Família x Grupo" And cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
    Exit Sub
End If

'se for resumido
If optResumido.Value = True Then

'If cmbfiltrarpor = "" Then
'    NomeCampo = "o texto para pesquisa"
'    ProcVerificaAcao
'    cmbTexto.SetFocus
'    Exit Sub
'End If

    ProcVerificaPeriodoMax
    If Permitido = False Then
        USMsgBox ("Só é permitido colocar um período de " & NomeCampo & "."), vbExclamation, "CAPRIND v5.0"
        msk_fltInicio.SetFocus
        Exit Sub
    End If
    If cmbPor = "Mês" Then
        If Cmb_mes_de = "" Then
            NomeCampo = "o mês"
            ProcVerificaAcao
            Cmb_mes_de.SetFocus
            Exit Sub
        End If
        If Cmb_mes_ate = "" Then
            NomeCampo = "o mês"
            ProcVerificaAcao
            Cmb_mes_ate.SetFocus
            Exit Sub
        End If
        If Cmb_ano_de = "" Then
            NomeCampo = "o ano"
            ProcVerificaAcao
            Cmb_ano_de.SetFocus
            Exit Sub
        End If
    ElseIf cmbPor = "Ano" Then
            If Cmb_ano_de1 = "" Then
                NomeCampo = "o ano"
                ProcVerificaAcao
                Cmb_ano_de1.SetFocus
                Exit Sub
            End If
            If Cmb_ano_ate1 = "" Then
                NomeCampo = "o ano"
                ProcVerificaAcao
                Cmb_ano_ate1.SetFocus
                Exit Sub
            End If
    End If
End If

'Se for Periodo
If cmbPor = "Período" Then
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With


If Txt_limite <> "" Then
    If Txt_limite < 10 Then
        USMsgBox ("O campo (Limitar em) não pode ser menor que 10."), vbExclamation, "CAPRIND v5.0"
        Txt_limite.SetFocus
        Exit Sub
    End If
End If

Inicio = Time
End If

'=======================================================
'Grava CNPJ nos pedidos que não tem.
'=======================================================
StrSql = "update compras_pedido set CPF_CNPJ = CF.CPF_CNPJ From Compras_pedido CP Inner join Compras_fornecedores CF on CP.idfornecedor = CF.IDCliente Where CP.CPF_CNPJ IS NULL"
Conexao.Execute StrSql
StrSql = ""
'=======================================================
'Limpa os campos dos totais
'=======================================================
ProcLimpaCamposTotais
'=======================================================
'Executa as consultas
'======================================================='
'Se for resumido executa filtro resumido  e carrega lista de resumido
'=======================================================
If optResumido.Value = True Then
    ProcAbrirTabelas
    
    ProcCriaColunas
    
    'Soma e grava o total geral
    Set TBLISTA = CreateObject("adodb.recordset")
    TBLISTA.Open "Select maquina, Sum(QtdeOK) as QtdeSaida from Producao_relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' Group by Maquina", Conexao, adOpenKeyset, adLockOptimistic
    If TBLISTA.EOF = False Then
        Do While TBLISTA.EOF = False
            quantidade = IIf(IsNull(TBLISTA!QtdeSaida), 0, TBLISTA!QtdeSaida) 'Qtde. comprada
            NovoValor = Replace(quantidade, ",", ".")
            Conexao.Execute "Update Producao_relatorios Set OS = " & NovoValor & " where Maquina = '" & TBLISTA!maquina & "'"
            TBLISTA.MoveNext
        Loop
    End If
    TBLISTA.Close

If Txt_limite <> "" Then ProcVerificaLimiteRegistros
If Permitido = True Then ProcGravarTotalizacoes

Set TBLISTA = CreateObject("adodb.recordset")
    
    If Opt_individual.Value = True And optDetalhado.Value = True Then
        TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Data, Maquina", Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' and Maquina <> 'Null' order by Maquina, Ordem", Conexao, adOpenKeyset, adLockOptimistic
    End If

'ProcCarregaLista
ProcCarregaListaResumida
End If

'=======================================================
'Carrega a ListaDetalhada com os dados filtrados
'=======================================================
If optDetalhado.Value = True Then
    ProcFiltrarDetalhado
    ProcCarregaListaDetalhada
End If

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaListaResumida()
On Error GoTo tratar_erro

Familiatext = ""
Contador1 = 1
Posicao = 0

Lista1.ListItems.Clear

If TBLISTA.EOF = False Then
    TBLISTA.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBLISTA.MoveFirst
    Do While TBLISTA.EOF = False
            If TBLISTA!maquina <> "" Then
                With Lista1.ListItems
                    Contador1 = 1
                    If cmbPor = "Período" Then
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> Format(TBLISTA!Execucaoprev, "dd/mm/yy")
                            Contador1 = Contador1 + 1
                        Loop
                    ElseIf cmbPor = "Mês" Then
                            Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev
                                Contador1 = Contador1 + 1
                            Loop
                        Else
                            Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> TBLISTA!Execucaoprev
                                Contador1 = Contador1 + 1
                            Loop
                    End If
                    
                    If TBLISTA!maquina <> Familiatext Then
                        .Add , , TBLISTA!maquina
                        Posicao = Posicao + 1
                    End If
                    .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!qtdeOK), "", Format(TBLISTA!qtdeOK, "###,##0.00"))
                    
                    'Carrega valor ou quantidade total
                    Contador1 = 1
                    If Opt_valor.Value = True Then
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Valor total"
                            Contador1 = Contador1 + 1
                        Loop
                    Else
                        Do While Lista1.ColumnHeaders(Contador1 + 1).Text <> "Qtde. total"
                            Contador1 = Contador1 + 1
                        Loop
                    End If
                    .Item(.Count).SubItems(Contador1) = IIf(IsNull(TBLISTA!OS), "", Format(TBLISTA!OS, "###,##0.00"))
                End With
        End If
        Familiatext = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLISTA.Close

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    txtTotal_produtos = Format(TBLISTA!qtdeNC, "###,##0.00")
    Txt_total_servicos = Format(TBLISTA!Totalutilizada, "###,##0.00")
    txtDesconto = Format(TBLISTA!Valor2, "###,##0.00")
    txtSubtotal = Format(TBLISTA!qtdeNC + TBLISTA!Totalutilizada - TBLISTA!Valor2, "###,##0.00")
    txtTotal_ipi = Format(TBLISTA!Totalprevista, "###,##0.00")
    txtICMS_ST = Format(TBLISTA!CustoMat, "###,##0.00")
    txtDesconto_percentual = Format(TBLISTA!Valor1, "###,##0.00") & "%"
    txtTotal_geral = Format((TBLISTA!qtdeNC + TBLISTA!Totalutilizada + TBLISTA!Totalprevista + TBLISTA!CustoMat) - TBLISTA!Valor2, "###,##0.00")
    Txt_qtde_total_vendido = Format(TBLISTA!QtdePrevista, "###,##0.0000")
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarDetalhado()
On Error GoTo tratar_erro

FamiliaAntiga = ""
'===============================================================
' Verifica primeiro filtro
'===============================================================
Select Case cmbfiltrarpor
    Case "Código interno":
        Grupo = "desenho, Descricao"
        GrupoRel = "{Compras_relatorios_historico_detalhado.desenho}, {Compras_relatorios_historico_detalhado.Descricao}"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "desenho = '" & cmbTexto & "' and "
            FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.desenho} = '" & cmbTexto & "' and "
        End If
        
    Case "Código de referência":
        Grupo = "n_referencia, Descricao"
        GrupoRel = "{Compras_relatorios_historico_detalhado.n_referencia}, {Compras_relatorios_historico_detalhado.Descricao}"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "n_referencia = '" & cmbTexto & "' and "
            FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.n_referencia} = '" & cmbTexto & "' and "
        End If
        
    Case "Descrição":
        Grupo = "Descricao"
        GrupoRel = "{Compras_relatorios_historico_detalhado.Descricao}"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "Descricao = '" & cmbTexto & "' and "
            FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.Descricao} = '" & cmbTexto & "' and "
        End If
        
    Case "Família x Grupo":
        Grupo = "Grupo, Familia"
        FamiliaAntiga = "{Compras_relatorios_historico_detalhado.Grupo} = '" & cmbTexto & "' and "
        
        GrupoRel = "Grupo, Familia"
        FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.Grupo} = '" & cmbTexto & "' and "
        
    Case "Família":
        Grupo = "familia"
        GrupoRel = "{Compras_relatorios_historico_detalhado.familia}"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "familia = '" & cmbTexto & "' and "
            FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.familia}= '" & cmbTexto & "' and "
        End If
        
    Case "Grupo"
        Grupo = "Grupo"
        GrupoRel = "{Compras_relatorios_historico_detalhado.Grupo}"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "Grupo = '" & cmbTexto & "' and "
            FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.Grupo} = '" & cmbTexto & "' and "
        End If
        
    Case "Posto de trabalho":
        Grupo = "maquina"
        GrupoRel = "{Compras_relatorios_historico_detalhado.maquina}"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "maquina = '" & cmbTexto & "' and "
            FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.maquina} = '" & cmbTexto & "' and "
        Else
            FamiliaAntiga = "maquina IS NOT NULL and "
            FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.maquina} IS NOT NULL and "
        End If
        
    Case "Código interno x Fornecedor":
        Grupo = "desenho, Descricao"
        GrupoRel = "{Compras_relatorios_historico_detalhado.desenho}, {Compras_relatorios_historico_detalhado.Descricao}"
    Case "Código de referência x Fornecedor":
        Grupo = "n_referencia, Descricao"
        GrupoRel = "{Compras_relatorios_historico_detalhado.n_referencia}, {Compras_relatorios_historico_detalhado.Descricao}"
End Select

'=========================================================
'Filtrar por fornecedor
'=========================================================
If cmbFornecedor.Text <> "" Then
    FamiliaAntiga = FamiliaAntiga & "Fornecedor = '" & cmbFornecedor & "' and "
    FamiliaAntigaRel = FamiliaAntigaRel & "{Compras_relatorios_historico_detalhado.Fornecedor} = '" & cmbFornecedor & "' and "
End If
'==========================================================
' Ser relatório detalhado
'==========================================================
        
  
    If cmbPor.Text = "Período" Then
        StrSql = "Select * from Compras_relatorios_historico_detalhado where " & FamiliaAntiga & " (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' order by Data, IDLista"
        FormulaRelatorio = FamiliaAntigaRel & "{Compras_relatorios_historico_detalhado.Data} >= Date('" & Format(msk_fltInicio.Value, "Short Date") & "') And {Compras_relatorios_historico_detalhado.Data} <= Date('" & Format(msk_fltFim.Value, "Short Date") & "')"
    End If
    
    If cmbPor.Text = "Mês" Then
        qt = FunVerificaMes(Cmb_mes_de)
        Qtd = FunVerificaMes(Cmb_mes_ate)

        StrSql = "Select * from Compras_relatorios_historico_detalhado where " & FamiliaAntiga & " Mes >= '" & qt & "' AND Mes <= '" & Qtd & "' And Year(Data) >= '" & Cmb_ano_de & "' And Ano <= '" & Cmb_ano_ate & "' order by Data, IDLista"
        FormulaRelatorio = FamiliaAntigaRel & "{Compras_relatorios_historico_detalhado.Mes} >= " & qt & " and {Compras_relatorios_historico_detalhado.Mes} <= " & Qtd & " And {Compras_relatorios_historico_detalhado.ano} >= " & Cmb_ano_de & " And {Compras_relatorios_historico_detalhado.Ano} <= " & Cmb_ano_ate & ""
    
    End If

    If cmbPor.Text = "Ano" Then
        StrSql = "Select * from Compras_relatorios_historico_detalhado where " & FamiliaAntiga & " Ano >= '" & Cmb_ano_de1 & "' And Ano <= '" & Cmb_ano_ate1 & "' order by Data, IDLista"
        FormulaRelatorio = FamiliaAntigaRel & "{Compras_relatorios_historico_detalhado.ano} >= " & Cmb_ano_de1 & " And {Compras_relatorios_historico_detalhado.Ano} <= " & Cmb_ano_ate1 & ""
    End If

    'Debug.print StrSql
    'Debug.print FormulaRelatorio

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrirTabelas()
On Error GoTo tratar_erro

'Deleta registros e adiciona novos
ProcExcluirDadosProducaoRelatorios
ProcExcluirDadosProducaoRelatoriosTotal

FamiliaAntiga = ""
Select Case cmbfiltrarpor
    Case "Código interno":
        Grupo = "desenho, Descricao"
        If cmbTexto <> "" Then FamiliaAntiga = "desenho = '" & cmbTexto & "' and "
    Case "Código de referência":
        Grupo = "n_referencia, Descricao"
        If cmbTexto <> "" Then FamiliaAntiga = "n_referencia = '" & cmbTexto & "' and "
    Case "Descrição":
        Grupo = "Descricao"
        If cmbTexto <> "" Then FamiliaAntiga = "Descricao = '" & cmbTexto & "' and "
    Case "Família x Grupo":
        Grupo = "Grupo, Familia"
        FamiliaAntiga = "Grupo = '" & cmbTexto & "' and "
    Case "Família":
        Grupo = "familia"
        If cmbTexto <> "" Then FamiliaAntiga = "familia = '" & cmbTexto & "' and "
    Case "Grupo"
        Grupo = "Grupo"
        If cmbTexto <> "" Then FamiliaAntiga = "Grupo = '" & cmbTexto & "' and "
    Case "Fornecedor":
        Grupo = "Fornecedor"
        If cmbFornecedor <> "" Then FamiliaAntiga = "Fornecedor = '" & cmbFornecedor & "' and "
    Case "Posto de trabalho":
        Grupo = "maquina"
        If cmbTexto <> "" Then FamiliaAntiga = "maquina = '" & cmbTexto & "' and " Else FamiliaAntiga = "maquina IS NOT NULL and "
    Case "Código interno x Fornecedor":
        Grupo = "desenho, Descricao"
    Case "Código de referência x Fornecedor":
        Grupo = "n_referencia, Descricao"
End Select

If cmbFornecedor.Text <> "" Then
    FamiliaAntiga = FamiliaAntiga & "Fornecedor = '" & cmbFornecedor & "' and "
End If

        
Set TBCarteira = CreateObject("adodb.recordset")
    
    If Opt_quantidade.Value = True Then
        TextoFiltro = "Quant_Comp"
    Else
        TextoFiltro = "preco_total"
    End If
    
    
    Par1 = ""
    Permitido = False
    Select Case cmbPor
        Case "Período":
            Dataini = msk_fltInicio
            DataFim = msk_fltFim
            Do While Dataini <= DataFim
                If Permitido = False Then Par1 = "[" & Dataini & "]" Else Par1 = Par1 & " , [" & Dataini & "]"
                Permitido = True
                Dataini = Dataini + 1
            Loop
            Pesquisa = "(Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Data In (" & Par1 & "))"
            Pesquisa2 = "Data"
        Case "Mês":
            qt = FunVerificaMes(Cmb_mes_de)
            Qtd = FunVerificaMes(Cmb_mes_ate)
            MesX = qt
            MesX1 = Qtd
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Month(Data) >= '" & MesX & "' and Year(Data) = '" & Cmb_ano_de & "' and Month(Data) <= '" & MesX1 & "' and Year(Data) = '" & Cmb_ano_ate & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Mes In (" & Par1 & "))"
            Pesquisa2 = "Mes"
        Case "Ano":
            qt = Cmb_ano_de1
            Qtd = Cmb_ano_ate1
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Year(Data) >= '" & Cmb_ano_de1 & "' and Year(Data) <= '" & Cmb_ano_ate1 & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Ano In (" & Par1 & "))"
            Pesquisa2 = "Ano"
    End Select
    
    If Opt_individual.Value = True Then
        StrSql = "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where " & FamiliaAntiga & Pesquisa & ") p " & Pesquisa1 & " pvt"
        'Debug.print StrSql
        TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
    Else
        Pesquisa3 = "Desenho is not null"
        If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Then
            TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where Fornecedor = '" & cmbTexto & "' and " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " pvt", Conexao, adOpenKeyset, adLockReadOnly
        ElseIf cmbfiltrarpor = "Família x Grupo" Then
                TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where Grupo = '" & cmbTexto & "' and " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " pvt", Conexao, adOpenKeyset, adLockReadOnly
            Else
                StrSql = "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " as pvt"

                TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockReadOnly
        End If
    End If
'Debug.print StrSql

ProcFiltrar1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarResumido()
On Error GoTo tratar_erro

FamiliaAntiga = ""
'===============================================================
' Verifica primeiro filtro
'===============================================================
Select Case cmbfiltrarpor
    Case "Código interno":
        Grupo = "desenho, Descricao"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "desenho = '" & cmbTexto & "' and "
        End If
        
    Case "Código de referência":
        Grupo = "n_referencia, Descricao"
        If cmbTexto <> "" Then
            FamiliaAntiga = "n_referencia = '" & cmbTexto & "' and "
            FamiliaAntigaRel = "{Compras_relatorios_historico_detalhado.n_referencia} = '" & cmbTexto & "' and "
        End If
        
    Case "Descrição":
        Grupo = "Descricao"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "Descricao = '" & cmbTexto & "' and "
        End If
        
    Case "Família x Grupo":
        Grupo = "Grupo, Familia"
        FamiliaAntiga = "{Compras_relatorios_historico_detalhado.Grupo} = '" & cmbTexto & "' and "
        
        
    Case "Família":
        Grupo = "familia"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "familia = '" & cmbTexto & "' and "
        End If
        
    Case "Grupo"
        Grupo = "Grupo"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "Grupo = '" & cmbTexto & "' and "
        End If
        
    Case "Posto de trabalho":
        Grupo = "maquina"
        
        If cmbTexto <> "" Then
            FamiliaAntiga = "maquina = '" & cmbTexto & "' and "
        Else
            FamiliaAntiga = "maquina IS NOT NULL and "
        End If
        
    Case "Código interno x Fornecedor":
        Grupo = "desenho, Descricao"
    Case "Código de referência x Fornecedor":
        Grupo = "n_referencia, Descricao"
End Select

'=========================================================
'Filtrar por fornecedor
'=========================================================
If cmbFornecedor.Text <> "" Then
    FamiliaAntiga = FamiliaAntiga & "Fornecedor = '" & cmbFornecedor & "' and "
End If

    
    If cmbPor.Text = "Período" Then
        StrSql = "Select * from Compras_relatorios_historico_detalhado where " & FamiliaAntiga & " (Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "' order by Data, IDListaDetalhada"
    End If
    
    If cmbPor.Text = "Mês" Then
        StrSql = "Select * from Compras_relatorios_historico_detalhado where " & FamiliaAntiga & " Mes >= '" & qt & "' AND Mes <= '" & Qtd & "' And Year(Data) >= '" & Cmb_ano_de & "' And Ano <= '" & Cmb_ano_ate & "' order by Data, IDListaDetalhada"
    End If

    If cmbPor.Text = "Ano" Then
        StrSql = "Select * from Compras_relatorios_historico_detalhado where " & FamiliaAntiga & " Ano >= '" & Cmb_ano_de1 & "' And Ano <= '" & Cmb_ano_ate1 & "' order by Data, IDListaDetalhada"
    End If

    'Debug.print StrSql
    'Debug.print FormulaRelatorio
    

    If Opt_quantidade.Value = True Then
        TextoFiltro = "Quant_Comp"
    Else
        TextoFiltro = "preco_total"
    End If
    
Par1 = ""
Permitido = False

    Select Case cmbPor
        Case "Período":
            Dataini = msk_fltInicio
            DataFim = msk_fltFim
            Do While Dataini <= DataFim
                If Permitido = False Then Par1 = "[" & Dataini & "]" Else Par1 = Par1 & " , [" & Dataini & "]"
                Permitido = True
                Dataini = Dataini + 1
            Loop
            Pesquisa = "(Data) Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Data In (" & Par1 & "))"
            Pesquisa2 = "Data"
        Case "Mês":
            qt = FunVerificaMes(Cmb_mes_de)
            Qtd = FunVerificaMes(Cmb_mes_ate)
            MesX = qt
            MesX1 = Qtd
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Month(Data) >= '" & MesX & "' and Year(Data) = '" & Cmb_ano_de & "' and Month(Data) <= '" & MesX1 & "' and Year(Data) = '" & Cmb_ano_ate & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Mes In (" & Par1 & "))"
            Pesquisa2 = "Mes"
        Case "Ano":
            qt = Cmb_ano_de1
            Qtd = Cmb_ano_ate1
            Do While qt <= Qtd
                If Permitido = False Then Par1 = "[" & qt & "]" Else Par1 = Par1 & ", [" & qt & "]"
                Permitido = True
                qt = qt + 1
            Loop
            Pesquisa = "Year(Data) >= '" & Cmb_ano_de1 & "' and Year(Data) <= '" & Cmb_ano_ate1 & "'"
            Pesquisa1 = "PIVOT (Sum(" & TextoFiltro & ") for Ano In (" & Par1 & "))"
            Pesquisa2 = "Ano"
    End Select
    If Opt_individual.Value = True Then
        TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where " & FamiliaAntiga & Pesquisa & ") p " & Pesquisa1 & " pvt", Conexao, adOpenKeyset, adLockReadOnly
    Else
        Pesquisa3 = "Desenho is not null"
        If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Then
            TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where Fornecedor = '" & cmbTexto & "' and " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " pvt", Conexao, adOpenKeyset, adLockReadOnly
        ElseIf cmbfiltrarpor = "Família x Grupo" Then
                TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where Grupo = '" & cmbTexto & "' and " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " pvt", Conexao, adOpenKeyset, adLockReadOnly
            Else
                TBCarteira.Open "SELECT " & Grupo & ", " & Par1 & " From (Select " & Grupo & ", " & Pesquisa2 & ", " & TextoFiltro & " from Compras_relatorios_historico_detalhado Where " & Pesquisa & " and " & Pesquisa3 & ") p " & Pesquisa1 & " as pvt", Conexao, adOpenKeyset, adLockReadOnly
        End If
    End If
    
ProcFiltrar1



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrar1()
On Error GoTo tratar_erro

Produto = ""
Familiatext = ""
quantidade = 0
QTLOTE = 0
Valor_Produto = 0
Valor_Cofins_Serv = 0
ValorIPI = 0
Valor_Cofins_Prod = 0
Valor_ICMS_SN = 0
Desconto = 0

If TBCarteira.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBCarteira.EOF = False
            ProcCriarResumido
        TBCarteira.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBCarteira.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCriarResumido()
On Error GoTo tratar_erro

Permitido = True
Select Case cmbPor
    Case "Período":
        qt = 0
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            qt = qt + 1
            ProcEnviaDadosResumido
            Dataini = Dataini + 1
        Loop
    Case "Mês":
        qt = MesX
        Qtd = MesX1
        Do While qt <= Qtd
            ProcEnviaDadosResumido
            qt = qt + 1
        Loop
    Case "Ano":
        qt = Cmb_ano_de1
        Qtd = Cmb_ano_ate1
        Do While qt <= Qtd
            ProcEnviaDadosResumido
            qt = qt + 1
        Loop
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

Select Case cmbfiltrarpor
    Case "Código interno": Familiatext = TBCarteira!Desenho
    Case "Código de referência": Familiatext = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    Case "Descrição": Familiatext = TBCarteira!Descricao
    Case "Família x Grupo": Familiatext = TBCarteira!Familia
    Case "Família": Familiatext = TBCarteira!Familia
    Case "Grupo": Familiatext = TBCarteira!Grupo
    Case "Fornecedor": Familiatext = TBCarteira!Fornecedor
    Case "Posto de trabalho": Familiatext = IIf(IsNull(TBCarteira!maquina), "", TBCarteira!maquina)
    Case "Código interno x Fornecedor": Familiatext = TBCarteira!Desenho
    Case "Código de referência x Fornecedor": Familiatext = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
End Select
Select Case cmbPor
    Case "Período": DataTexto = Dataini
    Case "Mês": DataTexto = "01/" & qt & "/" & Cmb_ano_de
    Case "Ano": DataTexto = "01" & "/01/" & qt
End Select
Set TBProdutividade = CreateObject("adodb.recordset")
TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
TBProdutividade.AddNew
TBProdutividade!Data = Format(DataTexto, "dd/mm/yy")
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
Select Case cmbPor
    Case "Período":
        DiaX = Dataini
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!qtdeOK = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00"))
        TBProdutividade!Execucaoprev = Format(Dataini, "dd/mm/yy")
    Case "Mês":
        Select Case qt
            Case 1: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![1]), 0, Format(TBCarteira![1], "###,##0.00"))
            Case 2: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![2]), 0, Format(TBCarteira![2], "###,##0.00"))
            Case 3: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![3]), 0, Format(TBCarteira![3], "###,##0.00"))
            Case 4: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![4]), 0, Format(TBCarteira![4], "###,##0.00"))
            Case 5: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![5]), 0, Format(TBCarteira![5], "###,##0.00"))
            Case 6: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![6]), 0, Format(TBCarteira![6], "###,##0.00"))
            Case 7: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![7]), 0, Format(TBCarteira![7], "###,##0.00"))
            Case 8: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![8]), 0, Format(TBCarteira![8], "###,##0.00"))
            Case 9: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![9]), 0, Format(TBCarteira![9], "###,##0.00"))
            Case 10: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![10]), 0, Format(TBCarteira![10], "###,##0.00"))
            Case 11: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![11]), 0, Format(TBCarteira![11], "###,##0.00"))
            Case 12: TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira![12]), 0, Format(TBCarteira![12], "###,##0.00"))
        End Select
        TBProdutividade!Execucaoprev = qt & "/" & Cmb_ano_de
    Case "Ano":
        DiaX = qt
        TotalCreditar = IIf(IsNull(TBCarteira(DiaX)), 0, TBCarteira(DiaX))
        TBProdutividade!qtdeOK = IIf(IsNull(TotalCreditar), 0, Format(TotalCreditar, "###,##0.00"))
        TBProdutividade!Execucaoprev = qt
End Select

TBProdutividade!Ordem = qt

If cmbfiltrarpor = "Código interno" Or cmbfiltrarpor = "Código de referência" Or cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Then
    TBProdutividade!maquina = Left(Familiatext & " " & TBCarteira!Descricao, 25)
Else
    TBProdutividade!maquina = Left(Familiatext, 25)
End If

TBProdutividade.Update
TBProdutividade.Close

Select Case cmbfiltrarpor
    Case "Código interno": Produto = TBCarteira!Desenho
    Case "Código de referência": Produto = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
    Case "Descrição": Produto = TBCarteira!Descricao
    Case "Família x Grupo": Produto = TBCarteira!Familia
    Case "Família": Produto = TBCarteira!Familia
    Case "Grupo": Produto = TBCarteira!Grupo
    Case "Fornecedor": Produto = TBCarteira!Fornecedor
    Case "Posto de trabalho": Produto = IIf(IsNull(TBCarteira!maquina), "", TBCarteira!maquina)
    Case "Código interno x Fornecedor": Produto = TBCarteira!Desenho
    Case "Código de referência x Fornecedor": Produto = IIf(IsNull(TBCarteira!N_referencia), "", TBCarteira!N_referencia)
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCriaColunas()
On Error GoTo tratar_erro

Lista1.ColumnHeaders.Clear
Contador = 1
With Lista1.ColumnHeaders
    .Add
    If cmbfiltrarpor <> "Código interno x Fornecedor" And cmbfiltrarpor <> "Código de referência x Fornecedor" And cmbfiltrarpor <> "Família x Grupo" Then
        .Item(Contador).Text = cmbfiltrarpor.Text
    Else
        If cmbfiltrarpor = "Código interno x Fornecedor" Then
            .Item(Contador).Text = "Código interno"
        ElseIf cmbfiltrarpor = "Código de referência x Fornecedor" Then
                .Item(Contador).Text = "Código de referência"
            Else
                .Item(Contador).Text = "Família"
        End If
    End If
    .Item(Contador).Width = 3500
    If cmbPor.Text = "Período" Then
        Dataini = msk_fltInicio
        DataFim = msk_fltFim
        Do While Dataini <= DataFim
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = Format(Dataini, "dd/mm/yy")
            .Item(Contador).Alignment = lvwColumnRight
            .Item(Contador).Width = 1000
            Dataini = Dataini + 1
        Loop
    End If
    If cmbPor.Text = "Mês" Then
        qt = FunVerificaMes(Cmb_mes_de)
        Qtd = FunVerificaMes(Cmb_mes_ate)
        Do While qt <= Qtd
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = qt & "/" & Cmb_ano_de
            .Item(Contador).Alignment = lvwColumnRight
            qt = qt + 1
        Loop
    End If
    If cmbPor.Text = "Ano" Then
        qt = Cmb_ano_de1
        Do While qt <= Cmb_ano_ate1
            .Add
            Contador = Contador + 1
            .Item(Contador).Text = qt
            .Item(Contador).Alignment = lvwColumnRight
            qt = qt + 1
        Loop
    End If
    .Add
    Contador = Contador + 1
    If Opt_valor.Value = True Then
        .Item(Contador).Text = "Valor total"
        .Item(Contador).Alignment = lvwColumnRight
    Else
        .Item(Contador).Text = "Qtde. total"
        .Item(Contador).Alignment = lvwColumnRight
    End If
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaLimiteRegistros()
On Error GoTo tratar_erro

Contador1 = Txt_limite
Valor_total = 0
Fornecedor = ""
Set TBListaDetalhada = CreateObject("adodb.recordset")
TBListaDetalhada.Open "Select maquina, OS from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' Group by Maquina, OS order by OS desc", Conexao, adOpenKeyset, adLockReadOnly
If TBListaDetalhada.EOF = False Then
    If TBListaDetalhada.RecordCount < 10 Then Exit Sub
    Do While Contador1 <> 0
        If Fornecedor <> TBListaDetalhada!maquina Then
            Valor_total = TBListaDetalhada!OS
            Contador1 = Contador1 - 1
        End If
        Fornecedor = TBListaDetalhada!maquina
        TBListaDetalhada.MoveNext
    Loop
End If
TBListaDetalhada.Close
NovoValor = Replace(Valor_total, ",", ".")
Conexao.Execute "DELETE from Producao_Relatorios where OS < " & NovoValor & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    If Opt_individual.Value = True Then
        TextoFiltroPadrao = FamiliaAntiga & Pesquisa
    Else
        If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Or cmbfiltrarpor = "Família x Grupo" Then
            If cmbfiltrarpor = "Código interno x Fornecedor" Or cmbfiltrarpor = "Código de referência x Fornecedor" Then TextoFiltro = "Fornecedor" Else TextoFiltro = "Grupo"
            TextoFiltroPadrao = TextoFiltro & " = '" & cmbTexto & "' and " & Pesquisa
        Else
            TextoFiltroPadrao = IIf(cmbfiltrarpor = "Posto de trabalho", FamiliaAntiga, "") & Pesquisa
        End If
    End If
    'Produtos e IPI
    Set TBListaDetalhada = CreateObject("adodb.recordset")
    TBListaDetalhada.Open "Select Sum(Quant_Comp * preco_unitario) as Valor, Sum(vlripi) as ValorIPI, Sum(valordesconto * Quant_Comp) as Desconto, Sum(Valor_ICMS_ST) as ICMS_ST from Compras_relatorios_historico_detalhado where " & TextoFiltroPadrao & " and Tipo = 'P'", Conexao, adOpenKeyset, adLockOptimistic
    If TBListaDetalhada.EOF = False Then
        Valor_Produto = IIf(IsNull(TBListaDetalhada!valor), 0, TBListaDetalhada!valor) 'Valor total produto
        ValorIPI = IIf(IsNull(TBListaDetalhada!ValorIPI), 0, TBListaDetalhada!ValorIPI) 'Valor total IPI
        Valor_Cofins_Prod = IIf(IsNull(TBListaDetalhada!Desconto), 0, TBListaDetalhada!Desconto)
        Valor_ICMS_SN = IIf(IsNull(TBListaDetalhada!ICMS_ST), 0, TBListaDetalhada!ICMS_ST) 'Valor total ICMS ST
        QTLOTE = IIf(IsNull(TBListaDetalhada!valor), 0, TBListaDetalhada!valor) + IIf(IsNull(TBListaDetalhada!ValorIPI), 0, TBListaDetalhada!ValorIPI) 'Valor total
    End If
        
    'Serviços
    Set TBListaDetalhada = CreateObject("adodb.recordset")
    TBListaDetalhada.Open "Select Sum(Quant_Comp * preco_unitario) as Valor, Sum(valordesconto * Quant_Comp) as Desconto from Compras_relatorios_historico_detalhado where " & TextoFiltroPadrao & " and Tipo = 'S'", Conexao, adOpenKeyset, adLockOptimistic
    If TBListaDetalhada.EOF = False Then
        Valor_Cofins_Serv = IIf(IsNull(TBListaDetalhada!valor), 0, TBListaDetalhada!valor) 'Valor total serviços
        Desconto = IIf(IsNull(TBListaDetalhada!Desconto), 0, TBListaDetalhada!Desconto)
    End If
    
    'Quantidade
    Set TBListaDetalhada = CreateObject("adodb.recordset")
    TBListaDetalhada.Open "Select Sum(Quant_Comp) as QtdeSaida from Compras_relatorios_historico_detalhado where " & TextoFiltroPadrao, Conexao, adOpenKeyset, adLockOptimistic
    If TBListaDetalhada.EOF = False Then
        quantidade = IIf(IsNull(TBListaDetalhada!QtdeSaida), 0, TBListaDetalhada!QtdeSaida) 'Qtde. comprada
    End If
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew

Select Case cmbPor
    Case "Período":
        Tipo = "D"
        TBAbrir!Data1 = Format(msk_fltInicio.Value, "dd/mm/yy")
        TBAbrir!Data2 = Format(msk_fltFim.Value, "dd/mm/yy")
    Case "Mês":
        Tipo = "M"
        TBAbrir!Data1 = Cmb_mes_de & "/" & Cmb_ano_de
        TBAbrir!Data2 = Cmb_mes_ate & "/" & Cmb_ano_ate
    Case "Ano":
        Tipo = "A"
        TBAbrir!Data1 = Cmb_ano_de1
        TBAbrir!Data2 = Cmb_ano_ate1
End Select

TBAbrir!Data3 = Tipo
If cmbfiltrarpor <> "Código interno x Fornecedor" And cmbfiltrarpor <> "Código de referência x Fornecedor" And cmbfiltrarpor <> "Família x Grupo" Then
    If Opt_individual.Value = True And cmbTexto <> "" Then TBAbrir!Texto = cmbfiltrarpor & " : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor
    TBAbrir!QtdeOrdem = "1"
Else
    If cmbfiltrarpor = "Código interno x Fornecedor" Then
        If Opt_individual.Value = True And cmbTexto <> "" Then
            TBAbrir!Texto = "Código interno" & " : " & cmbTexto
        Else
            TBAbrir!Texto = "Código interno"
        End If
        TBAbrir!QtdeOrdem = "2"
    ElseIf cmbfiltrarpor = "Código de referência x Fornecedor" Then
            If Opt_individual.Value = True And cmbTexto <> "" Then
                TBAbrir!Texto = "Código de referência" & " : " & cmbTexto
            Else
                TBAbrir!Texto = "Código de referência"
            End If
            TBAbrir!QtdeOrdem = "2"
        ElseIf cmbfiltrarpor = "Família x Grupo" Then
            If Opt_individual.Value = True And cmbTexto <> "" Then
                TBAbrir!Texto = "Fámilia" & " : " & cmbTexto
            Else
                TBAbrir!Texto = "Família"
            End If
            TBAbrir!QtdeOrdem = "3"
    End If
    TBAbrir!Texto1 = cmbTexto
End If

TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
If Opt_quantidade.Value = True Then TBAbrir!Turno = True Else TBAbrir!Turno = False
TBAbrir!QtdePrevista = quantidade 'Qtde. comprada
TBAbrir!QtdeProduzida = QTLOTE 'Valor total
TBAbrir!qtdeNC = Valor_Produto 'Valor total produtos
TBAbrir!Totalutilizada = Format(Valor_Cofins_Serv, "###,##0.00") 'Valor serviços
TBAbrir!Totalprevista = Format(ValorIPI, "###,##0.00") 'Valor total IPI
TBAbrir!Valor2 = Format(Valor_Cofins_Prod + Desconto, "###,##0.00") 'Valor total desconto
If TBAbrir!qtdeNC + TBAbrir!Totalutilizada = 0 Then TBAbrir!Valor1 = 0 Else TBAbrir!Valor1 = (TBAbrir!Valor2 * 100) / (TBAbrir!qtdeNC + TBAbrir!Totalutilizada)  'Percentual desconto
TBAbrir!CustoMat = Format(Valor_ICMS_SN, "###,##0.00") 'Valor total ICMS ST

If Opt_valor = True Then TBAbrir!TotalEficiencia = 1 Else TBAbrir!TotalEficiencia = 2
TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaPeriodoMax()
On Error GoTo tratar_erro

Permitido = True
If cmbPor = "Período" Then
    Dataini = msk_fltInicio
    DataFim = msk_fltFim
    If DataFim - Dataini > 10 Then
        Permitido = False
        NomeCampo = "10 dias"
    End If
ElseIf Cmb_ano_ate1 - Cmb_ano_de1 > 5 Then
        Permitido = False
        NomeCampo = "5 anos"
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_LostFocus()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
ProcCarregaCombo_Fornecedor
Else
ProcCarregaCombo_Texto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_LostFocus()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

If cmbfiltrarpor.Text = "Código interno x Fornecedor" Or cmbfiltrarpor.Text = "Código de referência x Fornecedor" Then
ProcCarregaCombo_Fornecedor
Else
ProcCarregaCombo_Texto
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
If Opt_comparativo.Value = True Then
    optDetalhado.Enabled = False
    optResumido.Value = True
    cmbTexto.ListIndex = -1
    cmbTexto.Enabled = False
    Txt_limite.Locked = False
    Txt_limite.TabStop = True
End If
With cmbfiltrarpor
    .Clear
        .Clear
        .AddItem " "
        .AddItem "Código interno"
        .AddItem "Fornecedor"
        .AddItem "Descrição"
        .AddItem "Família"
        .AddItem "Grupo"
        .AddItem "Posto de trabalho"
        .AddItem "Centro de custo"
        .AddItem "Código de referência"
        .AddItem "Detalhe"
        .AddItem "Código interno x Fornecedor"
        .AddItem "Código de referência x Fornecedor"
        .AddItem "Família x Grupo"
        .Text = " "
    End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
    Txt_limite = ""
    Txt_limite.Locked = True
    Txt_limite.TabStop = False
End If
With cmbfiltrarpor
        .Clear
        .AddItem " "
        .AddItem "Código interno"
        .AddItem "Fornecedor"
        .AddItem "Descrição"
        .AddItem "Família"
        .AddItem "Grupo"
        .AddItem "Posto de trabalho"
        .AddItem "Centro de custo"
        .AddItem "Código de referência"
        .AddItem "Detalhe"
        .AddItem "Código interno x Fornecedor"
        .AddItem "Código de referência x Fornecedor"
        .AddItem "Família x Grupo"
        .Text = " "
End With
ProcCarregaCombo_Texto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_quantidade_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_valor_Click()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    ListaDetalhada.ListItems.Clear
    ListaDetalhada.Visible = True
    Lista1.ListItems.Clear
    Lista1.Visible = False
    Opt_valor.Value = False
    Opt_valor.Enabled = False
    Opt_quantidade.Value = False
    Opt_quantidade.Enabled = False
    ProcMostrarEsconderCombosData
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    ListaDetalhada.ListItems.Clear
    ListaDetalhada.Visible = False
    Lista1.ListItems.Clear
    Lista1.Visible = True
    Opt_valor.Value = True
    Opt_valor.Enabled = True
    Opt_quantidade.Enabled = True
    ProcMostrarEsconderCombosData
    ProcLimpaCamposTotais
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcMostrarEsconderCombosData()
On Error GoTo tratar_erro

If cmbPor = "Período" Then
    msk_fltInicio.Visible = True
    msk_fltFim.Visible = True
    Cmb_mes_de.Visible = False
    Cmb_mes_ate.Visible = False
    Cmb_ano_de.Visible = False
    Cmb_ano_ate.Visible = False
    Cmb_ano_de1.Visible = False
    Cmb_ano_ate1.Visible = False
End If

If cmbPor = "Mês" Then
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = True
        Cmb_mes_ate.Visible = True
        Cmb_ano_de.Visible = True
        Cmb_ano_ate.Visible = True
        Cmb_ano_de1.Visible = False
        Cmb_ano_ate1.Visible = False
End If

If cmbPor = "Ano" Then
        msk_fltInicio.Visible = False
        msk_fltFim.Visible = False
        Cmb_mes_de.Visible = False
        Cmb_mes_ate.Visible = False
        Cmb_ano_de.Visible = False
        Cmb_ano_ate.Visible = False
        Cmb_ano_de1.Visible = True
        Cmb_ano_ate1.Visible = True
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_limite_Change()
On Error GoTo tratar_erro

ListaDetalhada.ListItems.Clear
Lista1.ListItems.Clear
ProcLimpaCamposTotais
If Txt_limite <> "" Then
    VerifNumero = Txt_limite
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_limite = ""
        txt_ValorPago.SetFocus
        Exit Sub
    End If
End If
    
Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    If optDetalhado.Value = True Then
        If ListaDetalhada.ListItems.Count = 0 Then Exit Sub
    Else
        If Lista1.ListItems.Count = 0 Then Exit Sub
    End If
Else
    If Lista1.ListItems.Count = 0 Then Exit Sub
End If
frmCompras_Relatorios_Historico_MenuImpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarTexto()
On Error GoTo tratar_erro
Dim CampoFiltro As String
CampoFiltro = ""

If txtTexto.Text <> "" Then

Select Case cmbfiltrarpor.Text
        Case "Status":  CampoFiltro = "Status_Item"
        Case "Pedido":  CampoFiltro = "Pedido"
        Case "Código interno":  CampoFiltro = "Desenho"
        Case "Fornecedor": CampoFiltro = "Fornecedor"
        Case "Código de referência": CampoFiltro = "n_referencia"
        Case "Descrição": CampoFiltro = "Descricao"
        Case "Família": CampoFiltro = "Familia"
        Case "Grupo":  CampoFiltro = "Grupo"
        Case "Posto de trabalho": CampoFiltro = "maquina"
        Case "Centro de custo": CampoFiltro = "Setor"
        Case "Detalhe": CampoFiltro = "Detalhe"
End Select

If Optinicio.Value = True Then
Filtro = " LIKE '" & txtTexto.Text & "%'"
FiltroRel = " LIKE '" & txtTexto.Text & "*'"
End If

If Optmeio.Value = True Then
Filtro = " LIKE '%" & txtTexto.Text & "%'"
FiltroRel = " LIKE '*" & txtTexto.Text & "*'"
End If

If Optfim.Value = True Then
Filtro = " LIKE '%" & txtTexto.Text & "'"
FiltroRel = " LIKE '*" & txtTexto.Text & "'"
End If

If optIgual.Value = True Then
Filtro = " = '" & txtTexto.Text & "'"
FiltroRel = " = '" & txtTexto.Text & "'"
End If



FiltroStatus = IIf(txtTexto.Text = "COMPRADO", "N_RECEBIDO", txtTexto.Text)

        Select Case cmbPor.Text
                Case "Período":    StrSql = "select * from Compras_relatorios_historico_detalhado where " & CampoFiltro & Filtro & "  and data >= '" & msk_fltInicio & "'  and data <= '" & msk_fltFim & "'   order by Ano, Mes"
                FormulaRelatorio = "{Compras_relatorios_historico_detalhado." & CampoFiltro & "}" & FiltroRel & Filtro2Rel & " and  {Compras_relatorios_historico_detalhado.Data} >= Date('" & Format(msk_fltInicio.Value, "Short Date") & "') And {Compras_relatorios_historico_detalhado.Data} <= Date('" & Format(msk_fltFim.Value, "Short Date") & "')"
                
                Case "Mês":
                    MesDe = FunVerificaMes(Cmb_mes_de)
                    MesAte = FunVerificaMes(Cmb_mes_ate)
                    StrSql = "select * from Compras_relatorios_historico_detalhado where " & CampoFiltro & Filtro & " and Ano >= " & Cmb_ano_de & "  and Ano <= " & Cmb_ano_ate & " and Mes >= " & MesDe & "   and  Mes <= " & MesAte & " order by Ano, Mes"
                 FormulaRelatorio = "{Compras_relatorios_historico_detalhado." & CampoFiltro & "} " & FiltroRel & Filtro2Rel & " and  {Compras_relatorios_historico_detalhado.Mes} >= " & qt & " and {Compras_relatorios_historico_detalhado.Mes} <= " & Qtd & " And {Compras_relatorios_historico_detalhado.ano} >= " & Cmb_ano_de & " And {Compras_relatorios_historico_detalhado.Ano} <= " & Cmb_ano_ate & ""
                
                Case "Ano":     StrSql = "select * from Compras_relatorios_historico_detalhado where " & CampoFiltro & Filtro & " and Ano >= " & Cmb_ano_de1 & "  and Ano <= " & Cmb_ano_ate1 & "   order by Ano, Mes"
                 FormulaRelatorio = "{Compras_relatorios_historico_detalhado." & CampoFiltro & "}" & FiltroRel & Filtro2Rel & " and  {Compras_relatorios_historico_detalhado.ano} >= " & Cmb_ano_de1 & " And {Compras_relatorios_historico_detalhado.Ano} <= " & Cmb_ano_ate1 & ""
        
        End Select
        
If StrSql <> "" Then
    'Debug.print FormulaRelatorio
    ProcCarregaListaDetalhada
    End If
Else
    ListaDetalhada.ListItems.Clear
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcFiltrarCombo()
On Error GoTo tratar_erro
Dim CampoFiltro, Filtro As String
FormulaRel = ""

If cmbTexto.Text <> "" Then

Select Case cmbfiltrarpor.Text
        Case "Status":  CampoFiltro = "Status_Item"
        Case "Pedido":  CampoFiltro = "Pedido"
        Case "Código interno":  CampoFiltro = "Desenho"
        Case "Fornecedor": CampoFiltro = "Fornecedor"
        Case "Código de referência": CampoFiltro = "n_referencia"
        Case "Descrição": CampoFiltro = "Descricao"
        Case "Família": CampoFiltro = "Familia"
        Case "Grupo":  CampoFiltro = "Grupo"
        Case "Posto de trabalho": CampoFiltro = "maquina"
        Case "Centro de custo": CampoFiltro = "Setor"
        Case "Detalhe": CampoFiltro = "Detalhe"
        Case "Código interno x Fornecedor":
        CampoFiltro = "Desenho"
        Filtro2 = "' and fornecedor = '" & cmbFornecedor.Text
        Filtro2Rel = "' and {Compras_relatorios_historico_detalhado.fornecedor} = '" & cmbFornecedor.Text
        Case "Código de referência x Fornecedor": CampoFiltro = "n_referencia"
        Case "Família x Grupo":  CampoFiltro = "Grupo"
End Select

Filtro = IIf(cmbTexto.Text = "COMPRADO", "N_RECEBIDO", cmbTexto.Text)

        Select Case cmbPor.Text
                Case "Período":
                StrSql = "select * from Compras_relatorios_historico_detalhado where " & CampoFiltro & " = '" & Filtro & Filtro2 & "' and data >= '" & msk_fltInicio & "'  and data <= '" & msk_fltFim & "'   order by Ano, Mes"
                FormulaRelatorio = "{Compras_relatorios_historico_detalhado." & CampoFiltro & "} = '" & Filtro & Filtro2Rel & "' and  {Compras_relatorios_historico_detalhado.Data} >= Date('" & Format(msk_fltInicio.Value, "Short Date") & "') And {Compras_relatorios_historico_detalhado.Data} <= Date('" & Format(msk_fltFim.Value, "Short Date") & "')"

                Case "Mês":
                    MesDe = FunVerificaMes(Cmb_mes_de)
                    MesAte = FunVerificaMes(Cmb_mes_ate)
                    StrSql = "select * from Compras_relatorios_historico_detalhado where " & CampoFiltro & " = '" & Filtro & Filtro2 & "' and Ano >= " & Cmb_ano_de & "  and Ano <= " & Cmb_ano_ate & " and Mes >= " & MesDe & "   and  Mes <= " & MesAte & " order by Ano, Mes"
                 FormulaRelatorio = "{Compras_relatorios_historico_detalhado." & CampoFiltro & "} = '" & Filtro & Filtro2Rel & "' and  {Compras_relatorios_historico_detalhado.Mes} >= " & qt & " and {Compras_relatorios_historico_detalhado.Mes} <= " & Qtd & " And {Compras_relatorios_historico_detalhado.ano} >= " & Cmb_ano_de & " And {Compras_relatorios_historico_detalhado.Ano} <= " & Cmb_ano_ate & ""
                Case "Ano":
                StrSql = "select * from Compras_relatorios_historico_detalhado where " & CampoFiltro & " = '" & Filtro & Filtro2 & "' and Ano >= " & Cmb_ano_de1 & "  and Ano <= " & Cmb_ano_ate1 & "   order by Ano, Mes"
                 FormulaRelatorio = "{Compras_relatorios_historico_detalhado." & CampoFiltro & "} = '" & Filtro & "'" & Filtro2Rel & "' and  {Compras_relatorios_historico_detalhado.ano} >= " & Cmb_ano_de1 & " And {Compras_relatorios_historico_detalhado.Ano} <= " & Cmb_ano_ate1 & ""

        End Select
        
If StrSql <> "" Then
    'Debug.print StrSql
    'Debug.print FormulaRelatorio
    ProcCarregaListaDetalhada
    End If
Else
    ListaDetalhada.ListItems.Clear
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1:
    
    If chkTexto.Value = 1 Then
        ProcFiltrarTexto
    Else
        ProcFiltrarCombo
    End If
    
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
