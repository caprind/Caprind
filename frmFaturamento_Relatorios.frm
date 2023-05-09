VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFaturamento_Relatorios 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Administrativo - Faturamento - Relatórios - Histórico"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbSerie 
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      ItemData        =   "frmFaturamento_Relatorios.frx":0000
      Left            =   150
      List            =   "frmFaturamento_Relatorios.frx":0025
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   113
      ToolTipText     =   "Série da nota fiscal para filtro"
      Top             =   1200
      Width           =   675
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2880
      TabIndex        =   66
      Top             =   990
      Width           =   2385
      Begin VB.OptionButton optTodas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todas"
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
         Height          =   210
         Left            =   1620
         TabIndex        =   114
         Top             =   300
         Width           =   735
      End
      Begin VB.OptionButton optEntrada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   750
         TabIndex        =   18
         Top             =   300
         Width           =   885
      End
      Begin VB.OptionButton optSaida 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saída"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   60
         TabIndex        =   17
         Top             =   300
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Emissão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   930
      TabIndex        =   67
      Top             =   990
      Width           =   1935
      Begin VB.OptionButton Opt_propria 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Própria"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1035
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.OptionButton Opt_terceiros 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Terceiros"
         DisabledPicture =   "frmFaturamento_Relatorios.frx":00CA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   60
         TabIndex        =   15
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Série"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   585
      Left            =   60
      TabIndex        =   112
      Top             =   990
      Width           =   855
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
   Begin VB.OptionButton Opt_dt_emissao 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dt. emissão"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   14250
      TabIndex        =   10
      Top             =   1770
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.OptionButton Opt_dt_entrada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dt. entrada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   165
      Left            =   13260
      TabIndex        =   9
      Top             =   1770
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Outros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   11490
      TabIndex        =   110
      Top             =   990
      Width           =   3855
      Begin VB.CheckBox Chk_retorno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Somar ret."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2730
         TabIndex        =   27
         Top             =   300
         Width           =   1095
      End
      Begin VB.CheckBox Chk_remessa 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Somar rem."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1620
         TabIndex        =   26
         Top             =   300
         Width           =   1275
      End
      Begin VB.CheckBox Chk_apenas_ST 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Apenas NF c/ ST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   25
         Top             =   300
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5280
      TabIndex        =   109
      Top             =   990
      Width           =   1875
      Begin VB.CheckBox Chk_ativa 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ativa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   300
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.CheckBox chkCanceladas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelada"
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
         Height          =   195
         Left            =   780
         TabIndex        =   20
         Top             =   300
         Value           =   1  'Checked
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Natureza de operação"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7170
      TabIndex        =   108
      Top             =   990
      Width           =   4305
      Begin VB.CheckBox chkMaoObra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mão de obra"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   900
         TabIndex        =   22
         Top             =   300
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkOutras 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Outras"
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
         Height          =   195
         Left            =   3450
         TabIndex        =   24
         Top             =   300
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox Chk_demonstracao 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Demonstração"
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
         Height          =   195
         Left            =   2100
         TabIndex        =   23
         Top             =   300
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chkVendas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vendas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   300
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1485
      Left            =   60
      TabIndex        =   73
      Top             =   6780
      Width           =   15285
      Begin VB.TextBox TxtValorRetorno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   4335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "0,000"
         ToolTipText     =   "Valor total de retorno."
         Top             =   450
         Width           =   1380
      End
      Begin VB.TextBox TxtValorRemessa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2910
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   30
         TabStop         =   0   'False
         Text            =   "0,000"
         ToolTipText     =   "Valor total de remessa."
         Top             =   450
         Width           =   1380
      End
      Begin VB.TextBox txtValorDesc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   13215
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do desconto."
         Top             =   1035
         Width           =   1800
      End
      Begin VB.TextBox txtVlrICMS_SN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   9060
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do ICMS simples nacional."
         Top             =   450
         Width           =   1680
      End
      Begin VB.TextBox txtVlrICMS_subst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   7155
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor do ICMS substituto."
         Top             =   450
         Width           =   1860
      End
      Begin VB.TextBox txtValorSeguro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   9490
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do seguro."
         Top             =   1035
         Width           =   1800
      End
      Begin VB.TextBox txtValorOutras 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   11352
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de outras despesas acessórias."
         Top             =   1035
         Width           =   1800
      End
      Begin VB.TextBox txtValorICMS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do ICMS."
         Top             =   450
         Width           =   1350
      End
      Begin VB.TextBox txtValorIPI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   10785
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do IPI."
         Top             =   450
         Width           =   1260
      End
      Begin VB.TextBox txtValorProdutos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   1485
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total dos produtos."
         Top             =   450
         Width           =   1380
      End
      Begin VB.TextBox txtQtdeProdutos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "0,000"
         ToolTipText     =   "Quantidade total de produtos."
         Top             =   450
         Width           =   1260
      End
      Begin VB.TextBox txtValorIRPJ_prod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2042
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do IRPJ."
         Top             =   1035
         Width           =   1800
      End
      Begin VB.TextBox txtValorPIS_prod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   12090
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do PIS."
         Top             =   450
         Width           =   1380
      End
      Begin VB.TextBox txtValorCofins_prod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   13515
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do Cofins."
         Top             =   450
         Width           =   1500
      End
      Begin VB.TextBox txtValorCSLL_prod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do CSLL."
         Top             =   1035
         Width           =   1800
      End
      Begin VB.TextBox txtValor_retencao_PIS_prod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   3904
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de retenção do PIS."
         Top             =   1035
         Width           =   1800
      End
      Begin VB.TextBox txtValor_retencao_Cofins_prod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   5766
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de retenção do Cofins."
         Top             =   1035
         Width           =   1800
      End
      Begin VB.TextBox txtValorFrete 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   7628
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do frete."
         Top             =   1035
         Width           =   1800
      End
      Begin VB.Label Label34 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total ret."
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
         Left            =   4485
         TabIndex        =   106
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label33 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total rem."
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
         Left            =   3015
         TabIndex        =   105
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label32 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total desc."
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
         Left            =   13515
         TabIndex        =   104
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total ICMS SN"
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
         Left            =   9165
         TabIndex        =   103
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total ICMS subst."
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
         Left            =   7200
         TabIndex        =   102
         Top             =   240
         Width           =   1530
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total outras desp."
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
         Left            =   11325
         TabIndex        =   101
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total seguro"
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
         Left            =   9705
         TabIndex        =   100
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total"
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
         Left            =   1815
         TabIndex        =   84
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label25 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total IPI"
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
         Index           =   0
         Left            =   10905
         TabIndex        =   83
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total ICMS"
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
         Left            =   5835
         TabIndex        =   82
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label30 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total"
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
         Left            =   360
         TabIndex        =   81
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label36 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total IRPJ"
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
         Left            =   2355
         TabIndex        =   80
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total PIS"
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
         Left            =   12255
         TabIndex        =   79
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total Cofins"
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
         Left            =   13620
         TabIndex        =   78
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total CSLL"
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
         Left            =   495
         TabIndex        =   77
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total ret. PIS"
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
         Left            =   4095
         TabIndex        =   76
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total ret. Cofins"
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
         Left            =   5850
         TabIndex        =   75
         Top             =   840
         Width           =   1425
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total frete"
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
         Left            =   7935
         TabIndex        =   74
         Top             =   840
         Width           =   1020
      End
   End
   Begin DrawSuite2022.USImageList USImageList1 
      Left            =   7140
      Top             =   210
      _ExtentX        =   900
      _ExtentY        =   767
      Img1            =   "frmFaturamento_Relatorios.frx":021C
      Count           =   1
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1425
      Left            =   13140
      TabIndex        =   58
      Top             =   1485
      Width           =   2205
      Begin MSComCtl2.DTPicker msk_fltFim 
         Height          =   315
         Left            =   690
         TabIndex        =   12
         ToolTipText     =   "Data final."
         Top             =   960
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Format          =   196411393
         CurrentDate     =   39057
      End
      Begin MSComCtl2.DTPicker msk_fltInicio 
         Height          =   315
         Left            =   690
         TabIndex        =   11
         ToolTipText     =   "Data inicio."
         Top             =   600
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Format          =   196411393
         CurrentDate     =   39057
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "De :"
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
         TabIndex        =   60
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Até :"
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
         Left            =   240
         TabIndex        =   59
         Top             =   960
         Width           =   360
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   1770
      TabIndex        =   61
      Top             =   1485
      Width           =   11355
      Begin VB.ComboBox Cmb_empresa 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmFaturamento_Relatorios.frx":3010
         Left            =   180
         List            =   "frmFaturamento_Relatorios.frx":3012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Empresa."
         Top             =   390
         Width           =   6000
      End
      Begin VB.CheckBox chkVlrTotal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vlr. total faturado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9720
         TabIndex        =   8
         Top             =   1080
         Width           =   1605
      End
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
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Limite de registros para carregar na lista."
         Top             =   990
         Width           =   555
      End
      Begin VB.ComboBox cmbTexto 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmFaturamento_Relatorios.frx":3014
         Left            =   180
         List            =   "frmFaturamento_Relatorios.frx":3016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Texto para pesquisa."
         Top             =   960
         Width           =   6015
      End
      Begin VB.ComboBox cmbfiltrarpor 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         ItemData        =   "frmFaturamento_Relatorios.frx":3018
         Left            =   6750
         List            =   "frmFaturamento_Relatorios.frx":303D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Opções para filtro."
         Top             =   390
         Width           =   4425
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "reg."
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
         Left            =   7905
         TabIndex        =   111
         Top             =   1050
         Width           =   300
      End
      Begin VB.Label Label44 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Left            =   2873
         TabIndex        =   107
         Top             =   180
         Width           =   615
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordenar por :"
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
         Left            =   8460
         TabIndex        =   69
         Top             =   1050
         Width           =   990
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limitar em"
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
         Left            =   6270
         TabIndex        =   68
         Top             =   1050
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar por"
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
         Left            =   8535
         TabIndex        =   63
         Top             =   180
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Texto para pesquisa"
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
         Left            =   2452
         TabIndex        =   62
         Top             =   750
         Width           =   1470
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   65
      Top             =   1485
      Width           =   1695
      Begin VB.OptionButton Opt_individual 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Individual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   0
         Top             =   180
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton Opt_comparativo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparativo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   450
         Width           =   1425
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   60
      TabIndex        =   64
      Top             =   2145
      Width           =   1695
      Begin VB.OptionButton optDetalhado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Detalhado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optResumido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resumido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   450
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView ListaNF 
      Height          =   3870
      Left            =   60
      TabIndex        =   13
      Top             =   2895
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   6826
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   40
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Empresa"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Object.Tag             =   "D"
         Text            =   "Dt. emissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Nota fiscal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Object.Tag             =   "T"
         Text            =   "Tipo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "CFOP"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Object.Tag             =   "T"
         Text            =   "Natureza operação"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Object.Tag             =   "T"
         Text            =   "Destinatário/Emitente"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Vlr. total faturado"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Tag             =   "T"
         Text            =   "Cód. interno"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "T"
         Text            =   "Descrição"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "T"
         Text            =   "CF"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Qtde. total produtos"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Vlr. total produtos"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Base de cálculo ICMS"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "Vlr. ICMS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Tag             =   "N"
         Text            =   "Base de calculo ICMS subst."
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Tag             =   "N"
         Text            =   "Vlr. ICMS subst."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Tag             =   "N"
         Text            =   "Vlr. ICMS SN"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Tag             =   "N"
         Text            =   "Vlr. total IPI"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Tag             =   "N"
         Text            =   "Vlr. total PIS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Tag             =   "N"
         Text            =   "Vlr. total Cofins"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Object.Tag             =   "N"
         Text            =   "Vlr. total CSLL"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Tag             =   "N"
         Text            =   "Vlr. total IRPJ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Object.Tag             =   "N"
         Text            =   "Vlr. total ret. PIS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Object.Tag             =   "N"
         Text            =   "Vlr. total ret. Cofins"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Object.Tag             =   "N"
         Text            =   "Qtde. total serviços"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Object.Tag             =   "N"
         Text            =   "Vlr. total serviços"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Object.Tag             =   "N"
         Text            =   "Vlr. total PIS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Object.Tag             =   "N"
         Text            =   "Vlr. total Cofins"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Object.Tag             =   "N"
         Text            =   "Vlr. total CSLL"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Object.Tag             =   "N"
         Text            =   "Vlr. total ISSQN"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   32
         Object.Tag             =   "N"
         Text            =   "Vlr. total INSS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   33
         Object.Tag             =   "N"
         Text            =   "Vlr. total IRPJ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   34
         Object.Tag             =   "N"
         Text            =   "Vlr. total IRRF"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   35
         Object.Tag             =   "N"
         Text            =   "Vlr. total frete"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   36
         Object.Tag             =   "N"
         Text            =   "Vlr. total seguro"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   37
         Object.Tag             =   "N"
         Text            =   "Vlr. total outras desp."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   38
         Text            =   "Vlr. total desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   39
         Object.Tag             =   "N"
         Text            =   "Vlr. total DAS"
         Object.Width           =   2646
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   71
      Top             =   9750
      Width           =   11625
      _ExtentX        =   20505
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
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   70
      Top             =   0
      Width           =   15285
      _ExtentX        =   26961
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   42
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft2     =   46
      ButtonTop2      =   2
      ButtonWidth2    =   60
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
      ButtonLeft3     =   108
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   112
      ButtonTop4      =   2
      ButtonWidth4    =   41
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft5     =   155
      ButtonTop5      =   2
      ButtonWidth5    =   30
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
      ButtonLeft6     =   187
      ButtonTop6      =   2
      ButtonWidth6    =   24
      ButtonHeight6   =   24
   End
   Begin MSComctlLib.ListView ListaNF1 
      Height          =   3870
      Left            =   60
      TabIndex        =   14
      Top             =   2895
      Visible         =   0   'False
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   6826
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   32
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Destinatário"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Destinatário"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "N"
         Text            =   "Vlr. total faturado"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "N"
         Text            =   "Qtde. total produtos"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "N"
         Text            =   "Vlr. total produtos"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "N"
         Text            =   "Base de cálculo ICMS"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "N"
         Text            =   "Vlr. ICMS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "N"
         Text            =   "Base de calculo ICMS subst."
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "N"
         Text            =   "Vlr. ICMS subst."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Object.Tag             =   "N"
         Text            =   "Vlr. ICMS SN"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Object.Tag             =   "N"
         Text            =   "Vlr. total IPI"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "N"
         Text            =   "Vlr. total PIS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "N"
         Text            =   "Vlr. total Cofins"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "N"
         Text            =   "Vlr. total CSLL"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "N"
         Text            =   "Vlr. total IRPJ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Tag             =   "N"
         Text            =   "Vlr. total ret. PIS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Tag             =   "N"
         Text            =   "Vlr. total ret. Cofins"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Tag             =   "N"
         Text            =   "Qtde. total serviços"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Tag             =   "N"
         Text            =   "Vlr. total serviços"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Tag             =   "N"
         Text            =   "Vlr. total PIS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Tag             =   "N"
         Text            =   "Vlr. total Cofins"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Object.Tag             =   "N"
         Text            =   "Vlr. total CSLL"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Tag             =   "N"
         Text            =   "Vlr. total ISSQN"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Object.Tag             =   "N"
         Text            =   "Vlr. total INSS"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Object.Tag             =   "N"
         Text            =   "Vlr. total IRPJ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   26
         Object.Tag             =   "N"
         Text            =   "Vlr. total IRRF"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Object.Tag             =   "N"
         Text            =   "Vlr. total frete"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   28
         Object.Tag             =   "N"
         Text            =   "Vlr. total seguro"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   29
         Object.Tag             =   "N"
         Text            =   "Vlr. total outras desp."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   30
         Object.Tag             =   "N"
         Text            =   "Vlr. total desc."
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   31
         Object.Tag             =   "N"
         Text            =   "Vlr. total DAS"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Serviços"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   60
      TabIndex        =   85
      Top             =   8250
      Width           =   11205
      Begin VB.TextBox txtQtdeServicos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "0,000"
         ToolTipText     =   "Quantidade total de serviços."
         Top             =   450
         Width           =   2100
      End
      Begin VB.TextBox txtValorServicos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2362
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total de serviços."
         Top             =   450
         Width           =   2100
      End
      Begin VB.TextBox txtValorISSQN_serv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do ISSQN."
         Top             =   1035
         Width           =   2670
      End
      Begin VB.TextBox txtValorINSS_serv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   2900
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do INSS."
         Top             =   1035
         Width           =   2670
      End
      Begin VB.TextBox txtValorIRPJ_serv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   5620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do IRPJ."
         Top             =   1035
         Width           =   2670
      End
      Begin VB.TextBox txtValorPIS_serv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   4544
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do PIS."
         Top             =   450
         Width           =   2100
      End
      Begin VB.TextBox txtValorCofins_serv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   6726
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do Cofins."
         Top             =   450
         Width           =   2100
      End
      Begin VB.TextBox txtValorCSLL_serv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do CSLL."
         Top             =   450
         Width           =   2100
      End
      Begin VB.TextBox txtValorIRRF_serv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   8340
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do IRRF."
         Top             =   1035
         Width           =   2670
      End
      Begin VB.Label Label31 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde. total"
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
         Left            =   780
         TabIndex        =   94
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label24 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total"
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
         Left            =   3045
         TabIndex        =   93
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total INSS"
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
         Left            =   3645
         TabIndex        =   92
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label38 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total ISSQN"
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
         Left            =   870
         TabIndex        =   91
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total IRPJ"
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
         Left            =   6360
         TabIndex        =   90
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total PIS"
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
         Left            =   5070
         TabIndex        =   89
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total Cofins"
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
         Left            =   7125
         TabIndex        =   88
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total CSLL"
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
         Left            =   9375
         TabIndex        =   87
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total IRRF"
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
         Left            =   9090
         TabIndex        =   86
         Top             =   840
         Width           =   1020
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Simples nacional"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   11280
      TabIndex        =   95
      Top             =   8250
      Width           =   1830
      Begin VB.TextBox txtValorDAS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   210
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total do DAS."
         Top             =   1035
         Width           =   1425
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor total DAS"
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
         Left            =   270
         TabIndex        =   96
         Top             =   840
         Width           =   1290
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totais"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   13125
      TabIndex        =   97
      Top             =   8250
      Width           =   2220
      Begin VB.TextBox txtValorTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "0,00"
         ToolTipText     =   "Valor total faturado."
         Top             =   450
         Width           =   1725
      End
      Begin VB.TextBox txtPercentual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "0"
         ToolTipText     =   "Percentual."
         Top             =   1050
         Width           =   1725
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Vlr. total faturado"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   285
         TabIndex        =   99
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentual (%)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   375
         TabIndex        =   98
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Label Lbl_relatorio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registros encontrados: 0000 - 00:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   11820
      TabIndex        =   72
      Top             =   9780
      Width           =   3315
   End
End
Attribute VB_Name = "frmFaturamento_Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub ProcAjuda()
On Error GoTo tratar_erro

If Faturamento = True Then FunAbrirVideoWeb ("http://www.youtube.com/watch?v=firkEb2mBBE&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=27&feature=plcp") Else FunAbrirVideoWeb ("http://www.youtube.com/watch?v=kYpW0IhHRoI&list=UUODGiDjQ-BCrxh0YXfJvoqg&index=29&feature=plcp")

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub


Private Sub ProcCarregaSerie()
On Error GoTo tratar_erro

With cmbSerie
    .Clear
    .AddItem ""
    .AddItem "0"
    .AddItem "1"
    .AddItem "2"
    .AddItem "3"
    .AddItem "4"
    .AddItem "5"
    .AddItem "6"
    .AddItem "7"
    .AddItem "8"
    .AddItem "9"
    .Refresh
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_ativa_Click()
On Error GoTo tratar_erro

'ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_apenas_ST_Click()
On Error GoTo tratar_erro

'ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_demonstracao_Click()
On Error GoTo tratar_erro

'ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_remessa_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Chk_retorno_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkCanceladas_Click()
On Error GoTo tratar_erro

'ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkMaoObra_Click()
On Error GoTo tratar_erro

'ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkOutras_Click()
On Error GoTo tratar_erro

'ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkVendas_Click()
On Error GoTo tratar_erro

'ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub chkVlrTotal_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Cmb_empresa_Click()
On Error GoTo tratar_erro

'ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbTexto_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcImprimir()
On Error GoTo tratar_erro

If ListaNF.ListItems.Count = 0 And ListaNF1.ListItems.Count = 0 Then Exit Sub
frmFaturamento_Relatorios_menuimpressao.Show 1

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

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
    Case vbKeyF2: ProcAbrir
    Case vbKeyF5: ProcImprimir
    Case vbKeyF1: ProcAjuda
    Case vbKeyEscape: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcCarregaLista()
On Error GoTo tratar_erro

Posicao = 0
ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
If TBLISTA.EOF = False Then
    
    Posicao = TBLISTA.RecordCount
    
    PBLista.Min = 0
    PBLista.Max = TBLISTA.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLISTA.EOF = False
        If optDetalhado.Value = True Then
            With ListaNF.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!Data3), "", TBLISTA!Data3) 'Empresa
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!Data), "", Format(TBLISTA!Data, "dd/mm/yy")) 'data emissão
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Nota), "", TBLISTA!Nota) 'Nota
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!Execucaoprev), "", TBLISTA!Execucaoprev) 'Tipo da nota
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!Totalhsprev), "", TBLISTA!Totalhsprev) 'CFOP
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Totalhsutil), "", TBLISTA!Totalhsutil) 'Natureza de operação
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!DescEvento), "", TBLISTA!DescEvento) 'Cliente
                
                'Valor total faturado
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Numero15), "0,00", Format(TBLISTA!Numero15, "###,##0.00"))
                
                'Produtos
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!Data1), "", TBLISTA!Data1) 'Código interno
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Data2), "", TBLISTA!Data2) 'Descrição
                .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Data4), "", TBLISTA!Data4) 'Descrição
                .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!qtdeOK), "0,0000", Format(TBLISTA!qtdeOK, "###,##0.0000")) 'Quantidade de produtos
                .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!qtdeNC), "0,00", Format(TBLISTA!qtdeNC, "###,##0.00")) 'Valor total de produtos
                .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!Qtdetotalprod), "0,00", Format(TBLISTA!Qtdetotalprod, "###,##0.00")) 'Base de calculo de icms
                .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!Eficiencia), "0,00", Format(TBLISTA!Eficiencia, "###,##0.00")) 'Valor do ICMS
                .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA!Terceiros), "0,00", Format(TBLISTA!Terceiros, "###,##0.00")) 'Base de calculo ICMS subst.
                .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA!impostos), "0,00", Format(TBLISTA!impostos, "###,##0.00")) 'Valor do ICMS subst.
                .Item(.Count).SubItems(18) = IIf(IsNull(TBLISTA!Valor5), "0,00", Format(TBLISTA!Valor5, "###,##0.00")) 'Valor do ICMS SN
                .Item(.Count).SubItems(19) = IIf(IsNull(TBLISTA!Lucro), "0,00", Format(TBLISTA!Lucro, "###,##0.00")) 'Valor total de ipi
                .Item(.Count).SubItems(20) = IIf(IsNull(TBLISTA!material), "0,00", Format(TBLISTA!material, "###,##0.00")) 'Valor total pis
                .Item(.Count).SubItems(21) = IIf(IsNull(TBLISTA!Servicos), "0,00", Format(TBLISTA!Servicos, "###,##0.00")) 'Valor total cofins
                .Item(.Count).SubItems(22) = IIf(IsNull(TBLISTA!Total), "0,00", Format(TBLISTA!Total, "###,##0.00")) 'Total CSLL
                .Item(.Count).SubItems(23) = IIf(IsNull(TBLISTA!Total_peca), "0,00", Format(TBLISTA!Total_peca, "###,##0.00")) 'Total IRPJ
                .Item(.Count).SubItems(24) = IIf(IsNull(TBLISTA!Refugo), "0,00", Format(TBLISTA!Refugo, "###,##0.00")) 'Total Pis
                .Item(.Count).SubItems(25) = IIf(IsNull(TBLISTA!Numero1), "0,00", Format(TBLISTA!Numero1, "###,##0.00")) 'Total retenção confins
                'Serviço
                .Item(.Count).SubItems(26) = IIf(IsNull(TBLISTA!Numero2), "0,0000", Format(TBLISTA!Numero2, "###,##0.0000")) 'Quantidade de serviços
                .Item(.Count).SubItems(27) = IIf(IsNull(TBLISTA!Numero3), "0,00", Format(TBLISTA!Numero3, "###,##0.00")) 'Valor total serviço
                .Item(.Count).SubItems(28) = IIf(IsNull(TBLISTA!Numero4), "0,00", Format(TBLISTA!Numero4, "###,##0.00")) 'Total pis serviços
                .Item(.Count).SubItems(29) = IIf(IsNull(TBLISTA!Numero5), "0,00", Format(TBLISTA!Numero5, "###,##0.00")) 'Total cofins serviços
                .Item(.Count).SubItems(30) = IIf(IsNull(TBLISTA!Numero6), "0,00", Format(TBLISTA!Numero6, "###,##0.00")) 'Total CSLL serviços
                .Item(.Count).SubItems(31) = IIf(IsNull(TBLISTA!Numero7), "0,00", Format(TBLISTA!Numero7, "###,##0.00")) 'Valor total de ISS
                .Item(.Count).SubItems(32) = IIf(IsNull(TBLISTA!Numero8), "0,00", Format(TBLISTA!Numero8, "###,##0.00")) 'Total de INSS serviços
                .Item(.Count).SubItems(33) = IIf(IsNull(TBLISTA!Numero9), "0,00", Format(TBLISTA!Numero9, "###,##0.00")) 'Total de IRPJ serviços
                .Item(.Count).SubItems(34) = IIf(IsNull(TBLISTA!Numero10), "0,00", Format(TBLISTA!Numero10, "###,##0.00")) 'Total de IRRF serviços
                'Totais
                .Item(.Count).SubItems(35) = IIf(IsNull(TBLISTA!Numero11), "0,00", Format(TBLISTA!Numero11, "###,##0.00")) 'Valor Frete
                .Item(.Count).SubItems(36) = IIf(IsNull(TBLISTA!Numero12), "0,00", Format(TBLISTA!Numero12, "###,##0.00")) 'Valor Seguro
                .Item(.Count).SubItems(37) = IIf(IsNull(TBLISTA!Numero13), "0,00", Format(TBLISTA!Numero13, "###,##0.00")) 'Despeças adicionais
                .Item(.Count).SubItems(38) = IIf(IsNull(TBLISTA!Valor6), "0,00", Format(TBLISTA!Valor6, "###,##0.00")) 'Desconto
                .Item(.Count).SubItems(39) = IIf(IsNull(TBLISTA!Numero14), "0,00", Format(TBLISTA!Numero14, "###,##0.00")) 'Total das
            End With
        Else
            With ListaNF1.ListItems
                .Add , , TBLISTA!ID
                .Item(.Count).SubItems(1) = IIf(IsNull(TBLISTA!maquina), "", TBLISTA!maquina)
                .Item(.Count).SubItems(2) = IIf(IsNull(TBLISTA!DescEvento), "", TBLISTA!DescEvento)
                
                'Valor total faturado
                .Item(.Count).SubItems(3) = IIf(IsNull(TBLISTA!Numero15), "0,00", Format(TBLISTA!Numero15, "###,##0.00"))
                
                'Produtos
                .Item(.Count).SubItems(4) = IIf(IsNull(TBLISTA!qtdeOK), "0,0000", Format(TBLISTA!qtdeOK, "###,##0.0000")) 'Quantidade de produtos
                .Item(.Count).SubItems(5) = IIf(IsNull(TBLISTA!qtdeNC), "0,00", Format(TBLISTA!qtdeNC, "###,##0.00")) 'Valor total de produtos
                .Item(.Count).SubItems(6) = IIf(IsNull(TBLISTA!Qtdetotalprod), "0,00", Format(TBLISTA!Qtdetotalprod, "###,##0.00")) 'Base de calculo de icms
                .Item(.Count).SubItems(7) = IIf(IsNull(TBLISTA!Eficiencia), "0,00", Format(TBLISTA!Eficiencia, "###,##0.00")) 'Valor do ICMS
                .Item(.Count).SubItems(8) = IIf(IsNull(TBLISTA!Terceiros), "0,00", Format(TBLISTA!Terceiros, "###,##0.00")) 'Base de calculo ICMS subst.
                .Item(.Count).SubItems(9) = IIf(IsNull(TBLISTA!impostos), "0,00", Format(TBLISTA!impostos, "###,##0.00")) 'Valor do ICMS subst.
                .Item(.Count).SubItems(10) = IIf(IsNull(TBLISTA!Valor5), "0,00", Format(TBLISTA!Valor5, "###,##0.00")) 'Valor do ICMS SN
                .Item(.Count).SubItems(11) = IIf(IsNull(TBLISTA!Lucro), "0,00", Format(TBLISTA!Lucro, "###,##0.00")) 'Valor total de ipi
                .Item(.Count).SubItems(12) = IIf(IsNull(TBLISTA!material), "0,00", Format(TBLISTA!material, "###,##0.00")) 'Valor total pis
                .Item(.Count).SubItems(13) = IIf(IsNull(TBLISTA!Servicos), "0,00", Format(TBLISTA!Servicos, "###,##0.00")) 'Valor total cofins
                .Item(.Count).SubItems(14) = IIf(IsNull(TBLISTA!Total), "0,00", Format(TBLISTA!Total, "###,##0.00")) 'Total CSLL
                .Item(.Count).SubItems(15) = IIf(IsNull(TBLISTA!Total_peca), "0,00", Format(TBLISTA!Total_peca, "###,##0.00")) 'Total IRPJ
                .Item(.Count).SubItems(16) = IIf(IsNull(TBLISTA!Refugo), "0,00", Format(TBLISTA!Refugo, "###,##0.00")) 'Total Pis
                .Item(.Count).SubItems(17) = IIf(IsNull(TBLISTA!Numero1), "0,00", Format(TBLISTA!Numero1, "###,##0.00")) 'Total retenção confins
                'Serviço
                .Item(.Count).SubItems(18) = IIf(IsNull(TBLISTA!Numero2), "0,0000", Format(TBLISTA!Numero2, "###,##0.0000")) 'Quantidade de serviços
                .Item(.Count).SubItems(19) = IIf(IsNull(TBLISTA!Numero3), "0,00", Format(TBLISTA!Numero3, "###,##0.00")) 'Valor total serviço
                .Item(.Count).SubItems(20) = IIf(IsNull(TBLISTA!Numero4), "0,00", Format(TBLISTA!Numero4, "###,##0.00")) 'Total pis serviços
                .Item(.Count).SubItems(21) = IIf(IsNull(TBLISTA!Numero5), "0,00", Format(TBLISTA!Numero5, "###,##0.00")) 'Total cofins serviços
                .Item(.Count).SubItems(22) = IIf(IsNull(TBLISTA!Numero6), "0,00", Format(TBLISTA!Numero6, "###,##0.00")) 'Total CSLL serviços
                .Item(.Count).SubItems(23) = IIf(IsNull(TBLISTA!Numero7), "0,00", Format(TBLISTA!Numero7, "###,##0.00")) 'Valor total de ISS
                .Item(.Count).SubItems(24) = IIf(IsNull(TBLISTA!Numero8), "0,00", Format(TBLISTA!Numero8, "###,##0.00")) 'Total de INSS serviços
                .Item(.Count).SubItems(25) = IIf(IsNull(TBLISTA!Numero9), "0,00", Format(TBLISTA!Numero9, "###,##0.00")) 'Total de IRPJ serviços
                .Item(.Count).SubItems(26) = IIf(IsNull(TBLISTA!Numero10), "0,00", Format(TBLISTA!Numero10, "###,##0.00")) 'Total de IRRF serviços
                'Totais
                .Item(.Count).SubItems(27) = IIf(IsNull(TBLISTA!Numero11), "0,00", Format(TBLISTA!Numero11, "###,##0.00")) 'Valor Frete
                .Item(.Count).SubItems(28) = IIf(IsNull(TBLISTA!Numero12), "0,00", Format(TBLISTA!Numero12, "###,##0.00")) 'Valor Seguro
                .Item(.Count).SubItems(29) = IIf(IsNull(TBLISTA!Numero13), "0,00", Format(TBLISTA!Numero13, "###,##0.00")) 'Despeças adicionais
                .Item(.Count).SubItems(30) = IIf(IsNull(TBLISTA!Valor6), "0,00", Format(TBLISTA!Valor6, "###,##0.00")) 'Desconto
                .Item(.Count).SubItems(31) = IIf(IsNull(TBLISTA!Numero14), "0,00", Format(TBLISTA!Numero14, "###,##0.00")) 'Total das
            End With
        End If
        TBLISTA.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    'Produtos
    txtQtdeProdutos = IIf(IsNull(TBLISTA!QtdePrevista), "0,0000", Format(TBLISTA!QtdePrevista, "###,##0.0000"))
    txtValorProdutos = IIf(IsNull(TBLISTA!QtdeProduzida), "0,00", Format(TBLISTA!QtdeProduzida, "###,##0.00"))
    TxtValorRemessa = IIf(IsNull(TBLISTA!Numero14), "0,00", Format(TBLISTA!Numero14, "###,##0.00"))
    TxtValorRetorno = IIf(IsNull(TBLISTA!Numero15), "0,00", Format(TBLISTA!Numero15, "###,##0.00"))
    txtValorICMS = IIf(IsNull(TBLISTA!qtdeNC), "0,00", Format(TBLISTA!qtdeNC, "###,##0.00"))
    txtVlrICMS_subst = IIf(IsNull(TBLISTA!Numero13), "0,00", Format(TBLISTA!Numero13, "###,##0.00"))
    txtVlrICMS_SN = IIf(IsNull(TBLISTA!Numero16), "0,00", Format(TBLISTA!Numero16, "###,##0.00"))
    txtValorIPI = IIf(IsNull(TBLISTA!QtdeOrdem), "0,00", Format(TBLISTA!QtdeOrdem, "###,##0.00"))
    txtValorPIS_prod = IIf(IsNull(TBLISTA!CustoMat), "0,00", Format(TBLISTA!CustoMat, "###,##0.00"))
    txtValorCofins_prod = IIf(IsNull(TBLISTA!CustoObra), "0,00", Format(TBLISTA!CustoObra, "###,##0.00"))
    txtValorCSLL_prod = IIf(IsNull(TBLISTA!Terceros), "0,00", Format(TBLISTA!Terceros, "###,##0.00"))
    txtValorIRPJ_prod = IIf(IsNull(TBLISTA!Lucro), "0,00", Format(TBLISTA!Lucro, "###,##0.00"))
    txtValor_retencao_PIS_prod = IIf(IsNull(TBLISTA!Valor1), "0,00", Format(TBLISTA!Valor1, "###,##0.00"))
    txtValor_retencao_Cofins_prod = IIf(IsNull(TBLISTA!Valor2), "0,00", Format(TBLISTA!Valor2, "###,##0.00"))
    txtValorFrete = IIf(IsNull(TBLISTA!Numero7), "0,00", Format(TBLISTA!Numero7, "###,##0.00"))
    txtValorSeguro = IIf(IsNull(TBLISTA!Numero8), "0,00", Format(TBLISTA!Numero8, "###,##0.00"))
    txtValorOutras = IIf(IsNull(TBLISTA!Numero9), "0,00", Format(TBLISTA!Numero9, "###,##0.00"))
    txtValorDesc = IIf(IsNull(TBLISTA!Numero17), "0,00", Format(TBLISTA!Numero17, "###,##0.00"))
    
    'Serviços
    txtQtdeServicos = IIf(IsNull(TBLISTA!Valor3), "0,00", Format(TBLISTA!Valor3, "###,##0.0000"))
    txtValorServicos = IIf(IsNull(TBLISTA!Total1), "0,00", Format(TBLISTA!Total1, "###,##0.00"))
    txtValorPIS_serv = IIf(IsNull(TBLISTA!Total2), "0,00", Format(TBLISTA!Total2, "###,##0.00"))
    txtValorCofins_serv = IIf(IsNull(TBLISTA!Numero1), "0,00", Format(TBLISTA!Numero1, "###,##0.00"))
    txtValorCSLL_serv = IIf(IsNull(TBLISTA!Numero2), "0,00", Format(TBLISTA!Numero2, "###,##0.00"))
    txtValorISSQN_serv = IIf(IsNull(TBLISTA!Numero3), "0,00", Format(TBLISTA!Numero3, "###,##0.00"))
    txtValorINSS_serv = IIf(IsNull(TBLISTA!Numero4), "0,00", Format(TBLISTA!Numero4, "###,##0.00"))
    txtValorIRPJ_serv = IIf(IsNull(TBLISTA!Numero5), "0,00", Format(TBLISTA!Numero5, "###,##0.00"))
    txtValorIRRF_serv = IIf(IsNull(TBLISTA!Numero6), "0,00", Format(TBLISTA!Numero6, "###,##0.00"))
    
    'Total
    txtValorDAS = IIf(IsNull(TBLISTA!Numero10), "0,00", Format(TBLISTA!Numero10, "###,##0.00"))
    txtValorTotal = IIf(IsNull(TBLISTA!Numero11), "0,00", Format(TBLISTA!Numero11, "###,##0.00"))
    txtPercentual = IIf(IsNull(TBLISTA!Numero12), "0,00", Format(TBLISTA!Numero12, "###,##0.00"))
End If
TBLISTA.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Sub ProcLimpaCamposTotais()
On Error GoTo tratar_erro

Lbl_relatorio.Caption = "Registros encontrados: 0000 - 00:00:00"
txtQtdeProdutos = ""
txtValorProdutos = ""
TxtValorRemessa = ""
TxtValorRetorno = ""
txtValorICMS = ""
txtVlrICMS_subst = ""
txtVlrICMS_SN = ""
txtValorIPI = ""
txtValorPIS_prod = ""
txtValorCofins_prod = ""
txtValorCSLL_prod = ""
txtValorIRPJ_prod = ""
txtValor_retencao_PIS_prod = ""
txtValor_retencao_Cofins_prod = ""
txtValorFrete = ""
txtValorSeguro = ""
txtValorOutras = ""
txtValorDesc = ""

txtQtdeServicos = ""
txtValorServicos = ""
txtValorPIS_serv = ""
txtValorCofins_serv = ""
txtValorCSLL_serv = ""
txtValorISSQN_serv = ""
txtValorINSS_serv = ""
txtValorIRPJ_serv = ""
txtValorIRRF_serv = ""

txtValorDAS = ""
txtValorTotal = ""
txtPercentual = ""

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 16895, 6, True
ProcCarregaSerie
cmbSerie.Text = 1
If Faturamento = True Then
    Formulario = "Faturamento/Relatórios/Histórico"
    Caption = "Faturamento - Relatórios - Histórico"
Else
    Formulario = "Administrativo/Vendas/Informações de faturamento"
    Caption = "Administrativo - Vendas - Informações de faturamento"
End If
Direitos
ProcLimpaVariaveisPrincipais
ProcCarregaComboEmpresa Cmb_empresa, True
msk_fltInicio.Value = Date
msk_fltFim.Value = Date

cmbfiltrarpor = "Destinatário"

ProcRemoveObjetosResize Me


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaComboFiltrarPor(Terceiros As Boolean, Comparativo As Boolean)
On Error GoTo tratar_erro

With cmbfiltrarpor
    .Clear
    .AddItem "CFOP"
    .AddItem "Classificação fiscal"
    .AddItem "Código de referência"
    .AddItem "Código interno"
    .AddItem "Descrição"
    If Terceiros = True Then
        .AddItem "Emitente"
        If Comparativo = True Then .AddItem "Nota fiscal x Emitente"
    Else
        .AddItem "Destinatário"
        If Comparativo = True Then .AddItem "Nota fiscal x Destinatário"
    End If
    .AddItem "Família"
    .AddItem "Grupo"
    .AddItem "Nota fiscal"
    .AddItem "UF"
    
    .Text = "Nota fiscal"
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

If Faturamento = True Then Formulario = "Faturamento/Relatórios/Histórico" Else Formulario = "Administrativo/Vendas/Informações de faturamento"
Direitos
ProcLimpaVariaveisPrincipais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNF_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaNF, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ListaNF1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo tratar_erro

ProcOrdenaListView ListaNF1, ColumnHeader

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltFim_Change()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmbfiltrarpor_Click()
On Error GoTo tratar_erro

ProcCarregarComboTexto

With cmbTexto
    If Opt_comparativo.Value = True Then
        If Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then
            Label9.Caption = "Texto para pesquisa*"
            .Enabled = True
        Else
            Label9.Caption = "Texto para pesquisa"
            .Enabled = False
            .ListIndex = -1
        End If
    End If
End With

ProcAcertaColunas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAcertaColunas()
On Error GoTo tratar_erro

With ListaNF
    If optDetalhado = True Then .ColumnHeaders(2).Width = 2200 Else .ColumnHeaders(2).Width = 0
    .ColumnHeaders(9).Width = 1500
    .ColumnHeaders(10).Width = 1200
    .ColumnHeaders(11).Width = 2800
    .ColumnHeaders(12).Width = 1200
    .ColumnHeaders(13).Width = 1700
    .ColumnHeaders(14).Width = 1500
    .ColumnHeaders(15).Width = 1800
    .ColumnHeaders(16).Width = 1500
    .ColumnHeaders(17).Width = 2250
    .ColumnHeaders(18).Width = 1500
    .ColumnHeaders(19).Width = 1500
    .ColumnHeaders(20).Width = 1500
    .ColumnHeaders(21).Width = 1500
    .ColumnHeaders(22).Width = 1500
    .ColumnHeaders(23).Width = 1500
    .ColumnHeaders(24).Width = 1500
    .ColumnHeaders(25).Width = 1500
    .ColumnHeaders(26).Width = 1600
    .ColumnHeaders(27).Width = 1700
    .ColumnHeaders(28).Width = 1500
    .ColumnHeaders(29).Width = 1500
    .ColumnHeaders(30).Width = 1500
    .ColumnHeaders(31).Width = 1500
    .ColumnHeaders(32).Width = 1500
    .ColumnHeaders(33).Width = 1500
    .ColumnHeaders(34).Width = 1500
    .ColumnHeaders(35).Width = 1500
    .ColumnHeaders(36).Width = 1500
    .ColumnHeaders(37).Width = 1500
    .ColumnHeaders(38).Width = 2000
    .ColumnHeaders(39).Width = 1500
    .ColumnHeaders(40).Width = 1500
End With

With ListaNF1
    .ColumnHeaders(4).Width = 1500
    .ColumnHeaders(5).Width = 1700
    .ColumnHeaders(6).Width = 1500
    .ColumnHeaders(7).Width = 1800
    .ColumnHeaders(8).Width = 1500
    .ColumnHeaders(9).Width = 2250
    .ColumnHeaders(10).Width = 1500
    .ColumnHeaders(11).Width = 1500
    .ColumnHeaders(12).Width = 1500
    .ColumnHeaders(13).Width = 1500
    .ColumnHeaders(14).Width = 1500
    .ColumnHeaders(15).Width = 1500
    .ColumnHeaders(16).Width = 1500
    .ColumnHeaders(17).Width = 1500
    .ColumnHeaders(18).Width = 1600
    .ColumnHeaders(19).Width = 1700
    .ColumnHeaders(20).Width = 1500
    .ColumnHeaders(21).Width = 1500
    .ColumnHeaders(22).Width = 1500
    .ColumnHeaders(23).Width = 1500
    .ColumnHeaders(24).Width = 1500
    .ColumnHeaders(25).Width = 1500
    .ColumnHeaders(26).Width = 1500
    .ColumnHeaders(27).Width = 1500
    .ColumnHeaders(28).Width = 1500
    .ColumnHeaders(29).Width = 1500
    .ColumnHeaders(30).Width = 2000
    .ColumnHeaders(31).Width = 1500
    .ColumnHeaders(32).Width = 1500
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregarComboTexto()
On Error GoTo tratar_erro

If Cmb_empresa = "" Or cmbfiltrarpor = "" Then Exit Sub

With ListaNF1
    If Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then .ColumnHeaders(2).Text = "Nota fiscal" Else .ColumnHeaders(2).Text = cmbfiltrarpor
    If Opt_individual.Value = True Then .ColumnHeaders(2).Width = 2800
    If cmbfiltrarpor = "Nota fiscal" Then .ColumnHeaders(3).Width = 3200 Else .ColumnHeaders(3).Width = 0
End With

If Opt_individual.Value = True Or Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then
    Desenho = ""
    cmbTexto.Clear
    
    If Opt_terceiros.Value = True Then Aplicacao_NF = "T" Else Aplicacao_NF = "P"
    
    If optSaida.Value = True Then
    Sit_Nota = 1
    End If
    
    If optEntrada.Value = True Then
    Sit_Nota = 2
    End If
    
    If opttodas.Value = True Then
    Sit_Nota = 0
    End If

    
    If Chk_ativa.Value = 1 And chkCanceladas.Value = 1 Then
        StatusTexto = "NF.int_status IS NOT NULL"
    ElseIf Chk_ativa.Value = 1 Then
            StatusTexto = "NF.int_status = 1"
        Else
            StatusTexto = "NF.int_status = 2"
    End If
    
    If chkVendas.Value = 1 Then CFOPVendas = "CFOP.Vendas = 'True'" Else CFOPVendas = ""
    If chkMaoObra.Value = 1 Then
        If CFOPVendas = "" Then CFOPMO = "CFOP.MaoObra = 'True'" Else CFOPMO = " or CFOP.MaoObra = 'True'"
    Else
        CFOPMO = ""
    End If
    If Chk_demonstracao.Value = 1 Then
        If CFOPVendas = "" And CFOPMO = "" Then CFOPDEM = "CFOP.Demonstracao = 'True'" Else CFOPDEM = " or CFOP.Demonstracao = 'True'"
    Else
        CFOPDEM = ""
    End If
    If chkOutras.Value = 1 Then
        If CFOPVendas = "" And CFOPMO = "" And CFOPDEM = "" Then CFOPOutros = "CFOP.Vendas = 'False' and CFOP.MaoObra = 'False' and CFOP.Demonstracao = 'False'" Else CFOPOutros = " or (CFOP.Vendas = 'False' and CFOP.MaoObra = 'False' and CFOP.Demonstracao = 'False')"
    Else
        CFOPOutros = ""
    End If
    
    If chkVendas.Value = 1 And chkMaoObra.Value = 1 And Chk_demonstracao.Value = 1 And chkOutras.Value = 1 Then
        CFOP = "CFOP.id_CFOP IS NOT NULL"
    ElseIf chkVendas.Value = 0 And chkMaoObra.Value = 0 And Chk_demonstracao.Value = 0 And chkOutras.Value = 0 Then
            CFOP = "CFOP.Vendas = 'False' and CFOP.MaoObra = 'False' and CFOP.Demonstracao = 'False'"
        Else
            CFOP = CFOPVendas & CFOPMO & CFOPDEM & CFOPOutros
    End If
    
    If Chk_apenas_ST.Value = 1 Then
        If cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Emitente" Or cmbfiltrarpor = "UF" Or Left(cmbfiltrarpor, 11) = "Nota fiscal" Then
            STTexto = " and TN.dbl_Base_ICMS_Subst IS NOT NULL and TN.dbl_Base_ICMS_Subst <> 0"
        Else
            STTexto = " and NFPCI.Valor_BC_ST IS NOT NULL and NFPCI.Valor_BC_ST <> 0"
        End If
    Else
        STTexto = ""
    End If
    
    If Cmb_empresa <> "Todas" Then TextoFiltroEmpresa = "and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) Else TextoFiltroEmpresa = ""
    
    Select Case cmbfiltrarpor
        Case "Código interno": NomeCampo = "NFP.int_Cod_Produto"
        Case "Código de referência": NomeCampo = "NFP.N_referencia"
        Case "CFOP": NomeCampo = "CFOP.id_CFOP"
        Case "Classificação fiscal": NomeCampo = "CF.IDIntClasse"
        Case "Descrição": NomeCampo = "NFP.txt_Descricao"
        Case "Família": NomeCampo = "NFP.Familia"
        Case "Grupo": NomeCampo = "PF.Grupo"
        Case "Destinatário": NomeCampo = "NF.txt_Razao_Nome"
        Case "Emitente": NomeCampo = "NF.txt_Razao_Nome"
        Case "UF": NomeCampo = "NF.txt_UF"
        Case "Nota fiscal": NomeCampo = "NF.int_NotaFiscal"
        Case "Nota fiscal x Destinatário": NomeCampo = "NF.txt_Razao_Nome"
        Case "Nota fiscal x Emitente": NomeCampo = "NF.txt_Razao_Nome"
    End Select
    
    NomeCampo1 = ""
    OrdenarTexto = ""
    If cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Emitente" Or Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then
        NomeCampo1 = ", NF.Id_Int_Cliente, NF.txt_Municipio"
        OrdenarTexto = ", NF.txt_Municipio"
    End If
    
    INNERJOINTEXTO = "Select " & NomeCampo & " as NomeCampo1 " & NomeCampo1 & " from (((((tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_Detalhes_Nota NFP on NFP.ID_nota = NF.ID) LEFT JOIN tbl_NaturezaOperacao CFOP ON NFP.ID_cfop = CFOP.IDCountCfop) LEFT JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF) LEFT JOIN Projfamilia PF ON PF.Familia = NFP.Familia) LEFT JOIN tbl_Detalhes_Nota_CST_ICMS NFPCI ON NFPCI.ID_item = NFP.Int_codigo) LEFT JOIN tbl_Totais_Nota TN ON TN.ID_nota = NF.ID"
    
    If Sit_Nota = 0 Then
      TipoNotaTexto = " (NF.int_TipoNota = 1 or NF.int_TipoNota = 2)"
    Else
      TipoNotaTexto = " NF.int_TipoNota = " & Sit_Nota
    End If
    
    If cmbSerie.Text <> "" Then
     ' SerieNotaTexto = " (Serie = '1' or Serie = '2')"
    'Else
      SerieNotaTexto = " Serie = '" & cmbSerie.Text & "' And "
    End If
        

    Set TBCarteira = CreateObject("adodb.recordset")
    StrSql = INNERJOINTEXTO & " where " & SerieNotaTexto & NomeCampo & " IS NOT NULL and NF.Aplicacao = '" & Aplicacao_NF & "' and " & StatusTexto & " and " & TipoNotaTexto & " and NF.DtValidacao IS NOT NULL " & TextoFiltroEmpresa & " and " & CFOP & STTexto & " and " & NomeCampo & " IS NOT NULL Group by " & NomeCampo & NomeCampo1 & ", NF.Aplicacao, NF.int_status, NF.int_TipoNota, CFOP.id_CFOP order by " & NomeCampo & OrdenarTexto
    'Debug.print StrSql
    
    TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic
    
    If TBCarteira.EOF = False Then
        With cmbTexto
            .AddItem ""
            Do While TBCarteira.EOF = False
                If cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Emitente" Or Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then
                    CampoTexto = TBCarteira!NomeCampo1 & " (" & TBCarteira!txt_Municipio & ")"
                Else
                    CampoTexto = TBCarteira!NomeCampo1
                End If
                If CampoTexto <> "" And Desenho <> CampoTexto Then
                    .AddItem CampoTexto
                    If cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Emitente" Or Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then .ItemData(.NewIndex) = TBCarteira!Id_Int_Cliente
                End If
                Desenho = CampoTexto
                TBCarteira.MoveNext
            Loop
        End With
    End If
    TBCarteira.Close
End If
ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcAbrir()
On Error GoTo tratar_erro

Acao = "filtrar"
If chkVendas.Value = 0 And chkMaoObra.Value = 0 And Chk_demonstracao.Value = 0 And chkOutras.Value = 0 Then
    NomeCampo = "a operação"
    ProcVerificaAcao
    Exit Sub
End If
If Left(cmbfiltrarpor, 13) = "Nota fiscal x" And cmbTexto = "" Then
    NomeCampo = "o texto para pesquisa"
    ProcVerificaAcao
    cmbTexto.SetFocus
    Exit Sub
End If
With msk_fltFim
    If FunVerificaDataFinal(msk_fltInicio.Value, .Value) = False Then
        .Value = Date
        .SetFocus
        Exit Sub
    End If
End With
Valor1 = 0
Desenho = ""
Desenho1 = ""

Inicio = Time
ProcLimpaCamposTotais
ProcAbrirTabelas
If Txt_limite <> "" And Txt_limite <> "0" Then ProcVerificaLimiteRegistros
If Permitido = True Then ProcGravarTotalizacoes
'ordenar
If chkVlrTotal.Value = 1 Then Ordenar = "Numero15 desc, maquina" Else Ordenar = "maquina, Nota"
Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by " & Ordenar, Conexao, adOpenKeyset, adLockOptimistic
ProcCarregaLista

intervalo = Time
ElapsedTime (intervalo - Inicio)
Lbl_relatorio.Caption = "Registros encontrados: " & FunTamanhoTextoZeroEsq(Posicao, 4) & " - " & HoraTotal

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

If Opt_terceiros.Value = True Then
    Aplicacao_NF = "T"
    If Opt_dt_entrada.Value = True Then DataCampo = "NF.dt_Saida_Entrada" Else DataCampo = "NF.dt_DataEmissao"
Else
    Aplicacao_NF = "P"
    DataCampo = "NF.dt_DataEmissao"
End If

If optSaida.Value = True Then
Sit_Nota = 1
End If

If optEntrada.Value = True Then
Sit_Nota = 2
End If

If opttodas.Value = True Then
Sit_Nota = 0
End If

'Escolhe notas ativas e canceladas
If Chk_ativa.Value = 1 And chkCanceladas.Value = 1 Then
    StatusTexto = "NF.int_status IS NOT NULL"
ElseIf Chk_ativa.Value = 1 Then
        StatusTexto = "NF.int_status = 1" 'Somente ativas
    Else
        StatusTexto = "NF.int_status = 2" 'Somente canceladas
End If

If chkVendas.Value = 1 Then CFOPVendas = "CFOP.Vendas = 'True'" Else CFOPVendas = ""
If chkMaoObra.Value = 1 Then
    If CFOPVendas = "" Then CFOPMO = "CFOP.MaoObra = 'True'" Else CFOPMO = " or CFOP.MaoObra = 'True'"
Else
    CFOPMO = ""
End If
If Chk_demonstracao.Value = 1 Then
    If CFOPVendas = "" And CFOPMO = "" Then CFOPDEM = "CFOP.Demonstracao = 'True'" Else CFOPDEM = " or CFOP.Demonstracao = 'True'"
Else
    CFOPDEM = ""
End If
If chkOutras.Value = 1 Then
    If CFOPVendas = "" And CFOPMO = "" And CFOPDEM = "" Then CFOPOutros = "CFOP.Vendas = 'False' and CFOP.MaoObra = 'False' and CFOP.Demonstracao = 'False'" Else CFOPOutros = " or (CFOP.Vendas = 'False' and CFOP.MaoObra = 'False' and CFOP.Demonstracao = 'False')"
Else
    CFOPOutros = ""
End If

If chkVendas.Value = 1 And chkMaoObra.Value = 1 And Chk_demonstracao.Value = 1 And chkOutras.Value = 1 Then
    CFOP = "CFOP.id_CFOP IS NOT NULL"
Else
    CFOP = "(" & CFOPVendas & CFOPMO & CFOPDEM & CFOPOutros & ")"
End If

'If cmbSerie.Text <> "" Then
'STTexto = " AND Serie = '" & cmbSerie.Text & "'"
'Else
'    STTexto = ""
'End If

'    If cmbSerie.Text = "" Then
'      SerieNotaTexto = " (Serie = 1 or Serie = 2)"
'    Else
'      SerieNotaTexto = " Serie = " & cmbSerie.Text
'    End If

If Chk_apenas_ST.Value = 1 Then
    If Opt_comparativo.Value = True And (cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Emitente" Or cmbfiltrarpor = "UF" Or Left(cmbfiltrarpor, 11) = "Nota fiscal") Then
        STTexto = STTexto & " and TN.dbl_Base_ICMS_Subst IS NOT NULL and TN.dbl_Base_ICMS_Subst <> 0"
    Else
        STTexto = STTexto & " and NFPCI.Valor_BC_ST IS NOT NULL and NFPCI.Valor_BC_ST <> 0"
    End If
End If

If Cmb_empresa <> "Todas" Then TextoFiltroEmpresa = "and NF.ID_empresa = " & Cmb_empresa.ItemData(Cmb_empresa.ListIndex) Else TextoFiltroEmpresa = ""
'DataFiltro = "(" & DataCampo & ") Between '" & Format(msk_fltInicio.Value, "Short Date") & "' And '" & Format(msk_fltFim.Value, "Short Date") & "'"
DataFiltro = "(" & DataCampo & ") >= '" & Format(msk_fltInicio.Value, "Short Date") & " 00:00:00.000' And (" & DataCampo & ") <= '" & Format(msk_fltFim.Value, "Short Date") & " 23:59:59.000'"
CamposFiltro = "NF.*, NFP.Tipo, NFP.Int_codigo, NFP.int_Cod_Produto, NFP.N_referencia, NFP.txt_Descricao, NFP.Familia, PF.Grupo, CF.IDIntClasse, (CFOP.id_CFOP) as CFOP, CFOP.Txt_descricao as Descricao_CFOP, CFOP.IDCountCfop, CFOP.Retem, NFP.int_Qtd, NFP.dbl_ValorTotal, NFP.int_ICMS, NFP.ICMS_SN, NFP.txt_CST, NFP.dbl_ValorUnitario, NFP.dbl_valoripi, NFP.Total_PIS_prod, NFP.Total_Cofins_prod, NFP.Total_CSLL_prod, NFP.Total_IRPJ_prod, NFP.valor_retencao_PIS, NFP.valor_retencao_Cofins, NFP.Valor_frete, NFP.Valor_seguro, NFP.Valor_acessorias, NFP.Valor_desconto, NFP.Valor_desconto_SUFRAMA, NFP.Remessa, NFP.Retorno, NFP.Total_PIS_serv, NFP.Total_Cofins_serv, NFP.Total_CSLL_serv, NFP.VlrISS, NFP.Total_INSS_serv, NFP.Total_IRPJ_serv, NFP.Total_IRRF_serv, NFP.Total_INSS_serv, NFPCI.Valor_BC, NFPCI.Valor_ICMS, NFPCI.Valor_BC_ST, NFPCI.Valor_ICMS_ST, NFPCI.Valor_ICMS_SN, TN.Total_DAS"
If Opt_individual.Value = True Then
    TextoFiltro = ""
    Select Case cmbfiltrarpor
        Case "Código interno": If cmbTexto <> "" Then TextoFiltro = "NFP.int_Cod_Produto = '" & cmbTexto & "' and "
        Case "Código de referência":   If cmbTexto <> "" Then TextoFiltro = "NFP.N_Referencia = '" & cmbTexto & "' and "
        Case "CFOP": If cmbTexto <> "" Then TextoFiltro = "CFOP.id_CFOP = '" & cmbTexto & "' and "
        Case "Classificação fiscal":  If cmbTexto <> "" Then TextoFiltro = "CF.IDIntClasse = '" & cmbTexto & "' and NFP.Tipo = 'P' and "
        Case "Descrição": If cmbTexto <> "" Then TextoFiltro = "NFP.txt_Descricao = '" & cmbTexto & "' and "
        Case "Família": If cmbTexto <> "" Then TextoFiltro = "NFP.Familia = '" & cmbTexto & "' and "
        Case "Grupo":  If cmbTexto <> "" Then TextoFiltro = "PF.Grupo = '" & cmbTexto & "' and "
        Case "Destinatário":
            If cmbTexto <> "" Then
                DestFiltro = FunVerifDestinatario(cmbTexto)
                TextoFiltro = "NF.Id_Int_Cliente = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and LEFT(NF.txt_Razao_Nome, " & Contador2 & ") = '" & DestFiltro & "' and "
            End If
        Case "Emitente":
            If cmbTexto <> "" Then
                DestFiltro = FunVerifDestinatario(cmbTexto)
                TextoFiltro = "NF.Id_Int_Cliente = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and LEFT(NF.txt_Razao_Nome, " & Contador2 & ") = '" & DestFiltro & "' and "
            End If
        Case "UF": If cmbTexto <> "" Then TextoFiltro = "NF.txt_UF = '" & cmbTexto & "' and "
        Case "Nota fiscal": If cmbTexto <> "" Then TextoFiltro = "NF.int_NotaFiscal = '" & cmbTexto & "' and "
    End Select
    OrdenarFiltro = "NF.dt_dataemissao, NF.int_NotaFiscal"
Else
    Select Case cmbfiltrarpor
        Case "Código interno": OrdenarFiltro = "NFP.Int_NotaFiscal, NFP.int_Cod_Produto"
        Case "Código de referência": OrdenarFiltro = "NFP.Int_NotaFiscal, NFP.N_Referencia"
        Case "CFOP": OrdenarFiltro = "NF.dt_dataemissao, CFOP.id_CFOP"
        Case "Classificação fiscal":
            TextoFiltro = "NFP.Tipo = 'P' and "
            OrdenarFiltro = "NFP.Int_NotaFiscal, CF.IDIntClasse"
        Case "Descrição": OrdenarFiltro = "NFP.Int_NotaFiscal, NFP.txt_Descricao"
        Case "Família": OrdenarFiltro = "NFP.Int_NotaFiscal, NFP.Familia"
        Case "Grupo": OrdenarFiltro = "NFP.Int_NotaFiscal, PF.Grupo"
        Case "Destinatário": OrdenarFiltro = "NF.dt_dataemissao, NF.txt_Razao_Nome"
        Case "Emitente": OrdenarFiltro = "NF.dt_dataemissao, NF.txt_Razao_Nome"
        Case "UF": OrdenarFiltro = "NF.dt_dataemissao, NF.txt_UF"
        Case "Nota fiscal": OrdenarFiltro = "NF.dt_dataemissao, NF.int_NotaFiscal"
        Case "Nota fiscal x Destinatário":
            DestFiltro = FunVerifDestinatario(cmbTexto)
            TextoFiltro = "NF.Id_Int_Cliente = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and LEFT(NF.txt_Razao_Nome, " & Contador2 & ") = '" & DestFiltro & "' and "
            OrdenarFiltro = "NF.dt_dataemissao, NF.int_NotaFiscal"
        Case "Nota fiscal x Emitente":
            DestFiltro = FunVerifDestinatario(cmbTexto)
            TextoFiltro = "NF.Id_Int_Cliente = " & cmbTexto.ItemData(cmbTexto.ListIndex) & " and LEFT(NF.txt_Razao_Nome, " & Contador2 & ") = '" & DestFiltro & "' and "
            OrdenarFiltro = "NF.dt_dataemissao, NF.int_NotaFiscal"
    End Select
End If

If cmbSerie.Text <> "" Then
  SerieNotaTexto = " Serie = '" & cmbSerie.Text & "' And "
End If


INNERJOINTEXTO = "Select " & CamposFiltro & " from (((((tbl_Dados_Nota_Fiscal NF LEFT JOIN tbl_Detalhes_Nota NFP on NFP.ID_nota = NF.ID) LEFT JOIN tbl_NaturezaOperacao CFOP ON NFP.ID_cfop = CFOP.IDCountCfop) LEFT JOIN tbl_ClassificacaoFiscal CF ON CF.Idclass = NFP.ID_CF) LEFT JOIN Projfamilia PF ON PF.Familia = NFP.Familia) LEFT JOIN tbl_Detalhes_Nota_CST_ICMS NFPCI ON NFPCI.ID_item = NFP.Int_codigo) LEFT JOIN tbl_Totais_Nota TN ON TN.ID_nota = NF.ID where " & SerieNotaTexto & TextoFiltro
Set TBCarteira = CreateObject("adodb.recordset")

If Sit_Nota = 0 Then
TipoNotaTexto = " (NF.int_TipoNota = 1 or NF.int_TipoNota = 2)"
Else
TipoNotaTexto = " NF.int_TipoNota = " & Sit_Nota
End If



StrSql = INNERJOINTEXTO & "NF.DtValidacao IS NOT NULL and NF.Aplicacao = '" & Aplicacao_NF & "' and " & StatusTexto & " and " & TipoNotaTexto & " " & TextoFiltroEmpresa & " and " & DataFiltro & " and " & CFOP & STTexto & " order by " & OrdenarFiltro
'Debug.print StrSql

TBCarteira.Open StrSql, Conexao, adOpenKeyset, adLockOptimistic

ProcFiltrar

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Function FunVerifDestinatario(Destinatario As String) As String
On Error GoTo tratar_erro

Texto = ""
Numero = 0
Numero1 = Len(Destinatario)
Do While Numero1 <> 0
    If Texto = "(" Then GoTo Pula
    Texto = Left(Destinatario, (Numero + 1))
    Texto = Right(Texto, Len(Texto) - Numero)
    Numero = Numero + 1
    Numero1 = Numero1 - 1
Loop
Pula:
Numero = Numero - 2
FunVerifDestinatario = Left(Destinatario, Numero)
Contador2 = Numero

Exit Function
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Function
End Function

Private Sub ProcFiltrar()
On Error GoTo tratar_erro

If TBCarteira.EOF = False Then
    Permitido = True
    PBLista.Min = 0
    PBLista.Max = TBCarteira.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBCarteira.EOF = False
        Set TBProdutividade = CreateObject("adodb.recordset")
        If optDetalhado.Value = True Then
            TBProdutividade.Open "Select * from Producao_Relatorios", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosDetalhado
        Else
            TextoFiltro2 = ""
            If cmbfiltrarpor = "Nota fiscal" Or Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then
                TextoFiltro = TBCarteira!int_NotaFiscal
                TextoFiltro1 = TBCarteira!txt_Razao_Nome & " (" & TBCarteira!txt_Municipio & ")"
                TextoFiltro2 = "and Descevento = '" & TextoFiltro1 & "'"
            Else
                Select Case cmbfiltrarpor
                    Case "Código interno": TextoFiltro = TBCarteira!int_Cod_Produto
                    Case "Código de referência": TextoFiltro = TBCarteira!N_referencia
                    Case "CFOP": TextoFiltro = TBCarteira!CFOP
                    Case "Classificação fiscal": TextoFiltro = TBCarteira!IDIntClasse
                    Case "Descrição": TextoFiltro = TBCarteira!Txt_descricao
                    Case "Família": TextoFiltro = TBCarteira!Familia
                    Case "Grupo": TextoFiltro = TBCarteira!Grupo
                    Case "Destinatário": TextoFiltro = TBCarteira!txt_Razao_Nome & " (" & TBCarteira!txt_Municipio & ")"
                    Case "Emitente": TextoFiltro = TBCarteira!txt_Razao_Nome & " (" & TBCarteira!txt_Municipio & ")"
                    Case "UF": TextoFiltro = TBCarteira!txt_UF
                End Select
            End If
            TBProdutividade.Open "Select * from Producao_Relatorios where maquina = '" & TextoFiltro & "' " & TextoFiltro2 & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
            ProcEnviaDadosResumido
        End If
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

Private Sub ProcEnviaDadosDetalhado()
On Error GoTo tratar_erro

Qtde = 0
Qtd = 0
TBProdutividade.AddNew
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
TBProdutividade!maquina = cmbTexto

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Empresa FROM Empresa WHERE Codigo = " & TBCarteira!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBProdutividade!Data3 = TBAbrir!Empresa 'Empresa
End If
TBAbrir.Close
TBProdutividade!Ordem = TBCarteira!ID 'ID da nota fiscal
TBProdutividade!Nota = TBCarteira!int_NotaFiscal 'Numero da nota fiscal

'Data de emissão
If Opt_terceiros.Value = True Then
    If Opt_dt_entrada.Value = True Then Data = TBCarteira!dt_Saida_Entrada Else Data = TBCarteira!dt_DataEmissao
Else
    Data = TBCarteira!dt_DataEmissao
End If
TBProdutividade!Data = Data

TBProdutividade!Execucaoprev = IIf(IsNull(TBCarteira!TipoNF), "", TBCarteira!TipoNF) 'Tipo da nota
TBProdutividade!DescEvento = IIf(IsNull(TBCarteira!txt_Razao_Nome), "", TBCarteira!txt_Razao_Nome) & " (" & IIf(IsNull(TBCarteira!txt_Municipio), "", TBCarteira!txt_Municipio) & ")" 'Cliente
TBProdutividade!Totalhsprev = IIf(IsNull(TBCarteira!CFOP), "", TBCarteira!CFOP) 'CFOP
TBProdutividade!Totalhsutil = IIf(IsNull(TBCarteira!Descricao_CFOP), "", TBCarteira!Descricao_CFOP) 'Natureza de operação

'Verifica se a NF é complementar e busca os dados dos totais da NF
'Set TBFIltro = CreateObject("adodb.recordset")
'TBFIltro.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & TBCarteira!ID & " and Finalidade_emissao <> 1", Conexao, adOpenKeyset, adLockOptimistic

If cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Emitente" Or cmbfiltrarpor = "UF" Or cmbfiltrarpor = "Nota fiscal" Or Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then
    
    TBProdutividade!ID_prod = TBCarteira!Int_codigo
    TBProdutividade!ID_serv = TBCarteira!Int_codigo
    TBProdutividade!Data1 = TBCarteira!int_Cod_Produto 'Código interno
    TBProdutividade!Data2 = TBCarteira!Txt_descricao 'Descrição
    
    If TBCarteira!Tipo = "P" Then
        TBProdutividade!Data4 = TBCarteira!IDIntClasse 'CF
        
        'If Opt_propria.Value = True And (TBFIltro.EOF = False Or TBCarteira!Alterar = True) Then
        'If Opt_propria.Value = True And TBFIltro.EOF = False Then
            'ProcEnviaDadosTotaisNF
        'Else
            'Quantidade de produtos
'            Permitido1 = False
'            If TBCarteira!Remessa = False And TBCarteira!Retorno = False Then Permitido1 = True
'            If Chk_remessa.Value = 1 And TBCarteira!Remessa = True Then Permitido1 = True
'            If Chk_retorno.Value = 1 And TBCarteira!Retorno = True Then Permitido1 = True
'            If Permitido1 = True Then
            TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira!int_Qtd), 0, TBCarteira!int_Qtd)
            TBProdutividade!qtdeNC = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos
            TBProdutividade!Numero15 = TBProdutividade!Numero15 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal)
            
            If IsNull(TBCarteira!txt_CST) = False And TBCarteira!txt_CST <> "" Then
                If Len(TBCarteira!txt_CST) = 4 Then FimCST = Right(TBCarteira!txt_CST, 3) Else FimCST = Right(TBCarteira!txt_CST, 2)
                
                If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                    If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Then
                        TBProdutividade!Qtdetotalprod = TBCarteira!Valor_BC 'Base de cálculo ICMS
                        TBProdutividade!Eficiencia = TBCarteira!Valor_ICMS 'Valor do ICMS
                    End If
                    
                    If FimCST = "10" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                        TBProdutividade!Terceiros = TBCarteira!Valor_BC_ST 'Base de calculo ICMS subst.
                        TBProdutividade!impostos = TBCarteira!Valor_ICMS_ST 'Valor do ICMS subst.
                        TBProdutividade!Numero15 = TBProdutividade!Numero15 + TBCarteira!Valor_ICMS_ST
                    End If
                    
                    If FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                        TBProdutividade!Qtdetotalprod = TBCarteira!Valor_BC 'Base de cálculo ICMS SN
                        TBProdutividade!Valor5 = TBCarteira!Valor_ICMS_SN 'Valor do ICMS SN
                    End If
                End If
            End If
            
            TBProdutividade!Lucro = IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) 'Valor total de ipi
            TBProdutividade!material = IIf(IsNull(TBCarteira!Total_PIS_prod), 0, TBCarteira!Total_PIS_prod) 'Valor total pis
            TBProdutividade!Servicos = IIf(IsNull(TBCarteira!Total_Cofins_prod), 0, TBCarteira!Total_Cofins_prod) 'Valor total cofins
            TBProdutividade!Total = IIf(IsNull(TBCarteira!Total_CSLL_prod), 0, TBCarteira!Total_CSLL_prod) 'Total CSLL
            TBProdutividade!Total_peca = IIf(IsNull(TBCarteira!Total_IRPJ_prod), 0, TBCarteira!Total_IRPJ_prod) 'Total IRPJ
            TBProdutividade!Refugo = IIf(IsNull(TBCarteira!Valor_Retencao_PIS), 0, TBCarteira!Valor_Retencao_PIS) 'Total retenção Pis
            TBProdutividade!Numero1 = IIf(IsNull(TBCarteira!Valor_Retencao_Cofins), 0, TBCarteira!Valor_Retencao_Cofins) 'Total retenção confins
            TBProdutividade!Numero11 = IIf(IsNull(TBCarteira!Valor_frete), 0, TBCarteira!Valor_frete) 'Total do frete
            TBProdutividade!Numero12 = IIf(IsNull(TBCarteira!Valor_seguro), 0, TBCarteira!Valor_seguro) 'Total do seguro
            TBProdutividade!Numero13 = IIf(IsNull(TBCarteira!Valor_acessorias), 0, TBCarteira!Valor_acessorias) 'Total outras despesas
            TBProdutividade!Valor6 = IIf(IsNull(TBCarteira!Valor_desconto), 0, TBCarteira!Valor_desconto) + IIf(IsNull(TBCarteira!Valor_desconto_SUFRAMA), 0, TBCarteira!Valor_desconto_SUFRAMA) 'Total do desconto
            If TBCarteira!retorno = True Then TBProdutividade!Valor4 = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos (retorno)
            If TBCarteira!Remessa = True Then TBProdutividade!Valor7 = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos (remessa)
            
            'Valor total faturado
            TBProdutividade!Numero15 = TBProdutividade!Numero15 + (IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) + IIf(IsNull(TBCarteira!Valor_frete), 0, TBCarteira!Valor_frete) + IIf(IsNull(TBCarteira!Valor_seguro), 0, TBCarteira!Valor_seguro) + IIf(IsNull(TBCarteira!Valor_acessorias), 0, TBCarteira!Valor_acessorias)) - (IIf(IsNull(TBCarteira!Valor_Retencao_PIS), 0, TBCarteira!Valor_Retencao_PIS) + IIf(IsNull(TBCarteira!Valor_Retencao_Cofins), 0, TBCarteira!Valor_Retencao_Cofins) + IIf(IsNull(TBCarteira!Valor_desconto), 0, TBCarteira!Valor_desconto) + IIf(IsNull(TBCarteira!Valor_desconto_SUFRAMA), 0, TBCarteira!Valor_desconto_SUFRAMA))
            TBProdutividade!Numero14 = TBProdutividade!Numero14 + IIf(IsNull(TBCarteira!Total_DAS), 0, TBCarteira!Total_DAS)
        'End If
    Else
        TBProdutividade!Numero2 = IIf(IsNull(TBCarteira!int_Qtd), 0, TBCarteira!int_Qtd) 'Quantidade de serviços
        TBProdutividade!Numero3 = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total serviço
        If TBCarteira!Retem = True Then
            TBProdutividade!Numero4 = IIf(IsNull(TBCarteira!Total_PIS_serv), 0, TBCarteira!Total_PIS_serv) 'Total pis serviços
            TBProdutividade!Numero5 = IIf(IsNull(TBCarteira!Total_Cofins_serv), 0, TBCarteira!Total_Cofins_serv) 'Total cofins serviços
            TBProdutividade!Numero6 = IIf(IsNull(TBCarteira!Total_CSLL_serv), 0, TBCarteira!Total_CSLL_serv) 'Total CSLL serviços
            TBProdutividade!Numero7 = IIf(IsNull(TBCarteira!VlrISS), 0, TBCarteira!VlrISS) 'Valor total de ISS
            TBProdutividade!Numero8 = IIf(IsNull(TBCarteira!Total_INSS_serv), 0, TBCarteira!Total_INSS_serv) 'Total de INSS serviços
            TBProdutividade!Numero9 = IIf(IsNull(TBCarteira!Total_IRPJ_serv), 0, TBCarteira!Total_IRPJ_serv) 'Total de IRPJ serviços
            TBProdutividade!Numero10 = IIf(IsNull(TBCarteira!Total_IRRF_serv), 0, TBCarteira!Total_IRRF_serv) 'Total de IRRF serviços
        End If
        TBProdutividade!Numero15 = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total faturado
        TBProdutividade!Numero14 = TBProdutividade!Numero14 + IIf(IsNull(TBCarteira!Total_DAS), 0, TBCarteira!Total_DAS)
    End If
Else
    TBProdutividade!ID_prod = TBCarteira!Int_codigo
    TBProdutividade!ID_serv = TBCarteira!Int_codigo
    TBProdutividade!Data1 = TBCarteira!int_Cod_Produto 'Código interno
    TBProdutividade!Data2 = TBCarteira!Txt_descricao 'Descrição
    If TBCarteira!Tipo = "P" Then
        TBProdutividade!Data4 = TBCarteira!IDIntClasse
        
        'If TBFIltro.EOF = False Then
            'ProcEnviaDadosTotaisNF
        'Else
            'Quantidade de produtos
'            Permitido1 = False
'            If TBCarteira!Remessa = False And TBCarteira!Retorno = False Then Permitido1 = True
'            If Chk_remessa.Value = 1 And TBCarteira!Remessa = True Then Permitido1 = True
'            If Chk_retorno.Value = 1 And TBCarteira!Retorno = True Then Permitido1 = True
'            If Permitido1 = True Then
            TBProdutividade!qtdeOK = IIf(IsNull(TBCarteira!int_Qtd), 0, TBCarteira!int_Qtd)
            TBProdutividade!qtdeNC = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos
            TBProdutividade!Numero15 = TBProdutividade!Numero15 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal)
            
            If IsNull(TBCarteira!txt_CST) = False And TBCarteira!txt_CST <> "" Then
                If Len(TBCarteira!txt_CST) = 4 Then FimCST = Right(TBCarteira!txt_CST, 3) Else FimCST = Right(TBCarteira!txt_CST, 2)
            Else
                FimCST = ""
            End If
            
            If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                Set TBCST = CreateObject("adodb.recordset")
                TBCST.Open "select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & TBCarteira!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                If TBCST.EOF = False Then
                    If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                        TBProdutividade!Qtdetotalprod = TBCST!Valor_BC 'Base de cálculo ICMS
                        TBProdutividade!Eficiencia = TBCST!Valor_ICMS 'Valor do ICMS
                    End If
                    
                    If FimCST = "10" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                        TBProdutividade!Terceiros = TBCST!Valor_BC_ST 'Base de calculo ICMS subst.
                        TBProdutividade!impostos = TBCST!Valor_ICMS_ST 'Valor do ICMS subst.
                        TBProdutividade!Numero15 = TBProdutividade!Numero15 + TBCST!Valor_ICMS_ST
                    End If
                    
                    If FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                        TBProdutividade!Qtdetotalprod = TBCST!Valor_BC 'Base de cálculo ICMS SN
                        TBProdutividade!Valor5 = TBCST!Valor_ICMS_SN 'Valor do ICMS SN
                    End If
                End If
                TBCST.Close
            End If
            
            TBProdutividade!Lucro = IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) 'Valor total de ipi
            TBProdutividade!material = IIf(IsNull(TBCarteira!Total_PIS_prod), 0, TBCarteira!Total_PIS_prod) 'Valor total pis
            TBProdutividade!Servicos = IIf(IsNull(TBCarteira!Total_Cofins_prod), 0, TBCarteira!Total_Cofins_prod) 'Valor total cofins
            TBProdutividade!Total = IIf(IsNull(TBCarteira!Total_CSLL_prod), 0, TBCarteira!Total_CSLL_prod) 'Total CSLL
            TBProdutividade!Total_peca = IIf(IsNull(TBCarteira!Total_IRPJ_prod), 0, TBCarteira!Total_IRPJ_prod) 'Total IRPJ
            TBProdutividade!Refugo = IIf(IsNull(TBCarteira!Valor_Retencao_PIS), 0, TBCarteira!Valor_Retencao_PIS) 'Total retenção Pis
            TBProdutividade!Numero1 = IIf(IsNull(TBCarteira!Valor_Retencao_Cofins), 0, TBCarteira!Valor_Retencao_Cofins) 'Total retenção confins
            TBProdutividade!Numero11 = IIf(IsNull(TBCarteira!Valor_frete), 0, TBCarteira!Valor_frete) 'Total do frete
            TBProdutividade!Numero12 = IIf(IsNull(TBCarteira!Valor_seguro), 0, TBCarteira!Valor_seguro) 'Total do seguro
            TBProdutividade!Numero13 = IIf(IsNull(TBCarteira!Valor_acessorias), 0, TBCarteira!Valor_acessorias) 'Total outras despesas
            TBProdutividade!Valor6 = IIf(IsNull(TBCarteira!Valor_desconto), 0, TBCarteira!Valor_desconto) + IIf(IsNull(TBCarteira!Valor_desconto_SUFRAMA), 0, TBCarteira!Valor_desconto_SUFRAMA) 'Total do desconto
            If TBCarteira!retorno = True Then TBProdutividade!Valor4 = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos (retorno)
            If TBCarteira!Remessa = True Then TBProdutividade!Valor7 = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos (remessa)
            
            'Valor total faturado
            TBProdutividade!Numero15 = TBProdutividade!Numero15 + (IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) + IIf(IsNull(TBCarteira!Valor_frete), 0, TBCarteira!Valor_frete) + IIf(IsNull(TBCarteira!Valor_seguro), 0, TBCarteira!Valor_seguro) + IIf(IsNull(TBCarteira!Valor_acessorias), 0, TBCarteira!Valor_acessorias)) - (IIf(IsNull(TBCarteira!Valor_Retencao_PIS), 0, TBCarteira!Valor_Retencao_PIS) + IIf(IsNull(TBCarteira!Valor_Retencao_Cofins), 0, TBCarteira!Valor_Retencao_Cofins) + IIf(IsNull(TBCarteira!Valor_desconto), 0, TBCarteira!Valor_desconto) + IIf(IsNull(TBCarteira!Valor_desconto_SUFRAMA), 0, TBCarteira!Valor_desconto_SUFRAMA))
            TBProdutividade!Numero14 = TBProdutividade!Numero14 + IIf(IsNull(TBCarteira!Total_DAS), 0, TBCarteira!Total_DAS)
        'End If
    Else
        TBProdutividade!Totalhsprev = IIf(IsNull(TBCarteira!CFOP), "", TBCarteira!CFOP) 'CFOP
        TBProdutividade!Totalhsutil = IIf(IsNull(TBCarteira!Descricao_CFOP), "", TBCarteira!Descricao_CFOP) 'Natureza de operação
        
        TBProdutividade!Numero2 = IIf(IsNull(TBCarteira!int_Qtd), 0, TBCarteira!int_Qtd) 'Quantidade de serviços
        TBProdutividade!Numero3 = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total serviço
        If TBCarteira!Retem = True Then
            TBProdutividade!Numero4 = IIf(IsNull(TBCarteira!Total_PIS_serv), 0, TBCarteira!Total_PIS_serv) 'Total pis serviços
            TBProdutividade!Numero5 = IIf(IsNull(TBCarteira!Total_Cofins_serv), 0, TBCarteira!Total_Cofins_serv) 'Total cofins serviços
            TBProdutividade!Numero6 = IIf(IsNull(TBCarteira!Total_CSLL_serv), 0, TBCarteira!Total_CSLL_serv) 'Total CSLL serviços
            TBProdutividade!Numero7 = IIf(IsNull(TBCarteira!VlrISS), 0, TBCarteira!VlrISS) 'Valor total de ISS
            TBProdutividade!Numero8 = IIf(IsNull(TBCarteira!Total_INSS_serv), 0, TBCarteira!Total_INSS_serv) 'Total de INSS serviços
            TBProdutividade!Numero9 = IIf(IsNull(TBCarteira!Total_IRPJ_serv), 0, TBCarteira!Total_IRPJ_serv) 'Total de IRPJ serviços
            TBProdutividade!Numero10 = IIf(IsNull(TBCarteira!Total_IRRF_serv), 0, TBCarteira!Total_IRRF_serv) 'Total de IRRF serviços
        End If
        TBProdutividade!Numero15 = IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total faturado
        TBProdutividade!Numero14 = TBProdutividade!Numero14 + IIf(IsNull(TBCarteira!Total_DAS), 0, TBCarteira!Total_DAS)
    End If
End If
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosResumido()
On Error GoTo tratar_erro

Qtd = 0
Qtde = 0
If TBProdutividade.EOF = True Then TBProdutividade.AddNew
If optDetalhado.Value = True Then
    Texto = cmbTexto
Else
    Select Case cmbfiltrarpor
        Case "Código interno": Texto = TBCarteira!int_Cod_Produto
        Case "Código de referência": Texto = TBCarteira!N_referencia
        Case "CFOP": Texto = TBCarteira!CFOP
        Case "Classificação fiscal": Texto = TBCarteira!IDIntClasse
        Case "Descrição": Texto = TBCarteira!Txt_descricao
        Case "Família": Texto = TBCarteira!Familia
        Case "Grupo": Texto = TBCarteira!Grupo
        Case "Destinatário": Texto = TBCarteira!txt_Razao_Nome & " (" & TBCarteira!txt_Municipio & ")"
        Case "Emitente": Texto = TBCarteira!txt_Razao_Nome & " (" & TBCarteira!txt_Municipio & ")"
        Case "UF": Texto = TBCarteira!txt_UF
        Case "Nota fiscal": Texto = TBCarteira!int_NotaFiscal
        Case "Nota fiscal x Destinatário": Texto = TBCarteira!int_NotaFiscal
        Case "Nota fiscal x Emitente": Texto = TBCarteira!int_NotaFiscal
    End Select
End If
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Empresa FROM Empresa WHERE Codigo = " & TBCarteira!ID_empresa, Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    TBProdutividade!Data3 = TBAbrir!Empresa 'Empresa
End If
TBAbrir.Close

If IsNumeric(TBCarteira!Serie) Then
TBProdutividade!OS = TBCarteira!Serie
End If

TBProdutividade!maquina = Texto
TBProdutividade!Responsavel = pubUsuario
TBProdutividade!Modulo = Formulario
If cmbfiltrarpor = "Nota fiscal" Or Left(cmbfiltrarpor, 13) = "Nota fiscal x" Then
    If Opt_terceiros.Value = True Then
        If Opt_dt_entrada.Value = True Then Data = TBCarteira!dt_Saida_Entrada Else Data = TBCarteira!dt_DataEmissao
    Else
        Data = TBCarteira!dt_DataEmissao
    End If
    TBProdutividade!Data = Data
    TBProdutividade!Ordem = TBCarteira!ID 'ID da nota fiscal
End If
TBProdutividade!Nota = TBCarteira!int_NotaFiscal 'Numero da nota fiscal
TBProdutividade!DescEvento = IIf(IsNull(TBCarteira!txt_Razao_Nome), "", TBCarteira!txt_Razao_Nome) & " (" & IIf(IsNull(TBCarteira!txt_Municipio), "", TBCarteira!txt_Municipio) & ")"  'Cliente
TBProdutividade!Totalhsprev = IIf(IsNull(TBCarteira!CFOP), "", TBCarteira!CFOP) 'CFOP
TBProdutividade!Totalhsutil = IIf(IsNull(TBCarteira!Descricao_CFOP), "", TBCarteira!Descricao_CFOP) 'Natureza de operação

'Verifica se a NF é complementar e busca os dados dos totais da NF
'Set TBFIltro = CreateObject("adodb.recordset")
'TBFIltro.Open "Select * FROM tbl_Dados_Nota_Fiscal_NFe where ID_nota = " & TBCarteira!ID & " and Finalidade_emissao <> 1", Conexao, adOpenKeyset, adLockOptimistic

If cmbfiltrarpor = "Destinatário" Or cmbfiltrarpor = "Emitente" Or cmbfiltrarpor = "UF" Or Left(cmbfiltrarpor, 11) = "Nota fiscal" Then
    If TBCarteira!Tipo = "P" Then
        TBProdutividade!Data4 = TBCarteira!IDIntClasse 'CF
        
        'If Opt_propria.Value = True And (TBFIltro.EOF = False Or TBCarteira!Alterar = True) Then
        'If Opt_propria.Value = True And TBFIltro.EOF = False Then
            'ProcEnviaDadosTotaisNF
        'Else
            'Quantidade de produtos
'            Permitido1 = False
'            If TBCarteira!Remessa = False And TBCarteira!Retorno = False Then Permitido1 = True
'            If Chk_remessa.Value = 1 And TBCarteira!Remessa = True Then Permitido1 = True
'            If Chk_retorno.Value = 1 And TBCarteira!Retorno = True Then Permitido1 = True
'            If Permitido1 = True Then
            TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + IIf(IsNull(TBCarteira!int_Qtd), 0, TBCarteira!int_Qtd)
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos
            TBProdutividade!Numero15 = TBProdutividade!Numero15 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal)
            
            If IsNull(TBCarteira!txt_CST) = False And TBCarteira!txt_CST <> "" Then
                If Len(TBCarteira!txt_CST) = 4 Then FimCST = Right(TBCarteira!txt_CST, 3) Else FimCST = Right(TBCarteira!txt_CST, 2)
                
                If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                    If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Then
                        TBProdutividade!Qtdetotalprod = TBProdutividade!Qtdetotalprod + TBCarteira!Valor_BC 'Base de cálculo ICMS
                        TBProdutividade!Eficiencia = TBProdutividade!Eficiencia + TBCarteira!Valor_ICMS 'Valor do ICMS
                    End If
                    
                    If FimCST = "10" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                        TBProdutividade!Terceiros = TBProdutividade!Terceiros + TBCarteira!Valor_BC_ST 'Base de calculo ICMS subst.
                        TBProdutividade!impostos = TBProdutividade!impostos + TBCarteira!Valor_ICMS_ST 'Valor do ICMS subst.
                        TBProdutividade!Numero15 = TBProdutividade!Numero15 + TBCarteira!Valor_ICMS_ST
                    End If
                    
                    If FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                        TBProdutividade!Qtdetotalprod = TBProdutividade!Qtdetotalprod + TBCarteira!Valor_BC 'Base de cálculo ICMS SN
                        TBProdutividade!Valor5 = TBProdutividade!Valor5 + TBCarteira!Valor_ICMS_SN 'Valor do ICMS SN
                    End If
                End If
            End If
            
            TBProdutividade!Lucro = TBProdutividade!Lucro + IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) 'Valor total de ipi
            TBProdutividade!material = TBProdutividade!material + IIf(IsNull(TBCarteira!Total_PIS_prod), 0, TBCarteira!Total_PIS_prod) 'Valor total pis
            TBProdutividade!Servicos = TBProdutividade!Servicos + IIf(IsNull(TBCarteira!Total_Cofins_prod), 0, TBCarteira!Total_Cofins_prod) 'Valor total cofins
            TBProdutividade!Total = TBProdutividade!Total + IIf(IsNull(TBCarteira!Total_CSLL_prod), 0, TBCarteira!Total_CSLL_prod) 'Total CSLL
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + IIf(IsNull(TBCarteira!Total_IRPJ_prod), 0, TBCarteira!Total_IRPJ_prod) 'Total IRPJ
            TBProdutividade!Refugo = TBProdutividade!Refugo + IIf(IsNull(TBCarteira!Valor_Retencao_PIS), 0, TBCarteira!Valor_Retencao_PIS) 'Total retenção Pis
            TBProdutividade!Numero1 = TBProdutividade!Numero1 + IIf(IsNull(TBCarteira!Valor_Retencao_Cofins), 0, TBCarteira!Valor_Retencao_Cofins) 'Total retenção confins
            TBProdutividade!Numero11 = TBProdutividade!Numero11 + IIf(IsNull(TBCarteira!Valor_frete), 0, TBCarteira!Valor_frete) 'Total do frete
            TBProdutividade!Numero12 = TBProdutividade!Numero12 + IIf(IsNull(TBCarteira!Valor_seguro), 0, TBCarteira!Valor_seguro) 'Total do seguro
            TBProdutividade!Numero13 = TBProdutividade!Numero13 + IIf(IsNull(TBCarteira!Valor_acessorias), 0, TBCarteira!Valor_acessorias) 'Total outras despesas
            TBProdutividade!Valor6 = TBProdutividade!Valor6 + IIf(IsNull(TBCarteira!Valor_desconto), 0, TBCarteira!Valor_desconto) + IIf(IsNull(TBCarteira!Valor_desconto_SUFRAMA), 0, TBCarteira!Valor_desconto_SUFRAMA) 'Total do desconto
            If TBCarteira!retorno = True Then TBProdutividade!Valor4 = TBProdutividade!Valor4 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos (retorno)
            If TBCarteira!Remessa = True Then TBProdutividade!Valor7 = TBProdutividade!Valor7 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos (remessa)
            
            'Valor total faturado
            TBProdutividade!Numero15 = TBProdutividade!Numero15 + (IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) + IIf(IsNull(TBCarteira!Valor_frete), 0, TBCarteira!Valor_frete) + IIf(IsNull(TBCarteira!Valor_seguro), 0, TBCarteira!Valor_seguro) + IIf(IsNull(TBCarteira!Valor_acessorias), 0, TBCarteira!Valor_acessorias)) - (IIf(IsNull(TBCarteira!Valor_Retencao_PIS), 0, TBCarteira!Valor_Retencao_PIS) + IIf(IsNull(TBCarteira!Valor_Retencao_Cofins), 0, TBCarteira!Valor_Retencao_Cofins) + IIf(IsNull(TBCarteira!Valor_desconto), 0, TBCarteira!Valor_desconto) + IIf(IsNull(TBCarteira!Valor_desconto_SUFRAMA), 0, TBCarteira!Valor_desconto_SUFRAMA))
        'End If
    Else
        TBProdutividade!Numero2 = TBProdutividade!Numero2 + IIf(IsNull(TBCarteira!int_Qtd), 0, TBCarteira!int_Qtd) 'Quantidade de serviços
        TBProdutividade!Numero3 = TBProdutividade!Numero3 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total serviço
        If TBCarteira!Retem = True Then
            TBProdutividade!Numero4 = TBProdutividade!Numero4 + IIf(IsNull(TBCarteira!Total_PIS_serv), 0, TBCarteira!Total_PIS_serv) 'Total pis serviços
            TBProdutividade!Numero5 = TBProdutividade!Numero5 + IIf(IsNull(TBCarteira!Total_Cofins_serv), 0, TBCarteira!Total_Cofins_serv) 'Total cofins serviços
            TBProdutividade!Numero6 = TBProdutividade!Numero6 + IIf(IsNull(TBCarteira!Total_CSLL_serv), 0, TBCarteira!Total_CSLL_serv) 'Total CSLL serviços
            TBProdutividade!Numero7 = TBProdutividade!Numero7 + IIf(IsNull(TBCarteira!VlrISS), 0, TBCarteira!VlrISS) 'Valor total de ISS
            TBProdutividade!Numero8 = TBProdutividade!Numero8 + IIf(IsNull(TBCarteira!Total_INSS_serv), 0, TBCarteira!Total_INSS_serv) 'Total de INSS serviços
            TBProdutividade!Numero9 = TBProdutividade!Numero9 + IIf(IsNull(TBCarteira!Total_IRPJ_serv), 0, TBCarteira!Total_IRPJ_serv) 'Total de IRPJ serviços
            TBProdutividade!Numero10 = TBProdutividade!Numero10 + IIf(IsNull(TBCarteira!Total_IRRF_serv), 0, TBCarteira!Total_IRRF_serv) 'Total de IRRF serviços
        End If
        TBProdutividade!Numero15 = TBProdutividade!Numero15 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total faturado
        TBProdutividade!Numero14 = TBProdutividade!Numero14 + IIf(IsNull(TBCarteira!Total_DAS), 0, TBCarteira!Total_DAS) 'Valor DAS
    End If
Else
    If TBCarteira!Tipo = "P" Then
        TBProdutividade!Data4 = TBCarteira!IDIntClasse 'CF
        
        'If TBFIltro.EOF = False Then
            'ProcEnviaDadosTotaisNF
        'Else
            'Quantidade de produtos
'            Permitido1 = False
'            If TBCarteira!Remessa = False And TBCarteira!Retorno = False Then Permitido1 = True
'            If Chk_remessa.Value = 1 And TBCarteira!Remessa = True Then Permitido1 = True
'            If Chk_retorno.Value = 1 And TBCarteira!Retorno = True Then Permitido1 = True
'            If Permitido1 = True Then
            TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + IIf(IsNull(TBCarteira!int_Qtd), 0, TBCarteira!int_Qtd)
            TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos
            TBProdutividade!Numero15 = TBProdutividade!Numero15 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal)
            
            If IsNull(TBCarteira!txt_CST) = False And TBCarteira!txt_CST <> "" Then
                If Len(TBCarteira!txt_CST) = 4 Then FimCST = Right(TBCarteira!txt_CST, 3) Else FimCST = Right(TBCarteira!txt_CST, 2)
                
                If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                    Set TBCST = CreateObject("adodb.recordset")
                    TBCST.Open "select * from tbl_Detalhes_Nota_CST_ICMS where id_item = " & TBCarteira!Int_codigo, Conexao, adOpenKeyset, adLockOptimistic
                    If TBCST.EOF = False Then
                        If FimCST = "00" Or FimCST = "10" Or FimCST = "20" Or FimCST = "51" Or FimCST = "70" Or FimCST = "90" Then
                            TBProdutividade!Qtdetotalprod = TBProdutividade!Qtdetotalprod + TBCST!Valor_BC 'Base de cálculo ICMS
                            TBProdutividade!Eficiencia = TBProdutividade!Eficiencia + TBCST!Valor_ICMS 'Valor do ICMS
                        End If
                        
                        If FimCST = "10" Or FimCST = "60" Or FimCST = "70" Or FimCST = "90" Or FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                            TBProdutividade!Terceiros = TBProdutividade!Terceiros + TBCST!Valor_BC_ST 'Base de calculo ICMS subst.
                            TBProdutividade!impostos = TBProdutividade!impostos + TBCST!Valor_ICMS_ST 'Valor do ICMS subst.
                            TBProdutividade!Numero15 = TBProdutividade!Numero15 + TBCST!Valor_ICMS_ST
                        End If
                        
                        If FimCST = "101" Or FimCST = "102" Or FimCST = "103" Or FimCST = "201" Or FimCST = "202" Or FimCST = "203" Or FimCST = "300" Or FimCST = "400" Or FimCST = "500" Or FimCST = "900" Then
                            TBProdutividade!Qtdetotalprod = TBProdutividade!Qtdetotalprod + TBCST!Valor_BC 'Base de cálculo ICMS SN
                            TBProdutividade!Valor5 = TBProdutividade!Valor5 + TBCST!Valor_ICMS_SN 'Valor do ICMS SN
                        End If
                    End If
                    TBCST.Close
                End If
            End If
            
            TBProdutividade!Lucro = TBProdutividade!Lucro + IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) 'Valor total de ipi
            TBProdutividade!material = TBProdutividade!material + IIf(IsNull(TBCarteira!Total_PIS_prod), 0, TBCarteira!Total_PIS_prod) 'Valor total pis
            TBProdutividade!Servicos = TBProdutividade!Servicos + IIf(IsNull(TBCarteira!Total_Cofins_prod), 0, TBCarteira!Total_Cofins_prod) 'Valor total cofins
            TBProdutividade!Total = TBProdutividade!Total + IIf(IsNull(TBCarteira!Total_CSLL_prod), 0, TBCarteira!Total_CSLL_prod) 'Total CSLL
            TBProdutividade!Total_peca = TBProdutividade!Total_peca + IIf(IsNull(TBCarteira!Total_IRPJ_prod), 0, TBCarteira!Total_IRPJ_prod) 'Total IRPJ
            TBProdutividade!Refugo = TBProdutividade!Refugo + IIf(IsNull(TBCarteira!Valor_Retencao_PIS), 0, TBCarteira!Valor_Retencao_PIS) 'Total retenção Pis
            TBProdutividade!Numero1 = TBProdutividade!Numero1 + IIf(IsNull(TBCarteira!Valor_Retencao_Cofins), 0, TBCarteira!Valor_Retencao_Cofins) 'Total retenção confins
            TBProdutividade!Numero11 = TBProdutividade!Numero11 + IIf(IsNull(TBCarteira!Valor_frete), 0, TBCarteira!Valor_frete) 'Total do frete
            TBProdutividade!Numero12 = TBProdutividade!Numero12 + IIf(IsNull(TBCarteira!Valor_seguro), 0, TBCarteira!Valor_seguro) 'Total do seguro
            TBProdutividade!Numero13 = TBProdutividade!Numero13 + IIf(IsNull(TBCarteira!Valor_acessorias), 0, TBCarteira!Valor_acessorias) 'Total outras despesas
            TBProdutividade!Valor6 = TBProdutividade!Valor6 + IIf(IsNull(TBCarteira!Valor_desconto), 0, TBCarteira!Valor_desconto) + IIf(IsNull(TBCarteira!Valor_desconto_SUFRAMA), 0, TBCarteira!Valor_desconto_SUFRAMA) 'Total do desconto
            If TBCarteira!retorno = True Then TBProdutividade!Valor4 = TBProdutividade!Valor4 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos (retorno)
            If TBCarteira!Remessa = True Then TBProdutividade!Valor7 = TBProdutividade!Valor7 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total de produtos (remessa)
            
            'Valor total faturado
            TBProdutividade!Numero15 = TBProdutividade!Numero15 + (IIf(IsNull(TBCarteira!dbl_valoripi), 0, TBCarteira!dbl_valoripi) + IIf(IsNull(TBCarteira!Valor_frete), 0, TBCarteira!Valor_frete) + IIf(IsNull(TBCarteira!Valor_seguro), 0, TBCarteira!Valor_seguro) + IIf(IsNull(TBCarteira!Valor_acessorias), 0, TBCarteira!Valor_acessorias)) - (IIf(IsNull(TBCarteira!Valor_Retencao_PIS), 0, TBCarteira!Valor_Retencao_PIS) + IIf(IsNull(TBCarteira!Valor_Retencao_Cofins), 0, TBCarteira!Valor_Retencao_Cofins) + IIf(IsNull(TBCarteira!Valor_desconto), 0, TBCarteira!Valor_desconto) + IIf(IsNull(TBCarteira!Valor_desconto_SUFRAMA), 0, TBCarteira!Valor_desconto_SUFRAMA))
        'End If
    Else
        TBProdutividade!Numero2 = TBProdutividade!Numero2 + IIf(IsNull(TBCarteira!int_Qtd), 0, TBCarteira!int_Qtd) 'Quantidade de serviços
        TBProdutividade!Numero3 = TBProdutividade!Numero3 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total serviço
        If TBCarteira!Retem = True Then
            TBProdutividade!Numero4 = TBProdutividade!Numero4 + IIf(IsNull(TBCarteira!Total_PIS_serv), 0, TBCarteira!Total_PIS_serv) 'Total pis serviços
            TBProdutividade!Numero5 = TBProdutividade!Numero5 + IIf(IsNull(TBCarteira!Total_Cofins_serv), 0, TBCarteira!Total_Cofins_serv) 'Total cofins serviços
            TBProdutividade!Numero6 = TBProdutividade!Numero6 + IIf(IsNull(TBCarteira!Total_CSLL_serv), 0, TBCarteira!Total_CSLL_serv) 'Total CSLL serviços
            TBProdutividade!Numero7 = TBProdutividade!Numero7 + IIf(IsNull(TBCarteira!VlrISS), 0, TBCarteira!VlrISS) 'Valor total de ISS
            TBProdutividade!Numero8 = TBProdutividade!Numero8 + IIf(IsNull(TBCarteira!Total_INSS_serv), 0, TBCarteira!Total_INSS_serv) 'Total de INSS serviços
            TBProdutividade!Numero9 = TBProdutividade!Numero9 + IIf(IsNull(TBCarteira!Total_IRPJ_serv), 0, TBCarteira!Total_IRPJ_serv) 'Total de IRPJ serviços
            TBProdutividade!Numero10 = TBProdutividade!Numero10 + IIf(IsNull(TBCarteira!Total_IRRF_serv), 0, TBCarteira!Total_IRRF_serv) 'Total de IRRF serviços
        End If
        TBProdutividade!Numero15 = TBProdutividade!Numero15 + IIf(IsNull(TBCarteira!dbl_ValorTotal), 0, TBCarteira!dbl_ValorTotal) 'Valor total faturado
    End If
End If
TBProdutividade.Update
TBProdutividade.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcEnviaDadosTotaisNF()
On Error GoTo tratar_erro

Set TBTotaisnota = CreateObject("adodb.recordset")
TBTotaisnota.Open "Select * from tbl_Totais_Nota where id_nota = " & TBCarteira!ID, Conexao, adOpenKeyset, adLockOptimistic
If TBTotaisnota.EOF = False Then
    'Remessa
    ValorPago = 0
    ValoresParcelas = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(dbl_ValorTotal) as ValorPago, Sum(int_Qtd) as ValoresParcelas from tbl_Detalhes_Nota where id_nota = " & TBCarteira!ID & " and Remessa = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ValorPago = IIf(IsNull(TBAbrir!ValorPago), 0, TBAbrir!ValorPago)
        ValoresParcelas = IIf(IsNull(TBAbrir!ValoresParcelas), 0, TBAbrir!ValoresParcelas)
    End If
    
    'Retorno
    ValorPagar = 0
    Valorparcela = 0
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Sum(dbl_ValorTotal) as ValorPagar, Sum(int_Qtd) as Valorparcela from tbl_Detalhes_Nota where id_nota = " & TBCarteira!ID & " and Retorno = 'True'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        ValorPagar = IIf(IsNull(TBAbrir!ValorPagar), 0, TBAbrir!ValorPagar)
        Valorparcela = IIf(IsNull(TBAbrir!Valorparcela), 0, TBAbrir!Valorparcela)
    End If
    
    If optResumido.Value = True And TBCarteira!Alterar = False Then
        'Produtos
        'Quantidade de produtos
        TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + IIf(IsNull(TBTotaisnota!Qtde_total_prod), 0, TBTotaisnota!Qtde_total_prod)
        If Chk_remessa.Value = 1 Then TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + ValoresParcelas
        If Chk_retorno.Value = 1 Then TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + Valorparcela
        
        'Valor total de produtos
        TBProdutividade!qtdeNC = TBProdutividade!qtdeNC + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Produtos), 0, TBTotaisnota!dbl_Valor_Total_Produtos)
        If Chk_remessa.Value = 0 Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC - ValorPago
        If Chk_retorno.Value = 0 Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC - ValorPagar
        
        If IsNull(TBTotaisnota!Valor_total_ICMS_SN) = False And TBTotaisnota!Valor_total_ICMS_SN > 0 Then
            TBProdutividade!Qtdetotalprod = TBProdutividade!Qtdetotalprod + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota) - ValorPagar 'Base de cálculo ICMS SN
        Else
            TBProdutividade!Qtdetotalprod = TBProdutividade!Qtdetotalprod + IIf(IsNull(TBTotaisnota!dbl_Base_ICMS), 0, TBTotaisnota!dbl_Base_ICMS) 'Base de cálculo ICMS
        End If
        
        TBProdutividade!Valor7 = TBProdutividade!Valor7 + ValorPago 'Remessa
        TBProdutividade!Valor4 = TBProdutividade!Valor4 + IIf(IsNull(TBTotaisnota!Total_retorno), 0, TBTotaisnota!Total_retorno)  'Retorno
        TBProdutividade!Eficiencia = TBProdutividade!Eficiencia + IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS), 0, TBTotaisnota!dbl_Valor_ICMS) 'Valor do ICMS
        TBProdutividade!Terceiros = TBProdutividade!Terceiros + IIf(IsNull(TBTotaisnota!dbl_Base_ICMS_Subst), 0, TBTotaisnota!dbl_Base_ICMS_Subst)  'Base de calculo ICMS subst.
        TBProdutividade!impostos = TBProdutividade!impostos + IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS_Subst), 0, TBTotaisnota!dbl_Valor_ICMS_Subst)  'Valor do ICMS subst.
        TBProdutividade!Valor5 = TBProdutividade!Valor5 + IIf(IsNull(TBTotaisnota!Valor_total_ICMS_SN), 0, TBTotaisnota!Valor_total_ICMS_SN)  'Valor do ICMS SN
        TBProdutividade!Lucro = TBProdutividade!Lucro + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), 0, TBTotaisnota!dbl_Valor_Total_IPI)  'Valor total de ipi
        TBProdutividade!material = TBProdutividade!material + IIf(IsNull(TBTotaisnota!Total_PIS_prod), 0, TBTotaisnota!Total_PIS_prod) 'Valor total pis
        TBProdutividade!Servicos = TBProdutividade!Servicos + IIf(IsNull(TBTotaisnota!Total_Cofins_prod), 0, TBTotaisnota!Total_Cofins_prod) 'Valor total cofins
        TBProdutividade!Total = TBProdutividade!Total + IIf(IsNull(TBTotaisnota!Total_CSLL_prod), 0, TBTotaisnota!Total_CSLL_prod) 'Total CSLL
        TBProdutividade!Total_peca = TBProdutividade!Total_peca + IIf(IsNull(TBTotaisnota!Total_IRPJ_prod), 0, TBTotaisnota!Total_IRPJ_prod) 'Total IRPJ
        TBProdutividade!Refugo = TBProdutividade!Refugo + IIf(IsNull(TBTotaisnota!Total_retencao_PIS), 0, TBTotaisnota!Total_retencao_PIS) 'Total retenção Pis
        TBProdutividade!Numero1 = TBProdutividade!Numero1 + IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), 0, TBTotaisnota!Total_retencao_Cofins) 'Total retenção confins
        TBProdutividade!Numero11 = TBProdutividade!Numero11 + IIf(IsNull(TBTotaisnota!dbl_Valor_Frete), 0, TBTotaisnota!dbl_Valor_Frete) 'Valor Frete
        TBProdutividade!Numero12 = TBProdutividade!Numero12 + IIf(IsNull(TBTotaisnota!dbl_Valor_Seguro), 0, TBTotaisnota!dbl_Valor_Seguro) 'Valor Seguro
        TBProdutividade!Numero13 = TBProdutividade!Numero13 + IIf(IsNull(TBTotaisnota!dbl_Desp_Adicionais), 0, TBTotaisnota!dbl_Desp_Adicionais) 'Despeças adicionais
        TBProdutividade!Valor6 = TBProdutividade!Valor6 + IIf(IsNull(TBTotaisnota!Valor_total_desconto), 0, TBTotaisnota!Valor_total_desconto) + IIf(IsNull(TBTotaisnota!Valor_total_desconto_SUFRAMA), 0, TBTotaisnota!Valor_total_desconto_SUFRAMA) 'Desconto
        
        'Serviços
        TBProdutividade!Numero2 = TBProdutividade!Numero2 + IIf(IsNull(TBTotaisnota!Qtde_total_serv), 0, TBTotaisnota!Qtde_total_serv) 'Quantidade de serviços
        TBProdutividade!Numero3 = TBProdutividade!Numero3 + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), 0, TBTotaisnota!dbl_Valor_Total_Nota_Serv) 'Valor total serviço
        TBProdutividade!Numero4 = TBProdutividade!Numero4 + IIf(IsNull(TBTotaisnota!Total_PIS_serv), 0, TBTotaisnota!Total_PIS_serv) 'Total pis serviços
        TBProdutividade!Numero5 = TBProdutividade!Numero5 + IIf(IsNull(TBTotaisnota!Total_Cofins_serv), 0, TBTotaisnota!Total_Cofins_serv) 'Total cofins serviços
        TBProdutividade!Numero6 = TBProdutividade!Numero6 + IIf(IsNull(TBTotaisnota!Total_CSLL_serv), 0, TBTotaisnota!Total_CSLL_serv) 'Total CSLL serviços
        TBProdutividade!Numero7 = TBProdutividade!Numero7 + IIf(IsNull(TBTotaisnota!dbl_valor_total_iss), 0, TBTotaisnota!dbl_valor_total_iss) 'Valor total de ISS
        TBProdutividade!Numero8 = TBProdutividade!Numero8 + IIf(IsNull(TBTotaisnota!Total_INSS_serv), 0, TBTotaisnota!Total_INSS_serv) 'Total de INSS serviços
        TBProdutividade!Numero9 = TBProdutividade!Numero9 + IIf(IsNull(TBTotaisnota!Total_IRPJ_serv), 0, TBTotaisnota!Total_IRPJ_serv) 'Total de IRPJ serviços
        TBProdutividade!Numero10 = TBProdutividade!Numero10 + IIf(IsNull(TBTotaisnota!Total_IRRF_serv), 0, TBTotaisnota!Total_IRRF_serv) 'Total de IRRF serviços
        
        'Totais
        TBProdutividade!Numero14 = TBProdutividade!Numero14 + IIf(IsNull(TBTotaisnota!Total_DAS), 0, TBTotaisnota!Total_DAS) 'Total das
        TBProdutividade!Valor1 = TBProdutividade!Valor1 + IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
        Valor1 = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
        Valor2 = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), 0, TBTotaisnota!Total_retencao_PIS)
        Valor3 = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), 0, TBTotaisnota!Total_retencao_Cofins)
        
        'Valor total faturado
        TBProdutividade!Numero15 = TBProdutividade!Numero15 + (Valor1 - Valor2 - Valor3)
        If Chk_retorno.Value = 0 Then TBProdutividade!Numero15 = TBProdutividade!Numero15 - ValorPagar
        If Chk_remessa.Value = 0 Then TBProdutividade!Numero15 = TBProdutividade!Numero15 - ValorPago
    Else
        'Produtos
        'Quantidade de produtos
        TBProdutividade!qtdeOK = IIf(IsNull(TBTotaisnota!Qtde_total_prod), 0, TBTotaisnota!Qtde_total_prod)
        If Chk_remessa.Value = 1 Then TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + ValoresParcelas
        If Chk_retorno.Value = 1 Then TBProdutividade!qtdeOK = TBProdutividade!qtdeOK + Valorparcela
        
        'Valor total de produtos
        TBProdutividade!qtdeNC = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Produtos), 0, TBTotaisnota!dbl_Valor_Total_Produtos)
        If Chk_remessa.Value = 0 Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC - ValorPago
        If Chk_retorno.Value = 0 Then TBProdutividade!qtdeNC = TBProdutividade!qtdeNC - ValorPagar
        
        TBProdutividade!Qtdetotalprod = IIf(IsNull(TBTotaisnota!dbl_Base_ICMS), 0, TBTotaisnota!dbl_Base_ICMS) 'Base de cálculo ICMS
        TBProdutividade!Valor7 = ValorPago 'Remessa
        TBProdutividade!Valor4 = IIf(IsNull(TBTotaisnota!Total_retorno), 0, TBTotaisnota!Total_retorno)  'Retorno
        TBProdutividade!Eficiencia = IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS), 0, TBTotaisnota!dbl_Valor_ICMS) 'Valor do ICMS
        TBProdutividade!Terceiros = IIf(IsNull(TBTotaisnota!dbl_Base_ICMS_Subst), 0, TBTotaisnota!dbl_Base_ICMS_Subst)  'Base de calculo ICMS subst.
        TBProdutividade!impostos = IIf(IsNull(TBTotaisnota!dbl_Valor_ICMS_Subst), 0, TBTotaisnota!dbl_Valor_ICMS_Subst)  'Valor do ICMS subst.
        TBProdutividade!Valor5 = IIf(IsNull(TBTotaisnota!Valor_total_ICMS_SN), 0, TBTotaisnota!Valor_total_ICMS_SN)  'Valor do ICMS SN
        TBProdutividade!Lucro = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_IPI), 0, TBTotaisnota!dbl_Valor_Total_IPI)  'Valor total de ipi
        TBProdutividade!material = IIf(IsNull(TBTotaisnota!Total_PIS_prod), 0, TBTotaisnota!Total_PIS_prod) 'Valor total pis
        TBProdutividade!Servicos = IIf(IsNull(TBTotaisnota!Total_Cofins_prod), 0, TBTotaisnota!Total_Cofins_prod) 'Valor total cofins
        TBProdutividade!Total = IIf(IsNull(TBTotaisnota!Total_CSLL_prod), 0, TBTotaisnota!Total_CSLL_prod) 'Total CSLL
        TBProdutividade!Total_peca = IIf(IsNull(TBTotaisnota!Total_IRPJ_prod), 0, TBTotaisnota!Total_IRPJ_prod) 'Total IRPJ
        TBProdutividade!Refugo = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), 0, TBTotaisnota!Total_retencao_PIS) 'Total retenção Pis
        TBProdutividade!Numero1 = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), 0, TBTotaisnota!Total_retencao_Cofins) 'Total retenção confins
        TBProdutividade!Numero11 = IIf(IsNull(TBTotaisnota!dbl_Valor_Frete), 0, TBTotaisnota!dbl_Valor_Frete) 'Valor Frete
        TBProdutividade!Numero12 = IIf(IsNull(TBTotaisnota!dbl_Valor_Seguro), 0, TBTotaisnota!dbl_Valor_Seguro) 'Valor Seguro
        TBProdutividade!Numero13 = IIf(IsNull(TBTotaisnota!dbl_Desp_Adicionais), 0, TBTotaisnota!dbl_Desp_Adicionais) 'Despeças adicionais
        TBProdutividade!Valor6 = IIf(IsNull(TBTotaisnota!Valor_total_desconto), 0, TBTotaisnota!Valor_total_desconto) + IIf(IsNull(TBTotaisnota!Valor_total_desconto_SUFRAMA), 0, TBTotaisnota!Valor_total_desconto_SUFRAMA) 'Desconto
        
        'Serviços
        TBProdutividade!Numero2 = IIf(IsNull(TBTotaisnota!Qtde_total_serv), 0, TBTotaisnota!Qtde_total_serv) 'Quantidade de serviços
        TBProdutividade!Numero3 = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota_Serv), 0, TBTotaisnota!dbl_Valor_Total_Nota_Serv) 'Valor total serviço
        TBProdutividade!Numero4 = IIf(IsNull(TBTotaisnota!Total_PIS_serv), 0, TBTotaisnota!Total_PIS_serv) 'Total pis serviços
        TBProdutividade!Numero5 = IIf(IsNull(TBTotaisnota!Total_Cofins_serv), 0, TBTotaisnota!Total_Cofins_serv) 'Total cofins serviços
        TBProdutividade!Numero6 = IIf(IsNull(TBTotaisnota!Total_CSLL_serv), 0, TBTotaisnota!Total_CSLL_serv) 'Total CSLL serviços
        TBProdutividade!Numero7 = IIf(IsNull(TBTotaisnota!dbl_valor_total_iss), 0, TBTotaisnota!dbl_valor_total_iss) 'Valor total de ISS
        TBProdutividade!Numero8 = IIf(IsNull(TBTotaisnota!Total_INSS_serv), 0, TBTotaisnota!Total_INSS_serv) 'Total de INSS serviços
        TBProdutividade!Numero9 = IIf(IsNull(TBTotaisnota!Total_IRPJ_serv), 0, TBTotaisnota!Total_IRPJ_serv) 'Total de IRPJ serviços
        TBProdutividade!Numero10 = IIf(IsNull(TBTotaisnota!Total_IRRF_serv), 0, TBTotaisnota!Total_IRRF_serv) 'Total de IRRF serviços
        
        'Totais
        TBProdutividade!Numero14 = IIf(IsNull(TBTotaisnota!Total_DAS), 0, TBTotaisnota!Total_DAS) 'Total das
        TBProdutividade!Valor1 = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
        Valor1 = IIf(IsNull(TBTotaisnota!dbl_Valor_Total_Nota), 0, TBTotaisnota!dbl_Valor_Total_Nota)
        Valor2 = IIf(IsNull(TBTotaisnota!Total_retencao_PIS), 0, TBTotaisnota!Total_retencao_PIS)
        Valor3 = IIf(IsNull(TBTotaisnota!Total_retencao_Cofins), 0, TBTotaisnota!Total_retencao_Cofins)
        
        'Valor total faturado
        TBProdutividade!Numero15 = TBProdutividade!Numero15 + (Valor1 - Valor2 - Valor3)
        If Chk_retorno.Value = 0 Then TBProdutividade!Numero15 = TBProdutividade!Numero15 - ValorPagar
        If Chk_remessa.Value = 0 Then TBProdutividade!Numero15 = TBProdutividade!Numero15 - ValorPago
    End If
End If
TBTotaisnota.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcGravarTotalizacoes()
On Error GoTo tratar_erro

Qtde = 0
vlrTotalProd = 0
ValorPago = 0
ValorPagoParcial = 0
ValorICMS = 0
quantestoque = 0
ICMS_SN = 0
VlrIPI = 0
Valor_PIS_Prod = 0
Valor_Cofins_Prod = 0
Valor_CSLL_Prod = 0
Valor_IRPJ_Prod = 0
Valor_Retencao_PIS = 0
Valor_Retencao_Cofins = 0
VLFRETE = 0
VLSEGURO = 0
VLOUTROS = 0
VLICMSOUTROS = 0

Qtd = 0
VlrTotalServ = 0
Valor_PIS_Serv = 0
Valor_Cofins_Serv = 0
Valor_CSLL_Serv = 0
Valor_ISS_Serv = 0
Valor_INSS_Serv = 0
Valor_IRPJ_Serv = 0
Valor_IRRF_Serv = 0

DAS = 0
Valor1 = 0
ValorTotal = 0

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from Producao_Relatorios_Total where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then TBAbrir.AddNew
If Opt_terceiros.Value = True Then
    If Opt_dt_entrada.Value = True Then TBAbrir!Totalutilizada = "T" Else TBAbrir!Totalutilizada = "T1"
Else
    TBAbrir!Totalutilizada = "P"
End If
If Cmb_empresa = "Todas" Then TBAbrir!Numero18 = 0 Else TBAbrir!Numero18 = Cmb_empresa.ItemData(Cmb_empresa.ListIndex)
TBAbrir!Data_inicial = msk_fltInicio.Value
TBAbrir!Data_final = msk_fltFim.Value
If Opt_individual.Value = True Then
    If cmbTexto <> "" Then TBAbrir!Texto = cmbfiltrarpor & ") : " & cmbTexto Else TBAbrir!Texto = cmbfiltrarpor & ")"
Else
    TBAbrir!Texto = cmbfiltrarpor
End If
TBAbrir!Responsavel = pubUsuario
TBAbrir!Modulo = Formulario
CamposFiltro = "Sum(QtdeOK) as Qtde, Sum(QtdeNC) as vlrTotalProd, Sum(Valor7) as ValorPagoParcial, Sum(Valor4) as ValorPago, Sum(Eficiencia) as ValorICMS, Sum(Impostos) as quantestoque, Sum(Valor5) as ICMS_SN, Sum(Lucro) as VlrIPI, Sum(material) as Valor_PIS_Prod, Sum(Servicos) as Valor_Cofins_Prod, Sum (Total) as Valor_CSLL_Prod, Sum(Total_peca) as Valor_IRPJ_Prod, Sum(Refugo) as Valor_Retencao_PIS, Sum(Numero1) as Valor_Retencao_Cofins,  Sum(Numero11) as VLFRETE, Sum(Numero12) as VLSEGURO, Sum(Numero13) as VLOUTROS, Sum(Valor6) as VLICMSOUTROS, Sum(Numero2) as Qtd, Sum(Numero3) as VlrTotalServ, Sum(Numero4) as Valor_PIS_Serv, Sum(Numero5) as Valor_Cofins_Serv, Sum(Numero6) as Valor_CSLL_Serv, Sum(Numero7) as vlriss, Sum(Numero8) as Valor_INSS_Serv, Sum(Numero9) as Valor_IRPJ_Serv, Sum(Numero10) as Valor_IRRF_Serv, Sum(Numero14) as DAS, Sum(Valor1) as Valor1, Sum(Numero15) as Valortotal"
Set TBproducao = CreateObject("adodb.recordset")
TBproducao.Open "Select " & CamposFiltro & " from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBproducao.EOF = False Then
    'Produtos
    Qtde = IIf(IsNull(TBproducao!Qtde), 0, TBproducao!Qtde) 'Quantidade total de produtos
    vlrTotalProd = IIf(IsNull(TBproducao!vlrTotalProd), 0, TBproducao!vlrTotalProd) 'Valor total de produtos
    ValorPagoParcial = IIf(IsNull(TBproducao!ValorPagoParcial), 0, TBproducao!ValorPagoParcial) 'Valor total de produtos (Remessa)
    ValorPago = IIf(IsNull(TBproducao!ValorPago), 0, TBproducao!ValorPago) 'Valor total de produtos (Retorno)
    ValorICMS = IIf(IsNull(TBproducao!ValorICMS), 0, TBproducao!ValorICMS) 'Valor total de icms
    quantestoque = IIf(IsNull(TBproducao!quantestoque), 0, TBproducao!quantestoque) 'valor icms subst
    ICMS_SN = IIf(IsNull(TBproducao!ICMS_SN), 0, TBproducao!ICMS_SN) 'valor icms SN
    VlrIPI = IIf(IsNull(TBproducao!VlrIPI), 0, TBproducao!VlrIPI) 'Valor total de ipi
    Valor_PIS_Prod = IIf(IsNull(TBproducao!Valor_PIS_Prod), 0, TBproducao!Valor_PIS_Prod) 'Valor total pis
    Valor_Cofins_Prod = IIf(IsNull(TBproducao!Valor_Cofins_Prod), 0, TBproducao!Valor_Cofins_Prod) 'Valor total cofins
    Valor_CSLL_Prod = IIf(IsNull(TBproducao!Valor_CSLL_Prod), 0, TBproducao!Valor_CSLL_Prod) 'Total CSLL
    Valor_IRPJ_Prod = IIf(IsNull(TBproducao!Valor_IRPJ_Prod), 0, TBproducao!Valor_IRPJ_Prod) 'Total IRPJ
    Valor_Retencao_PIS = IIf(IsNull(TBproducao!Valor_Retencao_PIS), 0, TBproducao!Valor_Retencao_PIS) 'Total retenção Pis
    Valor_Retencao_Cofins = IIf(IsNull(TBproducao!Valor_Retencao_Cofins), 0, TBproducao!Valor_Retencao_Cofins) 'Total retenção confins
    VLFRETE = IIf(IsNull(TBproducao!VLFRETE), 0, TBproducao!VLFRETE) 'Valor Frete
    VLSEGURO = IIf(IsNull(TBproducao!VLSEGURO), 0, TBproducao!VLSEGURO) 'Valor Seguro
    VLOUTROS = IIf(IsNull(TBproducao!VLOUTROS), 0, TBproducao!VLOUTROS) 'Despeças adicionais
    VLICMSOUTROS = IIf(IsNull(TBproducao!VLICMSOUTROS), 0, TBproducao!VLICMSOUTROS) 'Desconto
    
    'Serviços
    Qtd = IIf(IsNull(TBproducao!Qtd), 0, TBproducao!Qtd) 'Quantidade de serviços
    VlrTotalServ = IIf(IsNull(TBproducao!VlrTotalServ), 0, TBproducao!VlrTotalServ) 'Valor total serviço
    Valor_PIS_Serv = IIf(IsNull(TBproducao!Valor_PIS_Serv), 0, TBproducao!Valor_PIS_Serv) 'Total pis serviços
    Valor_Cofins_Serv = IIf(IsNull(TBproducao!Valor_Cofins_Serv), 0, TBproducao!Valor_Cofins_Serv) 'Total cofins serviços
    Valor_CSLL_Serv = IIf(IsNull(TBproducao!Valor_CSLL_Serv), 0, TBproducao!Valor_CSLL_Serv) 'Total CSLL serviços
    Valor_ISS_Serv = IIf(IsNull(TBproducao!VlrISS), 0, TBproducao!VlrISS) 'Valor total de ISS
    Valor_INSS_Serv = IIf(IsNull(TBproducao!Valor_INSS_Serv), 0, TBproducao!Valor_INSS_Serv) 'Total de INSS serviços
    Valor_IRPJ_Serv = IIf(IsNull(TBproducao!Valor_IRPJ_Serv), 0, TBproducao!Valor_IRPJ_Serv) 'Total de IRPJ serviços
    Valor_IRRF_Serv = IIf(IsNull(TBproducao!Valor_IRRF_Serv), 0, TBproducao!Valor_IRRF_Serv) 'Total de IRRF serviços
    
    'Totais
    DAS = IIf(IsNull(TBproducao!DAS), 0, TBproducao!DAS) 'Total das
    Valor1 = IIf(IsNull(TBproducao!Valor1), 0, TBproducao!Valor1) 'Valor total da nota
    
    ValorTotal = IIf(IsNull(TBproducao!ValorTotal), 0, TBproducao!ValorTotal) 'Valor total faturado
End If
TBproducao.Close

'Produtos
TBAbrir!QtdePrevista = Qtde
TBAbrir!QtdeProduzida = vlrTotalProd
TBAbrir!Numero14 = ValorPagoParcial
TBAbrir!Numero15 = ValorPago
TBAbrir!qtdeNC = ValorICMS
TBAbrir!Numero13 = quantestoque
TBAbrir!Numero16 = ICMS_SN
TBAbrir!QtdeOrdem = VlrIPI
TBAbrir!CustoMat = Valor_PIS_Prod
TBAbrir!CustoObra = Valor_Cofins_Prod
TBAbrir!Terceros = Valor_CSLL_Prod
TBAbrir!Lucro = Valor_IRPJ_Prod
TBAbrir!Valor1 = Valor_Retencao_PIS
TBAbrir!Valor2 = Valor_Retencao_Cofins
TBAbrir!Numero7 = VLFRETE
TBAbrir!Numero8 = VLSEGURO
TBAbrir!Numero9 = VLOUTROS
TBAbrir!Numero17 = VLICMSOUTROS

'Serviços
TBAbrir!Valor3 = Qtd
TBAbrir!Total1 = VlrTotalServ
TBAbrir!Total2 = Valor_PIS_Serv
TBAbrir!Numero1 = Valor_Cofins_Serv
TBAbrir!Numero2 = Valor_CSLL_Serv
TBAbrir!Numero3 = Valor_ISS_Serv
TBAbrir!Numero4 = Valor_INSS_Serv
TBAbrir!Numero5 = Valor_IRPJ_Serv
TBAbrir!Numero6 = Valor_IRRF_Serv

'Totais
TBAbrir!Numero10 = DAS

If Chk_remessa.Value = 0 And Chk_retorno.Value = 0 Then
    ValorTotal = ValorTotal - (ValorPagoParcial + ValorPago)
ElseIf Chk_remessa.Value = 1 Then
        ValorTotal = ValorTotal - ValorPago
    ElseIf Chk_retorno.Value = 1 Then
            ValorTotal = ValorTotal - ValorPagoParcial
End If
TBAbrir!Numero11 = ValorTotal

If Txt_limite <> "" And Txt_limite <> "0" Then
    Qtde = IIf(IsNull(TBAbrir!Numero11), 0, Format(TBAbrir!Numero11, "###,##0.00"))
    TBAbrir!Numero12 = (Qtde / Valor1) * 100
Else
    TBAbrir!Numero12 = 100
End If

TBAbrir.Update
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub msk_fltInicio_Change()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_comparativo_Click()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais
If Opt_comparativo.Value = True Then
    ProcCarregaComboFiltrarPor Opt_terceiros.Value, Opt_comparativo.Value
    
    optDetalhado.Enabled = False
    optResumido.Value = True
    With Txt_limite
        .Locked = False
        .TabStop = True
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_dt_emissao_Click()
On Error GoTo tratar_erro

ListaNF.ColumnHeaders.Item(3).Text = "Dt. emissão"
ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_dt_entrada_Click()
On Error GoTo tratar_erro

ListaNF.ColumnHeaders.Item(3).Text = "Dt. entrada"
ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_individual_Click()
On Error GoTo tratar_erro

If Opt_individual.Value = True Then
    ProcCarregaComboFiltrarPor Opt_terceiros.Value, Opt_comparativo.Value
    
    ProcCarregarComboTexto
    optDetalhado.Value = True
    optDetalhado.Enabled = True
    cmbTexto.Enabled = True
    With Txt_limite
        .Locked = True
        .TabStop = False
    End With
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_propria_Click()
On Error GoTo tratar_erro

optSaida.Value = True
optSaida.Enabled = True
Opt_dt_entrada.Visible = False
Opt_dt_emissao.Visible = False
With ListaNF.ColumnHeaders
    If Opt_terceiros.Value = True Then .Item(8).Text = "Emitente" Else .Item(8).Text = "Destinatário"
    If Opt_dt_entrada.Value = True Then .Item(3).Text = "Dt. entrada" Else .Item(3).Text = "Dt. emissão"
End With
ProcCarregaComboFiltrarPor Opt_terceiros.Value, Opt_comparativo.Value
ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Opt_terceiros_Click()
On Error GoTo tratar_erro

optSaida.Value = False
optSaida.Enabled = False
optEntrada.Value = True
Opt_dt_entrada.Visible = True
Opt_dt_emissao.Visible = True
With ListaNF.ColumnHeaders
    If Opt_terceiros.Value = True Then .Item(8).Text = "Emitente" Else .Item(8).Text = "Destinatário"
    If Opt_dt_entrada.Value = True Then .Item(3).Text = "Dt. entrada" Else .Item(3).Text = "Dt. emissão"
End With
ProcCarregaComboFiltrarPor Opt_terceiros.Value, Opt_comparativo.Value
ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optDetalhado_Click()
On Error GoTo tratar_erro

If optDetalhado.Value = True Then
    ListaNF.ListItems.Clear
    ListaNF.Visible = True
    ListaNF1.ListItems.Clear
    ListaNF1.Visible = False
    ProcLimpaCamposTotais
End If
ProcAcertaColunas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optEntrada_Click()
On Error GoTo tratar_erro
    
ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optResumido_Click()
On Error GoTo tratar_erro

If optResumido.Value = True Then
    ListaNF.ListItems.Clear
    ListaNF.Visible = False
    ListaNF1.ListItems.Clear
    ListaNF1.Visible = True
    ProcLimpaCamposTotais
End If
ProcAcertaColunas

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub optSaida_Click()
On Error GoTo tratar_erro
    
ProcCarregarComboTexto

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificaLimiteRegistros()
On Error GoTo tratar_erro

Contador1 = Txt_limite
Valor_total = 0
Valor1 = 0
Cliente = ""

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select Sum(Numero15) as Valor1 from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    Valor1 = IIf(IsNull(TBLISTA!Valor1), 0, TBLISTA!Valor1)
End If
TBLISTA.Close

Set TBLISTA = CreateObject("adodb.recordset")
TBLISTA.Open "Select * from Producao_Relatorios where Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "' order by Numero15 desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLISTA.EOF = False Then
    If TBLISTA.RecordCount < 10 Then Exit Sub
    Do While Contador1 <> 0
        If Cliente <> TBLISTA!maquina Then
            Valor_total = TBLISTA!Numero15
            Contador1 = Contador1 - 1
        End If
        Cliente = TBLISTA!maquina
        TBLISTA.MoveNext
    Loop
End If
TBLISTA.Close
NovoValor = Replace(Valor_total, ",", ".")
Conexao.Execute "DELETE from Producao_Relatorios where Numero15 < " & NovoValor & " and Responsavel = '" & pubUsuario & "' and Modulo = '" & Formulario & "'"

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Txt_limite_Change()
On Error GoTo tratar_erro

ListaNF.ListItems.Clear
ListaNF1.ListItems.Clear
ProcLimpaCamposTotais

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcAbrir
    Case 2: ProcImprimir
    Case 4: ProcAjuda
    Case 5: ProcSair
End Select

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub
